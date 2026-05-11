import csv
import json
import os
import re
import shutil
import socket
import sqlite3
import subprocess
import sys
import threading
import webbrowser
import zipfile
from contextlib import closing
from datetime import datetime, timedelta
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import mcp_server
from export_chat import _extract_content, _msg_type_str, _resolve_sender


ROOT = Path(__file__).resolve().parent
EXPORT_DIR = Path(os.environ.get("WECHAT_EXPORT_DIR", ROOT / "exports"))
STATE_FILE = EXPORT_DIR / "export_state.json"
SCHEDULE_FILE = EXPORT_DIR / "export_schedules.json"
SCHEDULE_CHECK_SECONDS = 30
SYSTEM_TYPES = {10000, 10002}
URL_RE = re.compile(r"https?://[^\s<>'\"]+")
RANGE_MODES = {"today", "2d", "3d", "7d", "30d", "all", "since_last", "custom_days", "custom"}
SCHEDULE_FREQUENCIES = {"daily", "weekly", "every_hours"}
WEEKDAY_LABELS = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
_SCHEDULE_LOCK = threading.RLock()
_SCHEDULE_WAKE = threading.Event()
_SCHEDULE_RUNNING = set()
_SCHEDULER_THREAD = None


def _ensure_export_dir():
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)


def _load_state():
    if not STATE_FILE.exists():
        return {}
    try:
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def _save_state(state):
    _ensure_export_dir()
    STATE_FILE.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def _load_schedules():
    if not SCHEDULE_FILE.exists():
        return {"jobs": []}
    try:
        data = json.loads(SCHEDULE_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"jobs": []}
    if isinstance(data, list):
        return {"jobs": data}
    if not isinstance(data, dict):
        return {"jobs": []}
    jobs = data.get("jobs", [])
    return {"jobs": jobs if isinstance(jobs, list) else []}


def _save_schedules(data):
    _ensure_export_dir()
    jobs = data.get("jobs", []) if isinstance(data, dict) else []
    SCHEDULE_FILE.write_text(json.dumps({"jobs": jobs}, ensure_ascii=False, indent=2), encoding="utf-8")


def _safe_filename(value):
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", value).strip(" ._")
    return value[:80] or "wechat_export"


def _safe_path(path):
    try:
        target = Path(path).expanduser().resolve()
        root = EXPORT_DIR.expanduser().resolve()
        if target == root or root in target.parents:
            return target
    except (OSError, RuntimeError, TypeError, ValueError):
        pass
    return None


def _safe_media_name(value):
    value = re.sub(r"[^A-Za-z0-9_.-]+", "_", str(value or "")).strip("._")
    return value or "media"


def _format_time(ts, minutes=False):
    if not ts:
        return ""
    fmt = "%Y-%m-%d %H:%M" if minutes else "%Y-%m-%d %H:%M:%S"
    return datetime.fromtimestamp(ts).strftime(fmt)


def _format_date(ts):
    return datetime.fromtimestamp(ts).strftime("%Y-%m-%d") if ts else ""


def _range_display(mode, start_ts=None, end_ts=None):
    if start_ts and end_ts:
        start = _format_date(start_ts)
        end = _format_date(end_ts)
        return start if start == end else f"{start} 至 {end}"
    if mode == "since_last":
        return f"{_format_date(start_ts)} 起" if start_ts else "继续导出"
    if mode == "all":
        return "全部记录"
    if start_ts:
        return f"{_format_date(start_ts)} 起"
    if end_ts:
        return f"截至 {_format_date(end_ts)}"
    return {
        "today": "今天",
        "2d": "最近2天",
        "3d": "最近3天",
        "7d": "最近7天",
        "30d": "最近30天",
        "custom_days": "自定义天数",
        "custom": "自定义范围",
    }.get(mode, "导出")


def _range_filename_part(mode, start_ts=None, end_ts=None):
    label = _range_display(mode, start_ts, end_ts)
    return _safe_filename(label.replace(" 至 ", "_到_").replace(" ", ""))


def _unique_path(path):
    path = Path(path)
    if not path.exists():
        return path
    counter = 2
    while True:
        candidate = path.with_name(f"{path.stem}_{counter}{path.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def _chat_export_dir(ctx):
    folder = EXPORT_DIR / _safe_filename(ctx["display_name"])
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def _export_path(ctx, mode, start_ts, end_ts, suffix):
    base = f"{_safe_filename(ctx['display_name'])}_{_range_filename_part(mode, start_ts, end_ts)}"
    return _unique_path(_chat_export_dir(ctx) / f"{base}{suffix}")


def _parse_ui_time(value, is_end=False):
    value = (value or "").strip()
    if not value:
        return None
    if "T" in value:
        value = value.replace("T", " ")
    if len(value) == 16:
        value += ":00"
    return mcp_server._parse_time_value(value, "time", is_end=is_end)


def _range_to_timestamps(mode, custom_days=None, custom_start="", custom_end="", last_ts=None):
    now = datetime.now()
    today_start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    end_ts = int(now.timestamp())

    if mode == "today":
        return int(today_start.timestamp()), end_ts
    if mode == "2d":
        return int((today_start - timedelta(days=1)).timestamp()), end_ts
    if mode == "3d":
        return int((today_start - timedelta(days=2)).timestamp()), end_ts
    if mode == "7d":
        return int((today_start - timedelta(days=6)).timestamp()), end_ts
    if mode == "30d":
        return int((today_start - timedelta(days=29)).timestamp()), end_ts
    if mode == "all":
        return None, None
    if mode == "since_last":
        return (int(last_ts) + 1 if last_ts else None), None
    if mode == "custom_days":
        try:
            days = max(1, int(custom_days or 1))
        except (TypeError, ValueError):
            days = 1
        return int((today_start - timedelta(days=days - 1)).timestamp()), end_ts
    if mode == "custom":
        return _parse_ui_time(custom_start, False), _parse_ui_time(custom_end, True)
    return int(today_start.timestamp()), end_ts


def _parse_schedule_time(value):
    value = (value or "08:00").strip()
    match = re.match(r"^([01]?\d|2[0-3]):([0-5]\d)$", value)
    if not match:
        raise ValueError("Schedule time must use HH:MM, for example 08:00")
    return int(match.group(1)), int(match.group(2))


def _compute_next_run(job, now=None):
    now = now or datetime.now()
    hour, minute = _parse_schedule_time(job.get("time", "08:00"))
    frequency = job.get("frequency", "daily")

    if frequency == "daily":
        candidate = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        if candidate <= now:
            candidate += timedelta(days=1)
        return candidate

    if frequency == "weekly":
        weekday = max(0, min(6, int(job.get("weekday", 0))))
        days_ahead = (weekday - now.weekday()) % 7
        candidate = (now + timedelta(days=days_ahead)).replace(hour=hour, minute=minute, second=0, microsecond=0)
        if candidate <= now:
            candidate += timedelta(days=7)
        return candidate

    if frequency == "every_hours":
        interval = max(1, min(168, int(job.get("interval_hours", 1))))
        candidate = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
        if candidate > now:
            candidate -= timedelta(hours=interval)
        while candidate <= now:
            candidate += timedelta(hours=interval)
        return candidate

    raise ValueError(f"Unsupported schedule frequency: {frequency}")


def _schedule_frequency_label(job):
    frequency = job.get("frequency", "daily")
    at = job.get("time", "08:00")
    if frequency == "weekly":
        weekday = max(0, min(6, int(job.get("weekday", 0))))
        return f"每周 {WEEKDAY_LABELS[weekday]} {at}"
    if frequency == "every_hours":
        return f"每 {job.get('interval_hours', 1)} 小时，{at[-2:]} 分执行"
    return f"每天 {at}"


def _schedule_view(job):
    item = dict(job)
    item["frequency_label"] = _schedule_frequency_label(job)
    item["next_run_time"] = _format_time(item.get("next_run"))
    item["last_run_time"] = _format_time(item.get("last_run_at"))
    item["last_started_time"] = _format_time(item.get("last_started_at"))
    return item


def list_schedules():
    with _SCHEDULE_LOCK:
        return [_schedule_view(job) for job in _load_schedules().get("jobs", [])]


def upsert_schedule(payload):
    ctx = mcp_server._resolve_chat_context((payload.get("username") or "").strip())
    if not ctx:
        raise ValueError("Chat not found")

    export_format = payload.get("format", "csv")
    if export_format not in {"csv", "ai_package"}:
        raise ValueError("Unsupported export format")

    range_mode = payload.get("range_mode", "since_last")
    if range_mode not in RANGE_MODES:
        raise ValueError("Unsupported export range")

    frequency = payload.get("frequency", "daily")
    if frequency not in SCHEDULE_FREQUENCIES:
        raise ValueError("Unsupported schedule frequency")

    schedule_time = payload.get("time", "08:00")
    _parse_schedule_time(schedule_time)

    job_id = ctx["username"]
    now_ts = int(datetime.now().timestamp())
    job = {
        "id": job_id,
        "username": ctx["username"],
        "display_name": ctx["display_name"],
        "active": bool(payload.get("active", True)),
        "format": export_format,
        "range_mode": range_mode,
        "custom_days": payload.get("custom_days") or "",
        "custom_start": payload.get("custom_start") or "",
        "custom_end": payload.get("custom_end") or "",
        "frequency": frequency,
        "time": schedule_time,
        "weekday": max(0, min(6, int(payload.get("weekday", 0) or 0))),
        "interval_hours": max(1, min(168, int(payload.get("interval_hours", 1) or 1))),
        "updated_at": now_ts,
    }
    job["next_run"] = int(_compute_next_run(job).timestamp()) if job["active"] else None

    with _SCHEDULE_LOCK:
        data = _load_schedules()
        jobs = data.get("jobs", [])
        old = next((item for item in jobs if item.get("id") == job_id), None)
        if old:
            for key in ("created_at", "last_run_at", "last_started_at", "last_status", "last_error", "last_file", "last_count"):
                if key in old:
                    job[key] = old[key]
            jobs = [job if item.get("id") == job_id else item for item in jobs]
        else:
            job["created_at"] = now_ts
            jobs.append(job)
        data["jobs"] = jobs
        _save_schedules(data)
    _SCHEDULE_WAKE.set()
    return _schedule_view(job)


def delete_schedule(job_id):
    with _SCHEDULE_LOCK:
        data = _load_schedules()
        before = len(data.get("jobs", []))
        data["jobs"] = [job for job in data.get("jobs", []) if job.get("id") != job_id]
        _save_schedules(data)
    _SCHEDULE_WAKE.set()
    return {"deleted": before != len(data.get("jobs", []))}


def _update_schedule_fields(job_id, fields):
    with _SCHEDULE_LOCK:
        data = _load_schedules()
        for job in data.get("jobs", []):
            if job.get("id") == job_id:
                job.update(fields)
                _save_schedules(data)
                return dict(job)
    return None


def _run_scheduled_export(job_id):
    try:
        with _SCHEDULE_LOCK:
            job = next((dict(item) for item in _load_schedules().get("jobs", []) if item.get("id") == job_id), None)
        if not job or not job.get("active"):
            return

        _update_schedule_fields(
            job_id,
            {
                "last_started_at": int(datetime.now().timestamp()),
                "last_status": "running",
                "last_error": "",
            },
        )
        export_func = export_chat_ai_package if job.get("format") == "ai_package" else export_chat_csv
        result = export_func(
            job.get("username", ""),
            range_mode=job.get("range_mode", "since_last"),
            custom_days=job.get("custom_days"),
            custom_start=job.get("custom_start", ""),
            custom_end=job.get("custom_end", ""),
            open_folder=False,
        )
        next_run = int(_compute_next_run(job).timestamp())
        _update_schedule_fields(
            job_id,
            {
                "last_run_at": int(datetime.now().timestamp()),
                "last_status": "ok",
                "last_error": "",
                "last_file": result.get("file", ""),
                "last_count": result.get("count", 0),
                "next_run": next_run,
            },
        )
    except Exception as exc:
        with _SCHEDULE_LOCK:
            job = next((dict(item) for item in _load_schedules().get("jobs", []) if item.get("id") == job_id), None)
        next_run = int(_compute_next_run(job or {"frequency": "daily", "time": "08:00"}).timestamp())
        _update_schedule_fields(
            job_id,
            {
                "last_run_at": int(datetime.now().timestamp()),
                "last_status": "error",
                "last_error": str(exc),
                "next_run": next_run,
            },
        )
    finally:
        with _SCHEDULE_LOCK:
            _SCHEDULE_RUNNING.discard(job_id)


def _trigger_due_schedules():
    now_ts = int(datetime.now().timestamp())
    due = []
    changed = False
    with _SCHEDULE_LOCK:
        data = _load_schedules()
        for job in data.get("jobs", []):
            if not job.get("active"):
                continue
            if not job.get("next_run"):
                job["next_run"] = int(_compute_next_run(job).timestamp())
                changed = True
            if job.get("next_run") <= now_ts and job.get("id") not in _SCHEDULE_RUNNING:
                _SCHEDULE_RUNNING.add(job.get("id"))
                due.append(dict(job))
        if changed:
            _save_schedules(data)

    for job in due:
        threading.Thread(target=_run_scheduled_export, args=(job["id"],), daemon=True).start()


def _scheduler_loop():
    while True:
        try:
            _trigger_due_schedules()
        except Exception as exc:
            print(f"Schedule check failed: {exc}", file=sys.stderr)
        _SCHEDULE_WAKE.wait(SCHEDULE_CHECK_SECONDS)
        _SCHEDULE_WAKE.clear()


def _start_scheduler():
    global _SCHEDULER_THREAD
    if _SCHEDULER_THREAD and _SCHEDULER_THREAD.is_alive():
        return
    _SCHEDULER_THREAD = threading.Thread(target=_scheduler_loop, name="wechat-export-scheduler", daemon=True)
    _SCHEDULER_THREAD.start()


def _strip_embedded_sender(text):
    if not text:
        return ""
    text = str(text)
    if ":\n" in text:
        possible_sender, rest = text.split(":\n", 1)
        if possible_sender.startswith("wxid_") or possible_sender.endswith("@chatroom"):
            text = rest
    text = re.sub(r"^\s*(wxid_[A-Za-z0-9_@.-]+|[0-9]{6,}@chatroom):\s*", "", text)
    return text.replace("\r\n", "\n").replace("\r", "\n").strip()


def _clean_content(row, ctx):
    local_id, local_type, _create_time, _real_sender_id, content, ct = row
    content = mcp_server._decompress_content(content, ct)
    if content is None:
        return ""

    base_type, _ = mcp_server._split_msg_type(local_type)
    if base_type in SYSTEM_TYPES:
        return ""

    if base_type == 1:
        _sender, text = mcp_server._parse_message_content(content, local_type, ctx["is_group"])
        return _strip_embedded_sender(text)
    if base_type == 3:
        return "[image]"
    if base_type == 34:
        return "[voice]"
    if base_type == 43:
        return "[video]"
    if base_type == 47:
        return "[sticker]"
    if base_type == 49:
        parsed = _format_link_or_file(content)
        if parsed:
            return parsed

    rendered = _extract_content(local_id, local_type, content, ct, ctx["username"], ctx["display_name"])
    rendered = _strip_embedded_sender(rendered or "")
    return re.sub(r"\s*\(local_id=\d+\)", "", rendered)


def _format_link_or_file(content):
    root = mcp_server._parse_xml_root(content) if content else None
    if root is None:
        return ""
    appmsg = root.find(".//appmsg")
    if appmsg is None:
        return ""
    title = mcp_server._collapse_text(appmsg.findtext("title") or "")
    url = (appmsg.findtext("url") or "").strip()
    app_type = mcp_server._parse_int((appmsg.findtext("type") or "").strip(), 0)
    if app_type == 6:
        return f"[file] {title}" if title else "[file]"
    if url:
        return f"[link] {title}\n{url}" if title else f"[link] {url}"
    return f"[link/file] {title}" if title else "[link/file]"


def _preview_kind(content):
    if not content:
        return "text"
    if content.startswith("[file]"):
        return "file"
    if content.startswith("[link]") or URL_RE.search(content):
        return "link"
    return "text"


def _preview_url(content):
    match = URL_RE.search(content or "")
    return match.group(0) if match else ""


def _markdown_text(value):
    value = str(value or "").replace("\r\n", "\n").replace("\r", "\n").strip()
    return value or "(empty)"


def _markdown_alt(value):
    return str(value or "image").replace("[", "(").replace("]", ")").replace("\n", " ")


def _copy_export_media(source_path, media_dir, timestamp, local_id, md5_value=""):
    source = Path(source_path)
    suffix = source.suffix.lower() or ".bin"
    time_part = datetime.fromtimestamp(timestamp).strftime("%Y%m%d_%H%M%S") if timestamp else "unknown_time"
    md5_part = _safe_media_name(md5_value)[:12]
    pieces = [time_part, str(local_id)]
    if md5_part:
        pieces.append(md5_part)
    filename = "_".join(pieces) + suffix
    target = media_dir / filename
    counter = 2
    while target.exists():
        target = media_dir / f"{target.stem}_{counter}{suffix}"
        counter += 1
    shutil.copy2(source, target)
    return target


def _write_zip(src_dir, zip_path):
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in src_dir.rglob("*"):
            if path.is_file():
                zf.write(path, path.relative_to(src_dir.parent))


def _is_port_in_use(host, port):
    with closing(socket.socket(socket.AF_INET, socket.SOCK_STREAM)) as sock:
        sock.settimeout(0.2)
        return sock.connect_ex((host, port)) == 0


def _find_available_port(host, start_port=8765):
    port = start_port
    while _is_port_in_use(host, port):
        port += 1
    return port


def _session_summary(summary):
    if isinstance(summary, bytes):
        try:
            summary = mcp_server._zstd_dctx.decompress(summary).decode("utf-8", errors="replace")
        except Exception:
            return ""
    if isinstance(summary, str) and ":\n" in summary:
        summary = summary.split(":\n", 1)[1]
    return _strip_embedded_sender(summary or "")


def _session_rows():
    path = mcp_server._cache.get(os.path.join("session", "session.db"))
    if not path:
        return []
    with closing(sqlite3.connect(path)) as conn:
        return conn.execute(
            """
            SELECT username, unread_count, summary, last_timestamp, sort_timestamp, last_sender_display_name
            FROM SessionTable
            WHERE username != '' AND last_timestamp > 0
            ORDER BY sort_timestamp DESC
            """
        ).fetchall()


_CONTACT_META = None


def _load_contact_meta():
    global _CONTACT_META
    if _CONTACT_META is not None:
        return _CONTACT_META

    meta = {}
    db_path = mcp_server._get_contact_db_path()
    if db_path:
        with closing(sqlite3.connect(db_path)) as conn:
            for row in conn.execute(
                """
                SELECT username, local_type, verify_flag, remark, nick_name
                FROM contact
                """
            ).fetchall():
                username, local_type, verify_flag, remark, nick_name = row
                meta[username] = {
                    "local_type": local_type or 0,
                    "verify_flag": verify_flag or 0,
                    "remark": remark or "",
                    "nick_name": nick_name or "",
                }
    _CONTACT_META = meta
    return meta


def _chat_type(username, contact_meta=None):
    if username.endswith("@chatroom"):
        return "group"
    if username.startswith("gh_"):
        return "official"
    if username in {
        "brandsessionholder",
        "brandservicesessionholder",
        "notifymessage",
        "qqmail",
        "weixin",
        "newsapp",
        "floatbottle",
    }:
        return "official"
    if username in {"filehelper", "medianote", "fmessage"}:
        return "friend"
    if username.startswith("@placeholder") or username.startswith("placeholder_"):
        return "friend"
    if username.endswith("@openim"):
        return "friend"

    contact_meta = contact_meta or {}
    meta = contact_meta.get(username)
    if meta:
        if meta.get("verify_flag", 0):
            return "official"
        return "friend"
    return "friend"


def _chat_type_label(chat_type):
    return {
        "group": "群聊",
        "friend": "好友",
        "official": "公众号",
        "other": "其他",
    }.get(chat_type, "其他")


def list_chats():
    names = mcp_server.get_contact_names()
    contact_meta = _load_contact_meta()
    state = _load_state()
    chats = []
    seen = set()
    for username, unread, summary, last_ts, _sort_ts, last_sender in _session_rows():
        seen.add(username)
        chat_type = _chat_type(username, contact_meta)
        chats.append(
            {
                "username": username,
                "display_name": names.get(username, username),
                "type": chat_type,
                "type_label": _chat_type_label(chat_type),
                "last_time": _format_time(last_ts, minutes=True),
                "last_timestamp": last_ts or 0,
                "last_sender": last_sender or "",
                "last_summary": _session_summary(summary),
                "unread": unread or 0,
                "last_export_time": _format_time(state.get(username, {}).get("last_timestamp"), minutes=True),
                "last_export_file": state.get(username, {}).get("file", ""),
            }
        )

    for username, display_name in names.items():
        if username.endswith("@chatroom") and username not in seen:
            chats.append(
                {
                    "username": username,
                    "display_name": display_name,
                    "type": "group",
                    "type_label": "group",
                    "last_time": "",
                    "last_timestamp": 0,
                    "last_sender": "",
                    "last_summary": "",
                    "unread": 0,
                    "last_export_time": _format_time(state.get(username, {}).get("last_timestamp"), minutes=True),
                    "last_export_file": state.get(username, {}).get("file", ""),
                }
            )
    return chats


def _collect_rows(ctx, start_ts=None, end_ts=None, limit=None, oldest_first=True):
    rows_with_maps = []
    names = mcp_server.get_contact_names()
    for table_info in ctx["message_tables"]:
        with closing(sqlite3.connect(table_info["db_path"])) as conn:
            id_to_username = mcp_server._load_name2id_maps(conn)
            rows = mcp_server._query_messages(
                conn,
                table_info["table_name"],
                start_ts=start_ts,
                end_ts=end_ts,
                limit=limit,
                oldest_first=oldest_first,
            )
            for row in rows:
                rows_with_maps.append((row, id_to_username))
    rows_with_maps.sort(key=lambda pair: pair[0][2] or 0, reverse=not oldest_first)
    return rows_with_maps, names


def preview_chat(username, limit=40, before_ts=None):
    ctx = mcp_server._resolve_chat_context(username)
    if not ctx or not ctx["message_tables"]:
        return {"messages": []}

    limit = max(1, int(limit))
    end_ts = int(before_ts) - 1 if before_ts else None
    rows_with_maps, names = _collect_rows(ctx, end_ts=end_ts, limit=limit, oldest_first=False)
    rows_with_maps = rows_with_maps[:limit]
    messages = []
    for row, id_to_username in reversed(rows_with_maps):
        base_type, _ = mcp_server._split_msg_type(row[1])
        if base_type in SYSTEM_TYPES:
            continue
        content = _clean_content(row, ctx)
        if not content:
            continue
        sender = _resolve_sender(row, ctx, names, id_to_username)
        messages.append(
            {
                "time": _format_time(row[2], minutes=True),
                "timestamp": row[2] or 0,
                "sender": sender,
                "content": content,
                "kind": _preview_kind(content),
                "url": _preview_url(content),
                "is_me": sender == "me",
            }
        )
    return {
        "chat": ctx["display_name"],
        "username": ctx["username"],
        "messages": messages,
        "has_more": len(rows_with_maps) >= limit,
        "oldest_timestamp": messages[0]["timestamp"] if messages else None,
    }


def export_chat_csv(username, range_mode="today", custom_days=None, custom_start="", custom_end="", open_folder=True):
    ctx = mcp_server._resolve_chat_context(username)
    if not ctx:
        raise ValueError(f"Chat not found: {username}")
    if not ctx["message_tables"]:
        raise ValueError(f"No message tables found for {ctx['display_name']}")

    state = _load_state()
    last_ts = state.get(ctx["username"], {}).get("last_timestamp")
    start_ts, end_ts = _range_to_timestamps(range_mode, custom_days, custom_start, custom_end, last_ts)
    if start_ts is not None and end_ts is not None and start_ts > end_ts:
        raise ValueError("&#24320;&#22987; time cannot be later than end time")

    rows_with_maps, names = _collect_rows(ctx, start_ts=start_ts, end_ts=end_ts, limit=None, oldest_first=True)
    out_rows = []
    skipped_system = 0
    max_ts = last_ts or 0
    for row, id_to_username in rows_with_maps:
        base_type, _ = mcp_server._split_msg_type(row[1])
        if base_type in SYSTEM_TYPES or _msg_type_str(row[1]) == "system":
            skipped_system += 1
            continue

        content = _clean_content(row, ctx)
        if not content:
            continue
        sender = _resolve_sender(row, ctx, names, id_to_username)
        ts = row[2] or 0
        max_ts = max(max_ts or 0, ts)
        out_rows.append({"sender": sender, "content": content})

    _ensure_export_dir()
    out_path = _export_path(ctx, range_mode, start_ts, end_ts, ".csv")
    with out_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["sender", "content"])
        writer.writeheader()
        writer.writerows(out_rows)

    if out_rows:
        state[ctx["username"]] = {
            "display_name": ctx["display_name"],
            "last_timestamp": max_ts,
            "last_time": _format_time(max_ts),
            "file": str(out_path),
            "exported_at": _format_time(int(datetime.now().timestamp())),
        }
        _save_state(state)
    if open_folder and out_path.exists():
        _open_export_target(out_path)

    return {
        "chat": ctx["display_name"],
        "username": ctx["username"],
        "file": str(out_path),
        "count": len(out_rows),
        "skipped_system": skipped_system,
        "start_time": _format_time(start_ts),
        "end_time": _format_time(end_ts),
        "last_export_time": _format_time(max_ts),
    }


def export_chat_ai_package(username, range_mode="today", custom_days=None, custom_start="", custom_end="", open_folder=True):
    ctx = mcp_server._resolve_chat_context(username)
    if not ctx:
        raise ValueError(f"Chat not found: {username}")
    if not ctx["message_tables"]:
        raise ValueError(f"No message tables found for {ctx['display_name']}")

    state = _load_state()
    last_ts = state.get(ctx["username"], {}).get("last_timestamp")
    start_ts, end_ts = _range_to_timestamps(range_mode, custom_days, custom_start, custom_end, last_ts)
    if start_ts is not None and end_ts is not None and start_ts > end_ts:
        raise ValueError("&#24320;&#22987; time cannot be later than end time")

    rows_with_maps, names = _collect_rows(ctx, start_ts=start_ts, end_ts=end_ts, limit=None, oldest_first=True)
    zip_path = _export_path(ctx, range_mode, start_ts, end_ts, "_压缩包.zip")
    package_dir = _unique_path(zip_path.with_suffix(""))
    media_dir = package_dir / "media"
    _ensure_export_dir()
    media_dir.mkdir(parents=True, exist_ok=True)

    messages = []
    markdown = [
        f"# {ctx['display_name']}",
        "",
        f"- username: `{ctx['username']}`",
        f"- exported_at: {_format_time(int(datetime.now().timestamp()))}",
        f"- range_start: {_format_time(start_ts) or 'beginning'}",
        f"- range_end: {_format_time(end_ts) or 'latest'}",
        "",
    ]
    skipped_system = 0
    image_count = 0
    image_failed = 0
    max_ts = last_ts or 0

    for row, id_to_username in rows_with_maps:
        local_id, local_type, create_time, _real_sender_id, _content, _ct = row
        base_type, _ = mcp_server._split_msg_type(local_type)
        type_str = _msg_type_str(local_type)
        if base_type in SYSTEM_TYPES or type_str == "system":
            skipped_system += 1
            continue

        ts = create_time or 0
        max_ts = max(max_ts or 0, ts)
        sender = _resolve_sender(row, ctx, names, id_to_username)
        msg = {
            "local_id": local_id,
            "sender": sender,
            "type": type_str,
        }

        markdown.append(f"## {sender or 'message'}")
        if base_type == 3:
            result = mcp_server._image_resolver.decode_image(ctx["username"], local_id)
            msg["content"] = "[image]"
            if result.get("success"):
                copied = _copy_export_media(result["path"], media_dir, ts, local_id, result.get("md5", ""))
                rel_path = copied.relative_to(package_dir).as_posix()
                msg["image_path"] = rel_path
                msg["image_format"] = result.get("format", "")
                msg["image_md5"] = result.get("md5", "")
                msg["image_size"] = result.get("size", 0)
                markdown.append(f"![{_markdown_alt(sender)}]({rel_path})")
                image_count += 1
            else:
                msg["image_error"] = result.get("error", "decode failed")
                if result.get("md5"):
                    msg["image_md5"] = result["md5"]
                markdown.append(f"[image decode failed: {_markdown_text(msg['image_error'])}]")
                image_failed += 1
        else:
            content = _clean_content(row, ctx)
            if not content:
                continue
            msg["content"] = content
            markdown.append(_markdown_text(content))

        markdown.append("")
        messages.append(msg)

    payload = {
        "chat": ctx["display_name"],
        "username": ctx["username"],
        "exported_at": _format_time(int(datetime.now().timestamp())),
        "range_start": _format_time(start_ts),
        "range_end": _format_time(end_ts),
        "message_count": len(messages),
        "image_count": image_count,
        "image_failed": image_failed,
        "messages": messages,
    }
    if ctx["is_group"]:
        payload["is_group"] = True

    (package_dir / "chat.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    (package_dir / "chat.md").write_text("\n".join(markdown).rstrip() + "\n", encoding="utf-8")
    _write_zip(package_dir, zip_path)

    if messages:
        state[ctx["username"]] = {
            "display_name": ctx["display_name"],
            "last_timestamp": max_ts,
            "last_time": _format_time(max_ts),
            "file": str(zip_path),
            "exported_at": _format_time(int(datetime.now().timestamp())),
        }
        _save_state(state)
    if open_folder and zip_path.exists():
        _open_export_target(zip_path)

    return {
        "chat": ctx["display_name"],
        "username": ctx["username"],
        "file": str(zip_path),
        "folder": str(package_dir),
        "count": len(messages),
        "image_count": image_count,
        "image_failed": image_failed,
        "skipped_system": skipped_system,
        "start_time": _format_time(start_ts),
        "end_time": _format_time(end_ts),
        "last_export_time": _format_time(max_ts),
    }


def _format_size(size):
    try:
        size = int(size or 0)
    except (TypeError, ValueError):
        size = 0
    units = ["B", "KB", "MB", "GB"]
    value = float(size)
    for unit in units:
        if value < 1024 or unit == units[-1]:
            return f"{value:.1f} {unit}" if unit != "B" else f"{int(value)} B"
        value /= 1024


def _artifact_candidates():
    if not EXPORT_DIR.exists():
        return []
    candidates = []

    def add_artifact(path):
        if path.name in {STATE_FILE.name, SCHEDULE_FILE.name}:
            return
        if path.is_file() and path.suffix.lower() not in {".csv", ".zip"}:
            return
        if path.is_dir() and not (path / "chat.json").exists():
            return
        if path.is_dir() and path.with_suffix(".zip").exists():
            return
        candidates.append(path)

    for item in EXPORT_DIR.iterdir():
        if item.is_file():
            add_artifact(item)
            continue
        if not item.is_dir():
            continue
        if (item / "chat.json").exists():
            add_artifact(item)
            continue
        for child in item.iterdir():
            if child.is_file() or child.is_dir():
                add_artifact(child)
    return candidates


def _artifact_meta_from_folder(path):
    chat_json = path / "chat.json"
    if not chat_json.exists():
        return {}
    try:
        data = json.loads(chat_json.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return {
        "display_name": data.get("chat") or "",
        "username": data.get("username") or "",
        "range_label": _range_display("custom", _parse_export_time(data.get("range_start")), _parse_export_time(data.get("range_end"))),
        "message_count": data.get("message_count"),
    }


def _parse_export_time(value):
    if not value:
        return None
    try:
        return int(datetime.strptime(str(value)[:19], "%Y-%m-%d %H:%M:%S").timestamp())
    except (TypeError, ValueError):
        return None


def _guess_display_name(path):
    stem = path.name
    if path.suffix:
        stem = path.stem
    stem = re.sub(r"_压缩包$", "", stem)
    match = re.match(r"(.+?)_\d+_chatroom_\d{8}", stem)
    if match:
        return match.group(1)
    parts = stem.split("_")
    return parts[0] if parts else stem


def _guess_range_from_filename(path):
    match = re.search(r"_(20\d{6})(?:_\d{6})?", path.stem if path.suffix else path.name)
    if not match:
        return ""
    try:
        return f"{datetime.strptime(match.group(1), '%Y%m%d').strftime('%Y-%m-%d')} 导出"
    except ValueError:
        return ""


def _export_record_view(path, state_item=None):
    target = Path(path)
    state_item = state_item or {}
    folder_meta = _artifact_meta_from_folder(target) if target.is_dir() else {}
    exists = target.exists()
    stat = target.stat() if exists else None
    display_name = state_item.get("display_name") or folder_meta.get("display_name") or _guess_display_name(target)
    username = state_item.get("username") or folder_meta.get("username") or ""
    state_last_ts = _parse_export_time(state_item.get("last_time"))
    state_range = f"截至 {_format_date(state_last_ts)}" if state_last_ts else ""
    range_label = folder_meta.get("range_label") or state_range or _guess_range_from_filename(target)
    exported_at = state_item.get("exported_at") or (_format_time(int(stat.st_mtime)) if stat else "")
    suffix = target.suffix.lower()
    if target.is_dir():
        kind = "压缩包文件夹"
    elif suffix == ".zip":
        kind = "压缩包"
    elif suffix == ".csv":
        kind = "CSV"
    else:
        kind = suffix.lstrip(".").upper() or "文件"
    size = 0
    size_label = ""
    if stat:
        if target.is_dir():
            size_label = "文件夹"
        else:
            size = stat.st_size
            size_label = _format_size(size)
    return {
        "path": str(target),
        "username": username,
        "display_name": display_name,
        "title": f"{display_name} · {range_label}" if range_label else display_name,
        "range_label": range_label,
        "exported_at": exported_at,
        "exists": exists,
        "kind": kind,
        "size": size,
        "size_label": size_label,
        "tracked": bool(state_item),
        "message_count": folder_meta.get("message_count"),
    }


def list_export_records():
    state = _load_state()
    by_path = {}
    for username, item in state.items():
        file_path = item.get("file")
        if not file_path:
            continue
        enriched = dict(item)
        enriched["username"] = username
        by_path[str(Path(file_path))] = _export_record_view(file_path, enriched)
    for path in _artifact_candidates():
        key = str(path)
        if key not in by_path:
            by_path[key] = _export_record_view(path)
    records = list(by_path.values())
    records.sort(key=lambda item: item.get("exported_at") or "", reverse=True)
    return records


def delete_export_record(payload):
    path = _safe_path(payload.get("path", ""))
    username = str(payload.get("username") or "")
    delete_file = bool(payload.get("delete_file"))

    state = _load_state()
    changed = False
    for key, item in list(state.items()):
        if (username and key == username) or (path and item.get("file") and Path(item["file"]) == path):
            state.pop(key, None)
            changed = True
    if changed:
        _save_state(state)

    deleted_file = False
    if delete_file and path and path.exists():
        if path.is_dir():
            shutil.rmtree(path)
            deleted_file = True
        elif path.is_file():
            path.unlink()
            deleted_file = True
            if path.suffix.lower() == ".zip":
                folder = path.with_suffix("")
                safe_folder = _safe_path(folder)
                if safe_folder and safe_folder.exists() and safe_folder.is_dir():
                    shutil.rmtree(safe_folder)
    return {"deleted_record": changed, "deleted_file": deleted_file, "records": list_export_records()}


def open_export_record(path):
    target = _safe_path(path)
    if not target or not target.exists():
        raise ValueError("Export file not found")
    _open_export_target(target)
    return {"ok": True}


def open_export_folder(path):
    target = _safe_path(path)
    if not target or not target.exists():
        raise ValueError("Export file not found")
    folder = target if target.is_dir() else target.parent
    _open_folder_target(folder)
    return {"ok": True}


def _bring_path_window_to_front(path):
    if os.name != "nt":
        return
    try:
        import ctypes
        import time

        time.sleep(0.8)
        user32 = ctypes.windll.user32
        hints = [Path(path).name.lower(), Path(path).parent.name.lower()]
        explorer_hints = ("文件资源管理器", "file explorer", "explorer")

        def callback(hwnd, _):
            if not user32.IsWindowVisible(hwnd):
                return True
            length = user32.GetWindowTextLengthW(hwnd)
            if length <= 0:
                return True
            buffer = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buffer, length + 1)
            title = buffer.value.lower()
            if any(hint and hint in title for hint in hints) or any(hint in title for hint in explorer_hints):
                user32.ShowWindow(hwnd, 9)
                user32.SetForegroundWindow(hwnd)
                return False
            return True

        user32.EnumWindows(ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)(callback), 0)
    except Exception:
        pass


def _open_export_target(path):
    target = Path(path)
    try:
        if os.name == "nt":
            if target.is_file():
                subprocess.Popen(["explorer", "/select,", str(target)])
            else:
                os.startfile(str(target))
            threading.Thread(target=_bring_path_window_to_front, args=(target,), daemon=True).start()
        else:
            folder = target if target.is_dir() else target.parent
            subprocess.Popen(["xdg-open", str(folder)])
    except Exception:
        pass


def _open_folder_target(path):
    folder = Path(path)
    try:
        if os.name == "nt":
            os.startfile(str(folder))
            threading.Thread(target=_bring_path_window_to_front, args=(folder,), daemon=True).start()
        else:
            subprocess.Popen(["xdg-open", str(folder)])
    except Exception:
        pass


INDEX_HTML = r"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>WeChat &#32842;&#22825;&#23548;&#20986;</title>
  <script>
    (function () {
      try {
        const versionKey = "wechat-export-theme-version";
        const saved = localStorage.getItem("wechat-export-theme");
        let next = saved === "dark" ? "dark" : "light";
        if (localStorage.getItem(versionKey) !== "2") {
          next = "light";
          localStorage.setItem("wechat-export-theme", "light");
          localStorage.setItem(versionKey, "2");
        }
        document.documentElement.dataset.theme = next;
      } catch (err) {
        document.documentElement.dataset.theme = "light";
      }
    })();
  </script>
  <style>
    :root {
      color-scheme: light;
      --bg: #f4f5f1;
      --pane: #fbfbf7;
      --ink: #1f2420;
      --muted: #697168;
      --faint: #8d958b;
      --line: #dfe4da;
      --accent: #24745c;
      --accent-soft: #e3eee8;
      --accent-ink: #134734;
      --accent-border: #b6c5bb;
      --canvas: #ffffff;
      --control: #ffffff;
      --control-soft: #f0f2ed;
      --control-hover: #f7faf6;
      --active: #dcebe5;
      --preview: #f1f2ed;
      --preview-head: #f9faf6;
      --detail: #fbfcf8;
      --bubble: #ffffff;
      --me: #9eea6a;
      --me-ink: #142414;
      --avatar: #d9dfd3;
      --avatar-ink: #596156;
      --avatar-me: #cde9ba;
      --time-bg: #e0e3dd;
      --time-ink: #8a9087;
      --link: #457461;
      --friend-ink: #5b4a19;
      --friend-bg: #eee8d2;
      --official-ink: #244f7a;
      --official-bg: #dfeaf3;
      --other-ink: #666b68;
      --other-bg: #e9ece6;
      --danger: #a64236;
      --shadow: 0 14px 38px rgba(36, 48, 38, .08);
    }
    :root[data-theme="dark"] {
      color-scheme: dark;
      --bg: #111611;
      --pane: #171d18;
      --ink: #edf2e9;
      --muted: #a7b1a5;
      --faint: #7d897b;
      --line: #2c382f;
      --accent: #7dd6ad;
      --accent-soft: #21382d;
      --accent-ink: #b8f1d5;
      --accent-border: #4d6a59;
      --canvas: #1b231d;
      --control: #202822;
      --control-soft: #151b16;
      --control-hover: #263329;
      --active: #254233;
      --preview: #151b16;
      --preview-head: #1d251f;
      --detail: #192119;
      --bubble: #232d25;
      --me: #7fd38c;
      --me-ink: #0c1d10;
      --avatar: #2c382f;
      --avatar-ink: #c6d0c3;
      --avatar-me: #35543b;
      --time-bg: #263027;
      --time-ink: #a7b1a5;
      --link: #8bd4b4;
      --friend-ink: #f0dc9a;
      --friend-bg: #3b321e;
      --official-ink: #b7d7f2;
      --official-bg: #203447;
      --other-ink: #c1cbc0;
      --other-bg: #2a322b;
      --danger: #ff9d8e;
      --shadow: 0 18px 48px rgba(0, 0, 0, .26);
    }
    * { box-sizing: border-box; }
    html {
      height: 100%;
      background: var(--bg);
    }
    body {
      margin: 0;
      height: 100%;
      overflow: hidden;
      background: var(--bg);
      color: var(--ink);
      font-family: Inter, "Segoe UI", "Microsoft YaHei", sans-serif;
      letter-spacing: 0;
    }
    main {
      height: 100vh;
      display: grid;
      grid-template-columns: minmax(280px, var(--sidebar-width, 390px)) 6px minmax(0, 1fr) var(--manager-width, 360px);
      overflow: hidden;
    }
    body.manager-collapsed {
      --manager-width: 48px;
    }
    .chat-sidebar {
      border-right: 1px solid var(--line);
      background: var(--pane);
      padding: 22px 18px;
      min-width: 0;
      min-height: 0;
      display: flex;
      flex-direction: column;
    }
    .sidebar-resizer {
      width: 6px;
      cursor: col-resize;
      background: linear-gradient(to right, transparent, var(--line), transparent);
      touch-action: none;
    }
    .sidebar-resizer:hover,
    .sidebar-resizer.dragging {
      background: var(--accent-soft);
    }
    .chat-sidebar, section, .manager-panel, .preview, .preview-head, .export-panel, .range-detail, .tab, .group, .range-option, input, button, .menu {
      transition: background-color .2s ease, border-color .2s ease, color .2s ease, box-shadow .2s ease;
    }
    section {
      padding: 24px 30px;
      min-width: 0;
      min-height: 0;
      overflow: hidden;
      background: var(--bg);
    }
    h1 {
      margin: 0;
      font-size: 21px;
      font-weight: 760;
    }
    .aside-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 14px;
    }
    .head-actions {
      display: flex;
      align-items: center;
      gap: 9px;
    }
    .count {
      color: var(--faint);
      font-size: 12px;
      white-space: nowrap;
    }
    .theme-toggle {
      width: 34px;
      height: 34px;
      border: 1px solid var(--line);
      border-radius: 7px;
      background: var(--control);
      color: var(--ink);
      cursor: pointer;
      display: grid;
      place-items: center;
      position: relative;
      box-shadow: 0 1px 0 rgba(30, 38, 31, .04);
    }
    .theme-toggle:hover { border-color: var(--accent-border); background: var(--control-hover); transform: translateY(-1px); }
    .theme-toggle::before {
      content: "";
      width: 14px;
      height: 14px;
      border-radius: 50%;
      background: var(--accent);
      box-shadow: 0 0 0 4px color-mix(in srgb, var(--accent) 18%, transparent);
    }
    :root[data-theme="dark"] .theme-toggle::before {
      background: transparent;
      box-shadow: inset -5px 0 0 1px var(--accent);
    }
    .search {
      width: 100%;
      margin: 16px 0 10px;
      padding: 12px 13px;
      border: 1px solid var(--line);
      background: var(--control);
      color: var(--ink);
      border-radius: 6px;
      font-size: 14px;
      outline: none;
    }
    .search::placeholder { color: var(--faint); }
    .tabs {
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 4px;
      padding: 4px;
      margin-bottom: 12px;
      border: 1px solid var(--line);
      border-radius: 7px;
      background: var(--control-soft);
    }
    .tab {
      border: 0;
      border-radius: 5px;
      background: transparent;
      color: var(--muted);
      padding: 8px 6px;
      font-size: 12px;
      cursor: pointer;
      transition: background .16s ease, color .16s ease, box-shadow .16s ease;
    }
    .tab.active {
      background: var(--control);
      color: var(--ink);
      box-shadow: 0 1px 5px rgba(30, 38, 31, .08);
    }
    .list {
      display: grid;
      gap: 3px;
      flex: 1;
      min-height: 0;
      max-height: none;
      overflow: auto;
      padding-right: 4px;
    }
    .group {
      width: 100%;
      text-align: left;
      border: 0;
      border-radius: 6px;
      background: transparent;
      padding: 11px 10px;
      cursor: pointer;
      color: var(--ink);
      transition: background .16s ease, transform .16s ease;
    }
    .group-row {
      display: grid;
      grid-template-columns: minmax(0, 1fr) 34px;
      gap: 4px;
      align-items: stretch;
    }
    .group:hover { background: var(--accent-soft); transform: translateX(2px); }
    .group.active { background: var(--active); }
    .group.pinned {
      background: var(--accent-soft);
      box-shadow: inset 3px 0 0 var(--accent);
    }
    .pin-button,
    .icon-action {
      display: grid;
      place-items: center;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: var(--control);
      color: var(--muted);
      cursor: pointer;
      font-size: 16px;
      line-height: 1;
    }
    .pin-button {
      width: 34px;
      min-height: 100%;
    }
    .pin-button:hover,
    .icon-action:hover {
      border-color: var(--accent-border);
      background: var(--control-hover);
      color: var(--accent);
      transform: translateY(-1px);
    }
    .pin-button.active,
    .icon-action.active {
      border-color: var(--accent);
      background: var(--accent-soft);
      color: var(--accent);
    }
    .rowtop {
      display: flex;
      justify-content: space-between;
      gap: 10px;
      align-items: center;
      min-width: 0;
    }
    .rowtop b {
      display: block;
      flex: 1 1 auto;
      font-size: 14px;
      min-width: 0;
      overflow: hidden;
      white-space: nowrap;
      text-overflow: ellipsis;
    }
    .rowtop time {
      color: var(--muted);
      flex: 0 0 auto;
      font-size: 12px;
    }
    .summary {
      margin-top: 5px;
      color: var(--muted);
      font-size: 12px;
      overflow: hidden;
      white-space: nowrap;
      text-overflow: ellipsis;
    }
    .chat-line {
      display: flex;
      flex: 1 1 auto;
      min-width: 0;
      overflow: hidden;
      align-items: center;
      gap: 7px;
    }
    .badge {
      flex: 0 0 auto;
      border-radius: 4px;
      padding: 2px 5px;
      color: var(--accent-ink);
      background: var(--accent-soft);
      font-size: 11px;
      font-weight: 680;
    }
    .badge.type-friend {
      color: var(--friend-ink);
      background: var(--friend-bg);
    }
    .badge.type-official {
      color: var(--official-ink);
      background: var(--official-bg);
    }
    .badge.type-other {
      color: var(--other-ink);
      background: var(--other-bg);
    }
    .workspace {
      height: 100%;
      max-width: none;
      margin: 0;
      display: grid;
      grid-template-rows: auto minmax(0, 1fr) auto;
      gap: 16px;
      min-height: 0;
    }
    .topline {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 18px;
      border-bottom: 1px solid var(--line);
      padding-bottom: 18px;
    }
    .selected {
      font-size: 26px;
      line-height: 1.2;
      font-weight: 760;
      word-break: break-word;
    }
    .meta {
      color: var(--muted);
      font-size: 13px;
    }
    .action-bar {
      display: flex;
      align-items: center;
      justify-content: flex-end;
      gap: 10px;
      flex: 0 0 auto;
    }
    .menu-wrap {
      position: relative;
    }
    .action-button {
      min-width: 112px;
      width: auto !important;
    }
    .icon-action {
      width: 42px !important;
      min-width: 42px;
      height: 42px;
      padding: 0 !important;
      font-size: 18px;
    }
    .menu {
      position: absolute;
      top: calc(100% + 8px);
      right: 0;
      z-index: 20;
      width: 310px;
      display: grid;
      gap: 8px;
      padding: 10px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--pane);
      box-shadow: var(--shadow);
    }
    .menu[hidden] { display: none; }
    .export-menu { right: 90px; width: 340px; }
    .schedule-menu { width: min(520px, calc(100vw - 36px)); }
    .menu-section {
      display: grid;
      gap: 8px;
      padding: 6px;
    }
    .menu-label {
      color: var(--faint);
      font-size: 12px;
      font-weight: 680;
    }
    .format-toggle {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 4px;
      padding: 4px;
      border: 1px solid var(--line);
      border-radius: 7px;
      background: var(--control-soft);
    }
    .format-choice,
    .menu-item {
      border: 0;
      border-radius: 6px;
      background: transparent;
      color: var(--ink);
      cursor: pointer;
      text-align: left;
      font-size: 14px;
    }
    .format-choice {
      text-align: center;
      padding: 9px 8px;
      color: var(--muted);
    }
    .format-choice.active {
      background: var(--control);
      color: var(--ink);
      box-shadow: 0 1px 5px rgba(30, 38, 31, .08);
    }
    .menu-item {
      width: 100%;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 11px 12px;
    }
    .menu-copy {
      display: grid;
      gap: 3px;
      min-width: 0;
    }
    .menu-title {
      color: var(--ink);
      font-weight: 720;
      line-height: 1.2;
    }
    .menu-hint {
      color: var(--muted);
      font-size: 12px;
      line-height: 1.35;
    }
    .menu-chevron {
      color: var(--faint);
      font-size: 18px;
      line-height: 1;
      transition: transform .16s ease;
    }
    .menu-item.active .menu-chevron {
      transform: rotate(180deg);
    }
    .menu-item:hover,
    .menu-item:focus {
      outline: none;
      background: var(--accent-soft);
    }
    .has-submenu {
      position: relative;
    }
    .submenu {
      position: absolute;
      top: 0;
      left: calc(100% + 8px);
      width: 180px;
      display: none;
      gap: 4px;
      padding: 8px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--pane);
      box-shadow: var(--shadow);
    }
    .has-submenu:hover .submenu,
    .has-submenu:focus-within .submenu {
      display: grid;
    }
    .custom-export {
      display: grid;
      gap: 10px;
      margin: -2px 6px 6px;
      padding: 12px;
      border: 1px solid var(--line);
      border-radius: 7px;
      background: var(--detail);
    }
    .custom-export[hidden] { display: none; }
    .schedule-form {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px;
    }
    .schedule-custom {
      display: none;
      grid-column: 1 / -1;
      grid-template-columns: 1fr 1fr;
      gap: 10px;
    }
    .schedule-custom.show { display: grid; }
    .controls {
      display: grid;
      grid-template-columns: minmax(0, 1fr) minmax(220px, 280px);
      gap: 16px;
      align-items: stretch;
    }
    label {
      display: grid;
      gap: 7px;
      color: var(--muted);
      font-size: 13px;
    }
    input[type="number"], input[type="datetime-local"] {
      width: 100%;
      padding: 11px 12px;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: var(--control);
      font-size: 14px;
      color: var(--ink);
    }
    .range-board {
      display: grid;
      gap: 13px;
    }
    .range-title {
      color: var(--ink);
      font-size: 15px;
      font-weight: 720;
    }
    .range-groups {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 10px;
    }
    .range-group {
      display: grid;
      align-content: start;
      gap: 7px;
    }
    .range-kicker {
      color: var(--faint);
      font-size: 12px;
    }
    .range-option {
      border: 1px solid var(--line);
      background: var(--control);
      color: var(--ink);
      border-radius: 7px;
      padding: 10px 11px;
      font-size: 13px;
      text-align: left;
      cursor: pointer;
      transition: background .16s ease, border-color .16s ease, color .16s ease, transform .16s ease, box-shadow .16s ease;
    }
    .range-option:hover { border-color: var(--accent-border); background: var(--control-hover); transform: translateY(-1px); }
    .range-option.active {
      background: var(--accent-soft);
      border-color: var(--accent);
      box-shadow: inset 3px 0 0 var(--accent);
    }
    .range-name {
      display: block;
      font-weight: 720;
    }
    .range-desc {
      display: block;
      margin-top: 3px;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.35;
    }
    .range-detail {
      display: none;
      grid-template-columns: minmax(120px, 180px) minmax(0, 1fr);
      gap: 12px;
      align-items: end;
      padding: 12px;
      border: 1px solid var(--line);
      border-radius: 7px;
      background: var(--detail);
    }
    .range-detail.show { display: grid; }
    .custom-row {
      display: none;
      grid-template-columns: 1fr 1fr;
      gap: 12px;
    }
    .custom-row.show { display: grid; }
    .export-panel {
      display: grid;
      align-content: space-between;
      gap: 12px;
      padding: 14px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--detail);
      box-shadow: var(--shadow);
    }
    .range-summary {
      color: var(--muted);
      font-size: 13px;
      line-height: 1.5;
    }
    button.primary {
      border: 0;
      border-radius: 6px;
      background: var(--accent);
      color: white;
      padding: 12px 18px;
      font-size: 14px;
      font-weight: 680;
      cursor: pointer;
      transition: filter .16s ease, transform .16s ease;
      min-height: 42px;
      width: 100%;
    }
    button.primary:hover { filter: brightness(.95); transform: translateY(-1px); }
    button.primary:disabled { opacity: .45; cursor: not-allowed; transform: none; }
    button.secondary {
      border: 1px solid var(--accent);
      border-radius: 6px;
      background: var(--control);
      color: var(--accent);
      padding: 12px 18px;
      font-size: 14px;
      font-weight: 680;
      cursor: pointer;
      transition: background .16s ease, transform .16s ease;
      min-height: 42px;
      width: 100%;
    }
    button.secondary:hover { background: var(--accent-soft); transform: translateY(-1px); }
    button.secondary:disabled { opacity: .45; cursor: not-allowed; transform: none; }
    .export-actions {
      display: grid;
      gap: 8px;
    }
    .schedule-panel {
      display: grid;
      gap: 12px;
      padding: 14px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--detail);
    }
    .schedule-grid {
      display: grid;
      grid-template-columns: repeat(5, minmax(0, 1fr));
      gap: 10px;
      align-items: end;
    }
    .schedule-grid select,
    .schedule-grid input {
      width: 100%;
      padding: 10px 11px;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: var(--control);
      color: var(--ink);
      font-size: 14px;
    }
    .schedule-actions {
      display: grid;
      grid-template-columns: minmax(0, 1fr) minmax(130px, 160px);
      gap: 10px;
    }
    .schedule-list {
      display: grid;
      gap: 6px;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.45;
      max-height: 240px;
      overflow: auto;
    }
    .schedule-item {
      display: grid;
      grid-template-columns: minmax(0, 1fr) 34px;
      gap: 8px;
      align-items: stretch;
      padding-top: 8px;
      border-top: 1px solid var(--line);
    }
    .schedule-item b { color: var(--ink); }
    .schedule-pick,
    .schedule-delete {
      border: 0;
      border-radius: 6px;
      background: transparent;
      color: inherit;
      cursor: pointer;
    }
    .schedule-pick {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      min-width: 0;
      padding: 8px;
      text-align: left;
    }
    .schedule-pick:hover,
    .schedule-pick.active {
      background: var(--accent-soft);
    }
    .schedule-pick > span {
      min-width: 0;
    }
    .schedule-delete {
      display: grid;
      place-items: center;
      width: 34px;
      color: var(--danger);
      font-size: 18px;
    }
    .schedule-delete:hover {
      background: var(--control-hover);
    }
    .preview {
      height: 100%;
      min-height: 0;
      background: var(--preview);
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: hidden;
      display: grid;
      grid-template-rows: auto 1fr;
    }
    .empty {
      color: var(--muted);
      padding: 18px 10px;
      font-size: 13px;
    }
    .preview-head {
      padding: 12px 16px;
      border-bottom: 1px solid var(--line);
      background: var(--preview-head);
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: center;
    }
    .messages {
      padding: 18px 18px 24px;
      overflow: auto;
      display: flex;
      flex-direction: column;
      gap: 12px;
    }
    .msg {
      display: grid;
      grid-template-columns: 34px minmax(0, 1fr);
      gap: 9px;
      max-width: 76%;
      align-self: flex-start;
    }
    .msg.me {
      grid-template-columns: minmax(0, 1fr) 34px;
      align-self: flex-end;
    }
    .avatar {
      width: 34px;
      height: 34px;
      border-radius: 5px;
      background: var(--avatar);
      display: grid;
      place-items: center;
      color: var(--avatar-ink);
      font-size: 13px;
      font-weight: 700;
    }
    .msg.me .avatar { grid-column: 2; background: var(--avatar-me); }
    .body { min-width: 0; }
    .name {
      color: var(--muted);
      font-size: 12px;
      margin-bottom: 4px;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .bubble {
      display: inline-block;
      background: var(--bubble);
      border-radius: 6px;
      padding: 9px 11px;
      line-height: 1.55;
      font-size: 14px;
      white-space: pre-wrap;
      word-break: break-word;
      box-shadow: 0 1px 0 rgba(0,0,0,.03);
    }
    .link-card, .file-card {
      display: grid;
      gap: 6px;
      min-width: 230px;
      max-width: 360px;
      text-decoration: none;
      color: inherit;
    }
    .card-kicker {
      color: var(--muted);
      font-size: 12px;
    }
    .card-title {
      font-weight: 650;
      line-height: 1.45;
    }
    .card-url {
      color: var(--link);
      font-size: 12px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .msg.me .body { text-align: right; grid-column: 1; grid-row: 1; }
    .msg.me .bubble { background: var(--me); color: var(--me-ink); text-align: left; }
    .time-sep {
      align-self: center;
      color: var(--time-ink);
      background: var(--time-bg);
      border-radius: 4px;
      padding: 3px 7px;
      font-size: 12px;
    }
    .status {
      min-height: 0;
      max-height: 70px;
      overflow: auto;
      color: var(--muted);
      line-height: 1.6;
      white-space: pre-wrap;
      word-break: break-word;
    }
    .ok { color: var(--accent); }
    .err { color: var(--danger); }
    .manager-panel {
      min-width: 0;
      min-height: 0;
      border-left: 1px solid var(--line);
      background: var(--pane);
      display: grid;
      grid-template-columns: 48px minmax(0, 1fr);
      overflow: hidden;
    }
    .manager-rail {
      border-right: 1px solid var(--line);
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 12px;
      padding: 14px 7px;
      color: var(--muted);
      background: var(--control-soft);
    }
    .manager-toggle {
      width: 34px;
      height: 34px;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: var(--control);
      color: var(--ink);
      cursor: pointer;
      font-size: 18px;
      line-height: 1;
    }
    .manager-toggle:hover {
      border-color: var(--accent-border);
      background: var(--control-hover);
    }
    .manager-rail span {
      writing-mode: vertical-rl;
      letter-spacing: 0;
      font-size: 12px;
      font-weight: 700;
    }
    .manager-content {
      min-width: 0;
      min-height: 0;
      display: grid;
      grid-template-rows: auto auto minmax(0, 1fr);
      gap: 14px;
      padding: 20px 16px 18px;
      overflow: hidden;
    }
    body.manager-collapsed .manager-panel {
      grid-template-columns: 48px 0;
    }
    body.manager-collapsed .manager-content {
      visibility: hidden;
      padding-left: 0;
      padding-right: 0;
    }
    body.manager-collapsed .manager-toggle {
      transform: rotate(180deg);
    }
    .manager-head {
      display: flex;
      align-items: start;
      justify-content: space-between;
      gap: 12px;
      border-bottom: 1px solid var(--line);
      padding-bottom: 13px;
    }
    .manager-title {
      font-size: 17px;
      font-weight: 760;
    }
    .manager-actions {
      display: flex;
      gap: 8px;
      flex: 0 0 auto;
    }
    .mini-button {
      border: 1px solid var(--line);
      border-radius: 6px;
      background: var(--control);
      color: var(--ink);
      cursor: pointer;
      min-height: 32px;
      padding: 6px 9px;
      font-size: 12px;
      white-space: nowrap;
    }
    .mini-button:hover {
      border-color: var(--accent-border);
      background: var(--control-hover);
    }
    .mini-button.danger {
      color: var(--danger);
    }
    .manager-stats {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px;
    }
    .manager-stat {
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 8px;
      text-align: left;
      cursor: pointer;
      color: var(--ink);
      background: transparent;
      font: inherit;
    }
    .manager-stat.active {
      border-color: var(--accent-border);
      background: var(--accent-soft);
      box-shadow: inset 0 0 0 1px var(--accent-border);
    }
    .manager-stat b {
      display: block;
      font-size: 18px;
      line-height: 1.2;
    }
    .manager-stat span {
      color: var(--muted);
      font-size: 12px;
    }
    .manager-section {
      min-height: 0;
      display: grid;
      grid-template-rows: auto minmax(0, 1fr);
      gap: 8px;
    }
    .manager-section-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 10px;
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
    }
    .manager-list {
      min-height: 0;
      overflow: auto;
      display: grid;
      align-content: start;
      gap: 7px;
      padding-right: 3px;
    }
    .manager-list[hidden] {
      display: none;
    }
    .managed-item {
      border-top: 1px solid var(--line);
      padding: 9px 0 8px;
      display: grid;
      gap: 7px;
    }
    .managed-item.active {
      box-shadow: inset 3px 0 0 var(--accent);
      padding-left: 8px;
    }
    .managed-title {
      color: var(--ink);
      font-size: 13px;
      line-height: 1.35;
      font-weight: 720;
      word-break: break-word;
    }
    .managed-meta {
      color: var(--muted);
      font-size: 12px;
      line-height: 1.45;
      word-break: break-word;
      overflow-wrap: anywhere;
    }
    .managed-actions {
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
    }
    @media (max-width: 1320px) {
      .topline { display: block; }
      .action-bar { justify-content: flex-start; margin-top: 14px; flex-wrap: wrap; }
    }
    @media (max-width: 860px) {
      body { overflow: auto; }
      main { grid-template-columns: 1fr; height: auto; min-height: 100vh; overflow: visible; }
      .chat-sidebar { border-right: 0; border-bottom: 1px solid var(--line); }
      .sidebar-resizer { display: none; }
      .list { max-height: 280px; }
      section { padding: 20px; overflow: visible; }
      .manager-panel { min-height: 520px; border-left: 0; border-top: 1px solid var(--line); }
      body.manager-collapsed .manager-panel { min-height: 48px; }
      .workspace { height: auto; min-height: 560px; }
      .action-bar { justify-content: flex-start; margin-top: 14px; }
      .menu { left: 0; right: auto; }
      .submenu { position: static; width: 100%; margin-top: 6px; box-shadow: none; }
      .controls { grid-template-columns: 1fr; }
      .range-groups { grid-template-columns: 1fr; }
      .custom-row { grid-template-columns: 1fr; }
      .range-detail { grid-template-columns: 1fr; }
      .schedule-grid { grid-template-columns: 1fr 1fr; }
      .schedule-form { grid-template-columns: 1fr; }
      .schedule-custom { grid-template-columns: 1fr; }
      .schedule-actions { grid-template-columns: 1fr; }
      .msg { max-width: 92%; }
      .topline { display: block; }
    }

    /* Reference-inspired UI refresh: semantic tokens, soft surfaces, restrained motion. */
    :root {
      --bg: #f5f7f4;
      --bg-deep: #e8eee8;
      --pane: rgba(255, 255, 255, .76);
      --pane-solid: #ffffff;
      --ink: #12201d;
      --muted: #41504b;
      --faint: #6e7a75;
      --line: rgba(18, 32, 29, .10);
      --line-strong: rgba(18, 32, 29, .20);
      --accent: #0f766e;
      --accent-soft: rgba(15, 118, 110, .10);
      --accent-ink: #0b514c;
      --accent-border: rgba(15, 118, 110, .34);
      --surface-strong: #10201d;
      --surface-strong-hover: #071311;
      --text-inverse: #f7fbf8;
      --canvas: rgba(255, 255, 255, .70);
      --control: rgba(255, 255, 255, .78);
      --control-soft: rgba(232, 238, 232, .72);
      --control-hover: rgba(255, 255, 255, .92);
      --active: rgba(15, 118, 110, .12);
      --preview: rgba(245, 247, 244, .78);
      --preview-head: rgba(255, 255, 255, .74);
      --detail: rgba(255, 255, 255, .66);
      --bubble: rgba(255, 255, 255, .88);
      --me: #10201d;
      --me-ink: #f7fbf8;
      --avatar: rgba(18, 32, 29, .08);
      --avatar-ink: #41504b;
      --avatar-me: rgba(15, 118, 110, .14);
      --time-bg: rgba(18, 32, 29, .06);
      --time-ink: #6e7a75;
      --link: #0f766e;
      --friend-ink: #7c4f0b;
      --friend-bg: rgba(183, 121, 31, .14);
      --official-ink: #26547c;
      --official-bg: rgba(38, 84, 124, .12);
      --other-ink: #5f6864;
      --other-bg: rgba(18, 32, 29, .06);
      --danger: #dc2626;
      --radius-sm: 10px;
      --radius-md: 14px;
      --radius-lg: 18px;
      --radius-xl: 24px;
      --radius-2xl: 30px;
      --radius-full: 999px;
      --shadow-xs: 0 1px 2px rgba(18, 32, 29, .04);
      --shadow-sm: 0 8px 24px rgba(18, 32, 29, .06);
      --shadow-md: 0 18px 48px rgba(18, 32, 29, .09);
      --shadow-lg: 0 30px 80px rgba(18, 32, 29, .13);
      --shadow: var(--shadow-md);
      --ease-out: cubic-bezier(.16, 1, .3, 1);
    }

    :root[data-theme="dark"] {
      --bg: #101612;
      --bg-deep: #172018;
      --pane: rgba(24, 33, 27, .76);
      --pane-solid: #18211b;
      --ink: #edf7f1;
      --muted: #b2beb8;
      --faint: #8b9891;
      --line: rgba(237, 247, 241, .10);
      --line-strong: rgba(237, 247, 241, .20);
      --accent: #62d3bd;
      --accent-soft: rgba(98, 211, 189, .14);
      --accent-ink: #aaf2e5;
      --accent-border: rgba(98, 211, 189, .42);
      --surface-strong: #e9f4ed;
      --surface-strong-hover: #ffffff;
      --text-inverse: #10201d;
      --canvas: rgba(26, 36, 30, .74);
      --control: rgba(31, 43, 36, .80);
      --control-soft: rgba(17, 25, 20, .74);
      --control-hover: rgba(40, 54, 45, .90);
      --active: rgba(98, 211, 189, .16);
      --preview: rgba(17, 25, 20, .80);
      --preview-head: rgba(30, 41, 34, .76);
      --detail: rgba(25, 35, 29, .72);
      --bubble: rgba(34, 46, 39, .88);
      --me: #e9f4ed;
      --me-ink: #10201d;
      --avatar: rgba(237, 247, 241, .10);
      --avatar-ink: #c8d5ce;
      --avatar-me: rgba(98, 211, 189, .18);
      --time-bg: rgba(237, 247, 241, .08);
      --time-ink: #a7b4ad;
      --friend-ink: #f5d38c;
      --friend-bg: rgba(183, 121, 31, .18);
      --official-ink: #b7d7f2;
      --official-bg: rgba(38, 84, 124, .22);
      --other-ink: #c1cbc0;
      --other-bg: rgba(237, 247, 241, .08);
      --danger: #ff9d8e;
      --shadow-xs: 0 1px 2px rgba(0, 0, 0, .18);
      --shadow-sm: 0 10px 26px rgba(0, 0, 0, .24);
      --shadow-md: 0 20px 52px rgba(0, 0, 0, .32);
      --shadow-lg: 0 34px 90px rgba(0, 0, 0, .42);
    }

    @keyframes panel-in {
      from { opacity: 0; transform: translateY(10px); }
      to { opacity: 1; transform: translateY(0); }
    }

    html {
      background: var(--bg);
    }

    body {
      position: relative;
      color: var(--ink);
      background:
        radial-gradient(circle at 12% 8%, rgba(15, 118, 110, .16), transparent 28rem),
        radial-gradient(circle at 88% 10%, rgba(183, 121, 31, .11), transparent 30rem),
        radial-gradient(circle at 50% 88%, rgba(18, 32, 29, .07), transparent 36rem),
        linear-gradient(180deg, #fbfcfa 0%, var(--bg) 48%, var(--bg-deep) 100%);
      font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
    }

    :root[data-theme="dark"] body {
      background:
        radial-gradient(circle at 10% 8%, rgba(98, 211, 189, .14), transparent 28rem),
        radial-gradient(circle at 86% 14%, rgba(183, 121, 31, .10), transparent 30rem),
        linear-gradient(180deg, #0c120f 0%, var(--bg) 56%, var(--bg-deep) 100%);
    }

    body::before {
      content: "";
      position: fixed;
      inset: 0;
      z-index: 0;
      pointer-events: none;
      background-image:
        linear-gradient(rgba(18, 32, 29, .035) 1px, transparent 1px),
        linear-gradient(90deg, rgba(18, 32, 29, .035) 1px, transparent 1px);
      background-size: 44px 44px;
      mask-image: linear-gradient(to bottom, black 0%, transparent 74%);
    }

    :root[data-theme="dark"] body::before {
      background-image:
        linear-gradient(rgba(237, 247, 241, .04) 1px, transparent 1px),
        linear-gradient(90deg, rgba(237, 247, 241, .04) 1px, transparent 1px);
    }

    main {
      position: relative;
      z-index: 1;
      background: transparent;
    }

    .chat-sidebar,
    section,
    .manager-panel {
      animation: panel-in 360ms var(--ease-out) both;
    }

    section { animation-delay: 40ms; }
    .manager-panel { animation-delay: 80ms; }

    .chat-sidebar {
      border-right: 1px solid var(--line);
      background: linear-gradient(180deg, rgba(255,255,255,.58), transparent 42%), var(--pane);
      box-shadow: 20px 0 70px rgba(18, 32, 29, .06);
      backdrop-filter: blur(22px);
      padding: 24px 18px 18px;
    }

    section {
      background: transparent;
      padding: 24px 26px;
    }

    h1 {
      display: inline-flex;
      align-items: center;
      gap: 10px;
      color: var(--ink);
      font-size: 20px;
      line-height: 1.1;
      letter-spacing: 0;
      font-weight: 760;
    }

    h1::before {
      content: "W";
      width: 34px;
      height: 34px;
      display: grid;
      place-items: center;
      border-radius: 13px;
      color: var(--text-inverse);
      background:
        radial-gradient(circle at 30% 20%, rgba(255,255,255,.22), transparent 36%),
        var(--surface-strong);
      box-shadow: var(--shadow-xs);
      font-size: 14px;
      font-weight: 800;
    }

    .count,
    .meta,
    .summary,
    .range-summary,
    .managed-meta,
    .menu-hint,
    .card-kicker {
      color: var(--muted);
    }

    .theme-toggle,
    .search,
    .tabs,
    .tab,
    .group,
    .pin-button,
    .icon-action,
    .menu,
    .format-toggle,
    .format-choice,
    .menu-item,
    .custom-export,
    .range-option,
    .range-detail,
    .export-panel,
    .schedule-panel,
    .preview,
    .preview-head,
    .bubble,
    .manager-panel,
    .manager-rail,
    .manager-toggle,
    .manager-stat,
    .mini-button,
    .managed-item,
    input,
    select,
    button {
      transition:
        background-color 180ms var(--ease-out),
        border-color 180ms var(--ease-out),
        box-shadow 180ms var(--ease-out),
        color 180ms var(--ease-out),
        transform 180ms var(--ease-out);
    }

    .theme-toggle,
    .pin-button,
    .icon-action,
    .manager-toggle,
    .manager-stat,
    .mini-button {
      border-color: var(--line);
      border-radius: var(--radius-md);
      background: var(--control);
      color: var(--muted);
      box-shadow: var(--shadow-xs);
    }

    .theme-toggle:hover,
    .pin-button:hover,
    .icon-action:hover,
    .manager-toggle:hover,
    .manager-stat:hover,
    .mini-button:hover {
      border-color: var(--accent-border);
      background: var(--control-hover);
      color: var(--accent-ink);
      transform: translateY(-1px);
      box-shadow: var(--shadow-sm);
    }

    .theme-toggle::before {
      background: var(--accent);
      box-shadow: 0 0 0 5px var(--accent-soft);
    }

    .search {
      min-height: 44px;
      margin: 18px 0 12px;
      padding: 0 14px;
      border-color: var(--line);
      border-radius: var(--radius-lg);
      background: var(--control);
      box-shadow: var(--shadow-xs);
    }

    .search:focus,
    input:focus,
    select:focus,
    button:focus-visible {
      outline: none;
      border-color: var(--accent-border);
      box-shadow: 0 0 0 4px var(--accent-soft), var(--shadow-xs);
    }

    .tabs,
    .format-toggle {
      gap: 5px;
      padding: 5px;
      border-color: var(--line);
      border-radius: var(--radius-lg);
      background: var(--control-soft);
    }

    .tab,
    .format-choice {
      border-radius: var(--radius-md);
      color: var(--muted);
      font-weight: 650;
    }

    .tab:hover,
    .format-choice:hover {
      background: rgba(255,255,255,.42);
      color: var(--ink);
    }

    .tab.active,
    .format-choice.active {
      background: var(--pane-solid);
      color: var(--ink);
      box-shadow: var(--shadow-sm);
    }

    .list {
      gap: 6px;
      padding-right: 6px;
      scrollbar-width: thin;
      scrollbar-color: var(--line-strong) transparent;
    }

    .group-row {
      grid-template-columns: minmax(0, 1fr) 36px;
      gap: 6px;
    }

    .group {
      border: 1px solid transparent;
      border-radius: var(--radius-lg);
      padding: 12px 11px;
    }

    .group:hover {
      background: rgba(255,255,255,.60);
      border-color: var(--line);
      transform: translateX(2px);
      box-shadow: var(--shadow-xs);
    }

    .group.active,
    .group.pinned {
      background: var(--active);
      border-color: var(--accent-border);
      box-shadow: inset 3px 0 0 var(--accent), var(--shadow-xs);
    }

    .pin-button {
      width: 36px;
      border-radius: var(--radius-md);
    }

    .pin-button.active,
    .icon-action.active {
      border-color: var(--accent-border);
      background: var(--accent-soft);
      color: var(--accent-ink);
    }

    .badge {
      border-radius: 8px;
      padding: 3px 6px;
      letter-spacing: 0;
      font-weight: 760;
    }

    .workspace {
      gap: 14px;
    }

    .topline {
      position: relative;
      z-index: 40;
      padding: 18px;
      border: 1px solid var(--line);
      border-radius: var(--radius-xl);
      background: var(--pane);
      box-shadow: var(--shadow-sm);
      backdrop-filter: blur(22px);
    }

    .selected {
      color: var(--ink);
      font-size: clamp(22px, 2vw, 30px);
      line-height: 1.08;
      letter-spacing: 0;
      font-weight: 760;
    }

    .action-bar {
      gap: 10px;
    }

    .action-button {
      min-width: 118px;
    }

    button.primary,
    button.secondary {
      border-radius: var(--radius-md);
      min-height: 44px;
      padding: 0 18px;
      font-weight: 700;
      box-shadow: var(--shadow-xs);
    }

    button.primary {
      background: var(--surface-strong);
      color: var(--text-inverse);
      box-shadow: var(--shadow-md);
    }

    button.primary:hover {
      filter: none;
      background: var(--surface-strong-hover);
      transform: translateY(-1px);
      box-shadow: var(--shadow-lg);
    }

    button.secondary {
      border-color: var(--line);
      background: var(--control);
      color: var(--ink);
    }

    button.secondary:hover {
      border-color: var(--accent-border);
      background: var(--control-hover);
      color: var(--accent-ink);
      transform: translateY(-1px);
      box-shadow: var(--shadow-sm);
    }

    button.primary:active,
    button.secondary:active,
    .mini-button:active,
    .group:active,
    .range-option:active {
      transform: scale(.98);
    }

    .menu {
      z-index: 80;
      gap: 8px;
      padding: 12px;
      border-color: var(--line);
      border-radius: var(--radius-xl);
      background: var(--pane);
      box-shadow: var(--shadow-lg);
      backdrop-filter: blur(22px);
    }

    .submenu {
      border-color: var(--line);
      border-radius: var(--radius-lg);
      background: var(--pane);
      box-shadow: var(--shadow-md);
      backdrop-filter: blur(18px);
    }

    .menu-label {
      color: var(--accent-ink);
      font-size: 12px;
      letter-spacing: .08em;
      text-transform: uppercase;
    }

    .menu-item {
      border-radius: var(--radius-lg);
    }

    .menu-item:hover,
    .menu-item:focus {
      background: var(--accent-soft);
    }

    .custom-export,
    .range-detail,
    .export-panel,
    .schedule-panel {
      border-color: var(--line);
      border-radius: var(--radius-xl);
      background: var(--detail);
      box-shadow: var(--shadow-sm);
      backdrop-filter: blur(18px);
    }

    label {
      color: var(--muted);
      font-weight: 620;
    }

    input[type="number"],
    input[type="datetime-local"],
    input[type="time"],
    select,
    .schedule-grid select,
    .schedule-grid input {
      min-height: 42px;
      border-color: var(--line);
      border-radius: var(--radius-md);
      background: var(--control);
      color: var(--ink);
      box-shadow: var(--shadow-xs);
    }

    .range-board {
      gap: 14px;
    }

    .range-title {
      color: var(--ink);
      font-size: 16px;
      letter-spacing: 0;
    }

    .range-option {
      border-color: var(--line);
      border-radius: var(--radius-lg);
      background: var(--control);
      box-shadow: var(--shadow-xs);
    }

    .range-option:hover {
      border-color: var(--accent-border);
      background: var(--control-hover);
      transform: translateY(-1px);
      box-shadow: var(--shadow-sm);
    }

    .range-option.active {
      background: var(--accent-soft);
      border-color: var(--accent-border);
      box-shadow: inset 3px 0 0 var(--accent), var(--shadow-xs);
    }

    .preview {
      position: relative;
      z-index: 1;
      border-color: var(--line);
      border-radius: var(--radius-2xl);
      background:
        linear-gradient(180deg, rgba(255,255,255,.48), transparent 38%),
        var(--preview);
      box-shadow: var(--shadow-md);
      backdrop-filter: blur(20px);
    }

    .preview-head {
      padding: 14px 18px;
      border-bottom-color: var(--line);
      background: var(--preview-head);
      backdrop-filter: blur(18px);
    }

    #previewTitle {
      color: var(--ink);
      font-weight: 760;
    }

    .messages {
      gap: 14px;
      padding: 22px 20px 26px;
      scrollbar-width: thin;
      scrollbar-color: var(--line-strong) transparent;
    }

    .msg {
      grid-template-columns: 38px minmax(0, 1fr);
      gap: 10px;
      max-width: 78%;
    }

    .msg.me {
      grid-template-columns: minmax(0, 1fr) 38px;
    }

    .avatar {
      width: 38px;
      height: 38px;
      border: 1px solid var(--line);
      border-radius: var(--radius-md);
      background: var(--avatar);
      color: var(--avatar-ink);
      box-shadow: var(--shadow-xs);
    }

    .msg.me .avatar {
      background: var(--avatar-me);
      color: var(--accent-ink);
    }

    .name {
      color: var(--faint);
    }

    .bubble {
      border: 1px solid var(--line);
      border-radius: 18px;
      padding: 10px 12px;
      background: var(--bubble);
      box-shadow: var(--shadow-sm);
    }

    .msg:not(.me) .bubble {
      border-top-left-radius: 8px;
    }

    .msg.me .bubble {
      border-color: transparent;
      border-top-right-radius: 8px;
      background: var(--me);
      color: var(--me-ink);
      box-shadow: var(--shadow-md);
    }

    .time-sep {
      border: 1px solid var(--line);
      border-radius: var(--radius-full);
      padding: 4px 10px;
      background: var(--time-bg);
      color: var(--time-ink);
      box-shadow: var(--shadow-xs);
    }

    .link-card,
    .file-card {
      min-width: min(280px, 100%);
      padding: 12px;
      border: 1px solid var(--line);
      border-radius: var(--radius-lg);
      background: rgba(255,255,255,.42);
    }

    .card-title {
      color: var(--ink);
      font-weight: 700;
    }

    .card-url,
    .bubble a {
      color: var(--link);
    }

    .status {
      min-height: 44px;
      max-height: 96px;
      padding: 10px 14px;
      border: 1px solid var(--line);
      border-radius: var(--radius-lg);
      background: var(--pane);
      box-shadow: var(--shadow-xs);
      backdrop-filter: blur(18px);
    }

    .ok { color: var(--accent-ink); }
    .err { color: var(--danger); }

    .manager-panel {
      border-left: 1px solid var(--line);
      background: var(--pane);
      box-shadow: -20px 0 70px rgba(18, 32, 29, .06);
      backdrop-filter: blur(22px);
    }

    .manager-rail {
      border-right-color: var(--line);
      background:
        radial-gradient(circle at 50% 0%, rgba(15,118,110,.20), transparent 10rem),
        var(--surface-strong);
      color: var(--text-inverse);
    }

    .manager-toggle {
      border-color: rgba(255,255,255,.14);
      background: rgba(255,255,255,.08);
      color: var(--text-inverse);
      box-shadow: none;
    }

    .manager-toggle:hover {
      border-color: rgba(255,255,255,.28);
      background: rgba(255,255,255,.14);
      color: var(--text-inverse);
    }

    .manager-content {
      gap: 14px;
      padding: 22px 18px 18px;
    }

    .manager-head {
      padding-bottom: 15px;
      border-bottom-color: var(--line);
    }

    .manager-title,
    .managed-title {
      color: var(--ink);
      letter-spacing: 0;
    }

    .manager-stats {
      gap: 10px;
    }

    .manager-stat {
      border: 1px solid var(--line);
      border-radius: var(--radius-lg);
      padding: 12px;
      background: var(--detail);
      box-shadow: var(--shadow-xs);
    }

    .manager-stat b {
      color: var(--ink);
      font-size: 22px;
      letter-spacing: 0;
    }

    .managed-item {
      border: 1px solid var(--line);
      border-radius: var(--radius-lg);
      padding: 11px;
      background: rgba(255,255,255,.40);
      box-shadow: var(--shadow-xs);
    }

    .managed-item.active {
      border-color: var(--accent-border);
      background: var(--accent-soft);
      box-shadow: inset 3px 0 0 var(--accent), var(--shadow-xs);
      padding-left: 11px;
    }

    .managed-actions {
      gap: 7px;
    }

    .mini-button {
      min-height: 34px;
      padding: 6px 10px;
      font-weight: 650;
    }

    .mini-button.danger {
      color: var(--danger);
    }

    .schedule-item {
      border-top-color: var(--line);
    }

    .schedule-pick {
      border-radius: var(--radius-md);
    }

    .schedule-pick:hover,
    .schedule-pick.active {
      background: var(--accent-soft);
    }

    .schedule-delete {
      border-radius: var(--radius-md);
    }

    .empty {
      color: var(--muted);
      border: 1px dashed var(--line);
      border-radius: var(--radius-lg);
      background: rgba(255,255,255,.34);
    }

    ::selection {
      color: var(--text-inverse);
      background: var(--accent);
    }

    @media (max-width: 1320px) {
      .topline {
        display: block;
      }

      .action-bar {
        justify-content: flex-start;
        margin-top: 14px;
        flex-wrap: wrap;
      }
    }

    @media (max-width: 860px) {
      body {
        overflow: auto;
      }

      main {
        grid-template-columns: 1fr;
        height: auto;
        min-height: 100vh;
        overflow: visible;
      }

      .chat-sidebar,
      section,
      .manager-panel {
        animation: none;
      }

      .chat-sidebar {
        border-right: 0;
        border-bottom: 1px solid var(--line);
        box-shadow: var(--shadow-sm);
      }

      section {
        padding: 18px;
      }

      .topline,
      .preview,
      .manager-panel {
        border-radius: var(--radius-xl);
      }

      .menu {
        left: 0;
        right: auto;
        width: min(340px, calc(100vw - 36px));
      }

      .schedule-menu {
        width: calc(100vw - 36px);
      }

      .msg {
        max-width: 94%;
      }
    }

    @media (prefers-reduced-motion: reduce) {
      *,
      *::before,
      *::after {
        animation-duration: .01ms !important;
        animation-iteration-count: 1 !important;
        scroll-behavior: auto !important;
        transition-duration: .01ms !important;
      }
    }

    /* Calmer production pass: neutral palette and export manager as an overlay drawer. */
    :root,
    :root[data-theme="dark"] {
      color-scheme: light;
      --bg: #f7f8fa;
      --bg-deep: #eef1f4;
      --pane: rgba(255, 255, 255, .94);
      --pane-solid: #ffffff;
      --ink: #17211f;
      --muted: #596561;
      --faint: #7f8a86;
      --line: rgba(23, 33, 31, .10);
      --line-strong: rgba(23, 33, 31, .18);
      --accent: #0f766e;
      --accent-soft: rgba(15, 118, 110, .08);
      --accent-ink: #0b5f59;
      --accent-border: rgba(15, 118, 110, .28);
      --surface-strong: #17211f;
      --surface-strong-hover: #0b1210;
      --text-inverse: #ffffff;
      --canvas: #ffffff;
      --control: #ffffff;
      --control-soft: #eef3f1;
      --control-hover: #f8fbfa;
      --active: rgba(15, 118, 110, .11);
      --preview: #ffffff;
      --preview-head: rgba(255, 255, 255, .96);
      --detail: #f8faf9;
      --bubble: #ffffff;
      --me: #17211f;
      --me-ink: #ffffff;
      --avatar: #edf2f0;
      --avatar-ink: #4c5a55;
      --avatar-me: #dcefed;
      --time-bg: #eef2f1;
      --time-ink: #73807b;
      --link: #0f766e;
      --friend-ink: #8a5b12;
      --friend-bg: rgba(183, 121, 31, .13);
      --official-ink: #2f5d7c;
      --official-bg: rgba(47, 93, 124, .10);
      --other-ink: #66716d;
      --other-bg: rgba(23, 33, 31, .06);
      --danger: #d92d20;
      --radius-sm: 8px;
      --radius-md: 12px;
      --radius-lg: 16px;
      --radius-xl: 20px;
      --radius-2xl: 24px;
      --shadow-xs: 0 1px 2px rgba(17, 24, 39, .04);
      --shadow-sm: 0 8px 22px rgba(17, 24, 39, .06);
      --shadow-md: 0 18px 44px rgba(17, 24, 39, .10);
      --shadow-lg: 0 28px 70px rgba(17, 24, 39, .14);
      --shadow: var(--shadow-md);
    }

    html,
    :root[data-theme="dark"] html {
      background: var(--bg);
    }

    body,
    :root[data-theme="dark"] body {
      background:
        linear-gradient(rgba(17, 24, 39, .026) 1px, transparent 1px),
        linear-gradient(90deg, rgba(17, 24, 39, .026) 1px, transparent 1px),
        linear-gradient(180deg, #fbfcfd 0%, var(--bg) 62%, var(--bg-deep) 100%);
      background-size: 40px 40px, 40px 40px, auto;
      color: var(--ink);
    }

    body::before,
    :root[data-theme="dark"] body::before {
      display: none;
    }

    main {
      grid-template-columns: minmax(300px, var(--sidebar-width, 372px)) 6px minmax(0, 1fr);
    }

    body.manager-collapsed {
      --manager-width: 0px;
    }

    .chat-sidebar {
      background: rgba(255, 255, 255, .90);
      box-shadow: none;
      backdrop-filter: blur(14px);
    }

    section {
      padding: 24px 28px 24px 26px;
    }

    h1::before {
      color: var(--accent-ink);
      background: #ffffff;
      border: 1px solid var(--line);
      box-shadow: var(--shadow-xs);
    }

    .theme-toggle {
      display: none;
    }

    .search,
    .tabs,
    .topline,
    .preview,
    .status,
    .menu,
    .custom-export,
    .range-detail,
    .export-panel,
    .schedule-panel,
    .managed-item,
    .manager-stat {
      background: var(--pane-solid);
      border-color: var(--line);
      box-shadow: var(--shadow-xs);
      backdrop-filter: none;
    }

    .topline {
      min-height: 78px;
      border-radius: var(--radius-xl);
    }

    .selected {
      font-size: clamp(24px, 2.2vw, 34px);
      line-height: 1.04;
    }

    .preview {
      border-radius: var(--radius-2xl);
      box-shadow: var(--shadow-sm);
    }

    .preview-head {
      background: var(--preview-head);
    }

    .messages {
      background:
        linear-gradient(180deg, rgba(248, 250, 252, .82), rgba(255, 255, 255, .92));
    }

    .group:hover {
      background: #f7faf9;
    }

    .group.active,
    .group.pinned {
      background: var(--active);
      border-color: var(--accent-border);
    }

    button.primary {
      background: var(--surface-strong);
      box-shadow: var(--shadow-sm);
    }

    button.primary:hover {
      background: var(--surface-strong-hover);
      box-shadow: var(--shadow-md);
    }

    .manager-panel {
      position: fixed;
      top: 64px;
      right: 18px;
      bottom: 18px;
      z-index: 120;
      width: min(430px, calc(100vw - 36px));
      min-width: 0;
      display: grid;
      grid-template-columns: 44px minmax(0, 1fr);
      border: 1px solid var(--line);
      border-radius: 18px;
      background: rgba(255, 255, 255, .96);
      box-shadow: var(--shadow-lg);
      overflow: hidden;
      animation: none;
      backdrop-filter: blur(16px);
    }

    body.manager-collapsed .manager-panel {
      top: 72px;
      right: 16px;
      bottom: auto;
      width: 44px;
      height: 140px;
      grid-template-columns: 44px 0;
      border: 0;
      border-radius: var(--radius-full);
      background: transparent;
      box-shadow: none;
      overflow: visible;
    }

    .manager-rail {
      width: 44px;
      border-right: 1px solid var(--line);
      background: #ffffff;
      color: var(--muted);
      padding: 8px 5px;
      box-shadow: none;
    }

    body.manager-collapsed .manager-rail {
      height: 140px;
      border: 1px solid var(--line);
      border-radius: var(--radius-full);
      box-shadow: var(--shadow-sm);
    }

    .manager-toggle {
      width: 32px;
      height: 32px;
      border-color: var(--line);
      background: #ffffff;
      color: var(--accent-ink);
      box-shadow: var(--shadow-xs);
      transform: none !important;
    }

    .manager-toggle:hover {
      border-color: var(--accent-border);
      background: var(--accent-soft);
      color: var(--accent-ink);
    }

    .manager-rail span {
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
    }

    .manager-content {
      min-width: 0;
      padding: 20px 18px 18px;
      background: #ffffff;
      opacity: 1;
      transition: opacity 160ms var(--ease-out), visibility 160ms var(--ease-out);
    }

    body.manager-collapsed .manager-content {
      visibility: hidden;
      opacity: 0;
      padding: 0;
      pointer-events: none;
    }

    .manager-head {
      align-items: center;
    }

    .manager-stat,
    .managed-item {
      background: #ffffff;
    }

    .managed-item.active {
      background: var(--active);
    }

    @media (max-width: 860px) {
      main {
        grid-template-columns: 1fr;
      }

      .manager-panel {
        top: 14px;
        right: 12px;
        bottom: 14px;
        width: calc(100vw - 24px);
      }

      body.manager-collapsed .manager-panel {
        top: 14px;
        right: 12px;
        width: 44px;
        height: 132px;
      }

      body.manager-collapsed .manager-rail {
        height: 132px;
      }
    }
  </style>
</head>
<body class="manager-collapsed">
<main>
  <aside class="chat-sidebar">
    <div class="aside-head">
      <h1>&#32842;&#22825;&#23548;&#20986;</h1>
      <div class="head-actions">
        <button id="themeToggle" class="theme-toggle" type="button" title="&#20999;&#25442;&#28145;&#27973;&#33394;" aria-label="&#20999;&#25442;&#28145;&#27973;&#33394;"></button>
        <span id="chatCount" class="count"></span>
      </div>
    </div>
    <input id="search" class="search" placeholder="&#25628;&#32034;&#32842;&#22825;&#23545;&#35937;" />
    <div class="tabs" aria-label="&#32842;&#22825;&#31867;&#22411;">
      <button class="tab active" data-filter="all">&#20840;&#37096;</button>
      <button class="tab" data-filter="group">&#32676;&#32842;</button>
      <button class="tab" data-filter="friend">&#22909;&#21451;</button>
      <button class="tab" data-filter="official">&#20844;&#20247;&#21495;</button>
    </div>
    <div id="groups" class="list"></div>
  </aside>
  <div id="sidebarResizer" class="sidebar-resizer" role="separator" aria-label="&#35843;&#25972;&#21015;&#34920;&#23485;&#24230;" tabindex="0"></div>
  <section>
    <div class="workspace">
      <div class="topline">
        <div>
          <div id="selectedName" class="selected">&#35831;&#36873;&#25321;&#19968;&#20010;&#32842;&#22825;</div>
          <div id="selectedMeta" class="meta"></div>
        </div>
        <div class="action-bar">
          <button id="pinSelectedBtn" class="secondary icon-action" type="button" disabled title="&#32622;&#39030;" aria-label="&#32622;&#39030;">&#9734;</button>
          <div class="menu-wrap">
            <button id="exportMenuBtn" class="primary action-button" type="button" disabled>&#23548;&#20986;</button>
            <div id="exportMenu" class="menu export-menu" hidden>
              <div class="menu-section">
                <div class="menu-label">&#26684;&#24335;</div>
                <div class="format-toggle" aria-label="&#23548;&#20986;&#26684;&#24335;">
                  <button class="format-choice active" type="button" data-format="csv">CSV</button>
                  <button class="format-choice" type="button" data-format="ai_package">&#21387;&#32553;&#21253;</button>
                </div>
              </div>
              <div class="menu-label">&#36873;&#25321;&#33539;&#22260;</div>
              <button class="menu-item" type="button" data-export-range="all">
                <span class="menu-copy"><span class="menu-title">&#20840;&#37096;&#23548;&#20986;</span><span class="menu-hint">&#21253;&#21547;&#27492;&#32842;&#22825;&#30340;&#20840;&#37096;&#35760;&#24405;</span></span>
              </button>
              <button class="menu-item" type="button" data-export-range="since_last">
                <span class="menu-copy"><span class="menu-title">&#32487;&#32493;&#23548;&#20986;</span><span class="menu-hint">&#20174;&#19978;&#27425;&#23548;&#20986;&#26102;&#38388;&#20043;&#21518;</span></span>
              </button>
              <div class="menu-item has-submenu" tabindex="0">
                <span class="menu-copy"><span class="menu-title">&#24555;&#36895;&#33539;&#22260;</span><span class="menu-hint">&#20170;&#22825;&#12289;7 &#22825;&#25110; 30 &#22825;</span></span><span class="menu-chevron">&#8250;</span>
                <div class="submenu">
                  <button class="menu-item" type="button" data-export-range="today">&#20170;&#22825;</button>
                  <button class="menu-item" type="button" data-export-range="7d">&#26368;&#36817; 7 &#22825;</button>
                  <button class="menu-item" type="button" data-export-range="30d">&#26368;&#36817; 30 &#22825;</button>
                </div>
              </div>
              <button id="customToggleBtn" class="menu-item" type="button" aria-expanded="false" aria-controls="customExportPanel">
                <span class="menu-copy"><span class="menu-title">&#33258;&#23450;&#20041;&#33539;&#22260;</span><span class="menu-hint">&#25163;&#21160;&#36873;&#25321;&#24320;&#22987;&#21644;&#32467;&#26463;&#26102;&#38388;</span></span><span class="menu-chevron">&#8964;</span>
              </button>
              <div id="customExportPanel" class="custom-export" hidden>
                <label>&#24320;&#22987;
                  <input id="customStart" type="datetime-local" />
                </label>
                <label>&#32467;&#26463;
                  <input id="customEnd" type="datetime-local" />
                </label>
                <button id="customExportBtn" class="primary" type="button" disabled>&#24320;&#22987;&#23548;&#20986;</button>
              </div>
            </div>
          </div>
          <div class="menu-wrap">
            <button id="scheduleMenuBtn" class="secondary action-button" type="button" disabled>&#23450;&#26102;&#23548;&#20986;</button>
            <div id="scheduleMenu" class="menu schedule-menu" hidden>
              <div>
                <div class="range-title">&#23450;&#26102;&#23548;&#20986;</div>
                <div id="scheduleSummary" class="range-summary"></div>
              </div>
              <div class="schedule-form">
                <label>&#39057;&#29575;
                  <select id="scheduleFrequency">
                    <option value="daily">&#27599;&#22825;</option>
                    <option value="weekly">&#27599;&#21608;</option>
                    <option value="every_hours">&#27599; N &#23567;&#26102;</option>
                  </select>
                </label>
                <label>&#26102;&#38388;
                  <input id="scheduleTime" type="time" value="08:00" />
                </label>
                <label id="scheduleWeekdayBox">&#21608;&#20960;
                  <select id="scheduleWeekday">
                    <option value="0">&#21608;&#19968;</option>
                    <option value="1">&#21608;&#20108;</option>
                    <option value="2">&#21608;&#19977;</option>
                    <option value="3">&#21608;&#22235;</option>
                    <option value="4">&#21608;&#20116;</option>
                    <option value="5">&#21608;&#20845;</option>
                    <option value="6">&#21608;&#26085;</option>
                  </select>
                </label>
                <label id="scheduleIntervalBox">&#38388;&#38548;&#23567;&#26102;
                  <input id="scheduleInterval" type="number" min="1" max="168" value="6" />
                </label>
                <label>&#26684;&#24335;
                  <select id="scheduleFormat">
                    <option value="csv">CSV</option>
                    <option value="ai_package">&#21387;&#32553;&#21253;</option>
                  </select>
                </label>
                <label>&#33539;&#22260;
                  <select id="scheduleRange">
                    <option value="since_last">&#32487;&#32493;&#23548;&#20986;</option>
                    <option value="today">&#20170;&#22825;</option>
                    <option value="7d">&#26368;&#36817; 7 &#22825;</option>
                    <option value="30d">&#26368;&#36817; 30 &#22825;</option>
                    <option value="all">&#20840;&#37096;</option>
                    <option value="custom">&#33258;&#23450;&#20041;</option>
                  </select>
                </label>
                <div id="scheduleCustomBox" class="schedule-custom">
                  <label>&#24320;&#22987;
                    <input id="scheduleCustomStart" type="datetime-local" />
                  </label>
                  <label>&#32467;&#26463;
                    <input id="scheduleCustomEnd" type="datetime-local" />
                  </label>
                </div>
              </div>
              <div class="schedule-actions">
                <button id="saveScheduleBtn" class="primary" type="button" disabled>&#20445;&#23384;&#24182;&#21551;&#29992;</button>
                <button id="deleteScheduleBtn" class="secondary" type="button" disabled>&#20572;&#27490;&#27492;&#32842;&#22825;</button>
              </div>
              <div id="scheduleList" class="schedule-list"></div>
            </div>
          </div>
        </div>
      </div>
      <div class="preview">
        <div class="preview-head">
          <span id="previewTitle">&#28040;&#24687;&#39044;&#35272;</span>
          <span id="previewMeta" class="meta"></span>
        </div>
        <div id="messages" class="messages"></div>
      </div>
      <div id="status" class="status"></div>
    </div>
  </section>
  <aside id="exportManager" class="manager-panel" aria-label="导出管理">
    <div class="manager-rail">
      <button id="managerToggle" class="manager-toggle" type="button" title="折叠导出管理" aria-label="折叠导出管理">›</button>
      <span>导出管理</span>
    </div>
    <div class="manager-content">
      <div class="manager-head">
        <div>
          <div class="manager-title">导出管理</div>
          <div class="meta">统一管理导出文件和定时任务</div>
        </div>
        <div class="manager-actions">
          <button id="refreshManagerBtn" class="mini-button" type="button">刷新</button>
          <button id="newManagedScheduleBtn" class="mini-button" type="button">新建任务</button>
        </div>
      </div>
      <div class="manager-stats">
        <button id="exportManagerTab" class="manager-stat active" type="button" data-manager-view="exports" aria-pressed="true"><b id="exportRecordCount">0</b><span>导出记录</span></button>
        <button id="scheduleManagerTab" class="manager-stat" type="button" data-manager-view="schedules" aria-pressed="false"><b id="scheduleRecordCount">0</b><span>定时任务</span></button>
      </div>
      <div class="manager-section">
        <div class="manager-section-head">
          <span id="managerListTitle">导出记录</span>
        </div>
        <div id="managedExports" class="manager-list"></div>
        <div id="managedSchedules" class="manager-list" hidden></div>
      </div>
    </div>
  </aside>
</main>
<script>
let chats = [];
let selected = null;
let selectedRange = "today";
let currentExportFormat = "csv";
let selectedFilter = "all";
let previewMessages = [];
let previewOldest = null;
let previewHasMore = false;
let previewLoading = false;
let schedules = [];
let exportRecords = [];
let pinnedChats = [];
let managerView = "exports";
const THEME_KEY = "wechat-export-theme";
const SIDEBAR_WIDTH_KEY = "wechat-export-sidebar-width";
const PINNED_CHATS_KEY = "wechat-export-pinned-chats";
const MANAGER_COLLAPSED_KEY = "wechat-export-manager-collapsed";
const MANAGER_VIEW_KEY = "wechat-export-manager-view";
const MANAGER_UI_VERSION_KEY = "wechat-export-manager-ui-version";
const $ = id => document.getElementById(id);

const RANGE_LABELS = {
  today: "今天",
  "7d": "最近 7 天",
  "30d": "最近 30 天",
  since_last: "继续导出",
  all: "全部记录",
  custom: "自定义时间"
};

function escapeHtml(s) {
  return String(s || "").replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function currentTheme() {
  return document.documentElement.dataset.theme === "dark" ? "dark" : "light";
}

function applyTheme(theme, persist = false) {
  const next = theme === "dark" ? "dark" : "light";
  document.documentElement.dataset.theme = next;
  $("themeToggle").setAttribute("aria-pressed", next === "dark" ? "true" : "false");
  $("themeToggle").setAttribute("aria-label", next === "dark" ? "切换到浅色" : "切换到深色");
  $("themeToggle").title = next === "dark" ? "切换到浅色" : "切换到深色";
  if (persist) {
    try { localStorage.setItem(THEME_KEY, next); } catch (err) {}
  }
}

function toggleTheme() {
  applyTheme(currentTheme() === "dark" ? "light" : "dark", true);
}

function loadPinnedChats() {
  try {
    const parsed = JSON.parse(localStorage.getItem(PINNED_CHATS_KEY) || "[]");
    pinnedChats = Array.isArray(parsed) ? parsed.filter(Boolean).map(String) : [];
  } catch (err) {
    pinnedChats = [];
  }
}

function savePinnedChats() {
  try { localStorage.setItem(PINNED_CHATS_KEY, JSON.stringify(pinnedChats)); } catch (err) {}
}

function isPinned(username) {
  return pinnedChats.includes(String(username || ""));
}

function setPinned(username, pinned) {
  username = String(username || "");
  if (!username) return;
  const next = new Set(pinnedChats);
  if (pinned) next.add(username);
  else next.delete(username);
  pinnedChats = Array.from(next);
  savePinnedChats();
  updatePinButton();
  if (selected && selected.username === username) {
    $("selectedMeta").textContent = selectedMeta(selected);
  }
  renderGroups();
}

function togglePinned(username) {
  setPinned(username, !isPinned(username));
}

function updatePinButton() {
  const button = $("pinSelectedBtn");
  if (!button) return;
  const pinned = selected && isPinned(selected.username);
  button.disabled = !selected;
  button.classList.toggle("active", !!pinned);
  button.textContent = pinned ? "\u2605" : "\u2606";
  button.title = pinned ? "\u53d6\u6d88\u7f6e\u9876" : "\u7f6e\u9876";
  button.setAttribute("aria-label", button.title);
  button.setAttribute("aria-pressed", pinned ? "true" : "false");
}

function groupSub(g) {
  const pieces = [];
  if (g.type === "group" && g.last_sender && g.last_summary) pieces.push(`${g.last_sender}: ${g.last_summary}`);
  else if (g.last_summary) pieces.push(g.last_summary);
  if (g.unread) pieces.push(`${g.unread} 条未读`);
  if (g.last_export_time) pieces.push(`已导出到 ${g.last_export_time}`);
  return pieces.join(" · ");
}

function selectedMeta(g) {
  const pieces = [];
  if (isPinned(g.username)) pieces.push("\u5df2\u7f6e\u9876");
  pieces.push(g.type_label || "其他");
  if (g.last_time) pieces.push(`最新 ${g.last_time}`);
  if (g.last_export_time) pieces.push(`已导出到 ${g.last_export_time}`);
  return pieces.join(" · ");
}

function renderGroups() {
  const q = $("search").value.trim().toLowerCase();
  const box = $("groups");
  const scoped = selectedFilter === "all" ? chats.filter(g => g.type !== "official") : chats.filter(g => g.type === selectedFilter);
  const visible = chats.filter(g => {
    const matchesType = selectedFilter === "all" ? g.type !== "official" : g.type === selectedFilter;
    const matchesQuery = !q || g.display_name.toLowerCase().includes(q) || g.username.toLowerCase().includes(q);
    return matchesType && matchesQuery;
  }).sort((a, b) => {
    const pinDiff = Number(isPinned(b.username)) - Number(isPinned(a.username));
    if (pinDiff) return pinDiff;
    return (b.last_timestamp || 0) - (a.last_timestamp || 0);
  });
  box.innerHTML = "";
  $("chatCount").textContent = `${visible.length} / ${scoped.length}`;
  if (!visible.length) {
    box.innerHTML = '<div class="empty">没有匹配的聊天。</div>';
    return;
  }
  visible.slice(0, 300).forEach(g => {
    const pinned = isPinned(g.username);
    const row = document.createElement("div");
    row.className = "group-row" + (pinned ? " pinned" : "");
    const btn = document.createElement("button");
    btn.className = "group" + (selected && selected.username === g.username ? " active" : "") + (pinned ? " pinned" : "");
    btn.innerHTML = `
      <div class="rowtop">
        <div class="chat-line"><span class="badge type-${g.type || "other"}">${escapeHtml(g.type_label || "")}</span><b>${escapeHtml(g.display_name)}</b></div>
        <time>${escapeHtml(g.last_time || "")}</time>
      </div>
      <div class="summary">${escapeHtml(groupSub(g))}</div>`;
    btn.onclick = () => selectGroup(g);
    const pin = document.createElement("button");
    pin.type = "button";
    pin.className = "pin-button" + (pinned ? " active" : "");
    pin.textContent = pinned ? "\u2605" : "\u2606";
    pin.title = pinned ? "\u53d6\u6d88\u7f6e\u9876" : "\u7f6e\u9876";
    pin.setAttribute("aria-label", pin.title);
    pin.setAttribute("aria-pressed", pinned ? "true" : "false");
    pin.onclick = event => {
      event.stopPropagation();
      togglePinned(g.username);
    };
    row.appendChild(btn);
    row.appendChild(pin);
    box.appendChild(row);
  });
}

function avatarText(name) {
  name = String(name || "?").trim();
  return escapeHtml(name === "me" ? "Me" : name.slice(0, 1));
}

function contentHtml(m) {
  const text = String(m.content || "");
  if (m.kind === "file") {
    return `<div class="file-card"><div class="card-kicker">File</div><div class="card-title">${escapeHtml(text.replace(/^\[file\]\s*/, ""))}</div></div>`;
  }
  if (m.kind === "link") {
    const url = m.url || "";
    const title = text.replace(/^\[link\]\s*/, "").replace(url, "").trim() || url;
    const body = `<div class="card-kicker">Link</div><div class="card-title">${escapeHtml(title)}</div><div class="card-url">${escapeHtml(url)}</div>`;
    return url ? `<a class="link-card" href="${escapeHtml(url)}" target="_blank" rel="noreferrer">${body}</a>` : `<div class="link-card">${body}</div>`;
  }
  return escapeHtml(text).replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1" target="_blank" rel="noreferrer">$1</a>');
}

function shouldShowTime(current, previous) {
  if (!current) return false;
  if (!previous) return true;
  return Math.abs((current.timestamp || 0) - (previous.timestamp || 0)) > 180;
}

function renderPreview(messages, keepScrollTop = false) {
  const box = $("messages");
  const oldHeight = box.scrollHeight;
  const oldTop = box.scrollTop;
  box.innerHTML = "";
  if (!messages.length) {
    box.innerHTML = '<div class="meta">No previewable messages.</div>';
    return;
  }
  messages.forEach((m, index) => {
    const prev = messages[index - 1];
    if (shouldShowTime(m, prev)) {
      const t = document.createElement("div");
      t.className = "time-sep";
      t.textContent = m.time;
      box.appendChild(t);
    }
    const el = document.createElement("div");
    el.className = "msg" + (m.is_me ? " me" : "");
    el.innerHTML = `
      <div class="avatar">${avatarText(m.sender)}</div>
      <div class="body">
        <div class="name">${escapeHtml(m.sender || "")}</div>
        <div class="bubble">${contentHtml(m)}</div>
      </div>`;
    box.appendChild(el);
  });
  box.scrollTop = keepScrollTop ? box.scrollHeight - oldHeight + oldTop : box.scrollHeight;
}

async function loadPreview(g, before = null) {
  if (previewLoading) return;
  previewLoading = true;
  $("previewMeta").textContent = "加载中";
  try {
    if (!before) renderPreview([]);
    const suffix = before ? `&before=${encodeURIComponent(before)}` : "";
    const res = await fetch(`/api/preview?username=${encodeURIComponent(g.username)}&limit=60${suffix}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "预览失败");
    const incoming = data.messages || [];
    if (before) {
      previewMessages = [...incoming, ...previewMessages];
      renderPreview(previewMessages, true);
    } else {
      previewMessages = incoming;
      renderPreview(previewMessages);
    }
    previewOldest = data.oldest_timestamp || (previewMessages[0] && previewMessages[0].timestamp);
    previewHasMore = !!data.has_more;
    $("previewMeta").textContent = previewHasMore ? `${previewMessages.length} 条，向上滚动加载更早消息` : `${previewMessages.length} 条消息`;
  } finally {
    previewLoading = false;
  }
}

async function selectGroup(g) {
  selected = g;
  $("selectedName").textContent = g.display_name;
  $("selectedMeta").textContent = selectedMeta(g);
  $("previewTitle").textContent = g.display_name;
  $("exportMenuBtn").disabled = false;
  $("scheduleMenuBtn").disabled = false;
  $("customExportBtn").disabled = false;
  updatePinButton();
  $("status").textContent = "";
  renderGroups();
  renderManagedExports();
  renderManagedSchedules();
  syncScheduleInputsFromJob(selectedSchedule());
  renderSchedulePanel();
  try {
    await loadPreview(g);
  } catch (err) {
    $("status").innerHTML = `<span class="err">${escapeHtml(err.message)}</span>`;
  }
}

function setMenuOpen(menu, open) {
  menu.hidden = !open;
}

function closeMenus(except = null) {
  [$("exportMenu"), $("scheduleMenu")].forEach(menu => {
    if (menu && menu !== except) menu.hidden = true;
  });
  if (except !== $("exportMenu")) setCustomExportOpen(false);
}

function toggleMenu(menu) {
  const nextOpen = menu.hidden;
  closeMenus(menu);
  setMenuOpen(menu, nextOpen);
  if (menu === $("exportMenu") && nextOpen) setCustomExportOpen(false);
}

function syncFormatButtons() {
  document.querySelectorAll(".format-choice").forEach(btn => {
    btn.classList.toggle("active", btn.dataset.format === currentExportFormat);
  });
}

function setCustomExportOpen(open) {
  const panel = $("customExportPanel");
  const button = $("customToggleBtn");
  if (!panel || !button) return;
  panel.hidden = !open;
  button.classList.toggle("active", open);
  button.setAttribute("aria-expanded", open ? "true" : "false");
}

function rangeText(mode, customStart = "", customEnd = "") {
  if (mode === "custom") {
    return `${customStart || "未设置开始"} 到 ${customEnd || "未设置结束"}`;
  }
  if (mode === "since_last" && selected && !selected.last_export_time) {
    return "继续导出（未找到上次记录时会导出全部）";
  }
  return RANGE_LABELS[mode] || mode;
}

function selectedSchedule() {
  return selected ? schedules.find(item => item.username === selected.username) : null;
}

function syncScheduleInputsFromJob(job) {
  if (!job) {
    $("scheduleRange").value = "since_last";
    $("scheduleCustomStart").value = "";
    $("scheduleCustomEnd").value = "";
    syncScheduleUI();
    return;
  }
  $("scheduleFrequency").value = job.frequency || "daily";
  $("scheduleTime").value = job.time || "08:00";
  $("scheduleWeekday").value = String(job.weekday || 0);
  $("scheduleInterval").value = String(job.interval_hours || 6);
  $("scheduleFormat").value = job.format || "csv";
  $("scheduleRange").value = job.range_mode || "since_last";
  $("scheduleCustomStart").value = job.custom_start || "";
  $("scheduleCustomEnd").value = job.custom_end || "";
  syncScheduleUI();
}

function syncScheduleUI() {
  const frequency = $("scheduleFrequency").value;
  $("scheduleWeekdayBox").style.display = frequency === "weekly" ? "grid" : "none";
  $("scheduleIntervalBox").style.display = frequency === "every_hours" ? "grid" : "none";
  $("scheduleCustomBox").classList.toggle("show", $("scheduleRange").value === "custom");
  renderSchedulePanel();
}

function renderSchedulePanel() {
  const job = selectedSchedule();
  $("saveScheduleBtn").disabled = !selected;
  $("deleteScheduleBtn").disabled = !job;
  $("saveScheduleBtn").textContent = job ? "\u4fdd\u5b58\u4fee\u6539" : "\u4fdd\u5b58\u5e76\u542f\u7528";
  $("deleteScheduleBtn").textContent = "\u5220\u9664\u6b64\u4efb\u52a1";
  if (!selected) {
    $("scheduleSummary").textContent = "先选择一个聊天。";
  } else if (job) {
    const state = job.active ? "已启用" : "已暂停";
    const next = job.next_run_time || "未排期";
    const last = job.last_run_time ? `，上次 ${job.last_run_time}` : "";
    const count = Number.isFinite(job.last_count) ? `，${job.last_count} 条消息` : "";
    const error = job.last_status === "error" && job.last_error ? `，错误：${job.last_error}` : "";
    $("scheduleSummary").textContent = `${state}：${job.frequency_label}，下次 ${next}${last}${count}${error}`;
  } else {
    $("scheduleSummary").textContent = "还没有为这个聊天设置定时导出。";
  }

  const box = $("scheduleList");
  if (!schedules.length) {
    box.innerHTML = '<div>暂无定时导出。</div>';
    return;
  }
  box.innerHTML = "";
  schedules.forEach(job => {
    const item = document.createElement("div");
    item.className = "schedule-item";
    const active = selected && selected.username === job.username;
    const status = job.last_status === "error" ? `\u9519\u8bef\uff1a${job.last_error || ""}` : (job.last_run_time ? `\u4e0a\u6b21 ${job.last_run_time}` : "\u7b49\u5f85\u4e2d");
    const pick = document.createElement("button");
    pick.type = "button";
    pick.className = "schedule-pick" + (active ? " active" : "");
    pick.innerHTML = `<span><b>${escapeHtml(job.display_name || job.username)}</b><br>${escapeHtml(job.frequency_label || "")}</span><span>${escapeHtml(job.next_run_time || "")}<br>${escapeHtml(status)}</span>`;
    pick.onclick = () => selectScheduleJob(job);
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "schedule-delete";
    remove.textContent = "\u00d7";
    remove.title = "\u5220\u9664\u8fd9\u4e2a\u5b9a\u65f6\u5bfc\u51fa";
    remove.setAttribute("aria-label", remove.title);
    remove.onclick = event => {
      event.stopPropagation();
      deleteSchedule(job.id);
    };
    item.appendChild(pick);
    item.appendChild(remove);
    box.appendChild(item);
  });
  renderManagedSchedules();
  /*
  schedules.forEach(job => {
    const item = document.createElement("div");
    item.className = "schedule-item";
    const status = job.last_status === "error" ? `错误：${job.last_error || ""}` : (job.last_run_time ? `上次 ${job.last_run_time}` : "等待中");
    item.innerHTML = `<div><b>${escapeHtml(job.display_name || job.username)}</b><br>${escapeHtml(job.frequency_label || "")}</div><div>${escapeHtml(job.next_run_time || "")}<br>${escapeHtml(status)}</div>`;
    box.appendChild(item);
  });
  */
}

function setManagerCollapsed(collapsed, persist = false) {
  document.body.classList.toggle("manager-collapsed", collapsed);
  const button = $("managerToggle");
  button.textContent = collapsed ? "\u2039" : "\u203a";
  button.title = collapsed ? "\u5c55\u5f00\u5bfc\u51fa\u7ba1\u7406" : "\u6536\u8d77\u5bfc\u51fa\u7ba1\u7406";
  button.setAttribute("aria-label", button.title);
  button.setAttribute("aria-expanded", collapsed ? "false" : "true");
  if (persist) {
    try { localStorage.setItem(MANAGER_COLLAPSED_KEY, collapsed ? "1" : "0"); } catch (err) {}
  }
}

function initManagerPanel() {
  let collapsed = true;
  try {
    if (localStorage.getItem(MANAGER_UI_VERSION_KEY) !== "2") {
      localStorage.setItem(MANAGER_COLLAPSED_KEY, "1");
      localStorage.setItem(MANAGER_UI_VERSION_KEY, "2");
    } else {
      collapsed = localStorage.getItem(MANAGER_COLLAPSED_KEY) !== "0";
    }
  } catch (err) {}
  setManagerCollapsed(collapsed);
  try {
    const savedView = localStorage.getItem(MANAGER_VIEW_KEY);
    if (savedView === "exports" || savedView === "schedules") managerView = savedView;
  } catch (err) {}
  setManagerView(managerView);
}

function setManagerView(view, persist = false) {
  managerView = view === "schedules" ? "schedules" : "exports";
  const exportsActive = managerView === "exports";
  const exportsBox = $("managedExports");
  const schedulesBox = $("managedSchedules");
  const title = $("managerListTitle");
  if (exportsBox) exportsBox.hidden = !exportsActive;
  if (schedulesBox) schedulesBox.hidden = exportsActive;
  if (title) title.textContent = exportsActive ? "导出记录" : "定时任务";
  document.querySelectorAll("[data-manager-view]").forEach(tab => {
    const active = tab.dataset.managerView === managerView;
    tab.classList.toggle("active", active);
    tab.setAttribute("aria-pressed", active ? "true" : "false");
  });
  if (persist) {
    try { localStorage.setItem(MANAGER_VIEW_KEY, managerView); } catch (err) {}
  }
}

function renderManagedExports() {
  const box = $("managedExports");
  if (!box) return;
  $("exportRecordCount").textContent = String(exportRecords.length);
  if (!exportRecords.length) {
    box.innerHTML = '<div class="empty">暂无导出记录。</div>';
    return;
  }
  box.innerHTML = "";
  exportRecords.forEach(record => {
    const item = document.createElement("div");
    const active = selected && record.username && selected.username === record.username;
    item.className = "managed-item" + (active ? " active" : "");
    const count = Number.isFinite(record.message_count) ? ` · ${record.message_count} 条` : "";
    const details = [record.kind || "文件", record.size_label || ""].filter(Boolean).join(" · ") + count;
    const exportedAt = record.exported_at ? `<br>${escapeHtml(record.exported_at)}` : "";
    const fallbackTitle = record.display_name || (record.path ? String(record.path).split(/[\\/]/).pop() : "");
    item.innerHTML = `
      <div class="managed-title">${escapeHtml(record.title || fallbackTitle)}</div>
      <div class="managed-meta">${escapeHtml(details)}${exportedAt}</div>
      <div class="managed-actions"></div>`;
    if (record.username) {
      item.querySelector(".managed-title").style.cursor = "pointer";
      item.querySelector(".managed-title").onclick = () => {
        const chat = chats.find(g => g.username === record.username);
        if (chat) selectGroup(chat);
      };
    }
    const actions = item.querySelector(".managed-actions");
    const open = document.createElement("button");
    open.type = "button";
    open.className = "mini-button";
    open.textContent = "打开";
    open.disabled = !record.exists;
    open.onclick = () => openExportRecord(record.path);
    actions.appendChild(open);
    const openFolder = document.createElement("button");
    openFolder.type = "button";
    openFolder.className = "mini-button";
    openFolder.textContent = "打开文件夹";
    openFolder.disabled = !record.exists;
    openFolder.onclick = () => openExportFolder(record.path);
    actions.appendChild(openFolder);
    const removeFile = document.createElement("button");
    removeFile.type = "button";
    removeFile.className = "mini-button danger";
    removeFile.textContent = "删除文件";
    removeFile.disabled = !record.exists;
    removeFile.onclick = () => deleteExportRecord(record);
    actions.appendChild(removeFile);
    box.appendChild(item);
  });
}

function renderManagedSchedules() {
  const box = $("managedSchedules");
  if (!box) return;
  $("scheduleRecordCount").textContent = String(schedules.length);
  if (!schedules.length) {
    box.innerHTML = '<div class="empty">暂无定时任务。</div>';
    return;
  }
  box.innerHTML = "";
  schedules.forEach(job => {
    const item = document.createElement("div");
    const active = selected && selected.username === job.username;
    item.className = "managed-item" + (active ? " active" : "");
    const status = job.last_status === "error" ? `错误：${job.last_error || ""}` : (job.last_run_time ? `上次 ${job.last_run_time}` : "等待中");
    item.innerHTML = `
      <div class="managed-title">${escapeHtml(job.display_name || job.username)}</div>
      <div class="managed-meta">${escapeHtml(job.frequency_label || "")}<br>下次 ${escapeHtml(job.next_run_time || "未排期")} · ${escapeHtml(status)}</div>
      <div class="managed-actions"></div>`;
    item.querySelector(".managed-title").style.cursor = "pointer";
    item.querySelector(".managed-title").onclick = () => selectScheduleJob(job);
    const actions = item.querySelector(".managed-actions");
    const edit = document.createElement("button");
    edit.type = "button";
    edit.className = "mini-button";
    edit.textContent = "编辑";
    edit.onclick = () => selectScheduleJob(job);
    actions.appendChild(edit);
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "mini-button danger";
    remove.textContent = "删除";
    remove.onclick = () => deleteSchedule(job.id);
    actions.appendChild(remove);
    box.appendChild(item);
  });
}

async function loadExportRecords() {
  const res = await fetch("/api/export-records");
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Export records load failed");
  exportRecords = data.records || [];
  renderManagedExports();
}

async function openExportRecord(path) {
  if (!path) return;
  $("status").textContent = "正在打开本地文件...";
  try {
    const res = await fetch("/api/export-records/open", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({path})
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Open failed");
    $("status").innerHTML = '<span class="ok">已打开本地位置</span>';
  } catch (err) {
    $("status").innerHTML = `<span class="err">${escapeHtml(err.message)}</span>`;
  }
}

async function openExportFolder(path) {
  if (!path) return;
  $("status").textContent = "正在打开导出文件夹...";
  try {
    const res = await fetch("/api/export-records/open-folder", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({path})
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Open folder failed");
    $("status").innerHTML = '<span class="ok">已打开导出文件夹</span>';
  } catch (err) {
    $("status").innerHTML = `<span class="err">${escapeHtml(err.message)}</span>`;
  }
}

async function deleteExportRecord(record) {
  if (!record) return;
  if (!confirm("删除本地导出文件？如果是压缩包，也会一起删除同名解压文件夹。")) return;
  $("status").textContent = "正在删除导出文件...";
  try {
    const res = await fetch("/api/export-records/delete", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({path: record.path, username: record.username, delete_file: true})
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Delete failed");
    exportRecords = data.records || [];
    renderManagedExports();
    await loadGroups();
    $("status").innerHTML = '<span class="ok">导出文件已删除</span>';
  } catch (err) {
    $("status").innerHTML = `<span class="err">${escapeHtml(err.message)}</span>`;
  }
}

async function refreshManager() {
  $("status").textContent = "正在刷新导出管理...";
  try {
    await Promise.all([loadExportRecords(), loadSchedules()]);
    $("status").innerHTML = '<span class="ok">导出管理已刷新</span>';
  } catch (err) {
    $("status").innerHTML = `<span class="err">${escapeHtml(err.message)}</span>`;
  }
}

function createManagedSchedule() {
  if (!selected) {
    $("status").textContent = "先在左侧选择一个聊天，再新建定时任务。";
    return;
  }
  syncScheduleInputsFromJob(null);
  closeMenus($("scheduleMenu"));
  setMenuOpen($("scheduleMenu"), true);
}

async function selectScheduleJob(job) {
  const chat = chats.find(item => item.username === job.username) || {
    username: job.username,
    display_name: job.display_name || job.username,
    type: "other",
    type_label: "Other",
    last_time: "",
    last_timestamp: 0,
    last_sender: "",
    last_summary: "",
    unread: 0,
    last_export_time: "",
    last_export_file: ""
  };
  await selectGroup(chat);
  syncScheduleInputsFromJob(job);
  setMenuOpen($("scheduleMenu"), true);
}

async function loadSchedules() {
  const res = await fetch("/api/schedules");
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Schedule load failed");
  schedules = data.schedules || [];
  renderSchedulePanel();
  renderManagedSchedules();
}

async function saveSchedule() {
  if (!selected) return;
  $("saveScheduleBtn").disabled = true;
  $("status").textContent = "正在保存定时导出...";
  const payload = {
    username: selected.username,
    format: $("scheduleFormat").value,
    range_mode: $("scheduleRange").value,
    custom_days: "",
    custom_start: $("scheduleRange").value === "custom" ? $("scheduleCustomStart").value : "",
    custom_end: $("scheduleRange").value === "custom" ? $("scheduleCustomEnd").value : "",
    frequency: $("scheduleFrequency").value,
    time: $("scheduleTime").value || "08:00",
    weekday: $("scheduleWeekday").value,
    interval_hours: $("scheduleInterval").value,
    active: true
  };
  try {
    const res = await fetch("/api/schedules", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify(payload)
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Schedule save failed");
    await loadSchedules();
    renderManagedSchedules();
    $("status").innerHTML = `<span class="ok">定时导出已保存</span>\n${escapeHtml(data.display_name)}\n下次：${escapeHtml(data.next_run_time || "")}`;
    closeMenus();
  } catch (err) {
    $("status").innerHTML = `<span class="err">${escapeHtml(err.message)}</span>`;
  } finally {
    $("saveScheduleBtn").disabled = !selected;
  }
}

async function deleteSelectedSchedule() {
  const job = selectedSchedule();
  if (!job) return;
  return deleteSchedule(job.id);
}

async function deleteSchedule(jobId) {
  if (!jobId) return;
  $("deleteScheduleBtn").disabled = true;
  $("status").textContent = "正在停止定时导出...";
  try {
    const res = await fetch("/api/schedules/delete", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({id: jobId})
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Schedule delete failed");
    await loadSchedules();
    renderManagedSchedules();
    if (selected && selected.username === jobId) syncScheduleInputsFromJob(null);
    $("status").innerHTML = `<span class="ok">已停止定时导出</span>`;
  } catch (err) {
    $("status").innerHTML = `<span class="err">${escapeHtml(err.message)}</span>`;
  } finally {
    $("deleteScheduleBtn").disabled = !selectedSchedule();
  }
}

async function loadGroups() {
  const res = await fetch("/api/chats");
  const data = await res.json();
  chats = data.chats || data.groups || [];
  if (selected) {
    const updated = chats.find(g => g.username === selected.username);
    if (updated) selected = updated;
    $("selectedMeta").textContent = selectedMeta(selected);
  }
  renderGroups();
  updatePinButton();
  renderManagedExports();
  renderSchedulePanel();
  if (!selected) $("status").textContent = `已加载 ${chats.length} 个聊天。`;
}

async function exportRange(mode) {
  selectedRange = mode;
  return exportSelected(currentExportFormat);
}

async function exportCustomRange() {
  selectedRange = "custom";
  return exportSelected(currentExportFormat);
}

async function exportSelected(format) {
  if (!selected) return;
  $("exportMenuBtn").disabled = true;
  $("scheduleMenuBtn").disabled = true;
  $("customExportBtn").disabled = true;
  const rangeLabel = rangeText(selectedRange, $("customStart").value, $("customEnd").value);
  $("status").textContent = format === "ai_package" ? `正在导出压缩包：${rangeLabel}` : `正在导出 CSV：${rangeLabel}`;
  const payload = {
    username: selected.username,
    format,
    range_mode: selectedRange,
    custom_days: "",
    custom_start: $("customStart").value,
    custom_end: $("customEnd").value
  };
  try {
    const res = await fetch("/api/export", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify(payload)
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Export failed");
    const imageLine = format === "ai_package" ? `\n图片：${data.image_count || 0}，失败：${data.image_failed || 0}\n文件夹：${escapeHtml(data.folder || "")}` : "";
    $("status").innerHTML = `<span class="ok">导出完成</span>\n${escapeHtml(data.chat)}\n${data.count} 条消息${imageLine}\n${escapeHtml(data.file)}`;
    closeMenus();
    await loadGroups();
    await loadExportRecords();
    const updated = chats.find(g => g.username === selected.username);
    if (updated) selected = updated;
  } catch (err) {
    $("status").innerHTML = `<span class="err">${escapeHtml(err.message)}</span>`;
  } finally {
    $("exportMenuBtn").disabled = !selected;
    $("scheduleMenuBtn").disabled = !selected;
    $("customExportBtn").disabled = !selected;
  }
}

function setSidebarWidth(width) {
  const max = Math.min(620, Math.max(320, window.innerWidth - 420));
  const next = Math.max(280, Math.min(max, Math.round(width)));
  document.documentElement.style.setProperty("--sidebar-width", `${next}px`);
  try { localStorage.setItem(SIDEBAR_WIDTH_KEY, String(next)); } catch (err) {}
}

function initSidebarResize() {
  try {
    const saved = parseInt(localStorage.getItem(SIDEBAR_WIDTH_KEY) || "", 10);
    if (Number.isFinite(saved)) setSidebarWidth(saved);
  } catch (err) {}

  const handle = $("sidebarResizer");
  let dragging = false;
  handle.addEventListener("pointerdown", event => {
    dragging = true;
    handle.classList.add("dragging");
    handle.setPointerCapture(event.pointerId);
    event.preventDefault();
  });
  handle.addEventListener("pointermove", event => {
    if (!dragging) return;
    setSidebarWidth(event.clientX);
  });
  handle.addEventListener("pointerup", event => {
    dragging = false;
    handle.classList.remove("dragging");
    try { handle.releasePointerCapture(event.pointerId); } catch (err) {}
  });
  handle.addEventListener("keydown", event => {
    const current = parseInt(getComputedStyle(document.documentElement).getPropertyValue("--sidebar-width"), 10) || 390;
    if (event.key === "ArrowLeft") {
      setSidebarWidth(current - 24);
      event.preventDefault();
    }
    if (event.key === "ArrowRight") {
      setSidebarWidth(current + 24);
      event.preventDefault();
    }
  });
}

$("search").addEventListener("input", renderGroups);
document.querySelectorAll(".tab").forEach(btn => {
  btn.addEventListener("click", () => {
    selectedFilter = btn.dataset.filter;
    document.querySelectorAll(".tab").forEach(tab => tab.classList.toggle("active", tab === btn));
    renderGroups();
  });
});

$("exportMenuBtn").addEventListener("click", event => {
  event.stopPropagation();
  toggleMenu($("exportMenu"));
});
$("scheduleMenuBtn").addEventListener("click", event => {
  event.stopPropagation();
  toggleMenu($("scheduleMenu"));
});
$("pinSelectedBtn").addEventListener("click", () => {
  if (selected) togglePinned(selected.username);
});
[$("exportMenu"), $("scheduleMenu")].forEach(menu => {
  menu.addEventListener("click", event => event.stopPropagation());
});
$("exportManager").addEventListener("click", event => event.stopPropagation());
document.addEventListener("click", () => closeMenus());
document.addEventListener("keydown", event => {
  if (event.key === "Escape") closeMenus();
});

document.querySelectorAll(".format-choice").forEach(btn => {
  btn.addEventListener("click", () => {
    currentExportFormat = btn.dataset.format;
    syncFormatButtons();
  });
});
document.querySelectorAll("[data-export-range]").forEach(btn => {
  btn.addEventListener("click", () => exportRange(btn.dataset.exportRange));
});
$("customToggleBtn").addEventListener("click", () => {
  setCustomExportOpen($("customExportPanel").hidden);
});
$("customExportBtn").addEventListener("click", exportCustomRange);
$("scheduleFrequency").addEventListener("change", syncScheduleUI);
$("scheduleTime").addEventListener("input", renderSchedulePanel);
$("scheduleWeekday").addEventListener("change", renderSchedulePanel);
$("scheduleInterval").addEventListener("input", renderSchedulePanel);
$("scheduleFormat").addEventListener("change", renderSchedulePanel);
$("scheduleRange").addEventListener("change", syncScheduleUI);
$("scheduleCustomStart").addEventListener("input", renderSchedulePanel);
$("scheduleCustomEnd").addEventListener("input", renderSchedulePanel);
$("messages").addEventListener("scroll", () => {
  if (!selected || !previewHasMore || previewLoading) return;
  if ($("messages").scrollTop < 80 && previewOldest) {
    loadPreview(selected, previewOldest).catch(err => {
      $("status").innerHTML = `<span class="err">${escapeHtml(err.message)}</span>`;
    });
  }
});
$("saveScheduleBtn").addEventListener("click", saveSchedule);
$("deleteScheduleBtn").addEventListener("click", deleteSelectedSchedule);
$("themeToggle").addEventListener("click", toggleTheme);
$("managerToggle").addEventListener("click", () => {
  setManagerCollapsed(!document.body.classList.contains("manager-collapsed"), true);
});
$("refreshManagerBtn").addEventListener("click", refreshManager);
$("newManagedScheduleBtn").addEventListener("click", createManagedSchedule);
document.querySelectorAll("[data-manager-view]").forEach(btn => {
  btn.addEventListener("click", () => setManagerView(btn.dataset.managerView, true));
});
initSidebarResize();
initManagerPanel();
applyTheme(currentTheme());
loadPinnedChats();
updatePinButton();
syncFormatButtons();
syncScheduleUI();
loadSchedules().catch(err => {
  $("status").innerHTML = `<span class="err">${escapeHtml(err.message)}</span>`;
});
loadExportRecords().catch(err => {
  $("status").innerHTML = `<span class="err">${escapeHtml(err.message)}</span>`;
});
loadGroups().catch(err => {
  $("status").innerHTML = `<span class="err">${escapeHtml(err.message)}</span>`;
});
</script>
</body>
</html>
"""


class Handler(BaseHTTPRequestHandler):
    def _send_json(self, payload, status=200):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_html(self, body):
        raw = body.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def _send_empty(self, status=204):
        self.send_response(status)
        self.send_header("Content-Length", "0")
        self.end_headers()

    def log_message(self, fmt, *args):
        sys.stderr.write("%s\n" % (fmt % args))

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self._send_html(INDEX_HTML)
            return
        if parsed.path == "/favicon.ico":
            self._send_empty()
            return
        if parsed.path in ("/api/chats", "/api/groups"):
            query = parse_qs(parsed.query)
            chats = list_chats()
            q = (query.get("q") or [""])[0].strip().lower()
            if q:
                chats = [g for g in chats if q in g["display_name"].lower() or q in g["username"].lower()]
            self._send_json({"chats": chats, "groups": chats})
            return
        if parsed.path == "/api/preview":
            query = parse_qs(parsed.query)
            try:
                payload = preview_chat(
                    (query.get("username") or [""])[0],
                    limit=int((query.get("limit") or ["30"])[0]),
                    before_ts=(query.get("before") or [None])[0],
                )
                self._send_json(payload)
            except Exception as exc:
                self._send_json({"error": str(exc)}, status=400)
            return
        if parsed.path == "/api/schedules":
            self._send_json({"schedules": list_schedules()})
            return
        if parsed.path == "/api/export-records":
            self._send_json({"records": list_export_records()})
            return
        self._send_json({"error": "Not found"}, status=404)

    def do_POST(self):
        path = urlparse(self.path).path
        if path == "/api/schedules":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                self._send_json(upsert_schedule(payload))
            except Exception as exc:
                self._send_json({"error": str(exc)}, status=400)
            return
        if path == "/api/schedules/delete":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                self._send_json(delete_schedule(payload.get("id", "")))
            except Exception as exc:
                self._send_json({"error": str(exc)}, status=400)
            return
        if path == "/api/export-records/delete":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                self._send_json(delete_export_record(payload))
            except Exception as exc:
                self._send_json({"error": str(exc)}, status=400)
            return
        if path == "/api/export-records/open":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                self._send_json(open_export_record(payload.get("path", "")))
            except Exception as exc:
                self._send_json({"error": str(exc)}, status=400)
            return
        if path == "/api/export-records/open-folder":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                self._send_json(open_export_folder(payload.get("path", "")))
            except Exception as exc:
                self._send_json({"error": str(exc)}, status=400)
            return
        if path != "/api/export":
            self._send_json({"error": "Not found"}, status=404)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            export_format = payload.get("format", "csv")
            export_func = export_chat_ai_package if export_format == "ai_package" else export_chat_csv
            result = export_func(
                payload.get("username", ""),
                range_mode=payload.get("range_mode", "today"),
                custom_days=payload.get("custom_days"),
                custom_start=payload.get("custom_start", ""),
                custom_end=payload.get("custom_end", ""),
            )
            self._send_json(result)
        except Exception as exc:
            self._send_json({"error": str(exc)}, status=400)


def run(host="127.0.0.1", port=None, open_browser=True):
    port = _find_available_port(host, 8765) if port is None else port
    _start_scheduler()
    server = ThreadingHTTPServer((host, port), Handler)
    url = f"http://{host}:{port}/"
    print(f"Export UI: {url}")
    if open_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    server.serve_forever()


if __name__ == "__main__":
    run()
