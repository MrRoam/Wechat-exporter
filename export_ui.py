import csv
import json
import os
import re
import sqlite3
import subprocess
import sys
import threading
import webbrowser
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
SYSTEM_TYPES = {10000, 10002}
URL_RE = re.compile(r"https?://[^\s<>'\"]+")


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


def _safe_filename(value):
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", value).strip(" ._")
    return value[:80] or "wechat_export"


def _format_time(ts, minutes=False):
    if not ts:
        return ""
    fmt = "%Y-%m-%d %H:%M" if minutes else "%Y-%m-%d %H:%M:%S"
    return datetime.fromtimestamp(ts).strftime(fmt)


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
        return "[图片]"
    if base_type == 34:
        return "[语音]"
    if base_type == 43:
        return "[视频]"
    if base_type == 47:
        return "[表情]"
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
        return f"[文件] {title}" if title else "[文件]"
    if url:
        return f"[链接] {title}\n{url}" if title else f"[链接] {url}"
    return f"[链接/文件] {title}" if title else "[链接/文件]"


def _preview_kind(content):
    if not content:
        return "text"
    if content.startswith("[文件]"):
        return "file"
    if content.startswith("[链接]") or URL_RE.search(content):
        return "link"
    return "text"


def _preview_url(content):
    match = URL_RE.search(content or "")
    return match.group(0) if match else ""


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
            WHERE username LIKE '%@chatroom%' AND last_timestamp > 0
            ORDER BY sort_timestamp DESC
            """
        ).fetchall()


def list_groups():
    names = mcp_server.get_contact_names()
    state = _load_state()
    groups = []
    seen = set()
    for username, unread, summary, last_ts, _sort_ts, last_sender in _session_rows():
        seen.add(username)
        groups.append(
            {
                "username": username,
                "display_name": names.get(username, username),
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
            groups.append(
                {
                    "username": username,
                    "display_name": display_name,
                    "last_time": "",
                    "last_timestamp": 0,
                    "last_sender": "",
                    "last_summary": "",
                    "unread": 0,
                    "last_export_time": _format_time(state.get(username, {}).get("last_timestamp"), minutes=True),
                    "last_export_file": state.get(username, {}).get("file", ""),
                }
            )
    return groups


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


def preview_group(username, limit=40, before_ts=None):
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


def export_group(username, range_mode="today", custom_days=None, custom_start="", custom_end=""):
    ctx = mcp_server._resolve_chat_context(username)
    if not ctx:
        raise ValueError(f"找不到群聊: {username}")
    if not ctx["is_group"]:
        raise ValueError("只能导出群聊")
    if not ctx["message_tables"]:
        raise ValueError(f"找不到 {ctx['display_name']} 的消息表")

    state = _load_state()
    last_ts = state.get(ctx["username"], {}).get("last_timestamp")
    start_ts, end_ts = _range_to_timestamps(range_mode, custom_days, custom_start, custom_end, last_ts)
    if start_ts is not None and end_ts is not None and start_ts > end_ts:
        raise ValueError("开始时间不能晚于结束时间")

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
        out_rows.append({"time": _format_time(ts), "sender": sender, "content": content})

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    name = _safe_filename(ctx["display_name"])
    filename = f"{name}_{ctx['username'].replace('@', '_')}_{stamp}.csv"
    out_path = EXPORT_DIR / filename
    _ensure_export_dir()
    with out_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["time", "sender", "content"])
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
    _open_export_folder()

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


def _open_export_folder():
    try:
        if os.name == "nt":
            os.startfile(str(EXPORT_DIR))
        else:
            subprocess.Popen(["xdg-open", str(EXPORT_DIR)])
    except Exception:
        pass


INDEX_HTML = r"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>WeChat Group Export</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #f6f7f2;
      --pane: #fbfcf8;
      --ink: #20231f;
      --muted: #737a70;
      --line: #dde2d8;
      --accent: #2f7d62;
      --accent-soft: #e4efe6;
      --bubble: #ffffff;
      --me: #9eea6a;
      --danger: #a64236;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: var(--bg);
      color: var(--ink);
      font-family: Inter, "Segoe UI", "Microsoft YaHei", sans-serif;
      letter-spacing: 0;
    }
    main {
      min-height: 100vh;
      display: grid;
      grid-template-columns: minmax(310px, 390px) 1fr;
    }
    aside {
      border-right: 1px solid var(--line);
      background: var(--pane);
      padding: 24px 20px;
      min-width: 0;
    }
    section {
      padding: 24px 30px;
      min-width: 0;
    }
    h1 {
      margin: 0;
      font-size: 22px;
      font-weight: 760;
    }
    .search {
      width: 100%;
      margin: 18px 0 12px;
      padding: 12px 13px;
      border: 1px solid var(--line);
      background: white;
      border-radius: 6px;
      font-size: 14px;
      outline: none;
    }
    .list {
      display: grid;
      gap: 3px;
      max-height: calc(100vh - 112px);
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
    .group:hover { background: var(--accent-soft); transform: translateX(2px); }
    .group.active { background: #dceada; }
    .rowtop {
      display: flex;
      justify-content: space-between;
      gap: 10px;
      align-items: baseline;
    }
    .rowtop b {
      font-size: 14px;
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
    .workspace {
      max-width: 1040px;
      margin: 0 auto;
      display: grid;
      gap: 18px;
    }
    .topline {
      display: flex;
      align-items: flex-start;
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
    .controls {
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      gap: 14px;
      align-items: start;
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
      background: white;
      font-size: 14px;
      color: var(--ink);
    }
    .custom-row {
      display: none;
      grid-template-columns: 1fr 1fr;
      gap: 12px;
    }
    .custom-row.show { display: grid; }
    .range-board {
      display: grid;
      gap: 10px;
    }
    .range-title {
      color: var(--muted);
      font-size: 13px;
    }
    .range-chips {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
    }
    .chip {
      border: 1px solid var(--line);
      background: white;
      color: var(--ink);
      border-radius: 999px;
      padding: 8px 12px;
      font-size: 13px;
      cursor: pointer;
      transition: background .16s ease, border-color .16s ease, color .16s ease;
    }
    .chip:hover { border-color: #b7c6b8; background: #f4f8f3; }
    .chip.active {
      background: var(--accent);
      border-color: var(--accent);
      color: white;
    }
    .range-extra {
      display: none;
      grid-template-columns: 150px 1fr;
      gap: 12px;
      align-items: end;
    }
    .range-extra.show { display: grid; }
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
    }
    button.primary:hover { filter: brightness(.95); transform: translateY(-1px); }
    button.primary:disabled { opacity: .45; cursor: not-allowed; transform: none; }
    .preview {
      height: min(62vh, 690px);
      min-height: 420px;
      background: #f1f2ed;
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: hidden;
      display: grid;
      grid-template-rows: auto 1fr;
    }
    .preview-head {
      padding: 12px 16px;
      border-bottom: 1px solid var(--line);
      background: #f9faf6;
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
      background: #d9dfd3;
      display: grid;
      place-items: center;
      color: #596156;
      font-size: 13px;
      font-weight: 700;
    }
    .msg.me .avatar { grid-column: 2; background: #cde9ba; }
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
      color: #457461;
      font-size: 12px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .msg.me .body { text-align: right; grid-column: 1; grid-row: 1; }
    .msg.me .bubble { background: var(--me); text-align: left; }
    .time-sep {
      align-self: center;
      color: #8a9087;
      background: #e0e3dd;
      border-radius: 4px;
      padding: 3px 7px;
      font-size: 12px;
    }
    .status {
      min-height: 48px;
      color: var(--muted);
      line-height: 1.6;
      white-space: pre-wrap;
      word-break: break-word;
    }
    .ok { color: var(--accent); }
    .err { color: var(--danger); }
    @media (max-width: 860px) {
      main { grid-template-columns: 1fr; }
      aside { border-right: 0; border-bottom: 1px solid var(--line); }
      .list { max-height: 280px; }
      section { padding: 20px; }
      .controls { grid-template-columns: 1fr; }
      .custom-row { grid-template-columns: 1fr; }
      .msg { max-width: 92%; }
      .topline { display: block; }
    }
  </style>
</head>
<body>
<main>
  <aside>
    <h1>群聊导出</h1>
    <input id="search" class="search" placeholder="搜索群聊" />
    <div id="groups" class="list"></div>
  </aside>
  <section>
    <div class="workspace">
      <div class="topline">
        <div>
          <div id="selectedName" class="selected">请选择一个群聊</div>
          <div id="selectedMeta" class="meta"></div>
        </div>
      </div>
      <div class="preview">
        <div class="preview-head">
          <span id="previewTitle">消息预览</span>
          <span id="previewMeta" class="meta"></span>
        </div>
      <div id="messages" class="messages"></div>
      </div>
      <div class="controls">
        <div class="range-board">
          <div class="range-title">导出范围</div>
          <div id="rangeChips" class="range-chips">
            <button class="chip active" data-range="today">今天</button>
            <button class="chip" data-range="2d">近两天</button>
            <button class="chip" data-range="3d">近三天</button>
            <button class="chip" data-range="7d">一周</button>
            <button class="chip" data-range="30d">近一个月</button>
            <button class="chip" data-range="since_last">接着上次</button>
            <button class="chip" data-range="custom_days">近 N 天</button>
            <button class="chip" data-range="custom">自定义</button>
            <button class="chip" data-range="all">全部</button>
          </div>
          <div id="daysBox" class="range-extra">
            <label>天数
              <input id="customDays" type="number" min="1" value="5" />
            </label>
            <span class="meta">从今天往前计算，包含今天。</span>
          </div>
        </div>
        <button id="exportBtn" class="primary" disabled>导出 CSV</button>
      </div>
      <div id="customRow" class="custom-row">
        <label>开始
          <input id="customStart" type="datetime-local" />
        </label>
        <label>结束
          <input id="customEnd" type="datetime-local" />
        </label>
      </div>
      <div id="status" class="status"></div>
    </div>
  </section>
</main>
<script>
let groups = [];
let selected = null;
let selectedRange = "today";
let previewMessages = [];
let previewOldest = null;
let previewHasMore = false;
let previewLoading = false;
const $ = id => document.getElementById(id);

function escapeHtml(s) {
  return String(s || "").replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function groupSub(g) {
  const pieces = [];
  if (g.last_sender && g.last_summary) pieces.push(`${g.last_sender}: ${g.last_summary}`);
  else if (g.last_summary) pieces.push(g.last_summary);
  if (g.unread) pieces.push(`${g.unread} 条未读`);
  if (g.last_export_time) pieces.push(`已导出到 ${g.last_export_time}`);
  return pieces.join(" · ");
}

function selectedMeta(g) {
  const pieces = [];
  if (g.last_time) pieces.push(`最近 ${g.last_time}`);
  if (g.last_export_time) pieces.push(`已导出到 ${g.last_export_time}`);
  return pieces.join(" · ");
}

function renderGroups() {
  const q = $("search").value.trim().toLowerCase();
  const box = $("groups");
  box.innerHTML = "";
  groups
    .filter(g => !q || g.display_name.toLowerCase().includes(q) || g.username.toLowerCase().includes(q))
    .slice(0, 300)
    .forEach(g => {
      const btn = document.createElement("button");
      btn.className = "group" + (selected && selected.username === g.username ? " active" : "");
      btn.innerHTML = `
        <div class="rowtop"><b>${escapeHtml(g.display_name)}</b><time>${escapeHtml(g.last_time || "")}</time></div>
        <div class="summary">${escapeHtml(groupSub(g))}</div>`;
      btn.onclick = () => selectGroup(g);
      box.appendChild(btn);
    });
}

function avatarText(name) {
  name = String(name || "?").trim();
  return escapeHtml(name === "me" ? "我" : name.slice(0, 1));
}

function contentHtml(m) {
  const text = String(m.content || "");
  if (m.kind === "file") {
    return `<div class="file-card"><div class="card-kicker">文件</div><div class="card-title">${escapeHtml(text.replace(/^\[文件\]\s*/, ""))}</div></div>`;
  }
  if (m.kind === "link") {
    const url = m.url || "";
    const title = text.replace(/^\[链接\]\s*/, "").replace(url, "").trim() || url;
    const body = `<div class="card-kicker">链接</div><div class="card-title">${escapeHtml(title)}</div><div class="card-url">${escapeHtml(url)}</div>`;
    return url ? `<a class="link-card" href="${escapeHtml(url)}" target="_blank" rel="noreferrer">${body}</a>` : `<div class="link-card">${body}</div>`;
  }
  const linked = escapeHtml(text).replace(/(https?:\/\/[^\s<]+)/g, '<a href="$1" target="_blank" rel="noreferrer">$1</a>');
  return linked;
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
    box.innerHTML = '<div class="meta">没有可预览的消息。</div>';
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
  if (keepScrollTop) {
    box.scrollTop = box.scrollHeight - oldHeight + oldTop;
  } else {
    box.scrollTop = box.scrollHeight;
  }
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
    $("previewMeta").textContent = previewHasMore ? `${previewMessages.length} 条，向上滚动加载更早消息` : `${previewMessages.length} 条`;
  } finally {
    previewLoading = false;
  }
}

async function selectGroup(g) {
  selected = g;
  $("selectedName").textContent = g.display_name;
  $("selectedMeta").textContent = selectedMeta(g);
  $("previewTitle").textContent = g.display_name;
  $("exportBtn").disabled = false;
  $("status").textContent = "";
  renderGroups();
  try {
    await loadPreview(g);
  } catch (err) {
    $("status").innerHTML = `<span class="err">${escapeHtml(err.message)}</span>`;
  }
}

function syncRangeUI() {
  document.querySelectorAll(".chip").forEach(btn => {
    btn.classList.toggle("active", btn.dataset.range === selectedRange);
  });
  $("daysBox").classList.toggle("show", selectedRange === "custom_days");
  $("customRow").classList.toggle("show", selectedRange === "custom");
}

async function loadGroups() {
  const res = await fetch("/api/groups");
  const data = await res.json();
  groups = data.groups || [];
  renderGroups();
  if (!selected) $("status").textContent = `已加载 ${groups.length} 个群聊。`;
}

async function exportCsv() {
  if (!selected) return;
  $("exportBtn").disabled = true;
  $("status").textContent = "正在导出...";
  const payload = {
    username: selected.username,
    range_mode: selectedRange,
    custom_days: $("customDays").value,
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
    if (!res.ok) throw new Error(data.error || "导出失败");
    $("status").innerHTML = `<span class="ok">导出成功</span>\n${escapeHtml(data.chat)}\n${data.count} 条\n${escapeHtml(data.file)}`;
    await loadGroups();
    const updated = groups.find(g => g.username === selected.username);
    if (updated) selected = updated;
  } catch (err) {
    $("status").innerHTML = `<span class="err">${escapeHtml(err.message)}</span>`;
  } finally {
    $("exportBtn").disabled = !selected;
  }
}

$("search").addEventListener("input", renderGroups);
document.querySelectorAll(".chip").forEach(btn => {
  btn.addEventListener("click", () => {
    selectedRange = btn.dataset.range;
    syncRangeUI();
  });
});
$("messages").addEventListener("scroll", () => {
  if (!selected || !previewHasMore || previewLoading) return;
  if ($("messages").scrollTop < 80 && previewOldest) {
    loadPreview(selected, previewOldest).catch(err => {
      $("status").innerHTML = `<span class="err">${escapeHtml(err.message)}</span>`;
    });
  }
});
$("exportBtn").addEventListener("click", exportCsv);
syncRangeUI();
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

    def log_message(self, fmt, *args):
        sys.stderr.write("%s\n" % (fmt % args))

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self._send_html(INDEX_HTML)
            return
        if parsed.path == "/api/groups":
            query = parse_qs(parsed.query)
            groups = list_groups()
            q = (query.get("q") or [""])[0].strip().lower()
            if q:
                groups = [g for g in groups if q in g["display_name"].lower() or q in g["username"].lower()]
            self._send_json({"groups": groups})
            return
        if parsed.path == "/api/preview":
            query = parse_qs(parsed.query)
            try:
                payload = preview_group(
                    (query.get("username") or [""])[0],
                    limit=int((query.get("limit") or ["30"])[0]),
                    before_ts=(query.get("before") or [None])[0],
                )
                self._send_json(payload)
            except Exception as exc:
                self._send_json({"error": str(exc)}, status=400)
            return
        self._send_json({"error": "Not found"}, status=404)

    def do_POST(self):
        if urlparse(self.path).path != "/api/export":
            self._send_json({"error": "Not found"}, status=404)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            result = export_group(
                payload.get("username", ""),
                range_mode=payload.get("range_mode", "today"),
                custom_days=payload.get("custom_days"),
                custom_start=payload.get("custom_start", ""),
                custom_end=payload.get("custom_end", ""),
            )
            self._send_json(result)
        except Exception as exc:
            self._send_json({"error": str(exc)}, status=400)


def run(host="127.0.0.1", port=8765, open_browser=True):
    server = ThreadingHTTPServer((host, port), Handler)
    url = f"http://{host}:{port}/"
    print(f"Export UI: {url}")
    if open_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    server.serve_forever()


if __name__ == "__main__":
    run()
