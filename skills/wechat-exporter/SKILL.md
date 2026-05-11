---
name: wechat-exporter
description: Export and query local WeChat chats from this repository. Use when an agent needs to initialize Wechat Exporter, launch the local export UI, export one or more chats, query WeChat messages through MCP, or schedule recurring local WeChat chat exports.
---

# Wechat Exporter

## Core Workflow

Work from the repository root. Keep WeChat logged in and running when refreshing data.

1. Install dependencies when needed:

```powershell
python -m pip install -r requirements.txt
```

2. Prepare or refresh local data:

```powershell
python prepare_data.py
```

This auto-detects the WeChat `db_storage` directory when possible, extracts keys, and decrypts the local databases. Only edit `config.json` if auto-detection fails.

3. Choose the task entrypoint:

- Browser UI: `python export_ui.py`, then open `http://127.0.0.1:8765/`.
- Single-chat CLI export: `python export_chat.py "<chat name>" output.json`.
- MCP access: run `python mcp_server.py` as a stdio MCP server for agents that support MCP.

## Configuration

Treat `config.json` as optional first-run state, not a required manual setup step. The loader creates or updates it after finding the local WeChat data directory.

If auto-detection fails, set only `db_dir` manually. Leave these defaults unless the user asks otherwise:

```json
{
  "keys_file": "all_keys.json",
  "decrypted_dir": "decrypted",
  "decoded_image_dir": "decoded_images",
  "wechat_process": "Weixin.exe"
}
```

## Recurring Exports

For scheduled exports, use the host agent's scheduler or automation system. The recurring job should:

1. Run `python prepare_data.py`.
2. Export the desired chat with `python export_chat.py "<chat name>" <output path>` or query through MCP.
3. Save outputs outside Git-tracked files when they contain private messages.

Do not upload chat logs. All operations should remain local unless the user explicitly asks for another destination.

## Troubleshooting

If `No module named 'Crypto'` appears, install dependencies into the same Python environment used to run the scripts:

```powershell
python -m pip install -r requirements.txt
```

For `.venv`, use:

```powershell
.\.venv\Scripts\python -m pip install -r requirements.txt
.\.venv\Scripts\python export_ui.py
```
