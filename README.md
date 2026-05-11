# Wechat Exporter

本项目用于在本机导出微信聊天记录，重点是把聊天内容导出为干净的 CSV，方便归档、检索、分析或交给 AI 继续处理。

读取、解密、预览和导出都在本机完成，不会上传聊天记录。

## 快速开始

先确认 Windows 微信已经登录并保持运行，然后在项目目录执行：

```powershell
python -m pip install -r requirements.txt
python prepare_data.py
python export_ui.py
```

`export_ui.py` 启动后会自动打开浏览器。默认地址是：

```text
http://127.0.0.1:8765/
```

如果 `8765` 已被占用，程序会自动换到下一个可用端口，请以终端里打印的 `Export UI: ...` 地址为准。

如果使用项目里的虚拟环境，把上面的 `python` 换成：

```powershell
.\.venv\Scripts\python
```

## 两个启动命令都要运行吗？

不一定。

`python prepare_data.py` 用来准备或刷新数据：自动定位微信数据目录、提取密钥、解密数据库。

`python export_ui.py` 用来启动本地浏览器导出界面。

建议这样用：

- 第一次运行：先 `python prepare_data.py`，再 `python export_ui.py`
- 想导出最新聊天记录：先重新跑 `python prepare_data.py`，再打开 UI
- 已经准备过数据，只是重新打开导出界面：只跑 `python export_ui.py`

## 配置

通常不需要手动编辑 `config.json`。程序会优先从微信本地配置中自动发现 `db_storage` 目录，并生成或更新 `config.json`。

只有自动发现失败时，才需要手动填写 `db_dir`：

```json
{
  "db_dir": "D:\\xwechat_files\\your_wxid\\db_storage",
  "keys_file": "all_keys.json",
  "decrypted_dir": "decrypted",
  "decoded_image_dir": "decoded_images",
  "wechat_process": "Weixin.exe"
}
```

一般只需要改 `db_dir`，其他字段保持默认即可。

## 常用命令

准备或刷新数据：

```powershell
python prepare_data.py
```

打开浏览器导出界面：

```powershell
python export_ui.py
```

命令行导出单个聊天为 JSON：

```powershell
python export_chat.py "聊天名称" output.json
```

指定 CSV 导出目录：

```powershell
$env:WECHAT_EXPORT_DIR="D:\wechat-exports"
python export_ui.py
```

## 常见问题

如果看到：

```text
ModuleNotFoundError: No module named 'Crypto'
```

说明当前 Python 环境缺少依赖，运行：

```powershell
python -m pip install -r requirements.txt
```

如果使用虚拟环境，安装依赖和运行脚本必须使用同一个 Python：

```powershell
.\.venv\Scripts\python -m pip install -r requirements.txt
.\.venv\Scripts\python prepare_data.py
.\.venv\Scripts\python export_ui.py
```

## 输出

CSV 默认导出到项目内的 `exports/` 目录，包含两列：

```csv
sender,content
```

`all_keys.json`、`config.json`、`decrypted/`、`decoded_images/`、`exports/` 等本地数据会被 Git 忽略，避免误提交私密数据。

## 致谢

感谢 [LC044/WeChatMsg](https://github.com/LC044/WeChatMsg) 和 [ylytdeng/wechat-decrypt](https://github.com/ylytdeng/wechat-decrypt) 在微信聊天记录展示、联系人解析和 WeChat 4.x 数据解密方向的探索。
