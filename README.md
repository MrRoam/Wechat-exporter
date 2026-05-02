# Wechat Exporter

一个本地运行的微信 4.x 群聊导出工具。它提供浏览器界面，用来预览群聊消息，并按时间范围导出干净的 CSV。

导出的 CSV 只包含三列：

```csv
time,sender,content
```

## 功能

- 在浏览器中选择要导出的群聊
- 搜索群聊
- 预览最近消息，向上滚动可加载更早记录
- 消息预览接近微信气泡样式
- 链接和文件消息在预览中以卡片形式展示
- 按时间范围导出：今天、近两天、近三天、一周、近一个月、全部、接着上次、近 N 天、自定义时间段
- 自动记录每个群聊上次导出的结尾时间
- 导出后自动打开 CSV 所在文件夹
- 默认过滤 system / recall 类型消息
- 默认清理群消息内容中内嵌的 wxid 前缀，只保留 sender

## 适用场景

这个项目适合把自己本机微信里的群聊记录导出为 CSV，用于归档、检索、分析或后续整理。

它不是云服务，不上传聊天记录。所有读取、预览和导出都在本机完成。

## 环境要求

- Python 3.10+
- 微信 4.x
- Windows 10/11 优先支持

项目依赖：

```bash
pip install -r requirements.txt
```

## 准备数据

首次使用前，需要先让项目能读取本机微信数据库。复制示例配置：

```bash
copy config.example.json config.json
```

编辑 `config.json`：

```json
{
  "db_dir": "D:\\xwechat_files\\your_wxid\\db_storage",
  "keys_file": "all_keys.json",
  "decrypted_dir": "decrypted",
  "wechat_process": "Weixin.exe"
}
```

然后在微信运行时执行：

```bash
python find_all_keys.py
python decrypt_db.py
```

生成的 `all_keys.json`、`config.json`、`decrypted/` 都包含本地敏感信息或数据库内容，已经在 `.gitignore` 中排除，不要提交。

## 启动导出界面

```bash
python export_ui.py
```

默认打开：

```text
http://127.0.0.1:8765/
```

如果浏览器没有自动打开，手动访问上面的地址即可。

## 导出目录

默认导出到项目内的 `exports/`。你也可以用环境变量指定输出位置：

```powershell
$env:WECHAT_EXPORT_DIR="D:\wechat-exports"
python export_ui.py
```

导出文件和导出状态文件都不会被 Git 跟踪。

## CSV 字段

| 字段 | 含义 |
| --- | --- |
| `time` | 本地时间，格式为 `YYYY-MM-DD HH:MM:SS` |
| `sender` | 群成员显示名，自己发送的消息显示为 `me` |
| `content` | 处理后的消息内容 |

## 隐私与安全

请只导出你有权访问的数据。这个项目设计为本地工具，不会主动上传任何聊天内容。

以下内容不要提交到公开仓库：

- `config.json`
- `all_keys.json`
- `decrypted/`
- `exports/`
- `.venv/`
- 任何 `.db`、`.db-wal`、`.db-shm` 文件

这些路径已在 `.gitignore` 中排除。

## 致谢与参考

本项目的本地数据库读取与消息解析思路参考了以下开源项目，在此感谢原作者：

- [ylytdeng/wechat-decrypt](https://github.com/ylytdeng/wechat-decrypt)：微信 4.x 本地数据库读取、解析与相关工具。
- [hicccc77/WeFlow](https://github.com/hicccc77/WeFlow)：微信聊天记录导出产品形态参考。

Wechat Exporter 在此基础上聚焦为一个本地群聊 CSV 导出界面。
