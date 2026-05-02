# Wechat Exporter

一个本地运行的微信群聊导出工具。它提供浏览器界面，用来预览群聊消息，并按时间范围导出干净的 CSV。

导出的 CSV 只包含三列：

```csv
time,sender,content
```

## 功能

- 选择性导出微信聊天记录
  - 可选择要导出的群聊，导出的时间范围（近一周、近一个月等，可自定义）
  - 可接着之前的导出结果继续导出，进而连续分析
- 可在浏览器ui中预览聊天记录
- 
  

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

| 字段        | 含义                             |
| --------- | ------------------------------ |
| `time`    | 本地时间，格式为 `YYYY-MM-DD HH:MM:SS` |
| `sender`  | 群成员显示名，自己发送的消息显示为 `me`         |
| `content` | 处理后的消息内容                       |

## 致谢与参考

感谢 [LC044/WeChatMsg](https://github.com/LC044/WeChatMsg) 在微信聊天记录导出、联系人和消息展示方面的探索。

后续实现本地微信 4.x 数据读取与消息解析时，参考了 [ylytdeng/wechat-decrypt](https://github.com/ylytdeng/wechat-decrypt)。

Wechat Exporter 在这些工作的基础上，聚焦为一个本地群聊 CSV 导出界面。
