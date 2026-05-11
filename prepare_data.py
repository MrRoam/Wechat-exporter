import subprocess
import sys


def _run(args):
    print(f"\n$ {' '.join(args)}", flush=True)
    completed = subprocess.run([sys.executable, *args])
    if completed.returncode != 0:
        raise SystemExit(completed.returncode)


def main():
    print("准备微信导出数据：自动定位数据目录、提取密钥、解密数据库。", flush=True)
    print("请确认 Windows 微信正在登录并保持运行。", flush=True)
    _run(["find_all_keys.py"])
    _run(["decrypt_db.py"])
    print("\n完成。现在可以运行：python export_ui.py", flush=True)


if __name__ == "__main__":
    main()
