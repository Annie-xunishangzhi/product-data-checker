"""Windows一键入口：准备项目环境，再启动仅监听本机的工作台。"""
import argparse
import os
from pathlib import Path
import socket
import subprocess
import sys
import threading
import time
import urllib.request
import venv
import webbrowser

ROOT = Path(__file__).resolve().parent


def choose_port(preferred):
    for port in range(preferred, min(preferred + 30, 65536)):
        with socket.socket() as connection:
            try:
                connection.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    raise RuntimeError("没有找到空闲本机端口，请关闭旧的工具窗口后重试。")


def open_when_ready(url, process):
    for _ in range(80):
        if process.poll() is not None:
            return
        try:
            with urllib.request.urlopen(url + "/_stcore/health", timeout=1) as response:
                if response.status == 200:
                    webbrowser.open(url)
                    return
        except OSError:
            time.sleep(.25)


def main():
    parser = argparse.ArgumentParser(description="启动商品资料工作台")
    parser.add_argument("--headless", action="store_true", help="不自动打开浏览器")
    parser.add_argument("--port", type=int, default=8501)
    args = parser.parse_args()
    if sys.version_info < (3, 11):
        parser.error("需要Python 3.11及以上版本，建议使用已验证的Python 3.12。")
    if not 1024 <= args.port <= 65506:
        parser.error("端口范围为1024至65506")
    os.chdir(ROOT)
    env_python = ROOT / ".venv" / ("Scripts/python.exe" if os.name == "nt" else "bin/python")
    if not env_python.is_file():
        print("首次启动：正在创建本项目的Python环境…", flush=True)
        venv.EnvBuilder(with_pip=True).create(ROOT / ".venv")
    probe = subprocess.run([str(env_python), "-c", "import streamlit, pandas, openpyxl"], capture_output=True)
    if probe.returncode:
        print("首次启动：正在安装依赖，需要联网。", flush=True)
        subprocess.run([str(env_python), "-m", "pip", "install", "--disable-pip-version-check", "-r", str(ROOT / "requirements.txt")], check=True)
    port = choose_port(args.port)
    url = f"http://127.0.0.1:{port}"
    print(f"商品资料工作台：{url}\n关闭此窗口或按 Ctrl+C 可停止服务。", flush=True)
    command = [str(env_python), "-m", "streamlit", "run", str(ROOT / "app.py"), "--server.address", "127.0.0.1", "--server.port", str(port), "--server.headless", "true", "--browser.gatherUsageStats", "false"]
    child_env = os.environ.copy()
    child_env["PYTHONUTF8"] = "1"
    child_env["PYTHONIOENCODING"] = "utf-8"
    process = subprocess.Popen(command, cwd=ROOT, env=child_env)
    if not args.headless:
        threading.Thread(target=open_when_ready, args=(url, process), daemon=True).start()
    try:
        return process.wait()
    except KeyboardInterrupt:
        process.terminate()
        try:
            process.wait(timeout=8)
        except subprocess.TimeoutExpired:
            process.kill()
        return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (OSError, RuntimeError, subprocess.CalledProcessError) as error:
        print(f"启动失败：{error}\n可查看同目录的使用说明.md。", file=sys.stderr)
        sys.exit(1)
