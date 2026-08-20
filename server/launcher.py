# -*- coding: utf-8 -*-
"""ERL库存管理 联网版服务器 一键启动/停止助手

用法：
  python launcher.py --start   一键部署并启动服务器
  python launcher.py --stop    一键安全停止服务器
"""
from __future__ import annotations

import json
import os
import socket
import sqlite3
import subprocess
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

BASE = Path(__file__).resolve().parent
PYTHON = BASE / "python" / "python.exe"
SERVER_DIR = BASE / "server"
MAIN_PY = SERVER_DIR / "main.py"
CONFIG_PATH = SERVER_DIR / "config.json"
DB_PATH = SERVER_DIR / "data" / "meta.db"
LOG_PATH = SERVER_DIR / "server.log"
PID_PATH = SERVER_DIR / "server.pid"
PORT = 18666
HEALTH_URL = f"http://127.0.0.1:{PORT}/api/health"
HEARTBEAT_PORT = 29998


def banner(text: str) -> None:
    print("=" * 58)
    print(text)
    print("=" * 58)


def local_ip() -> str:
    try:
        probe = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        probe.connect(("8.8.8.8", 80))
        address = probe.getsockname()[0]
        probe.close()
        return address
    except OSError:
        return "127.0.0.1"


def read_admin_password() -> str:
    if DB_PATH.exists():
        try:
            connection = sqlite3.connect(str(DB_PATH))
            row = connection.execute(
                "SELECT value FROM settings WHERE key = 'admin_password'"
            ).fetchone()
            connection.close()
            if row and row[0]:
                return str(row[0])
        except Exception:
            pass
    if CONFIG_PATH.exists():
        try:
            return str(
                json.loads(CONFIG_PATH.read_text(encoding="utf-8")).get(
                    "admin_password",
                    "",
                )
            )
        except Exception:
            pass
    return ""


def environment_ok() -> bool:
    problems = []
    if not PYTHON.exists():
        problems.append(f"缺少 Python 运行环境：{PYTHON.name} 文件夹不存在")
    if not MAIN_PY.exists():
        problems.append("缺少 server 服务文件夹")
    if problems:
        banner("环境检查未通过")
        for problem in problems:
            print("  - " + problem)
        print("请确认整个文件夹完整拷贝且未被破坏。")
        return False
    return True


def first_time_setup() -> bool:
    if CONFIG_PATH.exists() and DB_PATH.exists():
        return False
    banner("首次部署向导")
    print("欢迎使用 ERL 库存管理-联网版服务器！")
    print("本向导只需配置一次，之后每次双击「启动服务器.bat」即可运行。")
    print()
    password = (
        input("请输入管理员密码（回车使用默认 admin123）：").strip() or "admin123"
    )
    confirm = input("请再次输入管理员密码确认：").strip()
    if password != confirm:
        banner("配置失败")
        print("两次输入的密码不一致，请重新运行本程序。")
        sys.exit(1)
    email = (
        input("密保邮箱（找回密码用，回车默认 security@example.com）：").strip()
        or "security@example.com"
    )
    SERVER_DIR.mkdir(parents=True, exist_ok=True)
    config = {
        "admin_password": password,
        "security_email": email,
        "port": PORT,
        "version_keep_days": 3,
        "token_ttl_hours": 12,
        "max_version_bytes": 100 * 1024 * 1024,
        "smtp_host": "",
        "smtp_port": 465,
        "smtp_user": "",
        "smtp_password": "",
        "smtp_from": "",
        "smtp_use_ssl": True,
        "smtp_use_starttls": True,
    }
    CONFIG_PATH.write_text(
        json.dumps(config, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    banner("部署配置完成")
    print(f"  管理员密码：{password}")
    print(f"  密保邮箱：{email}")
    print("提示：如需邮件找回密码功能，请按 server\\README.md 补充 SMTP 发件配置。")
    print()
    return True


def verify_password() -> bool:
    expected = read_admin_password()
    if not expected:
        print("尚未配置服务器，请先运行「启动服务器.bat」完成部署。")
        return False
    for attempt in range(3):
        entered = input("请输入管理员密码：").strip()
        if entered == expected:
            return True
        print(f"密码错误（还可尝试 {2 - attempt} 次）。")
    print("密码验证失败次数过多，已退出。")
    return False


def port_in_use() -> bool:
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(1)
    try:
        result = sock.connect_ex(("127.0.0.1", PORT))
        sock.close()
        return result == 0
    except OSError:
        sock.close()
        return False


def find_other_servers() -> dict[str, str]:
    """监听局域网心跳 3 秒，返回其他在线服务器 {ip: 描述}。"""
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    try:
        sock.bind(("0.0.0.0", HEARTBEAT_PORT))
    except OSError:
        sock.close()
        return {}
    sock.settimeout(0.8)
    found: dict[str, str] = {}
    own = local_ip()
    deadline = time.time() + 3
    while time.time() < deadline:
        try:
            data, peer = sock.recvfrom(1024)
            text = data.decode("utf-8", "replace")
            if text.startswith("ERL-SERVER|") and peer[0] != own:
                found[peer[0]] = text
        except socket.timeout:
            continue
        except OSError:
            break
    sock.close()
    return found


def check_conflict() -> bool:
    if port_in_use():
        banner("检测到服务器已在运行")
        print("本机 18666 端口已有服务器在运行，无需重复启动。")
        print("如需停止，请直接运行「退出服务器.bat」。")
        return False
    others = find_other_servers()
    if others:
        banner("检测到其他服务器在线")
        print("为避免数据冲突，局域网中已检测到以下服务器：")
        for ip, text in others.items():
            hostname = text.split("|")[1] if "|" in text else "未知主机"
            print(f"  - {ip}（{hostname}）")
        print()
        print("同一时间只应保留一台服务器。请先在其他电脑上停止那台服务器，")
        print("或者确认那台是废弃的、要改用本机这台。")
        choice = input("仍要启动本机服务器吗？(y=启动 / 其他=退出)：").strip().lower()
        if choice != "y":
            print("已取消启动。")
            return False
    return True


def http_ok(url: str) -> bool:
    try:
        with urllib.request.urlopen(url, timeout=2) as response:
            return response.status == 200
    except Exception:
        return False


def start_server() -> None:
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    log_file = open(LOG_PATH, "a", encoding="utf-8")
    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    process = subprocess.Popen(
        [str(PYTHON), str(MAIN_PY)],
        cwd=str(SERVER_DIR),
        stdout=log_file,
        stderr=subprocess.STDOUT,
        creationflags=creation_flags,
    )
    ready = False
    for _ in range(30):
        time.sleep(1)
        if http_ok(HEALTH_URL):
            ready = True
            break
        if process.poll() is not None:
            break
    if not ready:
        banner("服务器启动失败")
        print("请查看日志文件：server\\server.log")
        print("常见原因：18666 端口被占用、文件被其他程序锁定。")
        sys.exit(1)
    banner("服务器运行中")
    print(f"  本机访问：http://127.0.0.1:{PORT}")
    print(f"  局域网地址：http://{local_ip()}:{PORT}")
    print("  其他电脑的客户端可自动发现本服务器，或在客户端「服务器设置」")
    print("  中手动填写上面的局域网地址。")
    print("  数据保存在：server\\data 文件夹，请勿删除。")
    print()
    print("  输入 Q 后回车可安全停止服务器，然后关闭本窗口。")
    while True:
        command = input("> ").strip().lower()
        if command in ("q", "quit", "exit", "stop", "停止", "退出"):
            break
        print("输入 Q 并回车即可停止服务器。")
    stop_server()


def stop_server() -> bool:
    password = read_admin_password()
    if not password:
        print("未找到服务器配置，无法停止。")
        return False
    if not http_ok(HEALTH_URL):
        print("服务器当前未在运行。")
        return False
    body = json.dumps({"password": password}).encode("utf-8")
    request = urllib.request.Request(
        f"http://127.0.0.1:{PORT}/api/shutdown",
        method="POST",
        data=body,
        headers={"Content-Type": "application/json"},
    )
    try:
        urllib.request.urlopen(request, timeout=5)
    except urllib.error.HTTPError as exc:
        print(f"停止失败：{exc.read().decode('utf-8', 'replace')}")
        return False
    except Exception as exc:
        print(f"停止失败：{exc}")
        return False
    for _ in range(10):
        time.sleep(1)
        if not http_ok(HEALTH_URL):
            break
    try:
        os.remove(PID_PATH)
    except OSError:
        pass
    print("服务器已安全停止。")
    return True


def main() -> None:
    banner("ERL 库存管理 联网版服务器")
    print("作者：塔菲  |  版本：2026-08")
    if len(sys.argv) > 1 and sys.argv[1] == "--stop":
        if not environment_ok():
            sys.exit(1)
        if not verify_password():
            sys.exit(1)
        stop_server()
        return
    if not environment_ok():
        sys.exit(1)
    first_time_setup()
    if not verify_password():
        sys.exit(1)
    if not check_conflict():
        sys.exit(1)
    start_server()


if __name__ == "__main__":
    main()
