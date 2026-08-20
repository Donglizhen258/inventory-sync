# -*- coding: utf-8 -*-
"""ERL库存管理 联网版服务器 控制台（本地网页版）

双击「服务器控制台.bat」后：本程序在 127.0.0.1:18667 提供控制台页面并自动打开浏览器。
功能：状态灯、启动/停止服务器、首次配置向导、内网穿透引导、日志查看、多服务器冲突检测。
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import socket
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from gui_window import run_gui

if getattr(sys, "frozen", False):
    BASE = Path(sys.executable).resolve().parent
else:
    BASE = Path(__file__).resolve().parent
PYTHON = BASE / "python" / "python.exe"
SERVER_DIR = BASE / "server"
DATA_DIR = SERVER_DIR / "data"
MAIN_PY = SERVER_DIR / "main.py"
CONFIG_PATH = SERVER_DIR / "config.json"
DB_PATH = SERVER_DIR / "data" / "meta.db"
LOG_PATH = SERVER_DIR / "server.log"
PID_PATH = SERVER_DIR / "server.pid"
CONSOLE_LOG = BASE / "console.log"
HTML_PATH = BASE / "console.html"

SERVER_PORT = 18666
CONSOLE_PORT = 18667
HEALTH_URL = f"http://127.0.0.1:{SERVER_PORT}/api/health"
HEARTBEAT_PORT = 29998

state_lock = threading.Lock()
server_process: subprocess.Popen | None = None
started_at: float | None = None

# 局域网/公网地址连通性检测缓存（后台线程每 5 秒刷新，避免拖慢状态轮询）
_link_status: dict = {"local": None, "public": None}


def _check_links_loop() -> None:
    """后台持续检测两个网址的连通性，供状态卡圆灯使用。"""
    while True:
        try:
            local_url = f"http://{local_ip()}:{SERVER_PORT}"
            _link_status["local"] = http_ok(f"{local_url}/api/health", timeout=2.0)
            config = read_config()
            public_url = str(config.get("public_url", "") or "").strip()
            if public_url:
                target = public_url if public_url.startswith("http") else f"http://{public_url}"
                _link_status["public"] = http_ok(f"{target}/api/health", timeout=3.0)
            else:
                _link_status["public"] = None
        except Exception:
            pass
        time.sleep(5)


# 内置内网穿透工具（免安装）：cpolar（需 authtoken）+ cloudflared（免账号临时隧道）
CPOLAR_PATH = BASE / "cpolar" / "cpolar.exe"
CLOUDFLARED_PATH = BASE / "cloudflared" / "cloudflared.exe"
TOKEN_HISTORY_PATH = BASE / "cpolar" / "token_history.json"
TOKEN_HISTORY_LIMIT = 20  # 最多记住 20 个历史令牌
tunnel_process: subprocess.Popen | None = None
tunnel_url: str = ""
tunnel_provider: str = ""
tunnel_error: str | None = None


def _mask_token(token: str) -> str:
    """令牌掩码展示：只露头尾各 4 位。"""
    token = (token or "").strip()
    if len(token) <= 8:
        return token
    return f"{token[:4]}…{token[-4:]}"


def load_token_history() -> list:
    """读取历史 cpolar 令牌列表（[{id,value,note,added}, ...]，新→旧）。"""
    try:
        if TOKEN_HISTORY_PATH.is_file():
            data = json.loads(TOKEN_HISTORY_PATH.read_text(encoding="utf-8"))
            if isinstance(data, list):
                return [e for e in data if isinstance(e, dict) and e.get("value")]
    except Exception:
        pass
    return []


def save_token_history(entries: list) -> None:
    try:
        TOKEN_HISTORY_PATH.parent.mkdir(parents=True, exist_ok=True)
        TOKEN_HISTORY_PATH.write_text(
            json.dumps(entries, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception as exc:
        log(f"保存历史令牌失败：{exc}")


def add_token_to_history(authtoken: str, note: str = "") -> list:
    """记住一个令牌：重复令牌只更新时间/备注，最多保留 20 条，新→旧。"""
    authtoken = (authtoken or "").strip()
    if not authtoken:
        return load_token_history()
    entries = load_token_history()
    now = time.strftime("%Y-%m-%d %H:%M")
    for entry in entries:
        if entry.get("value") == authtoken:
            entry["added"] = now
            if note:
                entry["note"] = note
            save_token_history(entries)
            return entries
    entries.insert(
        0,
        {
            "id": hashlib.sha1(authtoken.encode("utf-8")).hexdigest()[:12],
            "value": authtoken,
            "note": note or f"令牌 {len(entries) + 1}",
            "added": now,
        },
    )
    del entries[TOKEN_HISTORY_LIMIT:]
    save_token_history(entries)
    return entries


def delete_token_from_history(token_id: str) -> list:
    entries = [e for e in load_token_history() if e.get("id") != token_id]
    save_token_history(entries)
    return entries


def token_history_payload() -> dict:
    """供前端下拉菜单使用的历史令牌（display 为掩码，value 为完整令牌）。"""
    return {
        "ok": True,
        "tokens": [
            {
                "id": e.get("id", ""),
                "note": e.get("note", "") or "未命名",
                "added": e.get("added", "") or "",
                "display": _mask_token(e.get("value", "")),
                "value": e.get("value", ""),
            }
            for e in load_token_history()
        ],
    }


def cpolar_available() -> bool:
    return CPOLAR_PATH.is_file()


def cloudflared_available() -> bool:
    return CLOUDFLARED_PATH.is_file()


def tunnel_running() -> bool:
    return tunnel_process is not None and tunnel_process.poll() is None


def _tunnel_reader(pattern: re.Pattern, label: str) -> None:
    """后台持续读取隧道输出，解析公网地址；记录启动错误供前端展示。"""
    global tunnel_url, tunnel_error
    while tunnel_running():
        try:
            line = tunnel_process.stdout.readline()
        except Exception:
            break
        if not line:
            time.sleep(0.2)
            continue
        text = line.strip()
        if text:
            log(f"{label}: {text}")
        match = pattern.search(text)
        if match and not tunnel_url:
            tunnel_url = match.group(1)
            log(f"内网穿透公网地址：{tunnel_url}")
            continue
        # 记录明确失败信息（认证失败/连接失败等），避免 debug 行误报
        lower = text.lower()
        if not tunnel_url and tunnel_error is None and re.search(
            r"(auth[a-z]*\s+failed|failed\s+to\s+authenticate|invalid\s+auth|"
            r"connection\s+refused|cannot\s+connect|unable\s+to\s+connect|"
            r"tunnel\s+creation\s+failed)",
            lower,
        ):
            tunnel_error = text[:200]


def start_tunnel(provider: str, authtoken: str = "") -> dict:
    """启动内置穿透（provider: cpolar / cloudflared），后台解析公网地址。"""
    global tunnel_process, tunnel_url, tunnel_provider, tunnel_error
    tunnel_error = None
    provider = str(provider or "cpolar").strip().lower()
    authtoken = str(authtoken or "").strip()
    if tunnel_running():
        return {"ok": False, "error": "穿透已在运行中，请先停止。"}
    if provider == "cloudflared":
        if not cloudflared_available():
            return {"ok": False, "error": "未找到内置 cloudflared（cloudflared\\cloudflared.exe），请确认已完整拷贝服务器文件夹。"}
        try:
            tunnel_process = subprocess.Popen(
                [
                    str(CLOUDFLARED_PATH),
                    "tunnel",
                    "--url", f"http://127.0.0.1:{SERVER_PORT}",
                    "--protocol", "http2",
                    "--no-autoupdate",
                ],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=0x08000000,
            )
        except Exception as exc:
            tunnel_process = None
            log(f"启动 cloudflared 失败：{exc}")
            return {"ok": False, "error": f"启动 cloudflared 失败：{exc}"}
        tunnel_provider = "cloudflared"
        threading.Thread(
            target=_tunnel_reader,
            args=(re.compile(r"(https://[\w.-]+\.trycloudflare\.com)"), "cloudflared"),
            daemon=True,
        ).start()
        time.sleep(3)
        if not tunnel_running():
            detail = f"（{tunnel_error}）" if tunnel_error else ""
            tunnel_process = None
            tunnel_provider = ""
            tunnel_error = None
            return {"ok": False, "error": f"穿透启动失败：无法连接 Cloudflare{detail}"}
        message = f"穿透已开启，公网地址：{tunnel_url}" if tunnel_url else "穿透已启动，正在分配公网地址（约 10-30 秒）……"
        return {"ok": True, "url": tunnel_url, "message": message}
    # 默认 cpolar
    if not authtoken:
        return {"ok": False, "error": "请先粘贴 cpolar 的 authtoken（官网登录后在「验证」页复制）。"}
    if not cpolar_available():
        return {"ok": False, "error": "未找到内置 cpolar（cpolar\\cpolar.exe），请确认已完整拷贝服务器文件夹。"}
    # 保存 token 到 ~/.cpolar/cpolar.yml
    try:
        subprocess.run(
            [str(CPOLAR_PATH), "authtoken", authtoken],
            timeout=15,
            capture_output=True,
        )
    except Exception as exc:
        return {"ok": False, "error": f"保存 authtoken 失败：{exc}"}
    try:
        tunnel_process = subprocess.Popen(
            [str(CPOLAR_PATH), "http", str(SERVER_PORT), "-log", "stdout"],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=0x08000000,  # CREATE_NO_WINDOW
        )
    except Exception as exc:
        tunnel_process = None
        log(f"启动 cpolar 失败：{exc}")
        return {"ok": False, "error": f"启动 cpolar 失败：{exc}"}
    tunnel_provider = "cpolar"
    threading.Thread(
        target=_tunnel_reader,
        args=(re.compile(r"(?:Tunnel established at|Forwarding)\s+(https?://[^\s\"]+)"), "cpolar"),
        daemon=True,
    ).start()
    time.sleep(2)
    if not tunnel_running():
        detail = f"（{tunnel_error}）" if tunnel_error else ""
        tunnel_process = None
        tunnel_provider = ""
        tunnel_error = None
        return {"ok": False, "error": f"穿透启动失败：authtoken 无效或网络不通{detail}"}
    add_token_to_history(authtoken)  # 启动成功即记住该令牌，下次可直接从历史下拉选择
    message = f"穿透已开启，公网地址：{tunnel_url}" if tunnel_url else "穿透已启动，正在分配公网地址（约 5-20 秒）……"
    return {"ok": True, "url": tunnel_url, "message": message}


def stop_tunnel() -> dict:
    """停止当前穿透。"""
    global tunnel_process, tunnel_url, tunnel_provider, tunnel_error
    was_url = tunnel_url
    if tunnel_running():
        try:
            tunnel_process.terminate()
            tunnel_process.wait(timeout=5)
        except Exception:
            try:
                tunnel_process.kill()
            except Exception:
                pass
    tunnel_process = None
    tunnel_url = ""
    tunnel_provider = ""
    tunnel_error = None
    # 若保存的公网地址正是本次穿透分配的地址，一并清除，避免远程同事连接失效地址
    if was_url:
        try:
            config = read_config()
            if str(config.get("public_url", "") or "").rstrip("/") == was_url.rstrip("/"):
                config["public_url"] = ""
                write_config(config)
                log("公网地址已随穿透停止而清除")
        except Exception:
            pass
    log("内网穿透已停止")
    return {"ok": True, "message": "穿透已停止。"}


def _rotate_log(path: Path, max_bytes: int = 5_000_000, keep: int = 2) -> None:
    """日志轮转：超过 max_bytes 时归档为 .1/.2（新→旧），最多保留 keep 份。"""
    try:
        if path.is_file() and path.stat().st_size > max_bytes:
            for index in range(keep, 0, -1):
                old_file = Path(f"{path}.{index}")
                if old_file.is_file():
                    if index == keep:
                        old_file.unlink(missing_ok=True)
                    else:
                        old_file.rename(Path(f"{path}.{index + 1}"))
            path.rename(Path(f"{path}.1"))
    except OSError:
        pass


def log(message: str) -> None:
    line = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {message}"
    print(line, flush=True)
    try:
        _rotate_log(CONSOLE_LOG)
        with open(CONSOLE_LOG, "a", encoding="utf-8") as handle:
            handle.write(line + "\n")
    except OSError:
        pass


def local_ip() -> str:
    try:
        probe = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        probe.connect(("8.8.8.8", 80))
        address = probe.getsockname()[0]
        probe.close()
        return address
    except OSError:
        return "127.0.0.1"


def http_ok(url: str, timeout: float = 2.0) -> bool:
    try:
        with urllib.request.urlopen(url, timeout=timeout) as response:
            return response.status == 200
    except Exception:
        return False


def server_running() -> bool:
    return http_ok(HEALTH_URL)


def read_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def write_config(config: dict) -> None:
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(
        json.dumps(config, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def db_password() -> str:
    if DB_PATH.exists():
        try:
            connection = sqlite3_connect()
            row = connection.execute(
                "SELECT value FROM settings WHERE key = 'admin_password'"
            ).fetchone()
            connection.close()
            if row and row[0]:
                return str(row[0])
        except Exception:
            pass
    return ""


def sqlite3_connect():
    import sqlite3

    return sqlite3.connect(str(DB_PATH))


def current_admin_password() -> str:
    return db_password() or str(read_config().get("admin_password", ""))


def find_other_servers(duration: float = 3.0) -> list[dict[str, str]]:
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    try:
        sock.bind(("0.0.0.0", HEARTBEAT_PORT))
    except OSError:
        sock.close()
        return []
    sock.settimeout(0.8)
    found: dict[str, str] = {}
    own = local_ip()
    deadline = time.time() + duration
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
    return [
        {"ip": ip, "hostname": text.split("|")[1] if "|" in text else "未知主机"}
        for ip, text in found.items()
    ]


_PWD_PREFIX = "enc:"


def _pwd_key() -> bytes:
    """读取密码加密密钥（config.json 的 pwd_key），随备份整体迁移。"""
    try:
        key = str(read_config().get("pwd_key") or "")
        if key:
            return bytes.fromhex(key)
    except Exception:
        pass
    return b""


def _xor_bytes(data: bytes, key: bytes) -> bytes:
    return bytes(value ^ key[index % len(key)] for index, value in enumerate(data))


def _decrypt_user_password(stored: str) -> str:
    """解密用户密码（内联实现，与 server/main.py 同构，打包版不依赖外部模块）。"""
    import base64
    stored = (stored or "").strip()
    if not stored.startswith(_PWD_PREFIX):
        return stored
    key = _pwd_key()
    if not key:
        return stored
    try:
        raw = base64.b64decode(stored[len(_PWD_PREFIX):])
        return _xor_bytes(raw, key).decode("utf-8")
    except Exception:
        return stored


def list_users() -> list[dict]:
    if not DB_PATH.exists():
        return []
    try:
        connection = sqlite3_connect()
        rows = connection.execute(
            "SELECT password, note, created_at FROM users ORDER BY created_at"
        ).fetchall()
        connection.close()
        return [
            {
                "id": row[0],
                "password": _decrypt_user_password(row[0]),
                "note": row[1],
                "createdAt": time.strftime(
                    "%Y-%m-%d %H:%M", time.localtime(row[2])
                )
                if row[2]
                else "",
            }
            for row in rows
        ]
    except Exception:
        return []


def start_server() -> dict:
    global server_process, started_at
    with state_lock:
        if server_running():
            return {"ok": True, "message": "服务器已在运行。"}
        if not current_admin_password():
            return {
                "ok": False,
                "error": "尚未完成首次配置。请先在「首次配置」中设置管理员密码与密保邮箱，保存后再启动服务器。",
            }
        if not PYTHON.exists() or not MAIN_PY.exists():
            return {"ok": False, "error": "运行环境不完整，请确认文件夹完整拷贝。"}
        LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        _rotate_log(LOG_PATH)
        log_file = open(LOG_PATH, "a", encoding="utf-8")
        creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        server_process = subprocess.Popen(
            [str(PYTHON), str(MAIN_PY)],
            cwd=str(SERVER_DIR),
            stdout=log_file,
            stderr=subprocess.STDOUT,
            creationflags=creation_flags,
        )
        started_at = time.time()
    ready = False
    for _ in range(30):
        time.sleep(1)
        if server_running():
            ready = True
            break
        if server_process.poll() is not None:
            break
    if not ready:
        log("服务器启动失败，请查看 server.log")
        return {"ok": False, "error": "服务器启动失败，请查看下方日志。可能端口被占用或文件被锁定。"}
    log("服务器已启动")
    return {"ok": True, "message": "服务器已启动。"}


def backup_data(save_dir: str = "") -> dict:
    """一键备份数据：先停云端（释放 WAL 锁）再打包 server/data + config.json 为 zip。
    save_dir 非空时把压缩包写到指定目录，否则存到 data/backups。"""
    import zipfile
    backup_dir = DATA_DIR / "backups"
    if save_dir.strip():
        backup_dir = Path(str(save_dir).strip())
    was_running = server_running()
    stop_server()
    try:
        backup_dir.mkdir(parents=True, exist_ok=True)
        stamp = time.strftime("%Y%m%d_%H%M%S")
        target = backup_dir / f"ERL数据备份-{stamp}.zip"
        with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zf:
            for file_path in sorted(DATA_DIR.rglob("*")):
                if not file_path.is_file():
                    continue
                if "backups" in file_path.parts:
                    continue
                if file_path.suffix in (".wal", "-wal", ".shm", "-shm"):
                    continue  # SQLite 临时文件，停云端后数据已合并进主库
                zf.write(file_path, file_path.relative_to(SERVER_DIR))
            config_path = SERVER_DIR / "config.json"
            if config_path.is_file():
                zf.write(config_path, "config.json")
        log(f"数据已备份：{target}")
        result = {"ok": True, "path": str(target)}
    except Exception as exc:
        result = {"ok": False, "error": f"备份失败：{exc}"}
    if was_running:
        try:
            start_server()
        except Exception as exc:
            result = {"ok": False, "error": f"备份完成但重启云端失败：{exc}"}
    return result


def restore_data(zip_bytes: bytes) -> dict:
    """从备份 zip 恢复数据（岗位交接用）：先备份当前→停云端→校验→替换→重启云端。"""
    import io, zipfile, shutil, tempfile
    if not zip_bytes or len(zip_bytes) > 200_000_000:
        return {"ok": False, "error": "备份文件为空或过大。"}
    try:
        zf = zipfile.ZipFile(io.BytesIO(zip_bytes))
        names = zf.namelist()
    except zipfile.BadZipFile:
        return {"ok": False, "error": "不是有效的备份 zip 文件。"}
    has_db = any(n.endswith("meta.db") for n in names)
    has_versions = any("/versions/" in n for n in names)
    if not has_db:
        return {"ok": False, "error": "备份文件缺少数据库（meta.db），不是有效的 ERL 备份。"}
    if not has_versions:
        return {"ok": False, "error": "备份文件缺少版本目录（versions/），不是有效的 ERL 备份。"}
    # 解压到临时目录后做深度校验：meta.db 必须是合法 SQLite 且含用户表；versions 至少一个 Excel
    temp_dir = Path(tempfile.mkdtemp(prefix="erl_restore_"))
    try:
        zf.extractall(temp_dir)
        db_candidates = [Path(temp_dir) / n for n in names if n.endswith("meta.db")]
        valid_db = False
        if db_candidates and db_candidates[0].is_file():
            try:
                import sqlite3 as _sqlite3
                conn = _sqlite3.connect(str(db_candidates[0]))
                conn.execute("SELECT 1 FROM users LIMIT 1")
                conn.close()
                valid_db = True
            except Exception:
                valid_db = False
        xlsx_count = sum(1 for n in names if n.endswith(".xlsx") and "/versions/" in n)
        if not valid_db:
            return {"ok": False, "error": "备份中的数据库文件损坏或格式不正确，已中止恢复。"}
        if xlsx_count == 0:
            return {"ok": False, "error": "备份中没有任何 Excel 版本文件，已中止恢复。"}
        # 校验通过：备份文件列表（供后续替换使用）
        validated_names = names
    except zipfile.BadZipFile:
        return {"ok": False, "error": "不是有效的备份 zip 文件。"}
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
    # 恢复前先自动备份当前数据
    backup_data()
    # 保护密码密钥：备份 config 若无 pwd_key，恢复后沿用当前密钥，避免密码乱码
    try:
        current_pwd_key = str(read_config().get("pwd_key") or "")
    except Exception:
        current_pwd_key = ""
    was_running = server_running()
    stop_server()
    temp_dir = Path(tempfile.mkdtemp(prefix="erl_restore_"))
    try:
        zf.extractall(temp_dir)
        for name in names:
            source = temp_dir / name
            if not source.is_file():
                continue
            target = SERVER_DIR / name
            target.parent.mkdir(parents=True, exist_ok=True)
            if name == "config.json" and not current_pwd_key:
                # 当前无密钥则用备份的（原样覆盖）
                shutil.copy2(source, target)
                continue
            if name == "config.json":
                # 合并保留当前 pwd_key（备份 config 无密钥时补回）
                try:
                    import json as _json
                    new_cfg = _json.loads(source.read_text(encoding="utf-8"))
                    if not new_cfg.get("pwd_key") and current_pwd_key:
                        new_cfg["pwd_key"] = current_pwd_key
                        target.write_text(
                            _json.dumps(new_cfg, ensure_ascii=False, indent=2),
                            encoding="utf-8",
                        )
                        log("恢复 config 时已保留当前密码密钥（pwd_key）")
                        continue
                except Exception:
                    pass
            shutil.copy2(source, target)
        log("数据已从备份恢复")
        result = {"ok": True, "message": "数据已恢复。"}
    except Exception as exc:
        result = {"ok": False, "error": f"恢复失败：{exc}"}
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
    if was_running:
        try:
            start_server()
        except Exception as exc:
            result = {"ok": False, "error": f"数据已恢复，但重启云端失败：{exc}"}
    return result


def stop_server() -> dict:
    global server_process, started_at
    password = current_admin_password()
    if server_running() and password:
        body = json.dumps({"password": password}).encode("utf-8")
        request = urllib.request.Request(
            f"http://127.0.0.1:{SERVER_PORT}/api/shutdown",
            method="POST",
            data=body,
            headers={"Content-Type": "application/json"},
        )
        try:
            urllib.request.urlopen(request, timeout=5)
        except Exception:
            pass
        for _ in range(10):
            time.sleep(1)
            if not server_running():
                break
    with state_lock:
        if server_process is not None and server_process.poll() is None:
            try:
                server_process.terminate()
            except Exception:
                pass
            try:
                server_process.wait(timeout=5)
            except Exception:
                try:
                    server_process.kill()
                except Exception:
                    pass
        server_process = None
        started_at = None
    if server_running():
        return {"ok": False, "error": "服务器未能完全停止，请稍后重试。"}
    log("服务器已停止")
    return {"ok": True, "message": "服务器已安全停止。"}


def _stop_server_on_console_exit() -> None:
    """控制台退出时同步停止云端服务器，避免客户端仍能连接。"""
    try:
        if server_running():
            stop_server()
            log("控制台退出时已停止云端服务器")
    except Exception:
        pass


def status_payload() -> dict:
    config = read_config()
    running = server_running()
    password_set = bool(current_admin_password())
    return {
        "ok": True,
        "running": running,
        "configured": password_set,
        "localIp": local_ip(),
        "serverUrl": f"http://{local_ip()}:{SERVER_PORT}",
        "publicUrl": str(config.get("public_url", "") or ""),
        "adminEmail": str(config.get("security_email", "") or ""),
        "port": SERVER_PORT,
        "consolePort": CONSOLE_PORT,
        "uptime": int(time.time() - started_at) if running and started_at else 0,
        "cpolarAvailable": cpolar_available(),
        "cloudflaredAvailable": cloudflared_available(),
        "tunnelRunning": tunnel_running(),
        "tunnelProvider": tunnel_provider,
        "tunnelUrl": tunnel_url,
        "tunnelError": tunnel_error,
        "linkLocal": "ok" if _link_status.get("local") else ("fail" if _link_status.get("local") is not None else "none"),
        "linkPublic": "ok" if _link_status.get("public") else ("fail" if _link_status.get("public") is not None else "none"),
    }


def tail_log(max_bytes: int = 6000) -> str:
    try:
        content = LOG_PATH.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return ""
    return content[-max_bytes:]


class ConsoleHandler(BaseHTTPRequestHandler):
    def log_message(self, *args) -> None:
        pass

    def _send_json(self, payload: dict, status: int = HTTPStatus.OK) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _send_html(self) -> None:
        try:
            content = HTML_PATH.read_bytes()
        except OSError:
            content = "<h1>控制台页面缺失</h1>".encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(content)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(content)

    def do_GET(self) -> None:
        parsed_path = self.path.split("?", 1)[0]
        if parsed_path in ("/", "/index.html", "/console.html"):
            self._send_html()
        elif parsed_path == "/api/status":
            self._send_json(status_payload())
        elif parsed_path == "/api/log":
            self._send_json({"ok": True, "log": tail_log()})
        elif parsed_path == "/api/conflicts":
            self._send_json({"ok": True, "servers": find_other_servers()})
        elif parsed_path == "/api/users":
            self._send_json({"ok": True, "users": list_users()})
        elif parsed_path == "/api/get-password":
            self._send_json({"ok": True, "password": current_admin_password()})
        elif parsed_path == "/api/tunnel-status":
            self._send_json(
                {
                    "ok": True,
                    "available": cpolar_available(),
                    "cloudflaredAvailable": cloudflared_available(),
                    "running": tunnel_running(),
                    "provider": tunnel_provider,
                    "url": tunnel_url,
                    "tunnelError": tunnel_error,
                }
            )
        elif parsed_path == "/api/backup-path":
            self._send_json({"ok": True, "path": str(DATA_DIR / "backups")})
        elif parsed_path == "/api/token-history":
            self._send_json(token_history_payload())
        else:
            self._send_json({"ok": False, "error": "未知接口"}, HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        length = int(self.headers.get("Content-Length", 0) or 0)
        raw = self.rfile.read(length) if length else b""
        try:
            data = json.loads(raw.decode("utf-8")) if raw else {}
        except json.JSONDecodeError:
            data = {}
        parsed_path = self.path.split("?", 1)[0]
        if parsed_path == "/api/start":
            self._send_json(start_server())
        elif parsed_path == "/api/stop":
            self._send_json(stop_server())
        elif parsed_path == "/api/save-config":
            password = str(data.get("adminPassword", "")).strip()
            if not password:
                self._send_json({"ok": False, "error": "管理员密码不能为空。"}, HTTPStatus.BAD_REQUEST)
                return
            config = read_config()
            config["admin_password"] = password
            write_config(config)
            # 若数据库已存在，同步更新数据库中的管理员密码
            if DB_PATH.exists():
                try:
                    import sqlite3

                    connection = sqlite3.connect(str(DB_PATH))
                    connection.execute(
                        "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
                        ("admin_password", password),
                    )
                    connection.commit()
                    connection.close()
                except Exception:
                    pass
            log("控制台更新了管理员密码")
            self._send_json({"ok": True, "message": "密码已保存。"})
        elif parsed_path == "/api/set-password":
            password = str(data.get("adminPassword", "")).strip()
            if not password:
                self._send_json({"ok": False, "error": "管理员密码不能为空。"}, HTTPStatus.BAD_REQUEST)
                return
            config = read_config()
            config["admin_password"] = password
            write_config(config)
            if DB_PATH.exists():
                try:
                    import sqlite3

                    connection = sqlite3.connect(str(DB_PATH))
                    connection.execute(
                        "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
                        ("admin_password", password),
                    )
                    connection.commit()
                    connection.close()
                except Exception:
                    pass
            log("控制台修改了管理员密码")
            self._send_json({"ok": True, "message": "管理员密码已修改。"})
        elif parsed_path == "/api/add-user":
            password = str(data.get("password", "")).strip()
            note = str(data.get("note", "")).strip()[:40]
            if not password:
                self._send_json({"ok": False, "error": "密码不能为空。"}, HTTPStatus.BAD_REQUEST)
                return
            if not password.isdigit():
                self._send_json({"ok": False, "error": "密码需为纯数字。"}, HTTPStatus.BAD_REQUEST)
                return
            if DB_PATH.exists():
                try:
                    import sqlite3

                    connection = sqlite3.connect(str(DB_PATH))
                    existing = connection.execute(
                        "SELECT 1 FROM users WHERE password = ?", (password,)
                    ).fetchone()
                    if existing:
                        connection.close()
                        self._send_json({"ok": False, "error": "该密码已存在。"}, HTTPStatus.BAD_REQUEST)
                        return
                    connection.execute(
                        "INSERT INTO users (password, created_at, note) VALUES (?, ?, ?)",
                        (password, time.time(), note),
                    )
                    connection.commit()
                    connection.close()
                    log(f"控制台新增普通用户：{note or password}")
                    self._send_json({"ok": True, "message": "用户已添加。"})
                except Exception as exc:
                    self._send_json({"ok": False, "error": f"添加失败：{exc}"})
            else:
                self._send_json({"ok": False, "error": "数据库尚未初始化，请先启动一次服务器。"}, HTTPStatus.BAD_REQUEST)
        elif parsed_path == "/api/delete-user":
            user_id = str(data.get("id", "") or "").strip()
            if not user_id:
                self._send_json({"ok": False, "error": "缺少用户标识。"}, HTTPStatus.BAD_REQUEST)
                return
            if DB_PATH.exists():
                try:
                    import sqlite3

                    connection = sqlite3.connect(str(DB_PATH))
                    cursor = connection.execute(
                        "DELETE FROM users WHERE password = ?", (user_id,)
                    )
                    connection.commit()
                    connection.close()
                    if cursor.rowcount == 0:
                        self._send_json({"ok": False, "error": "该用户不存在。"}, HTTPStatus.BAD_REQUEST)
                        return
                    log(f"控制台删除普通用户：{password}")
                    self._send_json({"ok": True, "message": "用户已删除。"})
                except Exception as exc:
                    self._send_json({"ok": False, "error": f"删除失败：{exc}"})
            else:
                self._send_json({"ok": False, "error": "数据库尚未初始化，请先启动一次服务器。"}, HTTPStatus.BAD_REQUEST)
        elif parsed_path == "/api/open-data-dir":
            # 打开存放 Excel 表格的文件夹；版本目录为空时打开 data 根目录
            target_dir = SERVER_DIR / "data" / "versions"
            if not target_dir.exists() or not any(target_dir.iterdir()):
                target_dir = SERVER_DIR / "data"
            try:
                target_dir.mkdir(parents=True, exist_ok=True)
                os.startfile(str(target_dir))
                self._send_json({"ok": True, "message": "已打开数据文件夹。"})
            except Exception as exc:
                self._send_json({"ok": False, "error": f"打开失败：{exc}"})
        elif parsed_path == "/api/save-tunnel":
            public_url = str(data.get("publicUrl", "")).strip().rstrip("/")
            if not public_url:
                self._send_json({"ok": False, "error": "公网地址不能为空。"}, HTTPStatus.BAD_REQUEST)
                return
            config = read_config()
            config["public_url"] = public_url
            write_config(config)
            log(f"保存公网地址：{public_url}")
            self._send_json({"ok": True, "message": "公网地址已保存。"})
        elif parsed_path == "/api/backup-data":
            self._send_json(backup_data(str(data.get("saveDir", "") or "")))
        elif parsed_path == "/api/restore-data":
            import base64 as _b64
            raw_bytes = b""
            try:
                raw_bytes = _b64.b64decode(str(data.get("data") or ""))
            except Exception:
                raw_bytes = b""
            self._send_json(restore_data(raw_bytes))
        elif parsed_path == "/api/tunnel-start":
            provider = str(data.get("provider", "cpolar")).strip()
            authtoken = str(data.get("authtoken", "")).strip()
            self._send_json(start_tunnel(provider, authtoken))
        elif parsed_path == "/api/token-save":
            authtoken = str(data.get("authtoken", "")).strip()
            note = str(data.get("note", "")).strip()
            if not authtoken:
                self._send_json({"ok": False, "error": "令牌不能为空。"}, HTTPStatus.BAD_REQUEST)
                return
            add_token_to_history(authtoken, note)
            log("控制台保存了一个 cpolar 历史令牌")
            self._send_json({"ok": True, "message": "已保存到历史令牌。", **token_history_payload()})
        elif parsed_path == "/api/token-delete":
            token_id = str(data.get("id", "")).strip()
            delete_token_from_history(token_id)
            self._send_json({"ok": True, "message": "已删除。", **token_history_payload()})
        elif parsed_path == "/api/tunnel-stop":
            self._send_json(stop_tunnel())
        elif parsed_path == "/api/tunnel-test":
            public_url = str(data.get("publicUrl", "")).strip().rstrip("/")
            if not public_url:
                self._send_json({"ok": False, "error": "请先输入公网地址。"}, HTTPStatus.BAD_REQUEST)
                return
            target = public_url if public_url.startswith("http") else f"http://{public_url}"
            ok = http_ok(f"{target}/api/health", timeout=6.0)
            if ok:
                self._send_json({"ok": True, "message": "连通正常！"})
            else:
                # 区分原因：服务器未启动（本机 18666 无服务）或地址本身不可达
                local_ok = server_running()
                hint = (
                    "连通失败：请先点「启动服务器」再测试（穿透地址正常，但 18666 端口还没有服务）。"
                    if not local_ok
                    else "连通失败：穿透地址无法访问，请检查穿透是否开启、地址是否最新。"
                )
                self._send_json({"ok": False, "message": hint})
        elif parsed_path == "/api/quit":
            def stopper() -> None:
                time.sleep(0.5)
                # 退出控制台时同时停止云端服务器，避免客户端仍能连接
                try:
                    stop_server()
                except Exception:
                    pass
                os._exit(0)

            self._send_json({"ok": True, "message": "控制台即将退出（云端服务器将同时停止）。"})
            threading.Thread(target=stopper, daemon=True).start()
        else:
            self._send_json({"ok": False, "error": "未知接口"}, HTTPStatus.NOT_FOUND)


class ConsoleHTTPServer(ThreadingHTTPServer):
    allow_reuse_address = False  # 禁止多实例共享端口，避免多个控制台混在一起


def main() -> None:
    gui = "--browser" not in sys.argv[1:]
    try:
        server = ConsoleHTTPServer(("127.0.0.1", CONSOLE_PORT), ConsoleHandler)
    except OSError:
        log(f"控制台端口 {CONSOLE_PORT} 被占用，可能已有一个控制台在运行。")
        print("服务器控制台已在运行，请查看右下角托盘图标。", flush=True)
        return
    log(f"服务器控制台已启动：http://127.0.0.1:{CONSOLE_PORT}")
    threading.Thread(target=_check_links_loop, daemon=True).start()
    server_thread = threading.Thread(target=server.serve_forever, daemon=True)
    server_thread.start()
    url = f"http://127.0.0.1:{CONSOLE_PORT}"
    icon = str(BASE / "erl_server.ico") if (BASE / "erl_server.ico").exists() else None
    if gui:
        used_gui = run_gui(
            url,
            "ERL库存管理 服务器控制台",
            icon=icon,
            width=1080,
            height=880,
            on_quit=lambda: _stop_server_on_console_exit(),
        )
        if not used_gui:
            # WebView2 缺失/失败已由 gui_window 提示并自动安装，这里不再重复处理
            server.server_close()
            return
    else:
        threading.Timer(1.2, lambda: webbrowser.open(url, new=1)).start()
        server_thread.join()
    try:
        server.server_close()
    except Exception:
        pass
    log("控制台已退出（服务器运行状态不受影响）。")


if __name__ == "__main__":
    main()
