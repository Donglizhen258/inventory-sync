from __future__ import annotations

import hashlib
import hmac
import json
import os
import secrets
import socket
import sqlite3
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Any

import uvicorn
from fastapi import Depends, FastAPI, File, Header, HTTPException, UploadFile, Request
from fastapi.responses import FileResponse
from pydantic import BaseModel

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
VERSIONS_DIR = DATA_DIR / "versions"
DATA_DIR.mkdir(parents=True, exist_ok=True)
VERSIONS_DIR.mkdir(parents=True, exist_ok=True)

CONFIG_PATH = BASE_DIR / "config.json"

DEFAULT_CONFIG: dict[str, Any] = {
    "admin_password": "admin123",
    "token_ttl_hours": 12,
    "version_keep_days": 14,
    "max_version_bytes": 100 * 1024 * 1024,
    "port": 18666,
}

CONFIG: dict[str, Any] = DEFAULT_CONFIG.copy()
if CONFIG_PATH.is_file():
    try:
        CONFIG.update(json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
    except (json.JSONDecodeError, OSError):
        pass

TOKENS: dict[str, tuple[float, str]] = {}
TOKEN_LOCK = threading.Lock()
DB_LOCK = threading.Lock()

DB_PATH = DATA_DIR / "meta.db"


def sha256_hex(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


PWD_PREFIX = "enc:"  # 加密密码存储前缀


def _pwd_key() -> bytes:
    """获取密码加密密钥（存 config.json，随备份整体迁移）。"""
    key = str(CONFIG.get("pwd_key") or "")
    if not key:
        key = secrets.token_hex(32)
        CONFIG["pwd_key"] = key
        try:
            CONFIG_PATH.write_text(
                json.dumps(CONFIG, ensure_ascii=False, indent=2), encoding="utf-8"
            )
        except Exception:
            pass
    return bytes.fromhex(key)


def _xor_bytes(data: bytes, key: bytes) -> bytes:
    return bytes(value ^ key[index % len(key)] for index, value in enumerate(data))


def hash_password(password: str) -> str:
    """对称加密存储（标准库 XOR + base64）：管理员可解密查看，文件泄露无密钥读不出。"""
    import base64
    raw = _xor_bytes(password.encode("utf-8"), _pwd_key())
    return PWD_PREFIX + base64.b64encode(raw).decode()


def decrypt_password(stored: str) -> str:
    """解密存储的密码；无法解密返回空字符串。"""
    import base64
    stored = (stored or "").strip()
    if not stored.startswith(PWD_PREFIX):
        return ""
    try:
        raw = base64.b64decode(stored[len(PWD_PREFIX):])
        return _xor_bytes(raw, _pwd_key()).decode("utf-8")
    except Exception:
        return ""


def verify_password(stored: str, password: str) -> bool:
    """校验密码：兼容对称加密(Fernet)、加盐格式(salt$hash)与旧版无盐 sha256/明文。"""
    stored = (stored or "").strip()
    if not stored:
        return False
    if stored.startswith(PWD_PREFIX):
        return hmac.compare_digest(decrypt_password(stored), password)
    if "$" in stored:
        salt, _, digest = stored.partition("$")
        return hmac.compare_digest(
            hashlib.sha256((salt + password).encode("utf-8")).hexdigest(), digest
        )
    return hmac.compare_digest(stored, password) or hmac.compare_digest(
        sha256_hex(password), stored
    )


def db_connect() -> sqlite3.Connection:
    connection = sqlite3.connect(DB_PATH)
    connection.execute("PRAGMA journal_mode=WAL")
    return connection


def get_setting(connection: sqlite3.Connection, key: str) -> str | None:
    row = connection.execute(
        "SELECT value FROM settings WHERE key = ?",
        (key,),
    ).fetchone()
    return row[0] if row else None


def set_setting(connection: sqlite3.Connection, key: str, value: str) -> None:
    connection.execute(
        "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
        (key, value),
    )


def init_db() -> None:
    with db_connect() as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS versions (
                version_id TEXT PRIMARY KEY,
                file_name TEXT NOT NULL,
                size INTEGER NOT NULL,
                md5 TEXT NOT NULL,
                uploaded_at REAL NOT NULL,
                uploader TEXT NOT NULL
            )
            """
        )
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                password TEXT PRIMARY KEY,
                created_at REAL NOT NULL,
                note TEXT NOT NULL DEFAULT ''
            )
            """
        )
        if get_setting(connection, "admin_password") is None:
            set_setting(
                connection,
                "admin_password",
                str(CONFIG["admin_password"]),
            )
        if get_setting(connection, "admin_password_hash") is None:
            set_setting(
                connection,
                "admin_password_hash",
                sha256_hex(str(CONFIG["admin_password"])),
            )


def admin_password(connection: sqlite3.Connection) -> str:
    return get_setting(connection, "admin_password") or str(CONFIG["admin_password"])


def admin_password_hash(connection: sqlite3.Connection) -> str:
    return get_setting(
        connection,
        "admin_password_hash",
    ) or sha256_hex(admin_password(connection))


def password_matches_admin(password: str) -> bool:
    with DB_LOCK:
        with db_connect() as connection:
            stored = admin_password_hash(connection)
    return hmac.compare_digest(sha256_hex(password), stored)


def find_user_role(password: str) -> str | None:
    """Return 'admin' / 'user' if the password is valid, else None."""
    if password_matches_admin(password):
        return "admin"
    with DB_LOCK:
        with db_connect() as connection:
            rows = connection.execute("SELECT password FROM users").fetchall()
            for (stored,) in rows:
                if stored and verify_password(stored, password):
                    # 旧格式（明文/无盐哈希/加盐哈希）自动升级为对称加密
                    if not stored.startswith(PWD_PREFIX):
                        connection.execute(
                            "UPDATE users SET password = ? WHERE password = ?",
                            (hash_password(password), stored),
                        )
                        connection.commit()
                    return "user"
    return None


def issue_token(role: str) -> str:
    token = secrets.token_hex(24)
    ttl = float(CONFIG["token_ttl_hours"]) * 3600
    with TOKEN_LOCK:
        now = time.time()
        for existing in list(TOKENS):
            if TOKENS[existing][0] < now:
                TOKENS.pop(existing, None)
        TOKENS[token] = (now + ttl, role)
    return token


def bearer_role(authorization: str | None) -> str | None:
    if not authorization or not authorization.startswith("Bearer "):
        return None
    token = authorization[7:].strip()
    with TOKEN_LOCK:
        entry = TOKENS.get(token)
        if entry is None:
            return None
        expiry, role = entry
        if expiry < time.time():
            TOKENS.pop(token, None)
            return None
        return role


def require_login(authorization: str | None = Header(default=None)) -> str:
    role = bearer_role(authorization)
    if role is None:
        raise HTTPException(status_code=401, detail="请先登录。")
    return role


def require_admin(role: str = Depends(require_login)) -> str:
    if role != "admin":
        raise HTTPException(status_code=403, detail="该操作仅限管理员。")
    return role


def clean_expired_versions(now: float | None = None) -> int:
    if now is None:
        now = time.time()
    cutoff = now - float(CONFIG["version_keep_days"]) * 86400
    removed = 0
    with DB_LOCK:
        with db_connect() as connection:
            rows = connection.execute(
                "SELECT version_id FROM versions WHERE uploaded_at < ?",
                (cutoff,),
            ).fetchall()
            for (version_id,) in rows:
                target = VERSIONS_DIR / f"{version_id}.xlsx"
                try:
                    target.unlink(missing_ok=True)
                except OSError:
                    pass
                connection.execute(
                    "DELETE FROM versions WHERE version_id = ?",
                    (version_id,),
                )
                removed += 1
    return removed


def version_dict(row: tuple[Any, ...]) -> dict[str, Any]:
    version_id, file_name, size, md5, uploaded_at, uploader = row
    return {
        "versionId": version_id,
        "fileName": file_name,
        "size": size,
        "md5": md5,
        "uploadedAt": datetime.fromtimestamp(uploaded_at).strftime("%Y-%m-%d %H:%M:%S"),
        "uploader": uploader,
        "downloadUrl": f"/api/download/{version_id}",
    }


app = FastAPI(title="ERL库存云同步服务", version="1.1.0")


class LoginBody(BaseModel):
    password: str


class AddUserBody(BaseModel):
    password: str
    note: str = ""


class RemoveUserBody(BaseModel):
    password: str


class ChangeAdminPasswordBody(BaseModel):
    old_password: str
    new_password: str




@app.on_event("startup")
def startup() -> None:
    init_db()
    clean_expired_versions()


@app.get("/api/health")
def health() -> dict[str, Any]:
    return {"ok": True, "service": "ERL库存云同步", "time": datetime.now().isoformat()}


@app.post("/api/login")
def login(body: LoginBody) -> dict[str, Any]:
    role = find_user_role(body.password)
    if role is None:
        raise HTTPException(status_code=401, detail="密码错误。")
    payload: dict[str, Any] = {
        "ok": True,
        "token": issue_token(role),
        "role": role,
    }
    if role == "admin":
        with DB_LOCK:
            with db_connect() as connection:
                payload["adminPasswordHash"] = admin_password_hash(connection)
    return payload


@app.get("/api/state")
def state(role: str = Depends(require_login)) -> dict[str, Any]:
    clean_expired_versions()
    with DB_LOCK:
        with db_connect() as connection:
            rows = connection.execute(
                "SELECT version_id, file_name, size, md5, uploaded_at, uploader "
                "FROM versions ORDER BY uploaded_at DESC LIMIT 50"
            ).fetchall()
    versions = [version_dict(row) for row in rows]
    public_url = ""
    try:
        with open(CONFIG_PATH, encoding="utf-8") as handle:
            public_url = str(json.load(handle).get("public_url", "") or "")
    except Exception:
        pass
    return {
        "ok": True,
        "latest": versions[0] if versions else None,
        "versions": versions,
        "public_url": public_url,
    }


@app.post("/api/upload")
def upload(
    role: str = Depends(require_login),
    file: UploadFile = File(...),
) -> dict[str, Any]:
    raw = file.file.read()
    if not raw:
        raise HTTPException(status_code=400, detail="上传内容为空。")
    if len(raw) > int(CONFIG["max_version_bytes"]):
        raise HTTPException(status_code=413, detail="文件超过大小限制。")

    md5 = hashlib.md5(raw).hexdigest()
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    version_id = f"ERL_存储-{stamp}-{md5[:8]}"
    target = VERSIONS_DIR / f"{version_id}.xlsx"
    if not target.exists():
        target.write_bytes(raw)

    uploader = role
    with DB_LOCK:
        with db_connect() as connection:
            existing = connection.execute(
                "SELECT uploader FROM versions WHERE version_id = ?",
                (version_id,),
            ).fetchone()
            if existing:
                uploader = existing[0]
            else:
                connection.execute(
                    "INSERT INTO versions (version_id, file_name, size, md5, uploaded_at, uploader) "
                    "VALUES (?, ?, ?, ?, ?, ?)",
                    (
                        version_id,
                        f"ERL_存储-{stamp}.xlsx",
                        len(raw),
                        md5,
                        time.time(),
                        uploader,
                    ),
                )

    clean_expired_versions()
    return {"ok": True, "versionId": version_id, "md5": md5, "size": len(raw)}


@app.get("/api/latest-file")
def latest_file(role: str = Depends(require_login)) -> FileResponse:
    """所有登录用户可下载最新版本文件（日常同步用）。"""
    with DB_LOCK:
        with db_connect() as connection:
            row = connection.execute(
                "SELECT version_id, file_name FROM versions "
                "ORDER BY uploaded_at DESC LIMIT 1"
            ).fetchone()
    if not row:
        raise HTTPException(status_code=404, detail="云端还没有任何版本。")
    version_id, file_name = row
    target = VERSIONS_DIR / f"{version_id}.xlsx"
    if not target.is_file():
        raise HTTPException(status_code=404, detail="版本文件缺失。")
    return FileResponse(
        target,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=file_name,
    )


@app.get("/api/download/{version_id}")
def download(version_id: str, role: str = Depends(require_admin)) -> FileResponse:
    """历史版本下载（回溯）仅限管理员。"""
    if not all(char.isalnum() or char in "-_" for char in version_id):
        raise HTTPException(status_code=404, detail="版本不存在。")
    target = VERSIONS_DIR / f"{version_id}.xlsx"
    if not target.is_file():
        raise HTTPException(status_code=404, detail="版本不存在。")
    with DB_LOCK:
        with db_connect() as connection:
            row = connection.execute(
                "SELECT file_name FROM versions WHERE version_id = ?",
                (version_id,),
            ).fetchone()
    file_name = row[0] if row else "实验室库存管理.xlsx"
    return FileResponse(
        target,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=file_name,
    )


@app.get("/api/users")
def list_users(role: str = Depends(require_admin)) -> dict[str, Any]:
    with DB_LOCK:
        with db_connect() as connection:
            rows = connection.execute(
                "SELECT password, created_at, note FROM users ORDER BY created_at"
            ).fetchall()
    return {
        "ok": True,
        "users": [
            {
                "id": row[0],
                "password": decrypt_password(row[0]) or row[0],
                "masked": (row[0][:6] + "…" + row[0][-4:]) if len(row[0]) > 12 else "已加密",
                "createdAt": datetime.fromtimestamp(row[1]).strftime("%Y-%m-%d %H:%M:%S"),
                "note": row[2],
            }
            for row in rows
        ],
    }


@app.post("/api/users")
def add_user(body: AddUserBody, role: str = Depends(require_admin)) -> dict[str, Any]:
    password = body.password.strip()
    if not password:
        raise HTTPException(status_code=400, detail="数字密码不能为空。")
    if not password.isdigit():
        raise HTTPException(status_code=400, detail="普通用户密码必须为数字。")
    if password_matches_admin(password):
        raise HTTPException(status_code=400, detail="该密码与管理员密码相同。")
    stored = hash_password(password)
    with DB_LOCK:
        with db_connect() as connection:
            connection.execute(
                "INSERT OR IGNORE INTO users (password, created_at, note) VALUES (?, ?, ?)",
                (stored, time.time(), body.note.strip()),
            )
    return {"ok": True, "password": password}


@app.delete("/api/users")
def remove_user(body: RemoveUserBody, role: str = Depends(require_admin)) -> dict[str, Any]:
    user_id = (body.password or "").strip()
    if not user_id:
        raise HTTPException(status_code=400, detail="缺少用户标识。")
    with DB_LOCK:
        with db_connect() as connection:
            cursor = connection.execute(
                "DELETE FROM users WHERE password = ?",
                (user_id,),
            )
    if cursor.rowcount == 0:
        raise HTTPException(status_code=404, detail="该用户不存在。")
    return {"ok": True}


class ShutdownBody(BaseModel):
    password: str


PID_FILE = BASE_DIR / "server.pid"

UDP_HEARTBEAT_PORT = 29998
UDP_DISCOVER_PORT = 29997


def _local_ip() -> str:
    try:
        probe = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        probe.connect(("8.8.8.8", 80))
        address = probe.getsockname()[0]
        probe.close()
        return address
    except OSError:
        return "127.0.0.1"


def udp_heartbeat_loop() -> None:
    """周期性向局域网广播自身存在，供多服务器冲突检测。"""
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
    payload = f"ERL-SERVER|{socket.gethostname()}".encode("utf-8")
    while True:
        try:
            sock.sendto(payload, ("255.255.255.255", UDP_HEARTBEAT_PORT))
        except OSError:
            pass
        time.sleep(3)


def udp_discover_loop() -> None:
    """监听发现探测包并回复，供客户端/启动器自动发现本服务器。"""
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    try:
        sock.bind(("0.0.0.0", UDP_DISCOVER_PORT))
    except OSError:
        return
    reply = f"ERL-SERVER|{socket.gethostname()}|{_local_ip()}".encode("utf-8")
    while True:
        try:
            data, peer = sock.recvfrom(1024)
            if data.strip().upper() == b"ERL-DISCOVER":
                sock.sendto(reply, peer)
        except OSError:
            return


@app.post("/api/shutdown")
def shutdown_server(
    body: ShutdownBody,
    request: Request,
) -> dict[str, Any]:
    """仅允许本机调用，需管理员密码，用于傻瓜式退出。"""
    host = getattr(request.client, "host", "")
    if host not in ("127.0.0.1", "::1"):
        raise HTTPException(status_code=403, detail="仅允许本机停止服务器。")
    if not password_matches_admin(body.password):
        raise HTTPException(status_code=400, detail="管理员密码错误。")

    def stopper() -> None:
        try:
            os.remove(PID_FILE)
        except OSError:
            pass
        print("服务器已安全停止。", flush=True)
        time.sleep(0.5)
        os._exit(0)

    threading.Thread(target=stopper, daemon=True).start()
    return {"ok": True, "message": "服务器即将停止。"}


def main() -> None:
    init_db()
    clean_expired_versions()
    try:
        PID_FILE.write_text(str(os.getpid()), encoding="utf-8")
    except OSError:
        pass
    threading.Thread(target=udp_heartbeat_loop, daemon=True).start()
    threading.Thread(target=udp_discover_loop, daemon=True).start()
    print(
        f"ERL 库存云同步服务已启动：http://0.0.0.0:{int(CONFIG['port'])}，"
        "按客户端或启动器提示可安全停止。",
        flush=True,
    )
    try:
        uvicorn.run(
            app,
            host="0.0.0.0",
            port=int(CONFIG["port"]),
            log_level="info",
        )
    finally:
        try:
            PID_FILE.unlink(missing_ok=True)
        except OSError:
            pass


if __name__ == "__main__":
    main()
