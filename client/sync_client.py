from __future__ import annotations

import hashlib
import json
import socket
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

DEFAULT_SERVER_URL = "http://127.0.0.1:18666"
UDP_DISCOVER_PORT = 29997
SERVER_HTTP_PORT = 18666


def discover_servers(timeout: float = 2.0) -> list[dict[str, str]]:
    """通过 UDP 广播自动发现局域网内的 ERL 库存服务器。"""
    results: list[dict[str, str]] = []
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
    sock.settimeout(timeout)
    try:
        sock.sendto(b"ERL-DISCOVER", ("255.255.255.255", UDP_DISCOVER_PORT))
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            try:
                data, peer = sock.recvfrom(1024)
            except socket.timeout:
                break
            text = data.decode("utf-8", "replace")
            if not text.startswith("ERL-SERVER|"):
                continue
            parts = text.split("|")
            hostname = parts[1] if len(parts) > 1 else ""
            ip = parts[2] if len(parts) > 2 else peer[0]
            url = f"http://{ip}:{SERVER_HTTP_PORT}"
            if not any(item["url"] == url for item in results):
                results.append({"url": url, "ip": ip, "hostname": hostname})
    finally:
        sock.close()
    return results


class SyncError(Exception):
    """云端同步错误，message 可直接展示给用户。"""


class SyncClient:
    """轻量云端同步客户端，仅依赖标准库。"""

    def __init__(
        self,
        server_url: str = DEFAULT_SERVER_URL,
        token: str = "",
        timeout: float = 8.0,
    ) -> None:
        self.server_url = (server_url or DEFAULT_SERVER_URL).rstrip("/")
        self.token = token
        self.timeout = timeout

    def _request(
        self,
        path: str,
        *,
        method: str = "GET",
        json_body: dict[str, Any] | None = None,
        raw_body: bytes | None = None,
        content_type: str | None = None,
        timeout: float | None = None,
    ) -> tuple[int, bytes]:
        url = urllib.parse.quote(
            f"{self.server_url}{path}", safe=":/?#&=%+-_.@"
        )
        headers: dict[str, str] = {}
        data: bytes | None = None
        if json_body is not None:
            data = json.dumps(json_body, ensure_ascii=False).encode("utf-8")
            content_type = "application/json; charset=utf-8"
        elif raw_body is not None:
            data = raw_body
        if content_type:
            headers["Content-Type"] = content_type
        if self.token:
            headers["Authorization"] = f"Bearer {self.token}"
        request = urllib.request.Request(url, method=method, data=data, headers=headers)
        try:
            with urllib.request.urlopen(request, timeout=timeout or self.timeout) as response:
                return response.status, response.read()
        except urllib.error.HTTPError as error:
            body = error.read()
            message = _error_message(body) or f"云端返回错误（{error.code}）。"
            raise SyncError(message) from error
        except urllib.error.URLError as error:
            raise SyncError(_network_message(error)) from error
        except TimeoutError as error:
            raise SyncError("连接云端超时，请检查网络。") from error
        except OSError as error:
            raise SyncError(f"无法连接云端：{error}") from error

    def health(self) -> bool:
        try:
            status, body = self._request("/api/health", timeout=3.0)
            return status == 200 and json.loads(body.decode("utf-8")).get("ok") is True
        except (SyncError, ValueError, UnicodeDecodeError):
            return False

    def login(self, password: str) -> tuple[str, str]:
        """云端登录，返回 (token, role)。"""
        status, body = self._request(
            "/api/login",
            method="POST",
            json_body={"password": password},
        )
        payload = json.loads(body.decode("utf-8"))
        token = str(payload.get("token") or "").strip()
        role = str(payload.get("role") or "user").strip()
        if not token:
            raise SyncError("云端登录未返回有效凭证。")
        self.token = token
        return token, role

    def fetch_state(self) -> dict[str, Any]:
        status, body = self._request("/api/state")
        payload = json.loads(body.decode("utf-8"))
        if not payload.get("ok"):
            raise SyncError(str(payload.get("detail") or "云端状态读取失败。"))
        return payload

    def upload(self, file_path: Path) -> dict[str, Any]:
        file_path = Path(file_path)
        if not file_path.is_file():
            raise SyncError("待上传的库存文件不存在。")
        raw = file_path.read_bytes()
        boundary = uuid.uuid4().hex
        file_name = file_path.name
        parts = [
            (
                f"--{boundary}\r\n"
                f'Content-Disposition: form-data; name="file"; '
                f'filename="{file_name}"\r\n'
                f"Content-Type: application/octet-stream\r\n\r\n"
            ).encode("utf-8"),
            raw,
            f"\r\n--{boundary}--\r\n".encode("utf-8"),
        ]
        status, body = self._request(
            "/api/upload",
            method="POST",
            raw_body=b"".join(parts),
            content_type=f"multipart/form-data; boundary={boundary}",
        )
        payload = json.loads(body.decode("utf-8"))
        if not payload.get("ok"):
            raise SyncError(str(payload.get("detail") or "版本上传失败。"))
        return payload

    def download(self, version_id: str, target_path: Path) -> Path:
        """下载指定历史版本（回溯），云端仅允许管理员调用。"""
        target_path = Path(target_path)
        safe = all(char.isalnum() or char in "-_" for char in version_id)
        if not safe:
            raise SyncError("版本标识不合法。")
        status, body = self._request(f"/api/download/{version_id}")
        target_path.parent.mkdir(parents=True, exist_ok=True)
        target_path.write_bytes(body)
        return target_path

    def download_latest(self, target_path: Path) -> dict[str, Any]:
        """下载云端最新版本文件（日常同步，所有登录用户可用）。"""
        target_path = Path(target_path)
        status, body = self._request("/api/latest-file")
        target_path.parent.mkdir(parents=True, exist_ok=True)
        target_path.write_bytes(body)
        return {"size": len(body)}

    def list_versions(self) -> list[dict[str, Any]]:
        payload = self.fetch_state()
        return list(payload.get("versions") or [])

    def list_users(self) -> list[dict[str, Any]]:
        status, body = self._request("/api/users")
        payload = json.loads(body.decode("utf-8"))
        if not payload.get("ok"):
            raise SyncError(str(payload.get("detail") or "用户列表读取失败。"))
        return list(payload.get("users") or [])

    def add_user(self, password: str, note: str = "") -> None:
        self._request(
            "/api/users",
            method="POST",
            json_body={"password": password, "note": note},
        )

    def remove_user(self, password: str) -> None:
        self._request(
            "/api/users",
            method="DELETE",
            json_body={"password": password},
        )

    def change_admin_password(self, old_password: str, new_password: str) -> dict[str, Any]:
        status, body = self._request(
            "/api/admin/password",
            method="POST",
            json_body={"old_password": old_password, "new_password": new_password},
        )
        return json.loads(body.decode("utf-8"))

    def get_security_email(self) -> str:
        status, body = self._request("/api/admin/email")
        payload = json.loads(body.decode("utf-8"))
        return str(payload.get("email") or "")

    def update_security_email(self, email: str) -> None:
        self._request(
            "/api/admin/email",
            method="POST",
            json_body={"email": email},
        )

    def forgot_password(self, kind: str) -> str:
        status, body = self._request(
            "/api/forgot-password",
            method="POST",
            json_body={"type": kind},
        )
        payload = json.loads(body.decode("utf-8"))
        return str(payload.get("message") or "操作完成。")


def _error_message(body: bytes) -> str:
    try:
        payload = json.loads(body.decode("utf-8"))
        detail = payload.get("detail")
        if isinstance(detail, str) and detail:
            return detail
    except (ValueError, UnicodeDecodeError):
        pass
    return ""


def _network_message(error: urllib.error.URLError) -> str:
    reason = getattr(error, "reason", None)
    if isinstance(reason, (TimeoutError,)):
        return "连接云端超时，请检查网络。"
    if isinstance(reason, OSError):
        text = str(reason)
        if "timed out" in text.lower():
            return "连接云端超时，请检查网络。"
        if "connection refused" in text.lower():
            return "无法连接云端服务器，请检查服务是否已启动。"
    return f"无法连接云端：{reason or error}"


def version_label(version: dict[str, Any] | None) -> str:
    if not version:
        return "无"
    uploaded = str(version.get("uploadedAt") or "")
    version_id = str(version.get("versionId") or "")
    return f"{uploaded}（{version_id[:18]}…）" if uploaded else version_id
