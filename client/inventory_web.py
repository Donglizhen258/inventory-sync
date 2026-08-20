from __future__ import annotations

import atexit
import ctypes
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
import uuid
import webbrowser
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from inventory_core import (
    APP_DIR,
    CHANGE_TYPES,
    LOCATIONS,
    PROJECTS,
    REQUESTERS,
    SETTINGS_PATH,
    InventoryRecord,
    create_backup,
    create_project,
    current_balance,
    ensure_default_workbook,
    format_number,
    list_project_names,
    load_settings,
    parse_datetime,
    recent_records,
    remove_project,
    run_self_test,
    save_settings,
    search_materials,
    upgrade_workbook_schema,
    validate_project_name,
    validate_record,
    validate_workbook,
    workbook_suggestions,
    write_record,
)
from sync_client import (
    DEFAULT_SERVER_URL,
    SyncClient,
    SyncError,
    discover_servers,
)
from gui_window import run_gui


APP_NAME = "ERL库存管理"
RUNTIME_PATH = APP_DIR / ".inventory-runtime.json"
SYNC_STATE_PATH = APP_DIR / "sync.json"


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def desktop_folder() -> Path:
    """获取用户真实的桌面文件夹（跟随系统/OneDrive 重定向）。"""
    try:
        import uuid as _uuid

        shell32 = ctypes.windll.shell32

        class GUID(ctypes.Structure):
            _fields_ = [
                ("Data1", ctypes.c_ulong),
                ("Data2", ctypes.c_ushort),
                ("Data3", ctypes.c_ushort),
                ("Data4", ctypes.c_ubyte * 8),
            ]

        # FOLDERID_Desktop = {B4BFCC3A-DB2C-424C-B029-7FE99A87C641}
        guid = GUID.from_buffer_copy(
            _uuid.UUID("{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}").bytes_le
        )
        shell32.SHGetKnownFolderPath.argtypes = [
            ctypes.POINTER(GUID),
            ctypes.c_uint32,
            ctypes.wintypes.HANDLE,
            ctypes.POINTER(ctypes.c_wchar_p),
        ]
        shell32.SHGetKnownFolderPath.restype = ctypes.c_long
        path_ptr = ctypes.c_wchar_p()
        hr = shell32.SHGetKnownFolderPath(
            ctypes.byref(guid), 0, None, ctypes.byref(path_ptr)
        )
        if hr == 0 and path_ptr.value:
            return Path(path_ptr.value)
    except Exception:
        pass
    # 兜底：USERPROFILE\Desktop
    return Path(os.environ.get("USERPROFILE") or str(Path.home())) / "Desktop"


def choose_save_folder(initial_dir: Path) -> Path | None:
    """弹出系统文件夹选择框，返回用户选择的目录；取消返回 None。"""
    # 首选 tkinter 原生对话框（在 WebView2 环境更稳定）
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        try:
            folder = filedialog.askdirectory(
                initialdir=str(initial_dir),
                title="选择下载保存位置（默认桌面）",
                mustexist=True,
            )
        finally:
            try:
                root.destroy()
            except Exception:
                pass
        if folder:
            return Path(folder)
        return None
    except Exception:
        pass
    # 兜底：Win32 文件夹选择框
    try:
        from ctypes import wintypes as wt

        shell32 = ctypes.windll.shell32
        ole32 = ctypes.windll.ole32

        class BROWSEINFO(ctypes.Structure):
            _fields_ = [
                ("hwndOwner", wt.HWND),
                ("pidlRoot", ctypes.c_void_p),
                ("pszDisplayName", wt.LPWSTR),
                ("lpszTitle", wt.LPCWSTR),
                ("ulFlags", wt.UINT),
                ("lpfn", ctypes.c_void_p),
                ("lParam", ctypes.c_void_p),
                ("iImage", ctypes.c_int),
            ]

        shell32.SHBrowseForFolderW.argtypes = [ctypes.POINTER(BROWSEINFO)]
        shell32.SHBrowseForFolderW.restype = ctypes.c_void_p
        shell32.SHGetPathFromIDListW.argtypes = [ctypes.c_void_p, wt.LPWSTR]
        shell32.SHGetPathFromIDListW.restype = wt.BOOL

        display = ctypes.create_unicode_buffer(512)
        bi = BROWSEINFO()
        bi.hwndOwner = None
        bi.pszDisplayName = display
        bi.lpszTitle = "选择下载保存位置（默认桌面）"
        bi.ulFlags = 0x0001 | 0x0040  # RETURNONLYFSDIRS | NEWDIALOGSTYLE
        pidl = shell32.SHBrowseForFolderW(ctypes.byref(bi))
        if not pidl:
            return None
        path_buf = ctypes.create_unicode_buffer(1024)
        try:
            if shell32.SHGetPathFromIDListW(pidl, path_buf):
                return Path(path_buf.value)
        finally:
            try:
                ole32.CoTaskMemFree(ctypes.c_void_p(pidl))
            except Exception:
                pass
    except Exception:
        pass
    return None


def _write_settings(**fields) -> None:
    """向 settings.json 追加/更新字段（不影响已有字段）。"""
    APP_DIR.mkdir(parents=True, exist_ok=True)
    payload = load_settings()
    payload.update(fields)
    SETTINGS_PATH.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _record_server_history(url: str) -> None:
    """记录一次连接成功的服务器地址（最近优先，最多 10 条）。"""
    history = [item for item in (load_settings().get("server_history") or []) if item != url]
    history.insert(0, url)
    _write_settings(server_history=history[:10])


def _parse_multipart(raw: bytes, content_type: str) -> dict[str, bytes]:
    """解析 multipart/form-data（用于文件上传），返回字段名 -> 内容。"""
    marker = "boundary="
    if marker not in content_type:
        return {}
    boundary = content_type.split(marker, 1)[1].strip().strip('"')
    delimiter = f"--{boundary}".encode("utf-8")
    fields: dict[str, bytes] = {}
    for part in raw.split(delimiter):
        part = part.strip(b"\r\n")
        if not part or part == b"--":
            continue
        header, _, body = part.partition(b"\r\n\r\n")
        name = None
        for line in header.split(b"\r\n"):
            if b"Content-Disposition" in line:
                match = re.search(rb'name="([^"]+)"', line)
                if match:
                    name = match.group(1).decode("utf-8", "replace")
        if name:
            fields[name] = body
    return fields


def resource_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS")) / "web"
    return Path(__file__).resolve().parent / "web"


def native_message(text: str, error: bool = False) -> None:
    flags = 0x10 if error else 0x40
    ctypes.windll.user32.MessageBoxW(None, text, APP_NAME, flags)


class AppState:
    def __init__(self) -> None:
        self.workbook_path = ensure_default_workbook()
        validate_workbook(self.workbook_path)
        self.lock = threading.Lock()
        self.sync_lock = threading.RLock()
        self.public_url = ""
        self.server_url = str(
            load_settings().get("server_url") or DEFAULT_SERVER_URL
        )
        self.sync_token = ""
        self.logged_in = False
        self.online = False
        self.role = "user"
        self.baseline_version: str | None = None
        self.pending_push = False
        self.last_sync_at: str | None = None
        self._load_sync_state()
        self._restore_session()

    def _load_sync_state(self) -> None:
        if not SYNC_STATE_PATH.is_file():
            return
        try:
            payload = json.loads(SYNC_STATE_PATH.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return
        self.server_url = str(
            load_settings().get("server_url")
            or payload.get("server_url")
            or DEFAULT_SERVER_URL
        )
        self.sync_token = str(payload.get("token") or "")
        self.role = str(payload.get("role") or "user")
        self.baseline_version = payload.get("baseline_version")
        self.pending_push = bool(payload.get("pending_push", False))
        self.last_sync_at = payload.get("last_sync_at")

    def _persist_sync_state(self) -> None:
        with self.sync_lock:
            payload = {
                "server_url": self.server_url,
                "token": self.sync_token,
                "role": self.role,
                "baseline_version": self.baseline_version,
                "pending_push": self.pending_push,
                "last_sync_at": self.last_sync_at,
            }
            try:
                APP_DIR.mkdir(parents=True, exist_ok=True)
                SYNC_STATE_PATH.write_text(
                    json.dumps(payload, ensure_ascii=False, indent=2),
                    encoding="utf-8",
                )
            except OSError:
                pass

    def _restore_session(self) -> None:
        if not self.sync_token:
            return
        client = SyncClient(self.server_url, self.sync_token)
        if not client.health():
            self.online = False
            self.logged_in = False
            self._persist_sync_state()
            return
        self.online = True
        try:
            client.fetch_state()
            self.logged_in = True
        except SyncError:
            self.logged_in = False

    def sync_status_payload(self) -> dict[str, Any]:
        """实时探测云端连接：连通则保持在线并刷新公网地址，断开则标记离线。"""
        latest = None
        if self.logged_in and self.server_url:
            try:
                cloud_state = SyncClient(self.server_url, self.sync_token).fetch_state()
                latest = cloud_state.get("latest")
                self.public_url = str(cloud_state.get("public_url") or "")
                self.online = True  # 连上即恢复在线
            except SyncError:
                # 云端不可达：翻转为离线，前端据此提示「云端链接已断开」
                self.online = False
        return {
            "loggedIn": self.logged_in,
            "online": self.online,
            "serverUrl": self.server_url,
            "role": self.role,
            "baselineVersion": self.baseline_version,
            "latestVersion": latest,
            "pendingPush": self.pending_push,
            "lastSyncAt": self.last_sync_at,
            "publicUrl": self.public_url,
        }

    def login(self, password: str) -> dict[str, Any]:
        password = str(password or "").strip()
        if not password:
            raise ValueError("请输入登录密码。")
        client = SyncClient(self.server_url)
        if not client.health():
            raise ValueError("无法连接云端服务器，请检查网络后重试。")
        try:
            self.sync_token, self.role = client.login(password)
        except SyncError as exc:
            raise ValueError(str(exc)) from None
        self.online = True
        self.logged_in = True
        self._persist_sync_state()
        _record_server_history(self.server_url)
        return self.sync_status_payload()

    def pull_latest(self) -> dict[str, Any]:
        if not self.logged_in:
            raise ValueError("请先登录。")
        if not self.online or not self.sync_token:
            raise ValueError("当前为离线模式，无法从云端下载。")
        client = SyncClient(self.server_url, self.sync_token)
        state = client.fetch_state()
        latest = state.get("latest")
        if not latest:
            raise ValueError("云端还没有任何版本。")
        version_id = str(latest["versionId"])
        create_backup(self.workbook_path)
        client.download_latest(self.workbook_path)
        validate_workbook(self.workbook_path)
        self.baseline_version = version_id
        self.pending_push = False
        self.last_sync_at = now_text()
        self._persist_sync_state()
        return self.sync_status_payload()

    def rollback_to(self, version_id: str) -> dict[str, Any]:
        if not self.logged_in:
            raise ValueError("请先登录。")
        if self.role != "admin":
            raise ValueError("表格回溯仅限管理员操作。")
        if not self.online or not self.sync_token:
            raise ValueError("当前为离线模式，无法从云端下载。")
        version_id = str(version_id or "").strip()
        if not version_id:
            raise ValueError("缺少版本标识。")
        client = SyncClient(self.server_url, self.sync_token)
        create_backup(self.workbook_path)
        client.download(version_id, self.workbook_path)
        validate_workbook(self.workbook_path)
        self.baseline_version = version_id
        self.pending_push = False
        self.last_sync_at = now_text()
        self._persist_sync_state()
        return self.sync_status_payload()

    def push_current(self) -> dict[str, Any]:
        if not self.logged_in:
            raise ValueError("请先登录。")
        if not self.online or not self.sync_token:
            raise ValueError("当前为离线模式，无法上传到云端。")
        client = SyncClient(self.server_url, self.sync_token)
        result = client.upload(self.workbook_path)
        self.baseline_version = str(result["versionId"])
        self.pending_push = False
        self.last_sync_at = now_text()
        self._persist_sync_state()
        return self.sync_status_payload()

    def list_users(self) -> list[dict[str, Any]]:
        if self.role != "admin":
            raise ValueError("用户管理仅限管理员操作。")
        client = SyncClient(self.server_url, self.sync_token)
        return client.list_users()

    def add_user(self, password: str, note: str = "") -> None:
        if self.role != "admin":
            raise ValueError("用户管理仅限管理员操作。")
        client = SyncClient(self.server_url, self.sync_token)
        client.add_user(password, note)

    def remove_user(self, password: str) -> None:
        if self.role != "admin":
            raise ValueError("用户管理仅限管理员操作。")
        client = SyncClient(self.server_url, self.sync_token)
        client.remove_user(password)

    def check_conflict(self) -> str | None:
        if self.pending_push:
            return "本地有尚未上传的更改，请先完成同步。"
        if not self.logged_in:
            return "请先登录后再写入。"
        if not self.online or not self.sync_token:
            return "无法连接云端，已禁止写入。请检查网络后重试。"
        try:
            latest = SyncClient(self.server_url, self.sync_token).fetch_state().get("latest")
        except SyncError:
            return "无法连接云端，已禁止写入。请检查网络后重试。"
        if latest and self.baseline_version != str(latest["versionId"]):
            return "云端已被其他同事更新，请先同步最新版本后再操作。"
        return None

    def enqueue_push(self) -> None:
        def worker() -> None:
            try:
                self.push_current()
            except (SyncError, ValueError, OSError):
                with self.sync_lock:
                    self.pending_push = True
                    self._persist_sync_state()
        threading.Thread(target=worker, daemon=True).start()

    def list_projects(self) -> list[str]:
        return list_project_names(self.workbook_path)

    def set_server_url(self, url: str) -> None:
        url = str(url or "").strip().rstrip("/")
        if not url:
            raise ValueError("服务器地址不能为空。")
        if not url.startswith("http://") and not url.startswith("https://"):
            raise ValueError("服务器地址需以 http:// 或 https:// 开头。")
        with self.sync_lock:
            self.server_url = url
            self.sync_token = ""
            self.logged_in = False
            self.online = False
            self.role = "user"
            self.baseline_version = None
            self.pending_push = False
            self.last_sync_at = None
            self._persist_sync_state()
        try:
            settings = load_settings()
            settings["server_url"] = url
            SETTINGS_PATH.write_text(
                json.dumps(settings, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except OSError:
            pass

    def add_project(self, project: str) -> None:
        if self.role != "admin":
            raise ValueError("项目管理仅限管理员操作。")
        create_project(self.workbook_path, project)
        self._push_after_local_change()

    def remove_project(self, project: str) -> None:
        if self.role != "admin":
            raise ValueError("项目管理仅限管理员操作。")
        remove_project(self.workbook_path, project)
        self._push_after_local_change()

    def _push_after_local_change(self) -> None:
        try:
            self.push_current()
        except (SyncError, ValueError, OSError):
            with self.sync_lock:
                self.pending_push = True
                self._persist_sync_state()

    def change_admin_password(self, old_password: str, new_password: str) -> None:
        if self.role != "admin":
            raise ValueError("修改密码仅限管理员操作。")
        SyncClient(self.server_url, self.sync_token).change_admin_password(
            old_password,
            new_password,
        )
        with self.sync_lock:
            self.logged_in = False
            self.online = False
            self.sync_token = ""
            self._persist_sync_state()

    def get_security_email(self) -> str:
        if self.role != "admin":
            raise ValueError("密保邮箱管理仅限管理员操作。")
        return SyncClient(self.server_url, self.sync_token).get_security_email()

    def update_security_email(self, email: str) -> None:
        if self.role != "admin":
            raise ValueError("密保邮箱管理仅限管理员操作。")
        SyncClient(self.server_url, self.sync_token).update_security_email(email)

    def forgot_password(self, kind: str) -> str:
        if kind not in ("admin", "user"):
            raise ValueError("未知的找回类型。")
        return SyncClient(self.server_url).forgot_password(kind)

    def download_excel_bytes(self) -> tuple[bytes, str]:
        if not self.logged_in:
            raise ValueError("请先登录。")
        return self.workbook_path.read_bytes(), self.workbook_path.name

    def desktop_download(self, save_dir: str = "") -> str:
        """下载 Excel：保存到指定目录（前端确认），默认真实桌面。"""
        if not self.logged_in:
            raise ValueError("请先登录。")
        save_path = Path(str(save_dir).strip()) if str(save_dir).strip() else desktop_folder()
        try:
            save_path.mkdir(parents=True, exist_ok=True)
        except OSError:
            save_path = desktop_folder()
        content = self.workbook_path.read_bytes()
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target = save_path / f"ERL_存储-{stamp}.xlsx"
        target.write_bytes(content)
        return str(target)

    def download_version_to_desktop(self, version_id: str, save_dir: str = "") -> str:
        """下载云端指定历史版本到指定目录（前端确认），默认真实桌面。"""
        if not self.logged_in:
            raise ValueError("请先登录。")
        if not self.online or not self.sync_token:
            raise ValueError("未连接云端，无法下载历史版本。")
        save_path = Path(str(save_dir).strip()) if str(save_dir).strip() else desktop_folder()
        try:
            save_path.mkdir(parents=True, exist_ok=True)
        except OSError:
            save_path = desktop_folder()
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target = save_path / f"ERL_存储-历史-{stamp}.xlsx"
        SyncClient(self.server_url, self.sync_token).download(version_id, target)
        return str(target)

    def replace_workbook(self, file_bytes: bytes) -> dict[str, Any]:
        """管理员替换库存表：本地校验格式 → 替换本机副本 → 同步云端。"""
        if not self.logged_in:
            raise ValueError("请先登录。")
        if self.role != "admin":
            raise ValueError("仅管理员可替换库存表。")
        temp = APP_DIR / f".replace-{uuid.uuid4().hex}.xlsx"
        try:
            temp.write_bytes(file_bytes)
            validate_workbook(temp)
        except Exception as exc:
            try:
                temp.unlink(missing_ok=True)
            except Exception:
                pass
            raise ValueError(f"格式校验未通过：{exc}") from None
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_dir = APP_DIR / "backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        backup_path = backup_dir / f"库存表替换前-{stamp}.xlsx"
        try:
            shutil.copy2(self.workbook_path, backup_path)
        except OSError:
            backup_path = None
        with self.lock:
            shutil.copy2(temp, self.workbook_path)
        try:
            temp.unlink(missing_ok=True)
        except Exception:
            pass
        uploaded = False
        upload_error = None
        if self.online and self.sync_token:
            try:
                SyncClient(self.server_url, self.sync_token).upload(
                    self.workbook_path
                )
                uploaded = True
            except SyncError as exc:
                upload_error = str(exc)
        message = "库存表已替换"
        if uploaded:
            message += "，并已同步到云端。"
        elif upload_error:
            message += f"，但同步云端失败：{upload_error}"
        else:
            message += "（未连接云端，仅本机生效）。"
        if backup_path:
            message += f" 替换前已备份：{backup_path}"
        return {
            "ok": True,
            "message": message,
            "uploaded": uploaded,
            "projects": self.list_projects(),
        }

    def view_excel(self) -> str:
        if not self.logged_in:
            raise ValueError("请先登录。")
        if self.role == "admin":
            try:
                os.startfile(self.workbook_path)
                return "已打开正式库存文件（可编辑）。"
            except OSError:
                return f"无法自动打开 Excel，请手动打开：{self.workbook_path}"
        copy_dir = APP_DIR / "view-copies"
        copy_dir.mkdir(parents=True, exist_ok=True)
        copy_path = copy_dir / f"库存查看副本-{uuid.uuid4().hex[:8]}.xlsx"
        copied = False
        if self.online and self.sync_token:
            try:
                SyncClient(self.server_url, self.sync_token).download_latest(copy_path)
                copied = True
            except SyncError:
                copied = False
        if not copied:
            shutil.copy2(self.workbook_path, copy_path)
        workbook = load_workbook(copy_path, data_only=False)
        for sheet in workbook.worksheets:
            sheet.protection.sheet = True
            sheet.protection.password = "admin123"
        workbook.save(copy_path)
        workbook.close()
        try:
            os.startfile(copy_path)
        except OSError:
            return (
                f"已生成云端只读副本（无法自动打开，请手动打开：{copy_path}）；"
                "若要修改，需在 Excel 中取消工作表保护并输入管理员密码。"
            )
        return "已打开云端只读副本；若要修改，需在 Excel 中取消工作表保护并输入管理员密码。"

    def state_payload(
        self,
        material: str = "",
        quantity_text: str = "",
        material_code: str = "",
        specification: str = "",
    ) -> dict[str, Any]:
        before = (
            current_balance(
                self.workbook_path,
                material=material,
                material_code=material_code,
                specification=specification,
            )
            if material or material_code
            else 0.0
        )
        try:
            quantity = float(quantity_text.replace(",", "").strip()) if quantity_text.strip() else 0.0
        except ValueError:
            quantity = 0.0
        workbook_values = workbook_suggestions(self.workbook_path)
        return {
            "workbookPath": str(self.workbook_path),
            "requesters": list(REQUESTERS),
            "projects": workbook_values["projects"] or list(PROJECTS),
            "locations": merge_suggestions(
                LOCATIONS,
                workbook_values["locations"],
            ),
            "changeTypes": CHANGE_TYPES,
            "material": material,
            "materialCode": material_code,
            "specification": specification,
            "balanceBefore": before,
            "balanceAfter": before + quantity,
        }

    def choose_workbook(self) -> Path | None:
        selected = windows_file_picker(self.workbook_path.parent)
        if not selected:
            return None
        candidate = Path(selected).resolve()
        validate_workbook(candidate)
        upgrade_workbook_schema(candidate)
        self.workbook_path = candidate
        save_settings(candidate)
        return candidate


def merge_suggestions(defaults: list[str], workbook_values: list[str]) -> list[str]:
    merged = []
    seen = set()
    for raw_value in [*defaults, *workbook_values]:
        value = str(raw_value).strip()
        identity = " ".join(unicodedata.normalize("NFKC", value).split()).casefold()
        if not value or identity in seen:
            continue
        seen.add(identity)
        merged.append(value)
    return merged


class InventoryServer(ThreadingHTTPServer):
    daemon_threads = True

    def __init__(self, address, handler, state: AppState):
        super().__init__(address, handler)
        self.app_state = state


class RequestHandler(BaseHTTPRequestHandler):
    server: InventoryServer

    def log_message(self, _format: str, *_args) -> None:
        return

    def send_json(
        self,
        payload: dict[str, Any] | list[Any],
        status: HTTPStatus = HTTPStatus.OK,
    ) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(body)

    def send_asset(self, filename: str, content_type: str) -> None:
        path = resource_dir() / filename
        if not path.is_file():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        body = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-cache")
        self.send_header(
            "Content-Security-Policy",
            "default-src 'self'; script-src 'self'; style-src 'self'; "
            "img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'",
        )
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(body)

    def read_json(self) -> dict[str, Any]:
        length = int(self.headers.get("Content-Length", "0"))
        if length < 0 or length > 1_000_000:
            raise ValueError("提交内容过大。")
        raw = self.rfile.read(length)
        try:
            data = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ValueError("提交内容格式错误。") from exc
        if not isinstance(data, dict):
            raise ValueError("提交内容格式错误。")
        return data

    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path in ("/", "/index.html"):
            self.send_asset("index.html", "text/html; charset=utf-8")
            return
        if parsed.path == "/styles.css":
            self.send_asset("styles.css", "text/css; charset=utf-8")
            return
        if parsed.path == "/app.js":
            self.send_asset("app.js", "text/javascript; charset=utf-8")
            return
        if parsed.path == "/api/health":
            self.send_json({"ok": True, "pid": os.getpid()})
            return
        if parsed.path == "/api/state":
            query = urllib.parse.parse_qs(parsed.query)
            material = query.get("material", [""])[0].strip()
            material_code = query.get("materialCode", [""])[0].strip()
            specification = query.get("specification", [""])[0].strip()
            quantity = query.get("quantity", [""])[0].strip()
            try:
                self.send_json(
                    self.server.app_state.state_payload(
                        material,
                        quantity,
                        material_code,
                        specification,
                    )
                )
            except Exception as exc:
                self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_REQUEST)
            return
        if parsed.path == "/api/materials":
            query = urllib.parse.parse_qs(parsed.query)
            search_text = query.get("q", [""])[0].strip()
            try:
                limit = int(query.get("limit", ["20"])[0])
                records = search_materials(
                    self.server.app_state.workbook_path,
                    search_text,
                    limit,
                )
                self.send_json({"ok": True, "materials": records})
            except (TypeError, ValueError) as exc:
                self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_REQUEST)
            return
        if parsed.path == "/api/recent":
            try:
                records = recent_records(self.server.app_state.workbook_path, limit=10)
                self.send_json({"ok": True, "records": records})
            except Exception as exc:
                self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_REQUEST)
            return
        if parsed.path == "/api/sync-status":
            self.send_json(self.server.app_state.sync_status_payload())
            return
        if parsed.path == "/api/versions":
            try:
                versions = SyncClient(
                    self.server.app_state.server_url,
                    self.server.app_state.sync_token,
                ).list_versions()
                self.send_json({"ok": True, "versions": versions})
            except SyncError as exc:
                self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_REQUEST)
            return
        if parsed.path == "/api/users":
            try:
                users = self.server.app_state.list_users()
                self.send_json({"ok": True, "users": users})
            except (ValueError, SyncError) as exc:
                self.send_json(
                    {"ok": False, "error": str(exc)},
                    HTTPStatus.FORBIDDEN,
                )
            return
        if parsed.path == "/api/projects":
            try:
                projects = self.server.app_state.list_projects()
                self.send_json({"ok": True, "projects": projects})
            except Exception as exc:
                self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_REQUEST)
            return
        if parsed.path == "/api/admin/email":
            try:
                email = self.server.app_state.get_security_email()
                self.send_json({"ok": True, "email": email})
            except (ValueError, SyncError) as exc:
                self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.FORBIDDEN)
            return
        if parsed.path == "/api/discover-servers":
            try:
                servers = discover_servers()
                self.send_json({"ok": True, "servers": servers})
            except Exception as exc:
                self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_REQUEST)
            return
        if parsed.path == "/api/desktop-path":
            self.send_json({"ok": True, "path": str(desktop_folder())})
            return
        if parsed.path == "/api/download-excel":
            try:
                data = self.read_json() if self.headers.get("Content-Length") else {}
                saved_path = self.server.app_state.desktop_download(
                    str(data.get("saveDir", "") or "")
                )
                self.send_json(
                    {
                        "ok": True,
                        "path": saved_path,
                        "message": f"文件已保存：{saved_path}",
                    }
                )
            except (ValueError, OSError) as exc:
                self.send_json(
                    {"ok": False, "error": str(exc)},
                    HTTPStatus.FORBIDDEN,
                )
            return
        if parsed.path == "/api/server-history":
            self.send_json(
                {
                    "ok": True,
                    "history": load_settings().get("server_history") or [],
                }
            )
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        try:
            if parsed.path == "/api/download-excel":
                data = self.read_json()
                try:
                    saved_path = self.server.app_state.desktop_download(
                        str(data.get("saveDir", "") or "")
                    )
                    self.send_json(
                        {
                            "ok": True,
                            "path": saved_path,
                            "message": f"文件已保存：{saved_path}",
                        }
                    )
                except (ValueError, OSError) as exc:
                    self.send_json(
                        {"ok": False, "error": str(exc)},
                        HTTPStatus.FORBIDDEN,
                    )
                return
            if parsed.path == "/api/download-version":
                if self.server.app_state.role != "admin":
                    self.send_json(
                        {"ok": False, "error": "历史版本下载仅限管理员。"},
                        HTTPStatus.FORBIDDEN,
                    )
                    return
                data = self.read_json()
                version_id = str(data.get("versionId", "")).strip()
                if not version_id:
                    self.send_json(
                        {"ok": False, "error": "缺少版本号。"},
                        HTTPStatus.BAD_REQUEST,
                    )
                    return
                try:
                    saved_path = self.server.app_state.download_version_to_desktop(
                        version_id,
                        str(data.get("saveDir", "") or ""),
                    )
                    self.send_json(
                        {
                            "ok": True,
                            "path": saved_path,
                            "message": f"文件已保存：{saved_path}",
                        }
                    )
                except (ValueError, OSError) as exc:
                    self.send_json(
                        {"ok": False, "error": str(exc)},
                        HTTPStatus.FORBIDDEN,
                    )
                return
            if parsed.path == "/api/upload-replace":
                length = int(self.headers.get("Content-Length") or 0)
                if length <= 0 or length > 50_000_000:
                    self.send_json(
                        {"ok": False, "error": "文件大小异常或未收到数据。"},
                        HTTPStatus.BAD_REQUEST,
                    )
                    return
                content_type = self.headers.get("Content-Type") or ""
                raw = self.rfile.read(length)
                file_bytes = b""
                if "application/json" in content_type:
                    try:
                        body = json.loads(raw.decode("utf-8"))
                        import base64
                        file_bytes = base64.b64decode(str(body.get("data") or ""))
                    except (ValueError, TypeError, UnicodeDecodeError) as exc:
                        self.send_json(
                            {"ok": False, "error": f"文件数据格式错误：{exc}"},
                            HTTPStatus.BAD_REQUEST,
                        )
                        return
                else:
                    fields = _parse_multipart(raw, content_type)
                    file_bytes = fields.get("file") or fields.get("excel")
                if not file_bytes:
                    self.send_json(
                        {"ok": False, "error": "未收到文件。"},
                        HTTPStatus.BAD_REQUEST,
                    )
                    return
                try:
                    result = self.server.app_state.replace_workbook(file_bytes)
                    self.send_json(result)
                except ValueError as exc:
                    self.send_json(
                        {"ok": False, "error": str(exc)},
                        HTTPStatus.BAD_REQUEST,
                    )
                except Exception as exc:
                    self.send_json(
                        {"ok": False, "error": str(exc)},
                        HTTPStatus.INTERNAL_SERVER_ERROR,
                    )
                return
            if parsed.path == "/api/submit":
                if self.server.app_state.role != "admin":
                    self.send_json(
                        {"ok": False, "error": "录入操作仅限管理员。"},
                        HTTPStatus.FORBIDDEN,
                    )
                    return
                conflict = self.server.app_state.check_conflict()
                if conflict:
                    self.send_json(
                        {"ok": False, "error": conflict, "conflict": True},
                        HTTPStatus.CONFLICT,
                    )
                    return
                data = self.read_json()
                project = str(data.get("project", "")).strip()
                if project not in self.server.app_state.list_projects():
                    raise ValueError("所选项目不存在，请从项目列表中选择。")
                material = str(data.get("material", "")).strip()
                requester = str(data.get("requester", "")).strip()
                project = str(data.get("project", "")).strip()
                specification = str(data.get("specification", "")).strip()
                material_code = str(data.get("materialCode", "")).strip()
                location = str(data.get("location", "")).strip()
                change_type = str(data.get("changeType", "")).strip()
                quantity_text = str(data.get("quantity", "")).strip()
                quantity = validate_record(
                    material,
                    requester,
                    project,
                    change_type,
                    quantity_text,
                    specification,
                    material_code,
                    location,
                )
                record = InventoryRecord(
                    material=material,
                    requester=requester,
                    po=str(data.get("po", "")).strip(),
                    changed_at=parse_datetime(str(data.get("changedAt", "")).strip()),
                    quantity=quantity,
                    change_type=change_type,
                    project=project,
                    note=str(data.get("note", "")).strip(),
                    specification=specification,
                    material_code=material_code,
                    location=location,
                )
                before, after, row = write_record(
                    self.server.app_state.workbook_path,
                    record,
                )
                uploaded = False
                upload_error: str | None = None
                try:
                    self.server.app_state.push_current()
                    uploaded = True
                except (SyncError, ValueError, OSError) as exc:
                    upload_error = str(exc)
                    with self.server.app_state.sync_lock:
                        self.server.app_state.pending_push = True
                        self.server.app_state._persist_sync_state()
                self.send_json(
                    {
                        "ok": True,
                        "row": row,
                        "materialCode": record.material_code,
                        "balanceBefore": before,
                        "balanceAfter": after,
                        "location": record.location,
                        "uploaded": uploaded,
                        "uploadError": upload_error,
                        "message": (
                            f"{material}：库存 {format_number(before)} → "
                            f"{format_number(after)}"
                        ),
                    }
                )
                return

            if parsed.path == "/api/login":
                data = self.read_json()
                payload = self.server.app_state.login(
                    str(data.get("password", ""))
                )
                self.send_json({"ok": True, **payload})
                return

            if parsed.path == "/api/sync-pull":
                payload = self.server.app_state.pull_latest()
                self.send_json({"ok": True, **payload})
                return

            if parsed.path == "/api/sync-push":
                if self.server.app_state.role != "admin":
                    self.send_json(
                        {"ok": False, "error": "上传当前版本仅限管理员。"},
                        HTTPStatus.FORBIDDEN,
                    )
                    return
                payload = self.server.app_state.push_current()
                self.send_json({"ok": True, **payload})
                return

            if parsed.path == "/api/rollback":
                data = self.read_json()
                version_id = str(data.get("versionId", "")).strip()
                if not version_id:
                    raise ValueError("缺少版本标识。")
                payload = self.server.app_state.rollback_to(version_id)
                self.send_json({"ok": True, **payload})
                return

            if parsed.path == "/api/users":
                data = self.read_json()
                action = str(data.get("action", "")).strip()
                if action == "add":
                    self.server.app_state.add_user(
                        str(data.get("password", "")).strip(),
                        str(data.get("note", "")).strip(),
                    )
                    self.send_json({"ok": True})
                elif action == "remove":
                    self.server.app_state.remove_user(
                        str(data.get("password", "")).strip()
                    )
                    self.send_json({"ok": True})
                else:
                    raise ValueError("未知的用户管理操作。")
                return

            if parsed.path == "/api/projects":
                data = self.read_json()
                action = str(data.get("action", "")).strip()
                project = str(data.get("project", "")).strip()
                if action == "add":
                    self.server.app_state.add_project(project)
                elif action == "remove":
                    self.server.app_state.remove_project(project)
                else:
                    raise ValueError("未知的项目管理操作。")
                self.send_json(
                    {
                        "ok": True,
                        "projects": self.server.app_state.list_projects(),
                    }
                )
                return

            if parsed.path == "/api/admin/password":
                data = self.read_json()
                self.server.app_state.change_admin_password(
                    str(data.get("oldPassword", "")).strip(),
                    str(data.get("newPassword", "")).strip(),
                )
                self.send_json({"ok": True, "relogin": True})
                return

            if parsed.path == "/api/admin/email":
                data = self.read_json()
                self.server.app_state.update_security_email(
                    str(data.get("email", "")).strip()
                )
                self.send_json({"ok": True})
                return

            if parsed.path == "/api/forgot-password":
                data = self.read_json()
                message = self.server.app_state.forgot_password(
                    str(data.get("type", "")).strip()
                )
                self.send_json({"ok": True, "message": message})
                return

            if parsed.path == "/api/view-excel":
                message = self.server.app_state.view_excel()
                self.send_json({"ok": True, "message": message})
                return

            if parsed.path == "/api/set-server-url":
                data = self.read_json()
                self.server.app_state.set_server_url(
                    str(data.get("url", "")).strip()
                )
                self.send_json({"ok": True})

            if parsed.path == "/api/open-excel":
                os.startfile(self.server.app_state.workbook_path)
                self.send_json({"ok": True})
                return

            if parsed.path == "/api/choose-excel":
                selected = self.server.app_state.choose_workbook()
                self.send_json(
                    {
                        "ok": True,
                        "cancelled": selected is None,
                        "workbookPath": (
                            str(self.server.app_state.workbook_path)
                        ),
                    }
                )
                return

            if parsed.path == "/api/shutdown":
                self.send_json({"ok": True})
                threading.Thread(
                    target=self.server.shutdown,
                    daemon=True,
                ).start()
                return

            self.send_error(HTTPStatus.NOT_FOUND)
        except (ValueError, PermissionError, FileNotFoundError) as exc:
            self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_REQUEST)
        except Exception as exc:
            self.send_json(
                {"ok": False, "error": f"操作失败：{exc}"},
                HTTPStatus.INTERNAL_SERVER_ERROR,
            )


def windows_file_picker(initial_dir: Path) -> str | None:
    result_file = Path(tempfile.gettempdir()) / f"inventory-picker-{os.getpid()}.txt"
    if result_file.exists():
        try:
            result_file.unlink()
        except OSError:
            pass
    script = r"""
Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Filter = 'Excel 工作簿 (*.xlsx)|*.xlsx'
$dialog.Title = '选择库存 Excel'
$dialog.InitialDirectory = $env:LAB_PICKER_INITIAL
if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    [System.IO.File]::WriteAllText(
        $env:LAB_PICKER_RESULT,
        $dialog.FileName,
        [System.Text.Encoding]::UTF8
    )
}
"""
    environment = os.environ.copy()
    environment["LAB_PICKER_INITIAL"] = str(initial_dir)
    environment["LAB_PICKER_RESULT"] = str(result_file)
    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    subprocess.run(
        [
            "powershell.exe",
            "-NoProfile",
            "-STA",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            script,
        ],
        env=environment,
        check=False,
        timeout=300,
        creationflags=creation_flags,
    )
    if not result_file.exists():
        return None
    try:
        return result_file.read_text(encoding="utf-8").strip() or None
    finally:
        try:
            result_file.unlink()
        except OSError:
            pass


def existing_instance_port() -> int | None:
    if not RUNTIME_PATH.is_file():
        return None
    try:
        runtime = json.loads(RUNTIME_PATH.read_text(encoding="utf-8"))
        port = int(runtime["port"])
        with urllib.request.urlopen(
            f"http://127.0.0.1:{port}/api/health",
            timeout=1.0,
        ) as response:
            if response.status == 200:
                return port
    except (OSError, ValueError, KeyError, json.JSONDecodeError, urllib.error.URLError):
        return None
    return None


def write_runtime_file(port: int) -> None:
    APP_DIR.mkdir(parents=True, exist_ok=True)
    RUNTIME_PATH.write_text(
        json.dumps({"pid": os.getpid(), "port": port}, ensure_ascii=False),
        encoding="utf-8",
    )


def remove_runtime_file() -> None:
    if not RUNTIME_PATH.exists():
        return
    try:
        runtime = json.loads(RUNTIME_PATH.read_text(encoding="utf-8"))
        if int(runtime.get("pid", -1)) == os.getpid():
            RUNTIME_PATH.unlink()
    except (OSError, ValueError, json.JSONDecodeError):
        pass


def launch_server(
    *,
    open_browser: bool = True,
    requested_port: int = 0,
    gui: bool = True,
) -> None:
    existing_port = existing_instance_port()
    if existing_port:
        if open_browser:
            native_message(
                "ERL库存管理已在运行。\n"
                "请查看任务栏或右下角托盘图标，不要重复打开程序。",
                error=False,
            )
        return

    state = AppState()
    server = InventoryServer(("127.0.0.1", requested_port), RequestHandler, state)
    port = int(server.server_address[1])
    write_runtime_file(port)
    atexit.register(remove_runtime_file)
    url = f"http://127.0.0.1:{port}/"

    if gui:
        server_thread = threading.Thread(
            target=server.serve_forever,
            kwargs={"poll_interval": 0.25},
            daemon=True,
        )
        server_thread.start()
        icon_candidates = [
            Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
            / "assets"
            / "erl_client.ico",
            Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
            / "assets"
            / "erl_client.png",
        ]
        icon = next(
            (str(candidate) for candidate in icon_candidates if candidate.exists()),
            None,
        )
        used_gui = run_gui(
            url,
            "ERL库存管理",
            icon=icon,
            monitor_url=f"{url}api/health",
            on_quit=lambda: (_remove_runtime_safe(), server.shutdown()),
        )
        if not used_gui:
            # WebView2 缺失/失败已由 gui_window 提示并自动安装，这里不再重复弹窗
            server.shutdown()
            server_thread.join()
        server.server_close()
        remove_runtime_file()
        return

    if open_browser:
        threading.Timer(0.35, lambda: webbrowser.open(url, new=1)).start()
    try:
        server.serve_forever(poll_interval=0.25)
    finally:
        server.server_close()
        remove_runtime_file()


def _remove_runtime_safe() -> None:
    try:
        remove_runtime_file()
    except Exception:
        pass
    # 退出程序时清除登录状态，下次打开需重新登录
    try:
        if SYNC_STATE_PATH.exists():
            SYNC_STATE_PATH.unlink()
    except Exception:
        pass


def parse_port(arguments: list[str]) -> int:
    if "--port" not in arguments:
        return 0
    index = arguments.index("--port")
    try:
        return int(arguments[index + 1])
    except (IndexError, ValueError) as exc:
        raise ValueError("--port 后需要填写端口号。") from exc


def _migrate_legacy(legacy: Path, target: Path) -> None:
    """把旧目录里的数据复制到目标目录，成功后清理旧目录中的程序数据。"""
    has_legacy = (legacy / "settings.json").exists() or (legacy / "data").exists()
    if not has_legacy:
        return
    try:
        target.mkdir(parents=True, exist_ok=True)
        for name in ("settings.json", "sync.json"):
            src = legacy / name
            if src.exists() and not (target / name).exists():
                shutil.copy2(src, target / name)
        for sub in ("data", "backups", "view-copies"):
            src = legacy / sub
            if src.exists():
                dst = target / sub
                if not dst.exists():
                    shutil.copytree(src, dst)
        # 迁移成功后再清理旧目录中的程序数据（避免明文 Excel 被他人直接打开）
        for name in (
            "settings.json",
            "sync.json",
            "data",
            "backups",
            "view-copies",
            ".inventory-runtime.json",
        ):
            item = legacy / name
            if not item.exists():
                continue
            try:
                if item.is_dir():
                    shutil.rmtree(item)
                else:
                    item.unlink()
            except Exception:
                pass
    except Exception:
        pass


def migrate_legacy_data() -> None:
    """把旧版存在 exe 文件夹里的数据迁移到用户专属目录，避免明文 Excel 留在程序文件夹。"""
    if not getattr(sys, "frozen", False):
        return  # 源码模式不迁移
    legacy = Path(sys.executable).resolve().parent
    if APP_DIR.resolve() == legacy.resolve():
        return
    _migrate_legacy(legacy, APP_DIR)


def main() -> None:
    arguments = sys.argv[1:]
    if len(arguments) == 2 and arguments[0] == "--self-test":
        result = run_self_test(Path(arguments[1]).resolve())
        print(json.dumps(result, ensure_ascii=False))
        return
    migrate_legacy_data()
    try:
        launch_server(
            open_browser="--no-browser" not in arguments,
            requested_port=parse_port(arguments),
            gui="--browser" not in arguments,
        )
    except Exception as exc:
        native_message(str(exc), error=True)


if __name__ == "__main__":
    main()
