# -*- coding: utf-8 -*-
"""ERL库存管理：Windows 系统托盘（Shell_NotifyIcon 原生实现，不依赖 pywebview）

提供右下角托盘图标：左键单击/双击显示窗口；右键菜单「显示窗口 / 退出程序」。
在独立线程中运行消息循环，跨线程安全回调。
"""
from __future__ import annotations

import ctypes
import threading
from ctypes import wintypes

user32 = ctypes.windll.user32
shell32 = ctypes.windll.shell32
kernel32 = ctypes.windll.kernel32

WM_USER = 0x0400
WM_TRAYICON = WM_USER + 20
WM_RBUTTONUP = 0x0205
WM_LBUTTONUP = 0x0202
WM_LBUTTONDBLCLK = 0x0203
WM_COMMAND = 0x0111
WM_DESTROY = 0x0002

NIM_ADD = 0
NIM_MODIFY = 1
NIM_DELETE = 2
NIF_MESSAGE = 0x0001
NIF_ICON = 0x0002
NIF_TIP = 0x0004

MF_STRING = 0
TPM_RIGHTBUTTON = 0x0002
TPM_RETURNCMD = 0x0100

MENU_SHOW = 1
MENU_QUIT = 2

CS_HREDRAW = 0x0002
CS_VREDRAW = 0x0001


class NOTIFYICONDATAW(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("hWnd", wintypes.HWND),
        ("uID", wintypes.UINT),
        ("uFlags", wintypes.UINT),
        ("uCallbackMessage", wintypes.UINT),
        ("hIcon", wintypes.HANDLE),
        ("szTip", wintypes.WCHAR * 128),
        ("dwState", wintypes.DWORD),
        ("dwStateMask", wintypes.DWORD),
        ("szInfo", wintypes.WCHAR * 256),
        ("uTimeout", wintypes.UINT),
        ("szInfoTitle", wintypes.WCHAR * 64),
        ("dwInfoFlags", wintypes.DWORD),
        ("guidItem", ctypes.c_byte * 16),
        ("hBalloonIcon", wintypes.HANDLE),
    ]


class WNDCLASSEXW(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.UINT),
        ("style", wintypes.UINT),
        ("lpfnWndProc", ctypes.c_void_p),
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", wintypes.HINSTANCE),
        ("hIcon", wintypes.HANDLE),
        ("hCursor", wintypes.HANDLE),
        ("hbrBackground", wintypes.HANDLE),
        ("lpszMenuName", wintypes.LPCWSTR),
        ("lpszClassName", wintypes.LPCWSTR),
        ("hIconSm", wintypes.HANDLE),
    ]


LRESULT = ctypes.c_longlong  # Windows 64 位 LONG_PTR；ctypes.wintypes 未提供
WNDPROC = ctypes.WINFUNCTYPE(
    LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM
)

# 统一声明参数类型，避免 64 位指针被截断为 32 位
user32.DefWindowProcW.argtypes = [
    wintypes.HWND,
    wintypes.UINT,
    wintypes.WPARAM,
    wintypes.LPARAM,
]
user32.DefWindowProcW.restype = LRESULT
user32.PostMessageW.argtypes = [
    wintypes.HWND,
    wintypes.UINT,
    wintypes.WPARAM,
    wintypes.LPARAM,
]
user32.TrackPopupMenu.argtypes = [
    wintypes.HMENU,
    wintypes.UINT,
    ctypes.c_int,
    ctypes.c_int,
    ctypes.c_int,
    wintypes.HWND,
    wintypes.LPRECT,
]
user32.TrackPopupMenu.restype = wintypes.UINT
user32.LoadImageW.argtypes = [
    wintypes.HINSTANCE,
    wintypes.LPCWSTR,
    wintypes.UINT,
    ctypes.c_int,
    ctypes.c_int,
    wintypes.UINT,
]
user32.LoadImageW.restype = wintypes.HANDLE


class TrayIcon:
    """系统托盘图标。start() 后在其他线程调用 stop() 可移除图标并结束消息循环。"""

    def __init__(
        self,
        title: str,
        icon_path: str | None,
        on_show,
        on_quit,
    ) -> None:
        self.title = title
        self.icon_path = icon_path
        self.on_show = on_show
        self.on_quit = on_quit
        self.hwnd: int | None = None
        self.class_name = f"ERLTray_{id(self)}"
        self._wndproc_ref = None
        self._thread: threading.Thread | None = None
        self._running = threading.Event()

    def _wndproc(self, hwnd, msg, wparam, lparam) -> int:
        if msg == WM_TRAYICON:
            if lparam in (WM_LBUTTONUP, WM_LBUTTONDBLCLK):
                self._safe(self.on_show)
            elif lparam == WM_RBUTTONUP:
                self._show_menu()
        elif msg == WM_COMMAND:
            if wparam == MENU_SHOW:
                self._safe(self.on_show)
            elif wparam == MENU_QUIT:
                self._safe(self.on_quit)
        elif msg == WM_DESTROY:
            user32.PostQuitMessage(0)
        return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

    @staticmethod
    def _safe(fn) -> None:
        try:
            fn()
        except Exception:
            pass

    def _show_menu(self) -> None:
        menu = user32.CreatePopupMenu()
        user32.AppendMenuW(menu, MF_STRING, MENU_SHOW, "显示窗口")
        user32.AppendMenuW(menu, MF_STRING, MENU_QUIT, "退出程序")
        pos = wintypes.POINT()
        user32.GetCursorPos(ctypes.byref(pos))
        cmd = user32.TrackPopupMenu(
            menu,
            TPM_RIGHTBUTTON | TPM_RETURNCMD,
            pos.x,
            pos.y,
            0,
            self.hwnd,
            None,
        )
        user32.DestroyMenu(menu)
        if cmd == MENU_SHOW:
            self._safe(self.on_show)
        elif cmd == MENU_QUIT:
            self._safe(self.on_quit)

    def _load_icon(self):
        if self.icon_path:
            try:
                return user32.LoadImageW(
                    None,
                    self.icon_path,
                    1,  # IMAGE_ICON
                    32,
                    32,
                    0x00000010,  # LR_LOADFROMFILE
                )
            except Exception:
                pass
        return None

    def start(self) -> None:
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def _run(self) -> None:
        instance = kernel32.GetModuleHandleW(None)
        wndproc = WNDPROC(self._wndproc)
        self._wndproc_ref = wndproc  # 防止被 GC
        class_info = WNDCLASSEXW()
        class_info.cbSize = ctypes.sizeof(WNDCLASSEXW)
        class_info.style = CS_HREDRAW | CS_VREDRAW
        class_info.lpfnWndProc = ctypes.cast(wndproc, ctypes.c_void_p)
        class_info.hInstance = instance
        class_info.lpszClassName = self.class_name
        atom = user32.RegisterClassExW(ctypes.byref(class_info))
        if not atom:
            return
        self.hwnd = user32.CreateWindowExW(
            0,
            self.class_name,
            "",
            0,
            0,
            0,
            0,
            0,
            None,
            None,
            instance,
            None,
        )
        if not self.hwnd:
            return
        icon = self._load_icon()
        nid = NOTIFYICONDATAW()
        nid.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        nid.hWnd = self.hwnd
        nid.uID = 1
        nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP
        nid.uCallbackMessage = WM_TRAYICON
        nid.hIcon = icon
        nid.szTip = self.title[:127]
        self.nid = nid
        shell32.Shell_NotifyIconW(NIM_ADD, ctypes.byref(nid))
        self._running.set()
        msg = wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))
        # 退出清理
        if self.nid is not None:
            shell32.Shell_NotifyIconW(NIM_DELETE, ctypes.byref(self.nid))
        if self.hwnd:
            user32.DestroyWindow(self.hwnd)
        user32.UnregisterClassW(self.class_name, instance)

    def stop(self) -> None:
        if self.hwnd:
            user32.PostMessageW(self.hwnd, WM_DESTROY, 0, 0)
