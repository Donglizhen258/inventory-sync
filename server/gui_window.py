# -*- coding: utf-8 -*-
"""ERL库存管理：pywebview 独立窗口 + 系统托盘封装（客户端与服务端控制台共用）

- 关闭窗口（点 X）不会退出程序，而是隐藏到 Windows 右下角系统托盘；
- 托盘菜单可「显示窗口」或「退出程序」；
- monitor_url 传入后，后台监控该地址，连续不可达会自动退出整个程序
  （用于客户端页面点「退出程序」后自动关闭窗口并结束进程）。
- 若本机不支持 pywebview（缺 WebView2 运行库等），返回 False，
  调用方应降级为浏览器模式。
"""
from __future__ import annotations

import os
import subprocess
import sys
import threading
import time
import traceback
import urllib.request
from pathlib import Path
from typing import Callable

from tray import TrayIcon

WEBVIEW2_CLIENT_GUID = "{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"


def _debug_log(message: str) -> None:
    """GUI 启动失败时写诊断日志到程序所在目录，便于排查。"""
    try:
        base = Path(sys.executable).resolve().parent
        with open(base / "gui_debug.log", "a", encoding="utf-8") as handle:
            handle.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")
    except Exception:
        pass


def _native_message(message: str, error: bool = False) -> None:
    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(
            None,
            message,
            "ERL库存管理" + (" - 提示" if not error else " - 错误"),
            0x10 if error else 0x40,
        )
    except Exception:
        pass


def webview2_installed() -> bool:
    """通过注册表检测 WebView2 运行库是否已安装。"""
    import winreg

    candidates = [
        r"SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients" + WEBVIEW2_CLIENT_GUID,
        r"SOFTWARE\Microsoft\EdgeUpdate\Clients" + WEBVIEW2_CLIENT_GUID,
    ]
    for path in candidates:
        try:
            winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path)
            return True
        except OSError:
            continue
    return False


def _bundled_webview2_installer() -> Path | None:
    """返回打包自带的 WebView2 安装器路径（PyInstaller 资源或 exe 同目录）。"""
    bundle = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    candidate = bundle / "assets" / "MicrosoftEdgeWebview2Setup.exe"
    if candidate.exists():
        return candidate
    candidate2 = Path(sys.executable).resolve().parent / "assets" / "MicrosoftEdgeWebview2Setup.exe"
    return candidate2 if candidate2.exists() else None


def _handle_webview2_missing() -> None:
    """未检测到 WebView2 时：提示并自动安装（软件自带安装器），不再降级浏览器。"""
    if webview2_installed():
        return
    installer = _bundled_webview2_installer()
    if installer is not None:
        _debug_log("未检测到 WebView2，启动内置安装器自动安装")
        _native_message(
            "未检测到 Microsoft Edge WebView2 运行库，正在自动安装（需要联网），"
            "安装完成后请重新打开本软件。"
        )
        try:
            subprocess.Popen(
                [str(installer), "/silent", "/install"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            return
        except Exception as exc:
            _debug_log(f"自动安装启动失败: {exc}")
    _native_message(
        "未检测到 Microsoft Edge WebView2 运行库，且自动安装失败。\n"
        "请手动下载安装后重试：\nhttps://developer.microsoft.com/microsoft-edge/webview2/",
        error=True,
    )


def _http_ok(url: str, timeout: float = 1.5) -> bool:
    try:
        with urllib.request.urlopen(url, timeout=timeout) as response:
            return response.status == 200
    except Exception:
        return False


def run_gui(
    url: str,
    title: str,
    icon: str | None = None,
    width: int = 1200,
    height: int = 860,
    min_width: int = 940,
    min_height: int = 620,
    monitor_url: str | None = None,
    on_quit: Callable[[], None] | None = None,
) -> bool:
    """以独立窗口方式运行。返回 True=GUI 启动成功；False=不支持，需降级浏览器。"""
    try:
        import webview
    except Exception as exc:
        _debug_log(f"import webview 失败: {exc}\n{traceback.format_exc()}")
        return False

    window = None
    destroyed = threading.Event()

    try:
        window = webview.create_window(
            title,
            url,
            width=width,
            height=height,
            min_size=(min_width, min_height),
            icon=icon,
        )
    except Exception:
        try:
            window = webview.create_window(
                title,
                url,
                width=width,
                height=height,
                min_size=(min_width, min_height),
            )
        except Exception as exc:
            _debug_log(
                f"create_window 失败（无法使用独立窗口）: {exc}\n{traceback.format_exc()}"
            )
            return False

    def on_closing() -> bool:
        # 点 X 时隐藏到托盘，不退出
        try:
            window.hide()
        except Exception:
            pass
        return False

    try:
        window.events.closing += on_closing
    except Exception:
        pass

    def tray_show() -> None:
        try:
            window.show()
            window.restore()
        except Exception:
            pass

    def tray_quit() -> None:
        destroyed.set()
        try:
            window.destroy()
        except Exception:
            pass
        # 兜底：若 destroy 未让事件循环退出，2 秒后强制结束进程并执行清理
        def _force_exit() -> None:
            time.sleep(2)
            if on_quit is not None:
                try:
                    on_quit()
                except Exception:
                    pass
            try:
                import os

                os._exit(0)
            except Exception:
                pass

        threading.Thread(target=_force_exit, daemon=True).start()

    tray = None
    try:
        tray = TrayIcon(title, icon, tray_show, tray_quit)
        tray.start()
    except Exception as exc:
        _debug_log(f"系统托盘创建失败: {exc}\n{traceback.format_exc()}")
        tray = None

    def monitor() -> None:
        if not monitor_url:
            return
        misses = 0
        while not destroyed.is_set():
            time.sleep(2)
            if _http_ok(monitor_url):
                misses = 0
            else:
                misses += 1
                if misses >= 3:
                    tray_quit()
                    break

    if monitor_url:
        threading.Thread(target=monitor, daemon=True).start()

    started = time.time()
    try:
        webview.start(debug=False)
    except Exception as exc:
        _debug_log(f"webview.start 异常: {exc}\n{traceback.format_exc()}")
        _handle_webview2_missing()
        return False
    finally:
        destroyed.set()
        if tray is not None:
            try:
                tray.stop()
            except Exception:
                pass
    # WebView2 初始化失败时 start() 会几乎立即返回（窗口从未显示），
    # 此时不再降级浏览器：提示并自动安装 WebView2。
    if time.time() - started < 3:
        _debug_log("webview.start 快速返回（<3秒），疑似 WebView2 初始化失败")
        try:
            window.destroy()
        except Exception:
            pass
        _handle_webview2_missing()
        return False
    if on_quit is not None:
        try:
            on_quit()
        except Exception:
            pass
    return True
