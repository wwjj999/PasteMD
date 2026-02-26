"""Hotkey binding manager – native OS API implementation.

Windows: RegisterHotKey + WM_HOTKEY message loop
macOS:   CGEventTapCreate + CFRunLoop
"""

from typing import Optional, Callable
import threading

from ...utils.logging import log
from ...utils.system_detect import is_macos, is_windows


# ---------------------------------------------------------------------------
# Windows implementation
# ---------------------------------------------------------------------------

class _WinHotkeyThread:
    """Hidden-window message loop that listens for WM_HOTKEY."""

    WM_HOTKEY = 0x0312
    WM_USER = 0x0400
    # Custom messages dispatched via PostThreadMessage
    WM_REGISTER = WM_USER + 1
    WM_UNREGISTER = WM_USER + 2
    WM_QUIT_LOOP = WM_USER + 3

    HOTKEY_ID = 1  # single hotkey, single id

    def __init__(self):
        self._thread: Optional[threading.Thread] = None
        self._thread_id: Optional[int] = None
        self._ready = threading.Event()
        self._callback: Optional[Callable[[], None]] = None
        self._registered = False

    # -- public interface ---------------------------------------------------

    def start(self) -> None:
        if self._thread is not None:
            return
        self._ready.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        self._ready.wait(timeout=5.0)

    def stop(self) -> None:
        if self._thread is None:
            return
        self._post(self.WM_QUIT_LOOP)
        self._thread.join(timeout=2.0)
        self._thread = None
        self._thread_id = None
        self._registered = False

    def register(self, modifiers: int, vk: int, callback: Callable[[], None]) -> None:
        self._callback = callback
        # Pack mod+vk into wParam/lParam of the message
        self._post(self.WM_REGISTER, modifiers, vk)

    def unregister(self) -> None:
        self._callback = None
        self._post(self.WM_UNREGISTER)

    @property
    def is_registered(self) -> bool:
        return self._registered

    # -- internals ----------------------------------------------------------

    def _post(self, msg: int, wparam: int = 0, lparam: int = 0) -> None:
        if self._thread_id is None:
            return
        import ctypes
        ctypes.windll.user32.PostThreadMessageW(self._thread_id, msg, wparam, lparam)

    def _run(self) -> None:
        import ctypes
        import ctypes.wintypes as wt

        kernel32 = ctypes.windll.kernel32
        user32 = ctypes.windll.user32

        self._thread_id = kernel32.GetCurrentThreadId()

        # Force message queue creation
        msg = wt.MSG()
        user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 0)
        self._ready.set()

        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            if msg.message == self.WM_HOTKEY:
                cb = self._callback
                if cb:
                    try:
                        cb()
                    except Exception as e:
                        log(f"Hotkey callback error: {e}")

            elif msg.message == self.WM_REGISTER:
                if self._registered:
                    user32.UnregisterHotKey(None, self.HOTKEY_ID)
                    self._registered = False
                mod, vk = msg.wParam, msg.lParam
                ok = user32.RegisterHotKey(None, self.HOTKEY_ID, mod, vk)
                self._registered = bool(ok)
                if not ok:
                    log(f"RegisterHotKey failed: mod=0x{mod:X} vk=0x{vk:X}")

            elif msg.message == self.WM_UNREGISTER:
                if self._registered:
                    user32.UnregisterHotKey(None, self.HOTKEY_ID)
                    self._registered = False

            elif msg.message == self.WM_QUIT_LOOP:
                if self._registered:
                    user32.UnregisterHotKey(None, self.HOTKEY_ID)
                    self._registered = False
                break

            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))


# ---------------------------------------------------------------------------
# macOS implementation
# ---------------------------------------------------------------------------

class _MacHotkeyTap:
    """CGEventTap-based global hotkey listener for macOS."""

    def __init__(self):
        self._thread: Optional[threading.Thread] = None
        self._tap = None
        self._loop_ref = None
        self._target_modifiers: int = 0
        self._target_keycode: int = -1
        self._callback: Optional[Callable[[], None]] = None
        self._lock = threading.Lock()

    # -- public interface ---------------------------------------------------

    def start(self) -> None:
        if self._thread is not None:
            return
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        if self._thread is None:
            return
        if self._loop_ref is not None:
            import Quartz
            Quartz.CFRunLoopStop(self._loop_ref)
        self._thread.join(timeout=2.0)
        self._thread = None
        self._loop_ref = None

    def set_hotkey(self, modifiers: int, keycode: int, callback: Callable[[], None]) -> None:
        with self._lock:
            self._target_modifiers = modifiers
            self._target_keycode = keycode
            self._callback = callback

    def clear_hotkey(self) -> None:
        with self._lock:
            self._target_keycode = -1
            self._callback = None

    @property
    def has_hotkey(self) -> bool:
        with self._lock:
            return self._callback is not None and self._target_keycode >= 0

    # -- internals ----------------------------------------------------------

    def _run(self) -> None:
        import Quartz
        from Quartz import (
            CGEventTapCreate,
            CGEventGetIntegerValueField,
            CGEventGetFlags,
            CGEventTapEnable,
            kCGSessionEventTap,
            kCGHeadInsertEventTap,
            kCGEventKeyDown,
        )

        # Mask for modifier flags we care about (strip device-dependent bits)
        MOD_MASK = (
            Quartz.kCGEventFlagMaskShift
            | Quartz.kCGEventFlagMaskControl
            | Quartz.kCGEventFlagMaskAlternate
            | Quartz.kCGEventFlagMaskCommand
        )

        try:
            NSSystemDefined = Quartz.NSSystemDefined
        except AttributeError:
            NSSystemDefined = 14

        def _callback(proxy, event_type, event, refcon):
            # Skip NSSystemDefined events (input method, media keys, etc.)
            if event_type == NSSystemDefined:
                return event

            # We only match on keyDown
            if event_type != kCGEventKeyDown:
                return event

            keycode = CGEventGetIntegerValueField(event, Quartz.kCGKeyboardEventKeycode)
            flags = CGEventGetFlags(event) & MOD_MASK

            with self._lock:
                target_kc = self._target_keycode
                target_mod = self._target_modifiers
                cb = self._callback

            if target_kc < 0 or cb is None:
                return event

            if keycode == target_kc and flags == target_mod:
                try:
                    from ...core.state import app_state
                    ui_queue = getattr(app_state, "ui_queue", None)
                    if ui_queue is not None:
                        ui_queue.put(cb)
                    else:
                        cb()
                except Exception as e:
                    log(f"Hotkey callback error: {e}")

            return event

        event_mask = (1 << kCGEventKeyDown)

        tap = CGEventTapCreate(
            kCGSessionEventTap,
            kCGHeadInsertEventTap,
            0,  # listenOnly=0 but we return the event unchanged
            event_mask,
            _callback,
            None,
        )

        if tap is None:
            log("Failed to create CGEventTap – accessibility permission required")
            return

        self._tap = tap
        source = Quartz.CFMachPortCreateRunLoopSource(None, tap, 0)
        self._loop_ref = Quartz.CFRunLoopGetCurrent()
        Quartz.CFRunLoopAddSource(self._loop_ref, source, Quartz.kCFRunLoopCommonModes)
        CGEventTapEnable(tap, True)

        log("macOS CGEventTap hotkey listener started")
        Quartz.CFRunLoopRun()
        log("macOS CGEventTap hotkey listener stopped")


# ---------------------------------------------------------------------------
# HotkeyManager – unified public API
# ---------------------------------------------------------------------------

class HotkeyManager:
    """热键管理器 – 使用原生 OS API 实现全局热键监听"""

    def __init__(self):
        self.current_hotkey: Optional[str] = None
        self._backend: object = None  # _WinHotkeyThread | _MacHotkeyTap

    # -- helpers ------------------------------------------------------------

    @staticmethod
    def _parse(hotkey: str):
        """Parse hotkey string using the platform-specific HotkeyChecker."""
        from ...utils.hotkey_checker import HotkeyChecker
        return HotkeyChecker.parse_hotkey(hotkey)

    def _ensure_backend(self):
        if self._backend is not None:
            return
        if is_windows():
            self._backend = _WinHotkeyThread()
        elif is_macos():
            self._backend = _MacHotkeyTap()
        else:
            raise RuntimeError("Unsupported platform for hotkey listening")
        self._backend.start()

    # -- public API ---------------------------------------------------------

    def bind(self, hotkey: str, callback: Callable[[], None]) -> None:
        """
        绑定全局热键

        Args:
            hotkey: 热键字符串 (例如: "<ctrl>+<shift>+b")
            callback: 热键触发时的回调函数
        """
        parsed = self._parse(hotkey)
        if parsed is None:
            raise ValueError(f"Cannot parse hotkey: {hotkey}")

        modifiers, key = parsed
        self._ensure_backend()

        if is_windows():
            self._backend.register(modifiers, key, callback)
        elif is_macos():
            self._backend.set_hotkey(modifiers, key, callback)

        self.current_hotkey = hotkey
        log(f"Hotkey bound: {hotkey}")

    def unbind(self) -> None:
        """解绑当前热键"""
        if self._backend is None:
            return
        previous = self.current_hotkey

        if is_windows():
            self._backend.unregister()
        elif is_macos():
            self._backend.clear_hotkey()

        self.current_hotkey = None
        log(f"Hotkey unbound: {previous}")

    def restart(self, hotkey: str, callback: Callable[[], None]) -> None:
        """重启热键绑定"""
        self.unbind()
        self.bind(hotkey, callback)

    def is_bound(self) -> bool:
        """检查是否有热键绑定"""
        if self._backend is None or self.current_hotkey is None:
            return False
        if is_windows():
            return self._backend.is_registered
        elif is_macos():
            return self._backend.has_hotkey
        return False

    def pause(self) -> None:
        """暂停热键监听（用于录制时避免触发）"""
        if self._backend is None or not self.current_hotkey:
            return

        if is_windows():
            self._backend.unregister()
        elif is_macos():
            self._backend.clear_hotkey()

        log(f"Hotkey paused: {self.current_hotkey}")

    def resume(self, callback: Callable[[], None]) -> None:
        """恢复热键监听"""
        if not self.current_hotkey:
            return

        # Already active
        if self.is_bound():
            return

        parsed = self._parse(self.current_hotkey)
        if parsed is None:
            return

        modifiers, key = parsed
        self._ensure_backend()

        if is_windows():
            self._backend.register(modifiers, key, callback)
        elif is_macos():
            self._backend.set_hotkey(modifiers, key, callback)

        log(f"Hotkey resumed: {self.current_hotkey}")
