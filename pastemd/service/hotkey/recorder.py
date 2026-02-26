"""Hotkey recording functionality – native implementation.

Windows: SetWindowsHookEx(WH_KEYBOARD_LL) + message loop
macOS:   Recording uses Tk events in dialog.py; this module is not used on macOS.
"""

from typing import Optional, Callable, Set

from ...utils.logging import log
from ...utils.hotkey_checker import HotkeyChecker
from ...i18n import t


# Reverse VK map: vk_code → key_name (built from win32 hotkey_checker's VK_MAP)
def _build_reverse_vk_map():
    """Build reverse VK code → key name mapping for the recorder."""
    # Modifier VK codes
    _MODIFIER_VKS = {
        0xA0: "shift", 0xA1: "shift",  # VK_LSHIFT, VK_RSHIFT
        0x10: "shift",                  # VK_SHIFT
        0xA2: "ctrl", 0xA3: "ctrl",    # VK_LCONTROL, VK_RCONTROL
        0x11: "ctrl",                   # VK_CONTROL
        0xA4: "alt", 0xA5: "alt",      # VK_LMENU, VK_RMENU
        0x12: "alt",                    # VK_MENU
        0x5B: "cmd", 0x5C: "cmd",      # VK_LWIN, VK_RWIN
    }

    # Normal keys
    _NORMAL_VKS = {
        0x08: "backspace", 0x09: "tab", 0x0D: "enter", 0x13: "pause",
        0x14: "caps_lock", 0x1B: "esc",
        0x20: "space", 0x21: "page_up", 0x22: "page_down",
        0x23: "end", 0x24: "home",
        0x25: "left", 0x26: "up", 0x27: "right", 0x28: "down",
        0x2C: "print_screen", 0x2D: "insert", 0x2E: "delete",
        0x70: "f1", 0x71: "f2", 0x72: "f3", 0x73: "f4",
        0x74: "f5", 0x75: "f6", 0x76: "f7", 0x77: "f8",
        0x78: "f9", 0x79: "f10", 0x7A: "f11", 0x7B: "f12",
        0x60: "num0", 0x61: "num1", 0x62: "num2", 0x63: "num3",
        0x64: "num4", 0x65: "num5", 0x66: "num6", 0x67: "num7",
        0x68: "num8", 0x69: "num9",
        0x6A: "multiply", 0x6B: "add", 0x6C: "separator",
        0x6D: "subtract", 0x6E: "decimal", 0x6F: "divide",
        0xBA: ";", 0xBB: "=", 0xBC: ",", 0xBD: "-",
        0xBE: ".", 0xBF: "/", 0xC0: "`",
        0xDB: "[", 0xDC: "\\", 0xDD: "]", 0xDE: "'",
    }

    # A-Z  (0x41 – 0x5A)
    for vk in range(0x41, 0x5B):
        _NORMAL_VKS[vk] = chr(vk).lower()

    # 0-9  (0x30 – 0x39)
    for vk in range(0x30, 0x3A):
        _NORMAL_VKS[vk] = chr(vk)

    return _MODIFIER_VKS, _NORMAL_VKS


_MODIFIER_VKS, _NORMAL_VKS = _build_reverse_vk_map()


class HotkeyRecorder:
    """热键录制器 – Windows 使用低级键盘钩子"""

    def __init__(self):
        self.recording = False
        self.pressed_keys: Set[str] = set()
        self.released_keys: Set[str] = set()
        self.all_pressed_keys: Set[str] = set()
        self._hook_thread = None
        self._hook_handle = None
        self.on_update_callback: Optional[Callable[[str], None]] = None
        self.on_finish_callback: Optional[Callable[[Optional[str], Optional[str]], None]] = None

    def start_recording(
        self,
        on_update: Optional[Callable[[str], None]] = None,
        on_finish: Optional[Callable[[Optional[str], Optional[str]], None]] = None,
    ) -> None:
        """
        开始录制热键

        Args:
            on_update: 更新回调，参数为格式化的热键字符串（用于实时显示）
            on_finish: 完成回调，参数为(热键字符串, 错误信息)
        """
        if self.recording:
            return

        self.recording = True
        self.pressed_keys.clear()
        self.released_keys.clear()
        self.all_pressed_keys.clear()
        self.on_update_callback = on_update
        self.on_finish_callback = on_finish

        self._start_hook()
        log("Hotkey recording started")

    def stop_recording(self) -> None:
        """停止录制"""
        self.recording = False
        self._stop_hook()

        self.pressed_keys.clear()
        self.released_keys.clear()
        self.all_pressed_keys.clear()
        log("Hotkey recording stopped")

    # ------------------------------------------------------------------
    # Windows low-level keyboard hook
    # ------------------------------------------------------------------

    def _start_hook(self) -> None:
        import ctypes
        import ctypes.wintypes as wt
        import threading

        WH_KEYBOARD_LL = 13
        WM_KEYDOWN = 0x0100
        WM_SYSKEYDOWN = 0x0104
        WM_KEYUP = 0x0101
        WM_SYSKEYUP = 0x0105

        HOOKPROC = ctypes.WINFUNCTYPE(
            ctypes.c_long,          # LRESULT
            ctypes.c_int,           # nCode
            ctypes.wintypes.WPARAM, # wParam
            ctypes.wintypes.LPARAM, # lParam
        )

        class KBDLLHOOKSTRUCT(ctypes.Structure):
            _fields_ = [
                ("vkCode", wt.DWORD),
                ("scanCode", wt.DWORD),
                ("flags", wt.DWORD),
                ("time", wt.DWORD),
                ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
            ]

        user32 = ctypes.windll.user32

        def _hook_proc(nCode, wParam, lParam):
            if nCode >= 0 and self.recording:
                kb = ctypes.cast(lParam, ctypes.POINTER(KBDLLHOOKSTRUCT)).contents
                vk = kb.vkCode

                if wParam in (WM_KEYDOWN, WM_SYSKEYDOWN):
                    key_name = self._vk_to_name(vk)
                    if key_name:
                        self._on_key_press(key_name)

                elif wParam in (WM_KEYUP, WM_SYSKEYUP):
                    key_name = self._vk_to_name(vk)
                    if key_name:
                        self._on_key_release(key_name)

            return user32.CallNextHookEx(None, nCode, wParam, lParam)

        # Must keep a reference so the callback isn't GC'd
        self._hook_proc_ref = HOOKPROC(_hook_proc)

        def _run():
            self._hook_handle = user32.SetWindowsHookExW(
                WH_KEYBOARD_LL, self._hook_proc_ref, None, 0
            )
            if not self._hook_handle:
                log("Failed to install keyboard hook for recording")
                return

            msg = wt.MSG()
            while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))

        self._hook_thread = threading.Thread(target=_run, daemon=True)
        self._hook_thread.start()

    def _stop_hook(self) -> None:
        if self._hook_handle:
            import ctypes
            ctypes.windll.user32.UnhookWindowsHookEx(self._hook_handle)
            self._hook_handle = None
        # Post WM_QUIT to break the message loop
        if self._hook_thread and self._hook_thread.is_alive():
            import ctypes
            kernel32 = ctypes.windll.kernel32
            # Thread will exit when GetMessage returns 0 (WM_QUIT)
            # We can't easily post WM_QUIT without the thread id, so just let
            # the daemon thread die when recording stops.
            self._hook_thread = None

    @staticmethod
    def _vk_to_name(vk: int) -> Optional[str]:
        """Convert VK code to key name."""
        if vk in _MODIFIER_VKS:
            return _MODIFIER_VKS[vk]
        if vk in _NORMAL_VKS:
            return _NORMAL_VKS[vk]
        return None

    # ------------------------------------------------------------------
    # Key tracking logic (shared)
    # ------------------------------------------------------------------

    def _on_key_press(self, key_name: str) -> None:
        if not self.recording:
            return
        if key_name not in self.pressed_keys:
            self.pressed_keys.add(key_name)
            self.all_pressed_keys.add(key_name)
            self._notify_update()

    def _on_key_release(self, key_name: str) -> None:
        if not self.recording:
            return

        self.released_keys.add(key_name)
        self.pressed_keys.discard(key_name)

        if self.all_pressed_keys and self.all_pressed_keys == self.released_keys:
            self._finish_recording()

    def _notify_update(self) -> None:
        if self.on_update_callback and self.all_pressed_keys:
            display_text = self._format_keys_for_display()
            try:
                self.on_update_callback(display_text)
            except Exception as e:
                log(f"Error in update callback: {e}")

    def _format_keys_for_display(self) -> str:
        if not self.all_pressed_keys:
            return ""

        modifiers = []
        keys = []
        modifier_order = ["ctrl", "shift", "alt", "cmd"]
        for mod in modifier_order:
            if mod in self.all_pressed_keys:
                modifiers.append(mod)
        for key in self.all_pressed_keys:
            if key not in modifier_order:
                keys.append(key)
        all_keys = modifiers + sorted(keys)
        return " + ".join(k.title() for k in all_keys)

    def _finish_recording(self) -> None:
        if not self.all_pressed_keys:
            self.stop_recording()
            if self.on_finish_callback:
                self.on_finish_callback(None, t("hotkey.recorder.error.no_key_detected"))
            return

        error = self._validate_hotkey()
        hotkey_str = None if error else self._generate_hotkey_string()

        self.stop_recording()

        if self.on_finish_callback:
            try:
                self.on_finish_callback(hotkey_str, error)
            except Exception as e:
                log(f"Error in finish callback: {e}")

    def _validate_hotkey(self) -> Optional[str]:
        hotkey_preview = self._format_keys_for_display().replace(" + ", "+")
        return HotkeyChecker.validate_hotkey_keys(
            self.all_pressed_keys,
            hotkey_repr=hotkey_preview,
            detailed=True,
        )

    def _generate_hotkey_string(self) -> str:
        """生成热键字符串（pynput 兼容格式）"""
        modifiers = []
        keys = []
        modifier_order = ["ctrl", "shift", "alt", "cmd"]
        for mod in modifier_order:
            if mod in self.all_pressed_keys:
                modifiers.append(f"<{mod}>")
        for key in self.all_pressed_keys:
            if key not in modifier_order:
                if len(key) > 1:
                    keys.append(f"<{key}>")
                else:
                    keys.append(key)
        return "+".join(modifiers + sorted(keys))
