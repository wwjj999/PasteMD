"""Cross-platform hotkey checker."""

from typing import Optional, Set
from .system_detect import is_windows, is_macos
from .logging import log


class HotkeyChecker:
    """跨平台热键检查器"""
    
    _checker = None
    
    @classmethod
    def _get_checker(cls):
        """获取平台特定的热键检查器"""
        if cls._checker is None:
            if is_windows():
                try:
                    from .win32.hotkey_checker import HotkeyChecker as WinChecker
                    cls._checker = WinChecker
                    log("Using Windows hotkey checker")
                except ImportError as e:
                    log(f"Failed to import Windows hotkey checker: {e}")
                    cls._checker = None
            elif is_macos():
                try:
                    from .macos.hotkey_checker import HotkeyChecker as MacChecker
                    cls._checker = MacChecker
                    log("Using macOS hotkey checker")
                except ImportError as e:
                    log(f"Failed to import macOS hotkey checker: {e}")
                    cls._checker = None
            else:
                log("Unsupported platform for hotkey checking")
                cls._checker = None
        
        return cls._checker
    
    @classmethod
    def validate_hotkey_keys(
        cls,
        keys: Set[str],
        *,
        hotkey_repr: str = "",
        detailed: bool = False,
    ) -> Optional[str]:
        """
        验证热键键集合的有效性。
        返回本地化的错误消息，如果有效则返回 None。
        """
        checker = cls._get_checker()
        if checker is None:
            return None  # 不支持的平台，跳过验证
        
        return checker.validate_hotkey_keys(
            keys,
            hotkey_repr=hotkey_repr,
            detailed=detailed,
        )
    
    @classmethod
    def validate_hotkey_string(cls, hotkey_str: str, *, detailed: bool = False) -> Optional[str]:
        """
        验证 pynput 风格的热键字符串。返回错误文本或 None。
        """
        checker = cls._get_checker()
        if checker is None:
            return None  # 不支持的平台，跳过验证
        
        return checker.validate_hotkey_string(hotkey_str, detailed=detailed)
    
    @classmethod
    def is_hotkey_available(cls, hotkey_str: str) -> bool:
        """
        检查热键是否可用。
        
        Args:
            hotkey_str: pynput 格式的热键字符串
            
        Returns:
            True 表示可用，False 表示不可用
        """
        checker = cls._get_checker()
        if checker is None:
            return True  # 不支持的平台，假设可用
        
        return checker.is_hotkey_available(hotkey_str)

    @classmethod
    def parse_hotkey(cls, hotkey_str: str):
        """
        解析热键字符串为 (modifiers, key_code/vk_code)。

        Returns:
            平台特定的 (modifiers, key) 元组，解析失败返回 None
        """
        checker = cls._get_checker()
        if checker is None:
            return None
        
        return checker.parse_hotkey(hotkey_str)
