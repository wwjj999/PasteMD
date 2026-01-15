"""Cross-platform clipboard operations.

This module provides a unified interface for clipboard operations across different platforms.
It automatically detects the operating system and imports the appropriate implementation.
"""

import sys
from ..core.errors import ClipboardError


# 根据操作系统导入对应的实现
if sys.platform == "darwin":
    from .macos.clipboard import (
        get_clipboard_text,
        set_clipboard_text,
        is_clipboard_empty,
        is_clipboard_html,
        get_clipboard_html,
        set_clipboard_rich_text,
        copy_files_to_clipboard,
        is_clipboard_files,
        get_clipboard_files,
        get_markdown_files_from_clipboard,
        read_markdown_files_from_clipboard,
        preserve_clipboard,
    )
    from .macos.keystroke import simulate_paste
    # read_file_with_encoding 从共享模块导入
    from .clipboard_file_utils import read_file_with_encoding
elif sys.platform == "win32":
    from .win32.clipboard import (
        get_clipboard_text,
        set_clipboard_text,
        is_clipboard_empty,
        is_clipboard_html,
        get_clipboard_html,
        set_clipboard_rich_text,
        copy_files_to_clipboard,
        is_clipboard_files,
        get_clipboard_files,
        get_markdown_files_from_clipboard,
        read_markdown_files_from_clipboard,
        preserve_clipboard,
    )
    from .win32.keystroke import simulate_paste
    # read_file_with_encoding 从共享模块导入
    from .clipboard_file_utils import read_file_with_encoding
else:
    # 其他平台的后备实现（仅支持基本文本功能）
    import pyperclip

    def get_clipboard_text() -> str:
        """
        获取剪贴板文本内容

        Returns:
            剪贴板文本内容

        Raises:
            ClipboardError: 剪贴板操作失败时
        """
        try:
            text = pyperclip.paste()
            if text is None:
                return ""
            return text
        except Exception as e:
            raise ClipboardError(f"Failed to read clipboard: {e}")

    def is_clipboard_empty() -> bool:
        """
        检查剪贴板是否为空

        Returns:
            True 如果剪贴板为空或只包含空白字符
        """
        try:
            text = get_clipboard_text()
            return not text or not text.strip()
        except ClipboardError:
            return True

    def is_clipboard_html() -> bool:
        """
        检查剪切板内容是否为 HTML 富文本

        Note:
            在不支持的平台上始终返回 False

        Returns:
            False (不支持的平台)
        """
        return False

    def get_clipboard_html(config: dict | None = None) -> str:
        """
        获取剪贴板 HTML 富文本内容

        Note:
            在不支持的平台上会抛出异常

        Raises:
            ClipboardError: 不支持的平台
        """
        raise ClipboardError(f"HTML clipboard operations not supported on {sys.platform}")

    def set_clipboard_rich_text(
        *,
        html: str | None = None,
        rtf_bytes: bytes | None = None,
        docx_bytes: bytes | None = None,
        text: str | None = None,
    ) -> None:
        raise ClipboardError(
            f"Rich-text clipboard operations not supported on {sys.platform}"
        )

    def simulate_paste(*, timeout_s: float = 5.0) -> None:
        raise ClipboardError(f"Paste keystroke not supported on {sys.platform}")


# 导出公共接口
__all__ = [
    "get_clipboard_text",
    "set_clipboard_text",
    "is_clipboard_empty",
    "is_clipboard_html",
    "get_clipboard_html",
    "ClipboardError",
]

# 条件导出文件操作/富文本/粘贴快捷键 (Windows 和 macOS)
if sys.platform in ("win32", "darwin"):
    __all__.extend([
        "set_clipboard_rich_text",
        "simulate_paste",
        "copy_files_to_clipboard",
        "is_clipboard_files",
        "get_clipboard_files",
        "get_markdown_files_from_clipboard",
        "read_markdown_files_from_clipboard",
        "read_file_with_encoding",
        "preserve_clipboard",
    ])
