import time
from typing import Optional
from ...core.types import PlacementResult
from ...utils.clipboard import set_clipboard_text, simulate_paste, preserve_clipboard
from ...utils.logging import log
from .base import BasePastePlacer

class PlainTextPastePlacer(BasePastePlacer):
    """纯文本粘贴落地器

    仅使用纯文本格式粘贴，适用于 Obsidian、Overleaf 等纯文本编辑器。
    """

    def place(
        self,
        content: str,
        config: dict,
        **kwargs
    ) -> PlacementResult:
        """
        使用纯文本格式粘贴内容

        Args:
            content: 纯文本内容
            config: 配置字典

        Returns:
            PlacementResult: 落地结果
        """
        try:
            with preserve_clipboard():
                set_clipboard_text(content)
                time.sleep(0.1)
                simulate_paste()

            return PlacementResult(
                success=True,
                method="clipboard_plain_text",
            )
        except Exception as e:
            log(f"纯文本粘贴失败: {e}")
            return PlacementResult(
                success=False,
                method="clipboard_plain_text",
                error=str(e),
            )