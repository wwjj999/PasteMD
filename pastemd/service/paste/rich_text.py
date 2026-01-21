import time
from typing import Optional
from ...core.types import PlacementResult
from ...utils.clipboard import set_clipboard_rich_text, simulate_paste, preserve_clipboard
from ...utils.logging import log
from .base import BasePastePlacer

class RichTextPastePlacer(BasePastePlacer):
    """富文本粘贴落地器

    支持 HTML + 纯文本格式的粘贴，适用于 Notion、语雀等笔记应用。
    """

    def place(
        self,
        content: str,
        config: dict,
        html: Optional[str] = None,
        **kwargs
    ) -> PlacementResult:
        """
        使用富文本格式粘贴内容

        Args:
            content: 纯文本内容（作为后备）
            config: 配置字典
            html: HTML 格式内容（可选）

        Returns:
            PlacementResult: 落地结果
        """
        try:
            with preserve_clipboard():
                set_clipboard_rich_text(html=html, text=content)
                time.sleep(0.1)
                simulate_paste()

            return PlacementResult(
                success=True,
                method="clipboard_rich_text",
                metadata={"has_html": html is not None}
            )
        except Exception as e:
            log(f"富文本粘贴失败: {e}")
            return PlacementResult(
                success=False,
                method="clipboard_rich_text",
                error=str(e),
            )
