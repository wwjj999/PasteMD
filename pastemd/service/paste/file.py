import time
from typing import Optional

from ...core.types import PlacementResult
from ...utils.clipboard import copy_files_to_clipboard, simulate_paste, preserve_clipboard
from ...utils.logging import log
from .base import BasePastePlacer


class FilePastePlacer(BasePastePlacer):
    """文件粘贴落地器

    将文件路径写入剪贴板并模拟粘贴，适用于支持文件粘贴的应用。
    """

    def place(
        self,
        content: str,
        config: dict,
        file_paths: Optional[list[str]] = None,
        **kwargs
    ) -> PlacementResult:
        """
        使用文件格式粘贴内容

        Args:
            content: 单个文件路径（兼容 BasePastePlacer 的签名）
            config: 配置字典
            file_paths: 文件路径列表

        Returns:
            PlacementResult: 落地结果
        """
        paths = list(file_paths or [])
        if not paths and content:
            paths = [content]

        if not paths:
            return PlacementResult(
                success=False,
                method="clipboard_file",
                error="no_file_paths",
            )

        try:
            with preserve_clipboard():
                copy_files_to_clipboard(paths)
                time.sleep(0.1)
                simulate_paste()

            return PlacementResult(
                success=True,
                method="clipboard_file",
                metadata={"count": len(paths)},
            )
        except Exception as e:
            log(f"文件粘贴失败: {e}")
            return PlacementResult(
                success=False,
                method="clipboard_file",
                error=str(e),
            )
