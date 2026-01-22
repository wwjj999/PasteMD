"""Base classes for spreadsheet placement."""

import sys
from abc import ABC, abstractmethod
import time
from typing import List
from ...core.types import PlacementResult
from ...utils.logging import log
from ...i18n import t
from ...utils.clipboard import set_clipboard_rich_text, simulate_paste, preserve_clipboard
from .html_converter import table_to_html, table_to_tsv


class BaseSpreadsheetPlacer(ABC):
    """表格内容落地器基类"""
    
    @abstractmethod
    def place(self, table_data: List[List[str]], config: dict) -> PlacementResult:
        """
        将表格数据落地到目标应用
        
        Args:
            table_data: 二维数组表格数据
            config: 配置字典（包含 keep_format 等选项）
            
        Returns:
            PlacementResult: 落地结果
            
        Note:
            ❌ 不做优雅降级,失败即返回错误
            ✅ 由 Workflow 决定如何处理失败(通知用户/记录日志)
        """
        pass


class ClipboardHTMLSpreadsheetPlacer(BaseSpreadsheetPlacer):
    """基于剪贴板 HTML/TSV 粘贴的表格内容落地器通用实现
    
    适用于 Excel、WPS 等支持 HTML 剪贴板格式的表格应用。
    子类只需指定平台检查和国际化 key。
    """
    app_name: str = None  # 如 "macOS Excel"
    
    def place(self, table_data: List[List[str]], config: dict) -> PlacementResult:
        try:
            keep_format = config.get("excel_keep_format", config.get("keep_format", True))
            
            # 使用共享的 HTML 和 TSV 转换工具
            html_text = table_to_html(table_data, keep_format=keep_format)
            tsv_text = table_to_tsv(table_data)

            # Excel/WPS 可以处理 HTML table；Plain TSV 作为兜底
            with preserve_clipboard():
                set_clipboard_rich_text(html=html_text, text=tsv_text)
                time.sleep(0.3)
                simulate_paste()

            return PlacementResult(
                success=True,
                method="clipboard_html_table" if keep_format else "clipboard_tsv",
            )
        except Exception as e:
            log(f"{self.app_name} HTML 粘贴失败: {e}")
            return PlacementResult(
                success=False,
                method="clipboard_html_table",
                error=t(f"{self.i18n_prefix}.insert_failed", error=str(e)),
            )
