"""HTML content preprocessor."""

from bs4 import BeautifulSoup
from .base import BasePreprocessor
from ...utils.html_formatter import (
    clean_html_content,
    convert_css_font_to_semantic,
    convert_strikethrough_to_del,
    promote_bold_first_row_to_header,
)
from ...utils.logging import log


class HtmlPreprocessor(BasePreprocessor):
    """HTML 内容预处理器（无状态）"""

    def process(self, html: str, config: dict) -> str:
        """
        预处理 HTML 内容

        处理步骤:
        1. 清理无效元素（SVG等）
        2. 转换删除线标记
        3. 清理 LaTeX 公式块中的 br 标签
        4. 其他自定义处理...

        Args:
            html: 原始 HTML 内容
            config: 配置字典

        Returns:
            预处理后的 HTML 内容
        """
        log("Preprocessing HTML content")

        # 使用 html_formatter 进行清理
        soup = BeautifulSoup(html, "html.parser")
        clean_html_content(soup, config)

        html_formatting = config.get("html_formatting") or config.get("Html_formatting") or {}
        if not isinstance(html_formatting, dict):
            html_formatting = {}
        if html_formatting.get("strikethrough_to_del", True):
            convert_strikethrough_to_del(soup)
        if html_formatting.get("css_font_to_semantic", True):
            convert_css_font_to_semantic(soup)
        if html_formatting.get("bold_first_row_to_header", False):
            promote_bold_first_row_to_header(soup)

        # unwrap_li_paragraphs(soup)
        # remove_empty_paragraphs(soup)

        html_output = str(soup)
        
        # 仅在 HTML 不包含 DOCTYPE 时才添加
        if "<!DOCTYPE" not in html_output.upper():
            html_output = f"<!DOCTYPE html>\n<meta charset='utf-8'>\n{html_output}"
        
        return html_output
