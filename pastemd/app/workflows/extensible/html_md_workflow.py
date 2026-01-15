# -*- coding: utf-8 -*-
"""HTML+Markdown paste workflow for note-taking apps."""

import os
import subprocess
import re

from .extensible_base import ExtensibleWorkflow
from ....core.errors import ClipboardError, PandocError
from ....utils.clipboard import (
    get_clipboard_html,
    get_clipboard_text,
    is_clipboard_empty,
    set_clipboard_rich_text,
    simulate_paste,
    preserve_clipboard,
)
from ....utils.html_analyzer import is_plain_html_fragment
from ....config.paths import resource_path
from ....i18n import t

class HtmlWorkflow(ExtensibleWorkflow):
    """HTML 粘贴工作流
    
    适用于 Notion、语雀等笔记软件：
    - 读取剪贴板 HTML/Markdown
    - 转换为 Markdown
    - 再转换为纯净 HTML（公式可保持 $...$ 格式）
    - 同时设置剪贴板的 HTML 和纯文本格式
    - 模拟粘贴
    """
    
    @property
    def workflow_key(self) -> str:
        return "html"
    
    def execute(self) -> None:
        """执行 HTML+MD 粘贴工作流"""
        try:
            # 1. 读取剪贴板内容
            content_type, content = self._read_clipboard()
            self._log(f"HTML+MD workflow: content_type={content_type}")
            
            # 2. 转换为 Markdown（如果是 HTML）
            if content_type == "html":
                content = self.html_preprocessor.process(content, self.config)
                md_text = self.doc_generator.convert_html_to_markdown_text(
                    content, self.config
                )
            else:
                md_text = content
            
            # 3. 预处理 Markdown
            md_text = self.markdown_preprocessor.process(md_text, self.config)
            # 4. 转换为纯净 HTML
            keep_formula = self.workflow_config.get("keep_formula_latex", True)
            html_text = self.doc_generator.convert_markdown_to_html_text(
                md_text, 
                {
                    **self.config,
                    "Keep_original_formula": keep_formula, 
                }
            )
            # 5. 设置剪贴板（HTML + 纯文本 Markdown）
            with preserve_clipboard():
                set_clipboard_rich_text(html=html_text, text=md_text)
                self._log("Set clipboard with HTML and plain text")
                
                # 6. 模拟粘贴
                simulate_paste()
            
            # 7. 通知成功
            self._notify_success(t("workflow.html_md.paste_success"))
            
        except ClipboardError as e:
            self._log(f"Clipboard error: {e}")
            self._notify_error(t("workflow.clipboard.read_failed"))
        except PandocError as e:
            self._log(f"Pandoc error: {e}")
            self._notify_error(t("workflow.html.convert_failed_generic"))
        except Exception as e:
            self._log(f"HTML+MD workflow failed: {e}")
            import traceback
            traceback.print_exc()
            self._notify_error(t("workflow.generic.failure"))
    
    def _read_clipboard(self) -> tuple[str, str]:
        """读取剪贴板内容，返回 (类型, 内容)"""
        # 优先尝试 HTML
        try:
            html = get_clipboard_html(self.config)
            if not is_plain_html_fragment(html):
                return ("html", html)
        except ClipboardError:
            pass
        
        # 尝试纯文本
        if not is_clipboard_empty():
            return ("markdown", get_clipboard_text())
        
        raise ClipboardError("剪贴板为空或无有效内容")
    
    def _strip_html_wrapper(self, html: str) -> str:
        """移除 Pandoc 生成的 HTML 包装标签（如 DOCTYPE、html、body）"""
        # 尝试提取 body 内容
        body_match = re.search(r"<body[^>]*>(.*?)</body>", html, re.DOTALL | re.IGNORECASE)
        if body_match:
            return body_match.group(1).strip()
        
        # 移除 DOCTYPE 和 html/head 标签
        html = re.sub(r"<!DOCTYPE[^>]*>", "", html, flags=re.IGNORECASE)
        html = re.sub(r"<html[^>]*>|</html>", "", html, flags=re.IGNORECASE)
        html = re.sub(r"<head[^>]*>.*?</head>", "", html, flags=re.DOTALL | re.IGNORECASE)
        html = re.sub(r"<body[^>]*>|</body>", "", html, flags=re.IGNORECASE)
        
        return html.strip()

