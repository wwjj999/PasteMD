# -*- coding: utf-8 -*-
"""LaTeX paste workflow for Overleaf and other LaTeX editors."""

from .extensible_base import ExtensibleWorkflow
from ....core.errors import ClipboardError, PandocError
from ....utils.clipboard import (
    get_clipboard_html,
    get_clipboard_text,
    is_clipboard_empty,
    set_clipboard_text,
    simulate_paste,
    preserve_clipboard,
)
from ....utils.html_analyzer import is_plain_html_fragment
from ....i18n import t


class LatexWorkflow(ExtensibleWorkflow):
    """LaTeX 粘贴工作流
    
    适用于 Overleaf 等 LaTeX 编辑器：
    - 读取剪贴板 HTML/Markdown（自动识别）
    - 转换为 LaTeX（去除文档头部）
    - 设置剪贴板为纯文本 LaTeX
    - 模拟粘贴
    """
    
    @property
    def workflow_key(self) -> str:
        return "latex"
    
    def execute(self) -> None:
        """执行 LaTeX 粘贴工作流"""
        try:
            # 1. 读取剪贴板内容
            content_type, content = self._read_clipboard()
            self._log(f"LaTeX workflow: content_type={content_type}")
            
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
            
            # 4. 转换为 LaTeX
            latex_text = self.doc_generator.convert_markdown_to_latex_text(
                md_text, self.config
            )
            
            # 5. 设置剪贴板为纯文本 LaTeX
            with preserve_clipboard():
                set_clipboard_text(latex_text)
                self._log("Set clipboard with LaTeX text")
                
                # 6. 模拟粘贴
                simulate_paste()
            
            # 7. 通知成功
            self._notify_success(t("workflow.latex.paste_success"))
            
        except ClipboardError as e:
            self._log(f"Clipboard error: {e}")
            self._notify_error(t("workflow.clipboard.read_failed"))
        except PandocError as e:
            self._log(f"Pandoc error: {e}")
            self._notify_error(t("workflow.html.convert_failed_generic"))
        except Exception as e:
            self._log(f"LaTeX workflow failed: {e}")
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
