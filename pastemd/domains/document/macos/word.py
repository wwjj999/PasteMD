"""macOS Word document placer."""

import subprocess
import os
from ..base import BaseDocumentPlacer
from ....core.types import PlacementResult
from ....utils.logging import log
from ....i18n import t
from ....config.paths import get_user_data_dir


class WordPlacer(BaseDocumentPlacer):
    """macOS Word 内容落地器"""
    
    def __init__(self):
        """初始化并创建固定的临时文件路径"""
        super().__init__()
        # 使用固定的临时文件路径，避免每次都需要授权
        temp_dir = os.path.join(get_user_data_dir(), "temp")
        os.makedirs(temp_dir, exist_ok=True)
        self._fixed_temp_path = os.path.join(temp_dir, "pastemd_word_insert.docx")
        log(f"Word 临时文件路径: {self._fixed_temp_path}")
    
    def place(self, docx_bytes: bytes, config: dict) -> PlacementResult:
        """通过 AppleScript 插入,失败不降级"""
        try:
            # 使用固定路径写入临时文件（覆盖之前的）
            with open(self._fixed_temp_path, 'wb') as f:
                f.write(docx_bytes)
            
            move_cursor_to_end = config.get("move_cursor_to_end", True)
            success = self._applescript_insert(self._fixed_temp_path, move_cursor_to_end)
            
            if success:
                return PlacementResult(success=True, method="applescript")
            else:
                raise Exception(t("placer.macos_word.applescript_failed"))
        
        except Exception as e:
            log(f"Word AppleScript 插入失败: {e}")
            return PlacementResult(
                success=False,
                method="applescript",
                error=t("placer.macos_word.insert_failed", error=str(e))
            )
    
    def _applescript_insert(self, docx_path: str, move_cursor_to_end: bool = True) -> bool:
        """
        使用 AppleScript 插入文档
        
        Args:
            docx_path: DOCX 文件路径
            move_cursor_to_end: 插入后是否将光标移动到插入内容的末尾
            
        Returns:
            True 如果插入成功
            
        Raises:
            Exception: 插入失败时
        """
        # 将路径转换为 POSIX 格式
        posix_path = os.path.abspath(docx_path)
        
        # 根据 move_cursor_to_end 配置生成不同的 AppleScript
        if move_cursor_to_end:
            # 光标移动到文档末尾再插入
            script = f'''
            tell application "Microsoft Word"
                activate
                if (count of documents) is 0 then
                    make new document
                end if
                
                tell active document
                    -- 创建一个指向文档末尾的 range
                    set docEnd to count of characters of content
                    set myRange to create range start docEnd end docEnd
                    -- 在末尾插入文件
                    insert file at myRange file name "{posix_path}"
                    -- 尝试将光标移动到插入内容的末尾（失败不应影响插入结果）
                    try
                        set insertEnd to end of myRange
                        select (create range start insertEnd end insertEnd)
                    end try
                end tell
            end tell
            '''
        else:
            # 在当前光标位置插入
            script = f'''
            tell application "Microsoft Word"
                activate
                if (count of documents) is 0 then
                    make new document
                end if
                
                tell active document
                    -- 在当前光标位置插入
                    set curSel to selection
                    set curRange to text object of curSel
                    insert file at curRange file name "{posix_path}"
                end tell
            end tell
            '''
        
        try:
            subprocess.run(
                ["osascript", "-e", script],
                check=True,
                capture_output=True,
                text=True,
                timeout=30
            )
            log(f"AppleScript 插入成功: {docx_path} (move_cursor_to_end={move_cursor_to_end})")
            return True
        except subprocess.CalledProcessError as e:
            log(f"AppleScript 执行失败: {e.stderr}")
            raise Exception(f"AppleScript 错误: {e.stderr}")
        except subprocess.TimeoutExpired:
            log("AppleScript 执行超时")
            raise Exception(t("placer.macos_word.timeout"))
