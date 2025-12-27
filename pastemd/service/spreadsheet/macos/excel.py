"""macOS Excel spreadsheet placer (Optimized + Hyperlink + RichText + InlineCode BG)."""

import os
import subprocess
from typing import List, Dict, Any

from ..base import BaseSpreadsheetPlacer
from ..formatting import CellFormat
from ....core.types import PlacementResult
from ....utils.logging import log
from ....i18n import t
from ....config.paths import get_user_data_dir


class ExcelPlacer(BaseSpreadsheetPlacer):
    """macOS Excel 内容落地器（批量操作优化版 + 整格超链接 + 富文本 + 行内代码整格灰底）"""

    def __init__(self):
        temp_dir = os.path.join(get_user_data_dir(), "temp")
        os.makedirs(temp_dir, exist_ok=True)
        self._fixed_script_path = os.path.join(temp_dir, "pastemd_excel_insert.applescript")

    def place(self, table_data: List[List[str]], config: dict) -> PlacementResult:
        try:
            keep_format = config.get("excel_keep_format", config.get("keep_format", True))
            # 可选开关：默认开启“出现行内 code 就整格灰底”
            inline_code_cell_bg = config.get("excel_inline_code_cell_bg", True)

            processed_data = self._process_table_data(
                table_data=table_data,
                keep_format=keep_format,
                inline_code_cell_bg=inline_code_cell_bg,
            )

            success = self._applescript_insert_batch(processed_data, keep_format)

            if success:
                return PlacementResult(success=True, method="applescript")
            raise Exception(t("placer.macos_excel.applescript_failed"))

        except Exception as e:
            log(f"Excel AppleScript 插入失败: {e}")
            return PlacementResult(success=False, method="applescript", error=str(e))

    def _process_table_data(
        self,
        table_data: List[List[str]],
        keep_format: bool,
        inline_code_cell_bg: bool,
    ) -> Dict[str, Any]:
        """预处理数据：data + formats + links（并带 gray_bg / code_block 标记）"""
        rows_count = len(table_data)
        cols_count = max((len(r) for r in table_data), default=0)

        clean_data: List[List[str]] = []
        formats: List[Dict[str, Any]] = []
        links: List[Dict[str, Any]] = []

        for i, row in enumerate(table_data):
            clean_row: List[str] = []

            for j, cell_value in enumerate(row):
                cf = CellFormat(cell_value)
                text = cf.parse()
                clean_row.append(text)

                if not keep_format:
                    continue

                # 是否存在行内代码段
                has_inline_code = bool(cf.segments) and any(seg.is_code for seg in cf.segments)
                # 是否需要整格灰底：code block 一定灰底；或（开启开关时）出现行内 code 就灰底
                gray_bg = bool(cf.is_code_block or (inline_code_cell_bg and has_inline_code))

                # ---- 整格超链接：仅 1 段，且无其它样式（和你 Windows 版一致）----
                if cf.segments and len(cf.segments) == 1:
                    seg = cf.segments[0]
                    if (
                        seg.hyperlink_url
                        and not seg.bold
                        and not seg.italic
                        and not seg.strikethrough
                        and not seg.is_code
                        and not cf.has_newline
                        and not cf.is_code_block
                    ):
                        url = self._normalize_url(seg.hyperlink_url)
                        display = (seg.text or text or "").strip()
                        if url and display:
                            links.append({"r": i + 1, "c": j + 1, "url": url, "display": display})
                        # 超链接格子不做富文本（避免互相踩）
                        # 如果你希望链接格子也能灰底/换行，可以改这里：不要 continue，往下走 formats
                        continue

                # ---- 富文本格式信息（字符级）----
                segments_payload = []
                if cf.segments:
                    char_index = 1  # AppleScript 字符索引从 1 开始
                    for seg in cf.segments:
                        seg_text = seg.text or ""
                        seg_len = len(seg_text)
                        if seg_len <= 0:
                            continue

                        start = char_index
                        end = char_index + seg_len - 1

                        if seg.bold or seg.italic or seg.strikethrough or seg.is_code:
                            segments_payload.append(
                                {
                                    "start": start,
                                    "end": end,
                                    "b": bool(seg.bold),
                                    "i": bool(seg.italic),
                                    "s": bool(seg.strikethrough),
                                    "code": bool(seg.is_code),
                                }
                            )
                        char_index += seg_len

                need_wrap = bool(cf.has_newline or cf.is_code_block)
                needs_rich = bool(segments_payload)  # 有字符级格式才需要强制文本格式

                # 只要：有 wrap / 有富文本 / 有灰底 / 是 code_block，就记录一个 format entry
                if need_wrap or needs_rich or gray_bg or cf.is_code_block:
                    formats.append(
                        {
                            "r": i + 1,
                            "c": j + 1,
                            "wrap": need_wrap,
                            "segments": segments_payload,
                            "needs_rich": needs_rich,
                            "gray_bg": gray_bg,
                            "code_block": bool(cf.is_code_block),
                        }
                    )

            while len(clean_row) < cols_count:
                clean_row.append("")
            clean_data.append(clean_row)

        return {
            "data": clean_data,
            "rows": rows_count,
            "cols": cols_count,
            "formats": formats,
            "links": links,
        }

    def _applescript_insert_batch(self, processed_data: Dict[str, Any], keep_format: bool) -> bool:
        data = processed_data["data"]
        rows = processed_data["rows"]
        cols = processed_data["cols"]
        formats = processed_data.get("formats", []) or []
        links = processed_data.get("links", []) or []

        if rows <= 0 or cols <= 0:
            return True

        # AppleScript list-of-lists: {{"a","b"},{"c","d"}}
        as_data_list = "{" + ",".join(
            ["{" + ",".join([f'"{self._escape_as(cell)}"' for cell in row]) + "}" for row in data]
        ) + "}"

        # ---- 超链接脚本（整格链接）----
        link_cmds: List[str] = []
        for lk in links:
            r = int(lk["r"])
            c = int(lk["c"])
            url = self._escape_as(str(lk["url"]))
            display = self._escape_as(str(lk["display"]))
            link_cmds.append(
                f"""
try
    set theCell to cell (startC + {c - 1}) of row (startR + {r - 1}) of active sheet
    set addrStr to (get address theCell)
    set theRange to range addrStr of active sheet
    make new hyperlink of theRange at active sheet with properties {{address:"{url}", text to display:"{display}"}}
on error errMsg number errNum
    set end of warnMsgs to ("link({r},{c}): (" & errNum & ") " & errMsg)
end try
""".rstrip()
            )
        hyperlink_script = "\n".join(link_cmds)

        # ---- 格式脚本（关键：对 range 做 characters；并支持整格灰底）----
        format_cmds: List[str] = []
        if keep_format:
            # 表头加粗
            format_cmds.append(
                f"""
try
    set bold of font object of (get resize startCell row size 1 column size {cols}) to true
on error errMsg number errNum
    set end of warnMsgs to ("header-bold: (" & errNum & ") " & errMsg)
end try
""".rstrip()
            )

            for f in formats:
                r = int(f["r"])
                c = int(f["c"])
                wrap = bool(f.get("wrap", False))
                segments = f.get("segments") or []
                needs_rich = bool(f.get("needs_rich", False))
                gray_bg = bool(f.get("gray_bg", False))
                code_block = bool(f.get("code_block", False))

                block: List[str] = [
                    f"""
try
    set theCell to cell (startC + {c - 1}) of row (startR + {r - 1}) of active sheet
    set addrStr to (get address theCell)
    set theRange to range addrStr of active sheet
""".rstrip()
                ]

                # 只要要做字符级富文本，就强制文本格式，避免 Excel 自动转数字导致 characters 失败
                if needs_rich:
                    block.append(
                        """
    try
        set number format of theRange to "@"
        set value of theRange to ((value of theRange) as text)
    on error errMsg number errNum
        set end of warnMsgs to ("force-text: (" & errNum & ") " & errMsg)
    end try
""".rstrip()
                    )

                # ✅ 行内 code/代码块：整格灰底
                if gray_bg:
                    block.append(
                        """
    try
        tell interior object of theRange
            set pattern to pattern solid
            set color to {240, 240, 240}
        end tell
    on error errMsg number errNum
        set end of warnMsgs to ("bg: (" & errNum & ") " & errMsg)
    end try
""".rstrip()
                    )

                # 代码块：更像代码块的整体样式（可按需删）
                if code_block:
                    block.append(
                        """
    try
        set name of font object of theRange to "Menlo"
    on error errMsg number errNum
        set end of warnMsgs to ("codeblock-font: (" & errNum & ") " & errMsg)
    end try
    try
        set vertical alignment of theRange to vertical alignment top
    on error errMsg number errNum
        set end of warnMsgs to ("codeblock-vAlign: (" & errNum & ") " & errMsg)
    end try
""".rstrip()
                    )

                if wrap:
                    block.append(
                        """
    try
        set wrap text of theRange to true
    on error errMsg number errNum
        set end of warnMsgs to ("wrap: (" & errNum & ") " & errMsg)
    end try
""".rstrip()
                    )

                # 字符级（粗斜删/行内代码字体）
                block.append(
                    """
    set cellText to ((value of theRange) as text)
    set tLen to (count of characters of cellText)
""".rstrip()
                )

                for seg in segments:
                    start = int(seg["start"])
                    end = int(seg["end"])
                    b = bool(seg.get("b"))
                    i_ = bool(seg.get("i"))
                    s_ = bool(seg.get("s"))
                    code = bool(seg.get("code"))

                    # clamp
                    block.append(
                        f"""
    if tLen > 0 then
        set sIdx to {start}
        set eIdx to {end}
        if sIdx < 1 then set sIdx to 1
        if eIdx > tLen then set eIdx to tLen
        if eIdx ≥ sIdx then
            try
                {"set name of font object of (characters sIdx thru eIdx of theRange) to \"Menlo\"" if code else ""}
                {"set bold of font object of (characters sIdx thru eIdx of theRange) to true" if b else ""}
                {"set italic of font object of (characters sIdx thru eIdx of theRange) to true" if i_ else ""}
                {"set strikethrough of font object of (characters sIdx thru eIdx of theRange) to true" if s_ else ""}
            on error errMsg number errNum
                set end of warnMsgs to ("cell({r},{c}) chars(" & sIdx & "-" & eIdx & "): (" & errNum & ") " & errMsg)
            end try
        end if
    end if
""".rstrip()
                    )

                block.append(
                    f"""
on error errMsg number errNum
    set end of warnMsgs to ("cell({r},{c}) format: (" & errNum & ") " & errMsg)
end try
""".rstrip()
                )

                format_cmds.append("\n".join(block))

        format_script = "\n".join(format_cmds)

        script = f'''
set warnMsgs to {{}}

tell application "Microsoft Excel"
    activate
    if (count of workbooks) is 0 then make new workbook

    try
        set startCell to active cell
    on error
        set startCell to cell 1 of row 1 of active sheet
    end try

    set startR to first row index of startCell
    set startC to first column index of startCell
    set targetRange to (get resize startCell row size {rows} column size {cols})

    -- 批量写入
    set value of targetRange to {as_data_list}

    -- 超链接（整格）
    {hyperlink_script}

    -- 富文本 / wrap / 背景
    {format_script}

    select targetRange
end tell

if (count of warnMsgs) > 0 then
    return "WARN: " & (warnMsgs as string)
else
    return ""
end if
'''

        try:
            with open(self._fixed_script_path, "w", encoding="utf-8") as f:
                f.write(script)

            res = subprocess.run(
                ["osascript", self._fixed_script_path],
                check=True,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=30,
            )
            out = (res.stdout or "").strip()
            if out:
                log(f"Excel AppleScript warnings: {out}")
            return True

        except subprocess.CalledProcessError as e:
            log(f"AppleScript Error: {e.stderr}")
            raise Exception(f"AppleScript Error: {e.stderr}")
        except subprocess.TimeoutExpired:
            raise Exception(t("placer.macos_excel.script_timeout"))

    def _escape_as(self, s: str) -> str:
        if s is None:
            return ""
        return str(s).replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\r").replace("\r", "\\r")

    def _normalize_url(self, url: str) -> str:
        u = (url or "").strip()
        if not u:
            return u
        if u.lower().startswith(("http://", "https://", "mailto:", "ftp://", "file://")):
            return u
        if u.lower().startswith("www."):
            return "https://" + u
        return u
