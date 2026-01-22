"""Shared HTML and TSV conversion utilities for macOS spreadsheet placers."""

from __future__ import annotations

from html import escape
from typing import List, Tuple

from .formatting import CellFormat


def wrap_tag(tag: str, content: str) -> str:
    """Wrap content with HTML tag."""
    return f"<{tag}>{content}</{tag}>"


def cell_to_html(cell_value: str, *, keep_format: bool) -> Tuple[str, bool]:
    """
    Convert a Markdown-ish cell payload to HTML.

    Args:
        cell_value: The cell content (possibly with markdown formatting).
        keep_format: Whether to preserve formatting (bold, italic, code, etc.).

    Returns:
        A tuple of (html_content, needs_code_bg):
        - html_content: The HTML representation of the cell.
        - needs_code_bg: Whether the cell needs code background styling.
    """
    cf = CellFormat(cell_value)
    clean_text = cf.parse()

    # If not keeping format, just escape and convert newlines
    if not keep_format:
        return escape(clean_text).replace("\n", "<br />"), False

    # Handle code block cells
    if cf.is_code_block:
        inner = escape(clean_text)
        inner = inner.replace("\n", "<br />")
        return wrap_tag("code", inner), True

    # Process inline formatting segments
    parts: List[str] = []
    needs_code_bg = False

    for seg in cf.segments:
        seg_text = escape(seg.text or "").replace("\n", "<br />")
        chunk = seg_text

        # Apply formatting in order: code -> strikethrough -> italic -> bold -> hyperlink
        if seg.is_code:
            needs_code_bg = True
            chunk = wrap_tag("code", chunk)
        if seg.strikethrough:
            chunk = wrap_tag("s", chunk)
        if seg.italic:
            chunk = wrap_tag("i", chunk)
        if seg.bold:
            chunk = wrap_tag("b", chunk)
        if seg.hyperlink_url:
            url = escape(seg.hyperlink_url, quote=True)
            chunk = f'<a href="{url}">{chunk}</a>'

        parts.append(chunk)

    return "".join(parts) or escape(clean_text), needs_code_bg


def table_to_html(table_data: List[List[str]], *, keep_format: bool) -> str:
    """
    Convert table data to HTML table format.

    Args:
        table_data: 2D list of cell values.
        keep_format: Whether to preserve formatting.

    Returns:
        Complete HTML document with table.
    """
    rows_html: List[str] = []
    start_marker = "<!--StartFragment-->"
    end_marker = "<!--EndFragment-->"

    for r, row in enumerate(table_data):
        cell_tag = "th" if r == 0 else "td"
        cell_html: List[str] = []

        for cell_value in row:
            content_html, needs_code_bg = cell_to_html(
                cell_value, keep_format=keep_format
            )
            
            # Build style attributes
            style_parts = ["padding:2px 6px", "vertical-align:middle"]
            
            # Header row styling
            if r == 0:
                style_parts.extend(["font-weight:bold", "background-color:#D3D3D3"])
            
            # Code background styling
            if needs_code_bg:
                style_parts.extend([
                    "background-color:#F0F0F0",
                    "font-family:Menlo,Consolas,monospace"
                ])
            
            style_attr = ";".join(style_parts)
            cell_html.append(
                f"<{cell_tag} style=\"{style_attr}\">{content_html}</{cell_tag}>"
            )

        rows_html.append("<tr>" + "".join(cell_html) + "</tr>")

    # Build complete HTML document
    return (
        start_marker +
        "<html><head><meta charset=\"utf-8\" />"
        "<style>"
        "table{border-collapse:collapse}"
        "td,th{border:1px solid #D0D0D0}"
        "a{color:#0563C1;text-decoration:underline}"
        "</style>"
        "</head><body>"
        "<table>"
        + "".join(rows_html)
        + "</table>"
        "</body></html>"
        + end_marker
    )


def table_to_tsv(table_data: List[List[str]]) -> str:
    """
    Convert table data to TSV (Tab-Separated Values) format.

    Args:
        table_data: 2D list of cell values.

    Returns:
        TSV formatted string.
    """
    lines: List[str] = []
    
    for row in table_data:
        out_cells: List[str] = []
        for cell_value in row:
            cf = CellFormat(cell_value)
            text = cf.parse()
            # Replace newlines with spaces for TSV format
            text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", " ")
            out_cells.append(text)
        lines.append("\t".join(out_cells))
    
    return "\n".join(lines)
