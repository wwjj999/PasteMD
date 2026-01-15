"""Markdown processing utilities - pure functions without workflow dependencies."""


import re


def merge_markdown_contents(files_data: list[tuple[str, str]]) -> str:
    """
    合并多个 MD 文件内容
    
    Args:
        files_data: [(filename, content), ...] 列表
        
    Returns:
        合并后的 Markdown 内容
        
    Notes:
        - 单文件：直接返回内容
        - 多文件：按原顺序拼接 `<!-- Source: filename -->` 注释 + content.strip() + 空行分隔
    """
    if len(files_data) == 1:
        # 单个文件直接返回内容
        return files_data[0][1]
    
    # 多个文件用 HTML 注释标记来源
    merged_parts = []
    for filename, content in files_data:
        merged_parts.append(f"<!-- Source: {filename} -->")
        merged_parts.append(content.strip())
        merged_parts.append("")  # 空行分隔
    
    return "\n".join(merged_parts)

def has_backtick_fenced_code_block(text: str) -> bool:
    """
    检测 ``` 这种 fenced code block，并要求起始/结束围栏成对出现。
    """
    if not text:
        return False

    pattern = re.compile(
        r'^\s{0,3}(`{3,})[^\n]*\n'   # 开始：``` 或更多反引号，允许 ```python
        r'[\s\S]*?\n'               # 内容（非贪婪）
        r'^\s{0,3}\1\s*$',          # 结束：同样数量的反引号
        re.MULTILINE
    )
    return bool(pattern.search(text))


def has_latex_math(text: str) -> bool:
    """
    检测常见 LaTeX 数学公式：
    行内：$...$ 或 \\( ... \\)
    块级：$$...$$ 或 \\[ ... \\]
    这里对 $...$ 不做内容限制（更宽松，误判风险也更高，比如 $100）。
    """
    if not text:
        return False

    # 块级：$$...$$（允许跨行）
    if re.search(r'\$\$[\s\S]*?\$\$', text):
        return True

    # 块级：\[...\]（允许跨行）
    if re.search(r'\\\[[\s\S]*?\\\]', text):
        return True

    # 行内：\(...\)（不跨行）
    if re.search(r'\\\([^\n]*?\\\)', text):
        return True

    # 行内：$...$（不跨行；排除 $$...$$）
    if re.search(r'(?<!\$)\$(?!\$)[^\n$]+(?<!\$)\$(?!\$)', text):
        return True

    return False

def is_markdown(text: str) -> bool:
    if not text or not isinstance(text, str):
        return False

    if has_backtick_fenced_code_block(text):
        return True

    if has_latex_math(text):
        return True

    md_patterns = [
        r'^\s{0,3}#{1,6}\s+',        # 标题
        r'\[.+?\]\(.+?\)',           # 链接
        r'^\s*[-*+]\s+',             # 无序列表
        r'^\s*\d+\.\s+',             # 有序列表
        r'^>\s+',                    # 引用
        r'`[^`]+`',                  # 行内代码
        r'!\[.*?\]\(.+?\)',          # 图片
        r'(\*\*|__).+?(\*\*|__)',    # 粗体
        r'(\*|_).+?(\*|_)',          # 斜体（可能误判）
    ]

    for p in md_patterns:
        if re.search(p, text, re.MULTILINE):
            return True
    return False