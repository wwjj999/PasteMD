"""Windows clipboard operations using win32clipboard."""

import os
import re
import sys
import time
import contextlib
import pyperclip
import ctypes
from ctypes import wintypes
import win32clipboard as wc
from ...core.errors import ClipboardError
from ...core.state import app_state
from ..clipboard_file_utils import read_file_with_encoding, filter_markdown_files, read_markdown_files
from ...utils.logging import log
from ...core.constants import CLIPBOARD_HTML_WAIT_MS, CLIPBOARD_POLL_INTERVAL_MS


def _snapshot_clipboard() -> dict[int, bytes]:
    """
    Best-effort snapshot of all clipboard formats.
    
    Returns a dict of {format_id: data_bytes}.
    """
    snapshot: dict[int, bytes] = {}
    try:
        wc.OpenClipboard(None)
        try:
            # 枚举所有可用的格式
            fmt = 0
            while True:
                fmt = wc.EnumClipboardFormats(fmt)
                if fmt == 0:
                    break
                try:
                    data = wc.GetClipboardData(fmt)
                    if data is not None:
                        # 转换为 bytes
                        if isinstance(data, bytes):
                            snapshot[fmt] = data
                        elif isinstance(data, str):
                            snapshot[fmt] = data.encode('utf-16le')
                        elif isinstance(data, (list, tuple)):
                            # CF_HDROP 文件列表
                            snapshot[fmt] = "\0".join(data).encode('utf-16le') + b'\0\0'
                        else:
                            # 尝试转换其他类型
                            try:
                                snapshot[fmt] = bytes(data)
                            except Exception:
                                pass
                except Exception as e:
                    log(f"Failed to snapshot clipboard format {fmt}: {e}")
                    continue
        finally:
            wc.CloseClipboard()
    except Exception as e:
        log(f"Failed to snapshot clipboard: {e}")
    
    return snapshot


def _restore_clipboard(snapshot: dict[int, bytes]) -> None:
    """
    Restore clipboard from snapshot.
    """
    try:
        wc.OpenClipboard(None)
        try:
            wc.EmptyClipboard()
            
            for fmt, data in snapshot.items():
                try:
                    # 特殊处理某些格式
                    if fmt == wc.CF_UNICODETEXT:
                        # Unicode 文本需要解码后设置
                        text = data.decode('utf-16le', errors='ignore').rstrip('\0')
                        wc.SetClipboardData(fmt, text)
                    elif fmt == wc.CF_TEXT:
                        # ANSI 文本
                        text = data.decode('cp1252', errors='ignore').rstrip('\0')
                        wc.SetClipboardData(fmt, text)
                    elif fmt == wc.CF_HDROP:
                        # 文件列表，需要重建 DROPFILES 结构
                        files_str = data.decode('utf-16le', errors='ignore').rstrip('\0')
                        files = [f for f in files_str.split('\0') if f]
                        if files:
                            hdrop_data = _build_hdrop_data(files)
                            wc.SetClipboardData(fmt, hdrop_data)
                    else:
                        # 其他格式直接设置原始字节
                        wc.SetClipboardData(fmt, data)
                except Exception as e:
                    log(f"Failed to restore clipboard format {fmt}: {e}")
                    continue
        finally:
            wc.CloseClipboard()
    except Exception as e:
        log(f"Failed to restore clipboard: {e}")


@contextlib.contextmanager
def preserve_clipboard(*, restore_delay_s: float = 0.25):
    """
    Preserve the user's clipboard across a temporary clipboard write.
    
    Useful for apps that require clipboard-based rich-text paste.
    """
    snapshot: dict[int, bytes] | None = None
    try:
        snapshot = _snapshot_clipboard()
        yield
    finally:
        if restore_delay_s > 0:
            time.sleep(restore_delay_s)
        if snapshot is not None:
            try:
                _restore_clipboard(snapshot)
            except Exception as exc:
                log(f"Failed to restore clipboard: {exc}")


def _build_cf_html(html: str) -> bytes:
    """
    Build CF_HTML payload bytes for the Windows clipboard ("HTML Format").

    CF_HTML is an ASCII header + UTF-8 HTML. Offsets are byte offsets from the
    start of the payload.
    """
    start_marker = "<!--StartFragment-->"
    end_marker = "<!--EndFragment-->"

    if start_marker in html and end_marker in html:
        html_doc = html
    else:
        html_doc = (
            "<html><head><meta charset=\"utf-8\"></head><body>"
            f"{start_marker}{html}{end_marker}"
            "</body></html>"
        )

    html_bytes = html_doc.encode("utf-8")

    header_template = (
        "Version:1.0\r\n"
        "StartHTML:{:010d}\r\n"
        "EndHTML:{:010d}\r\n"
        "StartFragment:{:010d}\r\n"
        "EndFragment:{:010d}\r\n"
    )

    # Header length is stable because we always format 10-digit offsets.
    header_placeholder = header_template.format(0, 0, 0, 0).encode("ascii")
    start_html = len(header_placeholder)
    end_html = start_html + len(html_bytes)

    start_marker_b = start_marker.encode("ascii")
    end_marker_b = end_marker.encode("ascii")
    sf_index = html_bytes.find(start_marker_b)
    ef_index = html_bytes.find(end_marker_b)
    if sf_index == -1 or ef_index == -1 or ef_index < sf_index:
        start_fragment = start_html
        end_fragment = end_html
    else:
        start_fragment = start_html + sf_index + len(start_marker_b)
        end_fragment = start_html + ef_index

    header = header_template.format(start_html, end_html, start_fragment, end_fragment).encode(
        "ascii"
    )
    return header + html_bytes


def set_clipboard_rich_text(
    *,
    html: str | None = None,
    rtf_bytes: bytes | None = None,
    docx_bytes: bytes | None = None,
    text: str | None = None,
) -> None:
    """
    Write rich-text clipboard content (HTML/RTF/Plain) on Windows.

    Note:
        - The target application will choose its preferred available format.
        - `docx_bytes` is currently ignored on Windows (no reliable standard clipboard format).
    """
    try:
        fmt_html = wc.RegisterClipboardFormat("HTML Format")
        fmt_rtf = wc.RegisterClipboardFormat("Rich Text Format")

        wc.OpenClipboard(None)
        try:
            wc.EmptyClipboard()

            if html is not None:
                wc.SetClipboardData(fmt_html, _build_cf_html(html))
                log(f"set HTML type=HTML Format len={len(html.encode('utf-8'))}")

            if rtf_bytes is not None:
                wc.SetClipboardData(fmt_rtf, rtf_bytes)
                log(f"set RTF type=Rich Text Format len={len(rtf_bytes)}")

            if text is not None:
                wc.SetClipboardData(wc.CF_UNICODETEXT, text)
                log(f"set PLAIN type=CF_UNICODETEXT len={len(text)}")

            if docx_bytes is not None:
                log(f"docx_bytes provided (len={len(docx_bytes)}), ignored on Windows clipboard")
        finally:
            wc.CloseClipboard()
    except Exception as e:
        raise ClipboardError(f"Failed to write rich text to clipboard: {e}") from e


def _try_read_cf_html(wait_ms: int, interval_ms: int) -> bytes | str | None:
    """
    在短窗口内轮询读取 CF_HTML（"HTML Format"）。

    目的：
    - 避免 OpenClipboard 竞争/占用导致的瞬时失败
    - 避免延迟渲染（IsClipboardFormatAvailable 初始为 False）导致的误判

    Returns:
        CF_HTML 原始数据（bytes 或 str），失败返回 None（不抛异常）。
    """
    try:
        fmt = wc.RegisterClipboardFormat("HTML Format")
    except Exception:
        return None

    deadline = time.monotonic() + (wait_ms / 1000.0)
    interval_s = max(1, interval_ms) / 1000.0

    last_error: Exception | None = None
    while time.monotonic() < deadline:
        try:
            wc.OpenClipboard(None)
        except Exception as exc:
            last_error = exc
            time.sleep(interval_s)
            continue

        try:
            try:
                available: bool | None = bool(wc.IsClipboardFormatAvailable(fmt))
            except Exception as exc:
                last_error = exc
                available = None

            if available is False:
                return None

            if available:
                try:
                    data = wc.GetClipboardData(fmt)
                    if data is None:
                        last_error = ValueError("CF_HTML data is None")
                    else:
                        return data
                except Exception as exc:
                    last_error = exc
        finally:
            try:
                wc.CloseClipboard()
            except Exception:
                pass

        time.sleep(interval_s)

    if last_error is not None:
        log(f"CF_HTML read timed out after {wait_ms}ms: {last_error}")
    return None


def get_clipboard_text() -> str:
    """
    获取剪贴板文本内容
    
    Returns:
        剪贴板文本内容
        
    Raises:
        ClipboardError: 剪贴板操作失败时
    """
    try:
        text = pyperclip.paste()
        if text is None:
            return ""
        return text
    except Exception as e:
        raise ClipboardError(f"Failed to read clipboard: {e}")


def set_clipboard_text(text: str) -> None:
    """
    设置剪贴板纯文本内容
    
    Args:
        text: 要设置的文本内容
        
    Raises:
        ClipboardError: 剪贴板操作失败时
    """
    try:
        pyperclip.copy(text)
    except Exception as e:
        raise ClipboardError(f"Failed to set clipboard text: {e}")


def is_clipboard_empty() -> bool:
    """
    检查剪贴板是否为空
    
    Returns:
        True 如果剪贴板为空或只包含空白字符
    """
    try:
        if is_clipboard_files():
            return False
        text = get_clipboard_text()
        return not text or not text.strip()
    except ClipboardError:
        return True


def is_clipboard_html() -> bool:
    """
    检查剪贴板内容是否为 HTML 富文本 (CF_HTML / "HTML Format")

    Returns:
        True 如果剪贴板中存在 HTML 富文本格式；否则 False
    """
    data = _try_read_cf_html(CLIPBOARD_HTML_WAIT_MS, CLIPBOARD_POLL_INTERVAL_MS)
    return data is not None and data != "" and data != b""


def get_clipboard_html(config: dict | None = None) -> str:
    """
    获取剪贴板 HTML 富文本内容，并清理 SVG 等不可用内容
    
    返回 CF_HTML 格式中的 Fragment 部分（实际网页复制的内容），
    并自动移除 <svg> 标签和 .svg 图片引用。

    Returns:
        清理后的 HTML 富文本内容

    Raises:
        ClipboardError: 剪贴板操作失败时
    """
    config = config or getattr(app_state, "config", {})

    data = _try_read_cf_html(CLIPBOARD_HTML_WAIT_MS, CLIPBOARD_POLL_INTERVAL_MS)
    if data is None or data == "" or data == b"":
        raise ClipboardError("No HTML format data in clipboard")

    try:
        # 解析 CF_HTML 格式，提取 Fragment
        if isinstance(data, bytes):
            fragment = _extract_html_fragment_bytes(data)
        else:
            fragment = _extract_html_fragment(data)

        # 直接返回原始 HTML Fragment，不在剪贴板层进行清理
        return fragment
    except Exception as e:
        raise ClipboardError(f"Failed to read HTML from clipboard: {e}") from e


def _extract_html_fragment_bytes(cf_html_bytes: bytes) -> str:
    """
    从 CF_HTML bytes 中提取 Fragment（StartFragment/EndFragment 通常为字节偏移）。

    - 优先按 bytes 偏移截取，避免非 ASCII 导致的 str 偏移错位。
    - 失败时回退到 <!--StartFragment--> 锚点提取。
    """
    meta: dict[str, str] = {}
    for raw_line in cf_html_bytes.splitlines():
        stripped = raw_line.strip()
        if stripped.startswith(b"<!--"):
            break
        if b":" in raw_line:
            k, v = raw_line.split(b":", 1)
            key = k.decode("ascii", errors="ignore").strip()
            val = v.decode("ascii", errors="ignore").strip()
            if key:
                meta[key] = val

    sf = meta.get("StartFragment", "")
    ef = meta.get("EndFragment", "")
    if sf.isdigit() and ef.isdigit():
        try:
            start_fragment = int(sf)
            end_fragment = int(ef)
            if 0 <= start_fragment <= end_fragment <= len(cf_html_bytes):
                return cf_html_bytes[start_fragment:end_fragment].decode("utf-8", errors="ignore")
        except Exception:
            pass

    m = re.search(
        rb"<!--StartFragment-->(.*)<!--EndFragment-->",
        cf_html_bytes,
        flags=re.S,
    )
    if m:
        return m.group(1).decode("utf-8", errors="ignore")

    start_html = meta.get("StartHTML", "0")
    end_html = meta.get("EndHTML", str(len(cf_html_bytes)))
    try:
        start = int(start_html) if start_html.isdigit() else 0
        end = int(end_html) if end_html.isdigit() else len(cf_html_bytes)
        if 0 <= start <= end <= len(cf_html_bytes):
            return cf_html_bytes[start:end].decode("utf-8", errors="ignore")
    except Exception:
        pass

    return cf_html_bytes.decode("utf-8", errors="ignore")


def _extract_html_fragment(cf_html: str) -> str:
    """
    从 CF_HTML 格式中提取 Fragment 部分
    
    Args:
        cf_html: CF_HTML 格式的完整文本
        
    Returns:
        Fragment HTML 内容
    """
    # 提取元数据
    meta = {}
    for line in cf_html.splitlines():
        if line.strip().startswith("<!--"):
            break
        if ":" in line:
            k, v = line.split(":", 1)
            meta[k.strip()] = v.strip()
    
    # 尝试使用偏移量提取 Fragment
    sf = meta.get("StartFragment")
    ef = meta.get("EndFragment")
    if sf and ef and sf.isdigit() and ef.isdigit():
        try:
            start_fragment = int(sf)
            end_fragment = int(ef)
            return cf_html[start_fragment:end_fragment]
        except Exception:
            pass
    
    # 兜底：使用注释锚点提取
    m = re.search(r"<!--StartFragment-->(.*)<!--EndFragment-->", cf_html, flags=re.S)
    if m:
        return m.group(1)
    
    # 再兜底：提取完整 HTML
    start_html = int(meta.get("StartHTML", "0"))
    end_html = int(meta.get("EndHTML", str(len(cf_html))))
    try:
        return cf_html[start_html:end_html]
    except Exception:
        return cf_html


def copy_files_to_clipboard(file_paths: list) -> None:
    """
    将文件路径复制到剪贴板（CF_HDROP 格式）
    
    Args:
        file_paths: 文件路径列表
        
    Raises:
        ClipboardError: 剪贴板操作失败时
    """
    try:
        # 确保文件路径是绝对路径
        absolute_paths = [os.path.abspath(path) for path in file_paths if os.path.exists(path)]
        
        if not absolute_paths:
            raise ClipboardError("No valid files to copy to clipboard")
        
        # 使用最简单可靠的方法
        _copy_files_simple(absolute_paths)
        
    except Exception as e:
        raise ClipboardError(f"Failed to copy files to clipboard: {e}")


def _build_hdrop_data(file_paths: list) -> bytes:
    """构建 CF_HDROP 格式的数据"""
    # 定义 DROPFILES 结构体
    class DROPFILES(ctypes.Structure):
        _fields_ = [
            ("pFiles", wintypes.DWORD),
            ("pt", wintypes.POINT),
            ("fNC", wintypes.BOOL),
            ("fWide", wintypes.BOOL),
        ]

    # 准备文件列表数据 (UTF-16LE, double-null terminated)
    # 路径之间用 \0 分隔，整个列表以 \0\0 结尾
    files_text = "\0".join(file_paths) + "\0\0"
    files_data = files_text.encode("utf-16le")
    
    # 计算结构体大小
    struct_size = ctypes.sizeof(DROPFILES)
    
    # 创建缓冲区
    total_size = struct_size + len(files_data)
    buf = ctypes.create_string_buffer(total_size)
    
    # 填充结构体
    dropfiles = DROPFILES.from_buffer(buf)
    dropfiles.pFiles = struct_size  # 文件列表数据紧跟在结构体之后
    dropfiles.pt = wintypes.POINT(0, 0)
    dropfiles.fNC = False
    dropfiles.fWide = True  # 使用宽字符 (UTF-16)
    
    # 填充文件数据
    # 使用 ctypes.memmove 确保数据正确复制到缓冲区指定偏移位置
    ctypes.memmove(ctypes.byref(buf, struct_size), files_data, len(files_data))
    
    return buf.raw


def _copy_files_simple(file_paths: list) -> None:
    """使用最简单可靠的方法复制文件到剪贴板"""
    try:
        # 直接使用 None 作为 owner，避免依赖 tkinter 窗口生命周期
        wc.OpenClipboard(None)
        try:
            # 清空剪贴板
            wc.EmptyClipboard()
            
            # 构建 CF_HDROP 数据
            data = _build_hdrop_data(file_paths)
            
            # 设置 CF_HDROP 数据
            wc.SetClipboardData(wc.CF_HDROP, data)
            
            log(f"Successfully copied {len(file_paths)} files to clipboard using CF_HDROP")
            
        finally:
            wc.CloseClipboard()
            
    except Exception as e1:
        log(f"CF_HDROP method failed: {e1}")
        # 尝试文本方式作为备选
        try:
            _copy_files_as_text(file_paths)
        except Exception as e2:
            # 保留两个异常的完整信息
            raise ClipboardError(
                f"All clipboard methods failed. CF_HDROP: {e1}; Text fallback: {e2}"
            ) from e1


def _copy_files_as_text(file_paths: list) -> None:
    """作为备选：复制文件路径为文本"""
    try:
        wc.OpenClipboard(None)
        try:
            wc.EmptyClipboard()
            
            # 复制文件路径为 Unicode 文本
            text_data = "\r\n".join(file_paths)
            wc.SetClipboardData(wc.CF_UNICODETEXT, text_data)
            
            log("Copied file paths as text to clipboard as fallback")
            
        finally:
            wc.CloseClipboard()
            
    except Exception as e:
        log(f"Text fallback method failed: {e}")
        # 最后的兜底方案：使用 pyperclip
        text_data = "\r\n".join(file_paths)
        pyperclip.copy(text_data)
        log("Used pyperclip as final fallback")


# ============================================================
# 剪贴板文件检测与读取（仅 Windows）
# ============================================================

def is_clipboard_files() -> bool:
    """
    检测剪贴板是否包含文件（CF_HDROP 格式）
    
    Returns:
        True 如果剪贴板中存在文件；否则 False
    """
    try:
        # 某些应用会暂时占用剪贴板，这里做几次轻量重试
        for attempt in range(3):
            try:
                wc.OpenClipboard(None)
                try:
                    result = bool(wc.IsClipboardFormatAvailable(wc.CF_HDROP))
                    log(f"Clipboard files check: {result}")
                    return result
                finally:
                    wc.CloseClipboard()
            except Exception as e:
                log(f"Clipboard files check attempt {attempt + 1} failed: {e}")
                time.sleep(0.03)
        return False
    except Exception as e:
        log(f"Failed to check clipboard files: {e}")
        return False


def get_clipboard_files() -> list[str]:
    """
    获取剪贴板中的文件路径列表（Windows 实现）
    
    使用 wc.GetClipboardData(wc.CF_HDROP) 获取文件路径
    
    Returns:
        文件绝对路径列表
    """
    file_paths = []
    try:
        for attempt in range(3):
            try:
                wc.OpenClipboard(None)
                try:
                    if wc.IsClipboardFormatAvailable(wc.CF_HDROP):
                        data = wc.GetClipboardData(wc.CF_HDROP)
                        # data 是一个包含文件路径的元组
                        if data:
                            file_paths = list(data)
                            log(f"Got {len(file_paths)} files from clipboard")
                        break
                finally:
                    wc.CloseClipboard()
            except Exception as e:
                log(f"Get clipboard files attempt {attempt + 1} failed: {e}")
                time.sleep(0.03)
    except Exception as e:
        log(f"Failed to get clipboard files: {e}")
    
    return file_paths


# ============================================================
# 剪贴板文件检测与读取（仅 Windows）
# ============================================================

def get_markdown_files_from_clipboard() -> list[str]:
    """
    从剪贴板获取 Markdown 文件路径列表
    
    只返回扩展名为 .md 或 .markdown 的文件
    
    Returns:
        Markdown 文件的绝对路径列表（按文件名排序）
    """
    all_files = get_clipboard_files()
    return filter_markdown_files(all_files)


def read_markdown_files_from_clipboard() -> tuple[bool, list[tuple[str, str]], list[tuple[str, str]]]:
    """
    从剪贴板读取 Markdown 文件内容
    
    封装了"获取剪贴板 MD 文件路径 + 逐个读取内容"的完整逻辑。
    读取失败的文件会被跳过，继续处理其它文件。
    
    Returns:
        (found, files_data, errors) 元组：
        - found: 是否发现并成功读取至少一个 MD 文件
        - files_data: [(filename, content), ...] 成功读取的文件名和内容列表
        - errors: [(filename, error_message), ...] 读取失败的文件和错误信息
    """
    md_files = get_markdown_files_from_clipboard()
    return read_markdown_files(md_files)
