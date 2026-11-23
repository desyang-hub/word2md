"""通过 Word 自动化在 LaTeX 文本与 MathType 公式之间批量转换。

- ``convert_latex_to_mathtype``：把文档中的 LaTeX 文本公式转成 MathType 公式；
- ``convert_mathtype``：把 MathType 公式批量转回 TeX 文本。

两者都依赖 Windows + Word + MathType 插件（``MTCommand_TeXToggle``）。
"""

from __future__ import annotations

from contextlib import contextmanager
from collections.abc import Iterator
from pathlib import Path

__all__ = ["convert_latex_to_mathtype", "convert_mathtype", "find_latex_matches"]

_WORD_OLE_TYPE = 1  # wdInlineShapeEmbeddedOLEObject
_TEX_TOGGLE_COMMAND = "MTCommand_TeXToggle"  # 等价于 MathType 的 Alt+\

_DISPLAY_START = r"\["
_DISPLAY_END = r"\]"
_INLINE_MARK = "$"


def _load_win32():
    try:
        import win32com.client as win32
    except ImportError as exc:
        raise ImportError(
            "MathType 转换需要 pywin32，请执行: pip install pywin32"
        ) from exc
    return win32


@contextmanager
def _open_document(docx_path: str | Path) -> Iterator[tuple]:
    """用独立 Word 实例打开文档，结束后保存、关闭并退出 Word。"""
    win32 = _load_win32()
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.ScreenUpdating = False
    word.DisplayAlerts = 0
    doc = None
    try:
        doc = word.Documents.Open(str(Path(docx_path).resolve()))
        yield word, doc
        doc.Save()
    finally:
        if doc is not None:
            doc.Close(SaveChanges=0)
        word.Quit()


def find_latex_matches(text: str) -> list[tuple[int, int, str]]:
    """扫描文本中的 LaTeX 公式。

    Args:
        text: Word 文档正文（``doc.Content.Text``）。

    Returns:
        ``(起始偏移, 结束偏移, 公式原文)`` 列表，偏移相对于 ``text``。
    """
    matches: list[tuple[int, int, str]] = []
    i = 0
    length = len(text)

    while i < length:
        # 1. 块级公式 \[...\]
        if text.startswith(_DISPLAY_START, i):
            end = text.find(_DISPLAY_END, i + len(_DISPLAY_START))
            if end != -1:
                end += len(_DISPLAY_END)
                matches.append((i, end, text[i:end]))
                i = end
                continue

        # 2. 行内公式 $...$，不跨行
        if text[i] == _INLINE_MARK:
            end = text.find(_INLINE_MARK, i + 1)
            if end != -1:
                between = text[i + 1 : end]
                if "\r" not in between and "\n" not in between:
                    matches.append((i, end + 1, text[i : end + 1]))
                    i = end + 1
                    continue

        i += 1

    return matches


def convert_latex_to_mathtype(
    docx_path: str | Path, *, quiet: bool = False
) -> tuple[int, int, int]:
    """把文档中的 LaTeX 文本公式批量转换为 MathType 公式，并就地保存。

    Args:
        docx_path: 待处理的 Word 文档路径。
        quiet: 为 True 时不显示 tqdm 进度条。

    Returns:
        (匹配数量, 成功数量, 失败数量)。

    Raises:
        ImportError: 未安装 pywin32。
    """
    from tqdm import tqdm

    matched = 0
    ok = 0
    fail = 0

    with _open_document(docx_path) as (word, doc):
        content = doc.Content
        base = content.Start  # 通常是 0
        matches = find_latex_matches(content.Text)
        matched = len(matches)

        pbar = (
            None
            if quiet
            else tqdm(total=matched, desc="LaTeX → MathType", unit="公式")
        )
        try:
            # 倒序处理，避免替换后位置偏移
            for start, end, latex in reversed(matches):
                doc.Range(Start=base + start, End=base + end).Select()
                try:
                    word.Run(_TEX_TOGGLE_COMMAND)
                    ok += 1
                except Exception as exc:  # 单个公式失败时继续处理其余公式
                    fail += 1
                    if pbar is not None:
                        pbar.write(f"转换失败: {latex[:60]}... 错误: {exc}")
                finally:
                    if pbar is not None:
                        pbar.update(1)
                        pbar.set_postfix(成功=ok, 失败=fail)
        finally:
            if pbar is not None:
                pbar.close()

    return matched, ok, fail


def convert_mathtype(docx_path: str | Path, *, quiet: bool = False) -> tuple[int, int]:
    """将文档中的 MathType OLE 公式批量转换为 TeX 文本，并就地保存。

    Args:
        docx_path: 待处理的 Word 文档路径。
        quiet: 为 True 时不显示 tqdm 进度条。

    Returns:
        (成功数量, 失败数量)。

    Raises:
        ImportError: 未安装 pywin32。
    """
    from tqdm import tqdm

    ok = 0
    fail = 0

    with _open_document(docx_path) as (word, doc):
        # 先统计 MathType OLE 对象数量，作为进度条总数
        total = sum(
            1
            for i in range(1, doc.InlineShapes.Count + 1)
            if doc.InlineShapes(i).Type == _WORD_OLE_TYPE
        )

        pbar = (
            None if quiet else tqdm(total=total, desc="转换 MathType", unit="公式")
        )
        try:
            # 倒序遍历，避免转换过程中集合变化导致索引错乱
            for i in range(doc.InlineShapes.Count, 0, -1):
                shape = doc.InlineShapes(i)
                if shape.Type != _WORD_OLE_TYPE:
                    continue

                try:
                    shape.Select()
                    word.Run(_TEX_TOGGLE_COMMAND)
                    ok += 1
                except Exception as exc:  # 单个公式失败时继续处理其余公式
                    fail += 1
                    if pbar is not None:
                        pbar.write(f"第 {i} 个对象转换失败: {exc}")
                finally:
                    if pbar is not None:
                        pbar.update(1)
                        pbar.set_postfix(成功=ok, 失败=fail)
        finally:
            if pbar is not None:
                pbar.close()

    return ok, fail
