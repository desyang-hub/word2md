"""通过 Word 自动化批量把 MathType 公式转换为 TeX 文本。

替代原 VBA 宏（见 VBA_script.md），依赖 Windows + Word + MathType 插件。
"""

from __future__ import annotations

from pathlib import Path

__all__ = ["convert_mathtype"]

_WORD_OLE_TYPE = 1  # wdInlineShapeEmbeddedOLEObject
_TEX_TOGGLE_COMMAND = "MTCommand_TeXToggle"  # 等价于 MathType 的 Alt+\


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
    try:
        import win32com.client as win32
    except ImportError as exc:
        raise ImportError(
            "MathType 转换需要 pywin32，请执行: pip install pywin32"
        ) from exc

    from tqdm import tqdm

    path = str(Path(docx_path).resolve())
    word = None
    doc = None
    ok = 0
    fail = 0

    try:
        # DispatchEx 启动独立 Word 实例，Quit 时不会影响用户已打开的 Word
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(path)

        # 先统计 MathType OLE 对象数量，作为进度条总数
        total = sum(
            1
            for i in range(1, doc.InlineShapes.Count + 1)
            if doc.InlineShapes(i).Type == _WORD_OLE_TYPE
        )

        pbar = None if quiet else tqdm(total=total, desc="转换 MathType", unit="公式")
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

        doc.Save()
    finally:
        if doc is not None:
            doc.Close(SaveChanges=0)
        if word is not None:
            word.Quit()

    return ok, fail
