"""Word 文档到 Markdown 的转换流程。"""

from __future__ import annotations

from pathlib import Path

from .mathfix import fix_math

__all__ = ["convert_docx"]


def convert_docx(input_path: str | Path, output_path: str | Path) -> Path:
    """将 Word 文档转换为 Markdown，并修复其中的 LaTeX 公式。

    Args:
        input_path: 待转换的 Word 文档路径。
        output_path: 生成的 Markdown 文件路径，父目录不存在时自动创建。

    Returns:
        实际写入的 Markdown 文件路径。

    Raises:
        FileNotFoundError: 输入文件不存在。
    """
    from markitdown import MarkItDown

    source = Path(input_path)
    if not source.is_file():
        raise FileNotFoundError(f"输入文件不存在: {source}")

    target = Path(output_path)
    text = fix_math(MarkItDown().convert(str(source)).text_content)

    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(text, encoding="utf-8")
    return target
