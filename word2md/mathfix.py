"""Markdown 文本中 LaTeX 公式的规范化修复。

Word 经 MathType 转 TeX、再由 MarkItDown 转 Markdown 后，公式常出现
``$$$...$``、``\\[...\\]``、单行 ``$$...$$`` 以及多余转义等形态，
本模块统一将其整理为通用 Markdown 渲染器可识别的公式写法。
"""

from __future__ import annotations

import re

__all__ = ["fix_math"]

# $$$...$ → $...$（MathType 转换偶发的多余 $）
_TRIPLE_DOLLAR = re.compile(r"\$\$\$(.+?)\$")
# \[...\] → $$...$$
_DISPLAY_BRACKET = re.compile(r"\\\[\s*(.*?)\s*\\\]", re.DOTALL)
# 单行 $$...$$ → 独立的公式块
_SINGLE_LINE_DISPLAY = re.compile(r"\$\$([^\n]+?)\$\$")
# 行内 $...$
_INLINE = re.compile(r"\$([^$\n]+?)\$")
# MathType 插入的注释残留
_MATHTYPE_COMMENT = re.compile(r"%\s*MathType[^%]*%?")
# 连续 4 个及以上换行压缩为 3 个
_EXCESS_BLANK_LINES = re.compile(r"\n{4,}")


def _unescape(content: str, *, caret: bool = False) -> str:
    """去掉 math 内容中多余的转义并去除首尾空白。"""
    pairs = (("\\_", "_"), ("\\{", "{"), ("\\}", "}"))
    if caret:
        pairs += (("\\^", "^"),)
    for escaped, raw in pairs:
        content = content.replace(escaped, raw)
    return content.strip()


def fix_math(text: str) -> str:
    """修复 Markdown 文本中的 LaTeX 公式。

    Args:
        text: MarkItDown 输出的原始 Markdown 文本。

    Returns:
        公式已规范化的 Markdown 文本。
    """

    def fix_display_bracket(match: "re.Match[str]") -> str:
        return "\n$$\n" + _unescape(match.group(1), caret=True) + "\n$$\n"

    def fix_inline(match: "re.Match[str]") -> str:
        return "$" + _unescape(match.group(1)) + "$"

    text = _TRIPLE_DOLLAR.sub(r"$\1$", text)
    text = _DISPLAY_BRACKET.sub(fix_display_bracket, text)
    text = _SINGLE_LINE_DISPLAY.sub(r"\n$$\1$$\n", text)
    text = _INLINE.sub(fix_inline, text)
    text = _MATHTYPE_COMMENT.sub("", text)
    text = _EXCESS_BLANK_LINES.sub("\n\n\n", text)
    return text
