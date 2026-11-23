"""Markdown 文本中 LaTeX 公式的规范化修复（与原 word2md.py 逻辑一致）。"""

import re

__all__ = ["fix_math"]


def fix_math(text: str) -> str:
    """修复 Markdown 文本中的 LaTeX 公式。

    Args:
        text: MarkItDown 输出的原始 Markdown 文本。

    Returns:
        公式已规范化的 Markdown 文本。
    """
    # 1. 修复 $$$...$ bug
    text = re.sub(r'\$\$\$(.+?)\$', r'$\1$', text)

    # 2. \[...\] → $$...$$ 带空行 + 去转义
    def fix_display(match):
        content = match.group(1)
        content = content.replace('\\_', '_')
        # content = content.replace('\\{', '{')
        content = content.replace('\\*', '*')
        # content = content.replace('\\}', '}')
        content = content.replace('\\^', '^')
        content = content.strip()
        # return '\n$$\n' + content + '\n$$\n'
        return '\n\\[\n' + content + '\n\\]\n'

    text = re.sub(r'\\\[\s*(.*?)\s*\\\]', fix_display, text, flags=re.DOTALL)

    # 3. 单行 $$内容$$ → 拆开带空行
    text = re.sub(r'\$\$([^\n]+?)\$\$', r'\n$$\1$$\n', text)

    # 4. 已有的 $$ 块确保前后有空行（原脚本中未启用）
    # text = re.sub(r'([^\n])\n\$\$', r'\1\n\n$$', text)
    # text = re.sub(r'\$\$\n([^\n])', r'$$\n\n\1', text)

    # 5. 行内 $...$ 去转义
    def fix_inline(match):
        content = match.group(1)
        content = content.replace('\\_', '_')
        content = content.replace('\\{', '{')
        content = content.replace('\\}', '}')
        content = content.strip()
        return '$' + content + '$'

    text = re.sub(r'\$([^$\n]+?)\$', fix_inline, text)

    # 6. 清理 MathType 残留
    text = re.sub(r'%\s*MathType[^%]*%?', '', text)

    # 7. 多余空行
    text = re.sub(r'\n{4,}', '\n\n\n', text)

    return text
