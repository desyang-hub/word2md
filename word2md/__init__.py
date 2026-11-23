"""word2md：将 Word 文档转换为 Markdown，并修复其中的 LaTeX 公式。"""

__version__ = "0.1.0"

from .converter import convert_docx
from .mathtype import convert_mathtype
from .mathfix import fix_math

__all__ = ["__version__", "convert_docx", "convert_mathtype", "fix_math"]
