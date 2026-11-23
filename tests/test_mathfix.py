"""fix_math 的单元测试。"""

from word2md import fix_math


def test_inline_formula_unescaped():
    assert fix_math(r"$a\_b$") == "$a_b$"


def test_display_formula_from_brackets():
    assert fix_math("\\[\nE = mc^2\n\\]") == "\n$$\nE = mc^2\n$$\n"


def test_display_formula_single_line():
    assert fix_math("$$x^2$$") == "\n$$x^2$$\n"


def test_triple_dollar_fixed():
    assert fix_math("$$$x$") == "$x$"


def test_mathtype_comment_removed():
    assert "MathType" not in fix_math("% MathType: x %")


def test_excess_blank_lines_collapsed():
    assert fix_math("a\n\n\n\n\nb") == "a\n\n\nb"
