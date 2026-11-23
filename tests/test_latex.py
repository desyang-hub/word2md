"""LaTeX 公式扫描的单元测试。"""

from word2md.mathtype import find_latex_matches


def test_find_display_formula():
    text = "前 \\[x^2 + y^2\\] 后"
    matches = find_latex_matches(text)
    assert len(matches) == 1
    start, end, latex = matches[0]
    assert latex == "\\[x^2 + y^2\\]"
    assert text[start:end] == latex


def test_find_inline_formula():
    text = "设 $a\\_b$ 为变量"
    matches = find_latex_matches(text)
    assert [m[2] for m in matches] == ["$a\\_b$"]


def test_inline_formula_not_across_lines():
    text = "第一行 $a\nb$ 第二行"
    assert find_latex_matches(text) == []


def test_find_multiple_formulas_without_overlap():
    text = "$a$ 与 \\[b\\] 以及 $c$"
    assert [m[2] for m in find_latex_matches(text)] == ["$a$", "\\[b\\]", "$c$"]


def test_plain_text_has_no_match():
    assert find_latex_matches("没有任何公式") == []
