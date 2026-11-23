"""CLI 流程测试（MathType 转换与 docx 转换均使用替身函数）。"""

from pathlib import Path

import word2md.cli as cli


def test_cli_converts_tmp_copy(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "论文.docx").write_bytes(b"fake docx")

    calls = {}

    def fake_mathtype(path, *, quiet=False):
        calls["mathtype"] = Path(path)
        return 3, 0

    def fake_convert(path, output):
        calls["convert"] = (Path(path), Path(output))
        Path(output).write_text("markdown", encoding="utf-8")
        return Path(output)

    monkeypatch.setattr(cli, "convert_mathtype", fake_mathtype)
    monkeypatch.setattr(cli, "convert_docx", fake_convert)

    assert cli.main(["论文.docx"]) == 0

    copied = tmp_path / ".tmp" / "论文.docx"
    assert copied.read_bytes() == b"fake docx"
    assert calls["mathtype"] == Path(".tmp") / "论文.docx"
    assert calls["convert"] == (Path(".tmp") / "论文.docx", Path("论文_final.md"))
    assert (tmp_path / "论文_final.md").read_text(encoding="utf-8") == "markdown"


def test_cli_skip_mathtype(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "a.docx").write_bytes(b"x")

    calls = {"mathtype": False}

    def fake_mathtype(path, *, quiet=False):
        calls["mathtype"] = True
        return 0, 0

    monkeypatch.setattr(cli, "convert_mathtype", fake_mathtype)
    monkeypatch.setattr(
        cli, "convert_docx", lambda path, output: Path(output)
    )

    assert cli.main(["--skip-mathtype", "a.docx"]) == 0
    assert calls["mathtype"] is False


def test_cli_missing_input(tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)

    assert cli.main(["missing.docx"]) == 1
    assert "输入文件不存在" in capsys.readouterr().err
