"""CLI 流程测试（MathType 转换与 docx 转换均使用替身函数）。"""

from pathlib import Path

import word2md.cli as cli


def test_cli_converts_tmp_copy(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "论文.docx").write_bytes(b"fake docx")

    calls = {}

    def fake_mathtype(path, *, quiet=False):
        calls["mathtype"] = Path(path)
        calls["existed"] = Path(path).exists()
        return 3, 0

    def fake_convert(path, output):
        calls["convert"] = (Path(path), Path(output))
        Path(output).write_text("markdown", encoding="utf-8")
        return Path(output)

    monkeypatch.setattr(cli, "convert_mathtype", fake_mathtype)
    monkeypatch.setattr(cli, "convert_docx", fake_convert)

    assert cli.main(["论文.docx"]) == 0

    assert calls["mathtype"] == Path("tmp_论文.docx")
    assert calls["existed"] is True
    assert calls["convert"] == (Path("tmp_论文.docx"), Path("论文_final.md"))
    assert (tmp_path / "论文_final.md").read_text(encoding="utf-8") == "markdown"
    # 转换结束后删除临时副本
    assert not (tmp_path / "tmp_论文.docx").exists()
    # 原文件保持不变
    assert (tmp_path / "论文.docx").read_bytes() == b"fake docx"


def test_cli_skip_mathtype(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "a.docx").write_bytes(b"x")

    calls = {"mathtype": False}

    def fake_mathtype(path, *, quiet=False):
        calls["mathtype"] = True
        return 0, 0

    monkeypatch.setattr(cli, "convert_mathtype", fake_mathtype)
    monkeypatch.setattr(cli, "convert_docx", lambda path, output: Path(output))

    assert cli.main(["--skip-mathtype", "a.docx"]) == 0
    assert calls["mathtype"] is False
    assert not (tmp_path / "tmp_a.docx").exists()


def test_cli_removes_tmp_copy_on_error(tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "a.docx").write_bytes(b"x")

    def fake_convert(path, output):
        raise OSError("disk full")

    monkeypatch.setattr(cli, "convert_docx", fake_convert)

    assert cli.main(["--skip-mathtype", "a.docx"]) == 1
    assert "disk full" in capsys.readouterr().err
    assert not (tmp_path / "tmp_a.docx").exists()


def test_cli_latex2mathtype_mode(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "a.docx").write_bytes(b"docx")

    calls = {}

    def fake(path, *, quiet=False):
        calls["path"] = Path(path)
        Path(path).write_bytes(b"converted")
        return 5, 4, 1

    monkeypatch.setattr(cli, "convert_latex_to_mathtype", fake)

    assert cli.main(["--latex2mathtype", "a.docx"]) == 0

    assert calls["path"] == Path("tmp_a.docx")
    assert (tmp_path / "a_mathtype.docx").read_bytes() == b"converted"
    assert not (tmp_path / "tmp_a.docx").exists()


def test_cli_latex2mathtype_custom_output(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    (tmp_path / "a.docx").write_bytes(b"docx")

    monkeypatch.setattr(
        cli, "convert_latex_to_mathtype", lambda path, *, quiet=False: (1, 1, 0)
    )

    assert cli.main(["--latex2mathtype", "a.docx", "-o", "out/b.docx"]) == 0
    assert (tmp_path / "out" / "b.docx").exists()
    assert not (tmp_path / "tmp_a.docx").exists()


def test_cli_missing_input(tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)

    assert cli.main(["missing.docx"]) == 1
    assert "输入文件不存在" in capsys.readouterr().err
