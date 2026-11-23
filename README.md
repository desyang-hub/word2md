# word2md

将 Word 文档（`.docx`）转换为 Markdown，并修复 MathType 公式转 TeX 后产生的 LaTeX 公式问题。

## 功能

- 基于 [MarkItDown](https://github.com/microsoft/markitdown) 完成 `.docx` → Markdown 转换；
- 公式规范化：
  - `$$$...$` → `$...$`（修复多余 `$`）；
  - `\[ ... \]` 块级公式 → 独立的 `$$ ... $$` 公式块；
  - 单行 `$$...$$` → 独立的 `$$ ... $$` 公式块；
  - 行内 `$...$` 去除 `\_`、`\{`、`\}` 等多余转义；
  - 清理 MathType 注释残留，压缩多余空行；
- 安装后提供 `word2md` 命令，也支持 `python -m word2md` 与 Python API 调用。

## 环境要求

- Python >= 3.10（`markitdown` 0.1.0 起要求 Python 3.10+）；
- 注意：PyPI 上的 `markitdown 0.0.1a1` 是不含实现的占位包，请确认安装的是 `>=0.1.0` 版本。

## 安装

```powershell
git clone <仓库地址> word2md
cd word2md
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install .
```

开发模式（包含 pytest）：

```powershell
pip install -e ".[dev]"
```

也可以先用 `pip install -r requirements.txt` 安装运行依赖，再执行 `pip install .` 安装命令行工具。

安装完成后验证：

```powershell
word2md --version
```

## 使用

### 1. 预处理 Word 文档（推荐）

MathType 公式默认以 OLE 对象形式存在，先按 [VBA_script.md](VBA_script.md) 中的宏把公式批量转成 TeX 文本，再交给本工具转换，公式才能被保留为 LaTeX。

### 2. 转换文档

```powershell
word2md 论文.docx                  # 输出 论文_final.md（与原脚本默认一致）
word2md 论文.docx -o out/论文.md   # 指定输出路径（目录不存在会自动创建）
word2md 论文.docx -q               # 静默模式
python -m word2md 论文.docx        # 未加入 PATH 时也可用模块方式运行
python word2md.py 论文.docx        # 兼容原脚本的调用方式
```

| 参数 | 说明 |
| --- | --- |
| `input` | 待转换的 Word 文档路径 |
| `-o, --output FILE` | 输出 Markdown 路径，默认为 `<输入文件名>_final.md`（当前工作目录） |
| `-q, --quiet` | 不输出进度信息 |
| `-V, --version` | 显示版本号 |

### 3. Python API

```python
from word2md import convert_docx, fix_math

convert_docx("论文.docx", "论文.md")

text = fix_math(r"$a\_b$")  # "$a_b$"
```

## 项目结构

```text
word2md/
├── word2md/
│   ├── __init__.py     # 包入口，导出 convert_docx / fix_math / __version__
│   ├── __main__.py     # 支持 python -m word2md
│   ├── cli.py          # 命令行参数解析与入口
│   ├── converter.py    # MarkItDown 转换流程
│   └── mathfix.py      # LaTeX 公式修复（与原脚本逻辑一致）
├── word2md.py          # 兼容入口，等价于 word2md 命令
├── tests/
│   └── test_mathfix.py
├── pyproject.toml
├── requirements.txt
├── VBA_script.md       # MathType 公式转 TeX 的 VBA 宏说明
└── README.md
```

## 开发与测试

```powershell
pip install -e ".[dev]"
pytest
```
