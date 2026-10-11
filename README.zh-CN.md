# md2docx

[English](README.md) | **简体中文**

把 Markdown 转成排版干净的 Word（`.docx`）文档，支持自定义中文字体。

适用场景很日常：你手里有 ChatGPT、Obsidian 或自己笔记里的 Markdown，需要交一份 Word 给导师、客户或期刊。直接复制粘贴进 Word 格式会乱，这个脚本生成的是真正的 Word 标题、列表、表格和超链接。

![转换效果预览](docs/preview.png)

## 为什么不直接用 Pandoc？

[Pandoc](https://pandoc.org) 非常强大，支持公式、引用、脚注和几十种格式。如果你已经在用，继续用就好。md2docx 适合以下情况：

- **单个 Python 文件**，只依赖 `python-docx`，代码好读好改，也能直接嵌进你自己的脚本。
- **不需要 reference doc 默认就好看**：标题有样式，表头有底色，代码块有边框，引用有竖线。
- **中文支持**：显式设置 East Asian 字体并移除冲突的 theme font，源文件中文换行处不会多出空格。

## 功能

- 标题 H1–H6（`#` 和 `===` 两种写法）
- **加粗**、*斜体*、***粗斜体***、~~删除线~~、`行内代码`、转义字符（`\*`）
- 链接、`<https://…>` 和裸 URL 都会变成可点击的 Word 超链接
- 无序、有序、**多级嵌套**列表，使用 Word 原生编号
  - 每个新列表重新从 1 编号，并支持指定起始编号（`3.`）
  - 有序/无序混合嵌套，任务列表 `- [x]`
  - 列表项内部的段落和代码块不会打断编号
- 表格：表头底色、列对齐（`:--`、`:-:`、`--:`）、转义竖线 `\|`、单元格内 `<br>` 换行、跨页重复表头
- 代码块（```` ``` ```` 和 `~~~`），保留缩进
- 引用块，支持 GitHub alert（`> [!NOTE]`、`> [!WARNING]` 等）
- 本地图片（PNG/JPG/GIF/BMP/TIFF），自动缩放到页宽，alt 文字作为图注
- 分隔线（或用 `--hr-pagebreak` 变成分页符）
- 可选目录（`--toc`）
- 可以用你自己的 `.docx` 作为样式模板（`--template`），比如学校或期刊模板
- 自动跳过 YAML front matter 和 HTML 注释
- 支持通配符批量转换（Windows 也可以）

## 安装

需要 Python 3.8 及以上。

如果直接运行单个脚本，下载 `md2docx.py`，然后安装依赖：

```bash
python -m pip install "python-docx>=1.0"
```

如果想从 GitHub 安装成 `md2docx` 命令，需要先安装 Git，再运行：

```bash
python -m pip install git+https://github.com/izeeternal9911/md2docx.git
```

也可以下载并解压项目 ZIP，在包含 `pyproject.toml` 的项目文件夹中打开终端，运行：

```bash
python -m pip install .
```

用上面任一种方式安装命令后：

```bash
md2docx --version
md2docx report.md
```

## 用法

直接使用下载的脚本时，在包含 `md2docx.py` 的文件夹中运行：

```bash
python md2docx.py report.md                    # 输出 report.docx
python md2docx.py report.md final.docx         # 输出 final.docx
python md2docx.py report.md -o final.docx      # 同上
python md2docx.py notes/*.md                   # 每个文件各生成一个 .docx
python md2docx.py thesis.md --toc --cjk-font 宋体
python md2docx.py paper.md --template journal-template.docx
```

安装成命令后，把 `python md2docx.py` 换成 `md2docx` 即可，参数相同。

| 参数 | 默认值 | 说明 |
|---|---|---|
| `-o, --output` | `<输入文件名>.docx` | 输出路径（仅限单个输入文件） |
| `--font` | `Calibri` | 英文字体 |
| `--cjk-font` | `等线` | 中文字体，例如 `宋体`、`微软雅黑`、`仿宋` |
| `--code-font` | `Consolas` | 代码字体 |
| `--font-size` | `11` | 正文字号（pt），标题按比例缩放（不使用 `--template` 时） |
| `--toc` | 关闭 | 在文档开头插入目录 |
| `--toc-title` | `Contents` | 目录标题，可改成 `目录` |
| `--hr-pagebreak` | 关闭 | 把 `---` 当作分页符，而不是分隔线 |
| `--template` | 无 | 从已有 `.docx` 读取样式（字体、标题、页边距） |

**字体：** 查看文档的电脑需要安装所选字体。md2docx 不会安装或嵌入字体；缺少字体时，Word 可能使用其他字体替代。

**目录：** 包含一至三级标题。用 Microsoft Word 打开后，按提示更新域；如果目录没有自动生成，右键目录并选择“更新域”。其他编辑器的支持情况可能不同。

**模板：** 清空模板正文，保留样式、最后一节的页面设置及其页眉页脚。正文和标题字体默认沿用模板，除非传入 `--font` / `--cjk-font`。代码字体由 `--code-font` 指定，默认是 `Consolas`。

### 作为 Python 库使用

```python
from md2docx import md_to_docx, convert, Options

md_to_docx("report.md", "report.docx", toc=True, cjk_font="宋体")

doc = convert("# 标题\n\n你好，世界", Options(font_size=12))  # 返回 python-docx 的 Document
doc.save("hello.docx")
```

## 占位行

整行只有 `[**…**]` 时，会变成居中的灰色斜体占位符，适合在草稿里标记图片位置：

```markdown
[**图 2：概念框架（待插入）**]
```

## 暂不支持

- 数学公式 / LaTeX（`$…$`）会按原文输出。公式多的文档建议用 Pandoc。
- 脚注、定义列表、HTML 标签（`<br>` 除外）
- 缩进式（4 空格）代码块，请用 ``` 围栏代码块
- 列表项**内部**的表格和引用；嵌套引用会被展平
- 网络图片（显示为链接，不会下载）

欢迎提 Issue 和 Pull Request。

## 开发

```bash
pip install python-docx pytest
python -m pytest
```

## 许可证

[MIT](LICENSE)
