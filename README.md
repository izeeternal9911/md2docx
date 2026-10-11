# md2docx

**English** | [简体中文](README.zh-CN.md)

Convert Markdown to a clean, properly formatted Word (`.docx`) document, with explicit Chinese/CJK font settings.

Built for the everyday case: you have Markdown from ChatGPT, Obsidian or your own notes, and you need a Word file you can send to a supervisor, a client or a journal — with real Word lists, tables, headings and hyperlinks rather than pasted text.

![Preview of a converted document](docs/preview.png)

## Why not just use Pandoc?

[Pandoc](https://pandoc.org) is excellent and does far more (math, citations, footnotes, dozens of formats). If you already use it, keep using it. md2docx is for when you want:

- **A single Python file** with one dependency (`python-docx`) — easy to read, modify, or drop into your own script.
- **Good-looking defaults with no reference document** — styled headings, shaded table headers, bordered code blocks, quote bars.
- **Chinese text support** — East Asian font set explicitly, conflicting theme fonts removed, and no stray spaces where Chinese lines were wrapped in the source.

## Features

- Headings H1–H6 (ATX `#` and Setext `===`)
- **Bold**, *italic*, ***bold italic***, ~~strikethrough~~, `inline code`, escapes (`\*`)
- Links, autolinks `<https://…>` and bare URLs → clickable Word hyperlinks
- Bullet, numbered and **nested** lists as real Word numbering
  - numbering restarts for each new list and respects the start number (`3.`)
  - mixed bullet/numbered nesting, task lists `- [x]`
  - continuation paragraphs and fenced code **inside** list items
- Tables with header shading, column alignment (`:--`, `:-:`, `--:`), escaped pipes `\|`, `<br>` in cells, header row repeated across pages
- Fenced code blocks (```` ``` ```` and `~~~`) with preserved indentation
- Block quotes, including GitHub alerts (`> [!NOTE]`, `> [!WARNING]`, …)
- Local images (PNG/JPG/GIF/BMP/TIFF), scaled to page width, alt text as caption
- Horizontal rules (or page breaks with `--hr-pagebreak`)
- Optional table of contents (`--toc`)
- Use your own `.docx` as a style template (`--template`)
- YAML front matter and HTML comments are skipped
- Batch conversion with wildcards (also on Windows)

## Installation

Requires Python 3.8+.

To run the single script, download `md2docx.py` and install its dependency:

```bash
python -m pip install "python-docx>=1.0"
```

To install the `md2docx` command from GitHub, first install Git, then run:

```bash
python -m pip install git+https://github.com/izeeternal9911/md2docx.git
```

Alternatively, download and extract the project ZIP, open a terminal in the project folder containing `pyproject.toml`, and run:

```bash
python -m pip install .
```

After either command installation:

```bash
md2docx --version
md2docx report.md
```

## Usage

When using the downloaded script, run these commands in the folder containing `md2docx.py`:

```bash
python md2docx.py report.md                    # -> report.docx
python md2docx.py report.md final.docx         # -> final.docx
python md2docx.py report.md -o final.docx      # same as above
python md2docx.py notes/*.md                   # one .docx per file
python md2docx.py thesis.md --toc --cjk-font 宋体
python md2docx.py paper.md --template journal-template.docx
```

When installed as a command, replace `python md2docx.py` with `md2docx`; the options are the same.

| Option | Default | Description |
|---|---|---|
| `-o, --output` | `<input>.docx` | Output path (single input only) |
| `--font` | `Calibri` | Font for Latin text |
| `--cjk-font` | `等线` (DengXian) | Font for Chinese/Japanese text, e.g. `宋体`, `微软雅黑`, `仿宋` |
| `--code-font` | `Consolas` | Font for code |
| `--font-size` | `11` | Body size in pt; headings scale with it (without `--template`) |
| `--toc` | off | Insert a table of contents at the top |
| `--toc-title` | `Contents` | Heading for the table of contents |
| `--hr-pagebreak` | off | Treat `---` as a page break instead of a line |
| `--template` | none | Take styles (fonts, headings, margins) from an existing `.docx` |

**Fonts:** the computer used to view the document must have the selected fonts installed. md2docx does not install or embed fonts; Word may substitute missing fonts.

**Table of contents:** includes headings H1–H3. Open the document in Microsoft Word and accept any prompt to update fields. If the TOC does not refresh automatically, right-click it and choose *Update Field*. Support in other editors varies.

**Templates:** the template's body content is cleared. Its styles, final section's page setup, and associated headers and footers are retained. Body and heading fonts are inherited unless you pass `--font` / `--cjk-font`. Code uses `--code-font` (default: `Consolas`).

### As a Python library

```python
from md2docx import md_to_docx, convert, Options

md_to_docx("report.md", "report.docx", toc=True, cjk_font="宋体")

doc = convert("# Title\n\n你好，世界", Options(font_size=12))  # python-docx Document
doc.save("hello.docx")
```

## Placeholder lines

A line consisting only of `[**…**]` becomes a centred, grey italic placeholder — handy for marking where figures will go in a draft:

```markdown
[**Figure 2: Conceptual framework (to be inserted)**]
```

## Not supported (yet)

- Math / LaTeX (`$…$`) — left as plain text. Use Pandoc for math-heavy documents.
- Footnotes, definition lists, raw HTML (other than `<br>`)
- Indented (4-space) code blocks — use fenced blocks
- Tables or quotes nested *inside* list items; nested quotes are flattened
- Remote images (shown as a link instead of being downloaded)

Issues and pull requests are welcome.

## Development

```bash
pip install python-docx pytest
python -m pytest
```

## License

[MIT](LICENSE)
