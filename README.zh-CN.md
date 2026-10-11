# md2docx

[English](README.md) | **简体中文**

把 Markdown（`.md`）文件转成 Word（`.docx`），支持中文、标题、表格、列表、链接和本地图片。

准备好后，每次转换只要：**输入 Python 命令 → 拖入脚本 → 拖入要转换的文件 → 回车。**

**Mac 用 `python3`，Windows 用 `python`。** 下面按系统说明。

## 怎么用（Mac）

### 1. 下载工具

点这个页面上方的绿色 **Code → Download ZIP**，下载后双击解压，找到里面的 **`md2docx.py`**。

### 2. 首次准备（只做一次）

电脑需要安装 Python。如果还没有，打开 [Python 官网](https://www.python.org/downloads/macos/)，下载 macOS 安装包，双击后按提示安装。已经装过可以跳过。

按 **⌘ + 空格**，搜索“终端”并打开。复制下面这一行，粘贴进去，按回车，等它安装完成：

```bash
python3 -m pip install "python-docx>=1.0"
```

这是转换需要的组件，装好后不用每次重复安装。

### 3. 转换你的文件

打开终端，然后按顺序操作：

1. 输入 **`python3`**，按一下**空格**。
2. 把下载的 **`md2docx.py`** 从 Finder 拖进终端，文件位置会自动填入。
3. 确认后面有一个**空格**，再把你要转换的 **`.md` 文件**拖进去。
4. 按**回车**。

终端出现 `Converted:` 就表示转换完成。回到 **`.md` 文件所在的文件夹**，就能找到同名的 **`.docx` 文件**，用 Word 打开即可。

例如 `我的笔记.md` 会生成 `我的笔记.docx`。已有同名 Word 文件会被覆盖，想保留旧版本时先改名。

<details>
<summary>Windows 用户点这里</summary>

1. 同样先下载并解压工具；从 [Python 官网](https://www.python.org/downloads/windows/) 安装 Python。如果安装界面有 **Add Python to PATH**，勾选它。
2. 打开“命令提示符”，复制下面这一行并按回车，安装转换组件（只做一次）：

```bat
python -m pip install "python-docx>=1.0"
```

3. 把 **`md2docx.py` 复制到你要转换的 `.md` 文件所在的文件夹**，保留 Markdown 和图片原来的位置。
4. 在这个文件夹的地址栏输入 `cmd`，按回车，打开命令提示符。
5. 输入下面这一行，把 `我的笔记.md` 换成你的实际文件名，保留两边的英文双引号，按回车：

```bat
python md2docx.py "我的笔记.md"
```

转换后的 Word 文件会出现在同一个文件夹。

</details>

## 转换效果

![转换效果预览](docs/preview.png)

[查看示例 Markdown](examples/sample.md) · [下载示例 Word](examples/sample.docx)

<details>
<summary>遇到问题？点这里</summary>

**提示找不到 `python3` / `python`？**

先安装 Python，再关掉终端或命令提示符，重新打开后试一次。如果仍不行，提 [Issue](https://github.com/izeeternal9911/md2docx/issues)，附上报错文字和你使用的系统。

**提示 `No module named 'docx'`？**

重新运行“首次准备”里的组件安装命令，等它安装成功后再转换。

**提示 `file not found` 或 `can't open file`？**

Mac 上按第 3 步重新拖入两个文件；Windows 上确认已把脚本复制到 Markdown 所在文件夹，并填写实际文件名。文件位置或名称有空格时，手动输入的位置要用英文双引号包住。

**生成的 Word 在哪里？**

在原 `.md` 文件旁边。比如桌面上的 `笔记.md`，会生成桌面上的 `笔记.docx`。

**Word 字体和预览图不一样？**

查看文档的电脑需要安装对应字体；缺少字体时 Word 会替换。想换字体，可以展开下面的“更多用法”。

</details>

<details>
<summary>支持哪些格式？点这里</summary>

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

</details>

<details>
<summary>更多用法：换字体、目录、模板和批量转换</summary>

下面的命令示例假设你已在终端中打开 `md2docx.py` 和输入文件所在的文件夹。Mac 使用 `python3`，Windows 把它换成 `python`。

```bash
python3 md2docx.py "我的笔记.md"
python3 md2docx.py "我的笔记.md" -o "另一个文件名.docx"
python3 md2docx.py "我的笔记.md" --cjk-font 宋体
python3 md2docx.py "我的笔记.md" --toc --toc-title 目录
python3 md2docx.py "我的笔记.md" --template "学校模板.docx"
python3 md2docx.py "*.md"
```

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

### 占位行

整行只有 `[**…**]` 时，会变成居中的灰色斜体占位符，适合在草稿里标记图片位置：

```markdown
[**图 2：概念框架（待插入）**]
```

</details>

<details>
<summary>使用前了解：暂不支持的格式</summary>

- 数学公式 / LaTeX（`$…$`）会按原文输出。公式多的文档建议用 Pandoc。
- 脚注、定义列表、HTML 标签（`<br>` 除外）
- 缩进式（4 空格）代码块，请用 ``` 围栏代码块
- 列表项**内部**的表格和引用；嵌套引用会被展平
- 网络图片（显示为链接，不会下载）

</details>

<details>
<summary>给开发者：项目安装、Python 调用和测试</summary>

需要 Python 3.8 及以上。以下是可选的项目安装和开发方式。Mac 使用 `python3`；Windows 把下面命令中的 `python3` 换成 `python`。

### 安装成 md2docx 命令

已安装 Git 时，可以从仓库安装：

```bash
python3 -m pip install git+https://github.com/izeeternal9911/md2docx.git
```

也可以在包含 `pyproject.toml` 的解压文件夹中运行：

```bash
python3 -m pip install .
```

然后用 `md2docx "我的笔记.md"` 转换，参数与脚本相同。Windows 把安装命令中的 `python3` 换成 `python`。

### 作为 Python 库使用

```python
from md2docx import md_to_docx, convert, Options

md_to_docx("report.md", "report.docx", toc=True, cjk_font="宋体")

doc = convert("# 标题\n\n你好，世界", Options(font_size=12))  # 返回 python-docx 的 Document
doc.save("hello.docx")
```

### 运行测试

在项目文件夹中运行：

```bash
python3 -m pip install "python-docx>=1.0" pytest
python3 -m pytest
```

</details>

欢迎在 [Issues](https://github.com/izeeternal9911/md2docx/issues) 反馈问题。

许可证：[MIT](LICENSE)。
