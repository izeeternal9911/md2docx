---
title: md2docx sample
author: example
---

# md2docx 示例文档 / Sample Document

这是一段中文正文，用来测试**中文字体**是否正确显示。源文件里这一行在中间换了行，
转换后不应该出现多余的空格。English text is mixed in with *italic*, ***bold italic***,
~~strikethrough~~ and `inline code`.

Links work too: [Python documentation](https://docs.python.org/3/), an autolink
<https://github.com>, and a bare URL https://www.example.com/path.

This sentence has a pipe | in it, but it is not a table.

## 1. 列表 Lists

1. 第一步：安装依赖
2. 第二步：运行脚本
   - 子项目 A，包含 **加粗** 内容
   - 子项目 B
     1. 更深一层的编号
     2. 第二个
3. 第三步：检查结果

   This paragraph belongs to step 3 and is indented with it.

A paragraph between two lists.

1. This numbered list starts again at 1
2. instead of continuing from the list above

Another paragraph.

3. A list can also start at 3
4. like this

Task lists:

- [x] Finished task
- [ ] Unfinished task

## 2. 表格 Table

| 指标 Metric | 说明 | 数值 |
|:------------|:----:|-----:|
| Heart rate  | 心率 `HR` | 92 |
| Core temp   | 核心体温 | 37.8 |
| Escaped     | a \| b | — |

## 3. 代码 Code

Run the converter:

```bash
pip install python-docx
python md2docx.py input.md --toc
```

```python
def hello(name):
    # 中文注释
    if name:
        print(f"Hello, {name}")
```

Steps with code inside a list:

1. Create a virtual environment:
   ```bash
   python -m venv .venv
   ```
2. Activate it and install.

## 4. 引用 Quotes

> 这是一段引用。Quotes can span
> multiple lines and contain **formatting**.
>
> A second paragraph in the same quote.

> [!NOTE]
> GitHub-style alerts are supported.

## 5. Images & placeholders

![Figure 1: A sample chart](chart.png)

[**Figure 2: Conceptual framework (to be inserted)**]

---

### 6. Line breaks

First line with a hard break  
second line on its own.

Setext Heading
==============

The end. 完。
