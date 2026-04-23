---
name: en2cn_diy
description: >
  将英文 DOCX 文档翻译为中文，完整保留原始格式与排版（字体、表格、三线表、加粗、
  斜体、对齐等）。不依赖 OCR 或外部翻译 API，由 AI 直接提供译文，再通过 python-docx
  以"克隆源文件XML结构 + 逐 run 替换文字"的方式生成目标文件。
  触发词：翻译、translate、英译中、en2cn、DOCX翻译、英文文档翻译、保留格式翻译、
  翻译并保持格式、translate docx、英文转中文、将英文翻译为中文
version: 1.0.0
base_dir: C:\Users\G1381\.workbuddy\skills\en2cn_diy
---

# en2cn_diy — 英文 DOCX → 中文 DOCX 格式保留翻译工具

将英文 Word 文档翻译为中文，**完整保留原文档的所有排版格式**（字体样式、表格结构、
三线表、段落对齐、行距、背景色、边框等）。

---

## 核心工作流程（SOP）

### Phase 1：读取源文件结构

使用 `python-docx` 读取英文 DOCX，提取以下内容：
- 所有**段落**（含标题、正文、注释）
- 所有**表格**（含表头、数据行，逐单元格、逐 run 提取文字）
- 记录每段/每格的索引位置，用于后续定向替换

```python
from docx import Document
doc = Document(r"<输入文件路径>")

# 查看段落
for i, p in enumerate(doc.paragraphs):
    print(f"[{i}] {p.text[:80]}")

# 查看表格
for ti, table in enumerate(doc.tables):
    for ri, row in enumerate(table.rows):
        for ci, cell in enumerate(row.cells):
            print(f"表{ti} 行{ri} 列{ci}: {cell.text[:50]}")
```

### Phase 2：AI 翻译

**由调用方（AI 助手）提供所有中文译文**，无需外部翻译 API。  
翻译时注意：
- 保留专业术语缩写（如 BBB、MMP、SVZ、TNF-α 等）
- 保留上标符号（如 ⁺、⁻）和特殊字符（如 →、≥、–）
- 将译文按索引整理为字典或列表结构

### Phase 3：克隆源文件 + 逐 run 替换

**关键技术**：以源文件为模板克隆，只替换文字内容，不触碰 XML 格式节点，
从而实现"零格式损失"翻译。

```python
# 打开源文件作为模板（会复制全部格式）
doc = Document(src_path)

# 替换段落（保留 run 格式）
def replace_para_text(para, new_text):
    if para.runs:
        para.runs[0].text = new_text
        for run in para.runs[1:]:
            run.text = ""
    else:
        para.add_run(new_text)

# 替换表格单元格
def replace_cell_text(cell, new_text):
    para = cell.paragraphs[0]
    replace_para_text(para, new_text)

# 使用示例
replace_para_text(doc.paragraphs[0], "表1 缺血性脑卒中4+4分期框架概述")
replace_cell_text(doc.tables[0].cell(0, 0), "分期")

doc.save(dst_path)
```

### Phase 4：验证输出

```python
doc_out = Document(dst_path)
for p in doc_out.paragraphs:
    if p.text.strip():
        print(p.text[:80])
for row in doc_out.tables[0].rows:
    for cell in row.cells:
        print(cell.text[:40], end=" | ")
    print()
```

---

## 脚本模板

技能目录下提供通用模板脚本：`en2cn_docx.py`

```bash
# 用法：修改脚本顶部的 SRC_PATH、DST_PATH 和翻译内容后直接运行
python en2cn_docx.py
```

---

## 适用场景

| 文档类型 | 说明 |
|----------|------|
| 学术表格（三线表） | 完整保留表格样式、背景色、边框 |
| 正文段落 | 保留字体、大小、加粗、斜体 |
| 带注释的文档 | 段落索引精确定位，分别替换 |
| 多表格文档 | 按 `doc.tables[i]` 索引逐表处理 |
| 发表级格式文档 | 无任何格式损失，可直接投稿 |

---

## 注意事项

- **依赖**：`pip install python-docx`（Python 3.8+）
- **编码**：脚本必须 `# -*- coding: utf-8 -*-`，并 `sys.stdout.reconfigure(encoding='utf-8')`
- **多 run 单元格**：单元格内有多个 run 时，只保留第一个 run 的文字和格式，其余 run 清空
- **合并单元格 / _tc 循环引用**：某些 DOCX 表格 `_tc` 存在循环引用（同一行3格共享来自其他行的 `_tc`），此时**不要用 `seen` 集合去重**，直接按 `(ri, ci)` 坐标循环替换即可——每个坐标对应内容都是正确的
- **特殊字符**：上标（⁺⁻²³）直接用 Unicode，不需要额外处理
- **文件路径**：Windows 路径使用 `r""` 原始字符串，避免转义问题

---

## 文件结构

```
en2cn_diy/
├── SKILL.md          # 本说明文件（技能 SOP）
└── en2cn_docx.py     # 通用翻译脚本模板
```

---

## 成功案例

- `Table 1 Overview of the 4+4 Staging Framework for Ischemic Stroke.docx`
  → `表1 缺血性脑卒中4+4分期框架概述.docx`  
  格式零损失，5列×8行三线表，发表级排版，直接可投稿。

- `Table 2 Comparison Between the 4+4 Staging Framework and Traditional Staging Systems.docx`
  → `表2 4+4分期框架与传统分期系统的比较.docx`  
  格式零损失，3列×13行，含 _tc 循环引用表格（不用 seen 去重，直接按坐标替换成功）。

- `Table 3 Summary of Stem Cell Clinical Trials by Treatment Timing According to the 4+4 Staging Framework (51 Trials, 2006–2025).docx`
  → `表3 基于4+4分期框架的干细胞临床试验总结（按治疗时机分类，51项试验，2006–2025）.docx`
  格式零损失，1张表格 36行×7列，3个段落。踩坑：**纵向分期行**结构（分期标签在 col 0，每分期下多行数据 col 0 为空），不能用 `cell.text.strip()` guard 跳过，**必须显式构建 7 列完整数据列表**，分期行显式含标签，非分期行显式含空字符串 `""`。
