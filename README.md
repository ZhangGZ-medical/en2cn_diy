# en2cn_diy

英文 DOCX → 中文 DOCX 格式保留翻译工具

## 功能

将英文 Word 文档翻译为中文，**完整保留原文档的所有排版格式**（字体样式、表格结构、三线表、段落对齐、行距、背景色、边框等）。

## 核心特性

- 格式零损失：字体、表格、加粗、斜体、三线表
- 无需 OCR 或外部翻译 API
- 支持多表格文档
- 发表级排版，可直接投稿

## 文件结构

```
en2cn_diy/
├── SKILL.md          # 技能说明文件（完整SOP）
└── en2cn_docx.py     # 通用翻译脚本模板
```

## 快速开始

```bash
# 1. 安装依赖
pip install python-docx

# 2. 查看文档结构
python en2cn_docx.py --mode inspect --src your_english.docx

# 3. 执行翻译（填写翻译内容后）
python en2cn_docx.py --mode translate --src your_english.docx --dst chinese_output.docx

# 4. 验证输出
python en2cn_docx.py --mode verify --dst chinese_output.docx
```

## 适用场景

- 学术表格（三线表）
- 正文段落
- 带注释的文档
- 多表格文档
- 发表级格式文档

## 成功案例

- Table 1 Overview of the 4+4 Staging Framework → 表1 缺血性脑卒中4+4分期框架概述
- Table 2 Comparison Between Frameworks → 表2 4+4分期框架与传统分期系统的比较
- Table 3 Summary of Stem Cell Clinical Trials → 表3 基于4+4分期框架的干细胞临床试验总结

---
Author: ZhangGZ-medical
