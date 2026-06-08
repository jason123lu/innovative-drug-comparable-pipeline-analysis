# Innovative Drug Comparable Pipeline Analysis Skill

## 简介

用于创新药可比管线分析的 Codex Skill，支持竞品筛选、可比性评分、临床/临床前数据抽取和来源可追溯的 Excel 横比表生成。

A Codex Skill for comparable drug pipeline analysis, scoring, evidence extraction, and source-backed Excel benchmarking.

## 适用场景

- 创新药可比管线分析
- 竞品管线筛选与分类
- 临床前、1期、2期、3期和上市药物数据整理
- 疗效和安全性横向比较
- 投研报告底稿和 Excel 数据表生成
- 斑块管理、影像终点、降脂、肿瘤、补体等复杂终点项目分析

## 核心功能

- 按靶向、载荷、技术平台、机制、适应症、人群、终点和阶段筛选可比管线
- 对可比管线进行分类、说明和 10 分制量化评分
- 判断管线或所属机构是否仍在推进
- 构建研究级临床/临床前数据池
- 摘录主要终点、次要终点相关基线
- 整理疗效、安全性、上市标签和来源链接
- 生成量化横比表，支持同一管线的不同临床研究分列比较

## 目录结构

```text
.
├── SKILL.md
├── references/
│   ├── workflow.md
│   ├── workbook_schema.md
│   ├── comparable_selection_rules.md
│   ├── clinical_endpoint_rules.md
│   └── source_hierarchy.md
├── scripts/
│   ├── create_workbook_from_spec.py
│   ├── build_quant_matrix_from_raw.py
│   ├── validate_workbook.py
│   └── add_simplified_quant_tab.py
└── assets/
    ├── example_spec.json
    └── comparable_pipeline_template.xlsx
```

## 快速开始

根据 `assets/example_spec.json` 生成一个 Excel 模板：

```bash
python scripts/create_workbook_from_spec.py assets/example_spec.json output.xlsx
```

从原始数据池生成量化横比表：

```bash
python scripts/build_quant_matrix_from_raw.py output.xlsx 管线A_适应症A 管线A_适应症A_量化横比
```

校验工作簿结构和关键字段：

```bash
python scripts/validate_workbook.py output.xlsx
```

## 输出示例

默认生成的 Excel 工作簿包含：

- `管线A的可比管线`
- `管线A_适应症A`
- `管线A_适应症A_量化横比`

其中第一张表用于投资人快速理解“哪些管线可比、为什么可比、可比性高低”；第二张表用于沉淀研究级临床和临床前数据；第三张表用于横向比较不同研究的基线、疗效、安全性和来源。

## 说明

本 Skill 不包含任何真实项目的敏感数据。示例中的 `管线A`、`适应症A` 仅为脱敏占位符。

使用本 Skill 进行真实项目分析时，应优先使用监管文件、临床登记结果、同行评议论文、会议资料和公司官方披露，并对未披露数据明确标注，不应使用推测数据填补。
