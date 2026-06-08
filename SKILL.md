---
name: innovative-drug-comparable-analysis
description: Use this skill when the user wants an investor-facing comparable pipeline analysis for an innovative drug, including comparable pipeline discovery, comparability classification/scoring, clinical/preclinical data extraction, source-backed Excel workbooks, and quantitative cross-trial comparison tabs. Use for tasks involving drug pipelines, targets, mechanisms, technology platforms, indications, competitive landscapes, efficacy/safety benchmarking, clinical trial data, plaque/atherosclerosis imaging endpoints, or investment research tables.
---

# 创新药可比管线分析

## 核心原则

1. 先回答“谁最可比、为什么可比”，再做疗效和安全性横比。
2. 可比性按靶向/载荷/技术/机制/适应症/人群/终点/阶段综合判断，不只看靶点。
3. 每条可比管线都要说明可比性高低，给出分类、量化评分和命中维度。
4. 原始数据池按“一个临床研究或一条临床前数据一行”组织；同一药物的不同研究不要挤在同一行。
5. 基线必须摘录到 Excel 单元格中，尤其是主要终点和关键次要终点相关基线，不要写“见论文”。
6. 疗效数据尽量同时给绝对变化和相对变化；无法计算时明确写“未披露/无法计算”，不要用同类药物推断。
7. 临床登记不等于结果披露；最高阶段也不等于仍在推进，要单独判断公司/机构是否还在推进。
8. 每个关键数字必须能追溯来源，优先使用监管文件、临床注册结果、同行评议论文和会议资料。

## 默认工作流

完整流程见 `references/workflow.md`。执行时按以下顺序：

1. 定义目标管线：公司、管线、靶向/载荷、机制、技术形式、适应症、人群、阶段、核心投资问题。
2. 按 `references/comparable_selection_rules.md` 搜索并筛选可比管线，通常不超过 20 条。
3. 先生成 `{管线}的可比管线`：列出基础属性、可比性分类、评分、命中维度、推进状态和核心差异。
4. 按 `references/source_hierarchy.md` 搜索已上市药物和未上市管线的注册临床、论文、会议和公司披露。
5. 按 `references/clinical_endpoint_rules.md` 抽取适应症相关终点、基线、疗效和安全性。
6. 按 `references/workbook_schema.md` 更新 `{管线}_{适应症}` 原始数据池。
7. 用 `scripts/build_quant_matrix_from_raw.py` 从原始数据池生成 `{管线}_{适应症}_量化横比`。
8. 用 `scripts/validate_workbook.py` 校验 workbook 结构、关键字段、空基线、空疗效、来源链接和公式错误。

## 默认 Excel 交付结构

每个项目默认产出一个 `.xlsx`，推荐 sheet 顺序：

- `{管线}的可比管线`
- `{管线}_{适应症}`
- `{管线}_{适应症}_量化横比`

以下 sheet 仅在用户要求或项目需要时创建：

- `指标口径`
- `{管线}_{适应症}_披露总览`
- `{管线}_{适应症}_量化横比_精简版`

## 脚本

脚本用于机械落表和校验，不替代研究判断：

- `scripts/create_workbook_from_spec.py`：按 JSON spec 生成新版 Excel 骨架。
- `scripts/build_quant_matrix_from_raw.py`：从研究级原始数据池生成量化横比表。
- `scripts/add_simplified_quant_tab.py`：从完整版量化横比抽取可选精简表。
- `scripts/validate_workbook.py`：检查 workbook 结构、公式错误、关键字段和数据完整性。

## 输出要求

最终回复用户时说明：

- 最终文件路径。
- 覆盖了哪些目标管线和可比管线。
- 哪些数据来自监管文件、临床登记结果、论文、会议资料、公司新闻稿或官网。
- 哪些关键数据未披露或只能定性描述。
- 是否生成了量化横比和可选精简版。
