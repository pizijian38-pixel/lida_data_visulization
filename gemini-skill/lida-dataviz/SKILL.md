---
name: lida-dataviz
description: 把用户上传的原始数据（CSV / Excel / JSON 表格）和一句需求变成合适的图表：静态 PNG，以及带原生可编辑 Excel 图表的 .xlsx（可直接复制进 PPT）。当用户上传数据并要求画图、做图表、可视化、看趋势/对比/排名/占比/分布/相关性，或说"帮我看看这份数据"时使用。只问一个数字或只要表格时不要用。
---

# LIDA 数据可视化（Gemini 版）

流程来自 Microsoft LIDA：**摘要 → 目标 → 代码 → 执行 → 看图评估 → 修复 → 交付**。
确定性的部分由 `scripts/lida_viz.py` 完成，你负责定目标、写代码、看图评估。

## 0. 准备（每个对话第一次用时，在代码执行环境里做）

1. 找到 `lida_viz.py`：用 `os.walk` 在当前目录及其子目录里搜索这个文件名。
   **找不到时**，把本 skill 的 `scripts/lida_viz.py` 的完整内容写入当前目录的 `lida_viz.py`，再继续。
2. `import lida_viz`（如有需要先把它所在的目录加到 `sys.path`），然后运行 `lida_viz.check_env()`，并按返回的 `notes` 行事：
   - `can_export_xlsx` 为 false：只交付 PNG，告诉用户原因。
   - `cjk_font` 为 null：PNG 里的文字改用英文；XLSX 照常用中文。
   - **环境没有网络，永远不要 `pip install`。**
3. 找到用户上传的数据文件（`os.listdir` / `os.walk`）。没有文件时，请用户上传；如果用户把表格粘贴在聊天里，先把它保存成 CSV。

## 1. 数据摘要

`summary = lida_viz.summarize_file("<数据文件>")`。Excel 有多个工作表时传 `sheet=`。用户没说用哪个工作表就先问。

重点看：
- `dtype`：每列的类型
- `stored_as_text`：文本形式的数字或日期，代码里必须先用 `to_number()` / `to_date()` 转换
- `top_values`：类别分布，判断是否需要 Top N
- `notes`：脚本发现的数据陷阱，逐条处理

## 2. 定目标

把需求写成 `{question, visualization, rationale}`，visualization 必须用摘要里的**确切列名**。
- 需求明确：直接写 1 个目标。
- 需求模糊：按 `references/goal_generation.md` 列 3–5 个候选让用户选。
- 只有会改变结论的歧义（例如两个候选指标列、口径不清）才问用户，而且只问一次。

## 3. 写代码

选图看 `references/chart_selection.md`，代码严格按 `references/code_template.md`。代码里定义三样东西：
- `EXCEL_CHART`：Excel 图表的设置（类型、标题、轴标题、数字格式）
- `chart_data(data)`：汇总后的宽表，PNG 和 Excel 共用
- `plot(data)`：基于 `chart_data()` 作图

代码保存为 `<slug>.py`。

## 4. 渲染

```python
r = lida_viz.render_chart("<数据文件>", "<slug>.py", "<slug>.png", out_xlsx="<slug>.xlsx")
```

- `status == "error"`：按 `stage`、`chart_code_lines`、`hint`、`available_columns` 修改代码后重跑，最多 3 次。
  `stage == "xlsx"` 时 PNG 已经生成，问题出在 `chart_data` 或 `EXCEL_CHART`。
- `warnings`、`xlsx_warnings` 非空：必须处理。

## 5. 看图评估并修复

1. 显示 PNG，亲眼看。
2. 按 `references/evaluation_rubric.md` 的 6 个维度打分。
3. 任一维度低于 7 分或有 warnings 就修复，最多 2 轮。
4. 结论里要说的数字，用 pandas 从数据**算出来**，不要目测图表。

## 6. 交付（示例）

> **各大区月度销售额趋势**
> （展示 PNG）
>
> - 华东始终最高，2025 年 6 月达 622 万元，比 2024 年 1 月增长 40%。
> - 西北规模最小（2025 年 6 月 184 万元），同期增长 33%。
>
> **数据处理**：「销售额（元）」从文本转为数字；按月、按大区汇总；单位换算为万元。
>
> **文件**：`sales_trend.png`、`sales_trend.xlsx`（打开「图表」工作表 → 选中图表 → 复制 → 在 PPT 中粘贴，即为可编辑的原生图表）、`sales_trend.py`

- 把 PNG 和 XLSX 作为可下载文件提供给用户。如果当前环境无法提供下载，如实说明，并把 `chart_data` 的表格以 Markdown 形式给出，方便用户粘贴到 Excel。
- `EXCEL_CHART["type"]` 为 `image` 或 `heatmap` 时，要说明这张图在 PPT 中不是可编辑的原生图表。

## 后续修改

用户说"改成柱状图 / 只看华东"时：另存为 `<slug>_v2.py`，修改 `plot`、`chart_data`、`EXCEL_CHART` 来满足每一条指令，然后重复第 4–6 步。

## 常见错误（务必避免）

- 对比率直接求和：应该用加权平均，即 `sum(分子) / sum(分母)`。
- 文本数字没有先 `to_number()`。
- 代码里读文件或调用 `plt.show()`：`data` 已经注入，且是全量数据。
- 列名和摘要里不一致，比如全角括号写成了半角。
- 类别超过 12 个不做 Top N 合并；用饼图比较超过 5 个部分。
- 结论里的数字是目测图表得来的。
