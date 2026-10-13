---
name: lida-dataviz
description: 根据用户的需求描述和原始数据文件（CSV / Excel / JSON / Parquet 等）生成合适的静态数据可视化图表（PNG）。当用户提供数据并要求画图、做图表、可视化、看趋势/分布/对比/占比/相关性，或者只给数据让你"看看有什么发现"时使用。Generates data-faithful static PNG charts from raw data plus a natural-language request (chart, plot, graph, visualize, trend, distribution, comparison). Based on Microsoft LIDA's summarize → goal → code → execute → evaluate → repair pipeline.
---

# LIDA 数据可视化

本 skill 把 Microsoft LIDA 的流水线移植给你（Agent）执行：**你**扮演 LIDA 里的 LLM，
`scripts/` 里的两个脚本负责确定性的部分（数据摘要、执行代码出图）。

```
原始数据 ──summarize.py──▶ 数据摘要 ──你──▶ 可视化目标(goal) ──你──▶ 绘图代码
          ──render.py──▶ PNG + 质量告警 ──你看图评估──▶ 修复(≤2轮) ──▶ 交付
```

`<SKILL_DIR>` 指本文件所在目录（工作区安装时通常是 `.agents/skills/lida-dataviz`，
全局安装时在 `~/.gemini/...` 下）。不确定时用本文件的实际绝对路径。所有产物写到工作区根目录的
`dataviz_output/` 下。

## 第 0 步：环境检查（每个会话首次使用时）

```bash
python -c "import pandas, matplotlib, seaborn, openpyxl" || pip install -r <SKILL_DIR>/scripts/requirements.txt
```

## 第 1 步：生成数据摘要

```bash
python <SKILL_DIR>/scripts/summarize.py <数据文件> [--sheet 工作表名]
```

完整阅读输出的 JSON，重点看：
- `fields[].properties.dtype`：number / date / category / string / boolean
- `stored_as_text: true`：文本形式的数字或日期，代码里**必须**先用 `to_number()` / `to_date()` 转换
- `top_values`：类别列的取值分布，用来判断要不要 Top N
- `notes`：脚本发现的数据陷阱（宽表、缺失值多、类别过多等），逐条处理
- `preview_rows`：前 5 行，用来理解表的结构（长表 / 宽表）

脚本出错（如编码、工作表）时按 `message` 调整参数重试。Excel 有多个工作表且用户没指定时，先问用户。

## 第 2 步：把需求变成可视化目标

每个目标是一个 goal：`{question, visualization, rationale}`，其中 visualization **必须引用摘要里的确切列名**。

- **需求明确**（"各大区月度销售额趋势"）：直接写 1 个 goal；需求里有多个问题就拆成多个。
- **需求模糊或没有需求**（"帮我看看这份数据"）：按 `references/goal_generation.md` 生成 3–5 个候选目标，列给用户选；用户说"直接做"就做前 3 个。
- **会改变结论的歧义**（有两个候选指标列、口径不清、时间范围不明）才问用户，并且只问一次；其他情况自己选合理默认值，在交付时说明假设。

## 第 3 步：选图表类型

查 `references/chart_selection.md`。原则：先确定问题类型（趋势 / 比较 / 排名 / 分布 / 构成 / 关系 / 地理），再选图；默认不用饼图、双 Y 轴、3D 图。

## 第 4 步：写绘图代码

严格按 `references/code_template.md` 的模板和规则，写到 `dataviz_output/<简短英文slug>.py`。
要点：只写 `plot(data)` 函数，代码里**不读取任何文件**（`data` 由 render.py 注入，是全量数据）；
先在注释里写计划，再写代码。

## 第 5 步：渲染

```bash
python <SKILL_DIR>/scripts/render.py --data <数据文件> --code dataviz_output/<slug>.py --out dataviz_output/<slug>.png [--sheet 工作表名]
```

- `"status": "error"`：根据 `chart_code_lines`、`message`、`hint`、`available_columns` 修代码后重跑。最多 3 次，仍失败就把错误如实告诉用户。
- `"status": "ok"`：进入第 6 步。`warnings` 非空时必须处理。

## 第 6 步：看图评估并修复（LIDA 的 evaluate + repair）

1. **打开 PNG 图片亲眼看**，不要只看代码。
2. 按 `references/evaluation_rubric.md` 的 6 个维度打分（1–10）。
3. 出现以下任一情况就修改代码并重新渲染，最多 2 轮：render.py 有 `warnings`；任一维度低于 7 分；图上能看出问题（文字被截断、重叠、方块字、颜色难区分、图例遮挡数据）。
4. 修复时只改 `plot()` 函数内部，保持模板结构不变。

## 第 7 步：核实要引用的数字

交付时要说的任何具体数字（最大值、增长率、占比……）都必须用 pandas 从数据里**算出来**，不能目测图表：

```bash
python -c "import sys; sys.path.insert(0, '<SKILL_DIR>/scripts'); from dataio import read_data, to_number, to_date; df = read_data('<数据文件>'); ..."
```

## 第 8 步：交付

用用户的语言回复，包含：
1. 图片（`dataviz_output/<slug>.png`），在回复里展示或给出路径
2. 2–3 句结论，引用第 7 步算出的数字
3. 做了哪些数据处理和假设（例如"销售额从文本转成数字，按月汇总，单位换算为万元"）
4. 代码路径 `dataviz_output/<slug>.py`，方便复用

## 后续修改（LIDA 的 edit）

用户说"改成柱状图 / 只看华东 / 标题换成英文 / 颜色换成蓝色"时：复制成 `<slug>_v2.py`，只修改 `plot()` 内部来满足**每一条**指令，
然后重复第 5–8 步。保留旧版本，方便用户对比。

## 硬性规则

- 不修改、不覆盖用户的原始数据文件。
- 不抽样：图表必须基于全量数据（render.py 已保证），除非用户明确要求抽样。
- 只执行你自己写的绘图代码；render.py 会 `exec()` 代码，不要渲染来路不明的代码。
- 一个 `.py` 文件只出一张图；需要对比多个子群体时可以用子图（small multiples）。
- 图里的文字（标题、轴标签、图例）用用户的语言；单位要写清楚。
- 不要在代码里调用 `plt.show()`、`savefig()` 或读写文件。
