# lida-dataviz：Google Antigravity 用的数据可视化 skill

把 [Microsoft LIDA](https://github.com/microsoft/lida) 的流水线（摘要 → 目标 → 代码 → 执行 → 评估 → 修复）改造成一个 Antigravity Agent Skill。
你给出需求和原始数据，Antigravity 里的 Gemini Agent 就会生成合适的静态 PNG 图表。

与直接调用 `lida` Python 包的区别：

- **不需要额外的 LLM API Key**：LIDA 里由 LLM 完成的步骤（生成目标、写代码、评估、修复）由 Antigravity 的 Agent 自己完成，不再经过 `llmx`。
- **Agent 会看图**：评估阶段 Agent 会打开渲染出的 PNG 检查效果，而不只是读代码。
- **修复了 LIDA 在中文数据上的问题**：
  - 中文列名不再被替换成 `___`
  - 不再把数据抽样到 4500 行，图表基于全量数据
  - `.xlsx` 文件可以正常读取（原版向 `read_excel` 传了 `encoding` 参数，会报错）
  - 自动识别 GBK / UTF-8-BOM 编码
  - 自动配置中文字体，避免方块字
  - 识别并转换文本形式的数字和日期（`¥1,234`、`12%`、`3.2万`、`2024年1月5日`）

## 安装

**仅当前项目可用**：把整个 `lida-dataviz` 目录放到项目根目录的 `.agents/skills/` 下（本仓库已经放好）。
较早版本的 Antigravity 读取的是 `.agent/skills/`，新版本对它保持向后兼容。

**所有项目可用**：全局目录在不同的 Antigravity 版本和产品（App / IDE / CLI）之间不一致，
已知的有 `~/.gemini/config/skills/`、`~/.gemini/antigravity/skills/`、`~/.gemini/antigravity-cli/skills/`。
请以你所用版本的官方文档（antigravity.google/docs/skills）为准，例如：

```bash
mkdir -p ~/.gemini/config/skills
cp -r .agents/skills/lida-dataviz ~/.gemini/config/skills/
```

**Python 依赖**（Agent 首次使用时也会自动检查）：

```bash
pip install -r .agents/skills/lida-dataviz/scripts/requirements.txt
```

**中文字体**：
- macOS 和 Windows 自带中文字体，无需处理。
- Linux 需要安装，例如 `sudo apt install fonts-noto-cjk` 或 `fonts-wqy-zenhei`。

## 使用

在 Antigravity 里直接说，例如：

- 「用 data/sales.xlsx 画一下各大区每月销售额的趋势」
- 「这份 orders.csv 帮我看看有什么值得关注的」（Agent 会先给出 3–5 个分析方向让你选）
- 「把刚才那张图改成横向条形图，只保留前 10 名」

产物都在工作区的 `dataviz_output/` 目录下：
- `xxx.png`：图表
- `xxx.py`：绘图代码，可以复用
- `xxx_summary.json`：数据摘要

## 目录结构

```
lida-dataviz/
├── SKILL.md                     # Agent 读取的主说明（工作流）
├── scripts/
│   ├── summarize.py             # 数据摘要（对应 LIDA Summarizer，不调用 LLM）
│   ├── render.py                # 执行绘图代码，输出 PNG 和质量告警（对应 LIDA ChartExecutor）
│   ├── dataio.py                # 数据读取 + to_number / to_date 工具函数
│   └── requirements.txt
├── references/
│   ├── goal_generation.md       # 需求 → 可视化目标（对应 LIDA GoalExplorer / Persona）
│   ├── chart_selection.md       # 图表选型最佳实践
│   ├── code_template.md         # 代码模板与规则（对应 LIDA ChartScaffold / VizGenerator）
│   └── evaluation_rubric.md     # 6 维评估 + 修复（对应 LIDA VizEvaluator / VizRepairer）
├── examples/                    # 样例数据和样例图表代码
└── tests/smoke_test.py          # 冒烟测试：python tests/smoke_test.py
```

## 手动运行脚本

```bash
python scripts/summarize.py examples/sample_sales.csv
python scripts/render.py --data examples/sample_sales.csv --code examples/example_chart.py --out dataviz_output/example.png
```

## 安全说明

`render.py` 会用 `exec()` 执行绘图代码，和 LIDA 原版一样没有沙箱。
SKILL.md 要求 Agent 只渲染它自己写的代码。Antigravity 执行终端命令前会请求你的确认（取决于你的终端执行策略设置），请只在可信的工作区里使用。
