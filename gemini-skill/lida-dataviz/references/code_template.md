# 绘图代码模板与规则

移植自 LIDA `lida/components/scaffold.py` 和 `viz/vizgenerator.py`，改为默认输出静态 PNG。

## 模板（matplotlib / seaborn，默认）

```python
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns
# <其他需要的 import>

# 需求: <一句话复述用户需求>
# 计划:
# 1. 数据转换: <需要转换的列（to_number / to_date）、过滤、聚合、新建的列>
# 2. 图表类型: <图表类型 + 理由>
# 3. 编码: x=<列>, y=<列>, 颜色=<列>, 排序=<规则>
# 4. 美化: <标题、轴标签/单位、数字格式、标注>
def plot(data: pd.DataFrame):
    <只修改这里>
    ax.set_title("<结论式或描述式标题>")
    return fig

chart = plot(data)  # data 已经加载好；保留这一行，后面不要再写代码
```

需要输出 XLSX 时（默认），还要按下文「Excel 原生图表」一节定义 `EXCEL_CHART` 和 `chart_data(data)`，并让 `plot()` 基于 `chart_data()` 作图。

`lida_viz.render_chart` 会注入这些名字，可以直接使用：`data`（全量数据的副本）、`pd`、`np`、`plt`、`sns`、`to_number`、`to_date`。

## 必须遵守（LIDA 的通用指令，已改写）

1. **不要写任何加载数据的代码**：不用 `read_csv`、`open()`，也不要写 `data = ...` 去读文件。
2. **只用摘要 `field_names` 里存在的列**，或者由它们派生出来的新列。列名必须逐字一致（包括全角括号、空格）。
3. **文本数字**（摘要里 `stored_as_text: true` 且 `dtype: number`）：先写 `data["x"] = to_number(data["x"])`。
   `to_number` 会处理 `¥`、`,`、`%`（转成小数）、`万`、`亿`、`元` 和括号负数。
4. **日期**（`dtype: date`）：
   ```python
   data["日期"] = to_date(data["日期"])          # 支持 "2024年1月5日"、"20240105"、年份整数等
   data = data[pd.notna(data["日期"])]
   ```
   然后按需要的粒度汇总，例如 `pd.Grouper(key="日期", freq="MS")` 或 `dt.to_period("Q")`。
5. **先聚合再画图**：同一个 x 有多行时，必须先 `groupby` 并选择正确的聚合方式：
   金额、数量用 `sum`；单价、比率用 `mean`，或者用加权平均（`sum(分子)/sum(分母)`）。**不要对比率直接求和。**
6. 计算均值、中位数等统计量时使用 `numeric_only=True`。
7. **x 轴标签必须清晰可读**：标签长或多时旋转（`fig.autofmt_xdate()` 或 `plt.setp(ax.get_xticklabels(), rotation=30, ha="right")`），或者改用横向条形图。
8. **单一数值答案**：用 `axhline` / `axvline` 画参考线，图例里标注数值（`f"{v:,.2f}"`）。
9. 多系列时加图例，并用不同颜色区分；系列不超过 5 条时，优先在线尾直接标注，代替图例。
10. 标题写结论或清晰的描述；轴标签带单位；不要在标题里堆砌列名。
11. 返回 `fig`（推荐）、`ax` 或 seaborn 的 grid 对象。**不要**调用 `plt.show()`、`savefig()`、`sns.set_theme()`；字体和基础样式由 lida_viz 统一设置。
12. 用 `fig, ax = plt.subplots(figsize=(宽, 高))` 显式创建画布。
    默认尺寸 `(10, 6)`；横向条形图的高度随类别数增加，大约 `0.4 * n + 1.5`。
13. 需要的库都要 import。**环境没有网络，不能 pip install**：只用 `check_env()` 报告可用的库；seaborn 不可用时只用 matplotlib；地图等需要额外依赖的图退回到条形图。
14. `check_env()` 报告没有中文字体（`cjk_font` 为 null）时：PNG 里的标题、轴标签、图例改用英文（或拼音），否则会显示为方块；`EXCEL_CHART` 和 `chart_data` 的列名照常用中文，因为 Excel 图表由用户的 Office 渲染。

## 常用代码片段

**Top N + 其他**
```python
totals = data.groupby("品类")["销售额"].sum().sort_values(ascending=False)
top = totals.head(10)
if len(totals) > 10:
    top["其他"] = totals.iloc[10:].sum()
```

**横向条形图（降序，最大值在上）+ 数值标签**
```python
s = top.sort_values()
fig, ax = plt.subplots(figsize=(10, 0.4 * len(s) + 1.5))
bars = ax.barh(s.index, s.values, color="#4C72B0")
ax.bar_label(bars, labels=[f"{v:,.0f}" for v in s.values], padding=3)
ax.grid(axis="y", visible=False)
```

**万元单位**
```python
data["销售额（万元）"] = data["销售额"] / 1e4
ax.yaxis.set_major_formatter(mticker.StrMethodFormatter("{x:,.0f}"))
```

**百分比轴**
```python
ax.yaxis.set_major_formatter(mticker.PercentFormatter(1))
```

**日期刻度控制**
```python
import matplotlib.dates as mdates
ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
```

**宽表转长表**（列名是年份或月份时）
```python
long = data.melt(id_vars=["地区"], var_name="年份", value_name="值")
```

## Excel 原生图表（默认开启）

调用 `render_chart(..., out_xlsx=...)` 时，代码除了 `plot(data)`，还必须定义下面两样东西。lida_viz 会据此生成一个 `.xlsx`，里面有三个工作表：
- 「图表」：**原生 Excel 图表**，可编辑、矢量、无损缩放，复制到 PPT 后仍是原生图表
- 「数据」：图表引用的汇总表
- 「PNG预览」：渲染出的 PNG，方便对照

```python
EXCEL_CHART = {
    "type": "line",                 # 见下表
    "title": "各大区月度销售额趋势",
    "x_title": "月份",               # 可省略
    "y_title": "销售额（万元）",      # 可省略
    "number_format": "#,##0",       # 数值轴和数据单元格的 Excel 格式: "#,##0" / "#,##0.0" / "0.0%" / "0%"
    "data_labels": False,           # 在柱/条/点上显示数值
}

def chart_data(data: pd.DataFrame) -> pd.DataFrame:
    """PNG 和 Excel 共用的汇总表。宽表格式: 第 1 列 = 类别或 x 轴, 其余每列 = 一个系列。"""
    ...
    return table

def plot(data: pd.DataFrame):
    table = chart_data(data)        # 必须基于 chart_data 作图，保证 PNG 与 Excel 数字完全一致
    ...
```

**规则**
1. `chart_data` 返回的是**汇总后**的数据（通常几行到几十行），不是明细；行的顺序就是图上的显示顺序，排序在这里完成。
2. 列名就是 Excel 里的系列名和图例，用用户的语言，单位写在 `y_title` 里。
3. 百分比数据保持 0–1 的小数，用 `"number_format": "0.0%"` 显示；金额先换算成万元/亿元再返回。
4. 日期列可以直接返回 datetime：月度数据会显示为 `2024-01`，其他显示为 `2024-01-05`。
5. Excel 图表只能表达 `chart_data` 里的数据。PNG 上的额外标注（线尾标签、注释文字、参考线）不会出现在 Excel 里，这是正常的。
   参考线需要的话，在表里加一列常数值作为额外的系列（例如「平均值」）。

**`type` 与 PNG 图表类型的对应**

| PNG 里画的是 | `type` | 表格格式 / 说明 |
|---|---|---|
| 柱状图 | `col` | 第 1 列类别，每个系列一列 |
| 横向条形图 | `bar` | 同上；**表格第一行显示在最上面** |
| 堆积 / 百分比堆积柱 | `stacked_col` / `percent_col` | 第 1 列类别，每个组成部分一列 |
| 堆积 / 百分比堆积条形 | `stacked_bar` / `percent_bar` | 同上 |
| 折线图 | `line` | 第 1 列时间或类别，每条线一列 |
| 面积 / 堆积面积 | `area` / `stacked_area` | 同折线 |
| 饼图 / 环形图 | `pie` / `doughnut` | 第 1 列类别，第 2 列数值（≤ 8 块） |
| 散点图 | `scatter` | 宽表：第 1 列 x（数值），其余列 y；或长表加 `"x"`、`"y"`、`"group"` 三个键按组着色 |
| 柱 + 折线（次坐标轴） | `col` 加 `"secondary": ["列名"]` | 指定列画成右侧坐标轴上的折线，可加 `secondary_title`、`secondary_number_format`、`secondary_min` / `secondary_max` |
| 热力图 | `heatmap` | 第 1 列行标签，其余列为列标签；Excel 没有原生热力图，生成带色阶的表格 |
| 直方图 | `col` | 在 `chart_data` 里先分箱（`pd.cut` + `value_counts`），返回「区间, 频数」 |
| 箱线图、小提琴图、回归线、hexbin、小多图等 | `image` | Excel 没有原生对应，只嵌入 PNG 并附数据表；交付时告诉用户这张图在 PPT 里不可编辑 |

其他可选键：`colors`（十六进制颜色列表，默认与 matplotlib 的 tab10 一致）、`y_min` / `y_max`、`width_cm` / `height_cm`（默认 25.4 × 14.3 cm，16:9）。

## 其他库（仅在用户要求时使用）

- **plotnine**（ggplot 语法）：`plot(data)` 返回 `ggplot` 对象，lida_viz 会调用 `.draw()`。中文需要在主题里设置字体：
  `p9.theme(text=p9.element_text(family="<render 输出里的 cjk_font>"))`。
- **交互式图表**（plotly / altair）超出本 skill 的范围。
