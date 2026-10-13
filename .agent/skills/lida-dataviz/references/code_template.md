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

render.py 会注入这些名字，可以直接使用：`data`（全量数据的副本）、`pd`、`np`、`plt`、`sns`、`to_number`、`to_date`。

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
11. 返回 `fig`（推荐）、`ax` 或 seaborn 的 grid 对象。**不要**调用 `plt.show()`、`savefig()`、`sns.set_theme()`；字体和基础样式由 render.py 统一设置。
12. 用 `fig, ax = plt.subplots(figsize=(宽, 高))` 显式创建画布。
    默认尺寸 `(10, 6)`；横向条形图的高度随类别数增加，大约 `0.4 * n + 1.5`。
13. 需要的库都要 import。地图需要 geopandas 等额外依赖时，先检查是否已安装，没有就退回到条形图。

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

## 其他库（仅在用户要求时使用）

- **plotnine**（ggplot 语法）：`plot(data)` 返回 `ggplot` 对象，render.py 会调用 `.draw()`。中文需要在主题里设置字体：
  `p9.theme(text=p9.element_text(family="<render 输出里的 cjk_font>"))`。
- **交互式图表**（plotly / altair）超出本 skill 默认的 PNG 范围。用户要求时，另外写独立脚本导出 HTML。
