import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import matplotlib.dates as mdates

# 需求: 各大区月度销售额趋势
# 计划:
# 1. 数据转换: 销售额（元）是文本 -> to_number; 日期是中文文本 -> to_date, 丢弃 NaT
#    按 月份 x 大区 汇总销售额, 换算成「万元」, 转成宽表（每个大区一列）
# 2. 图表类型: 折线图（时间趋势, 5 条线）
# 3. 编码: x=月份, y=销售额(万元), 每个大区一条线, 按最后一期销售额排序
# 4. 美化: 线尾直接标注大区名代替图例, 千分位, 月份刻度 YYYY-MM

EXCEL_CHART = {
    "type": "line",
    "title": "各大区月度销售额趋势（2024.01–2025.06）",
    "x_title": "月份",
    "y_title": "销售额（万元）",
    "number_format": "#,##0",
}


def chart_data(data: pd.DataFrame) -> pd.DataFrame:
    """PNG 和 Excel 共用的汇总表: 第 1 列 = 月份, 其余每列 = 一个大区。"""
    data["销售额"] = to_number(data["销售额（元）"])
    data["日期"] = to_date(data["日期"])
    data = data[pd.notna(data["日期"])]
    monthly = (data.groupby([pd.Grouper(key="日期", freq="MS"), "大区"])["销售额"]
               .sum().div(1e4).unstack("大区"))
    order = monthly.iloc[-1].sort_values(ascending=False).index
    return monthly[order].reset_index().rename(columns={"日期": "月份"})


def plot(data: pd.DataFrame):
    table = chart_data(data)
    fig, ax = plt.subplots(figsize=(11, 6))
    for region in table.columns[1:]:
        line, = ax.plot(table["月份"], table[region], marker="o", markersize=3, linewidth=2)
        ax.annotate(f"{region} {table[region].iloc[-1]:,.0f}", (table["月份"].iloc[-1], table[region].iloc[-1]),
                    xytext=(6, 0), textcoords="offset points", va="center", color=line.get_color())
    ax.set_title(EXCEL_CHART["title"])
    ax.set_xlabel(EXCEL_CHART["x_title"])
    ax.set_ylabel(EXCEL_CHART["y_title"])
    ax.yaxis.set_major_formatter(mticker.StrMethodFormatter("{x:,.0f}"))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
    fig.autofmt_xdate()
    return fig


chart = plot(data)
