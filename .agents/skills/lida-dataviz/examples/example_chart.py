import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

# 需求: 各大区月度销售额趋势
# 计划:
# 1. 销售额（元）是文本 -> to_number; 日期是中文文本 -> to_date, 丢弃 NaT
# 2. 按 月份 x 大区 汇总销售额, 换算成「万元」
# 3. 折线图: x=月份, y=销售额(万元), 每个大区一条线, 线尾直接标注大区名, 不用图例
def plot(data: pd.DataFrame):
    data["销售额"] = to_number(data["销售额（元）"])
    data["日期"] = to_date(data["日期"])
    data = data[pd.notna(data["日期"])]
    monthly = (data.groupby([pd.Grouper(key="日期", freq="MS"), "大区"])["销售额"]
               .sum().div(1e4).unstack("大区"))
    order = monthly.iloc[-1].sort_values(ascending=False).index
    fig, ax = plt.subplots(figsize=(11, 6))
    for region in order:
        line, = ax.plot(monthly.index, monthly[region], marker="o", markersize=3, linewidth=2)
        ax.annotate(f"{region} {monthly[region].iloc[-1]:,.0f}", (monthly.index[-1], monthly[region].iloc[-1]),
                    xytext=(6, 0), textcoords="offset points", va="center", color=line.get_color())
    ax.set_title("各大区月度销售额趋势（2024.01–2025.06）")
    ax.set_xlabel("月份")
    ax.set_ylabel("销售额（万元）")
    ax.yaxis.set_major_formatter(mticker.StrMethodFormatter("{x:,.0f}"))
    ax.xaxis.set_major_formatter(plt.matplotlib.dates.DateFormatter("%Y-%m"))
    fig.autofmt_xdate()
    return fig

chart = plot(data)
