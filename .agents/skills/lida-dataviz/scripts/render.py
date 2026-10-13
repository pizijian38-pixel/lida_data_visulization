"""Execute chart code against the FULL dataset and save a static PNG (port of LIDA's ChartExecutor).

Usage:
    python render.py --data <data_file> --code <chart.py> --out <chart.png> [--sheet S] [--encoding E] [--dpi 200]

Chart code contract (same idea as LIDA's scaffold):
    - defines `plot(data: pd.DataFrame)` and returns a matplotlib Figure / Axes / plt,
      a seaborn grid, or a plotnine ggplot object;
    - must NOT load data itself; `data` is injected (a copy of the full dataset);
    - `pd`, `np`, `plt`, `sns`, `to_number`, `to_date` are pre-imported helpers.

Prints one JSON object to stdout:
    {"status": "ok", "png": ..., "warnings": [...]}            exit code 0
    {"status": "error", "stage": ..., "message": ..., "traceback": ...}   exit code 1
Warnings are quality problems the agent should fix (overlapping labels, missing CJK glyphs, ...).
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import re
import sys
import traceback
import warnings

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
import pandas as pd  # noqa: E402
from matplotlib import font_manager  # noqa: E402

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dataio import force_utf8_stdio, read_data, to_date, to_number  # noqa: E402

logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)

CJK_FONTS = [
    "Noto Sans CJK SC", "Source Han Sans SC", "Noto Sans SC", "Microsoft YaHei", "PingFang SC",
    "Hiragino Sans GB", "Heiti SC", "STHeiti", "SimHei", "WenQuanYi Zen Hei", "WenQuanYi Micro Hei",
    "Sarasa Gothic SC", "Arial Unicode MS",
]
CJK_CHAR = re.compile(r"[㐀-鿿豈-﫿　-〿＀-￯]")


# ---------------------------------------------------------------- styling
def available_cjk_font() -> str | None:
    installed = {f.name for f in font_manager.fontManager.ttflist}
    return next((name for name in CJK_FONTS if name in installed), None)


def apply_style(dpi: int, cjk_font: str | None) -> None:
    plt.rcParams.update({
        "figure.figsize": (10, 6),
        "figure.dpi": 100,
        "savefig.dpi": dpi,
        "axes.spines.top": False,
        "axes.spines.right": False,
        "axes.grid": True,
        "grid.alpha": 0.3,
        "grid.linestyle": "--",
        "axes.titlesize": 14,
        "axes.titleweight": "bold",
        "axes.labelsize": 11,
        "legend.frameon": False,
        "axes.unicode_minus": False,
    })
    apply_font(cjk_font)


def apply_font(cjk_font: str | None) -> None:
    # seaborn.set_theme() resets the font list, so this is re-applied after the chart code runs
    plt.rcParams["axes.unicode_minus"] = False
    if cjk_font:
        plt.rcParams["font.family"] = "sans-serif"
        sans = [f for f in plt.rcParams["font.sans-serif"] if f != cjk_font]
        plt.rcParams["font.sans-serif"] = [cjk_font] + sans


# ---------------------------------------------------------------- execution
def resolve_figure(chart):
    if chart is None or chart is plt:
        return plt.gcf()
    if isinstance(chart, matplotlib.figure.Figure):
        return chart
    if isinstance(chart, matplotlib.axes.Axes):
        return chart.figure
    if isinstance(chart, np.ndarray) and chart.size and isinstance(chart.flat[0], matplotlib.axes.Axes):
        return chart.flat[0].figure
    if hasattr(chart, "draw") and hasattr(chart, "save") and not hasattr(chart, "savefig"):
        return chart.draw()  # plotnine ggplot
    for attr in ("figure", "fig"):  # seaborn FacetGrid / JointGrid / PairGrid
        fig = getattr(chart, attr, None)
        if isinstance(fig, matplotlib.figure.Figure):
            return fig
    raise TypeError(f"plot(data) 返回了无法识别的对象类型 {type(chart).__name__}，请返回 fig、ax 或 plt。")


def run_chart_code(code: str, data: pd.DataFrame):
    namespace = {"__name__": "__chart__", "pd": pd, "np": np, "plt": plt, "data": data,
                 "to_number": to_number, "to_date": to_date}
    try:
        import seaborn as sns
        namespace["sns"] = sns
    except ImportError:
        pass
    exec(compile(code, "<chart_code>", "exec"), namespace)
    if "chart" in namespace:
        return namespace["chart"]
    if callable(namespace.get("plot")):
        return namespace["plot"](data)
    raise NameError("代码里没有定义 plot(data) 函数。")


# ---------------------------------------------------------------- quality checks
def overlapping(labels, renderer) -> bool:
    boxes = [t.get_window_extent(renderer) for t in labels if t.get_visible() and t.get_text().strip()]
    return any(a.overlaps(b) for a, b in zip(boxes, boxes[1:]))


def quality_warnings(fig, has_cjk_text: bool, cjk_font: str | None) -> list[str]:
    found = []
    renderer = fig.canvas.get_renderer()
    axes = [ax for ax in fig.axes if ax.get_visible() and ax.get_label() != "<colorbar>"]
    if not axes:
        return ["图中没有任何坐标轴，可能 plot() 没有画到返回的 figure 上。"]
    for i, ax in enumerate(axes):
        name = f"子图{i + 1}" if len(axes) > 1 else "图表"
        has_marks = bool(ax.lines or ax.patches or ax.collections or ax.images or ax.tables)
        if not has_marks and not ax.texts:
            found.append(f"{name} 是空的（没有任何数据标记）。检查过滤条件或列名。")
        if overlapping(ax.get_xticklabels(), renderer):
            found.append(f"{name} 的 x 轴标签互相重叠：旋转标签、减少刻度、改用横向条形图或缩写标签。")
        if overlapping(ax.get_yticklabels(), renderer):
            found.append(f"{name} 的 y 轴标签互相重叠：减少类别（Top N）或加大图高。")
        legend = ax.get_legend()
        if legend is not None and len(legend.get_texts()) > 12:
            found.append(f"{name} 的图例有 {len(legend.get_texts())} 项，太多了：合并为 Top N +「其他」或改用分面。")
    titles = [ax.get_title() for ax in axes] + [fig._suptitle.get_text() if fig._suptitle else ""]
    if not any(t.strip() for t in titles):
        found.append("图表没有标题。")
    if has_cjk_text and not cjk_font:
        found.append("系统未找到中文字体，中文可能显示为方块。请安装 Noto Sans CJK SC / 思源黑体，或把中文标签换成英文。")
    return found


def figure_text(fig) -> str:
    return " ".join(t.get_text() for t in fig.findobj(matplotlib.text.Text))


def execution_error(exc: Exception, code: str, data: pd.DataFrame) -> dict:
    """Report the error at the chart-code line that caused it, not deep inside pandas."""
    source = code.splitlines()
    frames = [f for f in traceback.extract_tb(exc.__traceback__) if f.filename == "<chart_code>"]
    if isinstance(exc, SyntaxError) and exc.lineno:
        where = [f"line {exc.lineno}: {source[exc.lineno - 1].strip() if exc.lineno <= len(source) else ''}"]
    else:
        where = [f"line {f.lineno}: {source[f.lineno - 1].strip()}" for f in frames if f.lineno <= len(source)]
    payload = {"status": "error", "stage": "execute", "error_type": type(exc).__name__,
               "message": str(exc), "chart_code_lines": where,
               "library_frame": traceback.format_exception_only(type(exc), exc)[-1].strip()}
    if isinstance(exc, KeyError):
        payload["hint"] = "列名不存在。只能使用 available_columns 中的列名（注意全角括号、空格），或代码里新建的列。"
        payload["available_columns"] = list(map(str, data.columns))
    return payload


def emit(payload: dict, code: int) -> int:
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return code


def main() -> int:
    force_utf8_stdio()
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--data", required=True)
    parser.add_argument("--code", required=True, help="path to the chart .py file")
    parser.add_argument("--out", required=True, help="output .png path")
    parser.add_argument("--sheet", default=None)
    parser.add_argument("--encoding", default=None)
    parser.add_argument("--dpi", type=int, default=200)
    args = parser.parse_args()

    try:
        sheet = int(args.sheet) if args.sheet and args.sheet.isdigit() else args.sheet
        data = read_data(args.data, sheet=sheet, encoding=args.encoding)
        with open(args.code, encoding="utf-8") as fh:
            code = fh.read()
    except Exception as exc:
        return emit({"status": "error", "stage": "load", "message": str(exc)}, 1)

    cjk_font = available_cjk_font()
    apply_style(args.dpi, cjk_font)
    plt.close("all")

    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            chart = run_chart_code(code, data.copy())
            fig = resolve_figure(chart)
    except Exception as exc:
        return emit(execution_error(exc, code, data), 1)

    apply_font(cjk_font)
    out_dir = os.path.dirname(os.path.abspath(args.out))
    os.makedirs(out_dir, exist_ok=True)
    try:
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            fig.canvas.draw()
            found = quality_warnings(fig, bool(CJK_CHAR.search(figure_text(fig))), cjk_font)
            fig.savefig(args.out, bbox_inches="tight", pad_inches=0.25, facecolor="white")
    except Exception as exc:
        return emit({"status": "error", "stage": "save", "error_type": type(exc).__name__,
                     "message": str(exc), "traceback": traceback.format_exc()[-3000:]}, 1)
    finally:
        plt.close("all")

    if any("missing from font" in str(w.message) for w in caught):
        found.append("有字符在当前字体中缺失（显示为方块）。检查中文字体或特殊符号。")
    width, height = fig.get_size_inches() * args.dpi
    return emit({"status": "ok", "png": os.path.abspath(args.out),
                 "size_px": [int(width), int(height)], "rows_used": int(len(data)),
                 "cjk_font": cjk_font, "warnings": found}, 0)


if __name__ == "__main__":
    sys.exit(main())
