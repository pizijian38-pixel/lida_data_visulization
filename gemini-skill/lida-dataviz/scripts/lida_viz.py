"""lida_viz: LIDA-style data visualization helpers for the Gemini app skill (single file, no internet).

One self-contained module so it works both ways:
  * as a script:   python lida_viz.py env | summarize <data> | render --data .. --code .. --out .. [--xlsx ..]
  * in a Python session:  from lida_viz import check_env, summarize_file, render_chart
    (if the skill files are not on disk, write this file's content to lida_viz.py first, then import it)

Ported from Microsoft LIDA: Summarizer (summarize_file), ChartExecutor (render_chart).
Differences from LIDA: unicode column names are kept, rows are never sampled, Excel is read without the
invalid `encoding` argument, CSV encodings are detected, CJK fonts are configured, and an .xlsx with a
native, editable Excel chart can be exported for PowerPoint. Only pandas + matplotlib are required;
seaborn and openpyxl are optional.
"""

from __future__ import annotations

import argparse
import json
import logging
import math
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

try:  # optional: only needed for the .xlsx export
    from openpyxl import Workbook
    from openpyxl.chart import (AreaChart, BarChart, DoughnutChart, LineChart, PieChart, Reference,
                                ScatterChart, Series)
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.marker import Marker
    from openpyxl.chart.series import DataPoint
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.drawing.line import LineProperties
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False



# ==============================================================================
# Data loading and parsing helpers (to_number / to_date are injected into chart code)
# ==============================================================================
CSV_ENCODINGS = ["utf-8-sig", "gb18030", "big5", "latin-1"]
UNIT_MULTIPLIER = {"万": 1e4, "亿": 1e8, "k": 1e3, "K": 1e3, "m": 1e6, "M": 1e6}
SUPPORTED = (".csv", ".tsv", ".txt", ".json", ".jsonl", ".xlsx", ".xlsm", ".xls",
             ".parquet", ".feather")


def force_utf8_stdio() -> None:
    """Emit UTF-8 regardless of platform. On Chinese Windows, piped stdout defaults to GBK,
    which cannot encode characters such as '¥' and makes the scripts crash with no output."""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def _read_text_table(path: str, sep, encoding: str | None) -> pd.DataFrame:
    encodings = [encoding] if encoding else CSV_ENCODINGS
    last_error = None
    for enc in encodings:
        try:
            # sep=None lets the python engine sniff , ; | or tab
            return pd.read_csv(path, sep=sep, encoding=enc,
                               engine="python" if sep is None else "c")
        except UnicodeDecodeError as exc:
            last_error = exc
    raise ValueError(f"无法识别文件编码，尝试过 {encodings}: {last_error}")


def read_data(path: str, sheet: str | int | None = None,
              encoding: str | None = None) -> pd.DataFrame:
    """Read a data file into a DataFrame with cleaned (but unicode-preserving) column names."""
    if not os.path.exists(path):
        raise FileNotFoundError(f"数据文件不存在: {path}")
    ext = os.path.splitext(path)[1].lower()

    if ext == ".csv":
        df = _read_text_table(path, sep=None, encoding=encoding)
    elif ext in (".tsv", ".txt"):
        df = _read_text_table(path, sep="\t" if ext == ".tsv" else None, encoding=encoding)
    elif ext == ".json":
        try:
            df = pd.read_json(path, orient="records", encoding=encoding or "utf-8")
        except ValueError:
            df = pd.read_json(path, lines=True, encoding=encoding or "utf-8")
    elif ext == ".jsonl":
        df = pd.read_json(path, lines=True, encoding=encoding or "utf-8")
    elif ext in (".xlsx", ".xlsm", ".xls"):
        df = pd.read_excel(path, sheet_name=0 if sheet is None else sheet)
    elif ext == ".parquet":
        df = pd.read_parquet(path)
    elif ext == ".feather":
        df = pd.read_feather(path)
    else:
        raise ValueError(f"不支持的文件类型 {ext}，支持: {', '.join(SUPPORTED)}")

    df.columns = clean_column_names(df.columns)
    # drop fully empty rows/columns that spreadsheets often carry
    df = df.dropna(how="all").dropna(axis=1, how="all")
    return df.reset_index(drop=True)


def clean_column_names(columns) -> list[str]:
    """Trim whitespace, name blank/'Unnamed' columns, and de-duplicate. Keeps Chinese etc."""
    cleaned, seen = [], {}
    for i, col in enumerate(columns):
        name = re.sub(r"\s+", " ", str(col).replace("\ufeff", "")).strip()
        if not name or name.lower().startswith("unnamed:"):
            name = f"列_{i + 1}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        cleaned.append(name)
    return cleaned


# Helpers also injected into chart code by render.py
def to_number(series: pd.Series) -> pd.Series:
    """'¥1,234.5' / '1234元' -> 1234.5, '12%' -> 0.12, '3.2万' -> 32000, '(50)' -> -50; unparseable -> NaN."""
    if pd.api.types.is_numeric_dtype(series):
        return series.astype(float)
    text = series.astype("string").str.strip()
    negative = text.str.match(r"^[(（].*[)）]$").fillna(False)
    text = text.str.replace(r"[()（）\s,，¥￥$€£+元]", "", regex=True)
    percent = text.str.endswith("%").fillna(False)
    unit = text.str.extract(r"(万|亿|k|K|m|M)$")[0]
    numeric = pd.to_numeric(text.str.replace(r"(%|万|亿|k|K|m|M)$", "", regex=True), errors="coerce")
    numeric = numeric.astype(float)
    numeric = numeric * unit.map(UNIT_MULTIPLIER).fillna(1.0).astype(float).values
    numeric = numeric.where(~percent.values, numeric / 100.0)
    numeric = numeric.where(~negative.values, -numeric)
    return pd.Series(numeric.values, index=series.index, name=series.name)


def to_date(series: pd.Series) -> pd.Series:
    """Parse dates robustly (mixed formats, 2024年1月5日, 20240105, year ints); unparseable -> NaT."""
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    if pd.api.types.is_numeric_dtype(series):
        s = series.dropna()
        if len(s) and s.between(1900, 2100).all():
            return pd.to_datetime(series.astype("Int64").astype("string"), format="%Y", errors="coerce")
        series = series.astype("Int64").astype("string")
    text = series.astype("string").str.strip()
    text = text.str.replace(r"(\d+)年(\d+)月(\d+)日?", r"\1-\2-\3", regex=True)
    text = text.str.replace(r"(\d+)年(\d+)月", r"\1-\2", regex=True)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            return pd.to_datetime(text, errors="coerce", format="mixed")
        except (TypeError, ValueError):
            return pd.to_datetime(text, errors="coerce")


# ==============================================================================
# Dataset summary (LIDA Summarizer, no LLM call)
# ==============================================================================
DATE_LIKE_NAME = re.compile(
    r"^(19|20)\d{2}([-/年.]?(0?[1-9]|1[0-2])月?)?$|^(0?[1-9]|1[0-2])月$|^Q[1-4]$|^(19|20)\d{2}\s*Q[1-4]$",
    re.IGNORECASE)
YEAR_NAME = re.compile(r"(year|年份|年度|^年$)", re.IGNORECASE)


def to_jsonable(value):
    if value is None:
        return None
    if isinstance(value, (np.integer,)):
        return int(value)
    if isinstance(value, (np.floating, float)):
        f = float(value)
        return None if math.isnan(f) or math.isinf(f) else round(f, 4)
    if isinstance(value, (np.bool_,)):
        return bool(value)
    if isinstance(value, (pd.Timestamp, np.datetime64)):
        return str(pd.Timestamp(value))
    if isinstance(value, (pd.Timedelta,)):
        return str(value)
    return value if isinstance(value, (str, int, bool)) else str(value)


def is_text_series(s: pd.Series) -> bool:
    # pandas 3 uses a dedicated string dtype, pandas 2 uses object
    return pd.api.types.is_object_dtype(s) or pd.api.types.is_string_dtype(s)


def sample_values(s: pd.Series, n: int) -> list:
    uniques = pd.Series(s.dropna().unique())
    if len(uniques) == 0:
        return []
    return [to_jsonable(v) for v in uniques.sample(min(n, len(uniques)), random_state=42).tolist()]


def number_stats(s: pd.Series) -> dict:
    return {
        "min": to_jsonable(s.min()), "max": to_jsonable(s.max()),
        "mean": to_jsonable(s.mean()), "median": to_jsonable(s.median()),
        "std": to_jsonable(s.std()),
    }


def describe_column(name: str, s: pd.Series, n_samples: int) -> tuple[dict, list[str]]:
    notes: list[str] = []
    non_null = s.dropna()
    n_non_null = len(non_null)
    props: dict = {}

    if pd.api.types.is_bool_dtype(s):
        props["dtype"] = "boolean"
    elif pd.api.types.is_numeric_dtype(s):
        props["dtype"] = "number"
        props.update(number_stats(non_null))
        if (YEAR_NAME.search(name) and n_non_null and (non_null % 1 == 0).all()
                and non_null.between(1900, 2100).all()):
            props["semantic_hint"] = "year"
            notes.append(f"「{name}」是年份数字，作为时间轴时用 pd.to_datetime(data['{name}'].astype(int).astype(str), format='%Y') 或直接当有序类别。")
        elif n_non_null and non_null.nunique() <= 12 and (non_null % 1 == 0).all():
            props["semantic_hint"] = "low-cardinality integer (可能是编码/等级/类别)"
    elif pd.api.types.is_datetime64_any_dtype(s):
        props["dtype"] = "date"
        props["min"], props["max"] = to_jsonable(non_null.min()), to_jsonable(non_null.max())
    elif is_text_series(s):
        text = non_null.astype(str).str.strip()
        as_number = to_number(text) if n_non_null else None
        numeric_like = as_number.notna().mean() if n_non_null else 0
        all_digits = text.str.fullmatch(r"\d+").mean() if n_non_null else 0
        parsed_dates = to_date(text) if n_non_null and all_digits < 0.9 else None

        if n_non_null and numeric_like >= 0.9:
            props["dtype"] = "number"
            props["stored_as_text"] = True
            props.update(number_stats(as_number.dropna()))
            pct = " 注意百分号会被转成小数（18.6% → 0.186），坐标轴用 PercentFormatter(1)。" if text.str.endswith("%").mean() > 0.5 else ""
            notes.append(f"「{name}」是以文本存储的数字（含千分位/货币符号/百分号/万等），绘图前必须用 to_number(data['{name}']) 转换。{pct}")
        elif parsed_dates is not None and parsed_dates.notna().mean() >= 0.9:
            props["dtype"] = "date"
            props["stored_as_text"] = True
            props["min"], props["max"] = to_jsonable(parsed_dates.min()), to_jsonable(parsed_dates.max())
            notes.append(f"「{name}」是文本格式的日期，绘图前用 to_date(data['{name}']) 转换并丢弃 NaT。")
        else:
            nunique = text.nunique()
            ratio = nunique / n_non_null if n_non_null else 1
            props["dtype"] = "category" if (nunique <= 30 or ratio < 0.5) else "string"
            if props["dtype"] == "category":
                top = text.value_counts().head(10)
                props["top_values"] = {str(k): int(v) for k, v in top.items()}
                if nunique > 12:
                    notes.append(f"「{name}」有 {nunique} 个类别，柱状图/图例中建议只保留 Top N，其余合并为「其他」。")
    else:
        props["dtype"] = str(s.dtype)

    props["samples"] = sample_values(s, n_samples)
    props["num_unique_values"] = int(s.nunique())
    missing = int(s.isna().sum())
    props["missing_count"] = missing
    if missing:
        props["missing_pct"] = round(missing / len(s) * 100, 1) if len(s) else 0.0
        if props["missing_pct"] >= 20:
            notes.append(f"「{name}」缺失 {props['missing_pct']}%，注意在图中说明或先过滤。")
    return {"column": name, "properties": props}, notes


def summarize(df: pd.DataFrame, file_name: str, n_samples: int = 3) -> dict:
    fields, notes = [], []
    for col in df.columns:
        field, col_notes = describe_column(col, df[col], n_samples)
        fields.append(field)
        notes.extend(col_notes)

    date_like_cols = [c for c in df.columns if DATE_LIKE_NAME.match(str(c).strip())]
    if len(date_like_cols) >= 3:
        notes.append(
            f"列名 {date_like_cols[:6]}... 看起来是时间点，这是宽表。做趋势图前先 "
            f"data.melt(id_vars=[...], value_vars={date_like_cols[:3]}..., var_name='时间', value_name='值') 转成长表。")

    preview = df.head(5).map(to_jsonable) if hasattr(df, "map") else df.head(5).applymap(to_jsonable)
    return {
        "name": os.path.splitext(file_name)[0],
        "file_name": file_name,
        "n_rows": int(len(df)),
        "n_columns": int(df.shape[1]),
        "field_names": [str(c) for c in df.columns],
        "fields": fields,
        "preview_rows": preview.to_dict(orient="records"),
        "notes": notes,
    }


# ==============================================================================
# Native Excel chart export (.xlsx)
# ==============================================================================
CHART_TYPES = {
    "col": "簇状柱形图", "stacked_col": "堆积柱形图", "percent_col": "百分比堆积柱形图",
    "bar": "簇状条形图（横向）", "stacked_bar": "堆积条形图", "percent_bar": "百分比堆积条形图",
    "line": "折线图", "area": "面积图", "stacked_area": "堆积面积图",
    "pie": "饼图", "doughnut": "环形图", "scatter": "散点图",
    "heatmap": "热力表（单元格色阶，Excel 没有原生热力图）",
    "image": "仅嵌入 PNG（箱线图、小提琴图、回归线等 Excel 无原生对应的图）",
}
SPEC_KEYS = {"type", "title", "x_title", "y_title", "number_format", "data_labels", "colors",
             "secondary", "secondary_title", "secondary_number_format", "secondary_min", "secondary_max",
             "y_min", "y_max",
             "x", "y", "group", "width_cm", "height_cm"}
# matplotlib's default tab10 cycle, so the Excel chart matches the PNG
DEFAULT_COLORS = ["1F77B4", "FF7F0E", "2CA02C", "D62728", "9467BD",
                  "8C564B", "E377C2", "7F7F7F", "BCBD22", "17BECF"]
HEADER_FILL = PatternFill("solid", fgColor="F2F2F2") if HAS_OPENPYXL else None
NOTE = "图表引用「数据」工作表中的单元格，修改数据后图表自动更新。复制图表并粘贴到 PowerPoint 即为可编辑的原生图表。"


def _labels(col: pd.Series) -> pd.Series:
    """Datetime categories -> readable labels (2024-01 for monthly data, 2024-01-05 otherwise)."""
    if pd.api.types.is_datetime64_any_dtype(col):
        fmt = "%Y-%m" if (col.dropna().dt.day == 1).all() else "%Y-%m-%d"
        return col.dt.strftime(fmt)
    if isinstance(col.dtype, pd.PeriodDtype):
        return col.astype(str)
    return col


def _cell(value):
    if value is None or (isinstance(value, float) and pd.isna(value)) or value is pd.NA or value is pd.NaT:
        return None
    if hasattr(value, "item"):  # numpy scalar
        return value.item()
    return value


def _normalize(table: pd.DataFrame, spec: dict, warnings: list[str]) -> pd.DataFrame:
    if not isinstance(table, pd.DataFrame):
        raise TypeError(f"chart_data(data) 必须返回 DataFrame，实际返回了 {type(table).__name__}。")
    table = table.reset_index() if not isinstance(table.index, pd.RangeIndex) else table.copy()
    table.columns = [str(c) for c in table.columns]
    kind = spec["type"]

    if kind == "scatter" and spec.get("group"):
        missing = [spec.get(k) for k in ("x", "y", "group") if spec.get(k) not in table.columns]
        if missing:
            raise ValueError(f"散点图分组模式需要 x/y/group 三列都在 chart_data 中，缺少 {missing}。")
        for key in ("x", "y"):
            table[spec[key]] = to_number(table[spec[key]])
        return table

    if table.shape[1] < 2:
        raise ValueError("chart_data(data) 至少需要 2 列：第 1 列是类别/x，其余列是数值系列。")
    first = table.columns[0]
    # assign by column name: iloc assignment would cast the labels back to datetime
    table[first] = to_number(table[first]) if kind == "scatter" else _labels(table[first]).astype(str)
    for col in table.columns[1:]:
        if not pd.api.types.is_numeric_dtype(table[col]):
            table[col] = to_number(table[col])
        if table[col].isna().all():
            raise ValueError(f"系列「{col}」没有任何数值，检查 chart_data 的聚合结果。")

    n = len(table)
    if kind in ("pie", "doughnut"):
        if table.shape[1] > 2:
            warnings.append(f"{CHART_TYPES[kind]}只使用第一个数值列「{table.columns[1]}」，其余列被忽略。")
        if n > 8:
            warnings.append(f"{CHART_TYPES[kind]}有 {n} 块，太多了：合并为 Top 5 +「其他」或改用条形图。")
    elif kind not in ("heatmap", "image", "scatter") and n > 60:
        warnings.append(f"Excel 图表有 {n} 个类别/时间点，会很拥挤：考虑按更粗的时间粒度汇总或保留 Top N。")
    if n > 5000:
        raise ValueError(f"chart_data 返回了 {n} 行。Excel 图表应基于汇总后的数据，而不是明细。")
    return table


def _write_table(ws, table: pd.DataFrame, number_format: str | None, start_row: int = 1) -> None:
    for j, name in enumerate(table.columns, start=1):
        cell = ws.cell(row=start_row, column=j, value=name)
        cell.font, cell.fill = Font(bold=True), HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
    for i, row in enumerate(table.itertuples(index=False), start=start_row + 1):
        for j, value in enumerate(row, start=1):
            cell = ws.cell(row=i, column=j, value=_cell(value))
            if number_format and j > 1 and isinstance(cell.value, (int, float)):
                cell.number_format = number_format
    for j, name in enumerate(table.columns, start=1):
        width = max([len(str(name))] + [len(str(v)) for v in table.iloc[:, j - 1].head(200)])
        ws.column_dimensions[get_column_letter(j)].width = min(max(10, width * 1.6 + 2), 40)
    ws.freeze_panes = ws.cell(row=start_row + 1, column=2)


def _color(spec: dict, i: int) -> str:
    colors = spec.get("colors") or DEFAULT_COLORS
    return str(colors[i % len(colors)]).lstrip("#").upper()


def _marker(color: str, size: int) -> Marker:
    marker = Marker(symbol="circle", size=size)
    marker.graphicalProperties = GraphicalProperties(solidFill=color, ln=LineProperties(solidFill=color))
    return marker


def _style_series(chart, spec: dict, offset: int = 0) -> None:
    for i, s in enumerate(chart.series):
        color = _color(spec, i + offset)
        if isinstance(chart, LineChart):
            s.graphicalProperties.line.solidFill = color
            s.graphicalProperties.line.width = 28575  # 2.25pt
            s.smooth = False
            s.marker = _marker(color, 5)
        elif isinstance(chart, ScatterChart):
            s.marker = _marker(color, 6)
            s.graphicalProperties.line.noFill = True
        else:
            s.graphicalProperties.solidFill = color
            s.graphicalProperties.line.solidFill = color


def _axes(chart, spec: dict, kind: str) -> None:
    chart.x_axis.delete = False  # openpyxl >= 3.1 hides axes unless this is set
    chart.y_axis.delete = False
    chart.x_axis.title = spec.get("x_title") or None
    chart.y_axis.title = spec.get("y_title") or None
    if spec.get("number_format"):
        chart.y_axis.number_format = spec["number_format"]
        chart.y_axis.numFmt.sourceLinked = False
    if spec.get("y_min") is not None:
        chart.y_axis.scaling.min = spec["y_min"]
    if spec.get("y_max") is not None:
        chart.y_axis.scaling.max = spec["y_max"]
    chart.x_axis.tickLblPos = "low"
    if kind in ("bar", "stacked_bar", "percent_bar"):
        # keep the table order top-to-bottom and the value axis at the bottom
        chart.x_axis.scaling.orientation = "maxMin"
        chart.y_axis.crosses = "max"


def _build_chart(ws_data, table: pd.DataFrame, spec: dict):
    kind = spec["type"]
    n_rows, n_cols = len(table), table.shape[1]
    cats = Reference(ws_data, min_col=1, min_row=2, max_row=n_rows + 1)
    secondary = [c for c in spec.get("secondary") or [] if c in table.columns[1:]]
    primary_cols = [j for j, c in enumerate(table.columns, start=1) if j > 1 and c not in secondary]

    if kind in ("pie", "doughnut"):
        chart = PieChart() if kind == "pie" else DoughnutChart()
        chart.add_data(Reference(ws_data, min_col=2, min_row=1, max_row=n_rows + 1), titles_from_data=True)
        chart.set_categories(cats)
        for i in range(n_rows):
            point = DataPoint(idx=i)
            point.graphicalProperties.solidFill = _color(spec, i)
            chart.series[0].dPt.append(point)
        chart.dataLabels = DataLabelList(showPercent=True, showCatName=True, showVal=False,
                                         showSerName=False, showLeaderLines=True)
        chart.legend.position = "r"
        return chart

    if kind == "scatter":
        chart = ScatterChart()
        chart.style = 13
        if spec.get("group"):
            x_col = list(table.columns).index(spec["x"]) + 1
            y_col = list(table.columns).index(spec["y"]) + 1
            g = table[spec["group"]]
            # table was written sorted by group, so each group is a contiguous block of rows
            row = 2
            for name, size in g.value_counts(sort=False).sort_index().items():
                xs = Reference(ws_data, min_col=x_col, min_row=row, max_row=row + size - 1)
                ys = Reference(ws_data, min_col=y_col, min_row=row, max_row=row + size - 1)
                chart.series.append(Series(ys, xs, title=str(name)))
                row += size
        else:
            xs = Reference(ws_data, min_col=1, min_row=2, max_row=n_rows + 1)
            for j in range(2, n_cols + 1):
                ys = Reference(ws_data, min_col=j, min_row=1, max_row=n_rows + 1)
                chart.series.append(Series(ys, xs, title_from_data=True))
        _style_series(chart, spec)
        _axes(chart, spec, kind)
        chart.x_axis.tickLblPos = "nextTo"
        return chart

    if kind in ("line",):
        chart = LineChart()
    elif kind in ("area", "stacked_area"):
        chart = AreaChart()
        chart.grouping = "stacked" if kind == "stacked_area" else "standard"
    else:
        chart = BarChart()
        chart.type = "bar" if kind.endswith("bar") else "col"
        chart.grouping = {"stacked_col": "stacked", "stacked_bar": "stacked",
                          "percent_col": "percentStacked", "percent_bar": "percentStacked"}.get(kind, "clustered")
        if chart.grouping != "clustered":
            chart.overlap = 100
        chart.gapWidth = 60

    for j in primary_cols:
        chart.add_data(Reference(ws_data, min_col=j, min_row=1, max_row=n_rows + 1), titles_from_data=True)
    chart.set_categories(cats)
    _style_series(chart, spec)
    _axes(chart, spec, kind)

    if spec.get("data_labels"):
        chart.dataLabels = DataLabelList(showVal=True, showSerName=False, showCatName=False,
                                         showLegendKey=False, showPercent=False)

    if secondary:
        line = LineChart()
        for c in secondary:
            j = list(table.columns).index(c) + 1
            line.add_data(Reference(ws_data, min_col=j, min_row=1, max_row=n_rows + 1), titles_from_data=True)
        line.set_categories(cats)
        _style_series(line, spec, offset=len(primary_cols))
        line.y_axis.axId = 200
        line.y_axis.delete = False
        line.y_axis.crosses = "max"
        line.y_axis.majorGridlines = None
        line.y_axis.title = spec.get("secondary_title") or None
        if spec.get("secondary_min") is not None:
            line.y_axis.scaling.min = spec["secondary_min"]
        if spec.get("secondary_max") is not None:
            line.y_axis.scaling.max = spec["secondary_max"]
        if spec.get("secondary_number_format"):
            line.y_axis.number_format = spec["secondary_number_format"]
            line.y_axis.numFmt.sourceLinked = False
        chart += line
    return chart


def build_workbook(table: pd.DataFrame, spec: dict, out_path: str, png_path: str | None = None) -> list[str]:
    """Write the workbook; return quality warnings. Raises ValueError/TypeError on an invalid spec."""
    if not HAS_OPENPYXL:
        raise RuntimeError("当前环境没有 openpyxl，无法生成 .xlsx。只交付 PNG，并告诉用户原因。")
    if not isinstance(spec, dict):
        raise TypeError("EXCEL_CHART 必须是 dict。")
    spec = dict(spec)
    kind = spec.get("type")
    if kind not in CHART_TYPES:
        raise ValueError(f"EXCEL_CHART['type'] = {kind!r} 不支持。可选: " +
                         ", ".join(f"{k}（{v}）" for k, v in CHART_TYPES.items()))
    unknown = set(spec) - SPEC_KEYS
    warnings: list[str] = [f"EXCEL_CHART 里有未识别的键 {sorted(unknown)}，已忽略。"] if unknown else []

    table = _normalize(table, spec, warnings)
    if kind == "scatter" and spec.get("group"):
        table = table.sort_values(spec["group"], kind="stable").reset_index(drop=True)

    wb = Workbook()
    ws_chart = wb.active
    ws_chart.title = "图表"
    ws_chart.page_setup.orientation = "landscape"
    ws_chart.page_setup.fitToWidth, ws_chart.page_setup.fitToHeight = 1, 0
    ws_chart.sheet_properties.pageSetUpPr.fitToPage = True
    ws_data = wb.create_sheet("数据")
    _write_table(ws_data, table, spec.get("number_format"))

    if kind == "heatmap":
        _write_table(ws_chart, table, spec.get("number_format"), start_row=3)
        ws_chart["A1"] = spec.get("title") or ""
        ws_chart["A1"].font = Font(bold=True, size=14)
        last = f"{get_column_letter(table.shape[1])}{len(table) + 3}"
        ws_chart.conditional_formatting.add(
            f"B4:{last}", ColorScaleRule(start_type="min", start_color="F7FBFF", mid_type="percentile",
                                         mid_value=50, mid_color="C6DBEF", end_type="max", end_color="4292C6"))
        ws_chart.cell(row=len(table) + 5, column=1, value="Excel 没有原生热力图，这里用单元格色阶代替；在 PPT 中可直接粘贴为表格。")
    elif kind == "image":
        if not png_path:
            raise ValueError("type='image' 需要先成功渲染 PNG。")
        warnings.append("该图表类型在 Excel 中没有原生对应，只嵌入了 PNG（不可编辑），数据仍在「数据」工作表。")
        img = XLImage(png_path)
        scale = 900 / img.width
        img.width, img.height = int(img.width * scale), int(img.height * scale)
        ws_chart.add_image(img, "A1")
    else:
        chart = _build_chart(ws_data, table, spec)
        chart.title = spec.get("title") or None
        chart.width = spec.get("width_cm", 25.4)   # 16:9, fits a PPT slide body
        chart.height = spec.get("height_cm", 14.3)
        if table.shape[1] <= 2 and kind not in ("pie", "doughnut"):
            chart.legend = None
        elif chart.legend is not None and kind not in ("pie", "doughnut"):
            chart.legend.position = "b"
        ws_chart.add_chart(chart, "A1")
        ws_chart.cell(row=31, column=1, value=NOTE).font = Font(color="808080")

    if png_path and os.path.exists(png_path) and kind != "image":
        ws_png = wb.create_sheet("PNG预览")
        img = XLImage(png_path)
        scale = 900 / img.width
        img.width, img.height = int(img.width * scale), int(img.height * scale)
        ws_png.add_image(img, "A1")

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    wb.save(out_path)
    return warnings


# ==============================================================================
# Chart execution, styling and quality checks (LIDA ChartExecutor)
# ==============================================================================
logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)

CJK_FONTS = [
    "Noto Sans CJK SC", "Source Han Sans SC", "Noto Sans SC", "Microsoft YaHei", "PingFang SC",
    "Hiragino Sans GB", "Heiti SC", "STHeiti", "SimHei", "WenQuanYi Zen Hei", "WenQuanYi Micro Hei",
    "Sarasa Gothic SC", "Arial Unicode MS",
]
CJK_CHAR = re.compile(r"[㐀-鿿豈-﫿　-〿＀-￯]")


# ---------------------------------------------------------------- styling
CJK_FONT_HINT = re.compile(r"CJK|Han Sans|Han Serif|Hei|Song|Ming|Kai|YaHei|Gothic|Mincho|Droid Sans Fallback|"
                           r"WenQuanYi|Noto Sans (SC|TC|JP|KR)|Unifont", re.IGNORECASE)


def available_cjk_font() -> str | None:
    """Preferred CJK fonts first, then any installed font whose name suggests CJK coverage."""
    installed = sorted({f.name for f in font_manager.fontManager.ttflist})
    preferred = next((name for name in CJK_FONTS if name in installed), None)
    return preferred or next((name for name in installed if CJK_FONT_HINT.search(name)), None)


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
        return namespace["chart"], namespace
    if callable(namespace.get("plot")):
        return namespace["plot"](data), namespace
    raise NameError("代码里没有定义 plot(data) 函数。")


def export_xlsx(namespace: dict, data: pd.DataFrame, code: str, xlsx_path: str, png_path: str) -> dict:
    """Build the native-Excel-chart workbook from chart_data(data) + EXCEL_CHART."""
    missing = [name for name in ("chart_data", "EXCEL_CHART") if name not in namespace]
    if missing:
        return {"status": "error", "stage": "xlsx",
                "message": f"使用 --xlsx 时，代码里必须定义 {' 和 '.join(missing)}（见 references/code_template.md 的「Excel 原生图表」一节）。",
                "png": os.path.abspath(png_path)}
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            table = namespace["chart_data"](data.copy())
        xlsx_warnings = build_workbook(table, namespace["EXCEL_CHART"], xlsx_path, png_path)
    except Exception as exc:
        payload = execution_error(exc, code, data)
        payload.update(stage="xlsx", png=os.path.abspath(png_path))
        return payload
    return {"status": "ok", "xlsx": os.path.abspath(xlsx_path), "xlsx_warnings": xlsx_warnings}


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


# ==============================================================================
# Public API (returns JSON-serializable dicts and never raises) and command line
# ==============================================================================
def check_env() -> dict:
    """Which libraries and CJK fonts are available. Run this first in a new environment."""
    def version(module_name):
        try:
            return __import__(module_name).__version__
        except Exception:
            return None
    info = {"python": sys.version.split()[0], "pandas": pd.__version__, "matplotlib": matplotlib.__version__,
            "seaborn": version("seaborn"), "openpyxl": version("openpyxl"), "cjk_font": available_cjk_font()}
    info["can_export_xlsx"] = HAS_OPENPYXL
    info["notes"] = []
    if not HAS_OPENPYXL:
        info["notes"].append("没有 openpyxl：只能输出 PNG，不能生成 .xlsx。不要尝试 pip install（无网络）。")
    if not info["cjk_font"]:
        info["notes"].append("没有中文字体：PNG 里的中文会显示为方块。PNG 的标题/标签改用英文；XLSX 里照常用中文（由 Office 渲染）。")
    if not info["seaborn"]:
        info["notes"].append("没有 seaborn：只用 matplotlib 作图。")
    return info


def _sheet_arg(sheet):
    return int(sheet) if isinstance(sheet, str) and sheet.isdigit() else sheet


def summarize_file(path: str, sheet=None, encoding: str | None = None, n_samples: int = 3,
                   out: str | None = None) -> dict:
    """Summarize a data file. Writes the summary JSON to `out` when given."""
    try:
        df = read_data(path, sheet=_sheet_arg(sheet), encoding=encoding)
    except Exception as exc:
        return {"status": "error", "stage": "load", "message": str(exc)}
    summary = summarize(df, os.path.basename(path), n_samples)
    if out:
        os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
        with open(out, "w", encoding="utf-8") as fh:
            json.dump(summary, fh, ensure_ascii=False, indent=2, default=to_jsonable)
    return json.loads(json.dumps(summary, default=to_jsonable))


def render_chart(data, code: str, out_png: str, out_xlsx: str | None = None, sheet=None,
                 encoding: str | None = None, dpi: int = 200) -> dict:
    """Run chart code on the FULL dataset; save a PNG and optionally an .xlsx with a native Excel chart.

    data: path to the data file, or a DataFrame.  code: path to a .py file, or the code itself.
    """
    try:
        if isinstance(data, pd.DataFrame):
            data = data.copy()
            data.columns = clean_column_names(data.columns)
        else:
            data = read_data(data, sheet=_sheet_arg(sheet), encoding=encoding)
        if "\n" not in code and code.endswith(".py") and os.path.exists(code):
            with open(code, encoding="utf-8") as fh:
                code = fh.read()
    except Exception as exc:
        return {"status": "error", "stage": "load", "message": str(exc)}

    cjk_font = available_cjk_font()
    apply_style(dpi, cjk_font)
    plt.close("all")
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            chart, namespace = run_chart_code(code, data.copy())
            fig = resolve_figure(chart)
    except Exception as exc:
        return execution_error(exc, code, data)

    apply_font(cjk_font)
    os.makedirs(os.path.dirname(os.path.abspath(out_png)), exist_ok=True)
    try:
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            fig.canvas.draw()
            found = quality_warnings(fig, bool(CJK_CHAR.search(figure_text(fig))), cjk_font)
            fig.savefig(out_png, bbox_inches="tight", pad_inches=0.25, facecolor="white")
    except Exception as exc:
        return {"status": "error", "stage": "save", "error_type": type(exc).__name__,
                "message": str(exc), "traceback": traceback.format_exc()[-3000:]}
    finally:
        plt.close("all")

    if any("missing from font" in str(w.message) for w in caught):
        found.append("有字符在当前字体中缺失（显示为方块）。没有中文字体时，PNG 的文字改用英文。")
    width, height = fig.get_size_inches() * dpi
    result = {"status": "ok", "png": os.path.abspath(out_png), "size_px": [int(width), int(height)],
              "rows_used": int(len(data)), "cjk_font": cjk_font, "warnings": found}
    if out_xlsx:
        xlsx = export_xlsx(namespace, data, code, out_xlsx, out_png)
        if xlsx["status"] != "ok":
            return xlsx
        result.update(xlsx=xlsx["xlsx"], xlsx_warnings=xlsx["xlsx_warnings"])
    return result


def main(argv=None) -> int:
    force_utf8_stdio()
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("env", help="report available libraries and CJK fonts")
    p_sum = sub.add_parser("summarize", help="summarize a data file")
    p_sum.add_argument("data")
    p_sum.add_argument("--sheet", default=None)
    p_sum.add_argument("--encoding", default=None)
    p_sum.add_argument("--n-samples", type=int, default=3)
    p_sum.add_argument("--out", default=None)
    p_ren = sub.add_parser("render", help="run chart code, save PNG (+ optional XLSX)")
    p_ren.add_argument("--data", required=True)
    p_ren.add_argument("--code", required=True)
    p_ren.add_argument("--out", required=True)
    p_ren.add_argument("--xlsx", default=None)
    p_ren.add_argument("--sheet", default=None)
    p_ren.add_argument("--encoding", default=None)
    p_ren.add_argument("--dpi", type=int, default=200)
    args = parser.parse_args(argv)

    if args.command == "env":
        result = check_env()
    elif args.command == "summarize":
        result = summarize_file(args.data, args.sheet, args.encoding, args.n_samples, args.out)
    else:
        result = render_chart(args.data, args.code, args.out, args.xlsx, args.sheet, args.encoding, args.dpi)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 1 if result.get("status") == "error" else 0


if __name__ == "__main__":
    sys.exit(main())
