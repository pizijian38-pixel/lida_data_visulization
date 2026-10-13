"""Build a compact, LLM-friendly summary of a dataset (port of LIDA's Summarizer, no LLM call).

Usage:
    python summarize.py <data_file> [--sheet NAME] [--encoding ENC] [--n-samples 3] [--out summary.json]

Prints the summary JSON to stdout and also writes it to --out
(default: dataviz_output/<file stem>_summary.json in the current directory).
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import sys

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dataio import read_data, to_date, to_number  # noqa: E402

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


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("data")
    parser.add_argument("--sheet", default=None, help="Excel sheet name or index")
    parser.add_argument("--encoding", default=None)
    parser.add_argument("--n-samples", type=int, default=3)
    parser.add_argument("--out", default=None)
    args = parser.parse_args()

    sheet = int(args.sheet) if args.sheet and args.sheet.isdigit() else args.sheet
    try:
        df = read_data(args.data, sheet=sheet, encoding=args.encoding)
    except Exception as exc:  # report loading problems as JSON so the agent can react
        print(json.dumps({"status": "error", "message": str(exc)}, ensure_ascii=False))
        return 1

    summary = summarize(df, os.path.basename(args.data), args.n_samples)
    out = args.out or os.path.join(
        "dataviz_output", os.path.splitext(os.path.basename(args.data))[0] + "_summary.json")
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    text = json.dumps(summary, ensure_ascii=False, indent=2, default=to_jsonable)
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(text)
    print(text)
    print(f"\n[summary written to {out}]", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
