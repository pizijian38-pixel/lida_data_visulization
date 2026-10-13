"""Shared data loading for summarize.py and render.py.

Differences from LIDA's lida.utils.read_dataframe, on purpose:
- column names keep non-ASCII characters (LIDA turned "销售额" into "___");
- rows are never sampled, so charts are drawn from the full dataset;
- Excel is read without an `encoding` argument (pandas rejects it);
- CSV encoding is detected (utf-8 with/without BOM, gb18030, big5, latin-1).
"""

from __future__ import annotations

import os
import re
import warnings

import pandas as pd

CSV_ENCODINGS = ["utf-8-sig", "gb18030", "big5", "latin-1"]
UNIT_MULTIPLIER = {"万": 1e4, "亿": 1e8, "k": 1e3, "K": 1e3, "m": 1e6, "M": 1e6}
SUPPORTED = (".csv", ".tsv", ".txt", ".json", ".jsonl", ".xlsx", ".xlsm", ".xls",
             ".parquet", ".feather")


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
