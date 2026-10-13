"""Smoke test: python tests/smoke_test.py (run from the skill directory or anywhere)."""

import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile

import pandas as pd

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(SKILL_DIR, "scripts")
SAMPLE = os.path.join(SKILL_DIR, "examples", "sample_sales.csv")
EXAMPLE_CODE = os.path.join(SKILL_DIR, "examples", "example_chart.py")
sys.path.insert(0, SCRIPTS)
from dataio import to_date, to_number  # noqa: E402
from excel_export import build_workbook  # noqa: E402
from openpyxl import load_workbook  # noqa: E402


def run(script, *args):
    # PYTHONIOENCODING=gbk reproduces piped stdout on Chinese Windows on every platform
    env = {**os.environ, "PYTHONIOENCODING": "gbk"}
    proc = subprocess.run([sys.executable, os.path.join(SCRIPTS, script), *args],
                          capture_output=True, text=True, encoding="utf-8", env=env)
    try:
        return proc.returncode, json.loads(proc.stdout)
    except json.JSONDecodeError:
        raise AssertionError(
            f"{script} did not print JSON (exit code {proc.returncode}).\n"
            f"--- stdout ---\n{proc.stdout}\n--- stderr ---\n{proc.stderr}") from None


def main():
    tmp = tempfile.mkdtemp()

    code, summary = run("summarize.py", SAMPLE, "--out", os.path.join(tmp, "s.json"))
    assert code == 0
    assert summary["field_names"][0] == "日期", "BOM / Chinese column names must be preserved"
    fields = {f["column"]: f["properties"] for f in summary["fields"]}
    assert fields["日期"]["dtype"] == "date"
    assert fields["销售额（元）"]["dtype"] == "number" and fields["销售额（元）"]["stored_as_text"]
    assert fields["大区"]["dtype"] == "category"

    png = os.path.join(tmp, "chart.png")
    code, result = run("render.py", "--data", SAMPLE, "--code", EXAMPLE_CODE, "--out", png)
    assert code == 0 and result["status"] == "ok" and os.path.getsize(png) > 10_000, result
    assert result["rows_used"] == 360, "render must use the full dataset"

    bad = os.path.join(tmp, "bad.py")
    with open(bad, "w", encoding="utf-8") as fh:
        fh.write("def plot(data):\n    return data.groupby('地区')['销量'].sum().plot.bar()\n")
    code, result = run("render.py", "--data", SAMPLE, "--code", bad, "--out", os.path.join(tmp, "b.png"))
    assert code == 1 and result["error_type"] == "KeyError" and "大区" in result["available_columns"]

    xlsx = os.path.join(tmp, "s.xlsx")
    pd.read_csv(SAMPLE, encoding="utf-8-sig").to_excel(xlsx, index=False)
    code, result = run("render.py", "--data", xlsx, "--code", EXAMPLE_CODE, "--out", png)
    assert code == 0, result

    nums = to_number(pd.Series(["¥1,234.5", "12%", "3.2万", "(50)", "88元", "abc"])).tolist()
    assert nums[:5] == [1234.5, 0.12, 32000.0, -50.0, 88.0] and pd.isna(nums[5])
    dates = to_date(pd.Series(["2024年1月5日", "20240105", "2024/3", "x"]))
    assert str(dates[0].date()) == "2024-01-05" and str(dates[1].date()) == "2024-01-05" and pd.isna(dates[3])

    # --- native Excel chart export
    xlsx = os.path.join(tmp, "chart.xlsx")
    code, result = run("render.py", "--data", SAMPLE, "--code", EXAMPLE_CODE, "--out", png, "--xlsx", xlsx)
    assert code == 0 and result["xlsx_warnings"] == [], result
    assert re.search("<lineChart>", chart_xml(xlsx))
    wb = load_workbook(xlsx)
    assert wb.sheetnames == ["图表", "数据", "PNG预览"], wb.sheetnames
    ws = wb["数据"]
    assert ws["A1"].value == "月份" and ws["A2"].value == "2024-01", "monthly dates must become 2024-01 labels"
    assert ws["B1"].value == "华东" and isinstance(ws["B2"].value, float)

    no_spec = os.path.join(tmp, "no_spec.py")
    with open(no_spec, "w", encoding="utf-8") as fh:
        fh.write("def plot(data):\n    fig, ax = plt.subplots(); ax.bar(['a'], [1]); ax.set_title('t'); return fig\n")
    code, result = run("render.py", "--data", SAMPLE, "--code", no_spec, "--out", png, "--xlsx", xlsx)
    assert code == 1 and result["stage"] == "xlsx" and "chart_data" in result["message"], result

    table = pd.DataFrame({"品类": ["配件", "手机", "平板"], "销售额": [3.0, 2.0, 1.0], "毛利率": [0.35, 0.18, 0.14]})
    cases = {
        "bar": ({"type": "bar", "title": "t", "number_format": "0.0%", "data_labels": True}, r'<barDir val="bar"\s*/>'),
        "pie": ({"type": "pie", "title": "t"}, "<pieChart>"),
        "combo": ({"type": "col", "title": "t", "secondary": ["毛利率"]}, "<lineChart>"),
        "heatmap": ({"type": "heatmap", "title": "t"}, None),
    }
    for name, (spec, marker) in cases.items():
        path = os.path.join(tmp, f"{name}.xlsx")
        build_workbook(table, spec, path, png)
        if marker:
            assert re.search(marker, chart_xml(path)), name  # lxml / stdlib serialize "/>" differently
    try:
        build_workbook(table, {"type": "radar"}, os.path.join(tmp, "bad.xlsx"))
        raise AssertionError("unsupported chart type must raise")
    except ValueError as exc:
        assert "不支持" in str(exc)

    print("all smoke tests passed")


def chart_xml(path: str) -> str:
    with zipfile.ZipFile(path) as zf:
        return zf.read("xl/charts/chart1.xml").decode("utf-8")


if __name__ == "__main__":
    main()
