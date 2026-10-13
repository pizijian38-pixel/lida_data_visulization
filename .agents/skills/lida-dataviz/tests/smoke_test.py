"""Smoke test: python tests/smoke_test.py (run from the skill directory or anywhere)."""

import json
import os
import subprocess
import sys
import tempfile

import pandas as pd

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(SKILL_DIR, "scripts")
SAMPLE = os.path.join(SKILL_DIR, "examples", "sample_sales.csv")
EXAMPLE_CODE = os.path.join(SKILL_DIR, "examples", "example_chart.py")
sys.path.insert(0, SCRIPTS)
from dataio import to_date, to_number  # noqa: E402


def run(script, *args):
    proc = subprocess.run([sys.executable, os.path.join(SCRIPTS, script), *args],
                          capture_output=True, text=True, encoding="utf-8")
    return proc.returncode, json.loads(proc.stdout)


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

    print("all smoke tests passed")


if __name__ == "__main__":
    main()
