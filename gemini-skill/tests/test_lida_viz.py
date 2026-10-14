"""Tests for the Gemini-app build: python gemini-skill/tests/test_lida_viz.py"""

import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
GEMINI = os.path.dirname(HERE)
SKILL = os.path.join(GEMINI, "lida-dataviz")
SCRIPT = os.path.join(SKILL, "scripts", "lida_viz.py")
SAMPLE = os.path.join(SKILL, "examples", "sample_sales.csv")
EXAMPLE_CODE = os.path.join(SKILL, "examples", "example_chart.py")
sys.path.insert(0, os.path.dirname(SCRIPT))
import lida_viz as L  # noqa: E402


def cli(*args):
    # PYTHONIOENCODING=gbk reproduces piped stdout on Chinese Windows on every platform
    env = {**os.environ, "PYTHONIOENCODING": "gbk"}
    proc = subprocess.run([sys.executable, SCRIPT, *args], capture_output=True, text=True,
                          encoding="utf-8", env=env)
    try:
        return proc.returncode, json.loads(proc.stdout)
    except json.JSONDecodeError:
        raise AssertionError(f"no JSON from {args}\n--- stdout ---\n{proc.stdout}\n--- stderr ---\n{proc.stderr}") from None


def chart_xml(path):
    with zipfile.ZipFile(path) as zf:
        return zf.read("xl/charts/chart1.xml").decode("utf-8")


def main():
    tmp = tempfile.mkdtemp()

    # environment report
    code, env = cli("env")
    assert code == 0 and "can_export_xlsx" in env
    original = L.available_cjk_font
    L.available_cjk_font = lambda: None
    try:
        assert any("中文字体" in n for n in L.check_env()["notes"]), "missing CJK font must be reported"
    finally:
        L.available_cjk_font = original

    # summary (CLI and in-session)
    code, summary = cli("summarize", SAMPLE)
    assert code == 0 and summary["field_names"][0] == "日期"
    fields = {f["column"]: f["properties"] for f in summary["fields"]}
    assert fields["日期"]["dtype"] == "date" and fields["销售额（元）"]["stored_as_text"]
    assert L.summarize_file(SAMPLE)["n_rows"] == 360
    assert L.summarize_file(os.path.join(tmp, "missing.csv"))["status"] == "error"

    # render via CLI with xlsx
    png, xlsx = os.path.join(tmp, "c.png"), os.path.join(tmp, "c.xlsx")
    code, r = cli("render", "--data", SAMPLE, "--code", EXAMPLE_CODE, "--out", png, "--xlsx", xlsx)
    assert code == 0 and r["rows_used"] == 360 and r["xlsx_warnings"] == [], r
    assert re.search("<lineChart>", chart_xml(xlsx))

    # render in-session from a DataFrame and a code string (the "write lida_viz.py, then import" path)
    df = pd.read_csv(SAMPLE, encoding="utf-8-sig")
    with open(EXAMPLE_CODE, encoding="utf-8") as fh:
        code_text = fh.read()
    r = L.render_chart(df, code_text, os.path.join(tmp, "s.png"), os.path.join(tmp, "s.xlsx"))
    assert r["status"] == "ok" and r["rows_used"] == 360, r

    # errors point at the chart code
    r = L.render_chart(SAMPLE, "def plot(data):\n    return data.groupby('地区')['销量'].sum().plot.bar()\n",
                       os.path.join(tmp, "b.png"))
    assert r["status"] == "error" and r["error_type"] == "KeyError" and "大区" in r["available_columns"]
    r = L.render_chart(SAMPLE, "def plot(data):\n    fig, ax = plt.subplots(); ax.bar(['a'], [1]); ax.set_title('t'); return fig\n",
                       os.path.join(tmp, "n.png"), os.path.join(tmp, "n.xlsx"))
    assert r["status"] == "error" and r["stage"] == "xlsx" and "chart_data" in r["message"], r

    # excel chart types
    table = pd.DataFrame({"品类": ["配件", "手机", "平板"], "销售额": [3.0, 2.0, 1.0], "毛利率": [0.35, 0.18, 0.14]})
    for name, spec, marker in [("bar", {"type": "bar", "data_labels": True}, r'<barDir val="bar"\s*/>'),
                               ("pie", {"type": "pie"}, "<pieChart>"),
                               ("combo", {"type": "col", "secondary": ["毛利率"]}, "<lineChart>")]:
        path = os.path.join(tmp, f"{name}.xlsx")
        L.build_workbook(table, spec, path, png)
        assert re.search(marker, chart_xml(path)), name

    # helpers
    nums = L.to_number(pd.Series(["¥1,234.5", "12%", "3.2万", "(50)", "88元"])).tolist()
    assert nums == [1234.5, 0.12, 32000.0, -50.0, 88.0]
    assert str(L.to_date(pd.Series(["2024年1月5日"]))[0].date()) == "2024-01-05"

    # the package passes Google's upload rules
    proc = subprocess.run([sys.executable, os.path.join(GEMINI, "build_zip.py")], capture_output=True,
                          text=True, encoding="utf-8")
    assert proc.returncode == 0, proc.stdout + proc.stderr
    names = zipfile.ZipFile(os.path.join(GEMINI, "dist", "lida-dataviz.zip")).namelist()
    assert "SKILL.md" in names and not any(n.endswith((".pyc", ".xlsx")) for n in names), names

    print("all gemini skill tests passed")


if __name__ == "__main__":
    main()
