"""Validate the Gemini-app skill against Google's upload rules and package it for upload.

Usage:  python gemini-skill/build_zip.py
Output: gemini-skill/dist/lida-dataviz/      clean folder (upload this folder) and
        gemini-skill/dist/lida-dataviz.zip   the same files with SKILL.md at the zip root

Rules checked (Gemini Apps help, "Create & manage skills"):
- SKILL.md in the root folder; skill name all lowercase, words separated by hyphens
- plain-text files plus pdf / images only (no .xlsx, .docx or other binaries)
- no hidden binary files such as .DS_Store or .pyc
- total size <= 100 MB
- scripts must not need internet access
"""

from __future__ import annotations

import os
import re
import shutil
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.join(HERE, "lida-dataviz")
DIST = os.path.join(HERE, "dist")
MAX_BYTES = 100 * 1024 * 1024

ALLOWED_EXT = {
    ".txt", ".md", ".rst", ".rtf", ".tex", ".log",            # text & docs
    ".py", ".sh",                                               # code
    ".json", ".yaml", ".yml", ".csv", ".toml", ".xml", ".env", ".sql",  # data & config
    ".html", ".css", ".svg",                                    # web
    ".pdf", ".jpg", ".jpeg", ".png", ".gif", ".webp",           # rich media
}
ALLOWED_NAMES = {"Makefile", "Dockerfile"}
SKIP_DIRS = {"__pycache__", ".git", ".ipynb_checkpoints"}
NETWORK = re.compile(r"\b(requests|urllib|http\.client|socket|httpx|aiohttp)\b|pip install|urlopen")
NAME = re.compile(r"^[a-z0-9]+(-[a-z0-9]+)*$")


def collect() -> tuple[list[str], list[str], list[str]]:
    """Return (files to package, skipped junk, errors)."""
    files, skipped, errors = [], [], []
    for root, dirs, names in os.walk(SKILL_DIR):
        for d in list(dirs):
            if d in SKIP_DIRS or d.startswith("."):
                dirs.remove(d)
                skipped.append(os.path.relpath(os.path.join(root, d), SKILL_DIR) + "/")
        for name in names:
            rel = os.path.relpath(os.path.join(root, name), SKILL_DIR)
            ext = os.path.splitext(name)[1].lower()
            if name.startswith(".") or ext in {".pyc", ".pyo"}:
                skipped.append(rel)
            elif ext in ALLOWED_EXT or name in ALLOWED_NAMES:
                files.append(rel)
            else:
                errors.append(f"不支持的文件类型（Gemini 只接受纯文本、PDF 和图片）: {rel}")
    return sorted(files), skipped, errors


def check(files: list[str]) -> tuple[list[str], list[str]]:
    errors, warnings = [], []
    if "SKILL.md" not in files:
        errors.append("根目录缺少 SKILL.md")
    else:
        text = open(os.path.join(SKILL_DIR, "SKILL.md"), encoding="utf-8").read()
        front = re.match(r"^---\n(.*?)\n---\n", text, re.S)
        if not front:
            errors.append("SKILL.md 缺少 --- 包围的 frontmatter")
        else:
            meta = dict(line.split(":", 1) for line in front.group(1).splitlines() if ":" in line)
            name = meta.get("name", "").strip()
            if not NAME.match(name):
                errors.append(f"skill 名称必须全小写、用连字符分隔，当前是 {name!r}")
            if not meta.get("description", "").strip():
                errors.append("SKILL.md 缺少 description（Gemini 靠 name + description 判断何时使用）")
    total = sum(os.path.getsize(os.path.join(SKILL_DIR, f)) for f in files)
    if total > MAX_BYTES:
        errors.append(f"总大小 {total / 1e6:.1f} MB 超过 100 MB")
    for f in files:
        if f.endswith((".py", ".sh")):
            for lineno, line in enumerate(open(os.path.join(SKILL_DIR, f), encoding="utf-8"), start=1):
                for m in NETWORK.finditer(line):
                    if m.group(0) == "pip install" and ("不要" in line or "never" in line.lower()):
                        continue  # the "never pip install" guidance itself
                    errors.append(f"{f}:{lineno} 可能需要联网（{m.group(0)}），Gemini 中不支持")
    return errors, warnings


def main() -> int:
    files, skipped, errors = collect()
    more_errors, warnings = check(files)
    errors += more_errors
    for s in skipped:
        print(f"跳过: {s}")
    for w in warnings:
        print(f"警告: {w}")
    if errors:
        for e in errors:
            print(f"错误: {e}")
        return 1

    out_dir = os.path.join(DIST, "lida-dataviz")
    shutil.rmtree(out_dir, ignore_errors=True)
    for f in files:
        os.makedirs(os.path.dirname(os.path.join(out_dir, f)), exist_ok=True)
        shutil.copy2(os.path.join(SKILL_DIR, f), os.path.join(out_dir, f))
    zip_path = os.path.join(DIST, "lida-dataviz.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(os.path.join(SKILL_DIR, f), f.replace(os.sep, "/"))
    total = sum(os.path.getsize(os.path.join(SKILL_DIR, f)) for f in files)
    print(f"检查通过：{len(files)} 个文件，共 {total / 1024:.0f} KB")
    print(f"上传文件夹: {out_dir}")
    print(f"或上传 zip: {zip_path}")
    return 0


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.exit(main())
