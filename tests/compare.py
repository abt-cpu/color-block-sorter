"""比對：同一批 .docx，Python 版與網頁版的分類結果必須逐組逐段完全一致。
用法: python3 compare.py 樣本資料夾 網頁結果資料夾
（先用 make_samples.py 產生樣本，再用 web_run.js 產生網頁結果）
"""
import json
import subprocess
import sys
from pathlib import Path
from docx import Document

samples, webdir = Path(sys.argv[1]), Path(sys.argv[2])
tool = Path(__file__).resolve().parent.parent / "python" / "sort_colors.py"
bad = 0
for src in sorted(samples.glob("*.docx")):
    if src.stem.endswith("_分類"):
        continue
    subprocess.run([sys.executable, str(tool), str(src)], check=True, capture_output=True)
    groups = []
    for p in Document(str(src.with_name(src.stem + "_分類.docx"))).paragraphs:
        if p.style.name.startswith("Heading"):
            groups.append([])
        elif groups:
            groups[-1].append(p.text)
    web = json.load(open(webdir / (src.stem + ".json"), encoding="utf-8"))
    same = [g for _, g in web["groups"]] == groups
    print("PASS" if same else "FAIL", src.name, f"({web['total']} 字, {len(groups)} 組)")
    bad += not same
sys.exit(1 if bad else 0)
