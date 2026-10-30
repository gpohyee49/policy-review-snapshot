"""Puts Policy_Review_Template_blank.xlsx inside the app so its "Download blank template" button works
everywhere (GitHub Pages, claude.ai, offline file).

Run after tools/build_template.py:  python tools/embed_template.py
"""
import base64
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
app = ROOT / "Policy Review Snapshot.html"
b64 = base64.b64encode((ROOT / "Policy_Review_Template_blank.xlsx").read_bytes()).decode("ascii")
html = app.read_text(encoding="utf-8")
html, n = re.subn(r"const BLANK_TEMPLATE_B64 = '[A-Za-z0-9+/=]*';", f"const BLANK_TEMPLATE_B64 = '{b64}';", html)
assert n == 1, "marker not found"
app.write_text(html, encoding="utf-8")
print(f"Embedded template ({len(b64):,} characters)")
