"""Builds the installable web-app (PWA) copy for GitHub Pages.

Run:  python tools/make_pwa.py
Output folder: ../policy-review-app  (its own git repo; only app files, template and the fictional sample —
never real-test/ or any client data).
"""
import shutil
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT.parent / "policy-review-app"
VERSION = "2026-10-30h"          # bump to make installed apps pick up a new version

OUT.mkdir(exist_ok=True)
(OUT / "icons").mkdir(exist_ok=True)

# ---- index.html: the offline app + manifest, icons and service worker
html = (ROOT / "Policy Review Snapshot.html").read_text(encoding="utf-8")
head_add = """<link rel="manifest" href="manifest.webmanifest">
<meta name="theme-color" content="#6B3FA0">
<meta name="description" content="Turn a client's policy list into a bilingual coverage dashboard and keepsake PDF for annual reviews.">
<link rel="icon" href="icons/icon-192-report.png">
<link rel="apple-touch-icon" sizes="180x180" href="icons/apple-touch-icon-report.png">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-title" content="Policy Review">
"""
assert html.count("<title>Policy Review Snapshot</title>\n") == 1
html = html.replace("<title>Policy Review Snapshot</title>\n", "<title>Policy Review Snapshot</title>\n" + head_add)
sw_reg = """<script>
/* Installable app: cache the app so it opens offline after the first visit */
if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js');
</script>
</body>"""
# Installable build: lock screen visible from the first paint (the script then asks to set or enter the PIN)
assert html.count('<div id="lock" class="lock" hidden>') == 1
html = html.replace('<div id="lock" class="lock" hidden>', '<div id="lock" class="lock">')
assert html.count("</body>") == 1
html = html.replace("</body>", sw_reg)
(OUT / "index.html").write_text(html, encoding="utf-8")

# ---- manifest
(OUT / "manifest.webmanifest").write_text("""{
  "name": "Policy Review Snapshot",
  "short_name": "Policy Review",
  "description": "Bilingual coverage dashboard and keepsake PDF for annual policy reviews.",
  "start_url": "./",
  "scope": "./",
  "display": "standalone",
  "background_color": "#F7F3FB",
  "theme_color": "#6B3FA0",
  "icons": [
    {"src": "icons/icon-192-report.png", "sizes": "192x192", "type": "image/png", "purpose": "any"},
    {"src": "icons/icon-512-report.png", "sizes": "512x512", "type": "image/png", "purpose": "any"},
    {"src": "icons/icon-maskable-512-report.png", "sizes": "512x512", "type": "image/png", "purpose": "maskable"}
  ]
}
""", encoding="utf-8")

# ---- service worker: app shell + CDN libraries, cache first, refreshed in the background
(OUT / "sw.js").write_text(f"""const CACHE = 'policy-review-{VERSION}';
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icons/icon-192-report.png', 'icons/icon-512-report.png', 'icons/apple-touch-icon-report.png',
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js',
  'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js'];
self.addEventListener('install', e => e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())));
self.addEventListener('activate', e => e.waitUntil(
  caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())));
self.addEventListener('fetch', e => {{
  if (e.request.method !== 'GET') return;
  e.respondWith(caches.match(e.request).then(hit => {{
    const net = fetch(e.request).then(res => {{
      if (res.ok || res.type === 'opaque') caches.open(CACHE).then(c => c.put(e.request, res.clone()));
      return res;
    }}).catch(() => hit);
    return hit || net;
  }}));
}});
""", encoding="utf-8")

# ---- icons: design A "report" (see tools/icons.py)
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent))
from icons import draw_icon
icon = lambda size, pad=0.0: draw_icon("report", size, pad)

# icon files carry the design name: a new design = new URLs, so installed apps pick it up
for old in (OUT / "icons").glob("*.png"):
    old.unlink()
icon(192).save(OUT / "icons" / "icon-192-report.png")
icon(180).save(OUT / "icons" / "apple-touch-icon-report.png")
icon(512).save(OUT / "icons" / "icon-512-report.png")
icon(512, pad=0.08).save(OUT / "icons" / "icon-maskable-512-report.png")

# ---- supporting files (no client data)
shutil.copy(ROOT / "Policy_Review_Template.xlsx", OUT / "Policy_Review_Template.xlsx")
shutil.copy(ROOT / "Policy_Review_Template_blank.xlsx", OUT / "Policy_Review_Template_blank.xlsx")
(OUT / "sample").mkdir(exist_ok=True)
shutil.copy(ROOT / "sample-output" / "SAMPLE - Policy Review - TAN MEI LING.pdf", OUT / "sample" / "sample-report-fictional.pdf")
(OUT / "tools").mkdir(exist_ok=True)
for f in ("build_template.py", "embed_template.py", "icons.py", "make_pwa.py"):
    shutil.copy(ROOT / "tools" / f, OUT / "tools" / f)
# README.md in the app folder is written for the public repo (HOW TO USE.md mentions private files, so it is not copied)
(OUT / ".nojekyll").write_text("", encoding="utf-8")
print("Built", OUT)
