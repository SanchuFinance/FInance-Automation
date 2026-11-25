"""One-off: log in to Gofrugal HQ and list how Smart Report fetches report data (no secrets printed)."""
import os, re, sys, requests, urllib.parse
base = os.environ["GF_URL"].split("/RayMedi_HQ")[0] + "/RayMedi_HQ/"
s = requests.Session(); s.headers["User-Agent"] = "Mozilla/5.0 Chrome/124"
s.get(base + "mainIndex.do", timeout=60)
r = s.post(base + "j_security_check", data={"j_username": os.environ["GF_USER"], "j_password": os.environ["GF_PASS"]}, timeout=60)
print("login ->", r.status_code, r.url, "len", len(r.text), "cookies", list(s.cookies.keys()))
print("login page title:", re.findall(r"<title>(.*?)</title>", r.text, re.S)[:1], "| error?", bool(re.search(r"invalid|incorrect", r.text, re.I)))
h = s.get(base + "smartreport/index.html", timeout=60)
print("smartreport:", h.status_code, h.url, len(h.text))
scripts = re.findall(r'src=["\']([^"\']+\.js[^"\']*)', h.text)
print("scripts:", scripts)
pat = re.compile(r'["\'`]([A-Za-z0-9_/\.\-]{3,80})["\'`]')
for sc in scripts:
    u = urllib.parse.urljoin(h.url, sc)
    t = s.get(u, timeout=60).text
    hits = sorted({m for m in pat.findall(t) if re.search(r"report|Report|export|Export|download|excel|Excel|\.do$|api/|filter|Filter|outlet|Outlet", m)})
    print(f"\n== {sc} ({len(t)} bytes) {len(hits)} hits")
    for x in hits[:150]: print("  ", x)
    for m in re.finditer(r"(\.(?:post|get)\(\s*[^,)]{0,120})", t):
        print("  CALL", m.group(1)[:160])
