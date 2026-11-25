"""One-off: log in to Gofrugal HQ and find how HQ opens Smart Report (no secrets printed)."""
import os, re, requests
base = (os.environ.get("GF_URL") or "https://sanchuhq.gofrugal.com/RayMedi_HQ/").split("/RayMedi_HQ")[0] + "/RayMedi_HQ/"
root = base.split("/RayMedi_HQ")[0]
s = requests.Session(); s.headers["User-Agent"] = "Mozilla/5.0 Chrome/124"
s.get(base + "mainIndex.do", timeout=60)
r = s.post(base + "j_security_check", data={"j_username": os.environ["GF_USER"], "j_password": os.environ["GF_PASS"]}, timeout=60)
print("login ->", r.status_code, r.url, list(s.cookies.keys()))
m = s.get(base + "mainIndex.do", params={"page": "/smartreport/index.html#/reports?reportId=3123&productId=4"}, timeout=60)
print("mainIndex", m.status_code, len(m.text))
for x in re.findall(r'src=["\']([^"\']+)', m.text): print("  src", x)
def ctx(t, words, n=40, w=300):
    out, last = [], -1
    for mm in re.finditer(words, t):
        if mm.start() < last: continue
        out.append(re.sub(r"\s+", " ", t[max(0, mm.start()-w): mm.start()+w])); last = mm.start()+w
        if len(out) >= n: break
    return out
for js in re.findall(r'src=["\']([^"\']+\.js[^"\']*)', m.text):
    u = js if js.startswith("http") else (root + js if js.startswith("/") else base + js)
    t = s.get(u, timeout=60).text
    hits = ctx(t, r"smartreport|smartReport|SmartReport|reportsApi|WebReporter|sso|Sso|token=|Token")
    print(f"\n==== {js} {len(t)} bytes, {len(hits)} contexts")
    for h in hits: print("  ..", h[:600])
for p in ("rest/hq/product-type", "rest/hq/smartreport", "rest/hq/sr-token", "rest/hq/user"):
    x = s.get(base + p, timeout=60); print("GET", p, x.status_code, x.text[:300].replace("\n", " "))
