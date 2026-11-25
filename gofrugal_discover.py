"""One-off: open Gofrugal HQ in a headless browser, run Smart Report 3123/3126 and log the data calls.
Secrets are never printed (login request bodies are skipped)."""
import os, re, json, asyncio
from playwright.async_api import async_playwright
BASE = (os.environ.get("GF_URL") or "https://sanchuhq.gofrugal.com/RayMedi_HQ/").split("/RayMedi_HQ")[0] + "/RayMedi_HQ/"
SKIP = re.compile(r"\.(js|css|png|jpg|svg|woff2?|ttf|ico|gif)(\?|$)")

async def main():
    async with async_playwright() as p:
        b = await p.chromium.launch()
        ctx = await b.new_context(viewport={"width": 1500, "height": 900}, accept_downloads=True)
        page = await ctx.new_page()
        log = []
        async def on_resp(r):
            u = r.url
            if SKIP.search(u) or "j_security_check" in u: return
            req = r.request
            body = (req.post_data or "")[:1500] if req.method == "POST" else ""
            try: txt = (await r.text())[:800]
            except Exception: txt = "<binary>"
            log.append((req.method, u, r.status, body, r.headers.get("content-type", ""), txt))
        ctx.on("response", lambda r: asyncio.ensure_future(on_resp(r)))
        await page.goto(BASE + "mainIndex.do?page=%2Fsmartreport%2Findex.html%23%2Freports%3FreportId%3D3123%26productId%3D4", wait_until="domcontentloaded", timeout=120000)
        await page.wait_for_selector("input[name=j_username]", timeout=120000)
        await page.fill("input[name=j_username]", os.environ["GF_USER"])
        await page.fill("input[name=j_password]", os.environ["GF_PASS"])
        await page.locator("button[type=submit]").first.click()
        await page.wait_for_timeout(25000)
        print("URL after login:", page.url)
        for f in page.frames: print("frame:", f.url[:200])
        await page.screenshot(path="output/gf_1.png", full_page=True)
        # dump visible text of the report frame
        for f in page.frames:
            if "smartreport" in f.url:
                t = await f.evaluate("document.body ? document.body.innerText : ''")
                print("SMARTREPORT TEXT:", re.sub(r"\s+", " ", t)[:2500])
                btns = await f.evaluate("[...document.querySelectorAll('button,a,[title]')].map(e=>(e.innerText||e.title||'').trim()).filter(Boolean).slice(0,120)")
                print("BUTTONS:", btns)
        await page.wait_for_timeout(5000)
        await b.close()
        print("\n==== NETWORK ====")
        for m, u, st, body, ct, txt in log:
            print(f"{m} {st} {u[:300]}\n   ct={ct}\n   body={body}\n   resp={re.sub(chr(10),' ',txt)[:600]}")

os.makedirs("output", exist_ok=True)
asyncio.run(main())
