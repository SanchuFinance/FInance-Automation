"""
Gofrugal HQ (Smart Report) downloader - Sanchu Animal Hospital Finance Automation.

Signs in to Gofrugal HQ with a headless browser (GF_USER / GF_PASS), opens each Smart Report
the same way HQ does, then pulls the rows through Smart Report's own data call
(smartreport/rfwDataProcess.view/reportdata) for a date range.

  3123  Sales Detail - Customer Itemwise    -> sales   (value column: Item Net Amt)
  3126  (sales returns)                      -> returns (value column auto-detected, see RETURN_VALUE_COLS)

GOFRUGAL figure per unit per day = Sales (3123) - Returns (3126).
"""
import asyncio, datetime as dt, json, os, re

HOST_DEFAULT = "https://sanchuhq.gofrugal.com"
SALES_REPORT, RETURN_REPORT = 3123, 3126
SALES_VALUE_COL = "Item Net Amt"
RETURN_VALUE_COLS = ["Item Net Amount", "Item Net Amt", "Return Net Amt", "Net Amt", "Net Amount", "Return Amount", "Total Amount", "Amount"]
DATE_COLS = ["Bill Dt.", "SR Dt.", "Return Dt", "Sales Return Date", "SR Date.", "Ret Dt.", "Bill Dt", "Date.", "Return Dt.", "Return Date", "Bill Date", "Date", "Tran Date", "SR Date"]
OUTLET_COL = "Outlet Name"

# Gofrugal outlet name -> POS unit. Outlets not listed are reported in the log and left out.
OUTLET_TO_UNIT = {
    "Sanchu Adyar": "Adyar",
    "Sanchu Anna Nagar": "Anna Nagar", "Sanchu Anna Nagar 2": "Anna Nagar", "Sanchu Anna Nagar 3": "Anna Nagar",
    "Sanchu Anna Nagar 4": "Anna Nagar", "Sanchu Anna Nagar 5": "Anna Nagar",
    "Sanchu Velachery": "Velachery",
    "Sanchu Neelangarai": "Neelankarai",
    "Sanchu Valasaravakkam": "Valasaravakkam",
    "Sanchu Coimbatore": "Coimbatore", "Sanchu Coimbatore 2": "Coimbatore", "Sanchu Coimbatore 3": "Coimbatore",
    "Sanchu Madurai": "Madurai",
}


def _log(*a):
    import time
    print(time.strftime("%H:%M:%S"), *a, flush=True)


def _num(v):
    try:
        return float(str(v).replace(",", "").strip() or 0)
    except ValueError:
        return 0.0


def _date(v):
    s = str(v).strip()[:10]
    for f in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return dt.datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def _host():
    m = re.search(r"https?://[A-Za-z0-9.\-]+", os.environ.get("GF_URL") or "")
    return (m.group(0).replace("http://", "https://") if m else HOST_DEFAULT)


def _body(report_id, name, start, length):
    return {"groupby": None, "groupByLevel": None, "reportId": report_id, "reportname": name, "productId": 4,
            "displayLength": length, "displayStart": start, "advFilters": [], "advSumFilters": [], "filters": [],
            "sumFilters": [], "sort": [], "needColumns": False, "isProcedureReport": False, "invtype": 1,
            "dipsQry": "", "currencyPattern": "##,##,##,##0.00", "currencyDecimal": "2", "dateFormat": "DD-MM-YYYY",
            "numericalDecimal": "2", "qtyDecimal": "0", "colDisplay": "", "colExport": "", "allFilterValue": "",
            "isSubGroup": False, "defaultGroupByEnabled": False, "isMultipleUse": 1, "timeFormat": "12",
            "isGrouped": False, "entireGroupBy": "", "useSessionCache": True, "isFromSetting": False,
            "isGridReload": True, "isProcedure": True}


async def _fetch_reports(report_ids, d_from, d_to, page_size=2000):
    from playwright.async_api import async_playwright
    host = _host()
    base = host + "/RayMedi_HQ/"
    out = {}
    async with async_playwright() as p:
        b = await p.chromium.launch()
        ctx = await b.new_context(viewport={"width": 1500, "height": 900})
        page = await ctx.new_page()
        await page.goto(base + "mainIndex.do", wait_until="domcontentloaded", timeout=120000)
        await page.wait_for_selector("input[name=j_username]", timeout=120000)
        await page.locator("input[name=j_username]:visible").first.fill(os.environ["GF_USER"])
        pw = page.locator("input[name=j_password]:visible").first
        await pw.fill(os.environ["GF_PASS"])
        await pw.press("Enter")
        await page.wait_for_timeout(12000)
        if await page.locator("input[name=j_password]:visible").count():
            raise RuntimeError("Gofrugal login failed - check GF_USER / GF_PASS")
        _log("Logged in to Gofrugal")
        for rid in report_ids:
            await page.goto(f"{base}mainIndex.do?page=%2Fsmartreport%2Findex.html%23%2Freports%3FreportId%3D{rid}%26productId%3D4",
                            wait_until="domcontentloaded", timeout=120000)
            fr = None
            for _ in range(40):
                await page.wait_for_timeout(1500)
                fr = next((f for f in page.frames if "smartreport/index.html" in f.url), None)
                if fr:
                    break
            if not fr:
                raise RuntimeError(f"Smart Report {rid} did not open")
            await page.wait_for_timeout(6000)
            meta = await (await ctx.request.get(f"{host}/smartreport/meta/?reportId={rid}&productId=4&isShortCut=false&referrerProduct=&companyCode=&isHQ=false")).json()
            name = meta.get("reportName", "")
            qs = (f"reportId={rid}&productId=4&HQ_DEFA_ROLE_ID=20007&popup=false&fromdate={d_from:%Y-%m-%d}"
                  f"&todate={d_to:%Y-%m-%d}&retail=-1&areaCode=-1&referrerProduct=&companyCode=&isHQ=false")
            summ = await (await ctx.request.post(f"{host}/smartreport/rfwDataProcess.view/summarydata?{qs}",
                                                 data=json.dumps(_body(rid, name, 0, 50)),
                                                 headers={"Content-Type": "application/json"})).json()
            total = summ.get("nettotal", {}) or {}
            n = int(_num(total.get("rowCount", 0)))
            rows, start = [], 0
            while start < max(n, 1):
                r = await ctx.request.post(f"{host}/smartreport/rfwDataProcess.view/reportdata?{qs}",
                                           data=json.dumps(_body(rid, name, start, page_size)),
                                           headers={"Content-Type": "application/json"}, timeout=300000)
                chunk = (await r.json()).get("data", []) or []
                rows += chunk
                if len(chunk) < page_size:
                    break
                start += page_size
            _log(f"Gofrugal {rid} '{name}': {len(rows)} rows (report says {n}) | totals {total}")
            out[rid] = {"name": name, "rows": rows, "summary": total}
        await b.close()
    return out


def _pick(cols, prefs):
    for c in prefs:
        if c in cols:
            return c
    return None


def daywise(d_from, d_to):
    """Returns ({(date, unit): gofrugal_net}, details) where net = sales(3123) - returns(3126)."""
    data = asyncio.run(_fetch_reports([SALES_REPORT, RETURN_REPORT], d_from, d_to))
    result, details, unmapped = {}, {}, {}
    for rid, sign in ((SALES_REPORT, 1), (RETURN_REPORT, -1)):
        rows = data[rid]["rows"]
        cols = list(rows[0].keys()) if rows else []
        vcol = SALES_VALUE_COL if rid == SALES_REPORT else _pick(cols, RETURN_VALUE_COLS)
        dcol = _pick(cols, DATE_COLS)
        if rows and (not vcol or not dcol or OUTLET_COL not in cols):
            raise RuntimeError(f"Gofrugal {rid}: columns not recognised. Columns: {cols}")
        _log(f"Gofrugal {rid}: value column '{vcol}', date column '{dcol}'")
        for x in rows:
            outlet = str(x.get(OUTLET_COL, "")).strip()
            unit = OUTLET_TO_UNIT.get(outlet)
            v = _num(x.get(vcol))
            if not unit:
                unmapped[outlet] = unmapped.get(outlet, 0) + sign * v
                continue
            day = _date(x.get(dcol))
            if day:
                result[(day, unit)] = result.get((day, unit), 0) + sign * v
                k = (unit, "sales" if sign > 0 else "returns")
                details[k] = details.get(k, 0) + v
    for o, v in sorted(unmapped.items()):
        _log(f"Gofrugal outlet not in POS units (left out): {o}  {v:,.2f}")
    for (u, kind), v in sorted(details.items()):
        _log(f"Gofrugal {u:15s} {kind:8s} {v:,.2f}")
    return result, data


if __name__ == "__main__":
    import sys
    y = dt.date.today() - dt.timedelta(days=1)
    a = dt.date.fromisoformat(sys.argv[1]) if len(sys.argv) > 1 else y
    b = dt.date.fromisoformat(sys.argv[2]) if len(sys.argv) > 2 else y
    res, raw = daywise(a, b)
    for rid in raw:
        rows = raw[rid]["rows"]
        print(f"\n{rid} {raw[rid]['name']} columns:", list(rows[0].keys()) if rows else "no rows")
    for k in sorted(res):
        print(k[0], k[1], round(res[k], 2))
