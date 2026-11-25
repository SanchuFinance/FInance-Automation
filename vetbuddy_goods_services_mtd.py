"""
VetBuddy - Detailed Goods & Services usage by Patients  (month start -> yesterday, all units)
=========================================================================================
Sanchu Animal Hospital - Finance Automation

VetBuddy path: Settings > Business Intelligence > Detailed Goods & Services usage by Patients
Report page  : https://sanchu.thevetbuddy.com/sanchu/sales_by_planitem.php

What it does
  1. Logs in to VetBuddy.
  2. For every unit, downloads the report (Excel) for 1st of the month -> yesterday (MTD, D-1).
     On the 1st of a month it pulls the whole previous month.
  3. Builds one workbook:
        Summary      - MTD totals per unit (Total Sales, Discount, Net Sales, Tax, Final Sales, Returns, Net)
        Daywise      - Final Sales net of returns, date x unit (feeds the VETBUDDY column of the POS report)
        <Unit> sheets- the full line-item data for each unit
  4. Optional: uploads it to SharePoint (Finance Hub) if the SP_* settings are present.

Usage
  python vetbuddy_goods_services_mtd.py                 normal run (MTD D-1)
  python vetbuddy_goods_services_mtd.py --from 2026-11-01 --to 2026-11-24
  python vetbuddy_goods_services_mtd.py --no-upload     build the file only

Settings (environment variables / GitHub secrets, or a .env file next to this script)
  VB_USER, VB_PASS                       VetBuddy login
  VB_BASE      (optional)                default https://sanchu.thevetbuddy.com
  SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET   Azure app (same as the A/R / sales scripts)
  SP_SITE_HOST (optional)                default sanchuhospital.sharepoint.com
  SP_SITE_PATH (optional)                default ""  (root site)
  SP_LIBRARY   (optional)                default Finance Hub
  GS_SP_FOLDER (optional)                default "Goods & Services MTD"
Exit code 0 = success, 1 = failure.
"""
import argparse, datetime as dt, io, os, re, sys, time, urllib.parse
from pathlib import Path

import requests, xlrd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
OUT = HERE / "output"

# VetBuddy clinic ids (from the report's clinic dropdown)
UNITS = {
    "Adyar": "5",
    "Anna Nagar": "6",
    "Velachery": "2",
    "Neelankarai": "1",
    "Valasaravakkam": "14",
    "Coimbatore": "4",
    "Madurai": "13",
}
NUM_COLS = ["Total Sales", "Discount", "Net Sales", "Tax", "Final Sales",
            "Total Return", "Net Return", "Return Tax", "Return Total"]


# ----------------------------------------------------------------- settings
def load_env():
    f = HERE / ".env"
    if f.exists():
        for line in f.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.split(" #")[0].strip())


def env(k, default=None, required=False):
    v = os.environ.get(k, default)
    if required and not v:
        sys.exit(f"Missing setting {k}")
    return v


def log(*a):
    print(time.strftime("%H:%M:%S"), *a, flush=True)


# ----------------------------------------------------------------- VetBuddy
class VetBuddy:
    def __init__(self, base, user, pwd):
        self.base = base.rstrip("/")
        self.s = requests.Session()
        self.s.headers["User-Agent"] = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                        "AppleWebKit/537.36 Chrome/124 Safari/537.36")
        self.user, self.pwd = user, pwd

    @staticmethod
    def _text(html):
        t = re.sub(r"(?is)<(script|style).*?</\1>", " ", html)
        t = re.sub(r"<[^>]+>", " ", t)
        return re.sub(r"\s+", " ", t).strip()

    def _report_page_ok(self):
        chk = self.s.get(self.base + "/sanchu/sales_by_planitem.php",
                         params={"txthidedate": "true"}, timeout=60,
                         headers={"Referer": self.base + "/business_menu.phtml"})
        return 'name="selclinic"' in chk.text

    def _post_login(self):
        # the login page calls this on load, then posts the form as multipart/form-data
        self.s.get(self.base + "/", timeout=60)
        self.s.get(self.base + "/vet_inc_session.html",
                   params={"calledfromlogin": "true", "viewport": "1920|1080"}, timeout=60)
        r = self.s.post(self.base + "/comm_whiteboard.html",
                        files={"userid": (None, self.user), "passwd": (None, self.pwd),
                               "login": (None, "login")},
                        headers={"Referer": self.base + "/", "Origin": self.base},
                        timeout=60)
        r.raise_for_status()
        self.s.get(self.base + "/vet_inc_session.html",
                   params={"calledfromlogin": "true", "viewport": "1920|1080"}, timeout=60)
        return r

    def login(self):
        r = self._post_login()
        if self._report_page_ok():
            log("Logged in to VetBuddy")
            return
        page = self._text(r.text)
        log("After sign-in VetBuddy showed:", r.url, "|", page[:400])
        # VetBuddy may refuse a second session for the same user ("logged in on another device")
        if re.search(r"another device|other device|already logged|logoutfromdevice", r.text, re.I):
            log("Account is signed in elsewhere - closing the other session and retrying")
            self.s.get(self.base + "/vetapp_start.html",
                       params={"logoutfromdevice": "true", "uid": ""}, timeout=60)
            self.s.cookies.set("seed", str(int(time.time() * 1000)))
            r = self._post_login()
            if self._report_page_ok():
                log("Logged in to VetBuddy")
                return
            log("After retry VetBuddy showed:", r.url, "|", self._text(r.text)[:400])
        raise RuntimeError("VetBuddy login failed (report page did not open). "
                           "See the message above; check VB_USER / VB_PASS.")

    def download(self, clinic_id, d_from, d_to):
        """Returns the .xls bytes of the report for one clinic and date range."""
        data = {
            "calledforbiller": "", "txthidedate": "true",
            "datefrom": d_from.strftime("%m/%d/%Y"), "dateTo": d_to.strftime("%m/%d/%Y"),
            "selclinic": clinic_id, "txtstockitem": "", "selstockid": "",
            "hidcheckforcheckbox": "", "selprovider": "", "selsalestype": "",
            "txtclient": "", "hidclient": "", "txtclient_hiddefaultvalue": "",
            "txtclient_hidnextcontrol": "", "chkgoods": "chkgoods", "chkservice": "service",
            "lstscheduledrug": "", "chkcomment": "All", "cmdxls": "Extract XLS Report",
        }
        for attempt in range(3):
            try:
                r = self.s.post(self.base + "/sanchu/sales_by_planitem.php", data=data, timeout=600)
                r.raise_for_status()
                m = re.search(r"doc/sanchu/sales_by_planitem_\d+\.xls", r.text)
                if not m:
                    raise RuntimeError("no Excel link in VetBuddy response")
                f = self.s.get(self.base + "/sanchu/" + m.group(0), timeout=600)
                f.raise_for_status()
                return f.content
            except Exception as e:
                log(f"  attempt {attempt + 1} failed: {e}")
                time.sleep(10 * (attempt + 1))
        raise RuntimeError(f"Could not download report for clinic {clinic_id}")


# ----------------------------------------------------------------- parsing
def to_num(v):
    if v in (None, ""):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").replace("%", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    s = str(v).strip()
    m = re.match(r"([A-Za-z]{3})[a-z]* (\d{1,2}), (\d{4})", s)      # "Nov 24, 2026 23:52"
    if m:
        return dt.datetime.strptime(" ".join(m.groups()), "%b %d %Y").date()
    for fmt in ("%d-%m-%Y", "%m/%d/%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(s[:10], fmt).date()
        except ValueError:
            pass
    return None


def parse_xls(content):
    """Find the header row (starts with 'Sr No.'), return (headers, rows as dicts)."""
    if content[:4] != b"\xd0\xcf\x11\xe0":
        raise RuntimeError("Downloaded file is not an Excel (.xls) file")
    book = xlrd.open_workbook(file_contents=content)
    sh = book.sheet_by_index(0)
    hdr_row = None
    for r in range(min(sh.nrows, 40)):
        vals = [str(c).strip() for c in sh.row_values(r)]
        if vals and vals[0].lower().startswith("sr no") and "Final Sales" in vals:
            hdr_row = r
            break
    if hdr_row is None:
        raise RuntimeError("Header row ('Sr No.' ... 'Final Sales') not found in the report")
    headers = [str(c).strip() for c in sh.row_values(hdr_row)]
    rows = []
    for r in range(hdr_row + 1, sh.nrows):
        vals = sh.row_values(r)
        first = str(vals[0]).strip()
        if not first or first.lower().startswith("total"):
            continue
        if not re.match(r"^\d+(\.0)?$", first):
            continue
        rec = dict(zip(headers, vals))
        if isinstance(rec.get("Date"), float) and rec["Date"] > 30000:
            rec["Date"] = xlrd.xldate_as_datetime(rec["Date"], book.datemode)
        rows.append(rec)
    return headers, rows


# ----------------------------------------------------------------- workbook
GREEN = PatternFill("solid", fgColor="274E13")
WH = Font(name="Arial", bold=True, color="FFFFFF", size=10)
BODY = Font(name="Arial", size=10)
BOLD = Font(name="Arial", size=10, bold=True)
thin = Side(style="thin", color="A6A6A6")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
INR = '#,##,##0.00;-#,##,##0.00;"-"'


def head(ws, row, values):
    for c, v in enumerate(values, 1):
        x = ws.cell(row, c, v)
        x.fill, x.font, x.border = GREEN, WH, BOX
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def build(data, d_from, d_to, path):
    wb = Workbook()
    s = wb.active
    s.title = "Summary"
    s["A1"] = "VetBuddy - Detailed Goods & Services usage by Patients"
    s["A1"].font = Font(name="Arial", bold=True, size=14)
    s["A2"] = f"Period: {d_from:%d-%m-%Y} to {d_to:%d-%m-%Y}   |   Generated {dt.datetime.now():%d-%m-%Y %I:%M %p}"
    s["A2"].font = Font(name="Arial", italic=True, size=9)
    cols = ["Unit", "Line items", "Invoices"] + NUM_COLS + ["Net (Final Sales - Return Total)"]
    head(s, 4, cols)
    r = 5
    for unit, (_, rows) in data.items():
        inv = {str(x.get("Invoice #", "")).strip() for x in rows}
        vals = [unit, len(rows), len(inv - {""})]
        vals += [sum(to_num(x.get(c)) for x in rows) for c in NUM_COLS]
        for c, v in enumerate(vals, 1):
            cell = s.cell(r, c, v)
            cell.font, cell.border = BODY, BOX
            if c > 3:
                cell.number_format = INR
        fs, rt = get_column_letter(3 + NUM_COLS.index("Final Sales") + 1), \
                 get_column_letter(3 + NUM_COLS.index("Return Total") + 1)
        x = s.cell(r, len(cols), f"={fs}{r}-{rt}{r}")
        x.number_format, x.font, x.border = INR, BOLD, BOX
        r += 1
    s.cell(r, 1, "TOTAL")
    for c in range(2, len(cols) + 1):
        L = get_column_letter(c)
        s.cell(r, c, f"=SUM({L}5:{L}{r - 1})").number_format = INR if c > 3 else "#,##0"
    for c in range(1, len(cols) + 1):
        s.cell(r, c).fill, s.cell(r, c).font, s.cell(r, c).border = GREEN, WH, BOX
    s.column_dimensions["A"].width = 16
    for c in range(2, len(cols) + 1):
        s.column_dimensions[get_column_letter(c)].width = 14
    s.row_dimensions[4].height = 45
    s.freeze_panes = "B5"

    # Daywise: Final Sales - Return Total, by date x unit
    d = wb.create_sheet("Daywise")
    d["A1"] = "Daywise VetBuddy sales (Final Sales - Return Total) - use for the VETBUDDY column of the POS report"
    d["A1"].font = Font(name="Arial", bold=True, size=11)
    units = list(data)
    head(d, 3, ["Date"] + units + ["TOTAL"])
    days = [d_from + dt.timedelta(n) for n in range((d_to - d_from).days + 1)]
    agg = {}
    for unit, (_, rows) in data.items():
        for x in rows:
            day = parse_date(x.get("Date"))
            if day:
                agg[(day, unit)] = agg.get((day, unit), 0) + to_num(x.get("Final Sales")) - to_num(x.get("Return Total"))
    for i, day in enumerate(days):
        rr = 4 + i
        d.cell(rr, 1, day).number_format = "DD-MM-YYYY"
        for j, u in enumerate(units, 2):
            d.cell(rr, j, round(agg.get((day, u), 0), 2)).number_format = INR
        d.cell(rr, len(units) + 2, f"=SUM(B{rr}:{get_column_letter(len(units) + 1)}{rr})").number_format = INR
        for c in range(1, len(units) + 3):
            d.cell(rr, c).font, d.cell(rr, c).border = BODY, BOX
    tr = 4 + len(days)
    d.cell(tr, 1, "TOTAL")
    for c in range(2, len(units) + 3):
        L = get_column_letter(c)
        d.cell(tr, c, f"=SUM({L}4:{L}{tr - 1})").number_format = INR
    for c in range(1, len(units) + 3):
        d.cell(tr, c).fill, d.cell(tr, c).font, d.cell(tr, c).border = GREEN, WH, BOX
    for c in range(1, len(units) + 3):
        d.column_dimensions[get_column_letter(c)].width = 15
    d.freeze_panes = "B4"

    # Raw data per unit
    for unit, (headers, rows) in data.items():
        w = wb.create_sheet(unit[:31])
        if not headers:
            w["A1"] = "No data for this period"
            continue
        head(w, 1, headers)
        for i, x in enumerate(rows, 2):
            for c, h in enumerate(headers, 1):
                v = x.get(h)
                if h in NUM_COLS or h in ("Clinic Cost", "MRP /Service Fee", "Qty", "Return Qty"):
                    v = to_num(v)
                elif isinstance(v, float) and v.is_integer():
                    v = int(v)
                w.cell(i, c, v).font = BODY
        w.freeze_panes = "B2"
        w.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows) + 1}"
        for c in range(1, len(headers) + 1):
            w.column_dimensions[get_column_letter(c)].width = 14
    wb.save(path)


# ----------------------------------------------------------------- SharePoint
def sp_upload(path, folder):
    tenant, cid, sec = env("SP_TENANT_ID"), env("SP_CLIENT_ID"), env("SP_CLIENT_SECRET")
    if not (tenant and cid and sec):
        log("SharePoint settings not set - skipping upload")
        return
    tok = requests.post(f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token", data={
        "client_id": cid, "client_secret": sec, "grant_type": "client_credentials",
        "scope": "https://graph.microsoft.com/.default"}, timeout=60).json()["access_token"]
    H = {"Authorization": f"Bearer {tok}"}
    g = "https://graph.microsoft.com/v1.0"
    host = env("SP_SITE_HOST", "sanchuhospital.sharepoint.com")
    sp = env("SP_SITE_PATH", "")
    site = requests.get(f"{g}/sites/{host}" + (f":/{sp.strip('/')}" if sp else ""), headers=H, timeout=60).json()
    lib = env("SP_LIBRARY", "Finance Hub")
    drives = requests.get(f"{g}/sites/{site['id']}/drives", headers=H, timeout=60).json()["value"]
    drive = next((x for x in drives if x["name"] == lib), None)
    if not drive:
        raise RuntimeError(f"SharePoint library '{lib}' not found")
    dest = urllib.parse.quote(f"{folder}/{path.name}")
    url = f"{g}/drives/{drive['id']}/root:/{dest}:/content"
    r = requests.put(url, headers=H, data=path.read_bytes(), timeout=300)
    r.raise_for_status()
    log("Uploaded to SharePoint:", r.json().get("webUrl"))


# ----------------------------------------------------------------- main
def period(args):
    today = dt.date.today()
    if args.date_from and args.date_to:
        return dt.date.fromisoformat(args.date_from), dt.date.fromisoformat(args.date_to)
    d_to = today - dt.timedelta(days=1)          # day minus one
    return d_to.replace(day=1), d_to             # on the 1st -> whole previous month


def main():
    load_env()
    ap = argparse.ArgumentParser()
    ap.add_argument("--from", dest="date_from")
    ap.add_argument("--to", dest="date_to")
    ap.add_argument("--no-upload", action="store_true")
    args = ap.parse_args()
    d_from, d_to = period(args)
    log(f"Period {d_from} -> {d_to}")

    vb = VetBuddy(env("VB_BASE", "https://sanchu.thevetbuddy.com"),
                  env("VB_USER", required=True), env("VB_PASS", required=True))
    vb.login()
    OUT.mkdir(exist_ok=True)
    data = {}
    for unit, cid in UNITS.items():
        log(f"Downloading {unit} (clinic {cid})")
        raw = vb.download(cid, d_from, d_to)
        (OUT / f"raw_{unit.replace(' ', '_')}.xls").write_bytes(raw)
        headers, rows = parse_xls(raw)
        log(f"  {len(rows)} line items")
        data[unit] = (headers, rows)

    path = OUT / f"VetBuddy_GoodsServices_MTD_{d_to:%Y-%m-%d}.xlsx"
    build(data, d_from, d_to, path)
    log("Saved", path)
    if not args.no_upload:
        sp_upload(path, f"{env('GS_SP_FOLDER', 'Goods & Services MTD')}/{d_to:%Y-%m}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log("FAILED:", e)
        sys.exit(1)
