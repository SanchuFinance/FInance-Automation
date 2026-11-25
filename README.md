# Finance Automation – Sanchu Animal Hospital

## VetBuddy – Detailed Goods & Services usage by Patients (MTD)
`vetbuddy_goods_services_mtd.py` downloads **Settings › Business Intelligence › Detailed Goods & Services usage by Patients** from VetBuddy for every unit (Adyar, Anna Nagar, Velachery, Neelankarai, Valasaravakkam, Coimbatore, Madurai), for **1st of the month → yesterday**, and builds one Excel workbook:

| Sheet | Contents |
|---|---|
| Summary | MTD totals per unit: line items, invoices, Total Sales, Discount, Net Sales, Tax, Final Sales, returns, Net (Final Sales − Return Total) |
| Daywise | Final Sales − Return Total by date × unit (for the VETBUDDY column of the POS report) |
| One sheet per unit | Full line-item data |

**Runs:** every day at 8:13 AM IST via GitHub Actions (`.github/workflows/vetbuddy_goods_services.yml`), or manually from *Actions › VetBuddy Goods & Services MTD › Run workflow* (optional custom dates).
**POS report:** each run also fills the **VETBUDDY** column of the month's unit-wise POS report, *Finance Hub / POS Report / POS_Report_Unitwise_<Mon_YYYY>.xlsx* (created automatically on the 1st). Only that column is written – Gofrugal, Paytm, deposits, credit and remarks typed by the team are kept. VETBUDDY = Final Sales − Return Total (checked against the manual POS sheet: Valasaravakkam 01-11 = 45,763, 02-11 = 17,762). If someone has the file open in Excel when the job runs, SharePoint may refuse the update; it is filled again on the next run.

**Output:** the workbook is attached to each run (Actions › run › Artifacts) and, if the SharePoint secrets are set, uploaded to *Finance Hub / Goods & Services MTD / YYYY-MM/*.

### Setup
Repository › Settings › Secrets and variables › Actions › New repository secret:
- `VB_USER`, `VB_PASS` – VetBuddy login
- `SP_TENANT_ID`, `SP_CLIENT_ID`, `SP_CLIENT_SECRET` – Azure app (same as the A/R script; needs Sites.ReadWrite.All). Optional.

Local run: `pip install -r requirements.txt`, copy `.env.example` to `.env`, then `python vetbuddy_goods_services_mtd.py --no-upload`.
