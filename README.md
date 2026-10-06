<<<<<<< HEAD
# Marketing Performance Hub

A light-theme Streamlit dashboard and Excel data-entry tool built for the supplied workbook.

## What it does

- Reads `MasterData`, `Companies`, and `Currency selector` from `wolfable_master.xlsx`.
- Filters by Client, Platform, Region, Date and Currency.
- Date tiles: Yesterday, Last Week, Custom, Lifetime.
- Displays 8 KPI cards: Total Spend, Leads, CPL, CTR, Impressions, Clicks, CPC and Conversion Rate.
- **Never adds unlike currencies together.** Monetary values are calculated and displayed independently by currency.
- Wolfable permits INR, GBP and USD; other clients are restricted to their specified primary currency.
- Provides a **Data Manager** to add/update/delete campaign records.
- Every Data Manager save synchronizes both:
  1. the selected individual company sheet, and
  2. `MasterData`.
- Automatically calculates `Avg. CPC = Total Spend / Clicks` and `Cost Per Lead = Total Spend / Total Leads`.
- Creates a timestamped workbook backup before every write and retains the newest 20 backups.

## Files

- `app.py` — dashboard + data-entry UI.
- `data_store.py` — Excel reading/writing, validation, backup and synchronization logic.
- `wolfable_master.xlsx` — a copy of your supplied workbook used by default.
- `.streamlit/config.toml` — light theme.
- `requirements.txt` — Python dependencies.
- `run_dashboard.bat` / `run_dashboard.sh` — convenience launchers.

## Run it

Use Python 3.10+.

```bash
cd marketing_dashboard_app
python -m venv .venv
```

### Windows

```bat
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Or double-click `run_dashboard.bat` after installing the requirements.

### macOS / Linux

```bash
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Then open the local address Streamlit prints, normally `http://localhost:8501`.

## Use a workbook stored elsewhere

Set `DASHBOARD_WORKBOOK` to the full path before starting the app.

Windows PowerShell:

```powershell
$env:DASHBOARD_WORKBOOK="C:\Data\wolfable_master.xlsx"
streamlit run app.py
```

macOS / Linux:

```bash
export DASHBOARD_WORKBOOK="/path/to/wolfable_master.xlsx"
streamlit run app.py
```

## Currency rules

| Company | Allowed currency |
|---|---|
| Wolfable | INR, GBP, USD |
| Wellbeing | USD |
| Keecare | GBP |
| Cricket Centurion | USD |
| M360 | USD |
| Elite Valuation | INR |
| Musk | INR |

For example, if Wolfable has `£20` and `₹30`, the dashboard shows **£20 GBP · ₹30 INR**. It never shows `50`.

## Important workbook behavior

The individual client sheets do not contain a `Report Date` column. The Data Manager therefore stores `Report Date` in `MasterData`, while the remaining campaign fields are synchronized into the individual company sheet.

The date filter uses `MasterData[Report Date]`. If you later switch to true daily-grain advertising data, the same dashboard structure can be retained, but each daily row should have its own date.

## Deployment note

If this app is deployed for multiple people, add authentication and store the workbook on a shared filesystem or database. Do not allow several independent app servers to write separate local copies of the workbook. The included file lock protects concurrent writes on a single shared filesystem.
=======
# marketing-performance-dashboard
>>>>>>> 703a6092301365dd6ebf3281123e464f5b26832f
