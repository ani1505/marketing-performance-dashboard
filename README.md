# Marketing Performance Dashboard

## Live Dashboard

The deployed Streamlit application is available here:

**[Open the Marketing Performance Dashboard](https://marketing-performance-dashboard-bwholv4ywfozrqsssthdp7.streamlit.app/)**

---

A Streamlit-based marketing analytics dashboard for managing and analyzing campaign performance across multiple clients, platforms, regions, currencies, and reporting periods.

The project combines:

- an interactive marketing performance dashboard
- a company-aware currency filtering system
- an Excel-backed data store
- a Data Manager for adding, updating, and deleting campaign records
- automatic synchronization between `MasterData` and individual company sheets
- Streamlit Community Cloud deployment support

---

## Overview

This dashboard was built to provide a clean, Power BI-inspired view of campaign performance while keeping the underlying data easy to maintain.

The application supports seven companies:

| Company | Primary / Allowed Currency |
|---|---|
| Wolfable | INR, GBP, USD |
| Wellbeing | USD |
| Keecare | GBP |
| Cricket Centurion | USD |
| M360 | USD |
| Elite Valuation | INR |
| Musk | INR |

### Important currency rule

The dashboard **does not perform currency conversion**.

For companies such as Wolfable that may have transactions in multiple currencies, values are always kept separate.

For example:

```text
£20 GBP + ₹30 INR
```

is displayed as:

```text
£20 GBP | ₹30 INR
```

and is never incorrectly shown as `50`.

---

## Features

### Dashboard

The main dashboard includes:

- Company selector
- Region selector
- Platform selector
- Currency buttons
- Date filters:
  - Yesterday
  - Last Week
  - Lifetime
  - Custom
- KPI cards for:
  - Total Spend
  - Leads
  - Cost per Lead
  - CTR
  - Impressions
  - Clicks
  - CPC
  - Conversion Rate

### Analysis visuals

The dashboard includes performance-focused visuals such as:

- Leads by Platform
- Total Spend by Platform
- Leads by Date
- Campaign performance table
- Leads by Campaign
- Campaign-level performance comparisons

The design follows a light Power BI-inspired layout with:

- light grey background
- white KPI cards
- orange chart accents
- grey secondary chart elements
- responsive buttons and hover states
- dynamic company logos

---

## Currency Filtering

Currency buttons are displayed according to the selected company.

### Example: Wolfable

```text
Currency
[ INR ] [ GBP ]
```

By default, all available currencies for the company are included.

Selecting a currency filters the complete dashboard to that currency.

Clicking the selected currency again returns the dashboard to the all-currency view.

### Example behavior

```text
Default
→ INR + GBP data

Click INR
→ INR only

Click GBP
→ GBP only

Clear selection
→ INR + GBP data
```

No foreign-exchange conversion is performed.

---

## Data Manager

The application includes a Data Manager for maintaining campaign records directly from the Streamlit interface.

Supported actions:

- Add new campaign records
- Update existing records
- Delete records
- Download the latest workbook

When a record is added or updated:

1. The relevant company sheet is updated.
2. `MasterData` is updated.
3. CPC is recalculated automatically.
4. Cost per Lead is recalculated automatically.
5. A timestamped backup is created.
6. A success message is displayed.
7. The form is reset so another operation can be performed immediately.

When a record is deleted, it is removed from both the company sheet and `MasterData`.

---

## Excel Workbook Structure

The application uses:

```text
wolfable_master.xlsx
```

The workbook contains:

- `MasterData`
- `Companies`
- `Currency selector`
- individual company sheets

### MasterData

`MasterData` acts as the main source used by the dashboard.

Typical fields include:

```text
Company ID
Company
Platform
S.No
Campaign Name
Campaign Status
Region
Report Date
Start Date
End Date
Currency
Daily Budget
Impressions
Clicks
Avg. CPC
Total Leads
Cost Per Lead
Total Spend
Notes
```

### Companies

The `Companies` table is used for:

- company selection
- company IDs
- primary/native currency
- company logo URLs
- company metadata

### Currency selector

The currency table stores currency codes and symbols used in the dashboard.

---

## Project Structure

```text
marketing-performance-dashboard/
│
├── app.py
├── data_store.py
├── requirements.txt
├── wolfable_master.xlsx
├── README.md
│
├── .streamlit/
│   └── config.toml
│
└── backups/
```

### Main files

`app.py`

Contains the Streamlit interface, dashboard layout, filters, KPI calculations, charts, and Data Manager UI.

`data_store.py`

Contains Excel read/write logic, company currency rules, record validation, backups, and synchronization logic.

`wolfable_master.xlsx`

Main workbook used by the application.

`requirements.txt`

Python dependencies required to run the project.

---

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/ani1505/marketing-performance-dashboard.git
cd marketing-performance-dashboard
```

### 2. Create a virtual environment

Windows:

```powershell
python -m venv .venv
```

### 3. Install dependencies

You can run Python directly from the virtual environment without activating it:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### 4. Start the application

```powershell
.\.venv\Scripts\python.exe -m streamlit run app.py
```

The application will normally open at:

```text
http://localhost:8501
```

---

## PowerShell Execution Policy

If PowerShell blocks virtual environment activation with an execution-policy error, activation is not required.

Instead of:

```powershell
.venv\Scripts\activate
```

use:

```powershell
.\.venv\Scripts\python.exe -m streamlit run app.py
```

This avoids changing PowerShell security settings.

---

## Deploying to Streamlit Community Cloud

The project can be deployed directly from GitHub.

### 1. Push the project to GitHub

Make sure the repository contains:

```text
app.py
data_store.py
requirements.txt
wolfable_master.xlsx
.streamlit/config.toml
```

### 2. Open Streamlit Community Cloud

Go to:

```text
https://share.streamlit.io
```

### 3. Create a new app

Use:

```text
Repository:
ani1505/marketing-performance-dashboard

Branch:
main

Main file path:
app.py
```

Then click **Deploy**.

Streamlit will install the dependencies from `requirements.txt` automatically.

---

## Updating the Dashboard

After changing `app.py` locally:

```powershell
git add app.py
git commit -m "Update dashboard"
git pull --rebase origin main
git push origin main
```

Streamlit Community Cloud should automatically redeploy after the GitHub push.

---

## Updating the Workbook

After updating `wolfable_master.xlsx` locally:

```powershell
git add wolfable_master.xlsx
git commit -m "Update master data"
git pull --rebase origin main
git push origin main
```

The deployed Streamlit app will then pick up the updated workbook during redeployment.

---

## Recommended `.gitignore`

Use:

```gitignore
.venv/
__pycache__/
*.pyc
backups/
.DS_Store
```

Keep `wolfable_master.xlsx` tracked if the deployed dashboard currently depends on the workbook stored in GitHub.

---

## Important Deployment Note

Streamlit Community Cloud should not be treated as permanent writable file storage.

The Data Manager can write to the workbook while the application is running, but local file changes inside a cloud container may be lost after a restart or redeployment.

For a production system with frequent multi-user updates, consider moving the writable data layer to:

- PostgreSQL
- Google Sheets
- Supabase
- another persistent database

The existing Excel workbook can still be retained as an export or reporting format.

---

## Running the Data Manager Safely

Before using the Data Manager:

- keep a backup copy of the workbook
- avoid editing the Excel file manually while the Streamlit app is writing to it
- verify the selected company and currency before saving
- use the generated backups if a record needs to be restored

The application creates timestamped backups before write operations.

---

## Technology Stack

- Python
- Streamlit
- Pandas
- Plotly
- OpenPyXL
- FileLock
- Excel
- GitHub
- Streamlit Community Cloud

---

## Repository

**GitHub Repository**

https://github.com/ani1505/marketing-performance-dashboard

**Live Streamlit Dashboard**

https://marketing-performance-dashboard-bwholv4ywfozrqsssthdp7.streamlit.app/

---

## Future Improvements

Potential next steps include:

- move data storage from Excel to PostgreSQL
- add user authentication
- add role-based access for dashboard and Data Manager
- add campaign status analytics
- add month-over-month comparison KPIs
- add downloadable PDF reports
- add automated data imports from ad platforms
- add audit history for record changes
- add scheduled data refreshes

---

## License

This project is currently intended for internal/private use unless a separate license is added to the repository.
