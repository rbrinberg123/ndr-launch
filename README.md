# NDR Launch — Contact Filtering App

A web app that filters BD Advanced investor contacts by CDF criteria to generate a targeted list for an NDR or roadshow.

---

## What it does

Upload your BD Advanced exports and optionally a company background document (10-K, investor deck). The app:

1. **Infers CDF criteria** from your company documents using AI, or lets you enter them manually
2. **Filters contacts** across four dimensions: Industry Focus, Investment Style, Market Cap, Geography
3. **Routes contacts** to city-based tabs or a single Virtual sheet
4. **Enriches results** with ownership data and meeting history from Activities.xlsx
5. **Splits output** into structured sheets: Contacts, HFs, DNC, Check, Quant, Activist, Excluded, Too Small
6. **Downloads** a formatted Excel file ready to use, with a Criteria summary tab as the first sheet

---

## Input files

| File | Required | Description |
| --- | --- | --- |
| `Contacts w CDFs.xlsx` | ✅ | Main contacts export from bdadvanced.ipreo.com (header row 3) |
| `Ownership.xlsx` | Optional | Adds `Shares` column (header row 5) |
| `Fund-Level Ownership.xlsx` | Optional | Adds four fund-level columns (header row 5) |
| `Activities.xlsx` | Optional | Adds six meeting history columns; used for meeting exclusion and activity-only contacts |
| Company document (PDF or text) | Optional | 10-K, investor deck, etc. — triggers AI CDF recommendations |

---

## Output sheets

Sheets are written in this order. Empty sheets are omitted.

| Sheet | Contents |
| --- | --- |
| `Criteria` | Always first — summarizes company name, tickers, CDF criteria, routing mode, and exclusion settings used for the run |
| `[City]` or `Contacts` | Filtered contacts routed by city, or all contacts if Virtual |
| `Virtual` | Contacts not matching any selected city (only when cities are chosen) |
| `HFs` | Hedge funds (if HF treatment = Separate or Low-turnover only) |
| `DNC` | Do Not Contact |
| `Check` | Contacts flagged for review (Check before calling) |
| `Quant` | Quantitative funds |
| `Activist` | Contacts where Activist = Often |
| `Excluded` | Contacts excluded by meeting history rule |
| `Too Small` | Contacts below the EAUM minimum threshold (if set) |

---

## Workflow options

### Hedge fund treatment

* **Include all** — keep all HFs in main results
* **Separate into HFs tab** — move all contacts where `Primary Institution Type` = Hedge Fund to the HFs sheet
* **Include low-turnover HFs only (T/O ≤ 100%)** — move high-turnover HFs (T/O > 100%) to HFs sheet; keep low-turnover in main results

### EAUM minimum

Optional threshold in $mm. Contacts with a non-blank EAUM below this value are moved to a `Too Small` sheet.

### City routing

* **Virtual** (default) — all contacts go on a single `Contacts` sheet
* **Select cities** — creates one tab per city plus a `Virtual` catch-all tab for unmatched contacts

Supported cities: New York, Boston, Chicago, Philadelphia, San Francisco, Los Angeles, Dallas, Houston, Minneapolis, South Florida, Denver, Atlanta, Nashville, Kansas City, Columbus, San Antonio, London, Paris, Amsterdam, Tokyo, Hong Kong, Toronto.

Also accepts shorthand input: NY/NYC, SF, LA, Philly, Miami, Florida.

### Meeting history exclusion

* **Include all** — keep all contacts
* **Exclude L12M** — move contacts met with the company in the last 12 months to `Excluded`
* **Exclude L24M** — move contacts met with the company in the last 24 months to `Excluded`
* **Exclude all meeting history** — move all contacts with any prior meeting to `Excluded`

---

## Filtering logic

Contacts are evaluated across four CDF dimensions. Each dimension independently returns **match**, **exclude**, or **neutral**:

| Outcome | Condition |
| --- | --- |
| `neutral` | Dimension was skipped (no values selected), OR the contact's field is blank |
| `match` | Field is populated and at least one value overlaps the target criteria |
| `exclude` | Field is populated but no values match |

**A contact is included if:** at least one dimension is a match AND no dimension is an exclude.

**Industry special rule:** contacts with `*Generalist` in their Industry field always count as a match, regardless of the target criteria.

### Split order (applied sequentially)

Splits are applied to the running main list only. Contacts already split off are not re-evaluated.

1. **Too Small** — EAUM below threshold (if set)
2. **HFs** — based on HF treatment setting
3. **DNC** — `CDF (Contact): Do Not Call` or `CDF (Firm): Do Not Call` is non-blank
4. **Check** — `CDF (Firm): Check before calling` = Yes
5. **Quant** — `CDF (Contact): Is Quant?` = Yes
6. **Activist** — `Activist` = Often
7. **Excluded** — meeting history exclusion (if set)

### Activity-only contacts

Contacts who appear in Activities.xlsx for the selected ticker but are not in the contacts file are automatically added to the output with `Source = Meeting History`. Their demographic and CDF data is populated from the most recent non-null value across their activity rows.

---

## Tech stack

* **Backend:** Python / Flask
* **Filtering logic:** pandas (`modules/filter.py`)
* **Activities enrichment:** pandas (`modules/activities.py`)
* **Excel output:** openpyxl — navy headers, alternating rows, freeze panes, auto-filter (`modules/excel_output.py`)
* **AI analysis:** Anthropic Claude (`claude-haiku-4-5`) via `modules/ai_analysis.py`
* **Hosting:** Render.com

---

## Deployment

### Prerequisites

* GitHub account
* Render.com account (free tier works)
* Anthropic API key from console.anthropic.com

### Steps

**1. Push to GitHub**

```
git clone https://github.com/YOUR_USER/ndr-launch.git
cd ndr-launch
git add .
git commit -m "Initial deploy"
git push
```

**2. Create Render Web Service**

* dashboard.render.com → New → Web Service
* Connect your GitHub repo
* Render auto-detects `render.yaml`
* Build command: `pip install -r requirements.txt`
* Start command: `gunicorn app:app --workers 2 --timeout 120`

**3. Set environment variables in Render dashboard**

| Variable | Value |
| --- | --- |
| `ANTHROPIC_API_KEY` | Your key from console.anthropic.com |
| `SECRET_KEY` | Click Generate |
| `ADMIN_PASSWORD` | Password for /admin page |

SharePoint upload is optional — leave these blank to disable:

| Variable | Notes |
| --- | --- |
| `AZURE_TENANT_ID` | From Azure Portal |
| `AZURE_CLIENT_ID` | From Azure Portal |
| `AZURE_CLIENT_SECRET` | From Azure Portal |
| `SHAREPOINT_SITE_ID` | From IT admin or Azure |
| `SHAREPOINT_FOLDER` | Default: `/NDR Launch` |

**4. Deploy** — click Create Web Service. First deploy takes ~2 minutes.

---

## File structure

```
ndr-launch/
├── app.py                    # Flask routes and city/IC mapping
├── requirements.txt
├── render.yaml               # Render deployment config
├── modules/
│   ├── filter.py             # Core CDF filtering and split logic
│   ├── activities.py         # Meeting history enrichment and activity-only contacts
│   ├── excel_output.py       # Formatted Excel generation and Criteria sheet
│   ├── ai_analysis.py        # Claude-powered CDF recommendations
│   └── sharepoint.py         # Optional SharePoint upload
├── static/
│   ├── app.js
│   └── style.css
└── templates/
    ├── base.html
    ├── index.html
    └── admin.html
```

---

## Local development

```
cp .env.example .env
# Fill in ANTHROPIC_API_KEY and SECRET_KEY in .env
pip install -r requirements.txt
python app.py
```

App runs at http://localhost:5000
