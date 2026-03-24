# NDR Launch — Contact Filtering App

A web app that filters BD Advanced investor contacts by CDF criteria to generate a targeted contact list for an NDR or roadshow.

---

## What it does

Upload your BD Advanced exports and optionally a company background document. The app:

1. **Infers CDF criteria** from your company documents using AI, or lets you enter them manually
2. **Filters contacts** across four dimensions: Industry Focus, Investment Style, Market Cap, Geography
3. **Enriches results** with ownership data and meeting history from Activities.xlsx
4. **Routes contacts** to tabs by Investment Center, City, or State — or a single Virtual sheet
5. **Splits output** into structured sheets and downloads a formatted Excel file ready to use

---

## Input files

| File | Required | Header Row | Description |
| --- | --- | --- | --- |
| `Contacts w CDFs.xlsx` | ✅ | Row 3 | Main contacts export from BD Advanced |
| `Ownership.xlsx` | Optional | Row 5 | Adds `Shares` column |
| `Fund-Level Ownership.xlsx` | Optional | Row 5 | Adds four fund-level ownership columns |
| `Activities.xlsx` | Optional | Row 2 | Meeting history; drives enrichment, activity-only contacts, and meeting exclusion |
| `Junior Mining.xlsx` (or any supplemental list) | Optional | Row 3 | Supplemental contacts added to output, bypassing CDF filter |
| Company document (PDF or text) | Optional | — | 10-K, investor deck, etc. — triggers AI CDF recommendations |

---

## Output sheets

Sheets are written in this order. Empty sheets are omitted.

| Sheet | Contents |
| --- | --- |
| `Summary` | Always first — company name, tickers, CDF criteria, routing mode, exclusion settings, and per-sheet contact counts |
| `[Investment Center]` or `Contacts` | Filtered contacts routed by investment center, or all contacts if Virtual |
| `Virtual` | Contacts not matching any selected investment center (only when routing is active) |
| `HFs` | Hedge funds separated out based on HF treatment setting |
| `DNC` | Do Not Contact |
| `Check` | Check before calling |
| `Quant` | Quantitative funds |
| `Activist` | Contacts where Activist = Often |
| `Excluded` | Contacts excluded by shareholder threshold or meeting history rule |
| `Too Small` | Contacts below the EAUM minimum threshold (if set) |

Sheet names containing `/` are sanitized (replaced with `-`) and truncated to 31 characters to comply with Excel requirements.

---

## Workflow options

### Company name and tickers

Enter the company name used for the output filename. If an Activities file is uploaded, tickers are detected automatically and split into two groups:

**Subject company** — one or more tickers identifying the company this NDR is for. All six meeting history columns are computed using only rows matching these tickers. Meeting exclusion logic applies only to subject company meetings. Multiple subject tickers are pooled together — this is intended for dual-listed companies (same company, different ticker symbols). Using multiple genuinely different companies as subject tickers will produce misleading meeting counts.

**Other companies** — contacts who appear in activities for these tickers but are not already in the output are appended with `Source = Other: TICK1, TICK2` listing which tickers they were met under. Their meeting columns are left blank and meeting exclusion does not apply to them.

### CDF criteria

Select values from four dimensions. Leaving a dimension blank skips it (treats all values as neutral):

- **Industry Focus** — grouped by sector with search. `*Generalist` always matches any company.
- **Investment Style** — single-select or multi-select
- **Market Cap** — Micro, Small, Mid, Large, Mega
- **Geography** — includes global and regional values

Optionally upload company documents (10-K, investor deck) and click **Analyze documents** to have Claude pre-fill recommended CDF criteria with reasoning.

### Hedge fund treatment

* **Move to HFs sheet** — contacts where `Type` = Hedge Fund go to the HFs sheet
* **Include in main results** — all HFs remain in main results
* **Low-turnover only in main** — HFs with T/O > 100% go to HFs sheet; low-turnover HFs stay in main

### EAUM minimum

Optional dollar threshold in $mm. Contacts with a non-blank EAUM below this value are moved to the `Too Small` sheet.

### NDR routing

Four modes:

* **Virtual** (default) — all contacts on a single `Contacts` sheet
* **Investment Center** — searchable checklist of all 44 investment centers; one tab per selection plus `Virtual` catch-all
* **City** — searchable checklist of cities ("City, ST"); each resolves to its investment center for routing
* **State** — searchable checklist of state codes; all investment centers with cities in selected states are included

Routing always operates at the investment center level regardless of which mode was used to select. The investment center → city → state mapping is stored in `city_map.json` and can be updated via the admin page.

### Shareholder exclusion

Optional filter based on the `% S/O` column from the Ownership file (falls back to column index 4 if not found by name). Raw values are divided by 100. Options:

* **Include all** (default) — no exclusion
* **Exclude all shareholders** — any contact with `% S/O` > 0 is moved to Excluded
* **Greater than 0.01%** / **0.02%** / **0.03%** / **0.4%** / **0.5%** — contacts exceeding the threshold are moved to Excluded with reason "Exceeds Shareholder Limit"

### Meeting history exclusion

Applies only to subject company contacts:

* **Include all** — no exclusion
* **Exclude last 12 months** — contacts who met with the subject company in the last 12 months → `Excluded`
* **Exclude last 24 months** — contacts who met with the subject company in the last 24 months → `Excluded`
* **Exclude all prior meetings** — any contact with any prior meeting with the subject company → `Excluded`

---

## Filtering logic

Each contact is evaluated across four CDF dimensions. Each returns **match**, **exclude**, or **neutral**:

| Outcome | Condition |
| --- | --- |
| `neutral` | Dimension was skipped (nothing selected) OR contact's field is blank |
| `match` | Field is populated and at least one value overlaps the target set |
| `exclude` | Field is populated but no values match |

**A contact is included if:** at least one dimension is `match` AND no dimension is `exclude`.

**Industry special rule:** `*Generalist` in a contact's Industry field always returns `match` regardless of the target criteria.

### Split order (applied sequentially — each split only operates on remaining main contacts)

1. **Too Small** — EAUM non-null and below threshold
2. **HFs** — based on HF treatment setting; T/O threshold is > 100% (stored as > 1.0 after ÷100 normalization)
3. **DNC** — `CDF (Contact): Do Not Call` or `CDF (Firm): Do Not Call` is non-blank
4. **Check** — `CDF (Firm): Check before calling` = Yes
5. **Quant** — `CDF (Contact): Is Quant?` = Yes
6. **Activist** — `Activist` = Often
7. **Excluded (shareholders)** — shareholder exclusion based on `% S/O` threshold
8. **Excluded (meetings)** — meeting history exclusion (subject company only)

### Contact sources

| Source | Meaning |
| --- | --- |
| `CDF Match` | Passed the CDF filter from the contacts file |
| `Meeting History` | In Activities for subject company but not in contacts file |
| `Other: TICK1, TICK2` | In Activities for other-company tickers, not already in output; lists specific tickers |
| `Mining List` | From the supplemental contacts file |

---

## Output column reference

**Renamed columns** (BD Advanced original → output name):

| Original | Output |
| --- | --- |
| `Account Equity Assets Under Management (USD, mm)` | `EAUM ($mm)` |
| `Account Reported Total Assets (USD, mm)` | `AUM ($mm)` |
| `CDF (Contact): Geography` | `Geo` |
| `CDF (Contact): Industry Focus` | `Industry` |
| `CDF (Contact): Investment Style` | `Style` |
| `CDF (Contact): Market Cap.` | `Mkt. Cap` |
| `CDF (Firm): Coverage` | `Coverage` |
| `Last Meeting` | `Last Mtg. w/ Any Co` |
| `Primary Institution Type` | `Type` |
| `Contact Investment Center` | `Investment Ctr` |
| `Notes` + `Contact Notes` | `CRM Notes` (merged with ` \| ` separator) |

**T/O %** — raw turnover values in the BD Advanced file are divided by 100 on load (e.g. `50` → `0.50`). Formatted as a percentage in output.

**Meeting history columns** (added from Activities, subject company only):

| Column | Description |
| --- | --- |
| `Last Mtg btwn Contact & Co` | Most recent meeting date between this contact and the subject company |
| `Last Mtg btwn firm & Co` | Most recent meeting date between any contact at this institution and the subject company |
| `L12M` | Count of meetings in the last 12 months (blank if zero) |
| `Total` | Total meeting count (blank if zero) |
| `3rd Party` | Count of meetings where Topic = 3rd Party (blank if zero) |
| `Rose & Co` | Count of meetings where Topic is blank or = *Rose & Company (blank if zero) |

**Placeholder columns** added to all output sheets for manual entry: `Out1`, `Out2`, `Status`, `As of`, `Last Mtg. w/ Any Co`.

**Excel formatting:** navy (`#1F3864`) header row, alternating row shading, freeze panes at A2, auto-filter on all columns, column auto-width (min 8, max 42 characters). `CRM Notes` column is read-only (sheet protection enabled). `Industry`, `Geo`, `Style`, `Mkt. Cap` use shrink-to-fit at fixed width.

---

## Tech stack

* **Backend:** Python / Flask (`app.py`)
* **Filtering and enrichment:** pandas (`modules/filter.py`)
* **Excel output:** openpyxl (`modules/excel_output.py`)
* **AI analysis:** Anthropic Claude `claude-haiku-4-5` (`modules/ai_analysis.py`)
* **SharePoint upload:** optional (`modules/sharepoint.py`)
* **Frontend:** Vanilla JS (`static/app.js`), city/IC data injected as `window.CITY_MAP`
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

**3. Set environment variables**

| Variable | Value |
| --- | --- |
| `ANTHROPIC_API_KEY` | Your key from console.anthropic.com |
| `SECRET_KEY` | Click Generate |
| `ADMIN_PASSWORD` | Password for /admin page |

SharePoint upload is optional — leave blank to disable:

| Variable | Notes |
| --- | --- |
| `AZURE_TENANT_ID` | From Azure Portal |
| `AZURE_CLIENT_ID` | From Azure Portal |
| `AZURE_CLIENT_SECRET` | From Azure Portal |
| `SHAREPOINT_SITE_ID` | From IT admin or Azure |
| `SHAREPOINT_FOLDER` | Default: `/NDR Launch` |

**4. Deploy** — click Create Web Service. First deploy takes ~2 minutes.

**5. Upload city map** — go to `/admin` and upload an Excel file with columns `Investment Center`, `Nearby City`, `State` to populate location routing data. The current map has 829 rows across 44 investment centers.

---

## Admin page (`/admin`)

Password-protected. Provides:

* **Upload taxonomy** — update CDF dropdown values from an Excel file (columns: Filter Type, Value, Description)
* **Upload city map** — update Investment Center → City → State mapping from an Excel file (columns: Investment Center, Nearby City, State)
* **Meetings database** — preview and upload meeting history (requires `DATABASE_URL` environment variable)

---

## File structure

```
ndr-launch/
├── app.py                    # Flask routes, city map, taxonomy, file loading
├── city_map.json             # Investment center → city → state (829 rows, 44 ICs)
├── taxonomy.json             # CDF taxonomy values (auto-generated if missing)
├── requirements.txt
├── render.yaml               # Render deployment config
├── modules/
│   ├── filter.py             # CDF filtering, splits, activity enrichment, mining append
│   ├── activities.py         # Legacy module (superseded by filter.py — not imported)
│   ├── excel_output.py       # Excel generation, Summary sheet, formatting, sheet name sanitization
│   ├── ai_analysis.py        # Claude-powered CDF recommendations
│   ├── sharepoint.py         # Optional SharePoint upload
│   └── meetings.py           # Optional meetings database (requires DATABASE_URL)
├── static/
│   ├── app.js                # UI: CDF selection, ticker groups, routing mode switching
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
# Fill in ANTHROPIC_API_KEY and SECRET_KEY
pip install -r requirements.txt
python app.py
```

App runs at http://localhost:5000
