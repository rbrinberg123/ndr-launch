# NDR Launch — Contact Filtering App

A web app that filters BD Advanced investor contacts by CDF criteria to generate a targeted list for an NDR or roadshow.

---

## What it does

Upload your BD Advanced exports and optionally a company background document (10-K, investor deck). The app:

1. **Infers CDF criteria** from your company documents using AI, or lets you enter them manually
2. **Filters contacts** across four dimensions: Industry Focus, Investment Style, Market Cap, Geography
3. **Routes contacts** to tabs by Investment Center, City, or State — or a single Virtual sheet
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
| `Activities.xlsx` | Optional | Adds meeting history columns; drives subject/other company contact enrichment |
| Company document (PDF or text) | Optional | 10-K, investor deck, etc. — triggers AI CDF recommendations |

---

## Output sheets

Sheets are written in this order. Empty sheets are omitted.

| Sheet | Contents |
| --- | --- |
| `Criteria` | Always first — summarizes company name, tickers, CDF criteria, routing mode, and exclusion settings |
| `[Investment Center]` or `Contacts` | Filtered contacts routed by investment center, or all contacts if Virtual |
| `Virtual` | Contacts not matching any selected investment center (only when routing is active) |
| `HFs` | Hedge funds (if HF treatment = Separate or Low-turnover only) |
| `DNC` | Do Not Contact |
| `Check` | Contacts flagged for review (Check before calling) |
| `Quant` | Quantitative funds |
| `Activist` | Contacts where Activist = Often |
| `Excluded` | Contacts excluded by meeting history rule |
| `Too Small` | Contacts below the EAUM minimum threshold (if set) |

---

## Workflow options

### Ticker selection (Activities file)

When an Activities file is uploaded, tickers are split into two groups:

* **Subject company** — drives all meeting history column calculations (`Last Mtg btwn Contact & Co`, `L12M`, `Total`, etc.) and meeting exclusion logic. Activity-only contacts for the subject company are added to the output and subject to all split rules.
* **Other companies** — contacts who appear in activities for these tickers but are not already in the output are appended with `Source = Meeting History (Other)`. Meeting history columns are left blank for these contacts and meeting exclusion does not apply to them.

### Hedge fund treatment

* **Include all** — keep all HFs in main results
* **Separate into HFs tab** — move all contacts where `Type` = Hedge Fund to the HFs sheet
* **Include low-turnover HFs only (T/O ≤ 100%)** — move high-turnover HFs to HFs sheet; keep low-turnover in main results

### EAUM minimum

Optional threshold in $mm. Contacts with a non-blank EAUM below this value are moved to a `Too Small` sheet.

### Location routing

Four modes, selected via radio buttons:

* **Virtual** (default) — all contacts go on a single `Contacts` sheet
* **Investment Center** — select from a searchable list of all 44 investment centers; creates one tab per selection plus a `Virtual` catch-all
* **City** — select individual cities; each city resolves to its investment center for tab routing
* **State** — select states (2-letter codes); all investment centers with cities in the selected states are included

In all non-Virtual modes, routing is always done at the investment center level. Tabs are named after the investment center (with `/` replaced by `-` to comply with Excel sheet name rules).

The investment center → city → state mapping is stored in `city_map.json` and can be updated via the admin page.

### Meeting history exclusion

* **Include all** — keep all contacts
* **Exclude L12M** — move contacts met with the subject company in the last 12 months to `Excluded`
* **Exclude L24M** — move contacts met with the subject company in the last 24 months to `Excluded`
* **Exclude all meeting history** — move all contacts with any prior meeting with the subject company to `Excluded`

Only applies to subject company contacts. Other-company contacts are never excluded by meeting history.

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

### Split order (applied sequentially to main results only)

1. **Too Small** — EAUM below threshold (if set)
2. **HFs** — based on HF treatment setting
3. **DNC** — `CDF (Contact): Do Not Call` or `CDF (Firm): Do Not Call` is non-blank
4. **Check** — `CDF (Firm): Check before calling` = Yes
5. **Quant** — `CDF (Contact): Is Quant?` = Yes
6. **Activist** — `Activist` = Often
7. **Excluded** — meeting history exclusion (subject company only)

### Contact sources

| Source value | Meaning |
| --- | --- |
| `CDF Match` | Contact passed the CDF filter from the contacts file |
| `Meeting History` | Contact found in Activities for the subject company but not in the contacts file |
| `Meeting History (Other)` | Contact found in Activities for an other-company ticker, not already in output |
| `Mining List` | Contact from an optional supplemental mining contacts file |

---

## Output column reference

Key renamed columns in the output (original BD Advanced names → output names):

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

Meeting history columns added from Activities:

| Column | Description |
| --- | --- |
| `Last Mtg btwn Contact & Co` | Most recent meeting date between this contact and the subject company |
| `Last Mtg btwn firm & Co` | Most recent meeting date between any contact at this institution and the subject company |
| `L12M` | Count of meetings with subject company in the last 12 months |
| `Total` | Total meeting count with subject company |
| `3rd Party` | Count of meetings where Topic = 3rd Party |
| `Rose & Co` | Count of meetings where Topic is blank or = *Rose & Company |

---

## Tech stack

* **Backend:** Python / Flask
* **Filtering logic:** pandas (`modules/filter.py`)
* **Activities enrichment:** pandas (`modules/activities.py`)
* **Excel output:** openpyxl — navy headers, alternating rows, freeze panes, auto-filter, Criteria sheet (`modules/excel_output.py`)
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

**5. Upload city map** — after first deploy, go to `/admin` and upload `investment_center_nearby_cities.xlsx` (columns: `Investment Center`, `Nearby City`, `State`) to populate the location routing data. Alternatively, add `city_map.json` directly to the repo root.

---

## Admin page (`/admin`)

Accessible at `/admin` with the `ADMIN_PASSWORD` environment variable. Allows:

* **Upload taxonomy** — update the CDF taxonomy (Industry Focus, Investment Style, Market Cap, Geography values) from an Excel file
* **Upload city map** — update the Investment Center → City → State mapping from an Excel file with columns `Investment Center`, `Nearby City`, `State`

---

## File structure

```
ndr-launch/
├── app.py                    # Flask routes, city map loading, city/IC mapping
├── city_map.json             # Investment center → city → state mapping (829 rows, 44 ICs)
├── taxonomy.json             # CDF taxonomy (auto-generated if missing)
├── requirements.txt
├── render.yaml               # Render deployment config
├── modules/
│   ├── filter.py             # Core CDF filtering, split logic, activity enrichment
│   ├── activities.py         # Activity file parsing and activity-only contact building
│   ├── excel_output.py       # Formatted Excel generation, Criteria sheet, sheet name sanitization
│   ├── ai_analysis.py        # Claude-powered CDF recommendations
│   └── sharepoint.py         # Optional SharePoint upload
├── static/
│   ├── app.js                # UI logic, CDF selection, city routing mode switching
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
