# NDR Launch — Contact Filtering App

A Flask web app that filters BD Advanced investor contacts by CDF criteria to generate a targeted, formatted contact list for an NDR or roadshow.

---

## What it does

Upload your BD Advanced exports and optionally company background documents. The app:

1. **Infers CDF criteria** from company documents using AI (Claude Haiku), or lets you enter them manually
2. **Filters contacts** across four dimensions: Industry Focus, Investment Style, Market Cap, Geography
3. **Enriches results** with institutional ownership data and meeting history from Activities.xlsx
4. **Routes contacts** to a single `Contacts` sheet or splits by Investment Center, City, or State — with Virtual sub-tabs for unmatched contacts in all routed modes
5. **Splits output** into structured sheets — HFs, Fixed Income, DNC, Check, Quant, Activist, Excluded, Too Small — and downloads a formatted Excel file

---

## Input files

| File | Required | Header Row | Description |
| --- | --- | --- | --- |
| `Contacts w CDFs.xlsx` | ✅ | Row 3 | Main contacts export from BD Advanced |
| `Activities.xlsx` | Optional | Row 2 | Meeting history; drives enrichment, activity-only contacts, and meeting exclusion |
| `Ownership.xlsx` | Optional | Row 5 | Adds `Shares` column (account-level) |
| `Fund-Level Ownership.xlsx` | Optional | Row 5 | Adds four fund-level columns: Fund Shares, Passive or Index Shares, Total Funds, Passive or Index Funds |
| Additional List(s) | Optional | Row 3 | One or more supplemental contact files added to output, bypassing CDF filter |
| Company documents (PDF, XLSX, TXT, DOCX) | Optional | — | 10-K, investor deck, etc. — used for AI CDF recommendations |

---

## Output sheets

Sheets are written in this order. Empty sheets are omitted.

| Sheet | Contents |
| --- | --- |
| `Summary` | Always first — company, tickers, CDF criteria, routing mode, exclusion settings, and per-sheet counts |
| `Contacts` | All filtered contacts when routing is Virtual |
| `[City / IC name]` | One tab per selected city or Investment Center when routing is active |
| `Virtual - NAM` | Contacts not matched to any selected IC/city, in North American countries |
| `Virtual - EUR` | Contacts not matched to any selected IC/city, in European countries |
| `Virtual - Other` | Contacts not matched to any selected IC/city, outside NAM/EUR |
| `Too Small` | Contacts below the EAUM minimum threshold (if set) |
| `Fixed Income` | Contacts where `CDF (Contact): Invests in Credit/HY` = Yes |
| `HFs` | Hedge funds, based on HF treatment setting |
| `DNC` | Do Not Call |
| `Check` | Check before calling |
| `Quant` | Quantitative funds |
| `Activist` | Contacts where Activist = Often |
| `Excluded` | Contacts excluded by shareholder threshold or meeting history rule |

Sheet names containing `/`, `\`, `?`, `*`, `[`, `]`, or `:` are sanitized (replaced with `-`) and truncated to 31 characters to comply with Excel requirements.

---

## Workflow options

### Subject ticker(s)

When Activities.xlsx is uploaded, all tickers in the file are detected automatically and presented as two checkbox groups:

**Subject ticker(s)** — one or more tickers identifying the company this NDR is for. All six meeting history columns are computed using only rows matching these tickers. Meeting exclusion applies only to subject company meetings. Multiple subject tickers are pooled together — intended for dual-listed companies (same company, different symbols).

**Other tickers** — contacts who appear in activities for these tickers but are not already in the output are appended with `Source = Other: TICK1, TICK2`. Their meeting columns are left blank and meeting exclusion does not apply to them.

### CDF criteria

Select values from four dimensions. Leaving a dimension blank skips it (all values treated as neutral):

- **Industry Focus** — grouped by sector with search. `*Generalist` in a contact's field always returns a match regardless of criteria.
- **Investment Style** — single or multi-select
- **Market Cap** — Micro, Small, Mid, Large, Mega
- **Geography** — includes global and regional values

Upload company documents (10-K, investor deck) and click **Analyze documents** to have Claude pre-fill recommended CDF criteria with reasoning.

### Hedge fund treatment

- **Move to HFs sheet** (default) — contacts where `Type` = Hedge Fund go to the HFs sheet
- **Include in main results** — all HFs remain in main results
- **Low-turnover only in main** — HFs with T/O > 100% go to HFs sheet; low-turnover HFs stay in main

### NDR routing

Four modes, all of which produce Virtual sub-tabs for unmatched contacts:

- **Virtual** (default) — all matched contacts on a single `Contacts` sheet
- **Investment Center** — searchable checklist of all investment centers from `city_map.json`; one tab per selection; unmatched contacts go to Virtual sub-tabs
- **City** — checklist of common cities; each resolves to its BD Advanced Investment Center value; unmatched contacts go to Virtual sub-tabs
- **State** — checklist of state/province codes from `city_map.json`; all ICs with cities in selected states are included; unmatched contacts go to Virtual sub-tabs

Routing operates at the Investment Center level regardless of which mode was used to select.

### Virtual scope

A separate setting that applies in **all routing modes**. Controls which Virtual tab(s) appear in the output:

- **Both (NAM + EUR)** (default) — all unmatched contacts retained across `Virtual - NAM`, `Virtual - EUR`, and `Virtual - Other`
- **NAM only** — only `Virtual - NAM` is written; EUR and Other tabs are dropped
- **EUR only** — only `Virtual - EUR` is written; NAM and Other tabs are dropped

In pure Virtual mode (single `Contacts` sheet), this filter is applied to the `Contacts` tab itself.

### Meeting history exclusion

Requires Activities.xlsx and a selected subject ticker:

- **Include all** (default)
- **Exclude last 12 months** — contacts who met the subject company in the last 12 months → `Excluded`
- **Exclude last 24 months** — contacts who met the subject company in the last 24 months → `Excluded`
- **Exclude all prior meetings** — any contact with any prior subject company meeting → `Excluded`

### Shareholder exclusion

Requires Ownership.xlsx. Based on the `% S/O` column (falls back to column index 4 if not found by name). Raw values are divided by 100:

- **Include all** (default)
- **Exclude all shareholders** — any contact with `% S/O` > 0 → `Excluded`
- **> 0.01% / 0.02% / 0.03% / 0.4% / 0.5%** — contacts exceeding the threshold → `Excluded` with reason "Exceeds Shareholder Limit"

### EAUM minimum

Optional threshold in $mm. Contacts with a non-blank EAUM below this value are moved to the `Too Small` sheet.

---

## Filtering logic

Each contact is evaluated across four CDF dimensions. Each dimension returns **match**, **exclude**, or **neutral**:

| Outcome | Condition |
| --- | --- |
| `neutral` | Dimension was skipped (nothing selected) OR contact's CDF field is blank |
| `match` | Contact's field is populated and at least one value overlaps the target set |
| `exclude` | Contact's field is populated but no values match the target set |

**A contact is included if:** at least one dimension returns `match` AND no dimension returns `exclude`.

**Industry special rule:** `*Generalist` in a contact's Industry field always returns `match` regardless of criteria.

### Split order (sequential — each split operates only on remaining main contacts)

1. **Too Small** — EAUM non-null and below EAUM minimum
2. **HFs** — based on HF treatment setting (T/O threshold is > 1.0 after ÷100 normalization)
3. **Fixed Income** — `CDF (Contact): Invests in Credit/HY` = Yes
4. **DNC** — `CDF (Contact): Do Not Call` or `CDF (Firm): Do Not Call` is non-blank
5. **Check** — `CDF (Firm): Check before calling` = Yes
6. **Quant** — `CDF (Contact): Is Quant?` = Yes
7. **Activist** — `Activist` = Often
8. **Excluded (shareholders)** — shareholder exclusion based on `% S/O` threshold
9. **Excluded (meetings)** — meeting history exclusion (subject company only)

### Contact sources

| Source | Meaning |
| --- | --- |
| `CDF Match` | Passed the CDF filter from the contacts file |
| `Meeting History` | In Activities for subject company but not in the contacts file |
| `Additional List` | From a supplemental contacts file (bypasses CDF filter) |
| `Other: TICK1, TICK2` | In Activities for other-company tickers, not already in output |

---

## Activities enrichment

When Activities.xlsx is uploaded and a subject ticker is selected, the app:

- Joins `Last Mtg btwn Contact & Co` and `Last Mtg btwn firm & Co` to each contact
- Computes `L12M`, `Total`, `3rd Party`, `Rose & Co` meeting counts (blank if zero)
- Appends **activity-only contacts** — people in meeting history for the subject company who are not in the contacts file, reconstructed from available activity fields
- Overrides `Type` to `Hedge Fund` for contacts whose Activities investment style is `Alternative`

---

## Investment Center derivation

If `Contact Investment Center` is blank, it is derived from `City`, `State/Province`, and `Country/Territory` using a three-level lookup in `filter.py`:

1. **City** — direct city → IC mapping (e.g. Greenwich → New York/Southern CT/Northern NJ)
2. **State** — state/province → IC mapping (e.g. Massachusetts → Boston MA)
3. **Country** — country → IC mapping for international contacts (e.g. United Kingdom → London)

Falls back to `City, State` or `City, Country` freeform if no mapping is found.

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

**T/O %** — raw turnover values from BD Advanced are divided by 100 on load (e.g. `50` → `0.50`). Formatted as `#,##0.0` in output.

**Meeting history columns** (populated from Activities, subject company only):

| Column | Description |
| --- | --- |
| `Last Mtg btwn Contact & Co` | Most recent meeting date between this contact and the subject company |
| `Last Mtg btwn firm & Co` | Most recent meeting date between any contact at this institution and the subject company |
| `L12M` | Count of meetings in the last 12 months (blank if zero) |
| `Total` | Total meeting count (blank if zero) |
| `3rd Party` | Count of meetings where Topic = 3rd Party (blank if zero) |
| `Rose & Co` | Count of meetings where Topic is blank or = *Rose & Company (blank if zero) |

**Placeholder columns** added to all output sheets for manual entry: `Out1`, `Out2`, `Status`, `As of`, `Last Mtg. w/ Any Co`.

**Excel formatting:** Navy (`#1F3864`) header row, alternating row fill (`#EEF2F7`), freeze panes at A2, auto-filter on all columns, auto-width (min 8, max 42 characters). `Industry`, `Geo`, `Style`, `Mkt. Cap` use shrink-to-fit at fixed 27-character width.

---

## Tech stack

| Layer | Technology |
| --- | --- |
| Backend | Python 3.11 / Flask |
| Filtering & enrichment | pandas (`modules/filter.py`) |
| Excel output | openpyxl (`modules/excel_output.py`) |
| AI analysis | Anthropic `claude-haiku-4-5` (`modules/ai_analysis.py`) |
| SharePoint upload | Microsoft Graph API via `requests` (`modules/sharepoint.py`) — optional |
| Frontend | Vanilla JS (`static/app.js`), Jinja2 templates; IC and state lists injected as `window.CITY_MAP_ICS` / `window.CITY_MAP_STATES` |
| Hosting | Render.com (Python web service, 2 Gunicorn workers) |

---

## Deployment

### Prerequisites

- GitHub account
- Render.com account (free tier works)
- Anthropic API key from [console.anthropic.com](https://console.anthropic.com)

### Steps

**1. Push to GitHub**

```bash
git add .
git commit -m "Initial deploy"
git push
```

**2. Create Render Web Service**

- dashboard.render.com → New → Web Service
- Connect your GitHub repo
- Render auto-detects `render.yaml`
- Build command: `pip install -r requirements.txt`
- Start command: `gunicorn app:app --workers 2 --timeout 120 --bind 0.0.0.0:$PORT`

**3. Set environment variables**

| Variable | Value |
| --- | --- |
| `ANTHROPIC_API_KEY` | Your key from console.anthropic.com |
| `SECRET_KEY` | Click Generate (auto-set by render.yaml) |
| `ADMIN_PASSWORD` | Password for /admin page |

SharePoint upload is optional — leave blank to disable:

| Variable | Notes |
| --- | --- |
| `AZURE_TENANT_ID` | From Azure Portal |
| `AZURE_CLIENT_ID` | From Azure Portal |
| `AZURE_CLIENT_SECRET` | From Azure Portal |
| `SHAREPOINT_SITE_ID` | From IT admin or Azure |
| `SHAREPOINT_FOLDER` | Default: `/NDR Launch` |

Meetings database is optional — requires a PostgreSQL add-on in Render:

| Variable | Notes |
| --- | --- |
| `DATABASE_URL` | Set automatically by Render PostgreSQL add-on |

**4. Deploy** — click Create Web Service. First deploy takes ~2–4 minutes.

**5. Upload city map** — go to `/admin` and upload an Excel file with columns `Investment Center`, `Nearby City`, `State` to populate location routing data. The bundled `city_map.json` has 829 rows across 44 investment centers.

---

## Admin page (`/admin`)

Password-protected. Provides:

- **CDF taxonomy** — update CDF dropdown values from an Excel file (columns: Field Type, Value, Description). Saved to `taxonomy.json`; falls back to hardcoded defaults if file is absent.
- **City map** — update Investment Center → City → State mapping from an Excel file (columns: Investment Center, Nearby City, State). Saved to `city_map.json` and immediately reflected in the IC and State routing pickers.
- **Meetings database** — upload Activities.xlsx to a PostgreSQL database. Supports incremental (add/update) and full replace modes. Only shown if `DATABASE_URL` is configured.

---

## File structure

```
ndr-launch/
├── app.py                    # Flask routes, city map, taxonomy, file loading
├── city_map.json             # Investment center → city → state (829 rows, 44 ICs)
├── taxonomy.json             # CDF taxonomy values (auto-generated from defaults if missing)
├── requirements.txt
├── render.yaml               # Render deployment config
├── runtime.txt               # python-3.11.9
├── modules/
│   ├── filter.py             # CDF filtering, splits, virtual classification, activity enrichment
│   ├── excel_output.py       # Excel generation, Summary sheet, sheet name sanitization, formatting
│   ├── ai_analysis.py        # Claude-powered CDF recommendations
│   ├── sharepoint.py         # Optional SharePoint upload via Microsoft Graph API
│   ├── meetings.py           # Optional PostgreSQL meetings database (requires DATABASE_URL)
│   └── activities.py         # Legacy module — not imported, superseded by filter.py
├── static/
│   ├── app.js                # UI: CDF selection, ticker checkboxes, routing mode, results display
│   ├── admin.js              # Admin page: taxonomy, city map, meetings upload
│   └── style.css
└── templates/
    ├── base.html
    ├── index.html
    └── admin.html
```

---

## Local development

```bash
pip install -r requirements.txt
ANTHROPIC_API_KEY=your_key SECRET_KEY=dev python app.py
```

App runs at http://localhost:5000
