# NDR Launch — Contact Filtering App

A web app that filters BD Advanced investor contacts by CDF criteria to generate a targeted list for an NDR or roadshow. Built from the `ndr-launch` skill (v10).

---

## What it does

Upload your BD Advanced exports and optionally a company background document (10-K, investor deck). The app:

1. **Infers CDF criteria** from your company documents using AI, or lets you enter them manually
2. **Filters contacts** across four dimensions: Industry Focus, Investment Style, Market Cap, Geography
3. **Routes contacts** to city-based tabs (New York, Boston, etc.) or a single Virtual sheet
4. **Enriches results** with ownership data and meeting history from Activities.xlsx
5. **Splits output** into structured sheets: Contacts, HFs, DNC, Check, Quant, Activist, Excluded
6. **Downloads** a formatted Excel file ready to use

---

## Input files

| File | Required | Description |
|---|---|---|
| `Contacts w CDFs.xlsx` | ✅ | Main contacts export from bdadvanced.ipreo.com (header row 3) |
| `Ownership.xlsx` | Optional | Adds `Shares` column (header row 5) |
| `Fund-Level Ownership.xlsx` | Optional | Adds four fund-level columns (header row 5) |
| `Activities.xlsx` | Optional | Adds six meeting history columns; used for meeting exclusion |
| Company document (PDF) | Optional | 10-K, investor deck, etc. — triggers AI CDF recommendations |

---

## Output sheets

| Sheet | Contents |
|---|---|
| `[City]` or `Contacts` | Filtered contacts routed by city, or all contacts if Virtual |
| `Virtual` | Contacts not matching any selected city (only when cities are chosen) |
| `HFs` | Hedge funds (if HF treatment = No or Low-turnover only) |
| `DNC` | Do Not Contact |
| `Check` | Contacts flagged for review |
| `Quant` | Quantitative funds |
| `Activist` | Contacts where Activist = Often |
| `Excluded` | Contacts excluded by meeting history rule |
| `Activity-Only` | Contacts found in Activities.xlsx but not in the contacts file |

---

## Workflow options

**Hedge fund treatment**
- Include all in main results
- Move all HFs to a separate sheet
- Keep low-turnover (T/O ≤ 100%) in main results, move high-turnover to HFs sheet

**City routing**
- Virtual (single Contacts sheet)
- 1–4 cities → creates one tab per city + a Virtual catch-all tab

Supported city shortcuts: NY/NYC, Boston, Chicago, Philly, SF, LA, Dallas, Houston, Minneapolis, Florida/Miami, London, Paris, Amsterdam, Tokyo, Hong Kong, Toronto, Columbus, Kansas City, San Antonio.

**Meeting history exclusion**
- Include all
- Exclude contacts met in L12M
- Exclude contacts met in L24M
- Exclude all contacts with any prior meeting history

---

## Tech stack

- **Backend:** Python / Flask
- **Filtering logic:** pandas
- **Excel output:** openpyxl (navy headers, alternating rows, freeze panes, auto-filter)
- **AI analysis:** Anthropic Claude (claude-sonnet-4-20250514)
- **Hosting:** Render.com

---

## Deployment

### Prerequisites
- GitHub account
- Render.com account (free tier works)
- Anthropic API key from console.anthropic.com

### Steps

**1. Push to GitHub**
```bash
git clone https://github.com/YOUR_USER/ndr-launch.git
cd ndr-launch
git add .
git commit -m "Deploy v10"
git push
```

**2. Create Render Web Service**
- dashboard.render.com → New → Web Service
- Connect your GitHub repo
- Render auto-detects `render.yaml`
- Build command: `pip install -r requirements.txt`
- Start command: `gunicorn app:app --workers 2 --timeout 120`

**3. Set environment variables in Render dashboard**

| Variable | Value |
|---|---|
| `ANTHROPIC_API_KEY` | Your key from console.anthropic.com |
| `SECRET_KEY` | Click Generate |
| `ADMIN_PASSWORD` | Password for /admin page |

SharePoint variables are optional (leave blank to skip):

| Variable | Notes |
|---|---|
| `AZURE_TENANT_ID` | From Azure Portal |
| `AZURE_CLIENT_ID` | From Azure Portal |
| `AZURE_CLIENT_SECRET` | From Azure Portal |
| `SHAREPOINT_SITE_ID` | From IT admin or Azure |
| `SHAREPOINT_FOLDER` | Default: `/NDR Launch` |

**4. Deploy** — click Create Web Service. First deploy takes ~2 minutes.

---

## File structure

```
ndr-launch-app/
├── app.py                    # Flask routes
├── requirements.txt
├── render.yaml               # Render deployment config
├── .env.example              # Local dev template
├── modules/
│   ├── filter.py             # Core CDF filtering logic
│   ├── activities.py         # Meeting history enrichment
│   ├── excel_output.py       # Formatted Excel generation
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

```bash
cp .env.example .env
# Fill in ANTHROPIC_API_KEY and SECRET_KEY in .env
pip install -r requirements.txt
python app.py
```

App runs at http://localhost:5000
