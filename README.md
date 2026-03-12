# NDR Launch — Contact Filter App

A web app that filters BD Advanced investor contacts by CDF criteria, enriches results with meeting history, and pushes output to SharePoint.

---

## What it does

- Uploads and filters BD Advanced contact exports by Industry, Style, Market Cap, and Geography
- AI-analyzes company documents (10-K, decks) to auto-recommend CDF criteria
- Enriches results with meeting history (last meeting with your company, last meeting with any company, full list of companies met)
- Splits output into Contacts / HFs / DNC / Check / Quant sheets
- Pushes formatted Excel directly to SharePoint
- Stores 180k+ meeting history rows in PostgreSQL with incremental update support

---

## Deploy to Render (step by step)

### 1. Create accounts
- [GitHub](https://github.com) — free account
- [Render](https://render.com) — free account (sign in with GitHub)

### 2. Create a GitHub repository
1. Go to github.com → click **New repository**
2. Name it `ndr-launch` (or anything you like)
3. Set to **Private**
4. Click **Create repository**

### 3. Upload these files to GitHub
1. In your new repo, click **Add file → Upload files**
2. Drag the entire contents of this folder into the upload area
3. Click **Commit changes**

### 4. Deploy on Render
1. Go to [render.com/dashboard](https://render.com/dashboard)
2. Click **New → Blueprint**
3. Connect your GitHub repo
4. Render reads `render.yaml` and sets up the app + database automatically
5. Click **Apply**

### 5. Add environment variables in Render
Go to your service → **Environment** and add:

| Variable | Value |
|---|---|
| `ADMIN_PASSWORD` | A password for the admin page |
| `ANTHROPIC_API_KEY` | Your key from console.anthropic.com |

The `DATABASE_URL` and `SECRET_KEY` are set automatically by Render.

### 6. SharePoint setup (optional — adds "Open in SharePoint" button)

You need a one-time Azure app registration:

1. Go to [portal.azure.com](https://portal.azure.com)
2. Search for **App registrations** → **New registration**
3. Name it `NDR Launch`, click **Register**
4. Copy the **Application (client) ID** and **Directory (tenant) ID**
5. Go to **Certificates & secrets → New client secret** — copy the value immediately
6. Go to **API permissions → Add a permission → Microsoft Graph → Application permissions**
7. Add `Sites.ReadWrite.All` — click **Grant admin consent**
8. Find your SharePoint site ID:
   - Visit: `https://graph.microsoft.com/v1.0/sites/{your-domain}.sharepoint.com:/sites/{site-name}`
   - Or ask your IT admin

Then add to Render environment variables:

| Variable | Value |
|---|---|
| `AZURE_TENANT_ID` | Directory (tenant) ID from Azure |
| `AZURE_CLIENT_ID` | Application (client) ID from Azure |
| `AZURE_CLIENT_SECRET` | Client secret value from Azure |
| `SHAREPOINT_SITE_ID` | Your SharePoint site ID |
| `SHAREPOINT_FOLDER` | Folder path, e.g. `/NDR Launch` |

---

## Admin page

Visit `/admin` to:
- **Upload meeting history** — full refresh or incremental update with preview
- **Update taxonomy** — upload a new CDFs.xlsx to update the picker values and AI guidance
- **View upload history** — log of all uploads with row counts

The meeting history file should contain columns (exact names detected automatically):
- Date / Meeting Date
- Email / Email Address
- First Name
- Last Name
- Account Name / Firm
- Ticker / Symbol

Incremental updates match on: `date + first_name + last_name + account_name + ticker`

---

## Local development

```bash
# Install dependencies
pip install -r requirements.txt

# Copy and fill in environment variables
cp .env.example .env

# Start the app
python app.py
```

Requires Python 3.10+ and a PostgreSQL database (or set DATABASE_URL to a local Postgres instance).

---

## File structure

```
ndr-launch-app/
├── app.py                  # Flask routes
├── requirements.txt
├── render.yaml             # Render deployment config
├── .env.example
├── static/
│   ├── style.css
│   ├── app.js              # Filter page logic
│   └── admin.js            # Admin page logic
├── templates/
│   ├── base.html
│   ├── index.html          # Main filter page
│   └── admin.html          # Admin page
└── modules/
    ├── filter.py           # Core CDF filtering logic
    ├── meetings.py         # PostgreSQL meeting history
    ├── sharepoint.py       # Microsoft Graph API
    ├── ai_analysis.py      # Claude document analysis
    └── excel_output.py     # Formatted Excel generation
```

---

## Changing column order

To adjust the column order in the output Excel, edit the `run_filter` function in `modules/filter.py`. After the filtering and splitting steps, reorder `df.columns` before returning. Example:

```python
desired_order = ['First Name', 'Last Name', 'CRM Account Name', 'Shares', ...]
filtered_df = filtered_df[[c for c in desired_order if c in filtered_df.columns]]
```
