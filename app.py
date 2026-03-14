import os
import io
import json
import uuid
import tempfile
import pandas as pd
from flask import (Flask, render_template, request, jsonify,
                   send_file, session, redirect, url_for)
from modules.filter import run_filter, load_activities
from modules.sharepoint import SharePointClient
from modules.ai_analysis import analyze_documents
from modules.excel_output import generate_excel

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'change-this-in-production')
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'admin')
TEMP_DIR       = tempfile.mkdtemp()

# City name → Investment Center mapping (mirrors skill Step 0-B2)
CITY_IC_MAP = {
    'new york': 'New York/Southern CT/Northern NJ',
    'ny': 'New York/Southern CT/Northern NJ',
    'nyc': 'New York/Southern CT/Northern NJ',
    'boston': 'Boston MA',
    'chicago': 'Chicago IL',
    'philadelphia': 'Philadelphia PA/Wilmington DE',
    'philly': 'Philadelphia PA/Wilmington DE',
    'san francisco': 'San Francisco/San Jose CA',
    'sf': 'San Francisco/San Jose CA',
    'los angeles': 'Los Angeles/Pasadena CA',
    'la': 'Los Angeles/Pasadena CA',
    'dallas': 'Dallas/Ft. Worth TX',
    'houston': 'Houston TX',
    'minneapolis': 'Minneapolis/St. Paul MN',
    'florida': 'South Florida/Orlando FL/Tampa-St.Pete FL',
    'miami': 'South Florida/Orlando FL/Tampa-St.Pete FL',
    'south florida': 'South Florida/Orlando FL/Tampa-St.Pete FL',
    'london': 'London',
    'paris': 'Paris',
    'amsterdam': 'Amsterdam',
    'tokyo': 'Tokyo',
    'hong kong': 'Hong Kong',
    'toronto': 'Toronto',
    'columbus': 'Columbus OH',
    'kansas city': 'Kansas City MO',
    'san antonio': 'San Antonio TX',
    'denver': 'Denver',
    'atlanta': 'Atlanta',
    'nashville': 'Nashville',
}

DISPLAY_CITIES = [
    ('New York',       'New York/Southern CT/Northern NJ'),
    ('Boston',         'Boston MA'),
    ('Chicago',        'Chicago IL'),
    ('Philadelphia',   'Philadelphia PA/Wilmington DE'),
    ('San Francisco',  'San Francisco/San Jose CA'),
    ('Los Angeles',    'Los Angeles/Pasadena CA'),
    ('Dallas',         'Dallas/Ft. Worth TX'),
    ('Houston',        'Houston TX'),
    ('Minneapolis',    'Minneapolis/St. Paul MN'),
    ('South Florida',  'South Florida/Orlando FL/Tampa-St.Pete FL'),
    ('Denver',         'Denver'),
    ('Atlanta',        'Atlanta'),
    ('Nashville',      'Nashville'),
    ('Kansas City',    'Kansas City MO'),
    ('Columbus',       'Columbus OH'),
    ('San Antonio',    'San Antonio TX'),
    ('London',         'London'),
    ('Paris',          'Paris'),
    ('Amsterdam',      'Amsterdam'),
    ('Tokyo',          'Tokyo'),
    ('Hong Kong',      'Hong Kong'),
    ('Toronto',        'Toronto'),
]


# ── Taxonomy ──────────────────────────────────────────────────────────────────

TAXONOMY_PATH = os.path.join(os.path.dirname(__file__), 'taxonomy.json')

def load_taxonomy():
    if os.path.exists(TAXONOMY_PATH):
        try:
            with open(TAXONOMY_PATH) as f:
                return json.load(f)
        except Exception:
            pass
    return _default_taxonomy()

def save_taxonomy(data):
    with open(TAXONOMY_PATH, 'w') as f:
        json.dump(data, f)

def _default_taxonomy():
    return {
        'Geography': [
            {'value': '***Intl ADR', 'description': ''}, {'value': '**Emerging Markets', 'description': ''},
            {'value': '*Global', 'description': ''}, {'value': '*Global (ex US)', 'description': ''},
            {'value': 'Africa', 'description': ''}, {'value': 'Africa \u2013 South Pacific', 'description': ''},
            {'value': 'APAC', 'description': ''}, {'value': 'Asia Pacific', 'description': ''},
            {'value': 'Asia Pacific: India', 'description': ''}, {'value': 'Australia', 'description': ''},
            {'value': 'Canada \u2013 North America', 'description': ''}, {'value': 'Europe', 'description': ''},
            {'value': 'Europe \u2013 Israel', 'description': ''}, {'value': 'Europe \u2013 Norway', 'description': ''},
            {'value': 'Europe \u2013 UK', 'description': ''}, {'value': 'Middle East', 'description': ''},
            {'value': 'North America', 'description': ''}, {'value': 'North America (US-listed only)', 'description': ''},
            {'value': 'South America', 'description': ''},
        ],
        'Market capitalization': [
            {'value': 'Micro', 'description': ''}, {'value': 'Small', 'description': ''},
            {'value': 'Mid', 'description': ''}, {'value': 'Large', 'description': ''}, {'value': 'Mega', 'description': ''},
        ],
        'Investment style': [
            {'value': v, 'description': ''} for v in [
                'Aggressive growth', 'Asset allocator', 'Blend', 'Convertibles', 'Deep Value', 'Distressed',
                'ESG administrator', 'ESG investor', 'GARP', 'Growth', 'Hedge fund', 'Macro', 'Event-driven',
                'Special situations', 'Real Assets', 'Shariah', 'SPAC', 'SPAC (pre-merger)', 'Value', 'Wealth Manager', 'Yield',
            ]
        ],
        'Industry Focus': [
            {'value': v, 'description': ''} for v in [
                '*Generalist', 'Agriculture', 'Basic Materials', 'Basic Materials: Aluminum/Steel',
                'Basic Materials: Chemicals', 'Basic Materials: Construction Materials', 'Basic Materials: Forest Products',
                'Basic Materials: Lithium', 'Basic Materials: Metals & Mining', 'Basic Materials: Precious Metals',
                'Basic Materials: Uranium', 'Consumer Discretionary: Branded Apparel', 'Consumer Discretionary: Restaurants',
                'Consumer Goods', 'Consumer Goods: Automotive', 'Consumer Goods: Consumer Durables',
                'Consumer Goods: Consumer Non-Durables', 'Consumer Goods: Discretionary',
                'Consumer Goods: Food, Beverage and Tobacco', 'Consumer Services', 'Consumer Services: Gaming',
                'Consumer Services: Health and Wellness', 'Consumer Services: Homebuilding', 'Consumer Services: Internet',
                'Consumer Services: Media', 'Consumer Services: Personal Services', 'Consumer Services: Retail',
                'Consumer Services: Transportation Services', 'Consumer Services: Travel, Services and Leisure',
                'Consumer Services: Wholesale', 'Consumer Staples: Food & Staples Retail',
                'Consumer Staples: Household & Personal Products', 'Energy', 'Energy: Clean & Renewables',
                'Energy: Downstream', 'Energy: Infrastructure', 'Energy: Midstream', 'Energy: MLP',
                'Energy: Oil, Gas and Coal', 'Energy: Renewable Energy Equipment and Services', 'Energy: Upstream',
                'Financials', 'Financials: Asset Management', 'Financials: Banking', 'Financials: BDC',
                'Financials: Exchanges', 'Financials: Financial Services', 'Financials: FinTech', 'Financials: Insurance',
                'Financials: Payment', 'Financials: Real Estate', 'Financials: Real Estate Tech', 'Financials: REITS',
                'Financials: Specialized', 'Healthcare', 'Healthcare: Biotechnology and Pharmaceuticals',
                'Healthcare: Health Services', 'Healthcare: Healthcare and Supplies Wholesale',
                'Healthcare: Information Technology', 'Healthcare: Medical Equipment', 'Industrials',
                'Industrials: Aerospace and Defense', 'Industrials: Building Products', 'Industrials: Business Services',
                'Industrials: Commercial and Professional Services', 'Industrials: Conglomerates', 'Industrials: E&C',
                'Industrials: Environment Services', 'Industrials: General Industrials', 'Industrials: Industrial Equipment',
                'Industrials: Industrial Goods and Services', 'Industrials: Marine', 'Industrials: Materials and Construction',
                'Industrials: Road/Rail', 'Industrials: Staffing', 'Industrials: Transportation', 'Infrastructure',
                'Packaging', 'Technology', 'Technology: Comm Equip', 'Technology: Computer Software and Services',
                'Technology: Internet', 'Technology: IT Services and Technology', 'Technology: Semiconductors',
                'Technology: Software', 'Technology: Technology Hardware and Equipment', 'Technology: Telecommunications',
                'Thematic - Clean Energy', 'Thematic - Climate Solutions', 'Thematic - Digital',
                'Thematic - Health and Wellness', 'Thematic - Nutrition', 'Thematic - Security', 'Thematic - Water',
                'Thematic - Global Franchise', 'Thematic - Innovation', 'Thematic - Mobility',
                'Thematic - Pet/Animal Related', 'Thematic - Population', 'Thematic - Robotics & AI',
                'Thematic - SmartCity', 'Thematic - Space Exploration', 'Thematic - Timber', 'Utilities',
            ]
        ],
    }


# ── Main routes ───────────────────────────────────────────────────────────────

@app.route('/')
def index():
    taxonomy = load_taxonomy()
    sp_configured = False
    try:
        sp_configured = SharePointClient().is_configured()
    except Exception:
        pass
    return render_template('index.html',
                           taxonomy_json=json.dumps(taxonomy),
                           sp_configured=sp_configured,
                           display_cities=DISPLAY_CITIES)


@app.route('/api/get-symbols', methods=['POST'])
@app.route('/api/detect-symbols', methods=['POST'])
def get_symbols():
    f = request.files.get('activities')
    if not f or f.filename == '':
        return jsonify({'symbols': []})
    try:
        df = pd.read_excel(io.BytesIO(f.read()), header=1)
        symbols = sorted(df['Symbols'].dropna().astype(str).str.strip().str.upper().unique().tolist())
        return jsonify({'symbols': symbols})
    except Exception as e:
        return jsonify({'error': str(e)}), 400


@app.route('/api/analyze', methods=['POST'])
def analyze():
    files = request.files.getlist('documents')
    if not files or all(f.filename == '' for f in files):
        return jsonify({'error': 'No documents uploaded'}), 400
    taxonomy  = load_taxonomy()
    file_data = [{'name': f.filename, 'data': f.read(), 'type': f.content_type or ''} for f in files]
    try:
        return jsonify(analyze_documents(file_data, taxonomy))
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/run', methods=['POST'])
def run():
    contacts_file  = request.files.get('contacts')
    ownership_file = request.files.get('ownership')
    fund_file      = request.files.get('fund_ownership')
    activities_file = request.files.get('activities')

    if not contacts_file or contacts_file.filename == '':
        return jsonify({'error': 'Contacts file is required'}), 400

    def to_set(vals):
        s = {v for v in vals if v}
        return s if s else None

    criteria = {
        'industry': to_set(request.form.getlist('industry')),
        'style':    to_set(request.form.getlist('style')),
        'mcap':     to_set(request.form.getlist('mcap')),
        'geo':      to_set(request.form.getlist('geo')),
    }

    hf_treatment      = request.form.get('hf_treatment', 'separate')
    eaum_min_raw      = request.form.get('eaum_min', '').strip()
    eaum_min          = float(eaum_min_raw) if eaum_min_raw else None
    meeting_exclusion = request.form.get('meeting_exclusion', 'include_all')
    company_name      = request.form.get('company_name', 'Company').strip() or 'Company'
    subject_symbols   = [s.strip().upper() for s in request.form.getlist('subject_symbols') if s.strip()]
    routing_mode      = request.form.get('city_mode', 'virtual')

    # Parse city selections
    city_selections = None
    if routing_mode == 'cities':
        selected_cities = request.form.getlist('selected_cities')
        app.logger.info(f'City routing: mode={routing_mode}, selected={selected_cities}')
        city_selections = []
        for city_key in selected_cities:
            ck = city_key.strip().lower()
            ic = CITY_IC_MAP.get(ck)
            if not ic:
                for tab_name, ic_val in DISPLAY_CITIES:
                    if tab_name.lower() == ck:
                        ic = ic_val
                        city_key = tab_name
                        break
            if ic:
                city_selections.append((city_key.title() if city_key == city_key.lower() else city_key, ic))
        app.logger.info(f'city_selections built: {city_selections}')

    # Load files
    try:
        contacts_df = pd.read_excel(io.BytesIO(contacts_file.read()), header=2)
    except Exception as e:
        return jsonify({'error': f'Could not read contacts file: {e}'}), 400

    ownership_df = fund_df = acts_named = None

    if ownership_file and ownership_file.filename:
        try:
            ownership_df = pd.read_excel(io.BytesIO(ownership_file.read()), header=4)
        except Exception:
            pass

    if fund_file and fund_file.filename:
        try:
            fund_df = pd.read_excel(io.BytesIO(fund_file.read()), header=4)
        except Exception:
            pass

    if activities_file and activities_file.filename and subject_symbols:
        try:
            acts_df = pd.read_excel(io.BytesIO(activities_file.read()), header=1)
            acts_named = load_activities(acts_df, subject_symbols)
        except Exception:
            pass

    try:
        results = run_filter(
            contacts_df, ownership_df, fund_df, acts_named,
            criteria, hf_treatment, meeting_exclusion,
            city_selections, subject_symbols, company_name,
            eaum_min=eaum_min
        )
    except Exception as e:
        return jsonify({'error': f'Filter error: {e}'}), 500

    try:
        excel_bytes = generate_excel(results, company_name)
    except Exception as e:
        return jsonify({'error': f'Excel generation error: {e}'}), 500

    # Save for download
    file_id  = str(uuid.uuid4())
    symbol_label = ' '.join(subject_symbols) if subject_symbols else company_name
    filename = f'{symbol_label} Contacts Mapping.xlsx'
    with open(os.path.join(TEMP_DIR, f'{file_id}.xlsx'), 'wb') as f:
        f.write(excel_bytes)
    session['download_id']   = file_id
    session['download_name'] = filename

    # Push to SharePoint
    sharepoint_url = None
    try:
        sp = SharePointClient()
        if sp.is_configured():
            sharepoint_url = sp.upload_file(excel_bytes, filename)
    except Exception:
        pass

    # Build sheet counts
    frames = results['frames']
    sheet_counts = {}
    for k, v in frames.items():
        if v is not None and len(v) > 0:
            sheet_counts[k] = len(v)

    # Match breakdown (main/city sheets only)
    main_frames = {k: v for k, v in frames.items()
                   if k not in ('Too Small', 'HFs', 'DNC', 'Check', 'Quant', 'Activist', 'Excluded')
                   and v is not None and len(v) > 0}
    combined = pd.concat(list(main_frames.values()), ignore_index=True) if main_frames else pd.DataFrame()
    match_breakdown = {}
    if len(combined) > 0 and 'Match Count' in combined.columns:
        for count, grp in combined.groupby('Match Count'):
            try:
                match_breakdown[int(count)] = len(grp)
            except Exception:
                pass

    # Build city_counts and main_count for JS
    excluded_sheets = {'Too Small', 'HFs', 'DNC', 'Check', 'Quant', 'Activist', 'Excluded'}
    if results['has_city_routing']:
        city_counts = {k: v for k, v in sheet_counts.items() if k not in excluded_sheets}
        main_count  = sum(city_counts.values())
    else:
        city_counts = {}
        main_count  = sheet_counts.get('Contacts', 0)

    return jsonify({
        'total_source':    results['total_source'],
        'total_matched':   results['total_matched'],
        'sheet_counts':    sheet_counts,
        'city_counts':     city_counts,
        'main_count':      main_count,
        'hf_count':        sheet_counts.get('HFs', 0),
        'dnc_count':       sheet_counts.get('DNC', 0),
        'check_count':     sheet_counts.get('Check', 0),
        'quant_count':     sheet_counts.get('Quant', 0),
        'activist_count':  sheet_counts.get('Activist', 0),
        'too_small_count': sheet_counts.get('Too Small', 0),
        'excluded_count':  sheet_counts.get('Excluded', 0),
        'match_breakdown': match_breakdown,
        'sharepoint_url':  sharepoint_url,
        'filename':        filename,
        'has_city_routing': results['has_city_routing'],
    })


@app.route('/download')
def download():
    file_id  = session.get('download_id')
    filename = session.get('download_name', 'Contacts.xlsx')
    if not file_id:
        return 'No file available', 404
    path = os.path.join(TEMP_DIR, f'{file_id}.xlsx')
    if not os.path.exists(path):
        return 'File expired — please re-run the filter', 404
    return send_file(path, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# ── Admin routes ──────────────────────────────────────────────────────────────

@app.route('/admin', methods=['GET', 'POST'])
def admin():
    if request.method == 'POST' and not session.get('admin'):
        if request.form.get('password') == ADMIN_PASSWORD:
            session['admin'] = True
        else:
            return render_template('admin.html', authenticated=False, error='Incorrect password')
    if not session.get('admin'):
        return render_template('admin.html', authenticated=False)
    return render_template('admin.html', authenticated=True)


@app.route('/admin/logout')
def admin_logout():
    session.pop('admin', None)
    return redirect(url_for('admin'))


@app.route('/api/admin/upload-taxonomy', methods=['POST'])
def upload_taxonomy():
    if not session.get('admin'):
        return jsonify({'error': 'Unauthorized'}), 401
    f = request.files.get('taxonomy_file')
    if not f:
        return jsonify({'error': 'No file provided'}), 400
    try:
        df = pd.read_excel(io.BytesIO(f.read()))
        taxonomy = {}
        for _, row in df.iterrows():
            ft  = str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else ''
            val = str(row.iloc[1]).strip() if len(row) > 1 and not pd.isna(row.iloc[1]) else ''
            desc = str(row.iloc[2]).strip() if len(row) > 2 and not pd.isna(row.iloc[2]) else ''
            if ft and val:
                if ft not in taxonomy:
                    taxonomy[ft] = []
                taxonomy[ft].append({'value': val, 'description': desc})
        save_taxonomy(taxonomy)
        total = sum(len(v) for v in taxonomy.values())
        return jsonify({'success': True, 'rows': total})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.template_filter('format_number')
def format_number(n):
    try:
        return f'{int(n):,}'
    except Exception:
        return n


if __name__ == '__main__':
    app.run(debug=os.environ.get('FLASK_ENV') == 'development', host='0.0.0.0', port=5000)
