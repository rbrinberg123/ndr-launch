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

# ── City map ───────────────────────────────────────────────────────────────────

CITY_MAP_PATH = os.path.join(os.path.dirname(__file__), 'city_map.json')

def load_city_map():
    if os.path.exists(CITY_MAP_PATH):
        try:
            with open(CITY_MAP_PATH) as f:
                data = json.load(f)
                return data.get('mappings', [])
        except Exception:
            pass
    return []

def save_city_map(mappings):
    with open(CITY_MAP_PATH, 'w') as f:
        json.dump({'mappings': mappings}, f)

def get_city_map_lists():
    """Return (sorted IC list, sorted state list) derived from city_map.json."""
    mappings    = load_city_map()
    seen_ics    = {}
    seen_states = {}
    for m in mappings:
        ic    = m.get('investment_center', '')
        state = m.get('state', '')
        if ic    and ic    not in seen_ics:    seen_ics[ic]       = True
        if state and state not in seen_states: seen_states[state] = True
    return sorted(seen_ics.keys()), sorted(seen_states.keys())

# City name → Investment Center mapping (fallback for city routing)
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


# ── Taxonomy ───────────────────────────────────────────────────────────────────

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


# ── Meetings DB helper ─────────────────────────────────────────────────────────

def get_meetings_db():
    db_url = os.environ.get('DATABASE_URL')
    if not db_url:
        return None
    try:
        from modules.meetings import MeetingsDB
        db = MeetingsDB(db_url)
        db.init_db()
        return db
    except Exception:
        return None


# ── Main routes ────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    taxonomy      = load_taxonomy()
    sp_configured = False
    try:
        sp_configured = SharePointClient().is_configured()
    except Exception:
        pass
    display_ics, display_states = get_city_map_lists()
    return render_template('index.html',
                           taxonomy_json=json.dumps(taxonomy),
                           sp_configured=sp_configured,
                           display_cities=DISPLAY_CITIES,
                           city_map_ics_json=json.dumps(display_ics),
                           city_map_states_json=json.dumps(display_states))


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
    contacts_file   = request.files.get('contacts')
    ownership_file  = request.files.get('ownership')
    fund_file       = request.files.get('fund_ownership')
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

    hf_treatment          = request.form.get('hf_treatment', 'separate')
    meeting_exclusion     = request.form.get('meeting_exclusion', 'include_all')
    shareholder_exclusion = request.form.get('shareholder_exclusion', 'include_all')
    company_name          = request.form.get('company_name', 'Company').strip() or 'Company'
    subject_symbols_raw   = request.form.getlist('subject_symbol')
    subject_symbol        = [s.strip().upper() for s in subject_symbols_raw if s.strip()] or ['']
    subject_symbol        = subject_symbol[0] if len(subject_symbol) == 1 else subject_symbol
    routing_mode          = request.form.get('city_mode', 'virtual')
    virtual_scope         = request.form.get('virtual_scope', 'both')

    eaum_min_raw = request.form.get('eaum_min', '').strip()
    try:
        eaum_min = float(eaum_min_raw) if eaum_min_raw else None
    except ValueError:
        eaum_min = None

    # Build city_selections from routing mode
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

    elif routing_mode == 'investment_center':
        selected_ics = request.form.getlist('selected_ics')
        if selected_ics:
            city_selections = [(ic, ic) for ic in selected_ics if ic]

    elif routing_mode == 'state':
        selected_states = set(request.form.getlist('selected_states'))
        if selected_states:
            mappings = load_city_map()
            seen = {}
            for m in mappings:
                if m.get('state') in selected_states:
                    ic = m['investment_center']
                    if ic not in seen:
                        seen[ic] = True
            city_selections = [(ic, ic) for ic in seen.keys()]

    # Load contacts
    try:
        contacts_df = pd.read_excel(io.BytesIO(contacts_file.read()), header=2)
    except Exception as e:
        return jsonify({'error': f'Could not read contacts file: {e}'}), 400

    ownership_df = fund_df = None

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

    # Load activities — preserve raw df for other_symbols enrichment
    acts_df_raw = None
    acts_named  = None
    if activities_file and activities_file.filename:
        try:
            acts_bytes  = activities_file.read()
            acts_df_raw = pd.read_excel(io.BytesIO(acts_bytes), header=1)
            del acts_bytes  # free raw bytes immediately
            sym_arg = subject_symbol if subject_symbol != [''] else None
            if sym_arg:
                acts_named = load_activities(acts_df_raw, sym_arg)
        except Exception:
            pass

    # Load additional / mining lists
    mining_dfs   = []
    mining_files = request.files.getlist('mining')
    for mf in mining_files:
        if mf and mf.filename:
            try:
                mdf = pd.read_excel(io.BytesIO(mf.read()), header=2)
                mining_dfs.append(mdf)
            except Exception:
                pass
    mining_df = pd.concat(mining_dfs, ignore_index=True) if mining_dfs else None

    # Other symbols — exclude any that are already subject symbols
    other_symbols_raw = request.form.getlist('other_symbols')
    subject_set = set(subject_symbol) if isinstance(subject_symbol, list) else {subject_symbol}
    other_symbols = [s for s in other_symbols_raw if s and s not in subject_set] or None
    if other_symbols and acts_df_raw is None:
        other_symbols = None

    try:
        results = run_filter(
            contacts_df, ownership_df, fund_df, acts_named,
            criteria, hf_treatment, meeting_exclusion,
            city_selections, subject_symbol, company_name,
            eaum_min=eaum_min,
            mining_df=mining_df,
            acts_df_raw=acts_df_raw,
            other_symbols=other_symbols,
            shareholder_exclusion=shareholder_exclusion,
            virtual_scope=virtual_scope,
        )
    except Exception as e:
        return jsonify({'error': f'Filter error: {e}'}), 500

    # Attach routing context for Summary sheet
    results['routing_mode']  = routing_mode
    results['virtual_scope'] = virtual_scope

    try:
        excel_bytes = generate_excel(results, company_name)
    except Exception as e:
        return jsonify({'error': f'Excel generation error: {e}'}), 500

    # Save for download
    file_id  = str(uuid.uuid4())
    sym_str  = (subject_symbol[0] if isinstance(subject_symbol, list) else subject_symbol) or ''
    filename = f'{sym_str or company_name} Contacts Mapping.xlsx'
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

    # Match breakdown (main / city sheets only)
    excluded_sheets = {'HFs', 'DNC', 'Check', 'Quant', 'Activist', 'Excluded',
                       'Too Small', 'Fixed Income'}
    main_frames = {k: v for k, v in frames.items()
                   if k not in excluded_sheets and v is not None and len(v) > 0}
    combined = pd.concat(list(main_frames.values()), ignore_index=True) if main_frames else pd.DataFrame()
    match_breakdown = {}
    if len(combined) > 0 and 'Match Count' in combined.columns:
        for count, grp in combined.groupby('Match Count'):
            try:
                match_breakdown[int(count)] = len(grp)
            except Exception:
                pass

    if results['has_city_routing']:
        city_counts = {k: v for k, v in sheet_counts.items() if k not in excluded_sheets}
        main_count  = sum(city_counts.values())
    else:
        city_counts = {}
        main_count  = sheet_counts.get('Contacts', 0)

    return jsonify({
        'total_source':     results['total_source'],
        'total_matched':    results['total_matched'],
        'sheet_counts':     sheet_counts,
        'city_counts':      city_counts,
        'main_count':       main_count,
        'hf_count':         sheet_counts.get('HFs', 0),
        'dnc_count':        sheet_counts.get('DNC', 0),
        'check_count':      sheet_counts.get('Check', 0),
        'quant_count':      sheet_counts.get('Quant', 0),
        'activist_count':   sheet_counts.get('Activist', 0),
        'excluded_count':   sheet_counts.get('Excluded', 0),
        'too_small_count':  sheet_counts.get('Too Small', 0),
        'fi_count':         sheet_counts.get('Fixed Income', 0),
        'match_breakdown':  match_breakdown,
        'sharepoint_url':   sharepoint_url,
        'filename':         filename,
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


# ── Admin routes ───────────────────────────────────────────────────────────────

@app.route('/admin', methods=['GET', 'POST'])
def admin():
    if request.method == 'POST' and not session.get('admin'):
        if request.form.get('password') == ADMIN_PASSWORD:
            session['admin'] = True
        else:
            return render_template('admin.html', authenticated=False, error='Incorrect password')
    if not session.get('admin'):
        return render_template('admin.html', authenticated=False)
    db_configured = bool(os.environ.get('DATABASE_URL'))
    return render_template('admin.html', authenticated=True, db_configured=db_configured)


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
            ft   = str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else ''
            val  = str(row.iloc[1]).strip() if len(row) > 1 and not pd.isna(row.iloc[1]) else ''
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


@app.route('/api/admin/upload-city-map', methods=['POST'])
def upload_city_map():
    if not session.get('admin'):
        return jsonify({'error': 'Unauthorized'}), 401
    f = request.files.get('city_map_file')
    if not f:
        return jsonify({'error': 'No file provided'}), 400
    try:
        df = pd.read_excel(io.BytesIO(f.read()))
        mappings = []
        for _, row in df.iterrows():
            ic    = str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else ''
            city  = str(row.iloc[1]).strip() if len(row) > 1 and not pd.isna(row.iloc[1]) else ''
            state = str(row.iloc[2]).strip() if len(row) > 2 and not pd.isna(row.iloc[2]) else ''
            if ic and city:
                mappings.append({'investment_center': ic, 'city': city, 'state': state})
        save_city_map(mappings)
        ics = len(set(m['investment_center'] for m in mappings))
        return jsonify({'success': True, 'rows': len(mappings), 'investment_centers': ics})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/admin/preview-meetings', methods=['POST'])
def preview_meetings():
    if not session.get('admin'):
        return jsonify({'error': 'Unauthorized'}), 401
    db = get_meetings_db()
    if not db:
        return jsonify({'error': 'DATABASE_URL not configured'}), 400
    f = request.files.get('meetings_file')
    if not f:
        return jsonify({'error': 'No file provided'}), 400
    try:
        df     = pd.read_excel(io.BytesIO(f.read()), header=1)
        result = db.preview_incremental(df)
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/admin/upload-meetings', methods=['POST'])
def upload_meetings():
    if not session.get('admin'):
        return jsonify({'error': 'Unauthorized'}), 401
    db = get_meetings_db()
    if not db:
        return jsonify({'error': 'DATABASE_URL not configured'}), 400
    f           = request.files.get('meetings_file')
    upload_type = request.form.get('upload_type', 'incremental')
    if not f:
        return jsonify({'error': 'No file provided'}), 400
    try:
        df = pd.read_excel(io.BytesIO(f.read()), header=1)
        if upload_type == 'full':
            result = db.upload_full(df, f.filename)
        else:
            result = db.upload_incremental(df, f.filename)
        return jsonify(result)
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
