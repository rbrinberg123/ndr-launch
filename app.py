import os
import io
import json
import uuid
import tempfile
import pandas as pd
from flask import (Flask, render_template, request, jsonify,
                   send_file, session, redirect, url_for)
from modules.filter import run_filter
from modules.meetings import MeetingsDB, enrich_with_meetings
from modules.sharepoint import SharePointClient
from modules.ai_analysis import analyze_documents
from modules.excel_output import generate_excel

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'change-this-in-production')

ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'admin')
TEMP_DIR = tempfile.mkdtemp()

db = MeetingsDB(os.environ.get('DATABASE_URL', ''))


# ── Helpers ──────────────────────────────────────────────────────────────────

def get_taxonomy():
    try:
        tax = db.load_taxonomy()
        if tax:
            return tax
    except Exception:
        pass
    # Fallback to bundled defaults
    return _default_taxonomy()


def _default_taxonomy():
    return {
        'Geography': [
            {'value': '***Intl ADR', 'description': ''},
            {'value': '**Emerging Markets', 'description': ''},
            {'value': '*Global', 'description': ''},
            {'value': '*Global (ex US)', 'description': ''},
            {'value': 'Africa', 'description': ''},
            {'value': 'Africa \u2013 South Pacific', 'description': ''},
            {'value': 'APAC', 'description': ''},
            {'value': 'Asia Pacific', 'description': ''},
            {'value': 'Asia Pacific: India', 'description': ''},
            {'value': 'Australia', 'description': ''},
            {'value': 'Canada \u2013 North America', 'description': ''},
            {'value': 'Europe', 'description': ''},
            {'value': 'Europe \u2013 Israel', 'description': ''},
            {'value': 'Europe \u2013 Norway', 'description': ''},
            {'value': 'Europe \u2013 UK', 'description': ''},
            {'value': 'Middle East', 'description': ''},
            {'value': 'North America', 'description': ''},
            {'value': 'North America (US-listed only)', 'description': ''},
            {'value': 'South America', 'description': ''},
        ],
        'Market capitalization': [
            {'value': 'Micro', 'description': ''},
            {'value': 'Small', 'description': ''},
            {'value': 'Mid', 'description': ''},
            {'value': 'Large', 'description': ''},
            {'value': 'Mega', 'description': ''},
        ],
        'Investment style': [
            {'value': v, 'description': ''} for v in [
                'Aggressive growth', 'Asset allocator', 'Blend', 'Convertibles',
                'Deep Value', 'Distressed', 'ESG administrator', 'ESG investor',
                'GARP', 'Growth', 'Hedge fund', 'Macro', 'Event-driven',
                'Special situations', 'Real Assets', 'Shariah', 'SPAC',
                'SPAC (pre-merger)', 'Value', 'Wealth Manager', 'Yield',
            ]
        ],
        'Industry Focus': [
            {'value': v, 'description': ''} for v in [
                '*Generalist', 'Agriculture', 'Basic Materials',
                'Basic Materials: Aluminum/Steel', 'Basic Materials: Chemicals',
                'Basic Materials: Construction Materials', 'Basic Materials: Forest Products',
                'Basic Materials: Lithium', 'Basic Materials: Metals & Mining',
                'Basic Materials: Precious Metals', 'Basic Materials: Uranium',
                'Consumer Discretionary: Branded Apparel', 'Consumer Discretionary: Restaurants',
                'Consumer Goods', 'Consumer Goods: Automotive', 'Consumer Goods: Consumer Durables',
                'Consumer Goods: Consumer Non-Durables', 'Consumer Goods: Discretionary',
                'Consumer Goods: Food, Beverage and Tobacco',
                'Consumer Services', 'Consumer Services: Gaming',
                'Consumer Services: Health and Wellness', 'Consumer Services: Homebuilding',
                'Consumer Services: Internet', 'Consumer Services: Media',
                'Consumer Services: Personal Services', 'Consumer Services: Retail',
                'Consumer Services: Transportation Services',
                'Consumer Services: Travel, Services and Leisure', 'Consumer Services: Wholesale',
                'Consumer Staples: Food & Staples Retail',
                'Consumer Staples: Household & Personal Products',
                'Energy', 'Energy: Clean & Renewables', 'Energy: Downstream',
                'Energy: Infrastructure', 'Energy: Midstream', 'Energy: MLP',
                'Energy: Oil, Gas and Coal', 'Energy: Renewable Energy Equipment and Services',
                'Energy: Upstream',
                'Financials', 'Financials: Asset Management', 'Financials: Banking',
                'Financials: BDC', 'Financials: Exchanges', 'Financials: Financial Services',
                'Financials: FinTech', 'Financials: Insurance', 'Financials: Payment',
                'Financials: Real Estate', 'Financials: Real Estate Tech', 'Financials: REITS',
                'Financials: Specialized',
                'Healthcare', 'Healthcare: Biotechnology and Pharmaceuticals',
                'Healthcare: Health Services', 'Healthcare: Healthcare and Supplies Wholesale',
                'Healthcare: Information Technology', 'Healthcare: Medical Equipment',
                'Industrials', 'Industrials: Aerospace and Defense', 'Industrials: Building Products',
                'Industrials: Business Services', 'Industrials: Commercial and Professional Services',
                'Industrials: Conglomerates', 'Industrials: E&C', 'Industrials: Environment Services',
                'Industrials: General Industrials', 'Industrials: Industrial Equipment',
                'Industrials: Industrial Goods and Services', 'Industrials: Marine',
                'Industrials: Materials and Construction', 'Industrials: Road/Rail',
                'Industrials: Staffing', 'Industrials: Transportation',
                'Infrastructure', 'Packaging',
                'Technology', 'Technology: Comm Equip', 'Technology: Computer Software and Services',
                'Technology: Internet', 'Technology: IT Services and Technology',
                'Technology: Semiconductors', 'Technology: Software',
                'Technology: Technology Hardware and Equipment', 'Technology: Telecommunications',
                'Thematic - Clean Energy', 'Thematic - Climate Solutions', 'Thematic - Digital',
                'Thematic - Health and Wellness', 'Thematic - Nutrition', 'Thematic - Security',
                'Thematic - Water', 'Thematic - Global Franchise', 'Thematic - Innovation',
                'Thematic - Mobility', 'Thematic - Pet/Animal Related', 'Thematic - Population',
                'Thematic - Robotics & AI', 'Thematic - SmartCity', 'Thematic - Space Exploration',
                'Thematic - Timber', 'Utilities',
            ]
        ],
    }


# ── Main routes ───────────────────────────────────────────────────────────────

@app.route('/')
def index():
    taxonomy = get_taxonomy()
    sp_configured = False
    try:
        sp_configured = SharePointClient().is_configured()
    except Exception:
        pass
    return render_template('index.html',
                           taxonomy_json=json.dumps(taxonomy),
                           sp_configured=sp_configured)


@app.route('/api/analyze', methods=['POST'])
def analyze():
    files = request.files.getlist('documents')
    if not files or all(f.filename == '' for f in files):
        return jsonify({'error': 'No documents uploaded'}), 400

    taxonomy = get_taxonomy()
    file_data = []
    for f in files:
        data = f.read()
        file_data.append({'name': f.filename, 'data': data, 'type': f.content_type or ''})

    try:
        result = analyze_documents(file_data, taxonomy)
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/run', methods=['POST'])
def run():
    contacts_file   = request.files.get('contacts')
    ownership_file  = request.files.get('ownership')
    fund_file       = request.files.get('fund_ownership')

    if not contacts_file or contacts_file.filename == '':
        return jsonify({'error': 'Contacts file is required'}), 400

    def to_set(vals):
        s = set(v for v in vals if v)
        return s if s else None

    criteria = {
        'industry': to_set(request.form.getlist('industry')),
        'style':    to_set(request.form.getlist('style')),
        'mcap':     to_set(request.form.getlist('mcap')),
        'geo':      to_set(request.form.getlist('geo')),
    }

    hf_treatment = request.form.get('hf_treatment', 'separate')
    company_name = request.form.get('company_name', 'Company').strip() or 'Company'
    ticker       = request.form.get('ticker', '').strip().upper()

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

    try:
        results = run_filter(contacts_df, ownership_df, fund_df, criteria,
                             hf_treatment, company_name, ticker)
    except Exception as e:
        return jsonify({'error': f'Filter error: {e}'}), 500

    # Enrich with meeting history
    if ticker:
        try:
            results = enrich_with_meetings(results, db, ticker)
        except Exception:
            pass

    try:
        excel_bytes = generate_excel(results, company_name)
    except Exception as e:
        return jsonify({'error': f'Excel generation error: {e}'}), 500

    # Save to temp file for download
    file_id  = str(uuid.uuid4())
    filename = f"{company_name} - Relevant Contacts.xlsx"
    tmp_path = os.path.join(TEMP_DIR, f"{file_id}.xlsx")
    with open(tmp_path, 'wb') as f:
        f.write(excel_bytes)
    session['download_id']   = file_id
    session['download_name'] = filename

    # Push to SharePoint if configured
    sharepoint_url = None
    try:
        sp = SharePointClient()
        if sp.is_configured():
            sharepoint_url = sp.upload_file(excel_bytes, filename)
    except Exception:
        pass

    main_df = results['main']
    match_breakdown = {}
    if len(main_df) > 0 and 'Match Count' in main_df.columns:
        for count, grp in main_df.groupby('Match Count'):
            match_breakdown[int(count)] = len(grp)

    return jsonify({
        'total_source':   results['total_source'],
        'main_count':     len(results['main']),
        'hf_count':       len(results.get('hf', pd.DataFrame())),
        'dnc_count':      len(results.get('dnc', pd.DataFrame())),
        'check_count':    len(results.get('check', pd.DataFrame())),
        'quant_count':    len(results.get('quant', pd.DataFrame())),
        'match_breakdown': match_breakdown,
        'sharepoint_url': sharepoint_url,
        'filename':       filename,
    })


@app.route('/download')
def download():
    file_id  = session.get('download_id')
    filename = session.get('download_name', 'Contacts.xlsx')
    if not file_id:
        return 'No file available', 404
    tmp_path = os.path.join(TEMP_DIR, f"{file_id}.xlsx")
    if not os.path.exists(tmp_path):
        return 'File expired — please re-run the filter', 404
    return send_file(
        tmp_path,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


# ── Admin routes ──────────────────────────────────────────────────────────────

@app.route('/admin', methods=['GET', 'POST'])
def admin():
    if request.method == 'POST':
        if request.form.get('password') == ADMIN_PASSWORD:
            session['admin'] = True
        else:
            return render_template('admin.html', authenticated=False, error='Incorrect password')

    if not session.get('admin'):
        return render_template('admin.html', authenticated=False)

    try:
        upload_log = db.get_upload_log()
        total_meetings = db.get_total_meetings()
    except Exception:
        upload_log = []
        total_meetings = 0

    return render_template('admin.html', authenticated=True,
                           upload_log=upload_log,
                           total_meetings=total_meetings)


@app.route('/admin/logout')
def admin_logout():
    session.pop('admin', None)
    return redirect(url_for('admin'))


@app.route('/api/admin/preview-meetings', methods=['POST'])
def preview_meetings():
    if not session.get('admin'):
        return jsonify({'error': 'Unauthorized'}), 401
    f = request.files.get('meetings_file')
    if not f:
        return jsonify({'error': 'No file provided'}), 400
    try:
        df = pd.read_excel(io.BytesIO(f.read()))
        result = db.preview_incremental(df)
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/admin/upload-meetings', methods=['POST'])
def upload_meetings():
    if not session.get('admin'):
        return jsonify({'error': 'Unauthorized'}), 401
    f           = request.files.get('meetings_file')
    upload_type = request.form.get('upload_type', 'incremental')
    if not f:
        return jsonify({'error': 'No file provided'}), 400
    try:
        df = pd.read_excel(io.BytesIO(f.read()))
        if upload_type == 'full':
            result = db.upload_full(df, filename=f.filename)
        else:
            result = db.upload_incremental(df, filename=f.filename)
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/admin/upload-taxonomy', methods=['POST'])
def upload_taxonomy():
    if not session.get('admin'):
        return jsonify({'error': 'Unauthorized'}), 401
    f = request.files.get('taxonomy_file')
    if not f:
        return jsonify({'error': 'No file provided'}), 400
    try:
        df = pd.read_excel(io.BytesIO(f.read()))
        db.save_taxonomy(df)
        return jsonify({'success': True, 'rows': len(df)})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ── Startup ───────────────────────────────────────────────────────────────────

@app.template_filter('format_number')
def format_number(n):
    try:
        return f'{int(n):,}'
    except Exception:
        return n


if __name__ == '__main__':
    try:
        db.init_db()
    except Exception as e:
        print(f'DB init warning: {e}')
    app.run(debug=os.environ.get('FLASK_ENV') == 'development', host='0.0.0.0', port=5000)
