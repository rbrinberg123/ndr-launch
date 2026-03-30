import pandas as pd
import traceback
from typing import Optional

MCAP_MAP = {
    'Micro': '****Micro',
    'Small': '***Small',
    'Mid':   '**Mid',
    'Large': '*Large',
    'Mega':  'Mega',
}

GENERALIST = {'*Generalist'}

EUR_COUNTRIES = {
    'United Kingdom', 'France', 'Germany', 'Netherlands', 'Belgium', 'Switzerland',
    'Sweden', 'Norway', 'Denmark', 'Poland', 'Austria', 'Ireland', 'Spain', 'Italy',
    'Finland', 'Portugal', 'Luxembourg', 'Czech Republic', 'Hungary', 'Romania',
    'Greece', 'Israel', 'Turkey', 'Russia', 'South Africa', 'Jersey', 'Guernsey',
    'Isle of Man', 'Channel Islands', 'Gibraltar', 'Malta', 'Cyprus', 'Bulgaria',
    'Croatia', 'Slovakia', 'Slovenia', 'Estonia', 'Latvia', 'Lithuania',
}

NAM_COUNTRIES = {
    'United States', 'Canada', 'Mexico', 'Puerto Rico', 'Bermuda',
    'Cayman Islands', 'British Virgin Islands', 'Bahamas',
}

RENAME_MAP = {
    'Account Equity Assets Under Management (USD, mm)': 'EAUM ($mm)',
    'Account Reported Total Assets (USD, mm)':          'AUM ($mm)',
    'CDF (Contact): Geography':                         'Geo',
    'CDF (Contact): Industry Focus':                    'Industry',
    'CDF (Contact): Investment Style':                  'Style',
    'CDF (Contact): Market Cap.':                       'Mkt. Cap',
    'CDF (Firm): Coverage':                             'Coverage',
    'Last Meeting':                                     'Last Mtg. w/ Any Co',
    'Primary Institution Type':                         'Type',
    'Contact Investment Center':                        'Investment Ctr',
}

FINAL_COLS = [
    'First Name', 'Last Name', 'CRM Account Name', 'Job Function',
    'Phone', 'Email', 'Coverage',
    'Out1', 'Out2', 'Status', 'CRM Notes',
    'Shares', 'As of', 'Last Mtg. w/ Any Co',
    'Last Mtg btwn Contact & Co', 'Last Mtg btwn firm & Co',
    'Industry', 'Geo', 'Style', 'Mkt. Cap',
    'EAUM ($mm)', 'AUM ($mm)', 'T/O %',
    'City', 'State/Province', 'Country/Territory',
    'Fund Shares', 'Passive or Index Shares', 'Total Funds', 'Passive or Index Funds',
    'Type', 'Activist', 'Investment Ctr',
    'L12M', 'Total', '3rd Party', 'Rose & Co',
    'Match Criteria', 'Match Count', 'Source',
    'Exclusion Reason',
]


# ── CDF matching helpers ──────────────────────────────────────────────────────

def parse_cdf(val):
    if pd.isna(val) or str(val).strip() == '':
        return None
    return {v.strip() for v in str(val).split(',')}


def check_geo(contact_geo_vals, target_geo_set):
    if target_geo_set is None:
        return 'neutral'
    if contact_geo_vals is None:
        return 'neutral'
    if contact_geo_vals & target_geo_set:
        return 'match'
    return 'exclude'


def check_standard(contact_vals, target_set):
    if target_set is None:
        return 'neutral'
    if contact_vals is None:
        return 'neutral'
    if contact_vals & target_set:
        return 'match'
    return 'exclude'


def check_industry(contact_vals, target_set):
    if target_set is None:
        return 'neutral'
    if contact_vals is None:
        return 'neutral'
    if contact_vals & GENERALIST:
        return 'match'
    if contact_vals & target_set:
        return 'match'
    return 'exclude'


def evaluate_contact(row, criteria):
    industry_vals = parse_cdf(row.get('Industry'))
    style_vals    = parse_cdf(row.get('Style'))
    mcap_vals     = parse_cdf(row.get('Mkt. Cap'))
    geo_vals      = parse_cdf(row.get('Geo'))

    results = [
        check_industry(industry_vals, criteria.get('industry')),
        check_standard(style_vals,    criteria.get('style')),
        check_standard(mcap_vals,     criteria.get('mcap')),
        check_geo(geo_vals,           criteria.get('geo')),
    ]

    dim_names    = ['Industry', 'Style', 'Market Cap', 'Geography']
    matched_dims = [n for n, r in zip(dim_names, results) if r == 'match']
    has_match    = any(r == 'match'   for r in results)
    has_exclude  = any(r == 'exclude' for r in results)

    return has_match and not has_exclude, ', '.join(matched_dims)


# ── Investment Center derivation ──────────────────────────────────────────────

def build_inv_center(city, state, country):
    city_s  = str(city).strip()    if pd.notna(city)    else ''
    state_s = str(state).strip()   if pd.notna(state)   else ''
    ctry_s  = str(country).strip() if pd.notna(country) else ''

    CITY_MAP = {
        'Dallas': 'Dallas/Ft. Worth TX', 'Fort Worth': 'Dallas/Ft. Worth TX',
        'Houston': 'Houston TX', 'San Antonio': 'San Antonio TX', 'Austin': 'Dallas/Ft. Worth TX',
        'San Francisco': 'San Francisco/San Jose CA', 'San Jose': 'San Francisco/San Jose CA',
        'Palo Alto': 'San Francisco/San Jose CA', 'Menlo Park': 'San Francisco/San Jose CA',
        'Los Angeles': 'Los Angeles/Pasadena CA', 'Pasadena': 'Los Angeles/Pasadena CA',
        'Irvine': 'Los Angeles/Pasadena CA', 'Costa Mesa': 'Los Angeles/Pasadena CA',
        'Santa Monica': 'Los Angeles/Pasadena CA', 'Newport Beach': 'Los Angeles/Pasadena CA',
        'Kansas City': 'Kansas City MO', 'St Louis': 'Kansas City MO', 'St. Louis': 'Kansas City MO',
        'New York': 'New York/Southern CT/Northern NJ',
        'Purchase': 'New York/Southern CT/Northern NJ',
        'Greenwich': 'New York/Southern CT/Northern NJ',
        'Stamford': 'New York/Southern CT/Northern NJ',
        'Fort Lee': 'New York/Southern CT/Northern NJ',
        'Fort Lauderdale': 'South Florida/Orlando FL/Tampa-St.Pete FL',
        'Miami': 'South Florida/Orlando FL/Tampa-St.Pete FL',
        'Miami Beach': 'South Florida/Orlando FL/Tampa-St.Pete FL',
        'West Palm Beach': 'South Florida/Orlando FL/Tampa-St.Pete FL',
        'Tampa': 'South Florida/Orlando FL/Tampa-St.Pete FL',
        'St. Petersburg': 'South Florida/Orlando FL/Tampa-St.Pete FL',
        'Orlando': 'South Florida/Orlando FL/Tampa-St.Pete FL',
        'Winter Park': 'South Florida/Orlando FL/Tampa-St.Pete FL',
    }

    STATE_MAP = {
        'New York': 'New York/Southern CT/Northern NJ',
        'Connecticut': 'New York/Southern CT/Northern NJ',
        'New Jersey': 'New York/Southern CT/Northern NJ',
        'Massachusetts': 'Boston MA',
        'Illinois': 'Chicago IL', 'Michigan': 'Chicago IL',
        'Wisconsin': 'Chicago IL', 'Indiana': 'Chicago IL',
        'Ohio': 'Columbus OH', 'Kentucky': 'Columbus OH',
        'Pennsylvania': 'Philadelphia PA/Wilmington DE',
        'Delaware': 'Philadelphia PA/Wilmington DE',
        'Maryland': 'Philadelphia PA/Wilmington DE',
        'Virginia': 'Philadelphia PA/Wilmington DE',
        'Minnesota': 'Minneapolis/St. Paul MN',
        'Florida': 'South Florida/Orlando FL/Tampa-St.Pete FL',
        'Georgia': 'Atlanta', 'Tennessee': 'Nashville',
        'Colorado': 'Denver', 'Nebraska': 'Kansas City MO',
        'Arkansas': 'Dallas/Ft. Worth TX',
        'Washington': 'San Francisco/San Jose CA', 'Oregon': 'San Francisco/San Jose CA',
        'Nevada': 'Los Angeles/Pasadena CA',
        'Alberta': 'Toronto', 'Ontario': 'Toronto', 'Quebec': 'Toronto',
        'Texas': None, 'California': None, 'Missouri': None,
    }

    COUNTRY_MAP = {
        'United Kingdom': 'London', 'France': 'Paris',
        'Germany': 'London', 'Netherlands': 'Amsterdam', 'Belgium': 'Amsterdam',
        'Switzerland': 'London', 'Sweden': 'London', 'Norway': 'London',
        'Denmark': 'London', 'Poland': 'London',
        'Japan': 'Tokyo', 'Hong Kong SAR': 'Hong Kong', 'Hong Kong': 'Hong Kong',
        'Australia': 'Sydney', 'New Zealand': 'Sydney',
        'Canada': 'Toronto', 'Puerto Rico': 'San Juan',
    }

    if city_s and city_s in CITY_MAP:
        return CITY_MAP[city_s]
    if state_s and state_s in STATE_MAP and STATE_MAP[state_s] is not None:
        return STATE_MAP[state_s]
    if ctry_s and ctry_s != 'United States' and ctry_s in COUNTRY_MAP:
        return COUNTRY_MAP[ctry_s]

    if city_s and state_s:
        return f'{city_s}, {state_s}'
    elif city_s and ctry_s:
        return f'{city_s}, {ctry_s}'
    elif city_s:
        return city_s
    elif state_s:
        return state_s
    return None


# ── Virtual region classification ─────────────────────────────────────────────

def classify_virtual_region(row):
    """Classify a contact as Virtual - NAM, Virtual - EUR, or Virtual - Other."""
    country = str(row.get('Country/Territory', '') or '').strip()
    ic = str(row.get('Investment Ctr', '') or '').strip()
    if country in NAM_COUNTRIES:
        return 'Virtual - NAM'
    if country in EUR_COUNTRIES:
        return 'Virtual - EUR'
    if not country:
        eur_ic_hints = ('London', 'Paris', 'Amsterdam', 'Stockholm', 'Zurich',
                        'Frankfurt', 'Milan', 'Madrid', 'Oslo', 'Copenhagen')
        if any(h in ic for h in eur_ic_hints):
            return 'Virtual - EUR'
        return 'Virtual - NAM'
    return 'Virtual - Other'


# ── Activities enrichment ─────────────────────────────────────────────────────

def load_activities(acts_df, subject_symbols):
    acts_df = acts_df.copy()
    acts_df['Date'] = pd.to_datetime(acts_df['Date'], errors='coerce')
    upper_symbols = {s.upper() for s in subject_symbols} if isinstance(subject_symbols, list) else {subject_symbols.upper()}
    acts = acts_df[
        acts_df['Symbols'].fillna('').str.strip().str.upper().isin(upper_symbols)
    ].copy()
    acts['_fname'] = acts['External Participant First Name'].fillna('').str.strip().str.lower()
    acts['_lname'] = acts['External Participant Last Name'].fillna('').str.strip().str.lower()
    acts['_inst']  = acts['External Participants (Institutions)'].fillna('').str.strip().str.lower()
    acts_named = acts[(acts['_fname'] != '') & (acts['_lname'] != '')].copy()
    return acts_named


def compute_activity_cols(frame, acts_named, cutoff_l12m):
    frame = frame.reset_index(drop=True)
    if acts_named is None or len(acts_named) == 0:
        for col in ['Last Mtg btwn Contact & Co', 'Last Mtg btwn firm & Co', 'L12M', 'Total', '3rd Party', 'Rose & Co']:
            frame[col] = None
        return frame
    contact_grp = acts_named.groupby(['_fname', '_lname'])
    last_contact = contact_grp['Date'].max()
    total = contact_grp.size()
    l12m = acts_named[acts_named['Date'] >= cutoff_l12m].groupby(['_fname', '_lname']).size()
    topic = acts_named['Topic'].fillna('').str.strip()
    third_party = acts_named[topic == '3rd Party'].groupby(['_fname', '_lname']).size()
    rose = acts_named[topic.isin(['', '*Rose & Company'])].groupby(['_fname', '_lname']).size()
    inst_last = acts_named.groupby('_inst')['Date'].max()
    key_index = pd.MultiIndex.from_tuples(list(zip(frame['_fname'], frame['_lname'])), names=['_fname', '_lname'])
    frame['Last Mtg btwn Contact & Co'] = pd.to_datetime(last_contact.reindex(key_index).values, errors='coerce')
    frame['Last Mtg btwn Contact & Co'] = frame['Last Mtg btwn Contact & Co'].dt.date
    total_vals = total.reindex(key_index).fillna(0).astype(int).values
    l12m_vals = l12m.reindex(key_index).fillna(0).astype(int).values
    tp_vals = third_party.reindex(key_index).fillna(0).astype(int).values
    rc_vals = rose.reindex(key_index).fillna(0).astype(int).values
    frame['Total'] = [int(v) if v > 0 else None for v in total_vals]
    frame['L12M'] = [int(v) if v > 0 else None for v in l12m_vals]
    frame['3rd Party'] = [int(v) if v > 0 else None for v in tp_vals]
    frame['Rose & Co'] = [int(v) if v > 0 else None for v in rc_vals]
    frame['Last Mtg btwn firm & Co'] = pd.to_datetime(inst_last.reindex(frame['_inst'].values).values, errors='coerce')
    frame['Last Mtg btwn firm & Co'] = frame['Last Mtg btwn firm & Co'].dt.date
    return frame


def best_value(person_rows, col):
    if col not in person_rows.columns:
        return None
    vals = person_rows[col].dropna()
    for v in vals:
        if isinstance(v, str) and v.strip() == '':
            continue
        return v
    return None


def build_activity_only_contacts(acts_named, df_contact_keys, cutoff_l12m):
    acts_sorted = acts_named.sort_values('Date', ascending=False)
    acts_only   = acts_sorted[
        acts_sorted.apply(lambda r: (r['_fname'], r['_lname']) not in df_contact_keys, axis=1)
    ]
    unique_people = acts_only.drop_duplicates(subset=['_fname', '_lname'])

    rows = []
    for _, ar in unique_people.iterrows():
        fn, ln = ar['_fname'], ar['_lname']
        person_rows = acts_sorted[(acts_sorted['_fname'] == fn) & (acts_sorted['_lname'] == ln)]

        eaum_raw = best_value(person_rows, 'Equity Assets Under Management')
        ata_raw  = best_value(person_rows, 'Reported Total Assets')
        turnover = best_value(person_rows, 'Turnover')
        eaum_mm  = round(float(eaum_raw) / 1_000_000) if eaum_raw and pd.notna(eaum_raw) else None
        ata_mm   = round(float(ata_raw)  / 1_000_000) if ata_raw  and pd.notna(ata_raw)  else None
        to_pct   = round(float(turnover) / 100, 6)    if turnover and pd.notna(turnover) else None

        city    = best_value(person_rows, 'City')
        state   = best_value(person_rows, 'State/Province')
        country = best_value(person_rows, 'Country/Territory')

        style = best_value(person_rows, 'CDF (Contact): Investment Style')
        rows.append({
            'First Name':       ar['External Participant First Name'],
            'Last Name':        ar['External Participant Last Name'],
            'CRM Account Name': best_value(person_rows, 'External Participants (Institutions)'),
            'Email':            best_value(person_rows, 'Email'),
            'Phone':            best_value(person_rows, 'CRM Phone'),
            'Job Function':     best_value(person_rows, 'Job Function'),
            'City':             city,
            'State/Province':   state,
            'Country/Territory': country,
            'Contact Investment Center': build_inv_center(city, state, country),
            'Coverage':         best_value(person_rows, 'CDF (Firm): Coverage'),
            'EAUM ($mm)':       eaum_mm,
            'AUM ($mm)':        ata_mm,
            'T/O %':            to_pct,
            'Primary Institution Type': 'Hedge Fund' if style == 'Alternative' else None,
            'Industry':         best_value(person_rows, 'CDF (Contact): Industry Focus'),
            'Geo':              best_value(person_rows, 'CDF (Contact): Geography'),
            'Style':            style,
            'Mkt. Cap':         best_value(person_rows, 'CDF (Contact): Market Cap.'),
            'CDF (Contact): Do Not Call':         best_value(person_rows, 'CDF (Contact): Do Not Call'),
            'CDF (Contact): Is Quant?':           best_value(person_rows, 'CDF (Contact): Is Quant?'),
            'CDF (Contact): Invests in Credit/HY': best_value(person_rows, 'CDF (Contact): Invests in Credit/HY'),
            'CDF (Firm): Check before calling':   best_value(person_rows, 'CDF (Firm): Check before calling'),
            '_fname': fn, '_lname': ln,
            '_inst': str(best_value(person_rows, 'External Participants (Institutions)') or '').strip().lower(),
            'Match Criteria': '', 'Match Count': None, 'Source': 'Meeting History',
        })

    if not rows:
        return pd.DataFrame()

    extra_df = pd.DataFrame(rows)
    extra_df  = compute_activity_cols(extra_df, acts_named, cutoff_l12m)
    return extra_df


# ── Split helpers ─────────────────────────────────────────────────────────────

def is_dnc(row):
    c = row.get('CDF (Contact): Do Not Call', None)
    f = row.get('CDF (Firm): Do Not Call', None)
    return (not pd.isna(c) and str(c).strip() != '') or (not pd.isna(f) and str(f).strip() != '')

def is_check(row):
    v = row.get('CDF (Firm): Check before calling', None)
    return not pd.isna(v) and str(v).strip().lower() == 'yes'

def is_quant(row):
    v = row.get('CDF (Contact): Is Quant?', None)
    return not pd.isna(v) and str(v).strip().lower() == 'yes'

def is_activist(row):
    v = row.get('Activist', None)
    return not pd.isna(v) and str(v).strip().lower() == 'often'

def is_fixed_income(row):
    v = row.get('CDF (Contact): Invests in Credit/HY', None)
    return not pd.isna(v) and str(v).strip().lower() == 'yes'


def split_df(df, mask_fn):
    mask = df.apply(mask_fn, axis=1)
    return df[mask].reset_index(drop=True).copy(), df[~mask].reset_index(drop=True).copy()


# ── Reorder + sort ────────────────────────────────────────────────────────────

def reorder(frame):
    if frame is None or len(frame) == 0:
        return frame if frame is not None else pd.DataFrame()
    cols = [c for c in FINAL_COLS if c in frame.columns]
    return frame[cols].copy()


def sort_frame(frame):
    if frame is None or len(frame) == 0:
        return frame if frame is not None else pd.DataFrame()
    return frame.sort_values(
        ['CRM Account Name', 'Last Name'],
        key=lambda s: s.fillna('').str.lower()
    ).reset_index(drop=True)


# ── Main run_filter ───────────────────────────────────────────────────────────

def run_filter(contacts_df, ownership_df, fund_df, acts_named,
               criteria, hf_treatment, meeting_exclusion,
               city_selections, subject_symbols, company_name,
               eaum_min=None, mining_df=None,
               acts_df_raw=None, other_symbols=None,
               shareholder_exclusion='include_all',
               virtual_scope='both'):
    df = contacts_df.copy()
    df = df.reset_index(drop=True)
    df = df.loc[:, ~df.columns.duplicated()]

    # Ownership lookup
    if ownership_df is not None:
        try:
            shares_col = [c for c in ownership_df.columns
                          if 'Shares' in str(c) and 'Change' not in str(c) and 'Post' not in str(c)][0]
            shares_lookup = ownership_df.drop_duplicates(subset='Account Name').set_index('Account Name')[shares_col]
            if 'CRM Account Name' in df.columns:
                crm_idx = df.columns.get_loc('CRM Account Name')
                df.insert(crm_idx + 1, 'Shares', df['Account Name'].map(shares_lookup) if 'Account Name' in df.columns else None)
        except Exception:
            pass

    # % S/O lookup for shareholder exclusion
    so_lookup = {}
    if shareholder_exclusion != 'include_all' and ownership_df is not None:
        try:
            so_col = next((c for c in ownership_df.columns if '% s/o' in str(c).lower()), None)
            if so_col is None:
                so_col = ownership_df.columns[4]
            deduped = ownership_df.drop_duplicates(subset='Account Name')
            so_lookup = dict(zip(
                deduped['Account Name'],
                pd.to_numeric(deduped[so_col], errors='coerce') / 100))
        except Exception:
            pass

    # Fund-level ownership lookup
    if fund_df is not None:
        try:
            sc = [c for c in fund_df.columns
                  if 'Shares' in str(c) and 'Change' not in str(c) and 'Post' not in str(c)][0]
            fa = fund_df.groupby('Account Name').agg(fund_shares=(sc, 'sum'), total_funds=(sc, 'count')).reset_index()
            idx_f = fund_df[fund_df['Dominant Style'] == 'Index'].groupby('Account Name').agg(
                index_shares=(sc, 'sum'), index_funds=(sc, 'count')).reset_index()
            fa = fa.merge(idx_f, on='Account Name', how='left')
            fa[['index_shares', 'index_funds']] = fa[['index_shares', 'index_funds']].fillna(0).astype(int)
            fa = fa.drop_duplicates(subset='Account Name').set_index('Account Name')

            si = df.columns.get_loc('Shares') + 1 if 'Shares' in df.columns else 1
            df.insert(si,     'Fund Shares',             df['Account Name'].map(fa['fund_shares'])     if 'Account Name' in df.columns else None)
            df.insert(si + 1, 'Passive or Index Shares', df['Account Name'].map(fa['index_shares'])    if 'Account Name' in df.columns else None)
            df.insert(si + 2, 'Total Funds',             df['Account Name'].map(fa['total_funds'])     if 'Account Name' in df.columns else None)
            df.insert(si + 3, 'Passive or Index Funds',  df['Account Name'].map(fa['index_funds'])     if 'Account Name' in df.columns else None)
        except Exception:
            pass

    # Rename columns
    df.rename(columns=RENAME_MAP, inplace=True)

    # Divide EAUM and AUM by 1,000,000 (source file values are in raw dollars)
    for col in ('EAUM ($mm)', 'AUM ($mm)'):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce') / 1_000_000

    # Rename T/O % and convert from raw number to percentage
    for old in ['Account Equity % Portfolio Turnover', 'Account Equity % T/O']:
        if old in df.columns:
            df[old] = pd.to_numeric(df[old], errors='coerce') / 100
            df.rename(columns={old: 'T/O %'}, inplace=True)
            break

    # Rename institution type
    if 'Type' in df.columns:
        df['Type'] = df['Type'].replace('Investment Manager-Mutual Fund', 'Mutual fund')

    # Build _fname/_lname/_inst keys for matching
    df['_fname'] = df['First Name'].fillna('').str.strip().str.lower()   if 'First Name' in df.columns else ''
    df['_lname'] = df['Last Name'].fillna('').str.strip().str.lower()    if 'Last Name' in df.columns else ''
    df['_inst']  = df['CRM Account Name'].fillna('').str.strip().str.lower() if 'CRM Account Name' in df.columns else ''

    # Activities enrichment
    cutoff_l12m = pd.Timestamp.today() - pd.DateOffset(months=12)
    if acts_named is not None and len(acts_named) > 0:
        df = compute_activity_cols(df, acts_named, cutoff_l12m)

        # Override institution type to Hedge Fund for contacts whose Style is Alternative
        style_col = 'CDF (Contact): Investment Style'
        if style_col in acts_named.columns:
            alt_keys = set()
            for _, ar in acts_named.iterrows():
                if str(ar.get(style_col, '') or '').strip() == 'Alternative':
                    alt_keys.add((ar['_fname'], ar['_lname']))
            if alt_keys:
                mask = df.apply(lambda r: (r['_fname'], r['_lname']) in alt_keys, axis=1)
                df.loc[mask, 'Type'] = 'Hedge Fund'

    # Derive Investment Ctr for main contacts
    def build_ic_row(row):
        return build_inv_center(row.get('City'), row.get('State/Province'), row.get('Country/Territory'))

    if 'Investment Ctr' not in df.columns:
        df['Investment Ctr'] = df.apply(build_ic_row, axis=1)
    else:
        blank = df['Investment Ctr'].isna() | (df['Investment Ctr'].astype(str).str.strip() == '')
        df.loc[blank, 'Investment Ctr'] = df[blank].apply(build_ic_row, axis=1)

    # Translate mcap values
    if criteria.get('mcap'):
        criteria = {**criteria, 'mcap': {MCAP_MAP.get(v, v) for v in criteria['mcap']}}

    # Apply CDF filter
    match_results = df.apply(lambda row: evaluate_contact(row, criteria), axis=1)
    df['Match Criteria'] = [r[1] for r in match_results]
    df['Match Count']    = df['Match Criteria'].apply(lambda x: len(x.split(', ')) if x else 0)
    filtered = df[[r[0] for r in match_results]].copy()
    filtered = filtered.reset_index(drop=True)
    filtered['Source'] = 'CDF Match'

    # Append activity-only contacts
    if acts_named is not None and len(acts_named) > 0:
        df_contact_keys = set(zip(df['_fname'], df['_lname']))
        extra_df = build_activity_only_contacts(acts_named, df_contact_keys, cutoff_l12m)
        if not extra_df.empty:
            filtered_keys = set(zip(filtered['_fname'], filtered['_lname']))
            extra_new = extra_df[extra_df.apply(
                lambda r: (r['_fname'], r['_lname']) not in filtered_keys, axis=1)]
            filtered = pd.concat([filtered, extra_new], ignore_index=True, sort=False)

    # Append additional list contacts (bypass CDF criteria, subject to other splits)
    if mining_df is not None and len(mining_df) > 0:
        mdf = mining_df.copy()
        mdf.rename(columns=RENAME_MAP, inplace=True)
        mdf = mdf.loc[:, ~mdf.columns.duplicated(keep='first')]

        # Divide EAUM and AUM by 1,000,000 (source file values are in raw dollars)
        for col in ('EAUM ($mm)', 'AUM ($mm)'):
            if col in mdf.columns:
                mdf[col] = pd.to_numeric(mdf[col], errors='coerce') / 1_000_000

        for old in ['Account Equity % Portfolio Turnover', 'Account Equity % T/O']:
            if old in mdf.columns:
                mdf[old] = pd.to_numeric(mdf[old], errors='coerce') / 100
                mdf.rename(columns={old: 'T/O %'}, inplace=True)
                break
        if 'Type' in mdf.columns:
            mdf['Type'] = mdf['Type'].replace('Investment Manager-Mutual Fund', 'Mutual fund')
        mdf['_fname'] = mdf['First Name'].fillna('').str.strip().str.lower() if 'First Name' in mdf.columns else ''
        mdf['_lname'] = mdf['Last Name'].fillna('').str.strip().str.lower()  if 'Last Name' in mdf.columns else ''
        mdf['_inst']  = mdf['CRM Account Name'].fillna('').str.strip().str.lower() if 'CRM Account Name' in mdf.columns else ''
        if 'Investment Ctr' not in mdf.columns:
            mdf['Investment Ctr'] = mdf.apply(
                lambda r: build_inv_center(r.get('City'), r.get('State/Province'), r.get('Country/Territory')), axis=1)
        mdf['Match Criteria'] = ''
        mdf['Match Count'] = None
        mdf['Source'] = 'Additional List'
        filtered_keys = set(zip(filtered['_fname'], filtered['_lname']))
        mdf_new = mdf[mdf.apply(lambda r: (r['_fname'], r['_lname']) not in filtered_keys, axis=1)]
        if len(mdf_new) > 0:
            filtered = pd.concat([filtered, mdf_new], ignore_index=True, sort=False)

    # Add placeholder columns
    for col in ['Out1', 'Out2', 'Status', 'As of', 'Last Mtg. w/ Any Co']:
        if col not in filtered.columns:
            filtered[col] = None

    # CRM Notes: populated from Contact Notes column only
    if 'Contact Notes' in filtered.columns:
        filtered['CRM Notes'] = filtered['Contact Notes'].apply(
            lambda v: str(v).strip() if pd.notna(v) and str(v).strip() != '' else None
        )
    else:
        filtered['CRM Notes'] = None

    main_df = filtered.copy()

    # EAUM minimum filter
    too_small_df = pd.DataFrame()
    if eaum_min is not None and 'EAUM ($mm)' in main_df.columns:
        eaum_vals = pd.to_numeric(main_df['EAUM ($mm)'], errors='coerce')
        below_mask = eaum_vals.notna() & (eaum_vals < eaum_min)
        too_small_df = main_df[below_mask].reset_index(drop=True).copy()
        too_small_df['Exclusion Reason'] = f'EAUM below ${eaum_min:,.0f}M minimum'
        main_df = main_df[~below_mask].reset_index(drop=True).copy()

    # HF split
    hf_df = pd.DataFrame()
    if hf_treatment == 'separate':
        hf_df, main_df = split_df(main_df, lambda r: r.get('Type') == 'Hedge Fund')
        if not hf_df.empty:
            hf_df['Exclusion Reason'] = 'Hedge Fund (separated)'
    elif hf_treatment == 'low_turnover':
        def hf_high_turn(r):
            if r.get('Type') != 'Hedge Fund':
                return False
            try:
                return float(r.get('T/O %', 0) or 0) > 1
            except Exception:
                return False
        hf_df, main_df = split_df(main_df, hf_high_turn)
        if not hf_df.empty:
            hf_df['Exclusion Reason'] = 'Hedge Fund with T/O > 100%'

    # Institution Type split — Broker, Venture Capital, Private Equity → Excluded
    EXCL_TYPES = {'broker', 'venture capital', 'private equity'}
    def is_excl_type(r):
        v = r.get('Type', None)
        return not pd.isna(v) and str(v).strip().lower() in EXCL_TYPES
    inst_type_df, main_df = split_df(main_df, is_excl_type)
    if not inst_type_df.empty:
        inst_type_df['Exclusion Reason'] = 'Institution Type'

    # Fixed Income split — contacts with Invests in Credit/HY = Yes
    fi_df = pd.DataFrame()
    fi_df, main_df = split_df(main_df, is_fixed_income)
    if not fi_df.empty:
        fi_df['Exclusion Reason'] = 'Fixed Income Investor'

    # DNC split
    dnc_df,      main_df = split_df(main_df, is_dnc)
    if not dnc_df.empty:
        dnc_df['Exclusion Reason'] = 'Do Not Call'
    check_df,    main_df = split_df(main_df, is_check)
    if not check_df.empty:
        check_df['Exclusion Reason'] = 'Check before calling'
    quant_df,    main_df = split_df(main_df, is_quant)
    if not quant_df.empty:
        quant_df['Exclusion Reason'] = 'Quant'
    activist_df, main_df = split_df(main_df, is_activist)
    if not activist_df.empty:
        activist_df['Exclusion Reason'] = 'Frequent activist'

    # Shareholder exclusion
    excluded_df = pd.DataFrame()
    if not inst_type_df.empty:
        excluded_df = pd.concat([excluded_df, inst_type_df], ignore_index=True, sort=False)
    if shareholder_exclusion != 'include_all' and so_lookup:
        thresholds = {
            'exclude_all': 0.0, 'gt_001': 0.0001, 'gt_002': 0.0002,
            'gt_003': 0.0003, 'gt_04': 0.004, 'gt_05': 0.005,
        }
        threshold = thresholds.get(shareholder_exclusion)
        if threshold is not None:
            def exceeds_sh(r):
                acct = r.get('Account Name', '')
                val = so_lookup.get(acct)
                if val is None or pd.isna(val):
                    return False
                return val > 0 if shareholder_exclusion == 'exclude_all' else val > threshold
            sh_mask = main_df.apply(exceeds_sh, axis=1)
            sh_excluded = main_df[sh_mask].copy()
            if not sh_excluded.empty:
                sh_excluded['Exclusion Reason'] = 'Exceeds Shareholder Limit'
                excluded_df = pd.concat([excluded_df, sh_excluded], ignore_index=True, sort=False)
            main_df = main_df[~sh_mask].reset_index(drop=True)

    # Meeting history exclusion — works with either activities column name
    mtg_col = None
    for c in ['Last Mtg btwn Contact & Co', 'Specifically with Co.']:
        if c in main_df.columns:
            mtg_col = c
            break
    if meeting_exclusion != 'include_all' and mtg_col:
        if meeting_exclusion == 'exclude_l12m':
            cutoff = pd.Timestamp.today() - pd.DateOffset(months=12)
            exc_mask = main_df[mtg_col].apply(
                lambda d: pd.notna(d) and pd.Timestamp(d) >= cutoff)
            exc_reason = 'Met with company in last 12 months'
        elif meeting_exclusion == 'exclude_l24m':
            cutoff = pd.Timestamp.today() - pd.DateOffset(months=24)
            exc_mask = main_df[mtg_col].apply(
                lambda d: pd.notna(d) and pd.Timestamp(d) >= cutoff)
            exc_reason = 'Met with company in last 24 months'
        else:
            exc_mask = main_df[mtg_col].apply(pd.notna)
            exc_reason = 'Prior meeting with company'
        meeting_excluded = main_df[exc_mask].reset_index(drop=True).copy()
        if not meeting_excluded.empty:
            meeting_excluded['Exclusion Reason'] = exc_reason
            excluded_df = pd.concat([excluded_df, meeting_excluded], ignore_index=True, sort=False)
        main_df = main_df[~exc_mask].reset_index(drop=True).copy()

    # Append other-company activity contacts
    if other_symbols and acts_df_raw is not None and len(acts_df_raw) > 0:
        contact_tickers = {}
        for sym in other_symbols:
            sym_acts = load_activities(acts_df_raw, [sym])
            for _, r in sym_acts.iterrows():
                key = (r['_fname'], r['_lname'])
                contact_tickers.setdefault(key, set()).add(sym)

        if contact_tickers:
            all_keys = set()
            for frame in [main_df, too_small_df, hf_df, fi_df, dnc_df, check_df, quant_df, activist_df, excluded_df]:
                if frame is not None and len(frame) > 0:
                    all_keys.update(zip(frame['_fname'], frame['_lname']))
            all_keys.update(zip(df['_fname'], df['_lname']))

            other_acts = load_activities(acts_df_raw, other_symbols)
            other_extra = build_activity_only_contacts(other_acts, all_keys, cutoff_l12m)
            if not other_extra.empty:
                other_extra.rename(columns=RENAME_MAP, inplace=True)
                for col in ['Last Mtg btwn Contact & Co', 'Last Mtg btwn firm & Co', 'L12M', 'Total', '3rd Party', 'Rose & Co']:
                    if col in other_extra.columns:
                        other_extra[col] = None
                other_extra['Source'] = other_extra.apply(
                    lambda r: 'Other: ' + ', '.join(sorted(contact_tickers.get((r['_fname'], r['_lname']), set()))),
                    axis=1)
                other_new = other_extra[other_extra.apply(
                    lambda r: (r['_fname'], r['_lname']) not in all_keys, axis=1)]
                if len(other_new) > 0:
                    main_df = pd.concat([main_df, other_new], ignore_index=True, sort=False)

    # City routing — split virtual into NAM / EUR / Other sub-tabs
    city_dfs = {}
    if city_selections:
        remaining = main_df.copy()
        ic_col = remaining['Investment Ctr'].fillna('').str.lower()
        for tab_name, ic_value in city_selections:
            ic_lower = ic_value.lower()
            mask = ic_col.str.contains(ic_lower, regex=False)
            if mask.sum() == 0:
                first_segment = ic_lower.split('/')[0].strip()
                mask = ic_col.str.contains(first_segment, regex=False)
            city_dfs[tab_name] = remaining[mask].copy()
            remaining = remaining[~mask].copy()
            ic_col = remaining['Investment Ctr'].fillna('').str.lower()

        # Split unmatched into Virtual - NAM / Virtual - EUR / Virtual - Other
        # then apply virtual_scope filter to overflow tabs
        if len(remaining) > 0:
            remaining = remaining.copy()
            remaining['_vregion'] = remaining.apply(classify_virtual_region, axis=1)
            for vname in ['Virtual - NAM', 'Virtual - EUR', 'Virtual - Other']:
                city_dfs[vname] = remaining[remaining['_vregion'] == vname].drop(columns=['_vregion']).copy()
        # Apply virtual_scope to city-routing overflow tabs
        if virtual_scope == 'nam':
            city_dfs.pop('Virtual - EUR',   None)
            city_dfs.pop('Virtual - Other', None)
        elif virtual_scope == 'eur':
            city_dfs.pop('Virtual - NAM',   None)
            city_dfs.pop('Virtual - Other', None)
        main_df = None

    # Virtual mode — apply virtual_scope filter (NAM / EUR / Both)
    elif virtual_scope and virtual_scope != 'both' and main_df is not None and len(main_df) > 0:
        main_df = main_df.copy()
        main_df['_vregion'] = main_df.apply(classify_virtual_region, axis=1)
        if virtual_scope == 'nam':
            main_df = main_df[main_df['_vregion'] == 'Virtual - NAM'].drop(columns=['_vregion']).reset_index(drop=True)
        elif virtual_scope == 'eur':
            main_df = main_df[main_df['_vregion'] == 'Virtual - EUR'].drop(columns=['_vregion']).reset_index(drop=True)
        else:
            main_df = main_df.drop(columns=['_vregion'], errors='ignore')

    # Reorder + sort all frames
    frames_to_process = {}
    if city_dfs:
        for k, v in city_dfs.items():
            frames_to_process[k] = sort_frame(reorder(v))
    else:
        frames_to_process['Contacts'] = sort_frame(reorder(main_df))

    frames_to_process['Too Small']    = sort_frame(reorder(too_small_df))
    frames_to_process['Fixed Income'] = sort_frame(reorder(fi_df))
    frames_to_process['HFs']          = sort_frame(reorder(hf_df))
    frames_to_process['DNC']          = sort_frame(reorder(dnc_df))
    frames_to_process['Check']        = sort_frame(reorder(check_df))
    frames_to_process['Quant']        = sort_frame(reorder(quant_df))
    frames_to_process['Activist']     = sort_frame(reorder(activist_df))
    frames_to_process['Excluded']     = sort_frame(reorder(excluded_df))

    total_matched = sum(len(v) for k, v in frames_to_process.items()
                        if k not in ('Too Small', 'Fixed Income', 'HFs', 'DNC', 'Check', 'Quant', 'Activist', 'Excluded')
                        and v is not None)

    return {
        'frames':                 frames_to_process,
        'city_selections':        city_selections,
        'total_source':           len(contacts_df),
        'total_matched':          total_matched,
        'has_city_routing':       bool(city_dfs),
        'criteria':               criteria,
        'hf_treatment':           hf_treatment,
        'meeting_exclusion':      meeting_exclusion,
        'shareholder_exclusion':  shareholder_exclusion,
        'eaum_min':               eaum_min,
        'subject_symbols':        subject_symbols,
        'company_name':           company_name,
    }
