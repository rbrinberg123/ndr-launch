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
    # Filter by symbol FIRST, then copy only the matching subset — avoids copying the full DataFrame
    upper_symbols = {s.upper() for s in subject_symbols} if isinstance(subject_symbols, list) else {subject_symbols.upper()}
    sym_mask = acts_df['Symbols'].fillna('').str.strip().str.upper().isin(upper_symbols)
    acts = acts_df[sym_mask].copy()  # copy only the filtered subset
    acts['Date'] = pd.to_datetime(acts['Date'], errors='coerce')
    acts['_fname'] = acts['External Participant First Name'].fillna('').str.strip().str.lower()
    acts['_lname'] = acts['External Participant Last Name'].fillna('').str.strip().str.lower()
    acts['_inst']  = acts['External Participants (Institutions)'].fillna('').str.strip().str.lower()
    acts_named = acts[(acts['_fname'] != '') & (acts['_lname'] != '')].copy()
    del acts
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


def build_activity_only_contacts(acts_named, df_contact_keys, cutoff_l12m, compute_meetings=True):
    acts_sorted = acts_named.sort_values('Date', ascending=False).reset_index(drop=True)

    # Vectorized exclusion: build a key column and use isin instead of row-wise apply
    acts_sorted['_key'] = list(zip(acts_sorted['_fname'], acts_sorted['_lname']))
    acts_only = acts_sorted[~acts_sorted['_key'].isin(df_contact_keys)].copy()
    acts_only = acts_only.drop(columns=['_key'])
    acts_sorted = acts_sorted.drop(columns=['_key'])

    if acts_only.empty:
        return pd.DataFrame()

    unique_people = acts_only.drop_duplicates(subset=['_fname', '_lname'])[['_fname', '_lname']].copy()
    del acts_only

    # Vectorized aggregation: group acts_sorted by (_fname, _lname) once
    grp = acts_sorted.groupby(['_fname', '_lname'], sort=False)
    # Note: do NOT del acts_sorted here — grp holds a reference to it,
    # so deletion doesn't free memory and causes UnboundLocalError in Python 3.12.
    # Memory is freed correctly when del grp runs after aggregation.

    def first_valid(series):
        for v in series:
            if v is not None and not (isinstance(v, float) and pd.isna(v)):
                if not (isinstance(v, str) and v.strip() == ''):
                    return v
        return None

    # Build per-person aggregated columns via groupby + first-valid
    agg_cols = {
        'External Participant First Name': 'first',
        'External Participant Last Name':  'first',
    }
    text_cols = [
        'External Participants (Institutions)', 'Email', 'CRM Phone', 'Job Function',
        'City', 'State/Province', 'Country/Territory', 'CDF (Firm): Coverage',
        'CDF (Contact): Investment Style', 'CDF (Contact): Industry Focus',
        'CDF (Contact): Geography', 'CDF (Contact): Market Cap.',
        'CDF (Contact): Do Not Call', 'CDF (Contact): Is Quant?',
        'CDF (Contact): Invests in Credit/HY', 'CDF (Firm): Check before calling',
        'Equity Assets Under Management', 'Reported Total Assets', 'Turnover',
    ]
    for col in text_cols:
        if col in acts_sorted.columns:
            agg_cols[col] = first_valid

    agg = grp.agg(agg_cols).reset_index()
    del grp

    # Merge unique_people with aggregated data
    result = unique_people.merge(agg, on=['_fname', '_lname'], how='left')
    del agg, unique_people

    # Vectorized construction — rename and derive columns directly on result DataFrame
    result = result.rename(columns={
        'External Participant First Name':  'First Name',
        'External Participant Last Name':   'Last Name',
        'External Participants (Institutions)': 'CRM Account Name',
        'CRM Phone':                        'Phone',
        'CDF (Contact): Industry Focus':    'Industry',
        'CDF (Contact): Geography':         'Geo',
        'CDF (Contact): Investment Style':  'Style',
        'CDF (Contact): Market Cap.':       'Mkt. Cap',
        'CDF (Firm): Coverage':             'Coverage',
    })

    # Numeric conversions (vectorized)
    for raw_col, out_col, divisor in [
        ('Equity Assets Under Management', 'EAUM ($mm)', 1_000_000),
        ('Reported Total Assets',          'AUM ($mm)',  1_000_000),
        ('Turnover',                       'T/O %',      100),
    ]:
        if raw_col in result.columns:
            result[out_col] = pd.to_numeric(result[raw_col], errors='coerce') / divisor
            result = result.drop(columns=[raw_col])
        else:
            result[out_col] = None

    # Investment Center — must remain row-wise (complex lookup logic)
    result['Contact Investment Center'] = result.apply(
        lambda r: build_inv_center(r.get('City'), r.get('State/Province'), r.get('Country/Territory')),
        axis=1
    )

    # Derived columns
    style_s = result['Style'].fillna('') if 'Style' in result.columns else pd.Series([''] * len(result))
    result['Primary Institution Type'] = style_s.apply(
        lambda v: 'Hedge Fund' if str(v).strip() == 'Alternative' else None
    )
    inst_s = result['CRM Account Name'] if 'CRM Account Name' in result.columns else pd.Series([''] * len(result))
    result['_inst'] = inst_s.fillna('').str.strip().str.lower()
    result['Match Criteria'] = ''
    result['Match Count']    = None
    result['Source']         = 'Meeting History'

    if result.empty:
        return pd.DataFrame()

    if compute_meetings:
        extra_df = compute_activity_cols(result, acts_named, cutoff_l12m)
    else:
        # Skip meeting history computation — caller will set meeting cols to None
        for col in ['Last Mtg btwn Contact & Co', 'Last Mtg btwn firm & Co',
                    'L12M', 'Total', '3rd Party', 'Rose & Co']:
            result[col] = None
        extra_df = result
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
            alt_mask = acts_named[style_col].fillna('').str.strip() == 'Alternative'
            alt_keys = set(zip(acts_named.loc[alt_mask, '_fname'], acts_named.loc[alt_mask, '_lname']))
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
        del df_contact_keys
        if not extra_df.empty:
            filtered_keys = set(zip(filtered['_fname'], filtered['_lname']))
            key_s = pd.Series(list(zip(extra_df['_fname'], extra_df['_lname'])))
            extra_new = extra_df[~key_s.isin(filtered_keys).values]
            del extra_df, key_s, filtered_keys
            filtered = pd.concat([filtered, extra_new], ignore_index=True, sort=False)
            del extra_new
    del acts_named  # no longer needed after activity enrichment + activity-only contacts
    del df

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

    main_df = filtered
    del filtered

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
        # Vectorized: load all other-symbol activities at once, then build contact_tickers via groupby
        other_acts_all = load_activities(acts_df_raw, other_symbols)
        if not other_acts_all.empty:
            # Map each (fname, lname) → set of tickers they were met under
            sym_col = acts_df_raw.columns[acts_df_raw.columns.str.strip().str.lower() == 'symbols']
            if len(sym_col) > 0:
                sym_col_name = sym_col[0]
            else:
                sym_col_name = 'Symbols'

            upper_other = {s.upper() for s in other_symbols}
            raw_other = acts_df_raw[
                acts_df_raw[sym_col_name].fillna('').str.strip().str.upper().isin(upper_other)
            ].copy()
            raw_other['_fname'] = raw_other['External Participant First Name'].fillna('').str.strip().str.lower()
            raw_other['_lname'] = raw_other['External Participant Last Name'].fillna('').str.strip().str.lower()
            raw_other['_sym']   = raw_other[sym_col_name].fillna('').str.strip().str.upper()
            raw_other = raw_other[(raw_other['_fname'] != '') & (raw_other['_lname'] != '')]

            contact_tickers = {}
            for row in raw_other[['_fname', '_lname', '_sym']].itertuples(index=False):
                contact_tickers.setdefault((row[0], row[1]), set()).add(row[2])
            del raw_other

            if contact_tickers:
                all_keys = set()
                for frame in [main_df, too_small_df, hf_df, fi_df, dnc_df, check_df, quant_df, activist_df, excluded_df]:
                    if frame is not None and len(frame) > 0:
                        all_keys.update(zip(frame['_fname'], frame['_lname']))
                all_keys.update(zip(df['_fname'], df['_lname']))

                other_extra = build_activity_only_contacts(other_acts_all, all_keys, cutoff_l12m, compute_meetings=False)
                if not other_extra.empty:
                    other_extra.rename(columns=RENAME_MAP, inplace=True)
                    other_extra['Source'] = other_extra.apply(
                        lambda r: 'Other: ' + ', '.join(sorted(contact_tickers.get((r['_fname'], r['_lname']), set()))),
                        axis=1)
                    # Dedup: fast set-based filter (all_keys already built above)
                    key_s = pd.Series(list(zip(other_extra['_fname'], other_extra['_lname'])))
                    other_new = other_extra[~key_s.isin(all_keys).values].copy()
                    if len(other_new) > 0:
                        main_df = pd.concat([main_df, other_new], ignore_index=True, sort=False)

    if acts_df_raw is not None:
        del acts_df_raw

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
