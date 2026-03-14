import pandas as pd
from typing import Optional

MCAP_MAP = {
    'Micro': '****Micro',
    'Small': '***Small',
    'Mid':   '**Mid',
    'Large': '*Large',
    'Mega':  'Mega',
}

US_ONLY_GEO  = {'North America', 'North America (US-listed only)'}
GLOBAL_EX_US = '*Global (ex US)'
GLOBAL       = '*Global'
GENERALIST   = {'*Generalist'}

RENAME_MAP = {
    'Account Equity Assets Under Management (USD, mm)': 'EAUM ($mm)',
    'Account Reported Total Assets (USD, mm)':          'AUM ($mm)',
    'CDF (Contact): Geography':                         'Geo',
    'CDF (Contact): Industry Focus':                    'Industry',
    'CDF (Contact): Investment Style':                  'Style',
    'CDF (Contact): Market Cap.':                       'Mkt. Cap',
    'CDF (Firm): Coverage':                             'Coverage',
}

FINAL_COLS = [
    'First Name', 'Last Name', 'CRM Account Name', 'Job Function',
    'Phone', 'Email', 'Coverage',
    'Out1', 'Out2', 'Status', 'Notes', 'Contact Notes',
    'Shares', 'As of', 'Last Meeting',
    'Specifically with Co.', 'Anyone at Inst. with Co',
    'Industry', 'Geo', 'Style', 'Mkt. Cap',
    'EAUM ($mm)', 'AUM ($mm)', 'T/O %',
    'City', 'State/Province', 'Country/Territory',
    'Fund Shares', 'Passive or Index Shares', 'Total Funds', 'Passive or Index Funds',
    'Primary Institution Type', 'Activist', 'Contact Investment Center',
    'L12M', 'Total', '3rd Party', 'Rose & Co',
    'Match Criteria', 'Match Count', 'Source',
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
    if contact_geo_vals & US_ONLY_GEO:
        return 'match'
    if GLOBAL in contact_geo_vals and GLOBAL_EX_US not in contact_geo_vals:
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
    specifically, anyone, l12m_vals, total_vals, tp_vals, rc_vals = [], [], [], [], [], []

    for _, row in frame.iterrows():
        fn   = row.get('_fname', '')
        ln   = row.get('_lname', '')
        inst = row.get('_inst', '')

        c_acts = acts_named[(acts_named['_fname'] == fn) & (acts_named['_lname'] == ln)]
        i_acts = acts_named[acts_named['_inst'] == inst] if inst else acts_named.iloc[0:0]

        specifically.append(c_acts['Date'].max() if len(c_acts) > 0 else pd.NaT)
        anyone.append(i_acts['Date'].max() if len(i_acts) > 0 else pd.NaT)

        l12m  = int((c_acts['Date'] >= cutoff_l12m).sum())
        total = len(c_acts)
        tp    = int((c_acts['Topic'].fillna('').str.strip() == '3rd Party').sum())
        rc    = int(c_acts['Topic'].apply(
            lambda t: pd.isna(t) or str(t).strip() == '' or str(t).strip() == '*Rose & Company'
        ).sum())

        l12m_vals.append(l12m  if l12m  > 0 else None)
        total_vals.append(total if total > 0 else None)
        tp_vals.append(tp       if tp    > 0 else None)
        rc_vals.append(rc       if rc    > 0 else None)

    frame = frame.copy()
    frame['Specifically with Co.']   = pd.to_datetime(pd.Series(specifically, index=frame.index)).dt.date
    frame['Anyone at Inst. with Co'] = pd.to_datetime(pd.Series(anyone, index=frame.index)).dt.date
    frame['L12M']      = l12m_vals
    frame['Total']     = total_vals
    frame['3rd Party'] = tp_vals
    frame['Rose & Co'] = rc_vals
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
        to_pct   = round(float(turnover) * 100, 1)    if turnover and pd.notna(turnover) else None

        city    = best_value(person_rows, 'City')
        state   = best_value(person_rows, 'State/Province')
        country = best_value(person_rows, 'Country/Territory')

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
            'Industry':         best_value(person_rows, 'CDF (Contact): Industry Focus'),
            'Geo':              best_value(person_rows, 'CDF (Contact): Geography'),
            'Style':            best_value(person_rows, 'CDF (Contact): Investment Style'),
            'Mkt. Cap':         best_value(person_rows, 'CDF (Contact): Market Cap.'),
            'CDF (Contact): Do Not Call':       best_value(person_rows, 'CDF (Contact): Do Not Call'),
            'CDF (Contact): Is Quant?':         best_value(person_rows, 'CDF (Contact): Is Quant?'),
            'CDF (Firm): Check before calling': best_value(person_rows, 'CDF (Firm): Check before calling'),
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


def split_df(df, mask_fn):
    mask = df.apply(mask_fn, axis=1)
    return df[mask].copy(), df[~mask].copy()


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
               city_selections, subject_symbols, company_name):
    df = contacts_df.copy()

    # Ownership lookup
    if ownership_df is not None:
        try:
            shares_col = [c for c in ownership_df.columns
                          if 'Shares' in str(c) and 'Change' not in str(c) and 'Post' not in str(c)][0]
            shares_lookup = ownership_df.set_index('Account Name')[shares_col]
            if 'CRM Account Name' in df.columns:
                crm_idx = df.columns.get_loc('CRM Account Name')
                df.insert(crm_idx + 1, 'Shares', df['Account Name'].map(shares_lookup) if 'Account Name' in df.columns else None)
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
            fa = fa.set_index('Account Name')

            si = df.columns.get_loc('Shares') + 1 if 'Shares' in df.columns else 1
            df.insert(si,     'Fund Shares',             df['Account Name'].map(fa['fund_shares'])     if 'Account Name' in df.columns else None)
            df.insert(si + 1, 'Passive or Index Shares', df['Account Name'].map(fa['index_shares'])    if 'Account Name' in df.columns else None)
            df.insert(si + 2, 'Total Funds',             df['Account Name'].map(fa['total_funds'])     if 'Account Name' in df.columns else None)
            df.insert(si + 3, 'Passive or Index Funds',  df['Account Name'].map(fa['index_funds'])     if 'Account Name' in df.columns else None)
        except Exception:
            pass

    # Rename columns
    df.rename(columns=RENAME_MAP, inplace=True)

    # Rename T/O %
    for old in ['Account Equity % Portfolio Turnover', 'Account Equity % T/O']:
        if old in df.columns:
            df.rename(columns={old: 'T/O %'}, inplace=True)
            break

    # Rename institution type
    if 'Primary Institution Type' in df.columns:
        df['Primary Institution Type'] = df['Primary Institution Type'].replace(
            'Investment Manager-Mutual Fund', 'Mutual fund')

    # Build _fname/_lname/_inst keys for matching
    df['_fname'] = df['First Name'].fillna('').str.strip().str.lower()   if 'First Name' in df.columns else ''
    df['_lname'] = df['Last Name'].fillna('').str.strip().str.lower()    if 'Last Name' in df.columns else ''
    df['_inst']  = df['CRM Account Name'].fillna('').str.strip().str.lower() if 'CRM Account Name' in df.columns else ''

    # Activities enrichment
    cutoff_l12m = pd.Timestamp.today() - pd.DateOffset(months=12)
    if acts_named is not None and len(acts_named) > 0:
        df = compute_activity_cols(df, acts_named, cutoff_l12m)

    # Derive Contact Investment Center for main contacts
    def build_ic_row(row):
        return build_inv_center(row.get('City'), row.get('State/Province'), row.get('Country/Territory'))

    if 'Contact Investment Center' not in df.columns:
        df['Contact Investment Center'] = df.apply(build_ic_row, axis=1)
    else:
        blank = df['Contact Investment Center'].isna() | (df['Contact Investment Center'].astype(str).str.strip() == '')
        df.loc[blank, 'Contact Investment Center'] = df[blank].apply(build_ic_row, axis=1)

    # Translate mcap values
    if criteria.get('mcap'):
        criteria = {**criteria, 'mcap': {MCAP_MAP.get(v, v) for v in criteria['mcap']}}

    # Apply CDF filter
    match_results = df.apply(lambda row: evaluate_contact(row, criteria), axis=1)
    df['Match Criteria'] = [r[1] for r in match_results]
    df['Match Count']    = df['Match Criteria'].apply(lambda x: len(x.split(', ')) if x else 0)
    filtered = df[[r[0] for r in match_results]].copy()
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

    # Add placeholder columns
    for col in ['Out1', 'Out2', 'Status', 'Notes', 'Contact Notes', 'As of', 'Last Meeting']:
        if col not in filtered.columns:
            filtered[col] = None

    main_df = filtered.copy()

    # HF split
    hf_df = pd.DataFrame()
    if hf_treatment == 'separate':
        hf_df, main_df = split_df(main_df, lambda r: r.get('Primary Institution Type') == 'Hedge Fund')
    elif hf_treatment == 'low_turnover':
        def hf_high_turn(r):
            if r.get('Primary Institution Type') != 'Hedge Fund':
                return False
            try:
                return float(r.get('T/O %', 0) or 0) > 100
            except Exception:
                return False
        hf_df, main_df = split_df(main_df, hf_high_turn)

    # DNC split
    dnc_df,      main_df = split_df(main_df, is_dnc)
    check_df,    main_df = split_df(main_df, is_check)
    quant_df,    main_df = split_df(main_df, is_quant)
    activist_df, main_df = split_df(main_df, is_activist)

    # Meeting history exclusion
    excluded_df = pd.DataFrame()
    if meeting_exclusion != 'include_all' and 'Specifically with Co.' in main_df.columns:
        if meeting_exclusion == 'exclude_l12m':
            cutoff = pd.Timestamp.today() - pd.DateOffset(months=12)
            exc_mask = main_df['Specifically with Co.'].apply(
                lambda d: pd.notna(d) and pd.Timestamp(d) >= cutoff)
        elif meeting_exclusion == 'exclude_l24m':
            cutoff = pd.Timestamp.today() - pd.DateOffset(months=24)
            exc_mask = main_df['Specifically with Co.'].apply(
                lambda d: pd.notna(d) and pd.Timestamp(d) >= cutoff)
        else:  # exclude_all
            exc_mask = main_df['Specifically with Co.'].apply(pd.notna)
        excluded_df = main_df[exc_mask].copy()
        main_df     = main_df[~exc_mask].copy()

    # City routing
    city_dfs = {}
    if city_selections:
        remaining = main_df.copy()
        ic_col = remaining['Contact Investment Center'].fillna('').str.lower()
        for tab_name, ic_value in city_selections:
            ic_lower = ic_value.lower()
            # Try full match first, then first segment before '/' for flexibility
            mask = ic_col.str.contains(ic_lower, regex=False)
            if mask.sum() == 0:
                first_segment = ic_lower.split('/')[0].strip()
                mask = ic_col.str.contains(first_segment, regex=False)
            city_dfs[tab_name] = remaining[mask].copy()
            remaining = remaining[~mask].copy()
            ic_col = remaining['Contact Investment Center'].fillna('').str.lower()
        city_dfs['Virtual'] = remaining
        main_df = None

    # Reorder + sort all frames
    frames_to_process = {}
    if city_dfs:
        for k, v in city_dfs.items():
            frames_to_process[k] = sort_frame(reorder(v))
    else:
        frames_to_process['Contacts'] = sort_frame(reorder(main_df))

    frames_to_process['HFs']      = sort_frame(reorder(hf_df))
    frames_to_process['DNC']      = sort_frame(reorder(dnc_df))
    frames_to_process['Check']    = sort_frame(reorder(check_df))
    frames_to_process['Quant']    = sort_frame(reorder(quant_df))
    frames_to_process['Activist'] = sort_frame(reorder(activist_df))
    frames_to_process['Excluded'] = sort_frame(reorder(excluded_df))

    total_matched = sum(len(v) for k, v in frames_to_process.items()
                        if k not in ('HFs', 'DNC', 'Check', 'Quant', 'Activist', 'Excluded')
                        and v is not None)

    return {
        'frames':          frames_to_process,
        'city_selections': city_selections,
        'total_source':    len(contacts_df),
        'total_matched':   total_matched,
        'has_city_routing': bool(city_dfs),
    }
