import pandas as pd

STATE_MAP = {
    'New York':'New York/Southern CT/Northern NJ','Connecticut':'New York/Southern CT/Northern NJ',
    'New Jersey':'New York/Southern CT/Northern NJ','Massachusetts':'Boston MA',
    'Illinois':'Chicago IL','Ohio':'Columbus OH',
    'Florida':'South Florida/Orlando FL/Tampa-St.Pete FL','Georgia':'Atlanta',
    'Pennsylvania':'Philadelphia PA/Wilmington DE','Delaware':'Philadelphia PA/Wilmington DE',
    'Minnesota':'Minneapolis/St. Paul MN','Michigan':'Chicago IL','Wisconsin':'Chicago IL',
    'Indiana':'Chicago IL','Maryland':'Philadelphia PA/Wilmington DE',
    'Virginia':'Philadelphia PA/Wilmington DE','Colorado':'Denver',
    'Washington':'San Francisco/San Jose CA','Oregon':'San Francisco/San Jose CA',
    'Nebraska':'Kansas City MO','Kentucky':'Columbus OH','Tennessee':'Nashville',
    'Arkansas':'Dallas/Ft. Worth TX','Nevada':'Los Angeles/Pasadena CA',
    'Alberta':'Toronto','Ontario':'Toronto','Quebec':'Toronto',
    'Texas':None,'California':None,'Missouri':None,
}
CITY_MAP = {
    'dallas':'Dallas/Ft. Worth TX','fort worth':'Dallas/Ft. Worth TX',
    'houston':'Houston TX','san antonio':'San Antonio TX','austin':'Dallas/Ft. Worth TX',
    'san francisco':'San Francisco/San Jose CA','san jose':'San Francisco/San Jose CA',
    'palo alto':'San Francisco/San Jose CA','menlo park':'San Francisco/San Jose CA',
    'los angeles':'Los Angeles/Pasadena CA','pasadena':'Los Angeles/Pasadena CA',
    'irvine':'Los Angeles/Pasadena CA','costa mesa':'Los Angeles/Pasadena CA',
    'santa monica':'Los Angeles/Pasadena CA','newport beach':'Los Angeles/Pasadena CA',
    'kansas city':'Kansas City MO','st louis':'Kansas City MO','st. louis':'Kansas City MO',
    'new york':'New York/Southern CT/Northern NJ','purchase':'New York/Southern CT/Northern NJ',
    'greenwich':'New York/Southern CT/Northern NJ','stamford':'New York/Southern CT/Northern NJ',
    'fort lee':'New York/Southern CT/Northern NJ',
    'fort lauderdale':'South Florida/Orlando FL/Tampa-St.Pete FL',
    'miami':'South Florida/Orlando FL/Tampa-St.Pete FL',
    'miami beach':'South Florida/Orlando FL/Tampa-St.Pete FL',
    'west palm beach':'South Florida/Orlando FL/Tampa-St.Pete FL',
    'tampa':'South Florida/Orlando FL/Tampa-St.Pete FL',
    'st. petersburg':'South Florida/Orlando FL/Tampa-St.Pete FL',
    'orlando':'South Florida/Orlando FL/Tampa-St.Pete FL',
    'winter park':'South Florida/Orlando FL/Tampa-St.Pete FL',
    'columbus':'Columbus OH',
}
COUNTRY_MAP = {
    'United Kingdom':'London','France':'Paris','Germany':'London',
    'Netherlands':'Amsterdam','Belgium':'Amsterdam','Switzerland':'London',
    'Sweden':'London','Norway':'London','Denmark':'London','Poland':'London',
    'Japan':'Tokyo','Hong Kong SAR':'Hong Kong','Hong Kong':'Hong Kong',
    'Australia':'Sydney','New Zealand':'Sydney','Canada':'Toronto','Puerto Rico':'San Juan',
}


def build_inv_center(city, state, country):
    city_s  = str(city).strip()    if pd.notna(city)    else ''
    state_s = str(state).strip()   if pd.notna(state)   else ''
    ctry_s  = str(country).strip() if pd.notna(country) else ''
    if city_s and city_s.lower() in CITY_MAP:
        return CITY_MAP[city_s.lower()]
    if state_s and state_s in STATE_MAP and STATE_MAP[state_s]:
        return STATE_MAP[state_s]
    if ctry_s and ctry_s != 'United States' and ctry_s in COUNTRY_MAP:
        return COUNTRY_MAP[ctry_s]
    if city_s and state_s: return f'{city_s}, {state_s}'
    if city_s and ctry_s:  return f'{city_s}, {ctry_s}'
    return city_s or state_s or None


def load_activities(file_bytes):
    df = pd.read_excel(file_bytes, header=1)
    df['Date'] = pd.to_datetime(df.get('Date'), errors='coerce')
    return df


def get_symbols(acts_df):
    col = next((c for c in acts_df.columns if str(c).strip().lower() == 'symbols'), None)
    if col is None: return []
    return sorted({s for s in acts_df[col].dropna().astype(str).str.strip().str.upper() if s})


def compute_activity_cols(df, acts_df, subject_symbol):
    sym_col = next((c for c in acts_df.columns if str(c).strip().lower() == 'symbols'), None)
    if sym_col is None:
        for col in ['Specifically with Co.','Anyone at Inst. with Co','L12M','Total','3rd Party','Rose & Co']:
            df[col] = None
        return df, None

    acts = acts_df[acts_df[sym_col].fillna('').astype(str).str.strip().str.upper() == subject_symbol.upper()].copy()
    acts['_fname'] = acts.get('External Participant First Name', pd.Series(dtype=str)).fillna('').str.strip().str.lower()
    acts['_lname'] = acts.get('External Participant Last Name',  pd.Series(dtype=str)).fillna('').str.strip().str.lower()
    acts['_inst']  = acts.get('External Participants (Institutions)', pd.Series(dtype=str)).fillna('').str.strip().str.lower()
    acts_named = acts[(acts['_fname'] != '') & (acts['_lname'] != '')].copy()

    df = df.copy()
    df['_fname'] = df.get('First Name',       pd.Series(dtype=str)).fillna('').str.strip().str.lower()
    df['_lname'] = df.get('Last Name',        pd.Series(dtype=str)).fillna('').str.strip().str.lower()
    df['_inst']  = df.get('CRM Account Name', pd.Series(dtype=str)).fillna('').str.strip().str.lower()

    cutoff = pd.Timestamp.today() - pd.DateOffset(months=12)
    topic_col = next((c for c in acts_named.columns if str(c).strip().lower() == 'topic'), None)

    spec_l, any_l, l12m_l, tot_l, tp_l, rc_l = [], [], [], [], [], []
    for _, row in df.iterrows():
        fn, ln, inst = row['_fname'], row['_lname'], row['_inst']
        ca = acts_named[(acts_named['_fname'] == fn) & (acts_named['_lname'] == ln)]
        ia = acts_named[acts_named['_inst'] == inst] if inst else acts_named.iloc[0:0]
        spec_l.append(ca['Date'].max() if len(ca) else pd.NaT)
        any_l.append(ia['Date'].max()  if len(ia) else pd.NaT)
        l12m = int((ca['Date'] >= cutoff).sum())
        tot  = int(len(ca))
        tp = rc = 0
        if topic_col:
            tp = int((ca[topic_col].fillna('').str.strip() == '3rd Party').sum())
            rc = int(ca[topic_col].apply(lambda t: pd.isna(t) or str(t).strip() in ('','*Rose & Company')).sum())
        l12m_l.append(l12m if l12m else None)
        tot_l.append(tot   if tot   else None)
        tp_l.append(tp     if tp    else None)
        rc_l.append(rc     if rc    else None)

    df['Specifically with Co.']   = [d.date() if pd.notna(d) else None for d in pd.to_datetime(spec_l)]
    df['Anyone at Inst. with Co'] = [d.date() if pd.notna(d) else None for d in pd.to_datetime(any_l)]
    df['L12M']      = l12m_l
    df['Total']     = tot_l
    df['3rd Party'] = tp_l
    df['Rose & Co'] = rc_l
    return df, acts_named


def build_activity_only_contacts(df, acts_named, subject_symbol):
    if acts_named is None or len(acts_named) == 0: return pd.DataFrame()
    sorted_acts = acts_named.sort_values('Date', ascending=False)
    df_keys   = set(zip(df.get('_fname', pd.Series(dtype=str)).fillna(''),
                        df.get('_lname', pd.Series(dtype=str)).fillna('')))
    acts_keys = set(zip(acts_named['_fname'], acts_named['_lname']))
    only_keys = acts_keys - df_keys

    def best(rows, col):
        if col not in rows.columns: return None
        for v in rows[col].dropna():
            if not (isinstance(v, str) and v.strip() == ''): return v
        return None

    rows = []
    for (fn, ln) in only_keys:
        pr = sorted_acts[(sorted_acts['_fname'] == fn) & (sorted_acts['_lname'] == ln)]
        if len(pr) == 0: continue
        eaum = best(pr, 'Equity Assets Under Management')
        ata  = best(pr, 'Reported Total Assets')
        turn = best(pr, 'Turnover')
        city = best(pr, 'City'); state = best(pr, 'State/Province'); country = best(pr, 'Country/Territory')
        style = best(pr, 'CDF (Contact): Investment Style')
        rows.append({
            'First Name': best(pr,'External Participant First Name'),
            'Last Name':  best(pr,'External Participant Last Name'),
            'CRM Account Name': best(pr,'External Participants (Institutions)'),
            'Email': best(pr,'Email'), 'Phone': best(pr,'CRM Phone'),
            'Job Function': best(pr,'Job Function'),
            'City': city, 'State/Province': state, 'Country/Territory': country,
            'Contact Investment Center': build_inv_center(city, state, country),
            'Coverage': best(pr,'CDF (Firm): Coverage'),
            'EAUM ($mm)': round(eaum/1_000_000) if eaum and pd.notna(eaum) else None,
            'AUM ($mm)':  round(ata /1_000_000) if ata  and pd.notna(ata)  else None,
            'T/O %': round(float(turn)*100,1)   if turn and pd.notna(turn) else None,
            'Primary Institution Type': 'Hedge Fund' if style == 'Alternative' else None,
            'Industry': best(pr,'CDF (Contact): Industry Focus'),
            'Geo':      best(pr,'CDF (Contact): Geography'),
            'Style':    style,
            'Mkt. Cap': best(pr,'CDF (Contact): Market Cap.'),
            'CDF (Contact): Do Not Call':       best(pr,'CDF (Contact): Do Not Call'),
            'CDF (Contact): Is Quant?':         best(pr,'CDF (Contact): Is Quant?'),
            'CDF (Firm): Check before calling': best(pr,'CDF (Firm): Check before calling'),
            '_fname': fn, '_lname': ln,
            '_inst': str(best(pr,'External Participants (Institutions)') or '').strip().lower(),
            'Match Criteria':'', 'Match Count': None, 'Source':'Meeting History',
        })
    return pd.DataFrame(rows) if rows else pd.DataFrame()
