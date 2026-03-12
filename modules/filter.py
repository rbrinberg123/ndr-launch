import pandas as pd
from typing import Optional, Set, Dict, Any

MCAP_MAP = {
    'Micro': '****Micro',
    'Small': '***Small',
    'Mid':   '**Mid',
    'Large': '*Large',
    'Mega':  'Mega',
}

US_ONLY_GEO = {'North America', 'North America (US-listed only)'}
GLOBAL_EX_US = '*Global (ex US)'
GLOBAL = '*Global'
GENERALIST = {'*Generalist'}

COLUMNS_TO_DROP = [
    'Account Name', 'CRM Institution Type', 'Type', 'Asset Class',
    'BD Job Title', 'CRM Job Title', 'Address',
    'CDF (Contact): Coverage', 'CDF (Contact): Do Not Call',
    'CDF (Contact): Exclude from Distribution', 'CDF (Contact): Invest in Gold?',
    'CDF (Contact): Invests in Credit/HY', 'CDF (Contact): Invests in Fixed Income',
    'CDF (Contact): Invests in pre-IPO companies', 'CDF (Contact): Invests in Private Equity',
    'CDF (Contact): Is Quant?', 'CDF (Contact): Junior Mining',
    'CDF (Contact): Junior Mining (Notes)', 'CDF (Contact): No Relevant Mandate',
    'CDF (Contact): Primary Contact', 'CDF (Firm): Check before calling',
    'CDF (Firm): Do Not Call', 'Account Stated Total Assets (USD, mm)',
    'Contact Focus Market Cap (EQ)', 'Investment Center',
    'Dominant Style', 'Dominant Orientation',
]

RENAME_MAP = {
    'Account Equity Assets Under Management (USD, mm)': 'EAUM ($mm)',
    'Account Reported Total Assets (USD, mm)':          'AUM ($mm)',
    'CDF (Contact): Geography':                         'Geo',
    'CDF (Contact): Industry Focus':                    'Industry',
    'CDF (Contact): Investment Style':                  'Style',
    'CDF (Contact): Market Cap.':                       'Mkt. Cap',
}


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

    has_match   = any(r == 'match'   for r in results)
    has_exclude = any(r == 'exclude' for r in results)

    matched_dims = []
    dim_names = ['Industry', 'Style', 'Market Cap', 'Geography']
    for name, result in zip(dim_names, results):
        if result == 'match':
            matched_dims.append(name)

    return has_match and not has_exclude, ', '.join(matched_dims)


def drop_columns(df):
    return df.drop(columns=[c for c in COLUMNS_TO_DROP if c in df.columns])


def is_dnc(row):
    contact_dnc = row.get('CDF (Contact): Do Not Call', None)
    firm_dnc    = row.get('CDF (Firm): Do Not Call', None)
    return (
        (not pd.isna(contact_dnc) and str(contact_dnc).strip() != '') or
        (not pd.isna(firm_dnc) and str(firm_dnc).strip() != '')
    )


def is_check(row):
    val = row.get('CDF (Firm): Check before calling', None)
    return not pd.isna(val) and str(val).strip().lower() == 'yes'


def is_quant(row):
    val = row.get('CDF (Contact): Is Quant?', None)
    return not pd.isna(val) and str(val).strip().lower() == 'yes'


def run_filter(contacts_df, ownership_df, fund_df, criteria, hf_treatment, company_name, ticker):
    df = contacts_df.copy()

    # --- Ownership lookup ---
    if ownership_df is not None:
        try:
            shares_col = [c for c in ownership_df.columns
                          if 'Shares' in str(c) and 'Change' not in str(c) and 'Post' not in str(c)][0]
            shares_lookup = ownership_df.set_index('Account Name')[shares_col]
            crm_idx = df.columns.get_loc('CRM Account Name') if 'CRM Account Name' in df.columns else 0
            df.insert(crm_idx + 1, 'Shares', df['Account Name'].map(shares_lookup) if 'Account Name' in df.columns else None)
        except Exception:
            pass

    # --- Fund-level ownership lookup ---
    if fund_df is not None:
        try:
            shares_col_fund = [c for c in fund_df.columns
                               if 'Shares' in str(c) and 'Change' not in str(c) and 'Post' not in str(c)][0]
            fund_agg = fund_df.groupby('Account Name').agg(
                fund_shares=(shares_col_fund, 'sum'),
                total_funds=(shares_col_fund, 'count'),
            ).reset_index()
            index_fund = fund_df[fund_df['Dominant Style'] == 'Index'].groupby('Account Name').agg(
                index_shares=(shares_col_fund, 'sum'),
                index_funds=(shares_col_fund, 'count'),
            ).reset_index()
            fund_agg = fund_agg.merge(index_fund, on='Account Name', how='left')
            fund_agg[['index_shares', 'index_funds']] = fund_agg[['index_shares', 'index_funds']].fillna(0).astype(int)
            fund_agg = fund_agg.set_index('Account Name')

            shares_idx = df.columns.get_loc('Shares') + 1 if 'Shares' in df.columns else 1
            df.insert(shares_idx,     'Fund Shares',             df['Account Name'].map(fund_agg['fund_shares']) if 'Account Name' in df.columns else None)
            df.insert(shares_idx + 1, 'Passive or Index Shares', df['Account Name'].map(fund_agg['index_shares']) if 'Account Name' in df.columns else None)
            df.insert(shares_idx + 2, 'Total Funds',             df['Account Name'].map(fund_agg['total_funds']) if 'Account Name' in df.columns else None)
            df.insert(shares_idx + 3, 'Passive or Index Funds',  df['Account Name'].map(fund_agg['index_funds']) if 'Account Name' in df.columns else None)
        except Exception:
            pass

    # --- Rename columns ---
    df.rename(columns=RENAME_MAP, inplace=True)

    # --- Translate mcap values ---
    if criteria.get('mcap'):
        criteria['mcap'] = {MCAP_MAP.get(v, v) for v in criteria['mcap']}

    # --- Apply filter ---
    match_results = df.apply(lambda row: evaluate_contact(row, criteria), axis=1)
    df['Match Criteria'] = [r[1] for r in match_results]
    df['Match Count']    = df['Match Criteria'].apply(lambda x: len(x.split(', ')) if x else 0)
    filtered_df = df[[r[0] for r in match_results]].copy()

    # --- HF split ---
    hf_df = pd.DataFrame()
    if hf_treatment == 'separate':
        hf_mask     = filtered_df.get('Primary Institution Type', pd.Series(dtype=str)) == 'Hedge Fund'
        hf_df       = filtered_df[hf_mask].copy()
        filtered_df = filtered_df[~hf_mask].copy()
    elif hf_treatment == 'low_turnover':
        hf_base = filtered_df.get('Primary Institution Type', pd.Series(dtype=str)) == 'Hedge Fund'
        turnover_col = 'Account Equity % Portfolio Turnover'
        if turnover_col in filtered_df.columns:
            high_turn = pd.to_numeric(filtered_df[turnover_col], errors='coerce').fillna(0) > 100
            hf_mask   = hf_base & high_turn
        else:
            hf_mask = hf_base
        hf_df       = filtered_df[hf_mask].copy()
        filtered_df = filtered_df[~hf_mask].copy()

    # --- DNC split ---
    dnc_mask    = filtered_df.apply(is_dnc, axis=1)
    dnc_df      = filtered_df[dnc_mask].copy()
    filtered_df = filtered_df[~dnc_mask].copy()

    # --- Check split ---
    check_mask  = filtered_df.apply(is_check, axis=1)
    check_df    = filtered_df[check_mask].copy()
    filtered_df = filtered_df[~check_mask].copy()

    # --- Quant split ---
    quant_mask  = filtered_df.apply(is_quant, axis=1)
    quant_df    = filtered_df[quant_mask].copy()
    filtered_df = filtered_df[~quant_mask].copy()

    # --- Drop internal columns ---
    for frame in [filtered_df, hf_df, dnc_df, check_df, quant_df]:
        drop_columns(frame)

    return {
        'main':  filtered_df,
        'hf':    hf_df,
        'dnc':   dnc_df,
        'check': check_df,
        'quant': quant_df,
        'total_source': len(contacts_df),
    }
