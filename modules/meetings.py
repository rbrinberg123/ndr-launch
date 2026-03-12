import psycopg2
import psycopg2.extras
import pandas as pd
from datetime import datetime
import json


class MeetingsDB:
    def __init__(self, database_url):
        self.database_url = database_url
        self._conn = None

    def get_conn(self):
        if self._conn is None or self._conn.closed:
            self._conn = psycopg2.connect(self.database_url, sslmode='require')
        return self._conn

    def init_db(self):
        conn = self.get_conn()
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS meetings (
                    id SERIAL PRIMARY KEY,
                    composite_key TEXT UNIQUE,
                    meeting_date DATE,
                    email TEXT,
                    first_name TEXT,
                    last_name TEXT,
                    account_name TEXT,
                    ticker TEXT,
                    raw_data JSONB,
                    created_at TIMESTAMP DEFAULT NOW(),
                    updated_at TIMESTAMP DEFAULT NOW()
                );
                CREATE INDEX IF NOT EXISTS idx_meetings_email ON meetings(email);
                CREATE INDEX IF NOT EXISTS idx_meetings_name_account ON meetings(first_name, last_name, account_name);
                CREATE INDEX IF NOT EXISTS idx_meetings_ticker ON meetings(ticker);
                CREATE INDEX IF NOT EXISTS idx_meetings_date ON meetings(meeting_date);

                CREATE TABLE IF NOT EXISTS upload_log (
                    id SERIAL PRIMARY KEY,
                    upload_type TEXT,
                    filename TEXT,
                    rows_added INTEGER,
                    rows_updated INTEGER,
                    rows_skipped INTEGER,
                    total_rows_after INTEGER,
                    uploaded_at TIMESTAMP DEFAULT NOW()
                );

                CREATE TABLE IF NOT EXISTS taxonomy (
                    id SERIAL PRIMARY KEY,
                    field_type TEXT,
                    value TEXT,
                    description TEXT,
                    updated_at TIMESTAMP DEFAULT NOW()
                );
            """)
            conn.commit()

    def _detect_columns(self, df):
        """Flexibly detect column names regardless of exact casing/spacing."""
        cols = {c.strip().lower(): c for c in df.columns}

        def find(candidates):
            for c in candidates:
                if c.lower() in cols:
                    return cols[c.lower()]
            return None

        return {
            'date':         find(['date', 'meeting date', 'meetingdate', 'meeting_date']),
            'email':        find(['email', 'email address', 'emailaddress']),
            'first_name':   find(['first name', 'firstname', 'first_name']),
            'last_name':    find(['last name', 'lastname', 'last_name']),
            'account_name': find(['account name', 'accountname', 'account_name', 'firm', 'firm name']),
            'ticker':       find(['ticker', 'symbol', 'tick']),
        }

    def _build_composite_key(self, row, col_map):
        parts = []
        for field in ['date', 'first_name', 'last_name', 'account_name', 'ticker']:
            col = col_map.get(field)
            val = str(row[col]).strip().lower() if col and col in row and not pd.isna(row[col]) else ''
            parts.append(val)
        return '|'.join(parts)

    def _df_to_records(self, df):
        col_map = self._detect_columns(df)
        records = []
        for _, row in df.iterrows():
            raw = {k: (None if pd.isna(v) else str(v)) for k, v in row.items()}
            record = {
                'composite_key': self._build_composite_key(row, col_map),
                'meeting_date':  pd.to_datetime(row[col_map['date']], errors='coerce').date() if col_map.get('date') else None,
                'email':         str(row[col_map['email']]).strip().lower() if col_map.get('email') and not pd.isna(row.get(col_map['email'], None)) else None,
                'first_name':    str(row[col_map['first_name']]).strip().lower() if col_map.get('first_name') and not pd.isna(row.get(col_map['first_name'], None)) else None,
                'last_name':     str(row[col_map['last_name']]).strip().lower() if col_map.get('last_name') and not pd.isna(row.get(col_map['last_name'], None)) else None,
                'account_name':  str(row[col_map['account_name']]).strip().lower() if col_map.get('account_name') and not pd.isna(row.get(col_map['account_name'], None)) else None,
                'ticker':        str(row[col_map['ticker']]).strip().upper() if col_map.get('ticker') and not pd.isna(row.get(col_map['ticker'], None)) else None,
                'raw_data':      json.dumps(raw),
            }
            records.append(record)
        return records

    def upload_full(self, df, filename=''):
        records = self._df_to_records(df)
        conn = self.get_conn()
        with conn.cursor() as cur:
            cur.execute("TRUNCATE TABLE meetings RESTART IDENTITY;")
            psycopg2.extras.execute_batch(cur, """
                INSERT INTO meetings (composite_key, meeting_date, email, first_name, last_name, account_name, ticker, raw_data)
                VALUES (%(composite_key)s, %(meeting_date)s, %(email)s, %(first_name)s, %(last_name)s, %(account_name)s, %(ticker)s, %(raw_data)s)
                ON CONFLICT (composite_key) DO UPDATE SET
                    meeting_date = EXCLUDED.meeting_date,
                    email = EXCLUDED.email,
                    ticker = EXCLUDED.ticker,
                    raw_data = EXCLUDED.raw_data,
                    updated_at = NOW()
            """, records, page_size=1000)
            cur.execute("SELECT COUNT(*) FROM meetings;")
            total = cur.fetchone()[0]
            cur.execute("""
                INSERT INTO upload_log (upload_type, filename, rows_added, rows_updated, rows_skipped, total_rows_after)
                VALUES (%s, %s, %s, %s, %s, %s)
            """, ('full', filename, len(records), 0, 0, total))
            conn.commit()
        return {'added': len(records), 'updated': 0, 'skipped': 0, 'total': total}

    def upload_incremental(self, df, filename=''):
        records = self._df_to_records(df)
        conn = self.get_conn()
        added = updated = skipped = 0
        with conn.cursor() as cur:
            for rec in records:
                if not rec['composite_key'].replace('|', ''):
                    skipped += 1
                    continue
                cur.execute("SELECT id FROM meetings WHERE composite_key = %s", (rec['composite_key'],))
                existing = cur.fetchone()
                if existing:
                    cur.execute("""
                        UPDATE meetings SET meeting_date=%(meeting_date)s, email=%(email)s,
                        first_name=%(first_name)s, last_name=%(last_name)s, account_name=%(account_name)s,
                        ticker=%(ticker)s, raw_data=%(raw_data)s, updated_at=NOW()
                        WHERE composite_key=%(composite_key)s
                    """, rec)
                    updated += 1
                else:
                    cur.execute("""
                        INSERT INTO meetings (composite_key, meeting_date, email, first_name, last_name, account_name, ticker, raw_data)
                        VALUES (%(composite_key)s, %(meeting_date)s, %(email)s, %(first_name)s, %(last_name)s, %(account_name)s, %(ticker)s, %(raw_data)s)
                    """, rec)
                    added += 1
            cur.execute("SELECT COUNT(*) FROM meetings;")
            total = cur.fetchone()[0]
            cur.execute("""
                INSERT INTO upload_log (upload_type, filename, rows_added, rows_updated, rows_skipped, total_rows_after)
                VALUES (%s, %s, %s, %s, %s, %s)
            """, ('incremental', filename, added, updated, skipped, total))
            conn.commit()
        return {'added': added, 'updated': updated, 'skipped': skipped, 'total': total}

    def preview_incremental(self, df):
        """Dry-run incremental upload — returns counts without committing."""
        records = self._df_to_records(df)
        conn = self.get_conn()
        added = updated = skipped = 0
        with conn.cursor() as cur:
            for rec in records:
                if not rec['composite_key'].replace('|', ''):
                    skipped += 1
                    continue
                cur.execute("SELECT id FROM meetings WHERE composite_key = %s", (rec['composite_key'],))
                if cur.fetchone():
                    updated += 1
                else:
                    added += 1
        return {'added': added, 'updated': updated, 'skipped': skipped, 'total_in_file': len(records)}

    def get_upload_log(self):
        conn = self.get_conn()
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cur:
            cur.execute("SELECT * FROM upload_log ORDER BY uploaded_at DESC LIMIT 50;")
            rows = cur.fetchall()
        return [dict(r) for r in rows]

    def get_total_meetings(self):
        conn = self.get_conn()
        with conn.cursor() as cur:
            cur.execute("SELECT COUNT(*) FROM meetings;")
            return cur.fetchone()[0]

    def save_taxonomy(self, df):
        conn = self.get_conn()
        with conn.cursor() as cur:
            cur.execute("TRUNCATE TABLE taxonomy;")
            records = []
            for _, row in df.iterrows():
                records.append((
                    str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else '',
                    str(row.iloc[1]).strip() if len(row) > 1 and not pd.isna(row.iloc[1]) else '',
                    str(row.iloc[2]).strip() if len(row) > 2 and not pd.isna(row.iloc[2]) else '',
                ))
            psycopg2.extras.execute_batch(cur,
                "INSERT INTO taxonomy (field_type, value, description) VALUES (%s, %s, %s)",
                records)
            conn.commit()

    def load_taxonomy(self):
        conn = self.get_conn()
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cur:
            cur.execute("SELECT field_type, value, description FROM taxonomy ORDER BY field_type, id;")
            rows = cur.fetchall()
        if not rows:
            return None
        result = {}
        for row in rows:
            ft = row['field_type']
            if ft not in result:
                result[ft] = []
            result[ft].append({'value': row['value'], 'description': row['description']})
        return result


def enrich_with_meetings(results, db, ticker):
    """Add meeting history columns to all result dataframes."""
    ticker_upper = ticker.strip().upper()

    def _enrich_df(df):
        if df is None or len(df) == 0:
            return df

        emails = []
        name_account_pairs = []

        for _, row in df.iterrows():
            email = None
            for col in ['Email', 'email', 'Email Address']:
                if col in row and not pd.isna(row[col]) and str(row[col]).strip():
                    email = str(row[col]).strip().lower()
                    break
            emails.append(email)

            fn = la = acct = ''
            for col in ['First Name', 'first_name']:
                if col in row and not pd.isna(row.get(col)):
                    fn = str(row[col]).strip().lower()
                    break
            for col in ['Last Name', 'last_name']:
                if col in row and not pd.isna(row.get(col)):
                    la = str(row[col]).strip().lower()
                    break
            for col in ['CRM Account Name', 'Account Name', 'account_name']:
                if col in row and not pd.isna(row.get(col)):
                    acct = str(row[col]).strip().lower()
                    break
            name_account_pairs.append((fn, la, acct))

        conn = db.get_conn()
        enriched = []
        with conn.cursor(cursor_factory=psycopg2.extras.DictCursor) as cur:
            for i, (_, row) in enumerate(df.iterrows()):
                email = emails[i]
                fn, la, acct = name_account_pairs[i]

                meeting_row = None
                if email:
                    cur.execute("""
                        SELECT meeting_date, ticker FROM meetings
                        WHERE email = %s
                        ORDER BY meeting_date DESC
                    """, (email,))
                    rows = cur.fetchall()
                    if rows:
                        meeting_row = rows

                if not meeting_row and fn and la and acct:
                    cur.execute("""
                        SELECT meeting_date, ticker FROM meetings
                        WHERE first_name = %s AND last_name = %s AND account_name = %s
                        ORDER BY meeting_date DESC
                    """, (fn, la, acct))
                    rows = cur.fetchall()
                    if rows:
                        meeting_row = rows

                if meeting_row:
                    all_dates = [r['meeting_date'] for r in meeting_row if r['meeting_date']]
                    ticker_dates = [r['meeting_date'] for r in meeting_row
                                    if r['ticker'] and r['ticker'].upper() == ticker_upper and r['meeting_date']]
                    tickers_met = list(dict.fromkeys(
                        r['ticker'] for r in meeting_row if r['ticker']
                    ))

                    last_any = max(all_dates).strftime('%Y-%m-%d') if all_dates else None
                    last_ticker = max(ticker_dates).strftime('%Y-%m-%d') if ticker_dates else None
                    companies = ', '.join(tickers_met) if tickers_met else None
                else:
                    last_ticker = last_any = companies = None

                enriched.append({
                    f'Last Meeting ({ticker_upper})': last_ticker,
                    'Last Meeting (Any)': last_any,
                    'Companies Met': companies,
                })

        enrich_df = pd.DataFrame(enriched, index=df.index)
        return pd.concat([df, enrich_df], axis=1)

    results['main']  = _enrich_df(results['main'])
    results['hf']    = _enrich_df(results['hf'])
    results['dnc']   = _enrich_df(results['dnc'])
    results['check'] = _enrich_df(results['check'])
    results['quant'] = _enrich_df(results['quant'])
    return results
