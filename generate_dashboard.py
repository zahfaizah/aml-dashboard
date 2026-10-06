"""AML Dashboard generator.

Each team publishes independently:

    python generate_dashboard.py --team sanction --publish
    python generate_dashboard.py --team nontms   --publish
    python generate_dashboard.py --team tms      --publish

A team run reads ONLY that team's Excel and writes data/<team>.json. The
dashboard (index.html) is then rebuilt from all three data/*.json files, so the
other teams' published data is kept exactly as it is on GitHub.

Without --publish, nothing is pushed: data/<team>.json and index.html are
updated locally so you can open index.html and check it.
"""
import argparse
import getpass
import json
import os
import re
import subprocess
import sys
from datetime import datetime

try:
    import pandas as pd
    import numpy as np
    import openpyxl  # noqa: F401  (used by the NON TMS / TMS readers)
except ImportError as e:
    print(f"ERROR: missing Python package ({e.name}). Run:  pip install -r requirements.txt")
    sys.exit(1)

# ─── CONFIG ───────────────────────────────────────────────
# HERE = the folder this script runs from (your Excels live here).
# REPO = the GitHub folder (data/, template, index.html). Usually the same folder;
# a PIC may also run a copy of this script from their own work folder, e.g.
# OneDrive, and the script then uses the GitHub folder in Documents.
HERE = os.path.dirname(os.path.abspath(__file__))

# Each team's Excel file. The script looks for it next to this script, unless
# excel_paths.local.json (not uploaded to GitHub) points somewhere else, e.g.
#   {"sanction": "C:/Users/me/OneDrive/AML/SS_Dashboard_Data_2026.xlsx"}
TEAMS = {
    # Team 1: Sanction Screening (Screening, Remittance, Newsletter, DTTOT, Sub-Merchant)
    'sanction': {'label': 'Sanction', 'excel': 'SS_Dashboard_Data_2026.xlsx'},
    # Team 2: NON TMS
    'nontms':   {'label': 'NON TMS',  'excel': 'NON_TMS_Dashboard_Data_2026.xlsx'},
    # Team 3: TMS (all sub-pages)
    'tms':      {'label': 'TMS',      'excel': 'TMS_Dashboard_Data_2026.xlsx'},
}
LOCAL_PATHS_FILE = os.path.join(HERE, 'excel_paths.local.json')


def _find_repo():
    if os.path.isdir(os.path.join(HERE, '.git')):
        return HERE
    try:
        with open(LOCAL_PATHS_FILE, encoding='utf-8') as f:
            custom = json.load(f).get('repo')
        if custom: return custom
    except (OSError, ValueError):
        pass
    return os.path.join(os.path.expanduser('~'), 'Documents', 'aml-dashboard')


REPO = _find_repo()
DATA_DIR = os.path.join(REPO, 'data')

# TMS → RFI Merchant: sheet inside the TMS Excel, same columns as the 'RFI KYB' sheet
#   Month | Success | Failed | On Process | Additional Notes
# If the sheet doesn't exist yet, the tab shows "No RFI Merchant data."
RFI_MERCHANT_SHEET = "RFI Merchant"

# Template and output (index.html is what GitHub Pages serves)
TEMPLATE_FILE  = os.path.join(REPO, "AML_Dashboard_template.html")
OUTPUT_FILE    = os.path.join(REPO, "index.html")
# When run from a work folder outside the repo, the latest dashboard is also
# copied there under this name so it can be opened from that folder.
LOCAL_COPY_NAME = "AML_Dashboard.html"

# "latest" = auto-detect, or set e.g. "April 2026"
HIGHLIGHT_MONTH = "latest"

# Per-section notes (optional)
NOTES = {
    # "remittance": "Note text here if needed.",
}
# ──────────────────────────────────────────────────────────

MONTH_ORDER = ['January','February','March','April','May','June',
               'July','August','September','October','November','December']

def safe_int(v, default=0):
    try:
        if pd.isna(v): return default
        return int(float(v))
    except: return default

def month_sort_key(label):
    parts = label.strip().split()
    if len(parts) == 2 and parts[0] in MONTH_ORDER:
        return int(parts[1]) * 100 + MONTH_ORDER.index(parts[0])
    return 0


# ═══════════════════════════════════════════════════════════
# SANCTION READERS
# ═══════════════════════════════════════════════════════════

def _cell_txt(v):
    """Normalised text of a cell ('' when blank)."""
    if pd.isna(v): return ''
    return str(v).strip()


def find_block_header(df, col, matcher, start_row=0):
    """Locate a block by its header labels instead of a fixed row number.

    Scans column `col` for a cell reading 'Month' whose neighbour in col+1
    satisfies `matcher`. Returns the header row index, or None.
    """
    for i in range(start_row, len(df)):
        a = _cell_txt(df.iloc[i, col])
        if a.lower() != 'month': continue
        b = _cell_txt(df.iloc[i, col + 1])
        if b and matcher(b.lower()):
            return i
    return None


def block_end(df, col, hdr):
    """Row index where a block stops: the next 'Month' header in the same
    column, or end of sheet. Blocks may grow to any length."""
    for i in range(hdr + 1, len(df)):
        if _cell_txt(df.iloc[i, col]).lower() == 'month':
            return i
    return len(df)


def block_year(df, col, hdr, default=2026):
    """Year sits in a title row a little above the header. Look upward."""
    for i in range(hdr - 1, max(hdr - 4, -1), -1):
        v = df.iloc[i, col]
        y = safe_int(v, 0)
        if 2000 <= y <= 2100:
            return y
    return default


def month_label(val, default_year=2026):
    """Accept either a month name ('July') or an Excel date. Returns
    'July 2026', or None when the cell is not a month."""
    if pd.isna(val): return None
    if isinstance(val, str):
        name = val.strip()
        if name in MONTH_ORDER:
            return f"{name} {default_year}"
    dt = pd.to_datetime(val, errors='coerce')
    if pd.notna(dt):
        return f"{MONTH_ORDER[dt.month - 1]} {dt.year}"
    name = _cell_txt(val)
    if name in MONTH_ORDER:
        return f"{name} {default_year}"
    return None


def read_screening(path):
    """Read the Screening sheet.

    Every block is located by its header labels, so blocks may move down the
    sheet or grow as months are added without the reader losing track.
    """
    df = pd.read_excel(path, sheet_name='Screening', header=None)
    data = {'realtime_hits': {}, 'realtime_generated': {}, 'realtime_update': {},
            'batch_alerts': {}, 'batch_truehit': {}, 'batch_update': {}}

    LEFT, RIGHT = 0, 19

    h_hits = find_block_header(df, LEFT,  lambda s: s.startswith('false hit'))
    h_gen  = find_block_header(df, LEFT,  lambda s: 'false hit individual' in s)
    h_upd  = find_block_header(df, LEFT,  lambda s: 'total alert generate' in s)
    h_bal  = find_block_header(df, RIGHT, lambda s: s.startswith('type alert'))
    h_bth  = find_block_header(df, RIGHT, lambda s: s == 'pep')
    h_bupd = find_block_header(df, RIGHT, lambda s: 'total alert generate' in s)

    missing = [n for n, h in [('Realtime hits', h_hits), ('Realtime generated', h_gen),
                              ('Realtime update', h_upd), ('Batch alerts', h_bal),
                              ('Batch true hit', h_bth), ('Batch update', h_bupd)] if h is None]
    if missing:
        print(f"  WARNING: Screening blocks not found: {', '.join(missing)}")

    # ── Realtime false/true hit by category ──
    if h_hits is not None:
        yr = block_year(df, LEFT, h_hits)
        for i in range(h_hits + 1, block_end(df, LEFT, h_hits)):
            label = month_label(df.iloc[i, LEFT], yr)
            if not label: continue
            vals = [safe_int(df.iloc[i, c]) for c in range(1, 17)]
            if sum(vals) == 0: continue
            data['realtime_hits'][label] = {
                'indiv_fh':  dict(zip(['PEP','Sanction','Internal List','Multitype'], vals[0:4])),
                'indiv_th':  dict(zip(['PEP','Sanction','Internal List','Multitype'], vals[4:8])),
                'entity_fh': dict(zip(['PEP','Sanction','Internal List','Multitype'], vals[8:12])),
                'entity_th': dict(zip(['PEP','Sanction','Internal List','Multitype'], vals[12:16])),
            }

    # ── Realtime generated alerts ──
    if h_gen is not None:
        yr = block_year(df, LEFT, h_gen)
        for i in range(h_gen + 1, block_end(df, LEFT, h_gen)):
            label = month_label(df.iloc[i, LEFT], yr)
            if not label: continue
            ifh, ith, itot = (safe_int(df.iloc[i, c]) for c in (1, 2, 3))
            efh, eth, etot = (safe_int(df.iloc[i, c]) for c in (4, 5, 6))
            if ifh + ith + itot + efh + eth + etot == 0: continue
            data['realtime_generated'][label] = {
                'indiv_fh': ifh, 'indiv_th': ith, 'indiv_total': itot,
                'entity_fh': efh, 'entity_th': eth, 'entity_total': etot,
            }

    # ── Realtime screening update ──
    if h_upd is not None:
        yr = block_year(df, LEFT, h_upd)
        for i in range(h_upd + 1, block_end(df, LEFT, h_upd)):
            label = month_label(df.iloc[i, LEFT], yr)
            if not label: continue
            gen, closed = safe_int(df.iloc[i, 1]), safe_int(df.iloc[i, 2])
            if gen == 0 and closed == 0: continue
            data['realtime_update'][label] = {'generated': gen, 'closed': closed}

    # ── Batch screening alerts (month cells are merged; carry forward) ──
    # Columns are found by their header text, so inserting a column in the
    # sheet (e.g. 'System Closed') doesn't shift the others.
    if h_bal is not None:
        yr = block_year(df, RIGHT, h_bal)
        BATCH_COLS = {'false_hit': 'false hit', 'true_hit': 'true hit',
                      'manually_closed': 'manually closed', 'system_closed': 'system closed',
                      'automation_closed': 'automation closed', 'alert_generated': 'alert generated'}
        hdr = {_cell_txt(df.iloc[h_bal, c]).lower(): c for c in range(RIGHT + 2, df.shape[1])}
        cols = {k: hdr.get(h) for k, h in BATCH_COLS.items()}
        missing_cols = [BATCH_COLS[k] for k, c in cols.items() if c is None and k != 'system_closed']
        if missing_cols:
            print(f"  WARNING: Batch Screening Alert columns not found: {', '.join(missing_cols)}")
        current = None
        for i in range(h_bal + 1, block_end(df, RIGHT, h_bal)):
            label = month_label(df.iloc[i, RIGHT], yr)
            if label:
                current = label
                data['batch_alerts'].setdefault(current, {})
            scenario = _cell_txt(df.iloc[i, RIGHT + 1])
            if not scenario.upper().startswith('BATCH') or not current: continue
            if all(pd.isna(df.iloc[i, c]) for c in cols.values() if c is not None): continue
            data['batch_alerts'][current][scenario] = {
                k: (safe_int(df.iloc[i, c]) if c is not None else 0) for k, c in cols.items()}
        data['batch_alerts'] = {k: v for k, v in data['batch_alerts'].items() if v}

    # ── Batch PEP individual true hit by type ──
    if h_bth is not None:
        yr = block_year(df, RIGHT, h_bth)
        for i in range(h_bth + 1, block_end(df, RIGHT, h_bth)):
            label = month_label(df.iloc[i, RIGHT], yr)
            if not label: continue
            pep, crime = safe_int(df.iloc[i, 20]), safe_int(df.iloc[i, 21])
            pepcrime, total = safe_int(df.iloc[i, 22]), safe_int(df.iloc[i, 23])
            if pep + crime + pepcrime + total == 0: continue
            data['batch_truehit'][label] = {
                'PEP': pep, 'Crime': crime, 'PEP_Crime': pepcrime, 'Total': total,
            }

    # ── Batch screening update ──
    if h_bupd is not None:
        yr = block_year(df, RIGHT, h_bupd)
        for i in range(h_bupd + 1, block_end(df, RIGHT, h_bupd)):
            label = month_label(df.iloc[i, RIGHT], yr)
            if not label: continue
            gen, closed = safe_int(df.iloc[i, 20]), safe_int(df.iloc[i, 21])
            if gen == 0 and closed == 0: continue
            data['batch_update'][label] = {'generated': gen, 'closed': closed}

    return data


def read_remittance(path):
    df = pd.read_excel(path, sheet_name='Remittance', header=0)
    result = {'monthly': {}, 'rows': {}}
    for _, row in df.iterrows():
        month = str(row.get('Month', '')).strip()
        if month not in MONTH_ORDER: continue
        label = f"{month} 2026"
        if label not in result['monthly']:
            result['monthly'][label] = {'data_received': 0, 'false_hit': 0, 'true_hit': 0, 'hit': 0, 'not_hit': 0}
            result['rows'][label] = []
        dr = safe_int(row.get('Total Data Received')); fh = safe_int(row.get('False Hit'))
        th = safe_int(row.get('True Hit')); hit = safe_int(row.get('HIT'))
        not_hit = safe_int(row.get('NOT HIT'))
        if dr + fh + th + hit + not_hit == 0: continue
        result['monthly'][label]['data_received'] += dr
        result['monthly'][label]['false_hit'] += fh
        result['monthly'][label]['true_hit'] += th
        result['monthly'][label]['hit'] += hit
        result['monthly'][label]['not_hit'] += not_hit
        email_date = row.get('Email Date', '')
        if pd.notna(email_date) and hasattr(email_date, 'strftime'):
            email_date = email_date.strftime('%d %b %Y')
        else:
            email_date = str(email_date)
        result['rows'][label].append({
            'email_date': email_date,
            'remittance_data': str(row.get('Remittance Data', '')),
            'false_hit': fh, 'true_hit': th, 'data_received': dr,
        })
    return result


def read_submerchant(path):
    df = pd.read_excel(path, sheet_name=' Sub Merchant', header=0)
    result = {}
    for _, row in df.iterrows():
        month = str(row.get('Month', '')).strip()
        if month not in MONTH_ORDER: continue
        label = f"{month} 2026"
        shop_id = safe_int(row.get('Internal Shop ID'))
        total_hit = safe_int(row.get('Total Hit'))
        if shop_id == 0 and total_hit == 0: continue
        result[label] = {
            'internal_shop_id': shop_id, 'total_hit': total_hit,
            'false_hit_indiv': safe_int(row.get('False Hit Individual')),
            'true_hit_indiv': safe_int(row.get('True Hit Individual')),
            'false_hit_entity': safe_int(row.get('False Hit Entity')),
            'true_hit_entity': safe_int(row.get('True Hit Entity')),
            'total_true_hit': safe_int(row.get('Total True Hit')),
            'total_false_hit': safe_int(row.get('Total False Hit')),
            'hit_rate': round(total_hit / shop_id * 100, 1) if shop_id > 0 else 0,
        }
    return result


def read_newsletter(path):
    df = pd.read_excel(path, sheet_name='Newsletter', header=0)
    df['Email Date'] = pd.to_datetime(df['Email Date'], errors='coerce')
    df = df[df['Email Date'].notna()].copy()
    df['Recorded Names'] = pd.to_numeric(df['Recorded Names'], errors='coerce').fillna(0)
    result = {'monthly': {}, 'rows': {}}
    for _, row in df.iterrows():
        dt = row['Email Date']; label = f"{MONTH_ORDER[dt.month-1]} {dt.year}"
        if label not in result['monthly']:
            result['monthly'][label] = {'total_names': 0, 'count': 0}
            result['rows'][label] = []
        names = int(row['Recorded Names'])
        result['monthly'][label]['total_names'] += names
        result['monthly'][label]['count'] += 1
        result['rows'][label].append({
            'email_date': dt.strftime('%d %b %Y'),
            'subject': str(row['Subject']),
            'recorded_names': names,
        })
    for label in result['monthly']:
        m = result['monthly'][label]
        m['avg'] = round(m['total_names'] / m['count'], 1) if m['count'] > 0 else 0
    return result


def read_dttot(path):
    df = pd.read_excel(path, sheet_name='DTTOT', header=0)
    df['Email Date'] = pd.to_datetime(df['Email Date'], errors='coerce')
    df = df[df['Email Date'].notna()].copy()
    df['Recorded Names'] = pd.to_numeric(df['Recorded Names'], errors='coerce').fillna(0)
    rows = []
    for _, row in df.iterrows():
        rows.append({
            'email_date': row['Email Date'].strftime('%d %b %Y'),
            'dttot_number': str(row['DTTOT Number']),
            'recorded_names': int(row['Recorded Names']),
        })
    return {'total_names': sum(r['recorded_names'] for r in rows), 'total_updates': len(rows), 'rows': rows}


def read_sanction(path):
    """Read all Sanction team data from one Excel file. Returns None if file missing."""
    if not os.path.exists(path):
        print(f"  [SANCTION] File not found: {path} — skipping Sanction team.")
        return None
    print(f"  [SANCTION] Reading {path}...")
    screening = read_screening(path)
    remittance = read_remittance(path)
    submerchant = read_submerchant(path)
    newsletter = read_newsletter(path)
    dttot = read_dttot(path)
    return {'screening': screening, 'remittance': remittance,
            'submerchant': submerchant, 'newsletter': newsletter, 'dttot': dttot}


# ═══════════════════════════════════════════════════════════
# NON TMS READER
# ═══════════════════════════════════════════════════════════

def read_nontms(path):
    if not os.path.exists(path):
        print(f"  [NON TMS] File not found: {path} — skipping NON TMS team.")
        return None
    print(f"  [NON TMS] Reading {path}...")
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['NON TMS']

    escalation = {}
    for r in range(4, 16):
        month = ws.cell(r, 2).value
        if not month or not isinstance(month, str): continue
        month = month.strip()
        fcc = safe_int(ws.cell(r, 3).value); risk = safe_int(ws.cell(r, 4).value)
        sanction = safe_int(ws.cell(r, 5).value); total = fcc + risk + sanction
        if total > 0:
            label = f"{month} 2026"
            escalation[label] = {'fcc': fcc, 'risk_ops': risk, 'sanction': sanction, 'total': total}

    by_user_type = {}
    current_month = None
    for r in range(20, 56):
        m = ws.cell(r, 2).value
        if m and isinstance(m, str) and m.strip(): current_month = m.strip()
        if not current_month: continue
        src = ws.cell(r, 3).value
        if not src or not isinstance(src, str): continue
        src = src.strip()
        kyc = safe_int(ws.cell(r, 4).value); non_kyc = safe_int(ws.cell(r, 5).value)
        merchant = safe_int(ws.cell(r, 6).value); non_dana = safe_int(ws.cell(r, 7).value)
        need_legal = safe_int(ws.cell(r, 8).value)
        total = kyc + non_kyc + merchant + non_dana + need_legal
        if total == 0: continue
        label = f"{current_month} 2026"
        if label not in by_user_type:
            by_user_type[label] = {'kyc':0,'non_kyc':0,'merchant':0,'non_dana':0,'need_legal':0,'total':0,
                                   'kyc_src':{},'non_kyc_src':{},'merchant_src':{},'non_dana_src':{}}
        sn = 'FCC' if 'FCC' in src else ('RiskOps' if 'RISK' in src else 'Sanction')
        by_user_type[label]['kyc'] += kyc; by_user_type[label]['non_kyc'] += non_kyc
        by_user_type[label]['merchant'] += merchant; by_user_type[label]['non_dana'] += non_dana
        by_user_type[label]['need_legal'] += need_legal; by_user_type[label]['total'] += total
        for fld, val in [('kyc_src',kyc),('non_kyc_src',non_kyc),('merchant_src',merchant),('non_dana_src',non_dana)]:
            by_user_type[label][fld][sn] = by_user_type[label][fld].get(sn, 0) + val

    decision_alert = {}
    current_month = None
    for r in range(60, 96):
        m = ws.cell(r, 2).value
        if m and isinstance(m, str) and m.strip(): current_month = m.strip()
        if not current_month: continue
        src = ws.cell(r, 3).value
        if not src or not isinstance(src, str): continue
        src = src.strip()
        monitor = safe_int(ws.cell(r, 4).value); utr = safe_int(ws.cell(r, 5).value)
        reported = safe_int(ws.cell(r, 6).value); data_provided = safe_int(ws.cell(r, 7).value)
        under_review = safe_int(ws.cell(r, 8).value)
        total = monitor + utr + reported + data_provided + under_review
        if total == 0: continue
        label = f"{current_month} 2026"
        sn = 'FCC' if 'FCC' in src else ('RiskOps' if 'RISK' in src else 'Sanction')
        if label not in decision_alert:
            decision_alert[label] = {'monitor':0,'utr':0,'reported':0,'data_provided':0,'under_review':0,'total':0,
                                     'monitor_src':{},'utr_src':{},'reported_src':{},'data_provided_src':{},'under_review_src':{}}
        decision_alert[label]['monitor'] += monitor; decision_alert[label]['utr'] += utr
        decision_alert[label]['reported'] += reported; decision_alert[label]['data_provided'] += data_provided
        decision_alert[label]['under_review'] += under_review; decision_alert[label]['total'] += total
        for fld, val in [('monitor_src',monitor),('utr_src',utr),('reported_src',reported),
                         ('data_provided_src',data_provided),('under_review_src',under_review)]:
            decision_alert[label][fld][sn] = decision_alert[label][fld].get(sn, 0) + val

    utr_data = {}
    # UTR Indicator table is on the same NON TMS sheet, find it by scanning for "SUMMARY OF UTR INDICATOR"
    utr_start = None
    for r in range(1, ws.max_row + 1):
        for c in range(1, 6):
            v = ws.cell(r, c).value
            if v and isinstance(v, str) and 'summary of utr indicator' in v.lower():
                utr_start = r
                break
        if utr_start: break

    # Also check separate sheets as fallback
    utr_ws = ws  # default to same sheet
    utr_header_row = utr_start + 1 if utr_start else None
    if not utr_start:
        for sn in ['UTR Indicator', 'SUMMARY OF UTR INDICATOR']:
            if sn in wb.sheetnames:
                utr_ws = wb[sn]
                utr_header_row = 1
                break

    if utr_header_row:
        # Find column positions from header row
        col_month = col_cat = col_freq = None
        for c in range(1, utr_ws.max_column + 1):
            h = str(utr_ws.cell(utr_header_row, c).value or '').strip().lower()
            if h == 'month': col_month = c
            elif 'breakdown' in h or 'category' in h: col_cat = c
            elif 'frequen' in h or 'count' in h or 'freq' in h: col_freq = c

        if col_month and col_cat and col_freq:
            current_month = None
            for r in range(utr_header_row + 1, utr_ws.max_row + 1):
                m = utr_ws.cell(r, col_month).value
                if m and isinstance(m, str) and m.strip():
                    month_str = m.strip()
                    if month_str in MONTH_ORDER:
                        current_month = f"{month_str} 2026"
                    else:
                        for mo in MONTH_ORDER:
                            if month_str.lower().startswith(mo.lower()[:3]):
                                current_month = f"{mo} 2026"
                                break
                if not current_month: continue
                cat = utr_ws.cell(r, col_cat).value
                if not cat or not isinstance(cat, str) or not cat.strip(): continue
                cat = cat.strip()
                freq = safe_int(utr_ws.cell(r, col_freq).value)
                if freq == 0: continue

                if current_month not in utr_data:
                    utr_data[current_month] = []
                utr_data[current_month].append({'category': cat, 'count': freq})

            # Sort each month's categories by count descending
            for label in utr_data:
                utr_data[label] = sorted(utr_data[label], key=lambda x: -x['count'])
    wb.close()
    return {'escalation': escalation, 'by_user_type': by_user_type,
            'decision_alert': decision_alert, 'utr': utr_data}


# ═══════════════════════════════════════════════════════════
# TMS READER
# ═══════════════════════════════════════════════════════════

_MONTH_SHORT_MAP = {
    'jan':'January','feb':'February','mar':'March','apr':'April',
    'may':'May','jun':'June','jul':'July','aug':'August',
    'sep':'September','oct':'October','nov':'November','dec':'December',
}

def _expand_month_label(s):
    import datetime as _dt
    if s is None: return None
    # Handle datetime objects (Excel stores date-formatted cells as datetime)
    if isinstance(s, (_dt.datetime, _dt.date)):
        mo_names=['January','February','March','April','May','June','July','August','September','October','November','December']
        return f'{mo_names[s.month-1]} {s.year}'
    if not isinstance(s, str): return None
    parts = s.strip().split()
    if len(parts) != 2: return None
    mo = _MONTH_SHORT_MAP.get(parts[0][:3].lower())
    if not mo: return None
    yr = parts[1]
    if len(yr) == 2 and yr.isdigit(): yr = '20' + yr
    return f'{mo} {yr}'

def _clean_notes(s):
    if s is None: return ''
    s = str(s).strip()
    if not s or s.upper() == 'N/A': return ''
    s = s.replace('\u200b', '').replace('\xa0', ' ')
    lines = [re.sub(r'[ \t]+', ' ', ln).strip() for ln in s.split('\n')]
    lines = [ln for ln in lines if ln]
    return '\n'.join(lines)


def read_tms(path):
    if not os.path.exists(path):
        print(f"  [TMS] File not found: {path} — skipping TMS team.")
        return None
    print(f"  [TMS] Reading {path}...")
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    out = {}

    def _rows(sheet_name, header_row=1):
        if sheet_name not in wb.sheetnames: return
        ws = wb[sheet_name]
        headers = [str(ws.cell(header_row, c).value or '').strip() for c in range(1, ws.max_column + 1)]
        for r in range(header_row + 1, ws.max_row + 1):
            vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            if all(v is None or (isinstance(v, str) and not v.strip()) for v in vals): continue
            yield {h: v for h, v in zip(headers, vals)}

    # Re-KYC User
    rekyc = {}
    for row in _rows('Re-KYC User'):
        label = _expand_month_label(row.get('Month'))
        if not label: continue
        rekyc[label] = {'changed':safe_int(row.get('Changed')),'unchanged':safe_int(row.get('Unchanged')),
                        'no_update':safe_int(row.get('No Update')),'revoked':safe_int(row.get('Revoked/Suspended')),
                        'notes':_clean_notes(row.get('Additional Notes'))}
    out['rekyc'] = rekyc

    # RFI User
    rfi_user = {}
    for row in _rows('RFI User'):
        label = _expand_month_label(row.get('Month'))
        if not label: continue
        rfi_user[label] = {'answer':safe_int(row.get('Answer')),'uncontactable':safe_int(row.get('Uncontactable')),
                           'on_process':safe_int(row.get('On Process')),'notes':_clean_notes(row.get('Additional Notes'))}
    out['rfi_user'] = rfi_user

    # RFI KYB
    rfi_kyb = {}
    for row in _rows('RFI KYB'):
        label = _expand_month_label(row.get('Month'))
        if not label: continue
        rfi_kyb[label] = {'success':safe_int(row.get('Success')),'failed':safe_int(row.get('Failed')),
                          'on_process':safe_int(row.get('On Process')),'notes':_clean_notes(row.get('Additional Notes'))}
    out['rfi_kyb'] = rfi_kyb

    # RFI Merchant (same structure as RFI KYB, separate sheet & separate data)
    rfi_merchant = {}
    if RFI_MERCHANT_SHEET not in wb.sheetnames:
        print(f"  [TMS] Sheet '{RFI_MERCHANT_SHEET}' not found — RFI Merchant tab will show 'No RFI Merchant data.'")
    for row in _rows(RFI_MERCHANT_SHEET):
        label = _expand_month_label(row.get('Month'))
        if not label: continue
        rfi_merchant[label] = {'success':safe_int(row.get('Success')),'failed':safe_int(row.get('Failed')),
                               'on_process':safe_int(row.get('On Process')),'notes':_clean_notes(row.get('Additional Notes'))}
    out['rfi_merchant'] = rfi_merchant

    # Risk Notes
    risknotes = {}
    for row in _rows('Risk Notes'):
        label = _expand_month_label(row.get('Month'))
        if not label: continue
        risknotes[label] = {'uid':safe_int(row.get('Number of UID')),'utr':safe_int(row.get('Number of UTR')),
                            'reported_case':safe_int(row.get('Reported Case')),'notes':_clean_notes(row.get('Additional Notes'))}
    out['risknotes'] = risknotes

    # P2P Receiver
    p2p = {}
    for row in _rows('P2P Receiver'):
        label = _expand_month_label(row.get('Month'))
        if not label: continue
        p2p[label] = {'uid':safe_int(row.get('Number of UID')),'utr':safe_int(row.get('Number of UTR')),
                      'notes':_clean_notes(row.get('Additional Notes'))}
    out['p2p'] = p2p

    # Anomaly User
    anomaly = {}
    for row in _rows('Anomaly User'):
        label = _expand_month_label(row.get('Month'))
        if not label: continue
        anomaly[label] = {'uid':safe_int(row.get('Number of UID')),'utr':safe_int(row.get('Number of UTR')),
                          'reported':safe_int(row.get('Reported')),'monitored_fv':safe_int(row.get('Monitored + Face Verification')),
                          'notes':_clean_notes(row.get('Additional Notes'))}
    out['anomaly'] = anomaly

    # Face Verification
    faceverif = {}
    for row in _rows('Face Verification'):
        label = _expand_month_label(row.get('Month'))
        if not label: continue
        faceverif[label] = {'uid':safe_int(row.get('Number of UID')),'source':str(row.get('Source') or ''),
                            'notes':_clean_notes(row.get('Additional Notes'))}
    out['faceverif'] = faceverif

    # Dormant
    dormant = {}
    for row in _rows('Dormant'):
        label = _expand_month_label(row.get('Month'))
        if not label: continue
        dormant[label] = {'monitored':safe_int(row.get('Number of Monitored')),'uid':safe_int(row.get('Number of UID')),
                          'utr':safe_int(row.get('Number of UTR')),'notes':_clean_notes(row.get('Additional Notes'))}
    out['dormant'] = dormant

    # Merchant Screening
    merchant = {}
    for row in _rows('Merchant Screening'):
        label = _expand_month_label(row.get('Month'))
        if not label: continue
        merchant[label] = {'alerts':safe_int(row.get('Number of Alert')),'notes':_clean_notes(row.get('Additional Notes'))}
    out['merchant'] = merchant

    # TMS
    tms = {}
    for row in _rows('TMS'):
        label = _expand_month_label(row.get('Month'))
        if not label: continue
        alerts = safe_int(row.get('Number of Alert ') or row.get('Number of Alert'))
        tms[label] = {'alerts':alerts,'monitored':safe_int(row.get('Closed as Monitored')),
                      'issue':safe_int(row.get('Closed as Issue')),'notes':_clean_notes(row.get('Additional Notes'))}
    out['tms'] = tms

    # Unmasking Request
    unmask = {}
    if 'Unmasking Request' in wb.sheetnames:
        ws = wb['Unmasking Request']; current_label = None
        for r in range(2, ws.max_row + 1):
            m = ws.cell(r, 1).value
            label = _expand_month_label(m) if m else None
            if label: current_label = label
            if not current_label: continue
            src = ws.cell(r, 2).value
            if not src or not isinstance(src, str) or not src.strip(): continue
            src = src.strip(); uid = safe_int(ws.cell(r, 3).value); recv = safe_int(ws.cell(r, 4).value)
            notes = _clean_notes(ws.cell(r, 5).value) if ws.cell(r, 5).value else ''
            if current_label not in unmask:
                unmask[current_label] = {'sources': [], 'uid_total': 0, 'recv_total': 0, 'notes': ''}
            # Skip N/A sources but still create the month entry for notes
            if src.upper() != 'N/A':
                unmask[current_label]['sources'].append({'source': src, 'uid': uid, 'receivers': recv})
                unmask[current_label]['uid_total'] += uid; unmask[current_label]['recv_total'] += recv
            if notes and not unmask[current_label]['notes']: unmask[current_label]['notes'] = notes
    out['unmask'] = unmask

    # CS Escalation
    cs = {}
    if 'Eskalasi to CS' in wb.sheetnames:
        ws = wb['Eskalasi to CS']; current_label = None; current_cat = None
        for r in range(2, ws.max_row + 1):
            m = ws.cell(r, 1).value
            label = _expand_month_label(m) if m else None
            if label: current_label = label
            if not current_label: continue
            cat = ws.cell(r, 2).value
            if cat and isinstance(cat, str) and cat.strip():
                current_cat = cat.strip()
            if not current_cat: continue
            remarks = ws.cell(r, 3).value
            number = safe_int(ws.cell(r, 4).value); action = ws.cell(r, 5).value
            if number == 0: continue
            if current_label not in cs:
                cs[current_label] = {'rows': [], 'by_category': {}, 'total': 0}
            cs[current_label]['rows'].append({'category':current_cat,'remarks':_clean_notes(remarks),
                                              'number':number,'action':_clean_notes(action)})
            disp = 'Fake ID Card' if 'fake' in current_cat.lower() and 'id' in current_cat.lower() else current_cat
            cs[current_label]['by_category'][disp] = cs[current_label]['by_category'].get(disp, 0) + number
            cs[current_label]['total'] += number
    out['cs'] = cs
    wb.close()
    return out


# ═══════════════════════════════════════════════════════════
# COMPUTE DERIVED (Sanction team only)
# ═══════════════════════════════════════════════════════════

def compute_derived(sanction_data, nontms_data, tms_data):
    """Build the full JSON store. Each team is optional (can be None)."""

    store = {'months': [], 'data': {}, 'dttot': {'total_names': 0, 'total_updates': 0, 'rows': []},
             'nontms': {'escalation': {}, 'by_user_type': {}, 'decision_alert': {}, 'utr': []},
             'tms': {}}

    if sanction_data:
        screening = sanction_data['screening']
        remittance = sanction_data['remittance']
        submerchant = sanction_data['submerchant']
        newsletter = sanction_data['newsletter']
        dttot = sanction_data['dttot']

        all_months = set()
        for src in [screening['realtime_hits'], screening['realtime_generated'],
                    screening['realtime_update'], screening['batch_alerts'],
                    screening['batch_truehit'], screening['batch_update'],
                    remittance['monthly'], submerchant, newsletter['monthly']]:
            all_months.update(src.keys())
        all_months = sorted(all_months, key=month_sort_key)

        months_data = {}
        for label in all_months:
            m = {}
            hits = screening['realtime_hits'].get(label, {})
            gen = screening['realtime_generated'].get(label, {})
            upd = screening['realtime_update'].get(label, {})

            if hits:
                ifh = hits['indiv_fh']; ith = hits['indiv_th']
                efh = hits['entity_fh']; eth = hits['entity_th']
                m['alert_pep'] = ifh['PEP'] + ith['PEP'] + efh['PEP'] + eth['PEP']
                m['alert_sanction'] = ifh['Sanction'] + ith['Sanction'] + efh['Sanction'] + eth['Sanction']
                m['alert_internal'] = ifh['Internal List'] + ith['Internal List'] + efh['Internal List'] + eth['Internal List']
                m['alert_multitype'] = ifh['Multitype'] + ith['Multitype'] + efh['Multitype'] + eth['Multitype']
                m['total_alerts'] = m['alert_pep'] + m['alert_sanction'] + m['alert_internal'] + m['alert_multitype']
                m['true_hit_indiv'] = sum(ith.values()); m['true_hit_entity'] = sum(eth.values())
                m['true_hit_total'] = m['true_hit_indiv'] + m['true_hit_entity']
                m['false_hit_indiv'] = sum(ifh.values()); m['false_hit_entity'] = sum(efh.values())
                m['indiv_fh'] = ifh; m['indiv_th'] = ith; m['entity_fh'] = efh; m['entity_th'] = eth

            if gen:
                m['indiv_generated'] = gen['indiv_total']; m['entity_generated'] = gen['entity_total']
                m['indiv_fh_total'] = gen['indiv_fh']; m['indiv_th_total'] = gen['indiv_th']
                m['entity_fh_total'] = gen['entity_fh']; m['entity_th_total'] = gen['entity_th']

            if upd:
                m['rt_generated'] = upd['generated']; m['rt_closed'] = upd['closed']
                m['closure_rate'] = round(upd['closed'] / upd['generated'] * 100, 1) if upd['generated'] > 0 else 0

            m['true_hit_rate'] = round(m.get('true_hit_total', 0) / m['rt_generated'] * 100, 2) if m.get('rt_generated') else 0

            ba = screening['batch_alerts'].get(label, {})
            if ba:
                m['batch_scenarios'] = ba
                m['batch_total_fh'] = sum(s['false_hit'] for s in ba.values())
                m['batch_total_th'] = sum(s['true_hit'] for s in ba.values())
                m['batch_total_mc'] = sum(s['manually_closed'] for s in ba.values())
                m['batch_total_sc'] = sum(s.get('system_closed', 0) for s in ba.values())
                m['batch_total_ac'] = sum(s['automation_closed'] for s in ba.values())
                m['batch_total_ag'] = sum(s['alert_generated'] for s in ba.values())

            bth = screening['batch_truehit'].get(label, {})
            if bth:
                m['batch_th_pep'] = bth['PEP']; m['batch_th_crime'] = bth['Crime']
                m['batch_th_pepcrime'] = bth['PEP_Crime']; m['batch_th_total'] = bth['Total']

            bupd = screening['batch_update'].get(label, {})
            if bupd:
                m['batch_generated'] = bupd['generated']; m['batch_closed'] = bupd['closed']

            rem = remittance['monthly'].get(label, {})
            if rem:
                m['rem_data_received'] = rem['data_received']; m['rem_false_hit'] = rem['false_hit']
                m['rem_true_hit'] = rem['true_hit']
                m['rem_hit_rate'] = round((rem['false_hit'] + rem['true_hit']) / rem['data_received'] * 100, 1) if rem['data_received'] > 0 else 0
            m['rem_rows'] = remittance['rows'].get(label, [])

            sm = submerchant.get(label, {})
            if sm:
                m['subm_shop_id'] = sm['internal_shop_id']; m['subm_total_hit'] = sm['total_hit']
                m['subm_hit_rate'] = sm['hit_rate']; m['subm_detail'] = sm

            news = newsletter['monthly'].get(label, {})
            if news:
                m['news_total'] = news['total_names']; m['news_count'] = news['count']; m['news_avg'] = news['avg']
            m['news_rows'] = newsletter['rows'].get(label, [])

            months_data[label] = m

        def has_real_data(d):
            """Return True if the month has meaningful screening/alert data (not just empty placeholders)."""
            # Check the key indicators that prove real data exists
            checks = ['total_alerts','rt_generated','rt_closed','batch_generated','batch_closed',
                      'batch_total_th','batch_total_fh','rem_data_received','rem_false_hit','rem_true_hit',
                      'news_total','subm_total_hit','subm_shop_id','true_hit_total','false_hit_indiv']
            for k in checks:
                v = d.get(k, 0)
                if isinstance(v, (int, float)) and v > 0:
                    return True
            return False

        real_months = [label for label in all_months if has_real_data(months_data.get(label, {}))]

        store['months'] = real_months
        store['data'] = {k: months_data[k] for k in real_months}
        store['dttot'] = dttot

    # NON TMS — independent, just attach if available
    if nontms_data:
        store['nontms'] = nontms_data
    else:
        store['nontms'] = {'escalation': {}, 'by_user_type': {}, 'decision_alert': {}, 'utr': {}}

    # TMS — independent, just attach if available
    if tms_data:
        store['tms'] = tms_data
    else:
        store['tms'] = {}

    return store


# ═══════════════════════════════════════════════════════════
# RFI MERCHANT — dashboard UI (TMS section)
# ═══════════════════════════════════════════════════════════
# The TMS "RFI Merchant" section mirrors "RFI KYB" (KPI cards, status doughnut,
# Month-by-Month table, At a Glance + Highlights) and reuses the same shared
# helpers. It is added to the template in memory at build time, so
# AML_Dashboard_template.html stays untouched and every regenerated
# AML_Dashboard.html keeps the tab.
# Each part is skipped if it's already in the template (safe to run repeatedly).
# If a required anchor is missing, the template is used as-is (no half-patch).

_RFI_MERCHANT_RENDERER = r'''

// ── Sub-page: RFI Merchant ──
// Same layout as RFI KYB; reads STORE.tms.rfi_merchant (sheet "RFI Merchant").
function renderTmsRfiMerchant(el,data){
  var months=tmsMonthsOf(data);
  if(!months.length){el.innerHTML='<div class="placeholder">No RFI Merchant data.</div>';return}
  var cm=curMonth;if(months.indexOf(cm)===-1)cm=months[months.length-1];
  var isMulti=months.length>1;
  var d=isMulti?tmsAgg(data,months):(data[cm]||{});
  var prevLbl=isMulti?null:(months.indexOf(cm)>0?months[months.indexOf(cm)-1]:null);
  var prev=prevLbl?data[prevLbl]:null;
  var deltas=tmsDeltas(d,prev,prevLbl,['success','failed','on_process']);
  var O=outliers(deltas);
  var h='';
  h+='<div class="kr" style="grid-template-columns:repeat(3,1fr)">';
  h+=kp('Success',fmt(N(d.success)),deltas[0],{hl:O[0]});
  h+=kp('Failed',fmt(N(d.failed)),deltas[1],{hl:O[1]});
  h+=kp('On Process',fmt(N(d.on_process)),deltas[2],{hl:O[2]});
  h+='</div>';
  h+='<div class="tms-grid">';
  // Left: doughnut + small monthly table
  h+='<div class="tms-stack">';
  h+='<div class="cd"><div class="cd-t">Status Distribution \u2014 '+shortMonth(cm)+'</div><div style="display:flex;align-items:center;justify-content:center;gap:1.5rem"><div style="width:180px;height:180px;flex-shrink:0"><canvas id="tms-rfim-d"></canvas></div><div style="font-size:.7rem;color:var(--text3);line-height:1.7"><div><span style="display:inline-block;width:9px;height:9px;border-radius:2px;background:#2196F3;margin-right:.5rem"></span>Success: <b style="color:var(--text)">'+fmt(N(d.success))+'</b></div><div><span style="display:inline-block;width:9px;height:9px;border-radius:2px;background:#5C9CE6;margin-right:.5rem"></span>Failed: <b style="color:var(--text)">'+fmt(N(d.failed))+'</b></div><div><span style="display:inline-block;width:9px;height:9px;border-radius:2px;background:#00ACC1;margin-right:.5rem"></span>On Process: <b style="color:var(--text)">'+fmt(N(d.on_process))+'</b></div></div></div></div>';
  // Monthly table (hide for single month)
  if(months.length>=2){
  h+='<div class="cd"><div class="cd-t">Month-by-Month</div><table class="tms-mini-tb"><thead><tr><th>Month</th><th class="num">Success</th><th class="num">Failed</th><th class="num">On Process</th><th class="num">Total</th></tr></thead><tbody>';
  months.forEach(function(m){var r=data[m]||{};var t=N(r.success)+N(r.failed)+N(r.on_process);h+='<tr><td>'+shortMonth(m)+'</td><td class="num">'+fmt(N(r.success))+'</td><td class="num">'+fmt(N(r.failed))+'</td><td class="num">'+fmt(N(r.on_process))+'</td><td class="num"><b>'+fmt(t)+'</b></td></tr>'});
  h+='</tbody></table></div>';
  }
  h+='</div>';
  // Right: narrative
  var rmTotal=N(d.success)+N(d.failed)+N(d.on_process);
  var rmSuccessRate=rmTotal>0?Math.round(N(d.success)/rmTotal*1000)/10:0;
  var rmTrend=tmsTrend(d,prev,function(r){return N(r.success)+N(r.failed)+N(r.on_process)});
  var rmGl=[{l:'Total Sent',v:fmt(rmTotal),sub:rmTrend?rmTrend.txt+(prevLbl?' vs '+shortMonth(prevLbl):''):'',subCls:rmTrend?rmTrend.cls:''},{l:'Success Rate',v:rmSuccessRate+'%',sub:fmt(N(d.success))+' Success'},{l:'In Progress',v:fmt(N(d.on_process)),sub:'Under monitoring'}];
  h+='<div class="tms-narr"><div class="cd-t" style="margin-bottom:.55rem">At a Glance</div>'+tmsGlance(rmGl,3)+'<div class="cd-t" style="margin-top:.7rem;margin-bottom:.45rem">Highlights</div>'+(isMulti?tmsBulletsMulti(data,months):tmsBullets(d.notes))+'</div>';
  h+='</div>';
  el.innerHTML=h;
  // Doughnut
  var pieCanvas=document.getElementById('tms-rfim-d');
  var pieCtx=pieCanvas.getContext('2d');
  var pieGrads=[['#2196F3','#90CAF9'],['#0D47A1','#5472D3'],['#00ACC1','#4DD0E1']].map(function(p){var g=pieCtx.createLinearGradient(0,0,0,200);g.addColorStop(0,p[0]);g.addColorStop(1,p[1]);return g});
  allCharts.push(new Chart(pieCanvas,{type:'doughnut',data:{labels:['Success','Failed','On Process'],datasets:[{data:[N(d.success),N(d.failed),N(d.on_process)],backgroundColor:pieGrads,borderWidth:2,borderColor:gcv('--card')||'#fff',hoverOffset:8}]},options:{maintainAspectRatio:false,cutout:'58%',plugins:{legend:{display:false},tooltip:tt,datalabels:mkDDL()},layout:{padding:4}}}));
}'''


def _ins_after_line(html, anchor, snippet):
    """Insert `snippet` as new line(s) right after the line containing `anchor`."""
    i = html.find(anchor)
    j = html.find('\n', i)
    if j < 0: j = len(html)
    return html[:j] + '\n' + snippet + html[j:]


def _ins_after_function(html, anchor, snippet):
    """Insert `snippet` right after the closing '}' of the top-level function starting at `anchor`."""
    i = html.find(anchor)
    j = html.find('\n}\n', i)
    if j < 0: return None
    j += 2
    return html[:j] + snippet + html[j:]


def _replace_once(html, anchor, snippet):
    return html.replace(anchor, snippet, 1)


# (name, already-present marker, anchor, insert function, snippet, required)
RFI_MERCHANT_UI_PATCHES = [
    ('mount div',
     'id="tms-rfi-merchant-content"',
     '<div id="tms-rfi-kyb-content"></div>',
     _ins_after_line,
     '  <div id="tms-rfi-merchant-content"></div>',
     True),
    ('section entry',
     "{id:'tms-rfi-merchant'",
     "{id:'tms-unmask',",
     _ins_after_line,
     "  {id:'tms-rfi-merchant',team:'tms',label:'RFI Merchant',renderer:function(){renderTms('tms-rfi-merchant')},mountId:'tms-rfi-merchant-content',group:'tms-e'},",
     True),
    ('sparkline map',
     "sp==='tms-rfi-merchant'",
     "else if(sp==='tms-rfi-kyb')",
     _ins_after_line,
     "    // RFI Merchant shares KPI labels with RFI KYB / RFI User: only map when those aren't on screen\n"
     "    else if(sp==='tms-rfi-merchant'&&!document.querySelector('#v2-stack #tms-rfi-kyb-content,#v2-stack #tms-rfi-user-content')){var src=tms.rfi_merchant||{};tmsMap['success']=function(m){return N((src[m]||{}).success)};tmsMap['failed']=function(m){return N((src[m]||{}).failed)};tmsMap['on process']=function(m){return N((src[m]||{}).on_process)}}",
     True),
    ('dispatch',
     "tabId==='tms-rfi-merchant'",
     "if(tabId==='tms-rfi-kyb')",
     _ins_after_line,
     "  if(tabId==='tms-rfi-merchant') return renderTmsRfiMerchant(el,tms.rfi_merchant||{});",
     True),
    ('renderer',
     'function renderTmsRfiMerchant(',
     'function renderTmsRfiKyb(',
     _ins_after_function,
     _RFI_MERCHANT_RENDERER,
     True),
    ('notes keyword',
     "'RFI KYB','RFI Merchant'",
     "'KYB','RFI KYB'",
     _replace_once,
     "'KYB','RFI KYB','RFI Merchant'",
     False),
]


def add_rfi_merchant_ui(template):
    """Return the template with the TMS → RFI Merchant section added (idempotent)."""
    html = template
    added = []
    for name, marker, anchor, fn, snippet, required in RFI_MERCHANT_UI_PATCHES:
        if marker in html:
            continue
        result = fn(html, anchor, snippet) if html.count(anchor) == 1 else None
        if result is None:
            if required:
                print(f"  WARNING: RFI Merchant UI anchor not found ({name}) — tab NOT added, template used as-is.")
                return template
            print(f"  NOTE: RFI Merchant optional part skipped ({name}).")
            continue
        html = result
        added.append(name)
    if added:
        print(f"  [TMS] RFI Merchant tab added to dashboard ({', '.join(added)}).")
    return html



# ═══════════════════════════════════════════════════════════
# PER-TEAM DATA FILES (data/<team>.json)
# ═══════════════════════════════════════════════════════════

EMPTY_TEAM = {
    'sanction': {'months': [], 'data': {}, 'dttot': {'total_names': 0, 'total_updates': 0, 'rows': []}},
    'nontms':   {'escalation': {}, 'by_user_type': {}, 'decision_alert': {}, 'utr': {}},
    'tms':      {},
}


def team_json_path(team):
    return os.path.join(DATA_DIR, f'{team}.json')


def excel_path(team):
    """Excel location: excel_paths.local.json override, else next to this script."""
    if os.path.exists(LOCAL_PATHS_FILE):
        try:
            with open(LOCAL_PATHS_FILE, encoding='utf-8') as f:
                custom = json.load(f).get(team)
            if custom:
                return custom
        except ValueError as e:
            print(f"ERROR: {os.path.basename(LOCAL_PATHS_FILE)} is not valid JSON ({e}).")
            sys.exit(1)
    return os.path.join(HERE, TEAMS[team]['excel'])


def read_team(team, path):
    """Read one team's Excel and return the data that goes into data/<team>.json."""
    if team == 'sanction':
        part = compute_derived(read_sanction(path), None, None)
        return {'months': part['months'], 'data': part['data'], 'dttot': part['dttot']}
    if team == 'nontms':
        return read_nontms(path)
    return read_tms(path)


def team_months(team, payload):
    """Every month that has data for this team."""
    if not payload: return []
    if team == 'sanction':
        months = set(payload.get('months', []))
    else:
        months = set()
        for sub in payload.values():
            if isinstance(sub, dict): months.update(sub.keys())
    return sorted(months, key=month_sort_key)


def load_team_file(team):
    """Published data for a team, or None when it has never been published."""
    path = team_json_path(team)
    if not os.path.exists(path): return None
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def write_team_file(team, payload, source):
    os.makedirs(DATA_DIR, exist_ok=True)
    doc = {
        '_meta': {
            'team': TEAMS[team]['label'],
            'updated_at': datetime.now().strftime('%Y-%m-%d %H:%M'),
            'updated_by': git_user(),
            'source_file': os.path.basename(source),
            'months': team_months(team, payload),
        },
        'data': payload,
    }
    with open(team_json_path(team), 'w', encoding='utf-8', newline='\n') as f:
        json.dump(doc, f, ensure_ascii=False, indent=1)
        f.write('\n')


# ═══════════════════════════════════════════════════════════
# BUILD index.html FROM ALL TEAMS
# ═══════════════════════════════════════════════════════════

def build_store(teams_data):
    """teams_data: {team: payload or None}. Returns the STORE object for the template."""
    s = teams_data.get('sanction') or EMPTY_TEAM['sanction']
    nontms_data = teams_data.get('nontms')
    tms_data = teams_data.get('tms')

    store = {'months': list(s['months']), 'data': dict(s['data']), 'dttot': s['dttot'],
             'nontms': nontms_data or EMPTY_TEAM['nontms'],
             'tms': tms_data or {}}

    # Months available across all teams
    all_available = set(store['months'])
    if nontms_data:
        all_available.update(nontms_data.get('escalation', {}).keys())
    if tms_data:
        for sub in tms_data.values():
            if isinstance(sub, dict):
                all_available.update(sub.keys())
    all_available = sorted(all_available, key=month_sort_key)

    # Merge months from ALL teams — not just Sanction
    merged = sorted(set(store['months']) | set(all_available), key=month_sort_key)
    # Add empty Sanction data placeholders for months only in TMS/NON-TMS
    for m in merged:
        if m not in store['data']:
            store['data'][m] = {}
    store['months'] = merged

    highlight = HIGHLIGHT_MONTH
    if highlight == "latest":
        highlight = all_available[-1] if all_available else ""
    elif highlight not in all_available:
        print(f"Warning: '{highlight}' not found, using latest.")
        highlight = all_available[-1] if all_available else ""

    store['highlight'] = highlight
    store['notes'] = NOTES
    return store


def write_dashboard(store, out_path):
    if not os.path.exists(TEMPLATE_FILE):
        print(f"ERROR: Template not found: {TEMPLATE_FILE}"); sys.exit(1)
    with open(TEMPLATE_FILE, 'r', encoding='utf-8') as f:
        template = f.read()
    template = add_rfi_merchant_ui(template)
    output = template.replace('/*STORE_JSON*/', json.dumps(store, ensure_ascii=False, indent=None))
    if '/*STORE_JSON*/' in output:
        print("ERROR: STORE_JSON marker not replaced!"); sys.exit(1)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(output)


def shortm(label):
    parts = label.split()
    return f"{parts[0][:3]} {parts[1]}" if len(parts) == 2 else label


def build(fresh, out_path):
    """Rebuild the dashboard. `fresh` = {team: payload} just read from Excel;
    every other team comes from its published data/<team>.json."""
    teams_data = {}
    print("\nTeams on the dashboard:")
    for team, cfg in TEAMS.items():
        if team in fresh:
            teams_data[team] = fresh[team]
            src = 'from your Excel'
        else:
            doc = load_team_file(team)
            teams_data[team] = doc['data'] if doc else None
            if doc:
                meta = doc.get('_meta', {})
                src = f"as published {meta.get('updated_at', '?')} by {meta.get('updated_by', '?')}"
            else:
                src = 'nothing published yet'
        months = team_months(team, teams_data[team])
        span = f"{shortm(months[0])} - {shortm(months[-1])}" if months else 'no months'
        print(f"  {cfg['label']:<9} {span:<19} {src}")
    store = build_store(teams_data)
    write_dashboard(store, out_path)
    print(f"\nHighlight month: {store['highlight']}")
    print(f"Built {out_path} ({os.path.getsize(out_path) / 1024:.0f} KB)")


# ═══════════════════════════════════════════════════════════
# GIT (publish to GitHub)
# ═══════════════════════════════════════════════════════════

PAGES_URL = "https://zahfaizah.github.io/aml-dashboard/"


def git(*args, check=True):
    try:
        r = subprocess.run(['git', *args], cwd=REPO, capture_output=True,
                           text=True, encoding='utf-8', errors='replace')
    except FileNotFoundError:
        print("ERROR: git is not installed (or not on PATH). See README.md -> Setup.")
        sys.exit(1)
    if check and r.returncode != 0:
        print(f"ERROR: git {' '.join(args)} failed:\n{(r.stderr or r.stdout).strip()}")
        sys.exit(1)
    return r


def git_user():
    try:
        name = subprocess.run(['git', 'config', 'user.name'], cwd=REPO, capture_output=True,
                              text=True, encoding='utf-8', errors='replace').stdout.strip()
    except FileNotFoundError:
        name = ''
    return name or getpass.getuser()


def git_pull():
    """Get everyone's latest published data. Returns the files that changed."""
    before = git('rev-parse', 'HEAD').stdout.strip()
    r = git('pull', '--rebase', '--autostash', check=False)
    if r.returncode != 0:
        git('rebase', '--abort', check=False)
        print("ERROR: could not get the latest version from GitHub:")
        print("  " + (r.stderr or r.stdout).strip().replace('\n', '\n  '))
        print("Check your internet / GitHub login and run again.")
        sys.exit(1)
    after = git('rev-parse', 'HEAD').stdout.strip()
    if before == after: return []
    return git('diff', '--name-only', before, after).stdout.split()


def answer_yes(question, default_yes):
    hint = '[Y/n]' if default_yes else '[y/N]'
    try:
        a = input(f"{question} {hint} ").strip().lower()
    except EOFError:
        return default_yes
    return default_yes if not a else a in ('y', 'yes')


def month_slice(team, payload, month):
    """Everything a team has for one month (to compare old vs new)."""
    if team == 'sanction':
        return payload.get('data', {}).get(month)
    return {k: v.get(month) for k, v in payload.items() if isinstance(v, dict)}


def stale_check(team, payload):
    """Stop a PIC from publishing an older Excel over newer published data."""
    doc = load_team_file(team)
    if not doc: return
    meta = doc.get('_meta', {})
    old = set(team_months(team, doc['data']))
    new = set(team_months(team, payload))
    lost = sorted(old - new, key=month_sort_key)
    # Someone else published this team last: also flag months whose numbers would change
    changed = []
    if meta.get('updated_by') and meta.get('updated_by') != git_user():
        fresh = json.loads(json.dumps(payload))   # same shape as the JSON file
        changed = [m for m in sorted(old & new, key=month_sort_key)
                   if month_slice(team, fresh, m) != month_slice(team, doc['data'], m)]
    if not lost and not changed: return
    print(f"\n  WARNING: {TEAMS[team]['label']} was last published "
          f"{meta.get('updated_at', '?')} by {meta.get('updated_by', '?')}.")
    if lost:
        print(f"    Months on GitHub that your Excel doesn't have: {', '.join(lost)}")
    if changed:
        print(f"    Months where your numbers differ from theirs:  {', '.join(changed)}")
    print("    Your Excel may be an older copy. Publishing would replace their data.")
    if not answer_yes("  Publish anyway?", default_yes=False):
        print("Cancelled - nothing was changed."); sys.exit(1)


def _same_folder(a, b):
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def sync_work_folder():
    """Running from a work folder outside the repo (e.g. OneDrive): keep this
    folder's script and template the same as the published ones."""
    if _same_folder(HERE, REPO): return
    for name in ('generate_dashboard.py', 'AML_Dashboard_template.html'):
        src, dst = os.path.join(REPO, name), os.path.join(HERE, name)
        if not os.path.exists(src): continue
        with open(src, 'rb') as f: new = f.read().replace(b'\r\n', b'\n')
        old = None
        if os.path.exists(dst):
            with open(dst, 'rb') as f: old = f.read().replace(b'\r\n', b'\n')
        if old == new: continue
        if old is not None and name.endswith('.html'):
            backup = dst.replace('.html', '_backup.html')
            with open(backup, 'wb') as f: f.write(old)
            print(f"  NOTE: {name} here was different from the published one; saved it as "
                  f"{os.path.basename(backup)}. Design changes go in {src}")
        with open(dst, 'wb') as f: f.write(new)
        if name.endswith('.py'):
            print(f"\nThe dashboard script was updated (copied from {REPO}). Please run the bat again.")
            sys.exit(1)


def copy_to_work_folder():
    """Also put the latest dashboard in the work folder, to open it from there."""
    if _same_folder(HERE, REPO): return
    dst = os.path.join(HERE, LOCAL_COPY_NAME)
    with open(OUTPUT_FILE, 'rb') as f: data = f.read()
    with open(dst, 'wb') as f: f.write(data)
    print(f"Copied the dashboard to {dst}")


def publish(fresh, sources, assume_yes):
    if not os.path.isdir(os.path.join(REPO, '.git')):
        print(f"ERROR: GitHub folder not found: {REPO}")
        print("Clone it first (README.md -> Setup), or set its path as \"repo\" in "
              f"{LOCAL_PATHS_FILE}")
        sys.exit(1)
    print(f"\nGetting the latest dashboard from GitHub ({REPO})...")
    changed = git_pull()
    if 'generate_dashboard.py' in changed and _same_folder(HERE, REPO):
        print("\nThe dashboard script was updated on GitHub. Please run the bat again.")
        sys.exit(1)
    sync_work_folder()

    for team, payload in fresh.items():
        stale_check(team, payload)

    labels = ' + '.join(TEAMS[t]['label'] for t in fresh)
    paths = ['index.html'] + [f'data/{t}.json' for t in fresh]
    was_tracked = {p: git('ls-files', '--error-unmatch', p, check=False).returncode == 0 for p in paths}

    def restore():
        """Put index.html and our data files back to the committed version."""
        for p in paths:
            if was_tracked[p]:
                git('checkout', '--', p, check=False)
            elif os.path.exists(os.path.join(REPO, p)):
                os.remove(os.path.join(REPO, p))

    for attempt in range(1, 4):
        for team, payload in fresh.items():
            write_team_file(team, payload, sources[team])
        build(fresh, OUTPUT_FILE)

        # Same numbers as already published (only the timestamp differs)? Nothing to do.
        index_same = git('diff', '--quiet', '--', 'index.html', check=False).returncode == 0
        if index_same and only_meta_changed(fresh):
            restore()
            copy_to_work_folder()
            print("\nNo changes - the dashboard on GitHub already has this data.")
            return

        if attempt == 1 and not assume_yes and not answer_yes(
                "\nYou can open index.html now to check it. Publish to GitHub?", default_yes=True):
            restore()
            print("Cancelled - nothing was published."); return

        git('add', '--', *paths)
        git('commit', '-q', '-m', f"{labels} update - {datetime.now():%Y-%m-%d} (by {git_user()})")
        r = git('push', check=False)
        if r.returncode == 0:
            copy_to_work_folder()
            print(f"\nPublished! Live in about a minute at {PAGES_URL}")
            return

        # Undo our commit so the local copy matches GitHub again
        git('reset', '-q', '--mixed', 'HEAD~1')
        restore()
        msg = r.stderr or r.stdout
        if any(k in msg for k in ('rejected', 'fetch first', 'non-fast-forward')):
            print("\nSomeone else published at the same time - getting their update and rebuilding...")
            git_pull()
            continue
        print("ERROR: push to GitHub failed:\n  " + msg.strip().replace('\n', '\n  '))
        print("Nothing was published. Check your internet / GitHub access and run again.")
        sys.exit(1)

    print("ERROR: could not publish after 3 tries (others kept publishing). Run again in a minute.")
    sys.exit(1)


def only_meta_changed(fresh):
    """True when every team file differs from the committed one only in _meta.updated_*."""
    for team in fresh:
        r = git('show', f'HEAD:data/{team}.json', check=False)
        if r.returncode != 0: return False
        try:
            old = json.loads(r.stdout)
        except ValueError:
            return False
        if old.get('data') != load_team_file(team)['data']: return False
    return True


# ═══════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════

def main():
    try: sys.stdout.reconfigure(errors='replace')
    except AttributeError: pass

    ap = argparse.ArgumentParser(description="Build / publish the AML dashboard, one team at a time.")
    ap.add_argument('--team', help="sanction, nontms or tms (comma-separated for more than one)")
    ap.add_argument('--excel', help="path to the team's Excel (only with a single --team)")
    ap.add_argument('--publish', action='store_true', help="pull, update, commit and push to GitHub")
    ap.add_argument('--rebuild', action='store_true',
                    help="no Excel: rebuild index.html from data/*.json (e.g. after a template change)")
    ap.add_argument('--yes', action='store_true', help="don't ask before publishing")
    args = ap.parse_args()

    print("=" * 50)
    print("AML Dashboard Generator v3")
    print("=" * 50)

    if args.rebuild:
        build({}, OUTPUT_FILE)
        print("\nRebuilt from data/*.json. Commit index.html (and the template) yourself to publish.")
        return

    teams = [t.strip().lower() for t in (args.team or '').split(',') if t.strip()]
    bad = [t for t in teams if t not in TEAMS]
    if not teams or bad:
        print(f"ERROR: choose --team from: {', '.join(TEAMS)}"); sys.exit(1)
    if args.excel and len(teams) > 1:
        print("ERROR: --excel works with a single --team only."); sys.exit(1)

    fresh, sources = {}, {}
    for team in teams:
        path = args.excel or excel_path(team)
        if not os.path.exists(path):
            print(f"ERROR: {TEAMS[team]['label']} Excel not found:\n  {path}")
            print(f"Put {TEAMS[team]['excel']} in this folder, or set its path in "
                  f"{os.path.basename(LOCAL_PATHS_FILE)} (see README.md).")
            sys.exit(1)
        payload = read_team(team, path)
        if not team_months(team, payload):
            print(f"ERROR: no months with data found in {path}. Nothing to publish."); sys.exit(1)
        fresh[team], sources[team] = payload, path

    if args.publish:
        publish(fresh, sources, args.yes)
    else:
        out = os.path.join(HERE, 'preview.html')
        build(fresh, out)
        print("\nPreview only - nothing saved to data/ or published. Open preview.html to check.")


if __name__ == '__main__':
    main()
