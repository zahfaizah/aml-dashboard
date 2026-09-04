#!/usr/bin/env python3
"""
generate_dashboard.py — AML Sanction Screening Dashboard Generator
Reads Monthly_SS_Dashboard_Template.xlsx, builds a single HTML dashboard
with all months embedded as JSON. The browser handles month switching live.
"""
import pandas as pd
import numpy as np
import json
import os
import sys
import re
from datetime import datetime

# ─── CONFIG ───────────────────────────────────────────────
EXCEL_FILE     = "Monthly_SS_Dashboard_Template.xlsx"
TEMPLATE_FILE  = "AML_Dashboard_template.html"
OUTPUT_FILE    = "Sanction_Screening_Dashboard.html"
HIGHLIGHT_MONTH = "latest"   # "latest" or e.g. "March 2026"
NOTES = {
    "remittance": "This March, there are three weeks worth of data that we haven't been able to process yet due to issues with the BO AML system for uploading data.",
    # "screening": "",  # add notes for other sections as needed
}   # "latest" or e.g. "March 2026"
# ──────────────────────────────────────────────────────────

MONTH_ORDER = ['January','February','March','April','May','June',
               'July','August','September','October','November','December']

def safe_int(v, default=0):
    try:
        if pd.isna(v): return default
        return int(float(v))
    except: return default

def month_sort_key(label):
    """Sort 'January 2026' chronologically."""
    parts = label.strip().split()
    if len(parts) == 2 and parts[0] in MONTH_ORDER:
        return int(parts[1]) * 100 + MONTH_ORDER.index(parts[0])
    return 0

# ═══════════════════════════════════════════════════════════
# READERS — each returns structured data
# ═══════════════════════════════════════════════════════════

def read_screening(path):
    """Parse the Screening sheet — both Realtime (left) and Batch (right)."""
    df = pd.read_excel(path, sheet_name='Screening', header=None)

    data = {'realtime_hits': {}, 'realtime_generated': {}, 'realtime_update': {},
            'batch_alerts': {}, 'batch_truehit': {}, 'batch_update': {}}

    # ── LEFT: Realtime False Hit / True Hit (rows 3-14, cols 0-16) ──
    for i in range(3, 15):
        month = str(df.iloc[i, 0]).strip() if pd.notna(df.iloc[i, 0]) else ''
        if month not in MONTH_ORDER: continue
        fh_pep = safe_int(df.iloc[i, 1]); fh_sanc = safe_int(df.iloc[i, 2])
        fh_int = safe_int(df.iloc[i, 3]); fh_multi = safe_int(df.iloc[i, 4])
        th_pep = safe_int(df.iloc[i, 5]); th_sanc = safe_int(df.iloc[i, 6])
        th_int = safe_int(df.iloc[i, 7]); th_multi = safe_int(df.iloc[i, 8])
        # Entity false/true hits
        eFH_pep = safe_int(df.iloc[i, 9]); eFH_sanc = safe_int(df.iloc[i, 10])
        eFH_int = safe_int(df.iloc[i, 11]); eFH_multi = safe_int(df.iloc[i, 12])
        eTH_pep = safe_int(df.iloc[i, 13]); eTH_sanc = safe_int(df.iloc[i, 14])
        eTH_int = safe_int(df.iloc[i, 15]); eTH_multi = safe_int(df.iloc[i, 16])

        # Year detection from header row — assume 2026 from cell (1,0)
        year = safe_int(df.iloc[1, 0], 2026)
        label = f"{month} {year}"

        data['realtime_hits'][label] = {
            'indiv_fh': {'PEP': fh_pep, 'Sanction': fh_sanc, 'Internal List': fh_int, 'Multitype': fh_multi},
            'indiv_th': {'PEP': th_pep, 'Sanction': th_sanc, 'Internal List': th_int, 'Multitype': th_multi},
            'entity_fh': {'PEP': eFH_pep, 'Sanction': eFH_sanc, 'Internal List': eFH_int, 'Multitype': eFH_multi},
            'entity_th': {'PEP': eTH_pep, 'Sanction': eTH_sanc, 'Internal List': eTH_int, 'Multitype': eTH_multi},
        }

    # ── LEFT: Realtime Generated (rows 20-31, cols 0-6) ──
    for i in range(20, 32):
        month = str(df.iloc[i, 0]).strip() if pd.notna(df.iloc[i, 0]) else ''
        if month not in MONTH_ORDER: continue
        year = safe_int(df.iloc[18, 0], 2026)
        label = f"{month} {year}"
        data['realtime_generated'][label] = {
            'indiv_fh': safe_int(df.iloc[i, 1]),
            'indiv_th': safe_int(df.iloc[i, 2]),
            'indiv_total': safe_int(df.iloc[i, 3]),
            'entity_fh': safe_int(df.iloc[i, 4]),
            'entity_th': safe_int(df.iloc[i, 5]),
            'entity_total': safe_int(df.iloc[i, 6]),
        }

    # ── LEFT: Realtime Screening Update (rows 36-47, cols 0-2) ──
    for i in range(36, 48):
        month = str(df.iloc[i, 0]).strip() if pd.notna(df.iloc[i, 0]) else ''
        if month not in MONTH_ORDER: continue
        gen = safe_int(df.iloc[i, 1])
        closed = safe_int(df.iloc[i, 2])
        if gen == 0 and closed == 0: continue
        year = safe_int(df.iloc[34, 0], 2026) if pd.notna(df.iloc[34, 0]) else 2026
        # Use year from the section header or infer
        label = f"{month} {year}"
        data['realtime_update'][label] = {'generated': gen, 'closed': closed}

    # ── RIGHT: Batch Screening Alert (cols 19-25) ──
    current_month_label = None
    for i in range(3, 22):
        month_val = df.iloc[i, 19]
        type_val = df.iloc[i, 20]

        if pd.notna(month_val):
            dt = pd.to_datetime(month_val, errors='coerce')
            if pd.notna(dt):
                current_month_label = f"{MONTH_ORDER[dt.month-1]} {dt.year}"
                if current_month_label not in data['batch_alerts']:
                    data['batch_alerts'][current_month_label] = {}

        if pd.notna(type_val) and str(type_val).startswith('BATCH') and current_month_label:
            scenario = str(type_val).strip()
            data['batch_alerts'][current_month_label][scenario] = {
                'false_hit': safe_int(df.iloc[i, 21]),
                'true_hit': safe_int(df.iloc[i, 22]),
                'manually_closed': safe_int(df.iloc[i, 23]),
                'automation_closed': safe_int(df.iloc[i, 24]),
                'alert_generated': safe_int(df.iloc[i, 25]),
            }

    # ── RIGHT: Batch True Hit by Type (rows 25-onward) ──
    for i in range(25, 35):
        month_val = df.iloc[i, 19]
        pep = df.iloc[i, 20]
        if pd.isna(month_val) or pd.isna(pep): continue
        dt = pd.to_datetime(month_val, errors='coerce')
        if pd.isna(dt): continue
        label = f"{MONTH_ORDER[dt.month-1]} {dt.year}"
        data['batch_truehit'][label] = {
            'PEP': safe_int(pep),
            'Crime': safe_int(df.iloc[i, 21]),
            'PEP_Crime': safe_int(df.iloc[i, 22]),
            'Total': safe_int(df.iloc[i, 23]),
        }

    # ── RIGHT: Batch Screening Update (rows 32-43) ──
    for i in range(32, 44):
        month = str(df.iloc[i, 19]).strip() if pd.notna(df.iloc[i, 19]) else ''
        if month not in MONTH_ORDER: continue
        gen = safe_int(df.iloc[i, 20])
        closed = safe_int(df.iloc[i, 21])
        if gen == 0 and closed == 0: continue
        # Infer year from batch_alerts keys or batch_truehit keys
        year = 2026  # default
        for existing_label in list(data['batch_alerts'].keys()) + list(data['batch_truehit'].keys()):
            if existing_label.startswith(month):
                year = int(existing_label.split()[-1])
                break
        label = f"{month} {year}"
        data['batch_update'][label] = {'generated': gen, 'closed': closed}

    return data


def read_remittance(path):
    """Parse Remittance sheet, group by Month text column."""
    df = pd.read_excel(path, sheet_name='Remittance', header=0)
    result = {'monthly': {}, 'rows': {}}

    for _, row in df.iterrows():
        month = str(row.get('Month', '')).strip()
        if month not in MONTH_ORDER: continue
        label = f"{month} 2026"

        if label not in result['monthly']:
            result['monthly'][label] = {'data_received': 0, 'false_hit': 0, 'true_hit': 0, 'hit': 0, 'not_hit': 0}
            result['rows'][label] = []

        dr = safe_int(row.get('Total Data Received'))
        fh = safe_int(row.get('False Hit'))
        th = safe_int(row.get('True Hit'))
        hit = safe_int(row.get('HIT'))
        not_hit = safe_int(row.get('NOT HIT'))

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
            'false_hit': fh, 'true_hit': th,
            'data_received': dr,
        })

    return result


def read_submerchant(path):
    """Parse Sub Merchant sheet."""
    df = pd.read_excel(path, sheet_name=' Sub Merchant', header=0)
    result = {}

    for _, row in df.iterrows():
        month = str(row.get('Month', '')).strip()
        if month not in MONTH_ORDER: continue
        label = f"{month} 2026"

        shop_id = safe_int(row.get('Internal Shop ID'))
        total_hit = safe_int(row.get('Total Hit'))

        result[label] = {
            'internal_shop_id': shop_id,
            'total_hit': total_hit,
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
    """Parse Newsletter sheet, group by month from Email Date."""
    df = pd.read_excel(path, sheet_name='Newsletter', header=0)
    df['Email Date'] = pd.to_datetime(df['Email Date'], errors='coerce')
    df = df[df['Email Date'].notna()].copy()
    df['Recorded Names'] = pd.to_numeric(df['Recorded Names'], errors='coerce').fillna(0)
    df['ym'] = df['Email Date'].dt.to_period('M')

    result = {'monthly': {}, 'rows': {}}

    for _, row in df.iterrows():
        dt = row['Email Date']
        label = f"{MONTH_ORDER[dt.month-1]} {dt.year}"

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

    # Compute averages
    for label in result['monthly']:
        m = result['monthly'][label]
        m['avg'] = round(m['total_names'] / m['count'], 1) if m['count'] > 0 else 0

    return result


def read_dttot(path):
    """Parse DTTOT sheet — all historical data."""
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

    return {
        'total_names': sum(r['recorded_names'] for r in rows),
        'total_updates': len(rows),
        'rows': rows,
    }


# ═══════════════════════════════════════════════════════════
# COMPUTE — build the full JSON data store
# ═══════════════════════════════════════════════════════════

def compute_derived(screening, remittance, submerchant, newsletter, dttot):
    """Compute all derived values per month for the JSON store."""
    # Collect all month labels
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

        # ── REALTIME SCREENING ──
        hits = screening['realtime_hits'].get(label, {})
        gen = screening['realtime_generated'].get(label, {})
        upd = screening['realtime_update'].get(label, {})

        if hits:
            ifh = hits['indiv_fh']; ith = hits['indiv_th']
            efh = hits['entity_fh']; eth = hits['entity_th']

            # Total alerts per category
            m['alert_pep'] = ifh['PEP'] + ith['PEP'] + efh['PEP'] + eth['PEP']
            m['alert_sanction'] = ifh['Sanction'] + ith['Sanction'] + efh['Sanction'] + eth['Sanction']
            m['alert_internal'] = ifh['Internal List'] + ith['Internal List'] + efh['Internal List'] + eth['Internal List']
            m['alert_multitype'] = ifh['Multitype'] + ith['Multitype'] + efh['Multitype'] + eth['Multitype']
            m['total_alerts'] = m['alert_pep'] + m['alert_sanction'] + m['alert_internal'] + m['alert_multitype']

            # True hits total
            m['true_hit_indiv'] = sum(ith.values())
            m['true_hit_entity'] = sum(eth.values())
            m['true_hit_total'] = m['true_hit_indiv'] + m['true_hit_entity']

            # False hits total
            m['false_hit_indiv'] = sum(ifh.values())
            m['false_hit_entity'] = sum(efh.values())

            # Category breakdowns
            m['indiv_fh'] = ifh; m['indiv_th'] = ith
            m['entity_fh'] = efh; m['entity_th'] = eth

        if gen:
            m['indiv_generated'] = gen['indiv_total']
            m['entity_generated'] = gen['entity_total']
            m['indiv_fh_total'] = gen['indiv_fh']
            m['indiv_th_total'] = gen['indiv_th']
            m['entity_fh_total'] = gen['entity_fh']
            m['entity_th_total'] = gen['entity_th']

        if upd:
            m['rt_generated'] = upd['generated']
            m['rt_closed'] = upd['closed']
            m['closure_rate'] = round(upd['closed'] / upd['generated'] * 100, 1) if upd['generated'] > 0 else 0

        # True hit rate
        if m.get('true_hit_total') is not None and m.get('rt_generated'):
            m['true_hit_rate'] = round(m['true_hit_total'] / m['rt_generated'] * 100, 2)
        else:
            m['true_hit_rate'] = 0

        # ── BATCH SCREENING ──
        ba = screening['batch_alerts'].get(label, {})
        if ba:
            m['batch_scenarios'] = ba
            m['batch_total_fh'] = sum(s['false_hit'] for s in ba.values())
            m['batch_total_th'] = sum(s['true_hit'] for s in ba.values())
            m['batch_total_mc'] = sum(s['manually_closed'] for s in ba.values())
            m['batch_total_ac'] = sum(s['automation_closed'] for s in ba.values())
            m['batch_total_ag'] = sum(s['alert_generated'] for s in ba.values())

        bth = screening['batch_truehit'].get(label, {})
        if bth:
            m['batch_th_pep'] = bth['PEP']
            m['batch_th_crime'] = bth['Crime']
            m['batch_th_pepcrime'] = bth['PEP_Crime']
            m['batch_th_total'] = bth['Total']

        bupd = screening['batch_update'].get(label, {})
        if bupd:
            m['batch_generated'] = bupd['generated']
            m['batch_closed'] = bupd['closed']

        # ── REMITTANCE ──
        rem = remittance['monthly'].get(label, {})
        if rem:
            m['rem_data_received'] = rem['data_received']
            m['rem_false_hit'] = rem['false_hit']
            m['rem_true_hit'] = rem['true_hit']
            m['rem_hit_rate'] = round((rem['false_hit'] + rem['true_hit']) / rem['data_received'] * 100, 1) if rem['data_received'] > 0 else 0
        m['rem_rows'] = remittance['rows'].get(label, [])

        # ── SUB-MERCHANT ──
        sm = submerchant.get(label, {})
        if sm:
            m['subm_shop_id'] = sm['internal_shop_id']
            m['subm_total_hit'] = sm['total_hit']
            m['subm_hit_rate'] = sm['hit_rate']
            m['subm_detail'] = sm

        # ── NEWSLETTER ──
        news = newsletter['monthly'].get(label, {})
        if news:
            m['news_total'] = news['total_names']
            m['news_count'] = news['count']
            m['news_avg'] = news['avg']
        m['news_rows'] = newsletter['rows'].get(label, [])

        months_data[label] = m

    # Filter out months with no real data
    real_months = []
    for label in all_months:
        d = months_data[label]
        has_data = (d.get('total_alerts', 0) > 0 or
                    d.get('rt_generated', 0) > 0 or
                    d.get('batch_th_total', 0) > 0 or
                    d.get('rem_data_received', 0) > 0 or
                    d.get('news_total', 0) > 0)
        if has_data:
            real_months.append(label)

    return {
        'months': real_months,
        'data': {k: months_data[k] for k in real_months},
        'dttot': dttot,
    }


# ═══════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════

def main():
    if not os.path.exists(EXCEL_FILE):
        print(f"Error: {EXCEL_FILE} not found."); sys.exit(1)
    if not os.path.exists(TEMPLATE_FILE):
        print(f"Error: {TEMPLATE_FILE} not found."); sys.exit(1)

    print("Reading Excel...")
    screening = read_screening(EXCEL_FILE)
    remittance = read_remittance(EXCEL_FILE)
    submerchant = read_submerchant(EXCEL_FILE)
    newsletter = read_newsletter(EXCEL_FILE)
    dttot = read_dttot(EXCEL_FILE)

    print("Computing dashboard data...")
    store = compute_derived(screening, remittance, submerchant, newsletter, dttot)

    # Determine highlight month
    highlight = HIGHLIGHT_MONTH
    if highlight == "latest":
        highlight = store['months'][-1] if store['months'] else ""
    elif highlight not in store['months']:
        print(f"Warning: '{highlight}' not found, using latest.")
        highlight = store['months'][-1] if store['months'] else ""

    store['highlight'] = highlight
    store['notes'] = NOTES
    print(f"Highlight month: {highlight}")
    print(f"Available months: {', '.join(store['months'])}")

    # Serialize to JSON
    json_str = json.dumps(store, ensure_ascii=False, indent=None)

    # Read template and inject JSON
    with open(TEMPLATE_FILE, 'r', encoding='utf-8') as f:
        template = f.read()

    output = template.replace('/*STORE_JSON*/', json_str)

    # Check
    if '/*STORE_JSON*/' in output:
        print("ERROR: STORE_JSON marker not replaced!")
        sys.exit(1)

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(output)

    print(f"\n✓ Dashboard: {OUTPUT_FILE} ({os.path.getsize(OUTPUT_FILE):,} bytes)")
    print(f"  Months: {len(store['months'])}")
    print(f"  DTTOT records: {store['dttot']['total_updates']}")


if __name__ == '__main__':
    main()
