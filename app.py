import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
import io
import re
from datetime import datetime
from collections import Counter

# ── Page config ───────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Data Processing Toolbox",
    page_icon="🚛",
    layout="wide"
)

hide_menu = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
"""
st.markdown(hide_menu, unsafe_allow_html=True)

st.title("🚛 Data Processing Toolbox")

# ── Constants ─────────────────────────────────────────────────────────────────

STATES   = ['VIC', 'NSW', 'QLD', 'SA', 'WA', 'TAS', 'NT', 'ACT']
ZONE_MAP = {
    'YV': 'VIC', 'YN': 'NSW', 'YQ': 'QLD', 'YS': 'SA',
    'YW': 'WA',  'YT': 'TAS', 'YD': 'NT',  'YA': 'ACT'
}
HEADER_COLOR = '1F4E79'

VIP_SERVICE_KEYWORDS   = ['vip', 'elite']
VIP_INSTRUCTION_PHRASES = ['timeslot', 'time slot', 'delivery required', 'required on site']

# ── Helper functions ──────────────────────────────────────────────────────────

def clean_number(val):
    if pd.isna(val):
        return 0.0
    v = str(val).replace(',', '').strip()
    if v.endswith('-'):
        v = '-' + v[:-1]
    try:
        return float(v)
    except:
        return 0.0

def get_state(row):
    origin = str(row['Origin'])
    for state in STATES:
        if state in origin:
            return state
    zone = str(row['Zone'])
    return ZONE_MAP.get(zone[:2].upper(), 'UNKNOWN')

def split_customer_refs(ref_str):
    parts = re.split(r'[,&]+', str(ref_str))
    return [p.strip() for p in parts if p.strip()]

def extract_all_refs(ref_str):
    """Extract all 1800xxxxxx numbers from a reference string."""
    if pd.isna(ref_str):
        return []
    return re.findall(r'1800\d+', str(ref_str))

def is_vip_service(service_val):
    if pd.isna(service_val):
        return False
    s = str(service_val).lower()
    return any(kw in s for kw in VIP_SERVICE_KEYWORDS)

def is_vip_instruction(instr_val):
    if pd.isna(instr_val):
        return False
    s = str(instr_val).lower()
    return any(phrase in s for phrase in VIP_INSTRUCTION_PHRASES)

def parse_txt(uploaded_file):
    df_raw = pd.read_csv(uploaded_file, sep='\t', dtype=str)
    summary_row_data = df_raw.iloc[-1]
    df = df_raw[:-1].copy()
    for col in ['Total (AUD)', 'GST (AUD)', 'Freight', 'LEVY', 'Load Qty', 'Paid Qty']:
        df[col] = df[col].apply(clean_number)
    df['Final Total'] = df['Total (AUD)'] + df['GST (AUD)']
    df['State']       = df.apply(get_state, axis=1)
    client_total      = clean_number(summary_row_data['Total (AUD)'])
    return df, client_total

def parse_consignment_csvs(uploaded_files):
    """Parse one or more consignment CSV files and return a combined DataFrame."""
    frames = []
    for f in uploaded_files:
        try:
            frames.append(pd.read_csv(f, dtype=str))
        except Exception as e:
            st.warning(f"⚠ Could not read {f.name}: {e}")
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)

def enrich_txt_with_csv(df_txt, df_csv):
    """
    Join consignment CSV data onto TXT rows using Delivery/Adjustment → Reference 1.
    Adds columns: Service, Delivery Special Instructions, Carrier, Machship #
    """
    if df_csv.empty:
        df_txt['Service'] = None
        df_txt['Delivery Special Instructions'] = None
        df_txt['Carrier'] = None
        df_txt['Machship #'] = None
        df_txt['VIP'] = False
        df_txt['VIP Reason'] = ''
        return df_txt

    # Build lookup: only index CSV refs that exist in the TXT (ignore unrelated CSV rows)
    txt_refs = set(df_txt['Delivery/Adjustment'].dropna().str.strip())
    ref_lookup = {}
    for _, row in df_csv.iterrows():
        for ref in extract_all_refs(str(row.get('Reference 1', ''))):
            if ref in txt_refs and ref not in ref_lookup:
                ref_lookup[ref] = row

    enriched_rows = []
    for _, txt_row in df_txt.iterrows():
        delivery = str(txt_row.get('Delivery/Adjustment', '')).strip()
        csv_match = ref_lookup.get(delivery)
        if csv_match is not None:
            txt_row = txt_row.copy()
            txt_row['Service']                      = csv_match.get('Service', None)
            txt_row['Delivery Special Instructions'] = csv_match.get('Delivery Special Instructions', None)
            txt_row['Carrier']                      = csv_match.get('Carrier', None)
            txt_row['Machship #']                   = csv_match.get('Machship #', None)
        else:
            txt_row = txt_row.copy()
            txt_row['Service']                      = None
            txt_row['Delivery Special Instructions'] = None
            txt_row['Carrier']                      = None
            txt_row['Machship #']                   = None
        enriched_rows.append(txt_row)

    df_out = pd.DataFrame(enriched_rows)

    # Determine VIP flag and reason
    vip_flags, vip_reasons = [], []
    for _, row in df_out.iterrows():
        reasons = []
        if is_vip_service(row.get('Service')):
            reasons.append(f"Service: {row['Service']}")
        if is_vip_instruction(row.get('Delivery Special Instructions')):
            reasons.append(f"Instructions: {str(row['Delivery Special Instructions'])[:60]}")
        vip_flags.append(bool(reasons))
        vip_reasons.append(' | '.join(reasons))

    df_out['VIP'] = vip_flags
    df_out['VIP Reason'] = vip_reasons
    return df_out

def parse_csv(uploaded_file):
    df = pd.read_csv(uploaded_file, dtype=str)
    for col in ['Quantity', 'Total Cubic', 'Rate Charge', 'Fuel Levy',
                'Rate Charge and Fuel Levy', 'Total Tax', 'Total']:
        df[col] = df[col].apply(clean_number)
    return df

def build_excel(df, client_total):
    """Build the split Excel. Sheets are per-state (not per truck/state)."""
    wb = Workbook()
    wb.remove(wb.active)

    # Determine which extra columns exist from CSV enrichment
    extra_cols = [c for c in ['Carrier', 'Machship #', 'Service',
                               'Delivery Special Instructions', 'VIP']
                  if c in df.columns]

    base_headers = [c for c in df.columns
                    if c not in ('State', 'Final Total', 'VIP', 'VIP Reason',
                                 'Carrier', 'Machship #', 'Service',
                                 'Delivery Special Instructions')]
    headers_with_state = base_headers + ['State'] + extra_cols

    for state in sorted(df['State'].unique()):
        state_df = df[df['State'] == state].copy()
        ws = wb.create_sheet(title=f"{state}")
        for col, header in enumerate(headers_with_state, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill('solid', start_color=HEADER_COLOR)
        for row_idx, (_, row) in enumerate(state_df.reindex(columns=headers_with_state).iterrows(), 2):
            for col_idx, val in enumerate(row, 1):
                ws.cell(row=row_idx, column=col_idx, value=val)
        total_row = len(state_df) + 2
        total_col = headers_with_state.index('Total (AUD)') + 1
        gst_col   = headers_with_state.index('GST (AUD)') + 1
        ws.cell(row=total_row, column=1, value='TOTAL').font = Font(bold=True)
        ws.cell(row=total_row, column=total_col,
                value=f'=SUM({get_column_letter(total_col)}2:{get_column_letter(total_col)}{total_row-1})'
                ).font = Font(bold=True)
        ws.cell(row=total_row, column=gst_col,
                value=f'=SUM({get_column_letter(gst_col)}2:{get_column_letter(gst_col)}{total_row-1})'
                ).font = Font(bold=True)

    # ── Summary sheet ──────────────────────────────────────────────────────
    ws_sum = wb.create_sheet(title='Summary', index=0)

    # Part A: State breakdown
    sum_headers = ['State', 'Rows', 'Total (AUD)', 'GST (AUD)', 'Total (inc GST)']
    for col, header in enumerate(sum_headers, 1):
        cell = ws_sum.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', start_color=HEADER_COLOR)
    data_row = 2
    for state in sorted(df['State'].unique()):
        state_df = df[df['State'] == state]
        total = state_df['Total (AUD)'].sum()
        gst   = state_df['GST (AUD)'].sum()
        ws_sum.cell(row=data_row, column=1, value=state)
        ws_sum.cell(row=data_row, column=2, value=len(state_df))
        ws_sum.cell(row=data_row, column=3, value=round(total, 2))
        ws_sum.cell(row=data_row, column=4, value=round(gst, 2))
        ws_sum.cell(row=data_row, column=5, value=round(total + gst, 2))
        data_row += 1

    grand_row = data_row
    ws_sum.cell(row=grand_row, column=1, value='TOTAL').font = Font(bold=True)
    for col in [2, 3, 4, 5]:
        ws_sum.cell(row=grand_row, column=col,
                    value=f'=SUM({get_column_letter(col)}2:{get_column_letter(col)}{grand_row-1})'
                    ).font = Font(bold=True)
    ws_sum.cell(row=grand_row+2, column=1, value='Statements Total (inc GST)').font = Font(bold=True)
    ws_sum.cell(row=grand_row+2, column=2, value=client_total)
    ws_sum.cell(row=grand_row+3, column=1, value='Check').font = Font(bold=True)
    ws_sum.cell(row=grand_row+3, column=2,
                value=f'=IF(E{grand_row}=B{grand_row+2},"✓ MATCH","⚠ DISCREPANCY")')

    # Part B: VIP breakdown per state (only if VIP column exists)
    if 'VIP' in df.columns:
        vip_headers = ['Delivery/Adjustment', 'Customer Name & Address',
                       'Service', 'Delivery Special Instructions',
                       'Total (AUD)', 'GST (AUD)', 'Total (inc GST)']
        vip_headers_filtered = [h for h in vip_headers if h in df.columns or h in ['Total (inc GST)']]
        num_cols = len(vip_headers_filtered)

        total_col_idx  = vip_headers_filtered.index('Total (AUD)') + 1
        gst_col_idx    = vip_headers_filtered.index('GST (AUD)') + 1
        incgst_col_idx = vip_headers_filtered.index('Total (inc GST)') + 1

        vip_df = df[df['VIP'] == True]
        vip_states = sorted(vip_df['State'].unique())

        vr = grand_row + 6  # start row for first state table

        for state in vip_states:
            state_vip = vip_df[vip_df['State'] == state]

            # Section heading
            heading_cell = ws_sum.cell(row=vr, column=1, value=f'VIP / Special Service Breakdown — {state}')
            heading_cell.font = Font(bold=True, size=12)
            vr += 1

            # Header row
            for col, header in enumerate(vip_headers_filtered, 1):
                cell = ws_sum.cell(row=vr, column=col, value=header)
                cell.font = Font(bold=True, color='FFFFFF')
                cell.fill = PatternFill('solid', start_color='8B0000')
            vr += 1

            data_start = vr
            # Data rows
            for _, row in state_vip.iterrows():
                for col, header in enumerate(vip_headers_filtered, 1):
                    if header == 'Total (inc GST)':
                        ws_sum.cell(row=vr, column=col, value=round(row['Total (AUD)'] + row['GST (AUD)'], 2))
                    else:
                        ws_sum.cell(row=vr, column=col, value=row.get(header, ''))
                vr += 1

            # Totals row
            tot_total  = round(state_vip['Total (AUD)'].sum(), 2)
            tot_gst    = round(state_vip['GST (AUD)'].sum(), 2)
            tot_incgst = round(tot_total + tot_gst, 2)

            ws_sum.cell(row=vr, column=1, value='TOTAL').font = Font(bold=True)
            ws_sum.cell(row=vr, column=total_col_idx,  value=tot_total).font  = Font(bold=True)
            ws_sum.cell(row=vr, column=gst_col_idx,    value=tot_gst).font    = Font(bold=True)
            ws_sum.cell(row=vr, column=incgst_col_idx, value=tot_incgst).font = Font(bold=True)
            vr += 3  # gap before next state table

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ════════════════════════════════════════════════════════════════════════════
# RATE-CARD RECONCILIATION HELPERS  (Tab 2)
# ════════════════════════════════════════════════════════════════════════════
#
# ONLINE / STREAMLIT-CLOUD BUILD
# ------------------------------
# This version stores NOTHING. There is no database. The rate card and all
# invoice/CSV files must be uploaded every session and live only in memory for
# the duration of the run. Nothing is written to disk.
#
# The only hardcoded values are the Fuel Surcharge (FSC) rates below — these
# are not confidential, so they are kept in code for convenience. Edit them
# here whenever the FSC changes.

ZONE_PREFIX_STATE = {
    'V': 'VIC', 'N': 'NSW', 'S': 'SA', 'Q': 'QLD',
    'W': 'WA',  'T': 'TAS', 'D': 'NT', 'A': 'ACT'
}

# ── Fuel Surcharge (FSC) config — the ONLY hardcoded metric ───────────────────
# CURRENT_FSC_PCT is the default applied to every invoice line whose date is not
# covered by a date-range override below. Update it when the current week's FSC
# changes. Date ranges use DD.MM.YYYY.

CURRENT_FSC_PCT = 29.0

FSC_OVERRIDES = [
    {"from": "01.05.2026", "to": "01.05.2026", "pct": 42.0, "note": "1 May"},
    {"from": "04.05.2026", "to": "08.05.2026", "pct": 33.0, "note": "4-8 May"},
    {"from": "11.05.2026", "to": "19.05.2026", "pct": 31.0, "note": "11-19 May"},
    {"from": "20.05.2026", "to": "31.05.2026", "pct": 29.0, "note": "20-31 May"},
]

def get_fsc_overrides_df():
    """Build the FSC override table (in memory) from the hardcoded config."""
    if not FSC_OVERRIDES:
        return pd.DataFrame(columns=['date_from', 'date_to', 'fsc_pct', 'note'])
    return pd.DataFrame([
        {'date_from': o['from'], 'date_to': o['to'], 'fsc_pct': o['pct'], 'note': o.get('note', '')}
        for o in FSC_OVERRIDES
    ])


# ── Rate-card parsing ─────────────────────────────────────────────────────────

def parse_rate_card(file_like):
    """Parse the Rohlig/Sugar transportation rate card → structured dict."""
    wb = load_workbook(file_like, data_only=True)
    parsed = {'effective': None, 'zone_bands': {}, 'metro_band': None,
              'cubic_factor': None, 'basic_charge_kg': None}

    if 'Transportation - Metro,Regional' not in wb.sheetnames:
        return parsed

    ws = wb['Transportation - Metro,Regional']
    eff = ws.cell(row=1, column=2).value or ''
    m = re.search(r'(\d{2}\.\d{2}\.\d{4})', str(eff))
    parsed['effective'] = m.group(1) if m else None

    cf = ws.cell(row=6, column=4).value
    if isinstance(cf, (int, float)):
        parsed['cubic_factor'] = cf

    note = ws.cell(row=3, column=32).value or ''
    bc = re.search(r'\$?\s*([\d.]+)', str(note))
    if bc:
        try: parsed['basic_charge_kg'] = float(bc.group(1))
        except: pass

    # Regional pallet block: cols J-R (10-18); zone=col 13, bands=cols 15-18
    zone_bands = {}
    for r in range(14, ws.max_row + 1):
        zone = ws.cell(row=r, column=13).value
        p1, p25, p611, p12 = (ws.cell(row=r, column=col).value for col in (15, 16, 17, 18))
        if zone and isinstance(p1, (int, float)):
            zone_bands.setdefault(str(zone).strip(), [p1, p25, p611, p12])
    parsed['zone_bands'] = zone_bands

    if zone_bands:
        band_counts = Counter(tuple(v) for v in zone_bands.values())
        parsed['metro_band'] = list(band_counts.most_common(1)[0][0])

    return parsed

# ── Reconciliation logic ────────────────────────────────────────────────────

def _origin_state(origin_text, zone):
    for s in STATES:
        if s in str(origin_text):
            return s
    z = str(zone)
    if len(z) >= 2:
        return ZONE_PREFIX_STATE.get(z[1].upper(), 'UNKNOWN')
    return 'UNKNOWN'

def _dest_state(zone):
    z = str(zone)
    if len(z) >= 2:
        return ZONE_PREFIX_STATE.get(z[1].upper(), 'UNKNOWN')
    return 'UNKNOWN'

def _band_rate(qty, band):
    q = round(qty)
    if q <= 1:  return band[0]
    if q <= 5:  return band[1]
    if q <= 11: return band[2]
    return band[3]

def _parse_date(d):
    """TXT dates are DD.MM.YYYY."""
    for fmt in ('%d.%m.%Y', '%d/%m/%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(str(d).strip(), fmt).date()
        except: pass
    return None

def fsc_for_date(date_obj, fsc_df, default_fsc=None):
    """
    Return (fsc_fraction, source) for a given date.
    Date-range entries take priority; otherwise fall back to the current default FSC.
    default_fsc is a percentage (e.g. 31.0) or None.
    """
    if date_obj is not None and not fsc_df.empty:
        for _, r in fsc_df.iterrows():
            df_ = _parse_date(r['date_from']); dt_ = _parse_date(r['date_to'])
            if df_ and dt_ and df_ <= date_obj <= dt_:
                return float(r['fsc_pct']) / 100.0, 'range'
    if default_fsc is not None:
        return float(default_fsc) / 100.0, 'default'
    return None, None

def reconcile_txt(df_txt, rate_card, fsc_df, default_fsc=None):
    """
    Compare each TXT line against the rate card.
    Auto-verify same-state metro PAL lines; flag interstate/regional/tonne for manual review.
    default_fsc (percentage) is applied to any line whose date isn't covered by a date-range entry.
    """
    band = rate_card.get('metro_band') or [60, 35, 30, 25]
    records = []
    for _, row in df_txt.iterrows():
        zone        = row.get('Zone', '')
        uom         = str(row.get('UOM', '')).strip()
        qty         = clean_number(row.get('Paid Qty', 0))
        chg_rate    = clean_number(row.get('Rate', 0))
        chg_freight = clean_number(row.get('Freight', 0))
        chg_levy    = clean_number(row.get('LEVY', 0))
        date_obj    = _parse_date(row.get('Date'))
        o_state     = _origin_state(row.get('Origin', ''), zone)
        d_state     = _dest_state(zone)
        fsc, fsc_src = fsc_for_date(date_obj, fsc_df, default_fsc)

        is_metro_pal = ('METRO' in str(zone).upper()) and uom == 'PAL' and (o_state == d_state)

        rec = {
            'Date': row.get('Date'),
            'Delivery/Adjustment': row.get('Delivery/Adjustment'),
            'Customer Name & Address': row.get('Customer Name & Address'),
            'Zone': zone, 'UOM': uom, 'Qty': qty,
            'Charged Rate': chg_rate, 'Charged Freight': round(chg_freight, 2),
            'Charged LEVY': round(chg_levy, 2),
            'Charged Total': round(chg_freight + chg_levy, 2),
        }

        if is_metro_pal:
            exp_rate    = _band_rate(qty, band)
            exp_freight = round(exp_rate * qty, 2)
            if fsc is not None:
                exp_levy  = round(exp_freight * fsc, 2)
                exp_total = round(exp_freight + exp_levy, 2)
                levy_diff = round(chg_levy - exp_levy, 2)
            else:
                exp_levy = exp_total = levy_diff = None
            rec.update({
                'Check': 'Auto (metro pallet)',
                'Expected Rate': exp_rate,
                'Expected Freight': exp_freight,
                'FSC %': round(fsc * 100, 2) if fsc is not None else None,
                'FSC Source': fsc_src,
                'Expected LEVY': exp_levy,
                'Expected Total': exp_total,
                'Rate Diff': round(chg_rate - exp_rate, 2),
                'Freight Diff': round(chg_freight - exp_freight, 2),
                'LEVY Diff': levy_diff,
                'Total Diff': round(chg_freight + chg_levy - exp_total, 2) if exp_total is not None else None,
            })
            ft_diff = abs(rec['Freight Diff']) > 0.01
            lv_diff = (levy_diff is not None and abs(levy_diff) > 0.01)
            if rec['FSC %'] is None:
                rec['Status'] = '⚠ No FSC for date'
            elif ft_diff or lv_diff:
                rec['Status'] = '❌ Mismatch'
            else:
                rec['Status'] = '✓ OK'
        else:
            reason = []
            if o_state != d_state: reason.append(f'interstate {o_state}→{d_state}')
            if uom != 'PAL':       reason.append(f'UOM={uom}')
            if 'METRO' not in str(zone).upper(): reason.append('non-metro zone')
            rec.update({
                'Check': 'Manual review (' + ', '.join(reason) + ')',
                'Expected Rate': None, 'Expected Freight': None,
                'FSC %': round(fsc * 100, 2) if fsc is not None else None,
                'FSC Source': fsc_src,
                'Expected LEVY': None, 'Expected Total': None,
                'Rate Diff': None, 'Freight Diff': None, 'LEVY Diff': None, 'Total Diff': None,
                'Status': '🔍 Manual review',
            })
        records.append(rec)

    return pd.DataFrame(records)

def reconcile_csv(df_csv, recon_txt_df):
    """
    Compare Machship CSV 'Total Sell' against the rate-card expected total
    (from the auto-checked TXT lines), matched by 1800xxxxxx delivery ref.
    """
    if df_csv.empty:
        return pd.DataFrame()

    # Map delivery ref → expected total from TXT reconciliation (auto lines only)
    exp_by_ref = {}
    for _, r in recon_txt_df.iterrows():
        ref = str(r.get('Delivery/Adjustment', '')).strip()
        if r.get('Expected Total') is not None:
            exp_by_ref[ref] = r['Expected Total']

    rows = []
    for _, row in df_csv.iterrows():
        refs = re.findall(r'1800\d+', str(row.get('Reference 1', '')))
        total_sell = clean_number(row.get('Total Sell', 0))
        matched_exp = None
        matched_ref = None
        for ref in refs:
            if ref in exp_by_ref:
                matched_exp = exp_by_ref[ref]; matched_ref = ref; break
        rows.append({
            'Machship #': row.get('Machship #'),
            'Reference 1': row.get('Reference 1'),
            'To Name': row.get('To Name'),
            'CSV Total Sell': round(total_sell, 2),
            'Rate Card Expected': matched_exp,
            'Diff': round(total_sell - matched_exp, 2) if matched_exp is not None else None,
            'Status': ('—' if matched_exp is None
                       else ('✓ OK' if abs(total_sell - matched_exp) <= 0.01 else '❌ Mismatch')),
        })
    return pd.DataFrame(rows)

# ── Tabs ──────────────────────────────────────────────────────────────────────

tab1, tab2, tab3 = st.tabs(["📊 RCTI Processor", "🔍 Invoice Reconciliation", "📸 Screenshot to Excel (Beta)"])

# ════════════════════════════════════════════════════════════════════════════
# TAB 1
# ════════════════════════════════════════════════════════════════════════════

with tab1:
    st.write("Upload your `.TXT` statement file and optionally one or more consignment `.CSV` files to enrich with shipment data.")

    col_up1, col_up2 = st.columns(2)
    with col_up1:
        uploaded_txt = st.file_uploader("Statement (.TXT)", type=['txt', 'TXT'], key="tab1_txt")
    with col_up2:
        uploaded_csvs = st.file_uploader(
            "Consignment report(s) (.CSV) — optional, select multiple if needed",
            type=['csv', 'CSV'], key="tab1_csvs", accept_multiple_files=True
        )

    if uploaded_txt:
        if uploaded_txt.size > 20 * 1024 * 1024:
            st.error("⚠ File exceeds the 20MB limit.")
            st.stop()
        with st.spinner("Processing..."):
            df, client_total = parse_txt(uploaded_txt)

            # Enrich with CSV if provided
            if uploaded_csvs:
                df_csv_combined = parse_consignment_csvs(uploaded_csvs)
                df = enrich_txt_with_csv(df, df_csv_combined)
                csv_matched = int(df['Machship #'].notna().sum())
                total = len(df)
                if csv_matched == total:
                    st.success(f"✓ {csv_matched}/{total} shipments on TXT matched with consignment data.")
                else:
                    st.warning(f"⚠ {csv_matched}/{total} shipments on TXT matched with consignment data.")
                    unmatched_df = df[df['Machship #'].isna()]
                    with st.expander(f"Show {len(unmatched_df)} TXT shipment(s) still needing consignment data"):
                        st.dataframe(
                            unmatched_df[['Date', 'Truck', 'State', 'Delivery/Adjustment', 'Customer Name & Address']],
                            use_container_width=True, hide_index=True
                        )
            else:
                df['VIP'] = False
                df['VIP Reason'] = ''

            our_total   = df['Total (AUD)'].sum()
            our_gst     = df['GST (AUD)'].sum()
            our_inc_gst = our_total + our_gst
            match       = round(our_inc_gst, 2) == round(client_total, 2)
            vendor      = df['Vendor'].iloc[0]

        st.divider()
        col1, col2, col3 = st.columns(3)
        col1.metric("Invoice Lines", len(df))
        col2.metric("Trucks", df['Truck'].nunique())
        col3.metric("States", df['State'].nunique())

        st.divider()
        st.subheader("Totals")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total (ex GST)",  f"${our_total:,.2f}")
        col2.metric("GST",             f"${our_gst:,.2f}")
        col3.metric("Total (inc GST)", f"${our_inc_gst:,.2f}")

        if match:
            st.success(f"✓ MATCH — Our total matches statement's total of ${client_total:,.2f}")
        else:
            diff = abs(our_inc_gst - client_total)
            st.error(f"⚠ DISCREPANCY — Difference of ${diff:,.2f} vs statement's total of ${client_total:,.2f}")

        # ── Breakdown by State ────────────────────────────────────────────
        st.divider()
        st.subheader("Breakdown by State")
        state_summary = []
        for state in sorted(df['State'].unique()):
            s = df[df['State'] == state]
            state_summary.append({
                'State': state,
                'Rows': len(s),
                'Total (AUD)': round(s['Total (AUD)'].sum(), 2),
                'GST (AUD)': round(s['GST (AUD)'].sum(), 2),
                'Total (inc GST)': round(s['Total (AUD)'].sum() + s['GST (AUD)'].sum(), 2)
            })
        st.dataframe(pd.DataFrame(state_summary), use_container_width=True, hide_index=True)

        # ── VIP breakdown ─────────────────────────────────────────────────
        st.divider()
        st.subheader("VIP / Special Service Breakdown")

        vip_df = df[df['VIP'] == True].copy()

        if vip_df.empty:
            if uploaded_csvs:
                st.info("No VIP or special service shipments detected.")
            else:
                st.info("Upload consignment CSV(s) above to detect VIP/special service shipments.")
        else:
            st.caption(
                f"**{len(vip_df)} shipment(s)** identified as VIP/special — "
                "uncheck any that should be treated as general service."
            )

            # Session state for exclusions
            if 'vip_excluded' not in st.session_state:
                st.session_state['vip_excluded'] = set()

            display_cols = ['State', 'Delivery/Adjustment', 'Customer Name & Address',
                            'Service', 'Delivery Special Instructions',
                            'Total (AUD)', 'GST (AUD)']
            display_cols = [c for c in display_cols if c in vip_df.columns]

            checked_rows = []
            for idx, row in vip_df.reset_index().iterrows():
                orig_idx = row['index']
                default_checked = orig_idx not in st.session_state['vip_excluded']
                label = (
                    f"**{row.get('Delivery/Adjustment','')}** | "
                    f"{str(row.get('Customer Name & Address',''))[:50]} | "
                    f"{row.get('State','')} | "
                    f"${row.get('Total (AUD)', 0):.2f}"
                )
                checked = st.checkbox(label, value=default_checked, key=f"vip_chk_{orig_idx}")
                if not checked:
                    st.session_state['vip_excluded'].add(orig_idx)
                else:
                    st.session_state['vip_excluded'].discard(orig_idx)
                if checked:
                    checked_rows.append(row)

            if checked_rows:
                active_vip = pd.DataFrame(checked_rows)[display_cols]
                st.divider()
                st.caption("**Active VIP shipments (included in VIP summary):**")

                # Group by State
                vip_state_summary = []
                for state in sorted(active_vip['State'].unique()):
                    sv = active_vip[active_vip['State'] == state]
                    vip_state_summary.append({
                        'State': state,
                        'VIP Shipments': len(sv),
                        'Total (AUD)': round(sv['Total (AUD)'].astype(float).sum(), 2),
                        'GST (AUD)': round(sv['GST (AUD)'].astype(float).sum(), 2),
                        'Total (inc GST)': round(sv['Total (AUD)'].astype(float).sum() + sv['GST (AUD)'].astype(float).sum(), 2)
                    })
                st.dataframe(pd.DataFrame(vip_state_summary), use_container_width=True, hide_index=True)

                with st.expander("Show full VIP shipment list"):
                    st.dataframe(active_vip, use_container_width=True, hide_index=True)

        # ── Download ──────────────────────────────────────────────────────
        st.divider()
        # Apply exclusions to df before building Excel
        if 'vip_excluded' in st.session_state and st.session_state['vip_excluded']:
            df_for_excel = df.copy()
            vip_indices = df[df['VIP'] == True].index
            for i, orig_idx in enumerate(vip_indices):
                if orig_idx in st.session_state['vip_excluded']:
                    df_for_excel.at[orig_idx, 'VIP'] = False
                    df_for_excel.at[orig_idx, 'VIP Reason'] = ''
        else:
            df_for_excel = df

        buffer = build_excel(df_for_excel, client_total)
        filename = f'{vendor}_split.xlsx'

        st.download_button(
            label="⬇️ Download Excel",
            data=buffer,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ════════════════════════════════════════════════════════════════════════════
# TAB 2
# ════════════════════════════════════════════════════════════════════════════

with tab2:
    st.write(
        "Compare the **carrier invoice (.TXT)** against the **agreed sell rate card** to catch over/undercharges. "
        "Optionally add the **Machship shipment extract (.CSV)** to also check invoiced sell totals."
    )
    st.caption("🔒 Online build — nothing is stored. All files must be uploaded each session; they're held in memory only.")

    # ── Uploaders ─────────────────────────────────────────────────────────
    c1, c2, c3 = st.columns(3)
    with c1:
        recon_txt = st.file_uploader("① Carrier invoice (.TXT)", type=['txt', 'TXT'], key="tab2_txt")
    with c2:
        recon_csvs = st.file_uploader(
            "② Machship extract (.CSV) — optional",
            type=['csv', 'CSV'], key="tab2_csv", accept_multiple_files=True
        )
    with c3:
        rate_card_file = st.file_uploader("③ Sell rate card (.xlsx) — required", type=['xlsx', 'XLSX'], key="tab2_ratecard")

    # ── Rate card: must be uploaded every session (no storage) ────────────
    rate_card = None
    if rate_card_file is not None:
        try:
            rate_card = parse_rate_card(rate_card_file)
            st.success(
                f"✓ Rate card loaded (this session only) — effective {rate_card.get('effective','?')}, "
                f"metro pallet bands {rate_card.get('metro_band')}."
            )
        except Exception as e:
            st.error(f"⚠ Could not parse rate card: {e}")
    else:
        st.warning("Upload the sell rate card in ③ to enable the checks. It is not stored, so it must be uploaded each session.")

    # ── FSC (hardcoded, not confidential) ─────────────────────────────────
    st.divider()
    st.subheader("⛽ Fuel Surcharge (FSC)")
    st.caption(
        "FSC is hardcoded in the app (the only non-confidential metric). "
        "To change it permanently, edit `CURRENT_FSC_PCT` / `FSC_OVERRIDES` at the top of the file. "
        "You can also override it just for this session below."
    )

    fcol1, fcol2 = st.columns([1, 3])
    with fcol1:
        session_default = st.number_input(
            "Default FSC % (this session)", min_value=0.0, max_value=100.0, step=0.01,
            value=float(CURRENT_FSC_PCT), key="fsc_session_default"
        )
    with fcol2:
        st.caption(
            f"Hardcoded default is **{CURRENT_FSC_PCT}%**. The value here applies to every line "
            "whose date isn't covered by a hardcoded date-range override (shown below). "
            "Changing it here affects this session only — nothing is saved."
        )

    fsc_df = get_fsc_overrides_df()
    default_fsc = float(session_default)

    if not fsc_df.empty:
        st.markdown("**Hardcoded date-range overrides:**")
        show = fsc_df[['date_from', 'date_to', 'fsc_pct', 'note']].rename(
            columns={'date_from': 'From', 'date_to': 'To', 'fsc_pct': 'FSC %', 'note': 'Note'})
        st.dataframe(show, use_container_width=True, hide_index=True)
    else:
        st.caption("No date-range overrides configured — the default FSC applies to all lines.")

    # ── Run reconciliation ────────────────────────────────────────────────
    if recon_txt and rate_card:
        if recon_txt.size > 20 * 1024 * 1024:
            st.error("⚠ TXT file exceeds the 20MB limit.")
            st.stop()

        with st.spinner("Reconciling against rate card..."):
            df_txt, client_total = parse_txt(recon_txt)
            recon_df = reconcile_txt(df_txt, rate_card, fsc_df, default_fsc)

            df_csv_combined = parse_consignment_csvs(recon_csvs) if recon_csvs else pd.DataFrame()
            csv_recon_df = reconcile_csv(df_csv_combined, recon_df) if not df_csv_combined.empty else pd.DataFrame()

        # Helper: implied FSC from invoice for cross-checking
        st.divider()
        with st.expander("🔎 Implied FSC from invoice (LEVY ÷ Freight) — to cross-check the hardcoded FSC"):
            tmp = df_txt.copy()
            tmp['_f'] = tmp['Freight'].apply(clean_number)
            tmp['_l'] = tmp['LEVY'].apply(clean_number)
            tmp = tmp[tmp['_f'] > 0]
            tmp['Implied FSC %'] = (tmp['_l'] / tmp['_f'] * 100).round(2)
            implied = tmp.groupby('Date')['Implied FSC %'].agg(lambda s: ', '.join(map(str, sorted(s.unique())))).reset_index()
            st.dataframe(implied, use_container_width=True, hide_index=True)

        # ── Metrics ────────────────────────────────────────────────────────
        auto = recon_df[recon_df['Check'].str.startswith('Auto')]
        mismatches = recon_df[recon_df['Status'] == '❌ Mismatch']
        st.divider()
        st.subheader("① Invoice (.TXT) vs Rate Card")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Invoice lines", len(recon_df))
        m2.metric("Auto-checked", len(auto))
        m3.metric("Mismatches", len(mismatches))
        manual_n = (recon_df['Status'] == '🔍 Manual review').sum()
        m4.metric("Manual review", int(manual_n))

        auto_var = auto.dropna(subset=['Total Diff'])
        if not auto_var.empty:
            tot_chg = round(auto_var['Charged Total'].sum(), 2)
            tot_exp = round(auto_var['Expected Total'].sum(), 2)
            var = round(tot_chg - tot_exp, 2)
            v1, v2, v3 = st.columns(3)
            v1.metric("Charged (auto lines)",  f"${tot_chg:,.2f}")
            v2.metric("Expected (auto lines)", f"${tot_exp:,.2f}")
            v3.metric("Variance", f"${var:,.2f}", delta=f"{'Overcharged' if var>0 else 'Undercharged' if var<0 else 'Even'}")
        else:
            tot_chg = tot_exp = var = 0.0

        # ── Highlighted table ────────────────────────────────────────────────
        st.caption("Rows where the invoice differs from the rate card are highlighted red. 🔍 = needs manual review (interstate / regional / tonne).")

        display_cols = ['Status', 'Check', 'Date', 'Delivery/Adjustment', 'Customer Name & Address',
                        'Zone', 'UOM', 'Qty',
                        'Charged Rate', 'Expected Rate', 'Rate Diff',
                        'Charged Freight', 'Expected Freight', 'Freight Diff',
                        'FSC %', 'FSC Source', 'Charged LEVY', 'Expected LEVY', 'LEVY Diff',
                        'Charged Total', 'Expected Total', 'Total Diff']
        display_cols = [c for c in display_cols if c in recon_df.columns]

        def style_recon(df):
            styles = pd.DataFrame('', index=df.index, columns=df.columns)
            for i in range(len(df)):
                status = df.iloc[i]['Status']
                if status == '❌ Mismatch':
                    styles.iloc[i, :] = 'background-color: #4a1a1a; color: #ffb3b3'
                elif status == '🔍 Manual review':
                    styles.iloc[i, :] = 'background-color: #3a3320; color: #e0d4a8'
                elif status == '⚠ No FSC for date':
                    styles.iloc[i, :] = 'background-color: #3a2e1a; color: #e8c98a'
                elif status == '✓ OK':
                    styles.iloc[i, :] = 'background-color: #14241a; color: #a8d4b4'
            return styles

        st.dataframe(recon_df[display_cols].style.apply(style_recon, axis=None),
                     use_container_width=True, hide_index=True)

        # ── CSV vs rate card ──────────────────────────────────────────────────
        if not csv_recon_df.empty:
            st.divider()
            st.subheader("② Machship (.CSV) Sell Totals vs Rate Card")
            matched_csv = csv_recon_df[csv_recon_df['Rate Card Expected'].notna()]
            csv_mism = matched_csv[matched_csv['Status'] == '❌ Mismatch']
            cm1, cm2, cm3 = st.columns(3)
            cm1.metric("CSV rows", len(csv_recon_df))
            cm2.metric("Matched to rate card", len(matched_csv))
            cm3.metric("Mismatches", len(csv_mism))
            st.caption("Compares CSV 'Total Sell' to the rate-card expected total (auto lines). Mismatches highlighted red.")

            def style_csv(df):
                styles = pd.DataFrame('', index=df.index, columns=df.columns)
                for i in range(len(df)):
                    if df.iloc[i]['Status'] == '❌ Mismatch':
                        styles.iloc[i, :] = 'background-color: #4a1a1a; color: #ffb3b3'
                    elif df.iloc[i]['Status'] == '✓ OK':
                        styles.iloc[i, :] = 'background-color: #14241a; color: #a8d4b4'
                return styles

            st.dataframe(csv_recon_df.style.apply(style_csv, axis=None),
                         use_container_width=True, hide_index=True)

        # ── Download Excel ──────────────────────────────────────────────────
        st.divider()

        def build_recon_card_excel(recon_df, csv_recon_df):
            wb = Workbook(); wb.remove(wb.active)
            red  = PatternFill('solid', start_color='C00000')
            grn  = PatternFill('solid', start_color='1F7A1F')
            yel  = PatternFill('solid', start_color='B8860B')
            hdrf = PatternFill('solid', start_color='1F4E79')

            def write_df(ws, df, title):
                ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=13)
                cols = list(df.columns)
                for c, col in enumerate(cols, 1):
                    cell = ws.cell(row=2, column=c, value=col)
                    cell.font = Font(bold=True, color='FFFFFF'); cell.fill = hdrf
                for r, (_, row) in enumerate(df.iterrows(), 3):
                    status = row.get('Status', '')
                    for c, col in enumerate(cols, 1):
                        val = row[col]
                        if isinstance(val, float): val = round(val, 2)
                        cell = ws.cell(row=r, column=c, value=val)
                        if status == '❌ Mismatch':
                            cell.fill = red; cell.font = Font(color='FFFFFF')
                        elif status == '🔍 Manual review':
                            cell.fill = yel; cell.font = Font(color='FFFFFF')
                        elif status == '✓ OK':
                            cell.fill = grn; cell.font = Font(color='FFFFFF')

            write_df(wb.create_sheet("TXT vs Rate Card"), recon_df, "Invoice (.TXT) vs Rate Card")
            if not csv_recon_df.empty:
                write_df(wb.create_sheet("CSV vs Rate Card"), csv_recon_df, "Machship (.CSV) vs Rate Card")
            buf = io.BytesIO(); wb.save(buf); buf.seek(0)
            return buf

        excel_buf = build_recon_card_excel(recon_df, csv_recon_df)
        st.download_button(
            "⬇️ Download Reconciliation Report (.xlsx)",
            data=excel_buf,
            file_name="rate_card_reconciliation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


# ════════════════════════════════════════════════════════════════════════════
# TAB 3 — Screenshot to Excel (free OCR via pytesseract)
# ════════════════════════════════════════════════════════════════════════════

with tab3:
    st.write("Upload one or more screenshots of your charge table and download the extracted data as Excel.")

    screenshots = st.file_uploader(
        "Upload screenshots (.png, .jpg)",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="tab3_screenshots"
    )

    COLUMNS = [
        "Charge", "Description", "Bran", "Depa", "Cost",
        "OS Cost Amt", "Estimated Cost", "Local Cost Amt", "Creditor",
        "Cost Recognition", "Posted", "Apt", "Sell", "OS Sell Amt",
        "Estimated Revenue", "Local Sell Amt", "Debtor",
        "Sell Recognition", "Posted (Sell)"
    ]

    COL_X = [109, 256, 799, 883, 978, 1072, 1289, 1521, 1759,
              1924, 2176, 2293, 2417, 2525, 2728, 2967, 3152, 3352, 3596]

    def assign_col(left):
        dists = [abs(left - cx) for cx in COL_X]
        return COLUMNS[dists.index(min(dists))]

    def clean_charge(val):
        if not isinstance(val, str): return val
        m = re.search(r'\b(W[A-Z]{3})\b', val)
        return m.group(1) if m else val.strip()

    def clean_numeric_ocr(val):
        if not isinstance(val, str): return None
        val = val.replace(",", ".").strip()
        try: return float(val)
        except: return None

    def extract_table_ocr(image_bytes):
        from PIL import Image, ImageEnhance
        import pytesseract

        img = Image.open(io.BytesIO(image_bytes))
        w, h = img.size
        img = img.resize((w * 2, h * 2), Image.LANCZOS)
        img = ImageEnhance.Contrast(img).enhance(1.5)

        df_ocr = pytesseract.image_to_data(img, output_type=pytesseract.Output.DATAFRAME, config="--psm 6")
        df_ocr = df_ocr[df_ocr["conf"] > 0].copy()
        df_ocr["text"] = df_ocr["text"].astype(str).str.strip()
        df_ocr = df_ocr[df_ocr["text"].str.len() > 0].reset_index(drop=True)

        df_ocr = df_ocr.sort_values("top")
        row_labels, current_label, last_top = [], 0, -100
        for _, r in df_ocr.iterrows():
            if r["top"] - last_top > 30:
                current_label += 1
                last_top = r["top"]
            row_labels.append(current_label)
        df_ocr["row_id"] = row_labels

        rows_out = []
        for rid in sorted(df_ocr["row_id"].unique()):
            if rid == 1: continue
            row_words = df_ocr[df_ocr["row_id"] == rid].sort_values("left")
            row_dict = {}
            for _, w in row_words.iterrows():
                col = assign_col(w["left"])
                row_dict[col] = (row_dict.get(col, "") + " " + w["text"]).strip()
            if row_dict:
                rows_out.append(row_dict)

        return pd.DataFrame(rows_out, columns=COLUMNS)

    if screenshots:
        if st.button("🔍 Extract Table Data", key="tab3_extract"):
            all_rows = []
            progress = st.progress(0)
            status   = st.empty()

            for i, img_file in enumerate(screenshots):
                status.text(f"Processing {img_file.name} ({i+1}/{len(screenshots)})...")
                try:
                    img_bytes = img_file.read()
                    df_img    = extract_table_ocr(img_bytes)
                    df_img["_source"] = img_file.name
                    all_rows.append(df_img)
                except Exception as e:
                    st.warning(f"⚠ Could not extract from {img_file.name}: {e}")
                progress.progress((i + 1) / len(screenshots))

            status.empty()
            progress.empty()

            if all_rows:
                df_shots = pd.concat(all_rows, ignore_index=True)
                df_shots["Charge"] = df_shots["Charge"].apply(clean_charge)
                for col in ["OS Cost Amt", "Estimated Cost", "Local Cost Amt",
                            "OS Sell Amt", "Estimated Revenue", "Local Sell Amt"]:
                    df_shots[col] = df_shots[col].apply(clean_numeric_ocr)
                    df_shots[col] = pd.to_numeric(df_shots[col], errors="coerce")
                st.session_state["screenshot_df"] = df_shots
                st.success(f"✓ Extracted {len(df_shots)} rows from {len(screenshots)} screenshot(s)")

    if "screenshot_df" in st.session_state:
        df_shots     = st.session_state["screenshot_df"]
        display_cols = [c for c in COLUMNS if c in df_shots.columns]

        st.divider()
        st.subheader("Extracted Data")
        st.dataframe(df_shots[display_cols], use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("OS Cost Amt by Creditor")
        cost_summary = (
            df_shots.groupby("Creditor")["OS Cost Amt"].sum().reset_index()
            .rename(columns={"OS Cost Amt": "Total OS Cost Amt"})
            .sort_values("Total OS Cost Amt", ascending=False)
        )
        cost_summary["Total OS Cost Amt"] = cost_summary["Total OS Cost Amt"].round(2)
        st.dataframe(cost_summary, use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("Estimated Revenue by Debtor")
        rev_summary = (
            df_shots.groupby("Debtor")["Estimated Revenue"].sum().reset_index()
            .rename(columns={"Estimated Revenue": "Total Estimated Revenue"})
            .sort_values("Total Estimated Revenue", ascending=False)
        )
        rev_summary["Total Estimated Revenue"] = rev_summary["Total Estimated Revenue"].round(2)
        st.dataframe(rev_summary, use_container_width=True, hide_index=True)

        st.divider()

        def build_screenshot_excel(df, cost_sum, rev_sum):
            wb = Workbook()
            wb.remove(wb.active)

            def write_sheet(ws, df_in, title):
                ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=13)
                for c, col in enumerate(df_in.columns, 1):
                    cell = ws.cell(row=2, column=c, value=col)
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", start_color="1F4E79")
                for r, (_, row) in enumerate(df_in.iterrows(), 3):
                    for c, val in enumerate(row, 1):
                        ws.cell(row=r, column=c, value=val)

            write_sheet(wb.create_sheet("Extracted Data"),    df[display_cols], "Extracted Charge Data")
            write_sheet(wb.create_sheet("Cost by Creditor"),  cost_sum,         "OS Cost Amt by Creditor")
            write_sheet(wb.create_sheet("Revenue by Debtor"), rev_sum,          "Estimated Revenue by Debtor")

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            return buf

        excel_buf = build_screenshot_excel(df_shots, cost_summary, rev_summary)
        st.download_button(
            label="⬇️ Download Excel",
            data=excel_buf,
            file_name="screenshot_charges.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
