import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
import io
import re

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

    # Build lookup: ref_number → first matching CSV row
    ref_lookup = {}
    for _, row in df_csv.iterrows():
        for ref in extract_all_refs(str(row.get('Reference 1', ''))):
            if ref not in ref_lookup:
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
                               'Delivery Special Instructions', 'VIP', 'VIP Reason']
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

    # Part B: VIP breakdown by state (only if VIP column exists)
    if 'VIP' in df.columns:
        vip_start = grand_row + 6
        ws_sum.cell(row=vip_start, column=1, value='VIP / Special Service Breakdown').font = Font(bold=True, size=12)
        vip_headers = ['State', 'Delivery/Adjustment', 'Customer Name & Address',
                       'Service', 'Delivery Special Instructions', 'VIP Reason',
                       'Total (AUD)', 'GST (AUD)', 'Total (inc GST)']
        vip_headers_filtered = [h for h in vip_headers if h in df.columns or h in ['Total (inc GST)']]
        for col, header in enumerate(vip_headers_filtered, 1):
            cell = ws_sum.cell(row=vip_start+1, column=col, value=header)
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill('solid', start_color='8B0000')

        vip_df = df[df['VIP'] == True].sort_values('State')
        vr = vip_start + 2
        for _, row in vip_df.iterrows():
            for col, header in enumerate(vip_headers_filtered, 1):
                if header == 'Total (inc GST)':
                    ws_sum.cell(row=vr, column=col, value=round(row['Total (AUD)'] + row['GST (AUD)'], 2))
                else:
                    ws_sum.cell(row=vr, column=col, value=row.get(header, ''))
            vr += 1

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

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
                csv_matched = df['Machship #'].notna().sum()
                st.info(f"📦 {len(uploaded_csvs)} consignment CSV(s) loaded — matched {csv_matched}/{len(df)} TXT rows with shipment data.")
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
                            'Service', 'Delivery Special Instructions', 'VIP Reason',
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
    st.write("Upload both files to reconcile charges (.TXT) against invoice (.CSV).")

    col1, col2 = st.columns(2)
    with col1:
        recon_txt = st.file_uploader("Statement (.TXT)", type=['txt', 'TXT'], key="tab2_txt")
    with col2:
        recon_csvs = st.file_uploader(
            "invoice(s) (.CSV) — select multiple if needed",
            type=['csv', 'CSV'], key="tab2_csv", accept_multiple_files=True
        )

    if recon_txt and recon_csvs:
        MAX_MB = 20
        oversized = [f.name for f in recon_csvs if f.size > MAX_MB * 1024 * 1024]
        if oversized:
            st.error(f"⚠ Files exceed 20MB limit: {', '.join(oversized)}")
            st.stop()

        with st.spinner("Reconciling..."):
            df_txt, _ = parse_txt(recon_txt)
            csv_frames = [parse_csv(f) for f in recon_csvs]
            df_csv = pd.concat(csv_frames, ignore_index=True)
            invoice_nums = df_csv['Invoice Number'].unique().tolist() if 'Invoice Number' in df_csv.columns else []
            st.info(f"📄 {len(recon_csvs)} CSV file(s) loaded — Invoice(s): {', '.join(str(i) for i in invoice_nums)}")

            txt_deliveries = set(df_txt['Delivery/Adjustment'].dropna().str.strip().unique())
            csv_deliveries = set()
            for ref in df_csv['Customer Reference'].dropna():
                for r in split_customer_refs(ref):
                    csv_deliveries.add(r)

            df_txt['Reconciliation Status'] = df_txt['Delivery/Adjustment'].apply(
                lambda x: '✓ Matched' if x in csv_deliveries else 'Invoice not found on warehouse invoice'
            )
            df_csv['Reconciliation Status'] = df_csv['Customer Reference'].apply(
                lambda x: '✓ Matched' if any(r in txt_deliveries for r in split_customer_refs(x))
                          else 'Invoice not found on warehouse invoice'
            )
            df_matched_txt = df_txt[df_txt['Reconciliation Status'] == '✓ Matched'].copy()
            df_matched_csv = df_csv[df_csv['Reconciliation Status'] == '✓ Matched'].copy()

        # State filter
        st.divider()
        available_states = sorted(df_txt['State'].dropna().unique().tolist())
        selected_state   = st.selectbox("Filter by State", ['All'] + available_states, index=0, key="tab2_state_filter")

        if selected_state == 'All':
            df_txt_filtered  = df_txt.copy()
            df_matched_txt_f = df_matched_txt.copy()
            df_matched_csv_f = df_matched_csv.copy()
        else:
            df_txt_filtered  = df_txt[df_txt['State'] == selected_state].copy()
            filtered_dels    = set(df_txt_filtered['Delivery/Adjustment'].dropna().str.strip())
            df_matched_txt_f = df_matched_txt[df_matched_txt['Delivery/Adjustment'].isin(filtered_dels)].copy()
            df_matched_csv_f = df_matched_csv[
                df_matched_csv['Customer Reference'].apply(
                    lambda x: any(r in filtered_dels for r in split_customer_refs(str(x)))
                )
            ].copy()

        # Section 1
        st.divider()
        st.subheader("① Line Item Check")
        unmatched_txt = df_txt_filtered[df_txt_filtered['Reconciliation Status'] != '✓ Matched']
        unmatched_csv = df_csv[df_csv['Reconciliation Status'] != '✓ Matched']

        if unmatched_txt.empty and unmatched_csv.empty:
            st.success("✓ All invoice lines found")
        else:
            if not unmatched_txt.empty:
                st.warning(f"⚠ {len(unmatched_txt)} TXT line(s) not found in .CSV invoice")
                st.dataframe(unmatched_txt[['Delivery/Adjustment','State','Origin','Total (AUD)','GST (AUD)','Reconciliation Status']], use_container_width=True, hide_index=True)
            if not unmatched_csv.empty:
                st.warning(f"⚠ {len(unmatched_csv)} CSV line(s) not found in .TXT statement")
                st.dataframe(unmatched_csv[['Customer Reference','Total','Reconciliation Status']], use_container_width=True, hide_index=True)

        # Section 2
        st.divider()
        st.subheader("② Comparison Table")
        st.caption("Matched lines only — TXT vs CSV side by side.")

        comparison_rows = []
        for _, csv_row in df_matched_csv_f.iterrows():
            refs        = split_customer_refs(csv_row['Customer Reference'])
            matched_txt = df_matched_txt_f[df_matched_txt_f['Delivery/Adjustment'].isin(refs)]
            comparison_rows.append({
                'Customer Reference':  csv_row['Customer Reference'],
                'TXT Load Qty':        round(matched_txt['Load Qty'].sum(), 3),
                'CSV Total Cubic':     csv_row['Total Cubic'],
                'TXT Paid Qty':        round(matched_txt['Paid Qty'].sum(), 3),
                'CSV Quantity':        csv_row['Quantity'],
                'TXT Freight':         round(matched_txt['Freight'].sum(), 2),
                'CSV Rate Charge':     csv_row['Rate Charge'],
                'TXT LEVY':            round(matched_txt['LEVY'].sum(), 2),
                'CSV Fuel Levy':       csv_row['Fuel Levy'],
                'TXT Total (AUD)':     round(matched_txt['Total (AUD)'].sum(), 2),
                'CSV Rate+Fuel Levy':  csv_row['Rate Charge and Fuel Levy'],
                'TXT GST':             round(matched_txt['GST (AUD)'].sum(), 2),
                'CSV Total Tax':       csv_row['Total Tax'],
                'TXT Final Total':     round(matched_txt['Final Total'].sum(), 2),
                'CSV Total':           csv_row['Total'],
            })

        df_compare = pd.DataFrame(comparison_rows)

        compare_pairs = [
            ('TXT Load Qty','CSV Total Cubic'),('TXT Paid Qty','CSV Quantity'),
            ('TXT Freight','CSV Rate Charge'),('TXT LEVY','CSV Fuel Levy'),
            ('TXT Total (AUD)','CSV Rate+Fuel Levy'),('TXT GST','CSV Total Tax'),
            ('TXT Final Total','CSV Total'),
        ]
        csv_to_txt = {c: t for t, c in compare_pairs}

        TXT_BG  = 'background-color: #1a2a45; color: #a8c4e0'
        CSV_BG  = 'background-color: #1a3a35; color: #a8d4cc'
        CSV_RED = 'background-color: #a83232; color: white'
        CSV_GRN = 'background-color: #1a7a1a; color: white'

        def style_comparison(df):
            styles = pd.DataFrame('', index=df.index, columns=df.columns)
            for col in df.columns:
                if col == 'Customer Reference': continue
                elif col.startswith('TXT'): styles[col] = TXT_BG
                elif col.startswith('CSV'):
                    txt_col = csv_to_txt.get(col)
                    if txt_col and txt_col in df.columns:
                        for i in range(len(df)):
                            try:
                                cv, tv = float(df[col].iloc[i]), float(df[txt_col].iloc[i])
                                styles.iloc[i, df.columns.get_loc(col)] = CSV_GRN if tv > cv else (CSV_RED if cv > tv else CSV_BG)
                            except:
                                styles.iloc[i, df.columns.get_loc(col)] = CSV_BG
                    else: styles[col] = CSV_BG
            return styles

        total_txt     = len(df_matched_txt_f)
        total_csv     = len(df_matched_csv_f)
        exact_matches = sum(
            1 for _, row in df_compare.iterrows()
            if all(float(row[cc]) == float(row[tc]) for tc, cc in [('TXT Freight','CSV Rate Charge'),('TXT LEVY','CSV Fuel Levy'),('TXT Final Total','CSV Total')])
        )
        col1, col2, col3 = st.columns(3)
        col1.metric("TXT Lines Matched", total_txt)
        col2.metric("CSV Lines Matched", total_csv)
        col3.metric("Exact Value Matches", f"{exact_matches} / {total_csv}")

        st.caption("🔵 TXT columns (dark blue)  |  🟩 CSV columns (teal = match, green = higher, red = lower)")
        st.dataframe(df_compare.style.apply(style_comparison, axis=None), use_container_width=True, hide_index=True)

        # Section 3
        st.divider()
        st.subheader("③ Discrepancies")

        pairs = [
            ('TXT Freight','CSV Rate Charge','Rate Charge'),
            ('TXT LEVY','CSV Fuel Levy','Fuel Levy'),
            ('TXT GST','CSV Total Tax','Total Tax'),
            ('TXT Final Total','CSV Total','Total'),
        ]
        discrepancy_rows = []
        for _, row in df_compare.iterrows():
            diffs = {}
            for tc, cc, label in pairs:
                diff = round(float(row[cc]) - float(row[tc]), 2)
                if diff != 0:
                    diffs[f'TXT {label}'] = row[tc]
                    diffs[f'CSV {label}'] = row[cc]
                    diffs[f'Diff {label}'] = diff
            if diffs:
                discrepancy_rows.append({'Customer Reference': row['Customer Reference'], **diffs})

        if not discrepancy_rows:
            st.success("✓ No discrepancies found")
        else:
            st.error(f"⚠ {len(discrepancy_rows)} row(s) with discrepancies")
            df_disc = pd.DataFrame(discrepancy_rows)

            def highlight_diffs(df):
                styles = pd.DataFrame('', index=df.index, columns=df.columns)
                for col in df.columns:
                    if col == 'Customer Reference': continue
                    elif col.startswith('TXT '): styles[col] = 'background-color: #1a2a45; color: #a8c4e0'
                    elif col.startswith('CSV '): styles[col] = 'background-color: #1a3a35; color: #a8d4cc'
                    elif col.startswith('Diff '):
                        for i, val in enumerate(df[col]):
                            if pd.notna(val):
                                try:
                                    styles.iloc[i, df.columns.get_loc(col)] = (
                                        'background-color: #1a7a1a; color: white' if float(val) < 0
                                        else 'background-color: #a83232; color: white' if float(val) > 0 else ''
                                    )
                                except: pass
                return styles

            st.caption("🔵 TXT  |  🟩 CSV  |  🟢 Diff negative (CSV lower)  |  🔴 Diff positive (CSV higher)")
            st.dataframe(df_disc.style.apply(highlight_diffs, axis=None), use_container_width=True, hide_index=True)

        # Download
        st.divider()
        st.subheader("⬇️ Download Reconciliation Report")

        def build_recon_tables(df_txt_in, df_csv_in, df_matched_txt_in, df_matched_csv_in):
            unmatched_t = df_txt_in[df_txt_in['Reconciliation Status'] != '✓ Matched'][['Delivery/Adjustment','State','Origin','Total (AUD)','GST (AUD)','Reconciliation Status']].copy()
            unmatched_c = df_csv_in[df_csv_in['Reconciliation Status'] != '✓ Matched'][['Customer Reference','Total','Reconciliation Status']].copy()
            comp_rows = []
            for _, csv_row in df_matched_csv_in.iterrows():
                refs = split_customer_refs(csv_row['Customer Reference'])
                mtxt = df_matched_txt_in[df_matched_txt_in['Delivery/Adjustment'].isin(refs)]
                comp_rows.append({
                    'Customer Reference': csv_row['Customer Reference'],
                    'TXT Load Qty': round(mtxt['Load Qty'].sum(), 3),
                    'CSV Total Cubic': csv_row['Total Cubic'],
                    'TXT Paid Qty': round(mtxt['Paid Qty'].sum(), 3),
                    'CSV Quantity': csv_row['Quantity'],
                    'TXT Freight': round(mtxt['Freight'].sum(), 2),
                    'CSV Rate Charge': csv_row['Rate Charge'],
                    'TXT LEVY': round(mtxt['LEVY'].sum(), 2),
                    'CSV Fuel Levy': csv_row['Fuel Levy'],
                    'TXT Total (AUD)': round(mtxt['Total (AUD)'].sum(), 2),
                    'CSV Rate+Fuel Levy': csv_row['Rate Charge and Fuel Levy'],
                    'TXT GST': round(mtxt['GST (AUD)'].sum(), 2),
                    'CSV Total Tax': csv_row['Total Tax'],
                    'TXT Final Total': round(mtxt['Final Total'].sum(), 2),
                    'CSV Total': csv_row['Total'],
                })
            df_comp = pd.DataFrame(comp_rows) if comp_rows else pd.DataFrame()
            disc_rows = []
            for _, row in df_comp.iterrows():
                diffs = {}
                for tc, cc, label in [('TXT Freight','CSV Rate Charge','Rate Charge'),('TXT LEVY','CSV Fuel Levy','Fuel Levy'),('TXT GST','CSV Total Tax','Total Tax'),('TXT Final Total','CSV Total','Total')]:
                    diff = round(float(row[cc]) - float(row[tc]), 2)
                    if diff != 0:
                        diffs[f'TXT {label}'] = row[tc]; diffs[f'CSV {label}'] = row[cc]; diffs[f'Diff {label}'] = diff
                if diffs: disc_rows.append({'Customer Reference': row['Customer Reference'], **diffs})
            return unmatched_t, unmatched_c, df_comp, pd.DataFrame(disc_rows) if disc_rows else pd.DataFrame()

        def write_recon_sheet(ws, label, ut, uc, dc, dd):
            ws.cell(row=1, column=1, value=f'Reconciliation — {label}').font = Font(bold=True, size=14)
            def write_sec(ws, title, df_in, row):
                ws.cell(row=row, column=1, value=title).font = Font(bold=True, size=12)
                row += 1
                if df_in is None or df_in.empty:
                    ws.cell(row=row, column=1, value='No data').font = Font(italic=True)
                    return row + 2
                for c, col in enumerate(df_in.columns, 1):
                    cell = ws.cell(row=row, column=c, value=col)
                    cell.font = Font(bold=True, color='FFFFFF')
                    cell.fill = PatternFill('solid', start_color='1F4E79')
                row += 1
                for _, r in df_in.iterrows():
                    for c, val in enumerate(r, 1):
                        cell = ws.cell(row=row, column=c, value=val)
                        if df_in.columns[c-1].startswith('Diff '):
                            try:
                                v = float(val)
                                cell.fill = PatternFill('solid', start_color='1a7a1a' if v < 0 else 'a83232')
                                cell.font = Font(color='FFFFFF')
                            except: pass
                    row += 1
                return row + 2
            r = 3
            r = write_sec(ws, '① TXT not in .CSV Invoice', ut, r)
            r = write_sec(ws, '① CSV not in .TXT Statement', uc, r)
            r = write_sec(ws, '② Comparison Table', dc, r)
            r = write_sec(ws, '③ Discrepancies', dd, r)

        def build_recon_excel(df_txt_f, df_csv_f, df_mt_f, df_mc_f):
            wb  = Workbook()
            wb.remove(wb.active)
            states = sorted(df_txt_f['State'].dropna().unique().tolist())
            for label, state_filter in [('All', None)] + [(s, s) for s in states]:
                if state_filter is None:
                    t, mt, mc = df_txt_f, df_mt_f, df_mc_f
                else:
                    t   = df_txt_f[df_txt_f['State'] == state_filter].copy()
                    fd  = set(t['Delivery/Adjustment'].dropna().str.strip())
                    mt  = df_mt_f[df_mt_f['Delivery/Adjustment'].isin(fd)].copy()
                    mc  = df_mc_f[df_mc_f['Customer Reference'].apply(lambda x: any(r in fd for r in split_customer_refs(str(x))))].copy()
                ut, uc, dc, dd = build_recon_tables(t, df_csv_f, mt, mc)
                write_recon_sheet(wb.create_sheet(title=f'Recon - {label}'), label, ut, uc, dc, dd)
            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            return buf

        recon_buf = build_recon_excel(df_txt, df_csv, df_matched_txt, df_matched_csv)
        st.download_button(
            label="⬇️ Download Reconciliation Report (.xlsx)",
            data=recon_buf,
            file_name=f'invoice_reconciliation.xlsx',
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
