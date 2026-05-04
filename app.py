import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
import io
import re

# ── Page config ───────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="RCTI Processor",
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

st.title("🚛 RCTI Processor")

# ── Constants ─────────────────────────────────────────────────────────────────

STATES   = ['VIC', 'NSW', 'QLD', 'SA', 'WA', 'TAS', 'NT', 'ACT']
ZONE_MAP = {
    'YV': 'VIC', 'YN': 'NSW', 'YQ': 'QLD', 'YS': 'SA',
    'YW': 'WA',  'YT': 'TAS', 'YD': 'NT',  'YA': 'ACT'
}
HEADER_COLOR = '1F4E79'

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

def parse_csv(uploaded_file):
    df = pd.read_csv(uploaded_file, dtype=str)
    for col in ['Quantity', 'Total Cubic', 'Rate Charge', 'Fuel Levy',
                'Rate Charge and Fuel Levy', 'Total Tax', 'Total']:
        df[col] = df[col].apply(clean_number)
    return df

def build_excel(df, client_total):
    wb = Workbook()
    wb.remove(wb.active)
    headers            = [c for c in df.columns if c not in ('State', 'Final Total')]
    headers_with_state = headers + ['State']

    for truck in sorted(df['Truck'].unique()):
        truck_df = df[df['Truck'] == truck]
        for state in sorted(truck_df['State'].unique()):
            state_df = truck_df[truck_df['State'] == state].copy()
            ws = wb.create_sheet(title=f"{truck} - {state}")
            for col, header in enumerate(headers_with_state, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = Font(bold=True, color='FFFFFF')
                cell.fill = PatternFill('solid', start_color=HEADER_COLOR)
            for row_idx, (_, row) in enumerate(state_df[headers_with_state].iterrows(), 2):
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

    ws_sum      = wb.create_sheet(title='Summary', index=0)
    sum_headers = ['Truck', 'State', 'Rows', 'Total (AUD)', 'GST (AUD)', 'Total (inc GST)']
    for col, header in enumerate(sum_headers, 1):
        cell = ws_sum.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', start_color=HEADER_COLOR)
    data_row = 2
    for truck in sorted(df['Truck'].unique()):
        for state in sorted(df[df['Truck'] == truck]['State'].unique()):
            state_df = df[(df['Truck'] == truck) & (df['State'] == state)]
            total    = state_df['Total (AUD)'].sum()
            gst      = state_df['GST (AUD)'].sum()
            ws_sum.cell(row=data_row, column=1, value=truck)
            ws_sum.cell(row=data_row, column=2, value=state)
            ws_sum.cell(row=data_row, column=3, value=len(state_df))
            ws_sum.cell(row=data_row, column=4, value=round(total, 2))
            ws_sum.cell(row=data_row, column=5, value=round(gst, 2))
            ws_sum.cell(row=data_row, column=6, value=round(total + gst, 2))
            data_row += 1
    grand_row = data_row
    ws_sum.cell(row=grand_row, column=1, value='TOTAL').font = Font(bold=True)
    for col in [3, 4, 5, 6]:
        ws_sum.cell(row=grand_row, column=col,
                    value=f'=SUM({get_column_letter(col)}2:{get_column_letter(col)}{grand_row-1})'
                    ).font = Font(bold=True)
    ws_sum.cell(row=grand_row+2, column=1, value='Statements Total (inc GST)').font = Font(bold=True)
    ws_sum.cell(row=grand_row+2, column=2, value=client_total)
    ws_sum.cell(row=grand_row+3, column=1, value='Check').font = Font(bold=True)
    ws_sum.cell(row=grand_row+3, column=2,
                value=f'=IF(F{grand_row}=B{grand_row+2},"✓ MATCH","⚠ DISCREPANCY")')
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ── Tabs ──────────────────────────────────────────────────────────────────────

tab1, tab2, tab3 = st.tabs(["📊 RCTI Processor", "🔍 Invoice Reconciliation", "📸 Screenshot to Excel"])

# ════════════════════════════════════════════════════════════════════════════
# TAB 1
# ════════════════════════════════════════════════════════════════════════════

with tab1:
    st.write("Upload your `.TXT` statement file and download the split `.xlsx` automatically.")
    uploaded_txt = st.file_uploader("Choose your .TXT file", type=['txt', 'TXT'], key="tab1_txt")

    if uploaded_txt:
        if uploaded_txt.size > 20 * 1024 * 1024:
            st.error("⚠ File exceeds the 20MB limit.")
            st.stop()
        with st.spinner("Processing..."):
            df, client_total = parse_txt(uploaded_txt)
            our_total        = df['Total (AUD)'].sum()
            our_gst          = df['GST (AUD)'].sum()
            our_inc_gst      = our_total + our_gst
            match            = round(our_inc_gst, 2) == round(client_total, 2)
            buffer           = build_excel(df, client_total)
            vendor           = df['Vendor'].iloc[0]
            filename         = f'{vendor}_split.xlsx'

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

        st.divider()
        st.subheader("Breakdown by Truck & State")
        summary = []
        for truck in sorted(df['Truck'].unique()):
            for state in sorted(df[df['Truck'] == truck]['State'].unique()):
                s = df[(df['Truck'] == truck) & (df['State'] == state)]
                summary.append({
                    'Truck': truck, 'State': state, 'Rows': len(s),
                    'Total (AUD)': round(s['Total (AUD)'].sum(), 2),
                    'GST (AUD)': round(s['GST (AUD)'].sum(), 2),
                    'Total (inc GST)': round(s['Total (AUD)'].sum() + s['GST (AUD)'].sum(), 2)
                })
        st.dataframe(pd.DataFrame(summary), use_container_width=True, hide_index=True)

        st.divider()
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

    # X positions of each column header at 2x image scale
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
