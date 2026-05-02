import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
import io
import re

# ── Page config ───────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Rohlig RCTI Processor",
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

st.title("🚛 Rohlig RCTI Processor")

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
    """Split combined refs like '1800737446, 1800737447 & 1800737444' into a list"""
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

    headers             = [c for c in df.columns if c not in ('State', 'Final Total')]
    headers_with_state  = headers + ['State']

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

    ws_sum.cell(row=grand_row+2, column=1, value='Client Total (inc GST)').font = Font(bold=True)
    ws_sum.cell(row=grand_row+2, column=2, value=client_total)
    ws_sum.cell(row=grand_row+3, column=1, value='Check').font = Font(bold=True)
    ws_sum.cell(row=grand_row+3, column=2,
                value=f'=IF(F{grand_row}=B{grand_row+2},"✓ MATCH","⚠ DISCREPANCY")')

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ── Tabs ──────────────────────────────────────────────────────────────────────

tab1, tab2 = st.tabs(["📊 RCTI Processor", "🔍 Invoice Reconciliation"])

# ════════════════════════════════════════════════════════════════════════════
# TAB 1 — RCTI Processor
# ════════════════════════════════════════════════════════════════════════════

with tab1:
    st.write("Upload your `.TXT` statement file and download the split `.xlsx` automatically.")
    uploaded_txt = st.file_uploader("Choose your .TXT file", type=['txt', 'TXT'], key="tab1_txt")

    if uploaded_txt:
        if uploaded_txt.size > 20 * 1024 * 1024:
            st.error("⚠ File exceeds the 20MB limit. Please upload a smaller file.")
            st.stop()
        with st.spinner("Processing..."):
            df, client_total = parse_txt(uploaded_txt)
            our_total        = df['Total (AUD)'].sum()
            our_gst          = df['GST (AUD)'].sum()
            our_inc_gst      = our_total + our_gst
            match            = round(our_inc_gst, 2) == round(client_total, 2)
            buffer           = build_excel(df, client_total)
            vendor           = df['Vendor'].iloc[0]
            filename         = f'Rohlig_{vendor}_split.xlsx'

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
            st.success(f"✓ MATCH — Our total matches client total of ${client_total:,.2f}")
        else:
            diff = abs(our_inc_gst - client_total)
            st.error(f"⚠ DISCREPANCY — Difference of ${diff:,.2f} vs client total of ${client_total:,.2f}")

        st.divider()
        st.subheader("Breakdown by Truck & State")
        summary = []
        for truck in sorted(df['Truck'].unique()):
            for state in sorted(df[df['Truck'] == truck]['State'].unique()):
                s = df[(df['Truck'] == truck) & (df['State'] == state)]
                summary.append({
                    'Truck':           truck,
                    'State':           state,
                    'Rows':            len(s),
                    'Total (AUD)':     round(s['Total (AUD)'].sum(), 2),
                    'GST (AUD)':       round(s['GST (AUD)'].sum(), 2),
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
# TAB 2 — Invoice Reconciliation
# ════════════════════════════════════════════════════════════════════════════

with tab2:
    st.write("Upload both files to reconcile client charges (.TXT) against Rohlig's invoice (.CSV).")

    col1, col2 = st.columns(2)
    with col1:
        recon_txt = st.file_uploader("Client charges (.TXT)", type=['txt', 'TXT'], key="tab2_txt")
    with col2:
        recon_csvs = st.file_uploader(
            "Rohlig invoice(s) (.CSV) — select multiple if needed",
            type=['csv', 'CSV'],
            key="tab2_csv",
            accept_multiple_files=True
        )

    if recon_txt and recon_csvs:

        # Check file sizes before processing
        MAX_MB = 20
        oversized = [f.name for f in recon_csvs if f.size > MAX_MB * 1024 * 1024]
        if oversized:
            st.error(f"⚠ The following file(s) exceed the {MAX_MB}MB limit: {', '.join(oversized)}")
            st.stop()

        with st.spinner("Reconciling..."):
            df_txt, _ = parse_txt(recon_txt)

            # Join all uploaded CSVs into one dataframe
            csv_frames = [parse_csv(f) for f in recon_csvs]
            df_csv = pd.concat(csv_frames, ignore_index=True)

            # Show which invoices were loaded
            invoice_nums = df_csv['Invoice Number'].unique().tolist() if 'Invoice Number' in df_csv.columns else []
            st.info(f"📄 {len(recon_csvs)} CSV file(s) loaded — Invoice(s): {', '.join(str(i) for i in invoice_nums)}")

            # Build delivery sets
            txt_deliveries = set(df_txt['Delivery/Adjustment'].dropna().str.strip().unique())
            csv_deliveries = set()
            for ref in df_csv['Customer Reference'].dropna():
                for r in split_customer_refs(ref):
                    csv_deliveries.add(r)

            # Label TXT rows
            df_txt['Reconciliation Status'] = df_txt['Delivery/Adjustment'].apply(
                lambda x: '✓ Matched' if x in csv_deliveries
                          else 'Invoice not found on warehouse invoice'
            )

            # Label CSV rows
            df_csv['Reconciliation Status'] = df_csv['Customer Reference'].apply(
                lambda x: '✓ Matched' if any(r in txt_deliveries for r in split_customer_refs(x))
                          else 'Invoice not found on warehouse invoice'
            )

            # Matched subsets only
            df_matched_txt = df_txt[df_txt['Reconciliation Status'] == '✓ Matched'].copy()
            df_matched_csv = df_csv[df_csv['Reconciliation Status'] == '✓ Matched'].copy()

        # ── Section 1: Line item check ────────────────────────────────────

        st.divider()
        st.subheader("① Line Item Check")

        unmatched_txt = df_txt[df_txt['Reconciliation Status'] != '✓ Matched']
        unmatched_csv = df_csv[df_csv['Reconciliation Status'] != '✓ Matched']

        if unmatched_txt.empty and unmatched_csv.empty:
            st.success("✓ All invoice lines found — every line matches across both files")
        else:
            if not unmatched_txt.empty:
                st.warning(f"⚠ {len(unmatched_txt)} TXT line(s) not found in Rohlig invoice")
                st.dataframe(
                    unmatched_txt[['Delivery/Adjustment', 'State', 'Origin',
                                   'Total (AUD)', 'GST (AUD)', 'Reconciliation Status']],
                    use_container_width=True, hide_index=True
                )
            if not unmatched_csv.empty:
                st.warning(f"⚠ {len(unmatched_csv)} CSV line(s) not found in client statement")
                st.dataframe(
                    unmatched_csv[['Customer Reference', 'Total', 'Reconciliation Status']],
                    use_container_width=True, hide_index=True
                )

        # ── Section 2: Comparison table ───────────────────────────────────

        st.divider()
        st.subheader("② Comparison Table")
        st.caption("Matched lines only — TXT vs CSV side by side. Multiple TXT rows for the same reference are summed.")

        comparison_rows = []
        for _, csv_row in df_matched_csv.iterrows():
            refs        = split_customer_refs(csv_row['Customer Reference'])
            matched_txt = df_matched_txt[df_matched_txt['Delivery/Adjustment'].isin(refs)]

            comparison_rows.append({
                'Customer Reference':   csv_row['Customer Reference'],
                'TXT Load Qty':         round(matched_txt['Load Qty'].sum(), 3),
                'CSV Total Cubic':      csv_row['Total Cubic'],
                'TXT Paid Qty':         round(matched_txt['Paid Qty'].sum(), 3),
                'CSV Quantity':         csv_row['Quantity'],
                'TXT Freight':          round(matched_txt['Freight'].sum(), 2),
                'CSV Rate Charge':      csv_row['Rate Charge'],
                'TXT LEVY':             round(matched_txt['LEVY'].sum(), 2),
                'CSV Fuel Levy':        csv_row['Fuel Levy'],
                'TXT Total (AUD)':      round(matched_txt['Total (AUD)'].sum(), 2),
                'CSV Rate+Fuel Levy':   csv_row['Rate Charge and Fuel Levy'],
                'TXT GST':              round(matched_txt['GST (AUD)'].sum(), 2),
                'CSV Total Tax':        csv_row['Total Tax'],
                'TXT Final Total':      round(matched_txt['Final Total'].sum(), 2),
                'CSV Total':            csv_row['Total'],
            })

        df_compare = pd.DataFrame(comparison_rows)
        st.dataframe(df_compare, use_container_width=True, hide_index=True)

        # ── Section 3: Discrepancy table ──────────────────────────────────

        st.divider()
        st.subheader("③ Discrepancies")
        st.caption("🟢 CSV higher than TXT  |  🔴 CSV lower than TXT")

        pairs = [
            ('TXT Freight',     'CSV Rate Charge',  'Rate Charge'),
            ('TXT LEVY',        'CSV Fuel Levy',     'Fuel Levy'),
            ('TXT GST',         'CSV Total Tax',     'Total Tax'),
            ('TXT Final Total', 'CSV Total',         'Total'),
        ]

        discrepancy_rows = []
        for _, row in df_compare.iterrows():
            diffs = {}
            for txt_col, csv_col, label in pairs:
                diff = round(float(row[csv_col]) - float(row[txt_col]), 2)
                if diff != 0:
                    diffs[f'TXT {label}'] = row[txt_col]
                    diffs[f'CSV {label}'] = row[csv_col]
                    diffs[f'Diff {label}'] = diff
            if diffs:
                entry = {'Customer Reference': row['Customer Reference']}
                entry.update(diffs)
                discrepancy_rows.append(entry)

        if not discrepancy_rows:
            st.success("✓ No discrepancies found — all matched values align")
        else:
            st.error(f"⚠ {len(discrepancy_rows)} row(s) with discrepancies")
            df_disc = pd.DataFrame(discrepancy_rows)

            def highlight_diffs(df):
                styles = pd.DataFrame('', index=df.index, columns=df.columns)
                for col in df.columns:
                    if col.startswith('Diff '):
                        for i, val in enumerate(df[col]):
                            if pd.notna(val) and val != '':
                                try:
                                    if float(val) > 0:
                                        # CSV higher than TXT — green
                                        styles.iloc[i, df.columns.get_loc(col)] = \
                                            'background-color: #1a7a1a; color: white'
                                    elif float(val) < 0:
                                        # CSV lower than TXT — red
                                        styles.iloc[i, df.columns.get_loc(col)] = \
                                            'background-color: #a83232; color: white'
                                except:
                                    pass
                return styles

            st.dataframe(
                df_disc.style.apply(highlight_diffs, axis=None),
                use_container_width=True,
                hide_index=True
            )
