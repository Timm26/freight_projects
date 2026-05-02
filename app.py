import streamlit as st
import pandas as pd
import base64
import json
import requests
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

tab1, tab2, tab3 = st.tabs(["📊 RCTI Processor", "🔍 Invoice Reconciliation", "📸 Screenshot to Excel"])

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

        # ── State filter ──────────────────────────────────────────────────

        st.divider()
        available_states = sorted(df_txt['State'].dropna().unique().tolist())
        state_options    = ['All'] + available_states
        selected_state   = st.selectbox(
            "Filter by State",
            options=state_options,
            index=0,
            key="tab2_state_filter"
        )

        # Apply filter to TXT data — CSV has no state so we filter via matched deliveries
        if selected_state == 'All':
            df_txt_filtered     = df_txt.copy()
            df_matched_txt_f    = df_matched_txt.copy()
        else:
            df_txt_filtered     = df_txt[df_txt['State'] == selected_state].copy()
            filtered_deliveries = set(df_txt_filtered['Delivery/Adjustment'].dropna().str.strip())
            df_matched_txt_f    = df_matched_txt[
                df_matched_txt['Delivery/Adjustment'].isin(filtered_deliveries)
            ].copy()

            # Filter matched CSV to only rows where at least one ref is in filtered deliveries
            def csv_in_filter(ref_str):
                return any(r in filtered_deliveries for r in split_customer_refs(str(ref_str)))
            df_matched_csv      = df_matched_csv[
                df_matched_csv['Customer Reference'].apply(csv_in_filter)
            ].copy()

        # ── Section 1: Line item check ────────────────────────────────────

        st.divider()
        st.subheader("① Line Item Check")

        unmatched_txt = df_txt_filtered[df_txt_filtered['Reconciliation Status'] != '✓ Matched']
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
            matched_txt = df_matched_txt_f[df_matched_txt_f['Delivery/Adjustment'].isin(refs)]

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

        # Column pairs for value comparison (TXT col, CSV col)
        compare_pairs = [
            ('TXT Load Qty',    'CSV Total Cubic'),
            ('TXT Paid Qty',    'CSV Quantity'),
            ('TXT Freight',     'CSV Rate Charge'),
            ('TXT LEVY',        'CSV Fuel Levy'),
            ('TXT Total (AUD)', 'CSV Rate+Fuel Levy'),
            ('TXT GST',         'CSV Total Tax'),
            ('TXT Final Total', 'CSV Total'),
        ]
        csv_to_txt = {csv_col: txt_col for txt_col, csv_col in compare_pairs}

        # TXT = dark blue, CSV = dark teal, green/red on CSV where values differ
        TXT_BG  = 'background-color: #1a2a45; color: #a8c4e0'
        CSV_BG  = 'background-color: #1a3a35; color: #a8d4cc'
        CSV_RED = 'background-color: #a83232; color: white'
        CSV_GRN = 'background-color: #1a7a1a; color: white'

        def style_comparison(df):
            styles = pd.DataFrame('', index=df.index, columns=df.columns)
            for col in df.columns:
                if col == 'Customer Reference':
                    continue
                elif col.startswith('TXT'):
                    styles[col] = TXT_BG
                elif col.startswith('CSV'):
                    txt_col = csv_to_txt.get(col)
                    if txt_col and txt_col in df.columns:
                        for i in range(len(df)):
                            try:
                                csv_val = float(df[col].iloc[i])
                                txt_val = float(df[txt_col].iloc[i])
                                if txt_val > csv_val:
                                    styles.iloc[i, df.columns.get_loc(col)] = CSV_GRN
                                elif csv_val > txt_val:
                                    styles.iloc[i, df.columns.get_loc(col)] = CSV_RED
                                else:
                                    styles.iloc[i, df.columns.get_loc(col)] = CSV_BG
                            except:
                                styles.iloc[i, df.columns.get_loc(col)] = CSV_BG
                    else:
                        styles[col] = CSV_BG
            return styles

        # Match stats
        total_txt     = len(df_matched_txt_f)
        total_csv     = len(df_matched_csv)
        exact_matches = sum(
            1 for _, row in df_compare.iterrows()
            if all(
                float(row[csv_col]) == float(row[txt_col])
                for txt_col, csv_col in [
                    ('TXT Freight', 'CSV Rate Charge'),
                    ('TXT LEVY', 'CSV Fuel Levy'),
                    ('TXT Final Total', 'CSV Total')
                ]
            )
        )

        col1, col2, col3 = st.columns(3)
        col1.metric("TXT Lines Matched", total_txt)
        col2.metric("CSV Lines Matched", total_csv)
        col3.metric("Exact Value Matches", f"{exact_matches} / {total_csv}")

        st.caption("🔵 TXT columns (dark blue)  |  🟩 CSV columns (teal = match, green = higher, red = lower)")
        st.dataframe(
            df_compare.style.apply(style_comparison, axis=None),
            use_container_width=True,
            hide_index=True
        )

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
                    if col == 'Customer Reference':
                        continue
                    elif col.startswith('TXT '):
                        styles[col] = 'background-color: #1a2a45; color: #a8c4e0'
                    elif col.startswith('CSV '):
                        styles[col] = 'background-color: #1a3a35; color: #a8d4cc'
                    elif col.startswith('Diff '):
                        for i, val in enumerate(df[col]):
                            if pd.notna(val) and val != '':
                                try:
                                    if float(val) < 0:
                                        styles.iloc[i, df.columns.get_loc(col)] = \
                                            'background-color: #1a7a1a; color: white'
                                    elif float(val) > 0:
                                        styles.iloc[i, df.columns.get_loc(col)] = \
                                            'background-color: #a83232; color: white'
                                except:
                                    pass
                return styles

            st.caption("🔵 TXT columns  |  🟩 CSV columns  |  🟢 Diff positive (CSV higher)  |  🔴 Diff negative (CSV lower)")
            st.dataframe(
                df_disc.style.apply(highlight_diffs, axis=None),
                use_container_width=True,
                hide_index=True
            )

        # ── Download reconciliation Excel ─────────────────────────────────

        st.divider()
        st.subheader("⬇️ Download Reconciliation Report")

        def build_recon_tables(df_txt_in, df_csv_in, df_matched_txt_in, df_matched_csv_in):
            """Build comparison and discrepancy dataframes for a given filtered dataset."""
            unmatched_t = df_txt_in[df_txt_in['Reconciliation Status'] != '✓ Matched'][[
                'Delivery/Adjustment', 'State', 'Origin', 'Total (AUD)', 'GST (AUD)', 'Reconciliation Status'
            ]].copy()

            unmatched_c = df_csv_in[df_csv_in['Reconciliation Status'] != '✓ Matched'][[
                'Customer Reference', 'Total', 'Reconciliation Status'
            ]].copy()

            comp_rows = []
            for _, csv_row in df_matched_csv_in.iterrows():
                refs        = split_customer_refs(csv_row['Customer Reference'])
                mtxt        = df_matched_txt_in[df_matched_txt_in['Delivery/Adjustment'].isin(refs)]
                comp_rows.append({
                    'Customer Reference':   csv_row['Customer Reference'],
                    'TXT Load Qty':         round(mtxt['Load Qty'].sum(), 3),
                    'CSV Total Cubic':      csv_row['Total Cubic'],
                    'TXT Paid Qty':         round(mtxt['Paid Qty'].sum(), 3),
                    'CSV Quantity':         csv_row['Quantity'],
                    'TXT Freight':          round(mtxt['Freight'].sum(), 2),
                    'CSV Rate Charge':      csv_row['Rate Charge'],
                    'TXT LEVY':             round(mtxt['LEVY'].sum(), 2),
                    'CSV Fuel Levy':        csv_row['Fuel Levy'],
                    'TXT Total (AUD)':      round(mtxt['Total (AUD)'].sum(), 2),
                    'CSV Rate+Fuel Levy':   csv_row['Rate Charge and Fuel Levy'],
                    'TXT GST':              round(mtxt['GST (AUD)'].sum(), 2),
                    'CSV Total Tax':        csv_row['Total Tax'],
                    'TXT Final Total':      round(mtxt['Final Total'].sum(), 2),
                    'CSV Total':            csv_row['Total'],
                })
            df_comp = pd.DataFrame(comp_rows) if comp_rows else pd.DataFrame()

            disc_pairs = [
                ('TXT Freight',     'CSV Rate Charge',  'Rate Charge'),
                ('TXT LEVY',        'CSV Fuel Levy',     'Fuel Levy'),
                ('TXT GST',         'CSV Total Tax',     'Total Tax'),
                ('TXT Final Total', 'CSV Total',         'Total'),
            ]
            disc_rows = []
            for _, row in df_comp.iterrows():
                diffs = {}
                for tc, cc, label in disc_pairs:
                    diff = round(float(row[cc]) - float(row[tc]), 2)
                    if diff != 0:
                        diffs[f'TXT {label}'] = row[tc]
                        diffs[f'CSV {label}'] = row[cc]
                        diffs[f'Diff {label}'] = diff
                if diffs:
                    disc_rows.append({'Customer Reference': row['Customer Reference'], **diffs})
            df_disc_out = pd.DataFrame(disc_rows) if disc_rows else pd.DataFrame()

            return unmatched_t, unmatched_c, df_comp, df_disc_out

        def write_recon_sheet(ws, sheet_label, unmatched_t, unmatched_c, df_comp, df_disc_out):
            """Write all three tables to a single worksheet with headers."""
            from openpyxl.styles import Font, PatternFill, Alignment

            HDR_BLUE = 'FF1a2a45'
            HDR_TEAL = 'FF1a3a35'
            HDR_GREY = 'FF2a2a3a'
            RED_BG   = 'FFa83232'
            GRN_BG   = 'FF1a7a1a'

            def write_section(ws, title, df, start_row, col_colors=None):
                # Section title
                ws.cell(row=start_row, column=1, value=title).font = Font(bold=True, size=12)
                start_row += 1

                if df is None or df.empty:
                    ws.cell(row=start_row, column=1, value='No data').font = Font(italic=True)
                    return start_row + 2

                # Headers
                for c, col in enumerate(df.columns, 1):
                    cell = ws.cell(row=start_row, column=c, value=col)
                    cell.font = Font(bold=True, color='FFFFFF')
                    bg = HDR_GREY
                    if col_colors:
                        if col.startswith('TXT'):  bg = HDR_BLUE
                        elif col.startswith('CSV'): bg = HDR_TEAL
                    cell.fill = PatternFill('solid', start_color=bg)
                start_row += 1

                # Data
                for _, row in df.iterrows():
                    for c, col in enumerate(df.columns, 1):
                        cell = ws.cell(row=start_row, column=c, value=row[col])
                        if col_colors and col.startswith('Diff '):
                            try:
                                v = float(row[col])
                                if v < 0:
                                    cell.fill = PatternFill('solid', start_color=GRN_BG)
                                    cell.font = Font(color='FFFFFF')
                                elif v > 0:
                                    cell.fill = PatternFill('solid', start_color=RED_BG)
                                    cell.font = Font(color='FFFFFF')
                            except:
                                pass
                    start_row += 1

                return start_row + 2  # gap between sections

            row = 1
            ws.cell(row=row, column=1, value=f'Reconciliation Report — {sheet_label}').font = Font(bold=True, size=14)
            row += 2

            row = write_section(ws, '① Line Item Check — TXT not in Rohlig Invoice', unmatched_t, row)
            row = write_section(ws, '① Line Item Check — CSV not in Client Statement', unmatched_c, row)
            row = write_section(ws, '② Comparison Table', df_comp, row, col_colors=True)
            row = write_section(ws, '③ Discrepancies', df_disc_out, row, col_colors=True)

        def build_recon_excel(df_txt_full, df_csv_full, df_matched_txt_full, df_matched_csv_full):
            wb  = Workbook()
            wb.remove(wb.active)
            states = sorted(df_txt_full['State'].dropna().unique().tolist())
            sheets = [('All', None)] + [(s, s) for s in states]

            for sheet_label, state_filter in sheets:
                if state_filter is None:
                    t, m_t = df_txt_full, df_matched_txt_full
                    m_c    = df_matched_csv_full
                else:
                    t      = df_txt_full[df_txt_full['State'] == state_filter].copy()
                    f_dels = set(t['Delivery/Adjustment'].dropna().str.strip())
                    m_t    = df_matched_txt_full[df_matched_txt_full['Delivery/Adjustment'].isin(f_dels)].copy()
                    def in_filter(ref_str):
                        return any(r in f_dels for r in split_customer_refs(str(ref_str)))
                    m_c    = df_matched_csv_full[df_matched_csv_full['Customer Reference'].apply(in_filter)].copy()

                ut, uc, dc, dd = build_recon_tables(t, df_csv_full, m_t, m_c)
                ws = wb.create_sheet(title=f'Recon - {sheet_label}')
                write_recon_sheet(ws, sheet_label, ut, uc, dc, dd)

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            return buf

        recon_buffer = build_recon_excel(df_txt, df_csv, df_matched_txt, df_matched_csv)
        vendor       = df_txt['Vendor'].iloc[0]

        st.download_button(
            label="⬇️ Download Reconciliation Report (.xlsx)",
            data=recon_buffer,
            file_name=f'Rohlig_{vendor}_reconciliation.xlsx',
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ════════════════════════════════════════════════════════════════════════════
# TAB 3 — Screenshot to Excel
# ════════════════════════════════════════════════════════════════════════════

with tab3:
    st.write("Upload one or more screenshots of your charge table. The AI will extract the data and build a downloadable Excel file.")

    screenshots = st.file_uploader(
        "Upload screenshots (.png, .jpg)",
        type=['png', 'jpg', 'jpeg'],
        accept_multiple_files=True,
        key="tab3_screenshots"
    )

    COLUMNS = [
        'Charge', 'Description', 'Bran', 'Depa', 'Cost',
        'OS Cost Amt', 'Estimated Cost', 'Local Cost Amt', 'Creditor',
        'Cost Recognition', 'Posted', 'Apt', 'Sell', 'OS Sell Amt',
        'Estimated Revenue', 'Local Sell Amt', 'Debtor',
        'Sell Recognition', 'Posted (Sell)'
    ]

    def extract_table_from_image(image_bytes, media_type):
        """Send image to Anthropic API and extract table rows as JSON."""
        b64 = base64.standard_b64encode(image_bytes).decode('utf-8')

        prompt = f"""This is a screenshot of a financial charge table.
Extract every data row as a JSON array of objects.
Use exactly these column names: {json.dumps(COLUMNS)}
- Include every visible row, skip the header row.
- Use null for any cell that is blank or not visible.
- Return ONLY the JSON array, no markdown, no explanation.
Example format: [{{"Charge":"WDEL","Description":"Delivery Charge",...}}]"""

        response = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"Content-Type": "application/json"},
            json={
                "model": "claude-sonnet-4-20250514",
                "max_tokens": 4096,
                "messages": [{
                    "role": "user",
                    "content": [
                        {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64}},
                        {"type": "text", "text": prompt}
                    ]
                }]
            }
        )

        result = response.json()
        raw = result['content'][0]['text'].strip()
        # Strip markdown fences if present
        raw = re.sub(r"^```[a-z]*\n?", "", raw)
        raw = re.sub(r"\n?```$", "", raw)
        return json.loads(raw.strip())

    if screenshots:
        if st.button("🔍 Extract Table Data", key="tab3_extract"):
            all_rows = []
            progress = st.progress(0)
            status   = st.empty()

            for i, img_file in enumerate(screenshots):
                status.text(f"Processing {img_file.name} ({i+1}/{len(screenshots)})...")
                img_bytes  = img_file.read()
                media_type = "image/png" if img_file.name.lower().endswith(".png") else "image/jpeg"

                try:
                    rows = extract_table_from_image(img_bytes, media_type)
                    for row in rows:
                        row['_source'] = img_file.name
                    all_rows.extend(rows)
                except Exception as e:
                    st.warning(f"⚠ Could not extract from {img_file.name}: {e}")

                progress.progress((i + 1) / len(screenshots))

            status.empty()
            progress.empty()

            if all_rows:
                df_shots = pd.DataFrame(all_rows)

                # Ensure all expected columns exist
                for col in COLUMNS:
                    if col not in df_shots.columns:
                        df_shots[col] = None

                # Clean numeric columns
                for col in ['OS Cost Amt', 'Estimated Cost', 'Local Cost Amt',
                            'OS Sell Amt', 'Estimated Revenue', 'Local Sell Amt']:
                    df_shots[col] = pd.to_numeric(df_shots[col], errors='coerce')

                st.session_state['screenshot_df'] = df_shots
                st.success(f"✓ Extracted {len(df_shots)} rows from {len(screenshots)} screenshot(s)")

    # Show results if we have data
    if 'screenshot_df' in st.session_state:
        df_shots = st.session_state['screenshot_df']

        st.divider()
        st.subheader("Extracted Data")
        display_cols = [c for c in COLUMNS if c in df_shots.columns]
        st.dataframe(df_shots[display_cols], use_container_width=True, hide_index=True)

        # ── Summary 1: OS Cost Amt by Creditor ───────────────────────────
        st.divider()
        st.subheader("OS Cost Amt by Creditor")
        if 'Creditor' in df_shots.columns and 'OS Cost Amt' in df_shots.columns:
            cost_summary = (
                df_shots.groupby('Creditor')['OS Cost Amt']
                .sum().reset_index()
                .rename(columns={'OS Cost Amt': 'Total OS Cost Amt'})
                .sort_values('Total OS Cost Amt', ascending=False)
            )
            cost_summary['Total OS Cost Amt'] = cost_summary['Total OS Cost Amt'].round(2)
            st.dataframe(cost_summary, use_container_width=True, hide_index=True)

        # ── Summary 2: Estimated Revenue by Debtor ───────────────────────
        st.divider()
        st.subheader("Estimated Revenue by Debtor")
        if 'Debtor' in df_shots.columns and 'Estimated Revenue' in df_shots.columns:
            rev_summary = (
                df_shots.groupby('Debtor')['Estimated Revenue']
                .sum().reset_index()
                .rename(columns={'Estimated Revenue': 'Total Estimated Revenue'})
                .sort_values('Total Estimated Revenue', ascending=False)
            )
            rev_summary['Total Estimated Revenue'] = rev_summary['Total Estimated Revenue'].round(2)
            st.dataframe(rev_summary, use_container_width=True, hide_index=True)

        # ── Download ──────────────────────────────────────────────────────
        st.divider()

        def build_screenshot_excel(df, cost_sum, rev_sum):
            wb  = Workbook()
            wb.remove(wb.active)

            def write_df_to_sheet(ws, df_in, title):
                ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=13)
                for c, col in enumerate(df_in.columns, 1):
                    cell = ws.cell(row=2, column=c, value=col)
                    cell.font = Font(bold=True, color='FFFFFF')
                    cell.fill = PatternFill('solid', start_color='1F4E79')
                for r, (_, row) in enumerate(df_in.iterrows(), 3):
                    for c, val in enumerate(row, 1):
                        ws.cell(row=r, column=c, value=val)

            # Sheet 1 — full data
            ws1 = wb.create_sheet("Extracted Data")
            write_df_to_sheet(ws1, df[display_cols], "Extracted Charge Data")

            # Sheet 2 — OS Cost by Creditor
            ws2 = wb.create_sheet("Cost by Creditor")
            write_df_to_sheet(ws2, cost_sum, "OS Cost Amt by Creditor")

            # Sheet 3 — Revenue by Debtor
            ws3 = wb.create_sheet("Revenue by Debtor")
            write_df_to_sheet(ws3, rev_sum, "Estimated Revenue by Debtor")

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
