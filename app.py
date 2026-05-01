import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
import io

# ── Page config ───────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Rohlig RCTI Processor",
    page_icon="🚛",
    layout="centered"
)

st.title("🚛 Rohlig RCTI Processor")
st.write("Upload your `.TXT` statement file and download the split `.xlsx` automatically.")

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

def process_file(uploaded_file):
    # Step 1: Read
    df_raw = pd.read_csv(uploaded_file, sep='\t', dtype=str)
    summary_row_data = df_raw.iloc[-1]
    df = df_raw[:-1].copy()

    # Step 2: Clean numbers
    df['Total (AUD)'] = df['Total (AUD)'].apply(clean_number)
    df['GST (AUD)']   = df['GST (AUD)'].apply(clean_number)
    client_total      = clean_number(summary_row_data['Total (AUD)'])

    # Step 3: State column
    df['State'] = df.apply(get_state, axis=1)

    # Step 4: Cross-check
    our_total   = df['Total (AUD)'].sum()
    our_gst     = df['GST (AUD)'].sum()
    our_inc_gst = our_total + our_gst
    match       = round(our_inc_gst, 2) == round(client_total, 2)

    # Step 5 & 6: Build Excel
    wb = Workbook()
    wb.remove(wb.active)

    headers = [c for c in df.columns if c != 'State']
    headers_with_state = headers + ['State']

    for truck in sorted(df['Truck'].unique()):
        truck_df = df[df['Truck'] == truck]
        for state in sorted(truck_df['State'].unique()):
            state_df = truck_df[truck_df['State'] == state].copy()
            ws = wb.create_sheet(title=f"{truck} - {state}")

            # Header
            for col, header in enumerate(headers_with_state, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = Font(bold=True, color='FFFFFF')
                cell.fill = PatternFill('solid', start_color=HEADER_COLOR)

            # Data
            for row_idx, (_, row) in enumerate(state_df[headers_with_state].iterrows(), 2):
                for col_idx, val in enumerate(row, 1):
                    ws.cell(row=row_idx, column=col_idx, value=val)

            # Total row
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

    # Summary sheet
    ws_sum = wb.create_sheet(title='Summary', index=0)
    sum_headers = ['Truck', 'State', 'Rows', 'Total (AUD)', 'GST (AUD)', 'Total (inc GST)']
    for col, header in enumerate(sum_headers, 1):
        cell = ws_sum.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', start_color=HEADER_COLOR)

    data_row = 2
    for truck in sorted(df['Truck'].unique()):
        for state in sorted(df[df['Truck'] == truck]['State'].unique()):
            state_df = df[(df['Truck'] == truck) & (df['State'] == state)]
            total = state_df['Total (AUD)'].sum()
            gst   = state_df['GST (AUD)'].sum()
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

    # Save to memory buffer (not a file on disk)
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    vendor = df['Vendor'].iloc[0]
    filename = f'Rohlig_{vendor}_split.xlsx'

    return df, client_total, our_total, our_gst, our_inc_gst, match, buffer, filename

# ── UI ────────────────────────────────────────────────────────────────────────

uploaded_file = st.file_uploader("Choose your .TXT file", type=['txt', 'TXT'])

if uploaded_file:
    with st.spinner("Processing..."):
        df, client_total, our_total, our_gst, our_inc_gst, match, buffer, filename = process_file(uploaded_file)

    # Stats
    st.divider()
    col1, col2, col3 = st.columns(3)
    col1.metric("Rows", len(df))
    col2.metric("Trucks", df['Truck'].nunique())
    col3.metric("States", df['State'].nunique())

    st.divider()

    # Totals breakdown
    st.subheader("Totals")
    col1, col2, col3 = st.columns(3)
    col1.metric("Total (ex GST)", f"${our_total:,.2f}")
    col2.metric("GST", f"${our_gst:,.2f}")
    col3.metric("Total (inc GST)", f"${our_inc_gst:,.2f}")

    # Cross-check
    if match:
        st.success(f"✓ MATCH — Our total matches client total of ${client_total:,.2f}")
    else:
        diff = abs(our_inc_gst - client_total)
        st.error(f"⚠ DISCREPANCY — Difference of ${diff:,.2f} vs client total of ${client_total:,.2f}")

    # Breakdown table
    st.divider()
    st.subheader("Breakdown by Truck & State")
    summary = []
    for truck in sorted(df['Truck'].unique()):
        for state in sorted(df[df['Truck'] == truck]['State'].unique()):
            s = df[(df['Truck'] == truck) & (df['State'] == state)]
            summary.append({
                'Truck': truck,
                'State': state,
                'Rows': len(s),
                'Total (AUD)': round(s['Total (AUD)'].sum(), 2),
                'GST (AUD)': round(s['GST (AUD)'].sum(), 2),
                'Total (inc GST)': round(s['Total (AUD)'].sum() + s['GST (AUD)'].sum(), 2)
            })
    st.dataframe(pd.DataFrame(summary), use_container_width=True, hide_index=True)

    # Download
    st.divider()
    st.download_button(
        label="⬇️ Download Excel",
        data=buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
