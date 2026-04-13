import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Universal MIS Tool", page_icon="📊", layout="wide")

st.title("📊 Universal Monthly Variance Analyzer")
st.markdown("Upload any Tally report (CSV or Excel). This tool automatically maps months for any client.")

# 1. SUPPORT BOTH FORMATS
uploaded_file = st.file_uploader("Upload Master Trial Balance", type=["csv", "xlsx"])

if uploaded_file:
    try:
        # 2. LOAD DATA SAFELY
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None).fillna("")
        else:
            df_raw = pd.read_excel(uploaded_file, header=None).fillna("")

        # 3. SCAN FOR DATA START
        # We look for 'Particulars' anywhere in the file
        header_row = None
        for i in range(len(df_raw)):
            if any("Particulars" in str(val) for val in df_raw.iloc[i].values):
                header_row = i
                break

        if header_row is None:
            st.error("Could not find 'Particulars' column. Please check the export format.")
            st.stop()

        # 4. DYNAMIC MONTH DETECTION
        # We look 1-3 rows above Particulars to find Month names
        possible_month_rows = [header_row-1, header_row-2, header_row-3]
        months_row = []
        for r in possible_month_rows:
            if r >= 0:
                row_vals = [str(v).strip() for v in df_raw.iloc[r] if str(v).strip() not in ["", "nan"]]
                if len(row_vals) > 1: # Found the row with month names
                    months_row = df_raw.iloc[r].tolist()
                    break

        sub_headers = df_raw.iloc[header_row].tolist()
        
        # Forward-fill month names (e.g., 'Feb' applies to its Debit, Credit, Balance)
        current_m = "Opening"
        final_cols = []
        for m, s in zip(months_row, sub_headers):
            m_str = str(m).strip()
            s_str = str(s).strip()
            if m_str and m_str.lower() != "nan":
                current_m = m_str
            
            if s_str == "Particulars":
                final_cols.append("Particulars")
            elif s_str and s_str.lower() != "nan":
                final_cols.append(f"{current_m} - {s_str}")
            else:
                final_cols.append(f"Empty_{len(final_cols)}")

        # 5. PREPARE DATAFRAME
        df_clean = df_raw.iloc[header_row + 1:].copy()
        df_clean.columns = final_cols
        # Remove garbage columns
        df_clean = df_clean.loc[:, ~df_clean.columns.str.contains('Empty_|^0$')]
        
        # 6. DYNAMIC DROPDOWNS
        # We filter for 'Balance' so the user only picks the final monthly figures
        bal_cols = [c for c in df_clean.columns if 'Balance' in c]
        
        st.sidebar.header("Select Months to Compare")
        m1 = st.sidebar.selectbox("Base Month (Older)", bal_cols, index=0)
        m2 = st.sidebar.selectbox("Comparison Month (Newer)", bal_cols, index=len(bal_cols)-1)

        # 7. CURRENCY CLEANING FUNCTION
        def clean_fin_val(x):
            try:
                s = str(x).replace(' Dr', '').replace(' Cr', '').replace(',', '').strip()
                return float(s)
            except: return 0.0

        # 8. PROCESS ANALYSIS
        report = df_clean[['Particulars', m1, m2]].copy()
        report[m1] = report[m1].apply(clean_fin_val)
        report[m2] = report[m2].apply(clean_fin_val)
        report['Variance'] = report[m2] - report[m1]
        report['% Change'] = (report['Variance'] / report[m1].replace(0, 1))

        # 9. DISPLAY & EXPORT
        st.subheader(f"Results for {m2} vs {m1}")
        st.dataframe(report, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            report.to_excel(writer, sheet_name='Variance_Report', index=False)
            wb, ws = writer.book, writer.sheets['Variance_Report']
            
            # Formats
            hdr_f = wb.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1})
            num_f = wb.add_format({'num_format': '#,##0.00', 'border': 1})
            pct_f = wb.add_format({'num_format': '0.0%', 'border': 1})
            red_f = wb.add_format({'bg_color': '#F4CCCC', 'font_color': '#990000'})
            grn_f = wb.add_format({'bg_color': '#D9EAD3', 'font_color': '#38761D'})

            ws.set_column('A:A', 45)
            ws.set_column('B:D', 18, num_f)
            ws.set_column('E:E', 12, pct_f)

            for i, col in enumerate(report.columns):
                ws.write(0, i, col, hdr_f)

            ws.conditional_format(1, 3, len(report), 3, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': red_f})
            ws.conditional_format(1, 3, len(report), 3, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': grn_f})

        st.download_button("📥 Download Excel Report", output.getvalue(), "MIS_Variance_Report.xlsx")

    except Exception as e:
        st.error(f"Processing Error: {e}")
