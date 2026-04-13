import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Universal MIS Tool", page_icon="📊", layout="wide")

st.title("📊 Universal Monthly Variance Analyzer")
st.markdown("Optimized for Master Client Files with Multiple Months.")

uploaded_file = st.file_uploader("Upload Master Trial Balance", type=["csv", "xlsx"])

if uploaded_file:
    try:
        # 1. LOAD DATA
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None).fillna("")
        else:
            df_raw = pd.read_excel(uploaded_file, header=None).fillna("")

        # 2. FIND 'PARTICULARS'
        header_row_idx = None
        for i in range(len(df_raw)):
            if any("Particulars" in str(val) for val in df_raw.iloc[i].values):
                header_row_idx = i
                break

        if header_row_idx is None:
            st.error("Could not find 'Particulars' column.")
            st.stop()

        # 3. THE MONTH HUNTER
        # We look for month names in the rows above 'Particulars'
        month_keywords = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
        
        months_row = [""] * len(df_raw.columns)
        for r in range(max(0, header_row_idx-5), header_row_idx):
            row_vals = df_raw.iloc[r].tolist()
            # If this row contains any month names, we use it
            if any(any(m in str(v).lower() for m in month_keywords) for v in row_vals):
                months_row = row_vals
                break

        sub_headers = [str(s).strip() for s in df_raw.iloc[header_row_idx].tolist()]
        
        # 4. COMBINE LABELS
        current_m = "Balance"
        final_cols = []
        for i, (m, s) in enumerate(zip(months_row, sub_headers)):
            m_str = str(m).strip()
            s_str = str(s).strip()
            
            # If we found a new month name in this column, update current_m
            if m_str and m_str.lower() != "nan" and len(m_str) > 2:
                current_m = m_str
            
            if s_str == "Particulars":
                final_cols.append("Particulars")
            elif s_str and s_str.lower() != "nan":
                # Label as "Oct - Closing Balance (4)"
                final_cols.append(f"{current_m} - {s_str} ({i})")
            else:
                final_cols.append(f"Hidden_{i}")

        # 5. DATA PREP
        df_clean = df_raw.iloc[header_row_idx + 1:].copy()
        df_clean.columns = final_cols
        df_clean = df_clean.loc[:, ~df_clean.columns.str.contains('Hidden_')]
        
        # Filter dropdown to show only "Balance" or "Closing" columns
        compare_options = [c for c in df_clean.columns if "Particulars" not in c]
        balance_options = [c for c in compare_options if "Balance" in c or "Closing" in c]
        
        # If no 'Balance' text found, show all options
        if not balance_options: balance_options = compare_options

        # 6. SIDEBAR
        st.sidebar.header("Select Months")
        m1 = st.sidebar.selectbox("Base Month (Older)", balance_options, index=0)
        m2 = st.sidebar.selectbox("Comparison Month (Newer)", balance_options, index=len(balance_options)-1)

        # 7. CLEAN & CALCULATE
        def clean_val(x):
            try:
                s = str(x).replace(' Dr', '').replace(' Cr', '').replace(',', '').strip()
                return float(s)
            except: return 0.0

        report = df_clean[['Particulars', m1, m2]].copy()
        report.columns = ['Particulars', 'Old', 'New']
        report['Old'] = report['Old'].apply(clean_val)
        report['New'] = report['New'].apply(clean_val)
        report['Variance'] = report['New'] - report['Old']
        report['% Change'] = (report['Variance'] / report['Old'].replace(0, 1))

        # Final Formatting for display
        final_view = report.copy()
        final_view.columns = ['Particulars', m1, m2, 'Variance', '% Change']

        st.subheader(f"Analyzing: {m1} vs {m2}")
        st.dataframe(final_view, use_container_width=True)

        # 8. EXCEL EXPORT
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_view.to_excel(writer, sheet_name='MIS_Report', index=False)
            wb, ws = writer.book, writer.sheets['MIS_Report']
            ws.set_column('A:A', 40)
            ws.set_column('B:D', 18, wb.add_format({'num_format': '#,##0.00'}))
            ws.set_column('E:E', 12, wb.add_format({'num_format': '0.0%'}))

        st.download_button("📥 Download MIS Report", output.getvalue(), "MIS_Analysis.xlsx")

    except Exception as e:
        st.error(f"Error: {e}")
