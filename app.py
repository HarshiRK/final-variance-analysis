import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Universal MIS Tool", page_icon="📊", layout="wide")

st.title("📊 Universal Monthly Variance Analyzer")
st.markdown("Upload any Tally report. This version is optimized to find months even in messy files.")

uploaded_file = st.file_uploader("Upload Master Trial Balance", type=["csv", "xlsx"])

if uploaded_file:
    try:
        # 1. LOAD DATA
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None).fillna("")
        else:
            df_raw = pd.read_excel(uploaded_file, header=None).fillna("")

        # 2. FIND THE DATA START (Search for 'Particulars')
        header_row_idx = None
        for i in range(len(df_raw)):
            if any("Particulars" in str(val) for val in df_raw.iloc[i].values):
                header_row_idx = i
                break

        if header_row_idx is None:
            st.error("Could not find 'Particulars' column. Please check your Tally export.")
            st.stop()

        # 3. ROBUST HEADER EXTRACTION
        # We check 3 rows above for months. If we find nothing, we use generic labels.
        sub_headers = [str(s).strip() for s in df_raw.iloc[header_row_idx].tolist()]
        
        # Try to find the most 'word-heavy' row above headers to use as months
        best_month_row = []
        max_words = 0
        for r in range(max(0, header_row_idx-3), header_row_idx):
            row_content = [str(v).strip() for v in df_raw.iloc[r] if len(str(v).strip()) > 2]
            if len(row_content) > max_words:
                max_words = len(row_content)
                best_month_row = df_raw.iloc[r].tolist()

        # Build final column names
        current_m = "Data"
        final_cols = []
        for i, (m, s) in enumerate(zip(best_month_row if best_month_row else [""]*len(sub_headers), sub_headers)):
            m_str = str(m).strip()
            s_str = str(s).strip()
            
            if m_str and m_str.lower() != "nan" and len(m_str) > 2:
                current_m = m_str
            
            if s_str == "Particulars":
                final_cols.append("Particulars")
            elif s_str and s_str.lower() != "nan":
                # Create a unique name: Month + Subheader + Index (to avoid duplicates)
                final_cols.append(f"{current_m} - {s_str} ({i})")
            else:
                final_cols.append(f"Empty_{i}")

        # 4. PREPARE DATAFRAME
        df_clean = df_raw.iloc[header_row_idx + 1:].copy()
        df_clean.columns = final_cols
        df_clean = df_clean.loc[:, ~df_clean.columns.str.contains('Empty_')]
        
        # Identify columns for the dropdown
        # We look for anything that looks like a Balance or Closing
        compare_options = [c for c in df_clean.columns if "Particulars" not in c]
        
        # 5. SIDEBAR
        st.sidebar.header("Select Columns")
        m1 = st.sidebar.selectbox("Base Column (Older)", compare_options, index=0)
        m2 = st.sidebar.selectbox("Comparison Column (Newer)", compare_options, index=len(compare_options)-1)

        # 6. MATH ENGINE
        def clean_val(x):
            try:
                s = str(x).replace(' Dr', '').replace(' Cr', '').replace(',', '').strip()
                return float(s)
            except: return 0.0

        report = df_clean[['Particulars', m1, m2]].copy()
        report.columns = ['Particulars', 'Old_Period', 'New_Period']
        report['Old_Period'] = report['Old_Period'].apply(clean_val)
        report['New_Period'] = report['New_Period'].apply(clean_val)
        report['Variance'] = report['New_Period'] - report['Old_Period']
        report['% Change'] = (report['Variance'] / report['Old_Period'].replace(0, 1))

        # Rename back for display
        display_report = report.copy()
        display_report.columns = ['Particulars', m1, m2, 'Variance', '% Change']

        # 7. OUTPUT
        st.subheader(f"Comparison: {m1} vs {m2}")
        st.dataframe(display_report, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            display_report.to_excel(writer, sheet_name='MIS_Variance', index=False)
            wb, ws = writer.book, writer.sheets['MIS_Variance']
            
            hdr = wb.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1})
            num = wb.add_format({'num_format': '#,##0.00', 'border': 1})
            pct = wb.add_format({'num_format': '0.0%', 'border': 1})
            
            ws.set_column('A:A', 45)
            ws.set_column('B:D', 18, num)
            ws.set_column('E:E', 12, pct)
            for i, col in enumerate(display_report.columns):
                ws.write(0, i, col, hdr)

        st.download_button("📥 Download Excel Report", output.getvalue(), "MIS_Variance.xlsx")

    except Exception as e:
        st.error(f"Something went wrong: {e}")
