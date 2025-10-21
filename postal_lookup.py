import streamlit as st
import pandas as pd
import io

st.title("📬 Postal Code Lookup Tool")

# Relative path to Excel in repo
FILE_PATH = "Book2.xlsx"

@st.cache_data
def load_data(file_path):
    try:
        # Read all sheets
        xls = pd.ExcelFile(file_path)
        df = None
        for sheet_name in xls.sheet_names:
            temp_df = pd.read_excel(xls, sheet_name=sheet_name)
            # Normalize column names
            temp_df.columns = temp_df.columns.str.strip().str.lower()
            if "postal_prefix" in temp_df.columns:
                df = temp_df
                break
        if df is None:
            st.error("❌ No sheet contains the column 'postal_prefix'.")
            st.stop()
        return df
    except FileNotFoundError:
        st.error(f"File not found: {file_path}")
        st.stop()

df = load_data(FILE_PATH)

st.write("Enter one or more postal codes or prefixes (one per line):")
input_text = st.text_area(
    "Postal Codes",
    height=150,
    placeholder="e.g.\nv5w2e4\nV5W 2E6\nv5v"
)

if st.button("🔍 Search"):
    postal_inputs = [x.strip().replace(" ", "").upper() for x in input_text.splitlines() if x.strip()]

    if not postal_inputs:
        st.warning("Please enter at least one postal code or prefix.")
    else:
        df["postal_prefix_clean"] = df["postal_prefix"].astype(str).str.replace(" ", "").str.upper()
        results = df[df["postal_prefix_clean"].apply(lambda x: any(x.startswith(p) for p in postal_inputs))]

        if results.empty:
            st.error("No matching postal codes found.")
        else:
            st.success(f"✅ Found {len(results)} matching records.")
            st.dataframe(results)

            # Prepare CSV download
            csv_data = results.to_csv(index=False)
            # Prepare Excel download
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                results.to_excel(writer, index=False, sheet_name='Matches')

            st.download_button(
                "⬇️ Download CSV",
                data=csv_data,
                file_name="postal_lookup_results.csv",
                mime="text/csv"
            )

            st.download_button(
                "⬇️ Download Excel",
                data=excel_buffer.getvalue(),
                file_name="postal_lookup_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
