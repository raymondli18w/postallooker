import streamlit as st
import pandas as pd
import io
import os

# Full Excel path
FILE_PATH = r"C:\Users\RaymondLi\OneDrive - 18wheels.ca\downloads may 30 2023\test6\postallook\Book2.xlsx"

st.title("📬 Postal Code Lookup Tool")

@st.cache_data
def load_data(file_path):
    if not os.path.exists(file_path):
        st.error(f"File not found: {file_path}")
        st.stop()
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()  # remove spaces in column names
    return df

df = load_data(FILE_PATH)

st.write("Enter one or more postal codes or prefixes (one per line):")
input_text = st.text_area("Postal Codes", height=150, placeholder="e.g.\nv5w2e4\nV5W 2E6\nv5v")

if st.button("🔍 Search"):
    postal_inputs = [x.strip().replace(" ", "").upper() for x in input_text.splitlines() if x.strip()]
    
    if not postal_inputs:
        st.warning("Please enter at least one postal code or prefix.")
    else:
        if "postal_prefix" not in df.columns:
            st.error("❌ Column 'postal_prefix' not found in Book2.xlsx. Please check the column name.")
        else:
            df["POSTAL_PREFIX_CLEAN"] = df["postal_prefix"].astype(str).str.replace(" ", "").str.upper()
            results = df[df["POSTAL_PREFIX_CLEAN"].apply(lambda x: any(x.startswith(p) for p in postal_inputs))]

            if results.empty:
                st.error("No matching postal codes found.")
            else:
                st.success(f"✅ Found {len(results)} matching records.")
                st.dataframe(results)

                # Download options
                csv_data = results.to_csv(index=False)
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
