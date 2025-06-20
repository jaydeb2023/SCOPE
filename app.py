import streamlit as st
import zipfile
import pandas as pd
import os
import tempfile
import re
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

# === Streamlit UI ===
st.set_page_config(page_title="BOQ Scope Extractor", layout="wide")
st.title("📦 BOQ Scope Extractor (DI Pipe - K7/K9)")
st.markdown("""
Upload a **ZIP file (max 1 GB)** containing **TDR-wise folders** with BOQ Excel files.  
The app will extract pipe-related items (DIA ≥ 80 mm) and export them as structured Excel output.
""")

uploaded_zip = st.file_uploader("Upload BOQ ZIP", type=["zip"])

# === Helper Functions ===
def clean_illegal_chars(value):
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub("", value)
    return value

def extract_dia(text):
    text = str(text).lower()
    match = re.search(r'(\d+)\s*mm', text)
    if match:
        return int(match.group(1))
    return None

def find_column(columns, keywords):
    for col in columns:
        if any(k in str(col).lower() for k in keywords):
            return col
    return None

def process_boq_file(file_path, tdr_name, file_name):
    try:
        df_raw = pd.read_excel(file_path, header=None)
    except Exception as e:
        return [{"TDR Folder": tdr_name, "BOQ File": file_name, "Item Description": f"❌ Failed to read file: {e}"}]

    header_row_index = None
    for i in range(min(30, len(df_raw))):
        row = df_raw.iloc[i].astype(str).str.lower()
        if any("desc" in cell or "item" in cell for cell in row) and any("unit" in cell for cell in row):
            header_row_index = i
            break

    if header_row_index is None:
        return [{"TDR Folder": tdr_name, "BOQ File": file_name, "Item Description": "⚠️ Header not found"}]

    try:
        df = pd.read_excel(file_path, header=header_row_index)
    except Exception as e:
        return [{"TDR Folder": tdr_name, "BOQ File": file_name, "Item Description": f"❌ Error loading structured data: {e}"}]

    desc_col = find_column(df.columns, ['desc', 'item'])
    unit_col = find_column(df.columns, ['unit'])
    qty_col = find_column(df.columns, ['qty', 'quantity'])
    rate_col = find_column(df.columns, ['rate', 'amount', 'estimate'])

    if not desc_col or not unit_col:
        return [{"TDR Folder": tdr_name, "BOQ File": file_name, "Item Description": "⚠️ Missing required columns"}]

    df[unit_col] = df[unit_col].astype(str).str.strip().str.lower().replace({
        "per metre": "meter", "rm": "meter", "rmt": "meter", "mtr": "meter", "mtrs": "meter"
    })

    def is_pipe_row(row):
        desc = str(row[desc_col]).lower()
        return any(x in desc for x in ["pipe", "di", "ci", "m.s", "k-7", "k-9", "hdpe", "pvc", "upvc"])

    filtered_df = df[df.apply(is_pipe_row, axis=1)].copy()
    if filtered_df.empty:
        return []

    filtered_df["Item Description"] = filtered_df[desc_col].astype(str)
    filtered_df["DIA"] = filtered_df["Item Description"].apply(extract_dia)
    filtered_df["estimate rate"] = pd.to_numeric(filtered_df[rate_col], errors='coerce') if rate_col else 0
    filtered_df["Units"] = df[unit_col].astype(str).str.strip()
    filtered_df["Quantity"] = pd.to_numeric(df[qty_col], errors='coerce') if qty_col else None
    filtered_df["K-9"] = filtered_df["Item Description"].str.lower().str.contains("k-9|k9").map({True: "Yes", False: ""})
    filtered_df["K-7"] = filtered_df["Item Description"].str.lower().str.contains("k-7").map({True: "Yes", False: ""})
    filtered_df["TDR Folder"] = tdr_name
    filtered_df["BOQ File"] = file_name

    final_df = filtered_df[[
        "TDR Folder", "BOQ File", "Item Description", "K-9", "K-7",
        "DIA", "estimate rate", "Units", "Quantity"
    ]].copy()

    return final_df.to_dict(orient="records")

# === Main Logic ===
if uploaded_zip:
    with tempfile.TemporaryDirectory() as tmp_dir:
        zip_path = os.path.join(tmp_dir, "uploaded.zip")
        with open(zip_path, "wb") as f:
            f.write(uploaded_zip.read())

        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(tmp_dir)

        combined_rows = []
        for root, dirs, files in os.walk(tmp_dir):
            for file in files:
                if file.lower().endswith(('.xls', '.xlsx')) and not file.startswith("~$"):
                    full_path = os.path.join(root, file)
                    tdr_folder = os.path.basename(os.path.dirname(full_path))
                    combined_rows.extend(process_boq_file(full_path, tdr_folder, file))

        if combined_rows:
            df_summary = pd.DataFrame(combined_rows)

            # Filter DIA >= 80
            df_summary["DIA"] = pd.to_numeric(df_summary["DIA"], errors="coerce")
            df_summary = df_summary[df_summary["DIA"] >= 80]

            # Insert serial number and reorder columns
            df_summary.insert(0, "SL No", range(1, len(df_summary) + 1))
            desired_order = [
                "SL No", "TDR Folder", "BOQ File", "Item Description", "K-9",
                "K-7", "DIA", "estimate rate", "Units", "Quantity"
            ]
            df_summary = df_summary[desired_order]
            df_summary = df_summary.astype(str).map(clean_illegal_chars)

            st.success(f"✅ Extracted {len(df_summary)} rows with DIA ≥ 80.")
            st.dataframe(df_summary)

            output_path = os.path.join(tmp_dir, "BOQ_All_Combined_Summary.xlsx")
            df_summary.to_excel(output_path, index=False)

            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 Download Excel Summary",
                    data=f,
                    file_name="BOQ_All_Combined_Summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("⚠️ No matching BOQ rows found.")
