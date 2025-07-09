import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Sportsduniya Attendance Summary Generator", page_icon="📝")

st.title("📝 Sportsduniya Attendance Summary Generator")
st.subheader("Upload Attendance Excel File (.xlsx or .xls)")

uploaded_file = st.file_uploader("Upload file", type=["xlsx", "xls"])

def extract_date_from_sheet(sheet):
    try:
        for i in range(10):
            row = sheet.iloc[i].astype(str).tolist()
            for item in row:
                if "202" in item:
                    dt = item.strip()
                    for fmt in ("%d-%b-%Y", "%d-%b-%y", "%d-%m-%Y"):
                        try:
                            return datetime.strptime(dt, fmt).strftime("%d-%m-%Y")
                        except:
                            continue
    except:
        pass
    return ""

def extract_kollegeapply_format(df, sheet):
    for i in range(10):
        row = df.iloc[i].astype(str).str.lower().tolist()
        if "e. code" in row or "e code" in row:
            df.columns = df.iloc[i]
            df = df[i + 1:].reset_index(drop=True)
            break

    df = df.loc[:, ~df.columns.duplicated()]
    df = df.rename(columns=lambda x: str(x).strip())

    expected_cols = [
        "E. Code", "Name", "Shift", "InTime", "OutTime",
        "Work Dur.", "OT", "Tot.  Dur.", "Status", "Remarks"
    ]

    try:
        df = df[expected_cols]
        df = df[df["E. Code"].notnull()]
        df["Date"] = extract_date_from_sheet(sheet)
    except Exception as e:
        st.error("⚠️ Header mismatch or missing expected columns.")
        st.stop()

    return df

if uploaded_file:
    try:
        sheet = pd.read_excel(uploaded_file, sheet_name=0, header=None)

        final_df = extract_kollegeapply_format(sheet.copy(), sheet)
        st.success("✅ File processed successfully!")
        st.dataframe(final_df)

        csv = final_df.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Download Processed CSV", data=csv, file_name="attendance_summary.csv", mime='text/csv')

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
