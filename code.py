import streamlit as st
import pandas as pd
import openpyxl

st.title("Spreadsheet Processor for Custom Columns")

# File upload
uploaded_file = st.file_uploader("Upload your spreadsheet", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    # Load the file
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file, engine='openpyxl')

    st.write("### Original Spreadsheet:")
    st.dataframe(df)

    # Ensure necessary columns are present
    if "规格属性" in df.columns and "SKCID" in df.columns:
        # Extracting columns
        df["款号编码"] = df["SKCID"].apply(lambda x: x[:2])
        df["颜色编码"] = df["规格属性"].apply(lambda x: x.split("/")[0] if pd.notna(x) else "")
        df["尺寸编码"] = df["规格属性"].apply(lambda x: x.split("/")[1] if pd.notna(x) else "")
        df["图片编码"] = df["SKCID"].apply(lambda x: x.split("-")[0])
        df["工艺类型"] = "白墨烫画"

        st.write("### Processed Spreadsheet:")
        st.dataframe(df)

        # Download button - Maintaining Excel format
        with pd.ExcelWriter(uploaded_file, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')

        with open(uploaded_file.name, "rb") as file:
            st.download_button(
                "Download Processed Spreadsheet",
                file.read(),
                uploaded_file.name,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    else:
        st.error("Uploaded file must contain '规格属性' and 'SKCID' columns.")
