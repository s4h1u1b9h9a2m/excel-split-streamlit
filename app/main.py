import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile

def split_excel_by_company(uploaded_file):
    try:
        # Read the Excel file
        df = pd.read_excel(uploaded_file, engine='openpyxl')

        if column_name not in df.columns:
            st.error(f"Column '{column_name}' not found in the uploaded file.")
            return None
        
        # Create a dictionary to store dataframes for each unique value in the selected column
        unique_values = df[column_name].unique()
        dataframes = {value: df[df[column_name] == value] for value in unique_values}

        # Create a dictionary to store dataframes for each company
        # company_dfs = {company: df[df['Company'] == company] for company in df['Company'].unique()}

        # Create a BytesIO object for each unique value's dataframe
        output_files = {}
        for value, data in dataframes.items():
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                data.to_excel(writer, index=False, sheet_name=str(value))
            output.seek(0)
            output_files[value] = output

        return output_files
    except Exception as e:
        st.error(f"An error occurred: {e}")
        return None

def create_zip(output_files):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zf:
        for value, data in output_files.items():
            zf.writestr(f"{value}.xlsx", data.getvalue())
    zip_buffer.seek(0)
    return zip_buffer

st.title("Split Excel Sheet by Column")
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    column_name = st.selectbox("Select the column to split by", df.columns)

    if st.button("Split Excel File"):
        output_files = split_excel_by_company(uploaded_file)
        if output_files:
            st.success(f"Excel file split into {len(output_files)} files based on '{column_name}' column.")

            # Create and provide download link for the ZIP file
            zip_buffer = create_zip(output_files)
            st.download_button(label="Download All as ZIP", data=zip_buffer, file_name="all_files.zip")

            # Provide download links for each file
            for value, output in output_files.items():
                st.download_button(label=f"Download {value}.xlsx", data=output, file_name=f"{value}.xlsx")
        else:
            st.error("An error occurred while splitting the Excel file.")
