from datetime import datetime

import pandas as pd
import pdfplumber
import streamlit as st

from excel_writer.retail import RetailExcelWriter
from excel_writer.wholesale import WholesaleExcelWriter
from pdf_parser.retail_parser import RetailPOParser
from pdf_parser.wholesale_parser import WholesalePOParser

# -------------------- Streamlit App --------------------
st.set_page_config(layout="centered")  # default

st.markdown(
    """
    <style>
        /* block-container */
        .block-container {
            max-width: 1000px;
            padding-left: 2rem;
            padding-right: 2rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Hi Angel! Welcome to Your Workspace!")
st.subheader("Purchase Order PDF Parser → Excel")

mode = st.radio(
    "Select order type:",
    options=["Wholesale", "Retail"],
    horizontal=True,
)
st.session_state["mode"] = mode

# Initialize session_state
if "df" not in st.session_state:
    st.session_state.df = None
if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None
if "failed_files" not in st.session_state:
    st.session_state.failed_files = []

uploaded_files = st.file_uploader(
    "Please Upload PDF files here", type="pdf", accept_multiple_files=True
)
st.write(f"Uploaded files: {len(uploaded_files) if uploaded_files else 0}")

if st.button("Run!"):
    if uploaded_files:
        if mode == "Wholesale":
            po_parser = WholesalePOParser()
            excel_writer = WholesaleExcelWriter()
        else:
            po_parser = RetailPOParser()
            excel_writer = RetailExcelWriter()

        result_list = []
        revised_result_list = []
        original_files = []
        failed_files = []
        revised_files = []
        required_keys = excel_writer.output_schema

        for upload_file in uploaded_files:
            if (mode == "Wholesale" and "KP" not in upload_file.name) or (
                mode == "Retail" and "DI" not in upload_file.name
            ):
                st.warning(
                    f"⚠️ Warning: PDF {upload_file.name} seems not a valid file type in mode {mode}, skipped."
                )
                failed_files.append(upload_file.name)
                continue

            file_type = "original"
            try:
                with pdfplumber.open(upload_file) as pdf:
                    full_text = ""
                    for page in pdf.pages:
                        text = page.extract_text()
                        if text:
                            full_text += text + "\n"

                    if mode == "Retail":
                        words = pdf.pages[0].extract_words()

            except Exception as e:
                st.warning(
                    f"⚠️ Warning: Failed to open/parse PDF: {upload_file.name} -> {e}"
                )
                failed_files.append(upload_file.name)
                continue

            if (
                "This Purchase Order has been changed. Specific changes are shown in red."
                in full_text
            ):
                file_type = "revised"
            if mode == "Wholesale":
                po_info = po_parser.parse_po_content(full_text, file_type=file_type)
            else:  # Retail
                po_info = po_parser.parse_po_content(full_text, words)

            if not po_info:
                st.warning(
                    f"⚠️ Warning: Can not parse/extract information from this PDF file: {upload_file.name}, file type: {file_type}"
                )
                failed_files.append(upload_file.name)
                continue

            missing_keys = [k for k in required_keys if k not in po_info[0]]
            if missing_keys:
                st.warning(
                    f"⚠️ Warning: PDF {upload_file.name} (file type: {file_type}) missing columns: {missing_keys}, skipped."
                )
                failed_files.append(upload_file.name)
                continue

            if file_type == "revised":
                revised_result_list.extend(po_info)
                revised_files.append(upload_file.name)
            else:
                result_list.extend(po_info)
                original_files.append(upload_file.name)

        if not (result_list or revised_result_list):
            st.error(
                "Error: Can not successfully parse ANY PDF files, no report will be generated."
            )
            st.session_state.df = None
            st.session_state.excel_bytes = None
        else:
            original_df = pd.DataFrame(result_list)
            revised_df = pd.DataFrame(revised_result_list)
            if not revised_df.empty:
                original_df = pd.concat([original_df, revised_df], ignore_index=True)
            id_cols = ["PO#"] if mode == "Wholesale" else ["Kohler PO", "Kohler SKU"]
            df = (
                original_df[required_keys]
                .drop_duplicates(subset=id_cols, keep="last")
                .sort_values(id_cols)
                .reset_index(drop=True)
            )

            st.session_state.file_info = {
                "original_files": original_files,
                "revised_files": revised_files,
                "failed_files": failed_files,
            }
            st.session_state.df = df
            st.session_state.excel_bytes = excel_writer.write_excel(df)

if st.session_state.df is not None:
    st.dataframe(st.session_state.df.head())

    st.download_button(
        label="📥 Download Excel",
        data=st.session_state.excel_bytes,
        file_name=f"{mode.lower()}_orders_{datetime.now().date().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if not st.session_state.file_info["failed_files"]:
        st.success("✅ All files parsed successfully!")
    else:
        st.info(
            f"Original files: {len(st.session_state.file_info['original_files'])}, revised files: {len(st.session_state.file_info['revised_files'])}"
        )
        st.warning(
            f"⚠️ Failed to parse some files below: {st.session_state.file_info['failed_files']}.\nPlease check them again."
        )

    st.success("Please download the Excel file.")
