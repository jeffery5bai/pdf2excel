from datetime import datetime

import pandas as pd
import pdfplumber
import streamlit as st

from excel_writer.write_style_excel import write_styled_excel
from pdf_parser.wholesale_parser import parse_po_text

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
        required_keys = [
            "PO#",
            "Material",
            "Description",
            "Qty",
            "Unit Price",
            "Create Date",
            "Due Date",
            "GT CRD",
        ]
        result_list = []
        revised_result_list = []
        original_files = []
        revised_files = []
        failed_files = []

        for upload_file in uploaded_files:
            file_type = "original"
            try:
                with pdfplumber.open(upload_file) as pdf:
                    full_text = ""
                    for page in pdf.pages:
                        text = page.extract_text()
                        if text:
                            full_text += text + "\n"
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
            po_info = parse_po_text(full_text, file_type=file_type)

            if not po_info:
                st.warning(
                    f"⚠️ Warning: Can not parse/extract information from this PDF file: {upload_file.name}, file type: {file_type}"
                )
                failed_files.append(upload_file.name)
                continue

            missing_keys = [k for k in required_keys if k not in po_info]
            if missing_keys:
                st.warning(
                    f"⚠️ Warning: PDF {upload_file.name} (file type: {file_type}) missing columns: {missing_keys}, skipped."
                )
                failed_files.append(upload_file.name)
                continue

            if file_type == "revised":
                revised_result_list.append(po_info)
                revised_files.append(upload_file.name)
            else:
                result_list.append(po_info)
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
            df = (
                original_df.drop_duplicates(subset=["PO#"], keep="last")
                .sort_values("PO#")
                .reset_index(drop=True)
            )

            st.session_state.info = {
                "original_files": original_files,
                "revised_files": revised_files,
                "failed_files": failed_files,
            }
            st.session_state.df = df
            st.session_state.excel_bytes = write_styled_excel(df)
            st.session_state.failed_files = failed_files

if st.session_state.df is not None:
    st.dataframe(st.session_state.df.head())

    st.download_button(
        label="📥 Download Excel",
        data=st.session_state.excel_bytes,
        file_name=f"new_orders_{datetime.now().date().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if not st.session_state.info["failed_files"]:
        st.success("✅ All files parsed successfully!")
    else:
        st.info(
            f"Original files: {len(st.session_state.info['original_files'])}, revised files: {len(st.session_state.info['revised_files'])}"
        )
        st.warning(
            f"⚠️ Failed to parse some files below: {st.session_state.info['failed_files']}.\nPlease check them again."
        )

    st.success("Please download the Excel file.")
