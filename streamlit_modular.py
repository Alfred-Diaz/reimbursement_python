import os
import uuid

import streamlit as st

from api_client import fetch_api_rows, load_api_config, save_api_config
from auth import is_admin, login_form, logout_button
from constants import OUTPUT_DIR, TEMPLATE_PATH, UPLOAD_DIR
from excel_writer import write_clean_output
from file_reader import read_input_file
from validation_engine import validate_and_clean


def show_results(cleaned_rows, error_rows, unmatched, output_path):
    st.success("Processing complete.")

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Rows", len(cleaned_rows))
    col2.metric("Rows With Errors", len(error_rows))
    col3.metric("Ready Rows", len(cleaned_rows) - len(error_rows))

    if unmatched:
        st.warning("Unmatched input columns: " + ", ".join(unmatched))

    if error_rows:
        st.subheader("Validation Error Preview")
        st.dataframe(error_rows[:50], use_container_width=True)

    with open(output_path, "rb") as file:
        st.download_button(
            "Download Cleaned BDO Upload File",
            data=file,
            file_name=os.path.basename(output_path),
            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
        )


def upload_tab():
    uploaded_file = st.file_uploader(
        "Upload CSV, XLSX, or XLSM",
        type=["csv", "xlsx", "xlsm"],
    )

    if uploaded_file and st.button("Validate Uploaded File", type="primary"):
        saved_path = os.path.join(
            UPLOAD_DIR,
            f"{uuid.uuid4().hex}_{uploaded_file.name}",
        )

        with open(saved_path, "wb") as file:
            file.write(uploaded_file.getbuffer())

        try:
            rows = read_input_file(saved_path)
            cleaned_rows, error_rows, unmatched = validate_and_clean(rows)
            output_path = write_clean_output(
                cleaned_rows,
                error_rows,
                uploaded_file.name,
            )
            show_results(cleaned_rows, error_rows, unmatched, output_path)
        except Exception as exc:
            st.error(f"Processing failed: {exc}")


def run_api_tab():
    st.subheader("Run API Import")
    st.caption("Pull data from the saved API configuration, validate it, and generate the BDO upload file.")

    config = load_api_config()
    if not config.get("api_url"):
        st.warning("API configuration has not been set yet. Please ask an admin to configure it.")
        return

    if st.button("Fetch and Validate API Data", type="primary"):
        try:
            rows = fetch_api_rows(config)
            cleaned_rows, error_rows, unmatched = validate_and_clean(rows)
            output_path = write_clean_output(
                cleaned_rows,
                error_rows,
                "Configured API",
            )
            show_results(cleaned_rows, error_rows, unmatched, output_path)
        except Exception as exc:
            st.error(f"API processing failed: {exc}")


def api_config_tab():
    if not is_admin():
        st.error("You do not have permission to access API settings.")
        return

    config = load_api_config()

    st.subheader("API Configuration")
    st.caption(
        "For Smartsheet, use: https://api.smartsheet.com/2.0/sheets/YOUR_SHEET_ID"
    )

    api_url = st.text_input("API URL", value=config.get("api_url", ""))
    bearer_token = st.text_input(
        "Bearer Token",
        value=config.get("bearer_token", ""),
        type="password",
    )
    json_data_path = st.text_input(
        "JSON Data Path",
        value=config.get("json_data_path", ""),
    )
    custom_headers = st.text_area(
        "Custom Headers",
        value=config.get("custom_headers", ""),
    )

    current_config = {
        "api_url": api_url,
        "method": "GET",
        "auth_type": "bearer",
        "bearer_token": bearer_token,
        "basic_username": "",
        "basic_password": "",
        "custom_headers": custom_headers,
        "json_data_path": json_data_path,
    }

    if st.button("Save API Config", type="primary"):
        try:
            save_api_config(current_config)
            st.success("API configuration saved.")
        except Exception as exc:
            st.error(f"Saving failed: {exc}")


def main():
    st.set_page_config(
        page_title="BDO Reimbursement Cleanup",
        layout="wide",
    )

    st.title("BDO Reimbursement Cleanup Tool")
    st.write(
        "Validate, clean, and export reimbursement data into the required BDO upload template."
    )

    if not login_form():
        return

    logout_button()

    if not os.path.exists(TEMPLATE_PATH):
        st.error("Missing reference/bdo_template.xlsm in the GitHub repository.")
        return

    os.makedirs(UPLOAD_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if is_admin():
        tab_upload, tab_run_api, tab_api_config = st.tabs(
            ["Upload File", "Run API", "API Config"]
        )
        with tab_upload:
            upload_tab()
        with tab_run_api:
            run_api_tab()
        with tab_api_config:
            api_config_tab()
    else:
        tab_upload, tab_run_api = st.tabs(["Upload File", "Run API"])
        with tab_upload:
            upload_tab()
        with tab_run_api:
            run_api_tab()


main()
