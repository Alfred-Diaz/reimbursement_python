import os
import uuid

import streamlit as st

from api_client import fetch_api_rows, load_api_config, save_api_config
from auth import is_admin, login_form, logout_button
from constants import BDO_ACA_TEMPLATE_PATH, OUTPUT_DIR, TEMPLATE_PATH, UPLOAD_DIR
from export_router import write_routed_outputs
from file_reader import read_input_file
from validation_engine import validate_and_clean


def download_file_button(label, output_path, key):
    with open(output_path, "rb") as file:
        st.download_button(
            label,
            data=file,
            file_name=os.path.basename(output_path),
            mime="application/vnd.ms-excel.sheet.macroEnabled.12",
            key=key,
        )


def render_output_section(title, row_count, errors, output_path, button_label, button_key):
    st.divider()
    st.subheader(title)
    st.metric("Rows Routed", row_count)

    if errors:
        st.warning("Please review the notes below before downloading the file.")
        st.dataframe(errors[:50], use_container_width=True)
    else:
        st.info("No validation notes for this output.")

    if output_path:
        download_file_button(button_label, output_path, button_key)
    else:
        st.caption("No downloadable file was generated for this output.")


def show_routed_results(cleaned_rows, error_rows, unmatched, outputs):
    st.success("Processing complete.")

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Rows", len(cleaned_rows))
    col2.metric("Rows With Errors", len(error_rows))
    col3.metric("Ready Rows", len(cleaned_rows) - len(error_rows))

    route1, route2 = st.columns(2)
    route1.metric("BDO Rows", outputs.get("bdo_rows", 0))
    route2.metric("Non-BDO Rows", outputs.get("non_bdo_rows", 0))

    if unmatched:
        st.warning("Unmatched input columns: " + ", ".join(unmatched))

    for warning in outputs.get("warnings", []):
        st.warning(warning)

    non_bdo_errors = outputs.get("non_bdo_errors", [])
    bdo_errors = outputs.get("bdo_errors", [])

    if outputs.get("non_bdo_rows", 0):
        render_output_section(
            "Non-BDO Upload File",
            outputs.get("non_bdo_rows", 0),
            non_bdo_errors,
            outputs.get("non_bdo_output_path"),
            "Download Non-BDO Upload File",
            "download_non_bdo",
        )

    if outputs.get("bdo_rows", 0):
        render_output_section(
            "BDO ACA Upload File",
            outputs.get("bdo_rows", 0),
            bdo_errors,
            outputs.get("bdo_output_path"),
            "Download BDO ACA Upload File",
            "download_bdo_aca",
        )

    if error_rows:
        with st.expander("View All Validation Errors"):
            st.dataframe(error_rows[:100], use_container_width=True)


def process_rows(rows, source_name):
    cleaned_rows, error_rows, unmatched = validate_and_clean(rows)
    outputs = write_routed_outputs(cleaned_rows, error_rows, source_name)
    show_routed_results(cleaned_rows, error_rows, unmatched, outputs)


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
            process_rows(rows, uploaded_file.name)
        except Exception as exc:
            st.error(f"Processing failed: {exc}")


def run_api_tab():
    st.subheader("Run API Import")
    st.caption("Pull data from the saved API configuration, validate it, and generate BDO and Non-BDO upload files.")

    config = load_api_config()
    if not config.get("api_url"):
        st.warning("API configuration has not been set yet. Please ask an admin to configure it.")
        return

    if st.button("Fetch and Validate API Data", type="primary"):
        try:
            rows = fetch_api_rows(config)
            process_rows(rows, "Configured API")
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
        "Validate, clean, auto-route, and export reimbursement data into BDO and Non-BDO upload templates."
    )

    if not login_form():
        return

    logout_button()

    if not os.path.exists(TEMPLATE_PATH):
        st.error("Missing reference/bdo_template.xlsm in the GitHub repository.")
        return

    if not os.path.exists(BDO_ACA_TEMPLATE_PATH):
        st.warning("BDO ACA template is not yet available. Upload reference/bdo_aca_template.xlsm to enable BDO output.")

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
