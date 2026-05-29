import os
import uuid

import streamlit as st

from api_client import fetch_api_rows, load_api_config, save_api_config
from auth import is_admin, login_form, logout_button
from constants import BDO_ACA_TEMPLATE_PATH, OUTPUT_DIR, TEMPLATE_PATH, UPLOAD_DIR
from export_router import write_routed_outputs
from file_reader import read_input_file
from validation_engine import validate_and_clean


def apply_custom_styles():
    st.markdown(
        """
        <style>
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(5, minmax(150px, 1fr));
            gap: 1rem;
            margin: 1.25rem 0 1.75rem 0;
        }
        .kpi-card {
            border: 1px solid #e5e7eb;
            border-radius: 14px;
            padding: 1rem 1.1rem;
            background: #ffffff;
            box-shadow: 0 1px 3px rgba(15, 23, 42, 0.08);
            min-height: 112px;
        }
        .kpi-label {
            font-size: 0.92rem;
            color: #475569;
            font-weight: 600;
            margin-bottom: 0.45rem;
        }
        .kpi-value {
            font-size: 2.15rem;
            line-height: 1;
            font-weight: 750;
            color: #0f172a;
        }
        .kpi-card.total { border-left: 7px solid #2563eb; }
        .kpi-card.error { border-left: 7px solid #dc2626; }
        .kpi-card.ready { border-left: 7px solid #16a34a; }
        .kpi-card.bdo { border-left: 7px solid #7c3aed; }
        .kpi-card.nonbdo { border-left: 7px solid #ea580c; }

        div.stButton > button[kind="primary"] {
            background: #dc2626;
            border: 1px solid #991b1b;
            color: white;
            font-weight: 700;
            border-radius: 10px;
            padding: 0.65rem 1.15rem;
        }
        div.stButton > button[kind="primary"]:hover {
            background: #b91c1c;
            border-color: #7f1d1d;
            color: white;
        }
        div.stDownloadButton > button {
            background: #0f766e;
            border: 1px solid #115e59;
            color: white;
            font-weight: 700;
            border-radius: 10px;
            padding: 0.65rem 1.15rem;
        }
        div.stDownloadButton > button:hover {
            background: #115e59;
            border-color: #134e4a;
            color: white;
        }
        @media (max-width: 1100px) {
            .kpi-grid { grid-template-columns: repeat(2, minmax(150px, 1fr)); }
        }
        @media (max-width: 650px) {
            .kpi-grid { grid-template-columns: 1fr; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_kpi_cards(total_rows, error_rows, ready_rows, bdo_rows, non_bdo_rows):
    st.markdown(
        f"""
        <div class="kpi-grid">
            <div class="kpi-card total">
                <div class="kpi-label">Total Rows</div>
                <div class="kpi-value">{total_rows}</div>
            </div>
            <div class="kpi-card error">
                <div class="kpi-label">Rows With Errors</div>
                <div class="kpi-value">{error_rows}</div>
            </div>
            <div class="kpi-card ready">
                <div class="kpi-label">Ready Rows</div>
                <div class="kpi-value">{ready_rows}</div>
            </div>
            <div class="kpi-card bdo">
                <div class="kpi-label">BDO Rows</div>
                <div class="kpi-value">{bdo_rows}</div>
            </div>
            <div class="kpi-card nonbdo">
                <div class="kpi-label">Non-BDO Rows</div>
                <div class="kpi-value">{non_bdo_rows}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


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

    render_kpi_cards(
        total_rows=len(cleaned_rows),
        error_rows=len(error_rows),
        ready_rows=len(cleaned_rows) - len(error_rows),
        bdo_rows=outputs.get("bdo_rows", 0),
        non_bdo_rows=outputs.get("non_bdo_rows", 0),
    )

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


def persist_results(cleaned_rows, error_rows, unmatched, outputs):
    st.session_state["last_reimbursement_results"] = {
        "cleaned_rows": cleaned_rows,
        "error_rows": error_rows,
        "unmatched": unmatched,
        "outputs": outputs,
    }


def render_persisted_results():
    results = st.session_state.get("last_reimbursement_results")
    if not results:
        return

    st.caption("Last processed reimbursement results are still available below.")
    show_routed_results(
        results["cleaned_rows"],
        results["error_rows"],
        results["unmatched"],
        results["outputs"],
    )


def process_rows(rows, source_name):
    cleaned_rows, error_rows, unmatched = validate_and_clean(rows)
    outputs = write_routed_outputs(cleaned_rows, error_rows, source_name)
    persist_results(cleaned_rows, error_rows, unmatched, outputs)
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


def reimbursement_upload_tab():
    st.subheader("Reimbursement Upload")
    st.caption("Pull reimbursement data from the saved configuration, validate it, and generate BDO and Non-BDO upload files.")

    config = load_api_config()
    if not config.get("api_url"):
        st.warning("Reimbursement source has not been configured yet. Please ask an admin to configure it.")
        return

    if st.button("Run Reimbursement", type="primary"):
        try:
            rows = fetch_api_rows(config)
            process_rows(rows, "Reimbursement Source")
        except Exception as exc:
            st.error(f"Reimbursement processing failed: {exc}")

    render_persisted_results()


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
    apply_custom_styles()

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
        tab_upload, tab_reimbursement, tab_api_config = st.tabs(
            ["Upload File", "Reimbursement Upload", "API Config"]
        )
        with tab_upload:
            upload_tab()
        with tab_reimbursement:
            reimbursement_upload_tab()
        with tab_api_config:
            api_config_tab()
    else:
        tab_upload, tab_reimbursement = st.tabs(["Upload File", "Reimbursement Upload"])
        with tab_upload:
            upload_tab()
        with tab_reimbursement:
            reimbursement_upload_tab()


main()
