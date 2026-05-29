# Streamlit Deployment

## Main file

Use:

streamlit_modular.py

## Streamlit Cloud setup

In Streamlit Cloud, set the app entrypoint / main file to:

streamlit_modular.py

Do not use app.py because that is the legacy Flask app.
Do not use streamlit_app.py because that is the older non-modular Streamlit version.

## Requirements

The repository requirements.txt should contain:
- streamlit
- openpyxl

Flask can remain temporarily while the legacy app is still in the repo, but Streamlit Cloud should run streamlit_modular.py.

## Required repository files

reference/bdo_template.xlsm
uploads/
outputs/

## Smartsheet API

Use this API URL format:

https://api.smartsheet.com/2.0/sheets/{sheet_id}

Authentication:
Bearer token

Leave JSON Data Path blank for the normal Smartsheet Get Sheet response.
