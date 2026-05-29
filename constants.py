import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "reference", "bdo_template.xlsm")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
CONFIG_PATH = os.path.join(BASE_DIR, "api_config.json")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

DATA_START_ROW = 12
TEMPLATE_SHEET = "BOB Outward Payments"
BANK_CODES_SHEET = "Bank Codes"

COLUMNS = [
    "Amount", "Type of Beneficiary", "Beneficiary First Name", "Beneficiary Middle Name",
    "Beneficiary Last Name or Corporate Beneficiary Name", "Beneficiary Address 1",
    "Beneficiary Address 2", "Beneficiary Address 3", "Beneficiary Account Number",
    "Beneficiary Bank", "Beneficiary Bank Address", "Beneficiary Information",
    "Bank to Bank Information", "Purpose Code", "Nature of Transfer", "Swift Code",
    "Country of Destination", "Importers Code", "Routing Number", "RTGS Purpose Code",
]

MANDATORY_FIELDS = [
    "Amount", "Type of Beneficiary", "Beneficiary Last Name or Corporate Beneficiary Name",
    "Beneficiary Address 1", "Beneficiary Address 2", "Beneficiary Address 3",
    "Beneficiary Account Number", "Beneficiary Bank", "Beneficiary Bank Address",
    "Purpose Code", "Country of Destination",
]

CLEAN_TEXT_FIELDS = [
    "Beneficiary First Name", "Beneficiary Middle Name",
    "Beneficiary Last Name or Corporate Beneficiary Name", "Beneficiary Address 1",
    "Beneficiary Address 2", "Beneficiary Address 3", "Beneficiary Bank Address",
    "Beneficiary Information", "Bank to Bank Information", "Nature of Transfer",
    "Country of Destination", "Swift Code", "Importers Code", "Routing Number",
]
ADDRESS_FIELDS = ["Beneficiary Address 1", "Beneficiary Address 2", "Beneficiary Address 3"]

HEADER_ALIASES = {
    "account number": "Beneficiary Account Number",
    "beneficiary account number": "Beneficiary Account Number",
    "beneficiary last nameor corporate beneficiary nam": "Beneficiary Last Name or Corporate Beneficiary Name",
    "beneficiary last name or corporate beneficiary name": "Beneficiary Last Name or Corporate Beneficiary Name",
    "beneficiary bank": "Beneficiary Bank",
    "beneficiary bank address": "Beneficiary Bank Address",
    "bank to bank information": "Bank to Bank Information",
    "country of destination": "Country of Destination",
    "rtgs purpose code": "RTGS Purpose Code",
    "swift code": "Swift Code",
    "purpose code": "Purpose Code",
}

BANK_ALIASES = {
    "METROBANK": "METROPOLITAN BANK AND TRUST CO",
    "METROPOLITAN BANK": "METROPOLITAN BANK AND TRUST CO",
    "BPI": "BANK OF THE PHILIPPINE ISLANDS",
    "RCBC": "RIZAL COMMERCIAL BANKING CORP",
    "SECURITY BANK": "SECURITY BANK CORPORATION",
    "CHINABANK SAVINGS": "CHINA BANK SAVINGS",
    "CHINA BANK SAVINGS": "CHINA BANK SAVINGS",
    "CHINABANK": "CHINA BANKING CORPORATION",
    "BDO": "BDO UNIBANK INC",
    "BDO UNIBANK": "BDO UNIBANK INC",
    "GCASH": "GCASH",
    "METRO BANK": "METROPOLITAN BANK AND TRUST CO",
    "BANK OF THE PHILIPPINE ISLAND": "BANK OF THE PHILIPPINE ISLANDS",
    "BANK OF THE PHILIPPINE ISLAND BPI": "BANK OF THE PHILIPPINE ISLANDS",
    "BPI BANK": "BANK OF THE PHILIPPINE ISLANDS",
    "HELLOMONEY AUB": "ASIA UNITED BANK",
    "MAYA PHILIPPINES": "PAYMAYA PHILIPPINES INC",
    "PAYMAYA": "PAYMAYA PHILIPPINES INC",
    "PS BANK": "PHILIPPINE SAVINGS BANK",
    "RCBC DISKARTECH": "RIZAL COMMERCIAL BANKING CORP",
    "RCBC SAVINGS BANK": "RIZAL COMMERCIAL BANKING CORP",
    "UNION BANK": "UNION BANK OF THE PHILIPPINES",
    "UNIONBANK": "UNION BANK OF THE PHILIPPINES",
    "UNIONBANK OF THE PHILIPPINES": "UNION BANK OF THE PHILIPPINES",
}

IGNORED_COLUMNS = {
    "date endorsed", "employer name", "status", "tracking no", "tracking no.",
    "upload date", "uploaded", "cells", "createdat", "expanded", "id",
    "modifiedat", "rownumber", "siblingid",
}
