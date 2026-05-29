import base64
import csv
import json
import urllib.error
import urllib.request

from constants import CONFIG_PATH


def default_api_config():
    return {
        "api_url": "",
        "method": "GET",
        "auth_type": "bearer",
        "bearer_token": "",
        "basic_username": "",
        "basic_password": "",
        "custom_headers": "",
        "json_data_path": "",
    }


def load_api_config():
    config = default_api_config()
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            config.update(json.load(f))
    except FileNotFoundError:
        pass
    except Exception:
        pass
    return config


def save_api_config(config):
    safe = {key: config.get(key, "") for key in default_api_config()}
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(safe, f, indent=2)


def parse_custom_headers(text):
    headers = {}
    for line in (text or "").splitlines():
        if not line.strip() or ":" not in line:
            continue
        key, value = line.split(":", 1)
        headers[key.strip()] = value.strip()
    return headers


def extract_json_path(data, path):
    current = data
    path = (path or "").strip()
    if not path:
        return current

    for part in path.split("."):
        part = part.strip()
        if not part:
            continue
        if isinstance(current, dict):
            current = current.get(part)
        elif isinstance(current, list) and part.isdigit():
            current = current[int(part)]
        else:
            raise ValueError(f"JSON path not found at '{part}'")
    return current


def convert_smartsheet_response(data):
    columns = {col.get("id"): col.get("title") for col in data.get("columns", [])}
    records = []

    for row in data.get("rows", []):
        record = {}
        for cell in row.get("cells", []):
            column_name = columns.get(cell.get("columnId"))
            if column_name:
                record[column_name] = cell.get("displayValue", cell.get("value", ""))
        records.append(record)

    return records


def fetch_api_rows(config):
    api_url = (config.get("api_url") or "").strip()
    if not api_url:
        raise ValueError("API URL is required.")

    method = (config.get("method") or "GET").upper()
    headers = {"Accept": "application/json"}
    headers.update(parse_custom_headers(config.get("custom_headers")))

    auth_type = config.get("auth_type", "none")
    if auth_type == "bearer" and config.get("bearer_token"):
        headers["Authorization"] = f"Bearer {config.get('bearer_token')}"
    elif auth_type == "basic" and config.get("basic_username"):
        raw = f"{config.get('basic_username')}:{config.get('basic_password', '')}".encode("utf-8")
        headers["Authorization"] = "Basic " + base64.b64encode(raw).decode("ascii")

    request = urllib.request.Request(api_url, method=method, headers=headers)
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            body = response.read().decode("utf-8-sig")
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")[:500]
        raise ValueError(f"API returned HTTP {exc.code}: {detail}")

    try:
        data = json.loads(body)
    except json.JSONDecodeError:
        return list(csv.DictReader(body.splitlines()))

    data = extract_json_path(data, config.get("json_data_path"))

    if isinstance(data, dict) and "columns" in data and "rows" in data:
        return convert_smartsheet_response(data)

    if isinstance(data, dict):
        for key in ("rows", "data", "items", "records", "results"):
            if isinstance(data.get(key), list):
                data = data[key]
                break

    if not isinstance(data, list):
        raise ValueError("API response must be a list of rows, a Smartsheet sheet response, or use JSON data path to point to the row list.")

    if not all(isinstance(row, dict) for row in data):
        raise ValueError("API row data must be a list of objects/dictionaries.")

    return data
