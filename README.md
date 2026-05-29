# BDO Reimbursement Cleanup App - Python Version with API Option

This version keeps the original upload workflow and adds an API input option.

## Run on Windows

1. Extract this ZIP first.
2. Open Command Prompt.
3. Go to the extracted folder:

```bat
cd /d "C:\Users\alfred\Downloads\bdo_cleanup_python_app_v3_api\bdo_cleanup_python_app_v3_api"
```

4. Create and activate a virtual environment:

```bat
python -m venv .venv
.venv\Scripts\activate
```

5. Install requirements:

```bat
pip install -r requirements.txt
```

6. Start the app:

```bat
python app.py
```

7. Open this in your browser:

```text
http://127.0.0.1:5000
```

## Input options

### Option 1: Upload file
Upload CSV, XLSX, or XLSM and click **Validate & Generate BDO File**.

### Option 2: API input
1. Open **API Config**.
2. Enter the API URL.
3. Choose authentication:
   - None
   - Bearer Token
   - Basic Username/Password
4. Optional: add custom headers, one per line:

```text
x-api-key: your-key
client-id: abc123
```

5. Optional: set JSON Data Path if the records are nested.

Examples:

If API returns:

```json
[
  {"Amount": 100, "Beneficiary Bank": "BPI"}
]
```

Leave JSON Data Path blank.

If API returns:

```json
{
  "data": {
    "items": [
      {"Amount": 100, "Beneficiary Bank": "BPI"}
    ]
  }
}
```

Use:

```text
data.items
```

Then open **Use API** and click **Fetch API Data & Generate BDO File**.

## API response format

The API must return one of these:

- JSON array of rows
- JSON object with a row array inside `data`, `rows`, `items`, `records`, or `results`
- CSV text

Each row should use headers/keys similar to the upload file or BDO template.
