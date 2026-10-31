# Python Data Automation Tools

Two practical Python utilities for Excel data cleaning and SQL database importing.

## Projects

### 1. Excel Data Cleaner (`excel_cleaner.py`)
- Removes completely empty rows and columns.
- Removes duplicate rows by default.
- Trims whitespace from text values and headers.
- Removes automatically generated `Unnamed` columns.
- Processes every worksheet and saves a separate cleaned workbook.
- Supports keeping duplicates with `--keep-duplicates`.

Run:
```bash
python excel_cleaner.py input.xlsx
python excel_cleaner.py input.xlsx --output cleaned.xlsx
```

### 2. Excel to SQL Converter (`excel_to_sql.py`)
- Imports an Excel worksheet into a SQLite database.
- Creates the SQL table automatically.
- Normalizes column names into SQL-friendly identifiers.
- Handles duplicate column names.
- Allows choosing the worksheet, database file, and table name.

Run:
```bash
python excel_to_sql.py products.xlsx
python excel_to_sql.py products.xlsx --table products --database store.db
python excel_to_sql.py products.xlsx --sheet Sheet2 --table customers
```

**Important:** The converter uses `if_exists="replace"` for the selected table. Running it again replaces that table's existing contents. Back up your database first if you need to preserve existing data.

## Requirements

```bash
pip install pandas openpyxl
```

SQLite is included with Python, so no separate database installation is required.

## Notes
- These tools work with `.xlsx` and `.xlsm` input workbooks.
- The cleaner exports cell values, not the original Excel formatting, formulas, charts, or macros.
- Test on a copy of important workbooks before using them in a production workflow.
