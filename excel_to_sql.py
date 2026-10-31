"""
Excel to SQL (SQLite) Converter
-------------------------------
Imports an Excel worksheet into a SQLite database.
The SQL table is created automatically.

Usage:
    python excel_to_sql.py products.xlsx
    python excel_to_sql.py products.xlsx --table products --database store.db
    python excel_to_sql.py products.xlsx --sheet Sheet2 --table customers
"""

import argparse
import re
import sqlite3
from pathlib import Path

import pandas as pd


def normalize_column_name(name: str) -> str:
    """Convert a column heading into a simple SQL-friendly identifier."""
    normalized = re.sub(r"\W+", "_", str(name).strip().lower())
    normalized = normalized.strip("_")

    if not normalized:
        normalized = "column"

    if normalized[0].isdigit():
        normalized = f"col_{normalized}"

    return normalized


def convert_excel_to_sqlite(
    excel_path: Path,
    database_path: Path,
    table_name: str,
    sheet_name: str | None = None,
) -> None:
    """Read Excel data and write it to a SQLite table."""

    if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", table_name):
        raise ValueError("Table name must contain only letters, numbers, and underscores.")

    dataframe = pd.read_excel(excel_path, sheet_name=sheet_name)

    # Remove fully empty rows and columns.
    dataframe = dataframe.dropna(axis=0, how="all").dropna(axis=1, how="all")

    if dataframe.empty:
        raise ValueError("The selected worksheet contains no data.")

    # Normalize headers and ensure they remain unique.
    normalized_columns = []
    used_names = set()

    for original_name in dataframe.columns:
        base_name = normalize_column_name(original_name)
        candidate = base_name
        suffix = 2

        while candidate in used_names:
            candidate = f"{base_name}_{suffix}"
            suffix += 1

        normalized_columns.append(candidate)
        used_names.add(candidate)

    dataframe.columns = normalized_columns

    # Create the database and replace the target table with the imported data.
    with sqlite3.connect(database_path) as connection:
        dataframe.to_sql(
            name=table_name,
            con=connection,
            if_exists="replace",
            index=False,
        )

    print(f"✓ Imported {len(dataframe)} rows.")
    print(f"✓ SQL table: {table_name}")
    print(f"✓ Database: {database_path}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert an Excel worksheet into a SQLite database table."
    )
    parser.add_argument("input", type=Path, help="Path to the input .xlsx file")
    parser.add_argument(
        "--database",
        type=Path,
        default=Path("output.db"),
        help="Output SQLite database (default: output.db)",
    )
    parser.add_argument(
        "--table",
        default="imported_data",
        help="Name of the SQL table (default: imported_data)",
    )
    parser.add_argument(
        "--sheet",
        default=None,
        help="Worksheet name (default: first worksheet)",
    )
    args = parser.parse_args()

    if not args.input.is_file():
        parser.error(f"Input file does not exist: {args.input}")

    if args.input.suffix.lower() not in {".xlsx", ".xlsm"}:
        parser.error("Please provide an .xlsx or .xlsm workbook.")

    try:
        convert_excel_to_sqlite(
            args.input,
            args.database,
            args.table,
            args.sheet,
        )
    except Exception as error:
        parser.exit(1, f"Error: {error}\n")


if __name__ == "__main__":
    main()
