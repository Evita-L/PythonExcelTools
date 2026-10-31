"""
Excel Data Cleaner
------------------
Cleans Excel workbooks by:
- Removing completely empty rows and columns
- Removing duplicate rows
- Trimming whitespace from text cells
- Removing columns with automatically generated "Unnamed" headers
- Cleaning every worksheet in the workbook

Usage:
    python excel_cleaner.py input.xlsx
    python excel_cleaner.py input.xlsx --output cleaned.xlsx
    python excel_cleaner.py input.xlsx --keep-duplicates
"""

import argparse
from pathlib import Path

import pandas as pd


def clean_dataframe(df: pd.DataFrame, remove_duplicates: bool = True) -> pd.DataFrame:
    """Clean a worksheet represented as a pandas DataFrame."""

    # Remove columns that pandas reads as unnamed Excel columns.
    df = df.loc[:, ~df.columns.astype(str).str.match(r"^\s*Unnamed", case=False)]

    # Remove rows and columns that are completely empty.
    df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")

    # Trim whitespace from column names.
    df.columns = [
        str(column).strip() if not pd.isna(column) else ""
        for column in df.columns
    ]

    # Trim whitespace in text cells; convert whitespace-only cells to missing values.
    for column in df.select_dtypes(include=["object", "string"]).columns:
        df[column] = df[column].apply(
            lambda value: value.strip() if isinstance(value, str) else value
        )
        df[column] = df[column].replace("", pd.NA)

    # Remove rows that became completely empty after trimming.
    df = df.dropna(axis=0, how="all")

    if remove_duplicates:
        df = df.drop_duplicates()

    return df.reset_index(drop=True)


def clean_workbook(input_path: Path, output_path: Path, remove_duplicates: bool) -> None:
    """Clean all worksheets and save them into a new Excel workbook."""

    sheets = pd.read_excel(input_path, sheet_name=None)

    if not sheets:
        raise ValueError("The workbook does not contain any worksheets.")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, dataframe in sheets.items():
            cleaned = clean_dataframe(dataframe, remove_duplicates)
            cleaned.to_excel(writer, sheet_name=sheet_name, index=False)
            print(
                f"✓ {sheet_name}: {len(dataframe)} -> {len(cleaned)} rows"
            )

    print(f"\nCleaning completed. Output saved to: {output_path}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Clean messy Excel workbooks.")
    parser.add_argument("input", type=Path, help="Path to the input .xlsx file")
    parser.add_argument(
        "--output",
        type=Path,
        help="Path for the cleaned workbook (default: <name>_cleaned.xlsx)",
    )
    parser.add_argument(
        "--keep-duplicates",
        action="store_true",
        help="Keep duplicate rows instead of removing them",
    )
    args = parser.parse_args()

    if not args.input.is_file():
        parser.error(f"Input file does not exist: {args.input}")

    if args.input.suffix.lower() not in {".xlsx", ".xlsm"}:
        parser.error("Please provide an .xlsx or .xlsm workbook.")

    output = args.output or args.input.with_name(
        f"{args.input.stem}_cleaned.xlsx"
    )

    if output.resolve() == args.input.resolve():
        parser.error("Output file must be different from the input file.")

    try:
        clean_workbook(args.input, output, not args.keep_duplicates)
    except Exception as error:
        parser.exit(1, f"Error: {error}\n")


if __name__ == "__main__":
    main()
