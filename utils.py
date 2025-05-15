import pandas as pd

EXCEL_FILE_PATH = "Data/capbudg.xls"

def load_excel_file(path=EXCEL_FILE_PATH):
    """
    Loads the Excel file and returns a dictionary of sheets.
    """
    try:
        return pd.read_excel(path, sheet_name=None)
    except Exception as e:
        raise RuntimeError(f"Failed to read Excel file: {e}")

def list_tables():
    """
    Lists all sheet names (tables) in the Excel file.
    """
    data = load_excel_file()
    return list(data.keys())

def get_table_details(table_name: str):
    """
    Returns first-column values (row names) of a given sheet.
    """
    data = load_excel_file()
    if table_name not in data:
        return None
    df = data[table_name]
    if df.empty:
        return []
    first_column = df.iloc[:, 0].dropna().astype(str).tolist()
    return first_column

def row_sum(table_name: str, row_name: str):
    """
    Returns the sum of numeric values in a specified row by its first-column label.
    """
    data = load_excel_file()
    if table_name not in data:
        return None
    df = data[table_name]
    df = df.dropna(how='all')
    
    # Match row where first column matches row_name
    match = df[df.iloc[:, 0].astype(str).str.strip() == row_name.strip()]
    if match.empty:
        return None
    
    row = match.iloc[0, 1:]  # exclude first column (label)
    cleaned_row = row.replace('%', '', regex=True)
    numeric_row = pd.to_numeric(cleaned_row, errors='coerce')
    total = numeric_row.sum(skipna=True)
    return total
