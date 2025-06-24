import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font

import os
import re

def excel_to_dataframe(file_path, sheet_name=None):
    """
    Reads an Excel sheet and converts it into a Pandas DataFrame.

    :param file_path: str - Path to the .xlsx file
    :param sheet_name: str (optional) - Name of the sheet to be read. If None, reads the first sheet.
    :return: DataFrame - Content of the sheet in a Pandas DataFrame
    """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def dataframe_to_excel(df, directory, file_name):
    """
    Exports a DataFrame to an Excel file formatted as a table.

    :param df: DataFrame to be exported.
    :param file_name: Name of the file (e.g., 'daily_report.xlsx').
    :param directory: Directory where the file will be saved.
    :return: Full path of the saved file.
    """
    try:
        # Create directory if it doesn't exist
        os.makedirs(directory, exist_ok=True)

        # Define the FULL path of the file
        file_path = os.path.join(directory, file_name)

        # Create Excel file
        wb = Workbook()
        ws = wb.active
        ws.title = "extracted_data"

        # Add headers
        ws.append(df.columns.tolist())
        for cell in ws[1]:
            cell.font = Font(bold=True)

        # Add data
        for row in dataframe_to_rows(df, index=False, header=False):
            ws.append(row)

        # Create formatted table in Excel
        num_rows = len(df) + 1  # +1 to include header
        table_ref = f"A1:{chr(65 + len(df.columns) - 1)}{num_rows}"
        table = Table(displayName="DataTable", ref=table_ref)

        style = TableStyleInfo(
            name="TableStyleMedium9", showFirstColumn=False,
            showLastColumn=False, showRowStripes=True, showColumnStripes=False
        )
        table.tableStyleInfo = style
        ws.add_table(table)

        # Auto-adjust column widths
        for col in ws.columns:
            max_length = max(len(str(cell.value)) for cell in col if cell.value)
            ws.column_dimensions[col[0].column_letter].width = max_length + 2

        # Save the file
        wb.save(file_path)

        return file_path  # Return the correct path of the generated file

    except Exception as e:
        print(f"Error exporting to Excel: {e}")
        return None    

def remove_duplicates(df, column: str):
    """
    Removes duplicates from the DataFrame based on the values of a specific column.

    Parameters:
    ----------
    df : pd.DataFrame
        The original DataFrame from which duplicates will be removed.

    column : str
        The name of the column used as a reference to identify duplicates.

    Returns:
    -------
    pd.DataFrame
        A new DataFrame without duplicate rows based on the specified column.
    """
    if column not in df.columns:
        raise ValueError(f"The column '{column}' does not exist in the DataFrame.")

    df = df.drop_duplicates(subset=[column], keep='first')

    return df

def strip_dataframe(df):
    """
    Strips whitespace from the beginning and end of each string in the DataFrame.

    :param df: DataFrame to be processed.
    :return: DataFrame with stripped strings.
    """
    return df.apply(lambda col: col.str.strip() if col.dtypes == 'object' else col)

def filter_dataframe(df, column, parameter):
    """
    Filters the DataFrame based on a specific column and parameter.

    :param df: DataFrame to be filtered.
    :param column: Column name to filter by.
    :param parameter: Parameter to filter the column.
    :return: Filtered DataFrame or None if the column does not exist.
    """
    if column not in df.columns:
        print(f"The column '{column}' does not exist in the DataFrame.")
        return None
    else:
        return df[df[column] == parameter]

def clean_special_characters(df, column: str):
    """
    Removes special characters from a DataFrame column, keeping only letters, numbers, and spaces.
    Replaces '&' with 'E' before cleaning.

    Parameters:
        df (pd.DataFrame): The original DataFrame.
        column (str): Name of the column to be processed.

    Returns:
        pd.DataFrame: The DataFrame with the processed column.
    """
    if column not in df.columns:
        print(f"The column '{column}' does not exist in the DataFrame.")
        return None

    def clean_text(value):
        if not isinstance(value, str):
            return value  # Ignore if not a string
        value = value.replace('&', 'E')              # Replace '&' with 'E'
        value = re.sub(r'[^a-zA-Z0-9\s]', '', value) # Remove everything that is not a letter, number, or space
        return value

    df.loc[:, column] = df[column].apply(clean_text)

    return df

def lowercase_dataframe(df, column: str):
    """
    Converts all string entries in a specified column of the DataFrame to lowercase.

    :param df: DataFrame to be processed.
    :param column: Column name to convert to lowercase.
    :return: DataFrame with the specified column in lowercase.
    """
    if column not in df.columns:
        print(f"The column '{column}' does not exist in the DataFrame.")
        return None

    df[column] = df[column].apply(lambda x: x.lower() if isinstance(x, str) else x)
    return df

def format_document(df, column: str, size: int):
    """
    Formats numeric documents (CPF or CNPJ) by removing invalid characters
    and adjusting the size with leading zeros or right trimming.
    Keeps empty values as NaN.

    :param df: DataFrame containing the documents to format.
    :param column: Column name with the documents to format.
    :param size: Desired size of the formatted document.
    :return: DataFrame with formatted documents.
    """
    if column not in df.columns:
        print(f"⚠ The column '{column}' does not exist in the DataFrame.")
        return None

    df = df.copy()

    def process(doc):
        # If null or empty, return NaN (keep as empty)
        if pd.isna(doc) or str(doc).strip().lower() in ["", "nan", "none"]:
            return None

        # Convert to string and clean non-numeric characters
        doc_str = str(doc).strip()
        doc_numeric = re.sub(r'\D', '', doc_str)

        # Adjust size as needed
        if len(doc_numeric) < size:
            doc_numeric = doc_numeric.zfill(size)
        elif len(doc_numeric) > size:
            doc_numeric = doc_numeric[:size]

        return doc_numeric

    df[column] = df[column].apply(process)
    return df

def drop_column(df, column: str):
    """
    Removes a column from the DataFrame, if it exists.

    :param df: The original DataFrame.
    :param column: The name of the column to be removed.
    :return: DataFrame without the specified column.
    """
    if column in df.columns:
        df = df.drop(columns=[column])
    else:
        return None
    
    return df

def filter_by_different_char_count(df, column: str, char_count: int) -> pd.DataFrame:
    """
    Filters the DataFrame, returning only records where the character count
    in the specified column is different from the given value.

    :param df: The original DataFrame.
    :param column: The name of the column to be evaluated.
    :param char_count: The character count to be avoided (will be removed if equal).
    :return: A new DataFrame with records where the length of the column values
             is different from the specified count.
    """
    # Check if the column exists in the DataFrame
    if column not in df.columns:
        raise ValueError(f"The column '{column}' does not exist in the DataFrame.")
    df[column] = df[column].astype(str)
    filter_mask = df[column].apply(len) != char_count
    return df[filter_mask]

def limit_column_size(df, column: str, limit: int):
    """
    Limits the size of the values in a DataFrame column.

    :param df: The input DataFrame
    :param column: Name of the column to be truncated
    :param limit: Maximum number of characters allowed
    :return: A copy of the DataFrame with the adjusted column
    """
    df_copy = df.copy()

    if column not in df_copy.columns:
        raise ValueError(f"The column '{column}' does not exist in the DataFrame.")

    df_copy[column] = df_copy[column].astype(str)

    df_copy[column] = df_copy[column].apply(lambda x: x[:limit] if x else x)

    return df_copy

def format_phone_number(df, column: str):
    """
    Formats phone numbers in a DataFrame column to contain only 12 numeric digits.
    Removes any non-numeric characters, pads with leading zeros, or trims excess.

    :param df: Input DataFrame
    :param column: Name of the column with phone numbers
    :return: DataFrame with the formatted column
    """
    if column not in df.columns:
        print(f"The column '{column}' does not exist in the DataFrame.")
        return None

    df = df.copy()
    df[column] = df[column].astype(str)  # Ensure values are strings

    def process_phone(phone):
        # Remove everything that is not a number
        clean_phone = re.sub(r'\D', '', phone)

        return clean_phone

    df[column] = df[column].apply(process_phone)
    return df

def rename_column(df, current_name: str, new_name: str):
    """
    Renames a column in the DataFrame.

    :param df: The DataFrame containing the column to rename.
    :param current_name: The current name of the column.
    :param new_name: The new name for the column.
    :return: DataFrame with the renamed column.
    """
    try:
        if current_name not in df.columns:
            print(f"The column '{current_name}' does not exist in the DataFrame.")
            raise Exception
        renamed_df = df.rename(columns={current_name: new_name})
        return renamed_df
    except Exception as e:
        print(f'Error renaming column: {e}')
