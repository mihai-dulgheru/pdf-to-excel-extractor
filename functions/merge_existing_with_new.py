import os

import pandas as pd

from config import Constants
from functions import enforce_column_formats


def merge_existing_with_new(existing_excel_path, processed_df):
    """
    Merge an existing Excel file with newly processed invoice data while ensuring correct data types.

    :param existing_excel_path: Path to the existing Excel file.
    :param processed_df: DataFrame with newly extracted invoices.
    :return: Merged DataFrame with correctly formatted data.
    """
    # Check if the existing Excel file path is valid
    if not existing_excel_path or not os.path.exists(existing_excel_path):
        print(f"[LOG] Existing Excel file not found: {existing_excel_path}")
        return processed_df  # If the file is missing, return only the new data

    try:
        # Read the existing Excel file as a DataFrame (all values as strings to avoid type issues)
        existing_df = pd.read_excel(existing_excel_path, dtype=str)

        # Reverse lookup to rename columns back to their expected internal names
        reverse_headers = {v: k for k, v in Constants.HEADERS.items()}
        renamed_columns = {col: reverse_headers[col] for col in existing_df.columns if col in reverse_headers}
        if renamed_columns:
            existing_df = existing_df.rename(columns=renamed_columns)

        # Filter out invalid rows in 'nr_crt' (only keep numeric values)
        if 'nr_crt' in existing_df.columns:
            existing_df = existing_df[existing_df['nr_crt'].apply(
                lambda x: pd.notna(x) and (isinstance(x, (int, float)) or (isinstance(x, str) and x.isdigit())))]
            existing_df['nr_crt'] = pd.to_numeric(existing_df['nr_crt'], errors='coerce')

        # Ensure correct data types based on predefined column formats
        formatted_df = enforce_column_formats(existing_df, Constants.COLUMN_FORMATS)

        # Keep only columns that exist in the new dataset to ensure consistency
        common_columns = list(processed_df.columns)
        formatted_df = formatted_df[common_columns]

        # Combine the old and new datasets into a single DataFrame
        combined_df = pd.concat([formatted_df, processed_df], ignore_index=True)

        return combined_df

    except Exception as e:
        print(f"[LOG] Error loading existing Excel: {e}")
        return processed_df  # If an error occurs, return only the new data
