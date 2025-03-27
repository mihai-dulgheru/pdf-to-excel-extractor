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
    if not existing_excel_path or not os.path.exists(existing_excel_path):
        print(f"[LOG] Existing Excel file not found: {existing_excel_path}")
        return processed_df

    try:
        existing_df = pd.read_excel(existing_excel_path, dtype=str)

        nr_crt_display_name = Constants.HEADERS.get('nr_crt')
        if nr_crt_display_name in existing_df.columns:
            existing_df = existing_df[existing_df[nr_crt_display_name].apply(
                lambda x: pd.notna(x) and (isinstance(x, (int, float)) or (isinstance(x, str) and x.isdigit())))]

        reverse_headers = {v: k for k, v in Constants.HEADERS.items()}
        renamed = {col: reverse_headers[col] for col in existing_df.columns if col in reverse_headers}
        if renamed:
            existing_df = existing_df.rename(columns=renamed)

        common_columns = list(processed_df.columns)
        existing_df = existing_df[common_columns]

        formatted_df = enforce_column_formats(existing_df)

        combined_df = pd.concat([formatted_df, processed_df], ignore_index=True)

        return combined_df

    except Exception as e:
        print(f"[LOG] Error loading existing Excel: {e}")
        return processed_df
