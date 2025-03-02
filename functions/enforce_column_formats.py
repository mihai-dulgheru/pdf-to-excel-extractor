import pandas as pd

from functions import convert_to_date


def enforce_column_formats(df, column_formats):
    """
    Ensures that DataFrame columns follow the expected data types based on COLUMN_FORMATS.

    :param df: DataFrame to be formatted.
    :param column_formats: Dictionary containing expected formats.
    :return: Formatted DataFrame with the correct data types.
    """

    # Mapping of format strings to corresponding conversion functions
    format_mappings = {"General": lambda x: x.astype(str).fillna(""),
                       "0": lambda x: pd.to_numeric(x, errors="coerce").fillna(0).astype("Int64"),
                       "#,##0.00": lambda x: pd.to_numeric(x, errors="coerce").fillna(0),
                       "#,##0.00;-#,##0.00": lambda x: pd.to_numeric(x, errors="coerce").fillna(0),
                       "#,##0": lambda x: pd.to_numeric(x, errors="coerce").fillna(0),
                       "#,##0;-#,##0": lambda x: pd.to_numeric(x, errors="coerce").fillna(0),
                       "dd.mmm": lambda x: x.map(lambda v: convert_to_date(v) if pd.notna(v) else None),
                       "#,##0.0000;-#,##0.0000": lambda x: pd.to_numeric(x, errors="coerce").fillna(0).astype(float),
                       "0.00": lambda x: pd.to_numeric(x, errors="coerce").fillna(0).astype(float)}

    # Iterate through all columns in the DataFrame and apply the corresponding format transformation
    for column, format_str in column_formats.items():
        if column in df.columns and format_str in format_mappings:
            df[column] = format_mappings[format_str](df[column])  # Apply the appropriate formatting function

    return df
