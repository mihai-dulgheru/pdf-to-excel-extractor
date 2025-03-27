import pandas as pd

from config import Constants
from functions import convert_to_date

# Map each format string to a conversion function
FORMAT_MAPPINGS = {"General": lambda x: x.astype(str).fillna(""), "@": lambda x: x.astype(str).fillna(""),
                   "0": lambda x: pd.to_numeric(x, errors="coerce").fillna(0).astype("Int64"),
                   "#,##0.00": lambda x: pd.to_numeric(x, errors="coerce").fillna(0).astype(float),
                   "#,##0": lambda x: pd.to_numeric(x, errors="coerce").fillna(0).astype(float),
                   "#,##0.0000": lambda x: pd.to_numeric(x, errors="coerce").fillna(0).astype(float),
                   "dd.mmm": lambda x: x.map(lambda v: convert_to_date(v) if pd.notna(v) else None),
                   "0.00": lambda x: pd.to_numeric(x, errors="coerce").fillna(0).astype(float)}


def enforce_column_formats(df):
    """
    Convert columns in the given DataFrame to types/formats specified in
    Constants.COLUMN_FORMATS, using FORMAT_MAPPINGS.

    Only columns present both in the DataFrame and in COLUMN_FORMATS are processed.
    If a column's assigned format string exists in FORMAT_MAPPINGS, the corresponding
    conversion function is applied.

    :param df: A pandas DataFrame to be converted.
    :return: The same DataFrame with updated column data types.
    """
    # Intersect the DataFrame columns with those defined in COLUMN_FORMATS
    common_cols = set(df.columns) & set(Constants.COLUMN_FORMATS.keys())

    for col in common_cols:
        fmt = Constants.COLUMN_FORMATS[col]
        if fmt in FORMAT_MAPPINGS:
            df[col] = FORMAT_MAPPINGS[fmt](df[col])

    return df
