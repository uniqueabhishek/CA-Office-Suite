"""
DataFrame cleaning and formatting transformations.

Shared by the desktop Excel Formatter and the Flask web formatter so both
apply identical rules. Contains no UI or threading dependencies.
"""

import logging
import warnings

import pandas as pd

logger = logging.getLogger(__name__)

# Excel strftime patterns keyed by the option the UI offers.
DATE_FORMATS = {
    "dd-mm-yyyy": "%d-%m-%Y",
    "yyyy-mm-dd": "%Y-%m-%d",
    "mm/dd/yyyy": "%m/%d/%Y",
}


def detect_and_convert_numbers(df):
    """
    Detect columns where many entries are numbers
    encoded as strings and convert them.
    Strategy:
    - For each column, try pd.to_numeric on the
    series with errors='coerce'.
    - If a reasonable proportion (>=30%) convertable or
    header suggests numeric type, coerce.
    """
    df2 = df.copy()
    conversions = {}
    for col in df2.columns:
        series = df2[col].replace("", pd.NA)
        # attempt numeric conversion
        converted = pd.to_numeric(series, errors="coerce")
        notnull = converted.notnull().sum()
        total = len(series) - series.isna().sum()

        # choose threshold: if at least 30% values
        # convert to numeric AND at least 2 values

        if total > 0 and notnull >= max(2, int(0.3 * total)):
            # Use converted values where possible
            df2[col] = converted.where(converted.notnull(), series)
            conversions[col] = True
    return df2, conversions


def trim_whitespace(df):
    """Trims whitespace from string columns using vectorized operations."""
    df2 = df.copy()
    for col in df2.columns:
        if df2[col].dtype == object:
            # Vectorized string operation - 10-50x faster than apply(lambda)
            df2[col] = df2[col].str.strip()
    return df2


def normalize_dates(df, target_format="dd-mm-yyyy"):
    """
    Detects and normalizes date columns to a consistent string format.
    """
    df2 = df.copy()
    fmt = DATE_FORMATS.get(target_format, "%d-%m-%Y")
    conversions = {}
    for col in df2.columns:
        # errors="coerce" turns unparseable values into NaT rather than raising,
        # so only a column pandas cannot interpret as dates at all lands here.
        try:
            with warnings.catch_warnings():
                # Columns of mixed or ambiguous date strings make pandas warn
                # that it cannot infer one format and is falling back to
                # dateutil. That fallback is the intent here: the data is
                # arbitrary user input, and dayfirst=True already settles the
                # ambiguous cases the warning is about.
                warnings.simplefilter("ignore", UserWarning)
                parsed = pd.to_datetime(df2[col], errors="coerce", dayfirst=True)
        except (TypeError, ValueError, OverflowError) as exc:
            logger.debug("Column %r left untouched, not parseable as dates: %s", col, exc)
            continue
        num_parsed = parsed.notna().sum()
        if num_parsed >= 2:  # threshold
            # format back to string in target format
            df2[col] = parsed.dt.strftime(fmt)
            conversions[col] = True
    return df2, conversions


def apply_number_formatting(df, option="2_decimals", currency_symbol=None):
    """
    Applies number formatting to numeric columns.
    Returns df (string representation may be applied) and a
    number_format_map for openpyxl.
    """
    df2 = df.copy()
    number_format_map = {}
    for col in df2.columns:
        # try detect numeric column
        converted = pd.to_numeric(df2[col], errors="coerce")
        if converted.notna().sum() >= 1:
            if option == "no_decimals":
                df2[col] = converted.round(0).astype("Int64").astype(object)
                number_format_map[col] = "#,##0"
            elif option == "2_decimals":
                df2[col] = converted.round(2)
                number_format_map[col] = "#,##0.00"
            elif option == "currency":
                # put numeric (float) in df; formatting will show currency in Excel
                df2[col] = converted.round(2)
                symbol = currency_symbol if currency_symbol else "₹"
                # Excel currency format example: '₹#,##0.00'
                number_format_map[col] = f'"{symbol}"#,##0.00'
    return df2, number_format_map


def apply_text_case(df, case_option="none"):
    """
    Applies a specified text case (upper, lower, title) to string columns using vectorized operations.
    """
    df2 = df.copy()
    if case_option == "none":
        return df2
    for col in df2.columns:
        if df2[col].dtype == object:
            # Vectorized string operations - 10-50x faster than apply(lambda)
            if case_option == "upper":
                df2[col] = df2[col].str.upper()
            elif case_option == "lower":
                df2[col] = df2[col].str.lower()
            elif case_option == "title":
                df2[col] = df2[col].str.title()
    return df2


def remove_duplicates(df):
    """Removes duplicate rows from the DataFrame."""
    return df.drop_duplicates()


def apply_all_transformations(df, options):
    """
    Run the enabled transformations in the order the UI presents them.

    This delegates to the single-purpose functions above rather than repeating
    their bodies. The desktop preview calls those functions one at a time while
    the batch and web paths call this one, so a second copy of the rules here
    would let the preview and the saved output drift apart.

    Each step returns a new frame and the previous one is released immediately,
    so at most two frames are alive at once regardless of how many options are
    enabled. The input is never modified.

    Args:
        df: Input DataFrame
        options: Dict with keys: trim, numbers, dates, date_format, text_case,
                 remove_dups, number_format, number_format_option, currency_symbol

    Returns:
        (processed_df, conversions_dict, number_format_map)
    """
    result = df
    conversions = {}
    number_format_map = {}

    if options.get("trim", False):
        result = trim_whitespace(result)

    if options.get("numbers", False):
        result, numeric_conversions = detect_and_convert_numbers(result)
        conversions.update(numeric_conversions)

    if options.get("dates", False):
        result, date_conversions = normalize_dates(result, options.get("date_format", "dd-mm-yyyy"))
        conversions.update(date_conversions)

    case_option = options.get("text_case") or "none"
    if case_option != "none":
        result = apply_text_case(result, case_option)

    if options.get("number_format", False):
        result, number_format_map = apply_number_formatting(
            result,
            options.get("number_format_option", "2_decimals"),
            options.get("currency_symbol", "₹"),
        )

    if options.get("remove_dups", False):
        result = remove_duplicates(result)

    # Callers may modify what they get back, so never hand out the input itself.
    if result is df:
        result = df.copy()
    return result, conversions, number_format_map
