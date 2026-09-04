"""
Tests for core.transformations - the cleaning rules shared by the desktop
formatter and the web formatter.
"""

import pandas as pd

from core.transformations import (
    apply_all_transformations,
    apply_number_formatting,
    apply_text_case,
    detect_and_convert_numbers,
    normalize_dates,
    remove_duplicates,
    trim_whitespace,
)


class TestTrimWhitespace:
    def test_strips_both_ends(self):
        df = pd.DataFrame({"A": ["  x  ", " y"]})
        assert list(trim_whitespace(df)["A"]) == ["x", "y"]

    def test_leaves_the_input_unchanged(self):
        df = pd.DataFrame({"A": ["  x  "]})
        trim_whitespace(df)
        assert df.loc[0, "A"] == "  x  "


class TestDetectAndConvertNumbers:
    def test_converts_a_numeric_text_column(self):
        df = pd.DataFrame({"Amount": ["10", "20", "30"]})
        out, conversions = detect_and_convert_numbers(df)
        assert conversions.get("Amount") is True
        assert list(out["Amount"]) == [10, 20, 30]

    def test_leaves_a_text_column_alone(self):
        df = pd.DataFrame({"Name": ["alice", "bob", "carol"]})
        out, conversions = detect_and_convert_numbers(df)
        assert "Name" not in conversions
        assert list(out["Name"]) == ["alice", "bob", "carol"]

    def test_keeps_unconvertible_values_in_a_mixed_column(self):
        df = pd.DataFrame({"Mixed": ["10", "20", "N/A", "40"]})
        out, _ = detect_and_convert_numbers(df)
        assert out.loc[2, "Mixed"] == "N/A"
        assert out.loc[0, "Mixed"] == 10


class TestNormalizeDates:
    def test_reformats_to_the_target_format(self):
        df = pd.DataFrame({"Date": ["01-02-2024", "15-03-2024"]})
        out, conversions = normalize_dates(df, target_format="yyyy-mm-dd")
        assert conversions.get("Date") is True
        assert list(out["Date"]) == ["2024-02-01", "2024-03-15"]

    def test_defaults_to_day_first(self):
        df = pd.DataFrame({"Date": ["01-02-2024", "15-03-2024"]})
        out, _ = normalize_dates(df)
        assert out.loc[0, "Date"] == "01-02-2024"

    def test_ignores_a_column_with_too_few_dates(self):
        df = pd.DataFrame({"Name": ["alice", "bob"]})
        out, conversions = normalize_dates(df)
        assert "Name" not in conversions
        assert list(out["Name"]) == ["alice", "bob"]


class TestApplyTextCase:
    def test_upper(self):
        df = pd.DataFrame({"A": ["abc"]})
        assert apply_text_case(df, "upper").loc[0, "A"] == "ABC"

    def test_lower(self):
        df = pd.DataFrame({"A": ["ABC"]})
        assert apply_text_case(df, "lower").loc[0, "A"] == "abc"

    def test_title(self):
        df = pd.DataFrame({"A": ["bob smith"]})
        assert apply_text_case(df, "title").loc[0, "A"] == "Bob Smith"

    def test_none_is_a_no_op(self):
        df = pd.DataFrame({"A": ["bob smith"]})
        assert apply_text_case(df, "none").loc[0, "A"] == "bob smith"


class TestApplyNumberFormatting:
    def test_two_decimals(self):
        df = pd.DataFrame({"A": ["1.567", "2.001"]})
        out, nf_map = apply_number_formatting(df, option="2_decimals")
        assert nf_map["A"] == "#,##0.00"
        assert out.loc[0, "A"] == 1.57

    def test_no_decimals(self):
        df = pd.DataFrame({"A": ["1.6", "2.4"]})
        out, nf_map = apply_number_formatting(df, option="no_decimals")
        assert nf_map["A"] == "#,##0"
        assert out.loc[0, "A"] == 2

    def test_currency_symbol_lands_in_the_format(self):
        df = pd.DataFrame({"A": ["10"]})
        _, nf_map = apply_number_formatting(df, option="currency", currency_symbol="$")
        assert nf_map["A"] == '"$"#,##0.00'

    def test_currency_defaults_to_rupee(self):
        df = pd.DataFrame({"A": ["10"]})
        _, nf_map = apply_number_formatting(df, option="currency")
        assert nf_map["A"] == '"₹"#,##0.00'


class TestRemoveDuplicates:
    def test_drops_repeated_rows(self):
        df = pd.DataFrame({"A": [1, 1, 2], "B": ["x", "x", "y"]})
        assert len(remove_duplicates(df)) == 2


class TestApplyAllTransformations:
    """The combined pipeline both front-ends actually call."""

    def test_disabled_options_change_nothing(self, messy_df):
        out, conversions, nf_map = apply_all_transformations(messy_df, {})
        assert out.equals(messy_df)
        assert conversions == {}
        assert nf_map == {}

    def test_trim_and_case_together(self, messy_df):
        out, _, _ = apply_all_transformations(messy_df, {"trim": True, "text_case": "upper"})
        assert out.loc[0, "Name"] == "ALICE"

    def test_numbers_are_converted(self, messy_df):
        out, conversions, _ = apply_all_transformations(messy_df, {"numbers": True})
        assert conversions.get("Amount") is True
        assert out.loc[0, "Amount"] == 1000

    def test_dates_are_normalized(self, messy_df):
        out, _, _ = apply_all_transformations(messy_df, {"dates": True, "date_format": "yyyy-mm-dd"})
        assert out.loc[0, "Date"] == "2024-02-01"

    def test_number_format_map_is_returned(self, messy_df):
        _, _, nf_map = apply_all_transformations(
            messy_df, {"number_format": True, "number_format_option": "2_decimals"}
        )
        assert nf_map.get("Amount") == "#,##0.00"

    def test_duplicates_removed(self):
        df = pd.DataFrame({"A": ["1", "1", "2"]})
        out, _, _ = apply_all_transformations(df, {"remove_dups": True})
        assert len(out) == 2

    def test_input_is_never_mutated(self, messy_df):
        before = messy_df.copy()
        apply_all_transformations(
            messy_df,
            {"trim": True, "numbers": True, "dates": True, "text_case": "upper"},
        )
        assert messy_df.equals(before)

    def test_matches_the_individual_helpers(self, messy_df):
        """The combined pipeline must agree with the step-by-step functions."""
        combined, _, _ = apply_all_transformations(messy_df, {"trim": True, "text_case": "title"})
        stepwise = apply_text_case(trim_whitespace(messy_df), "title")
        assert combined.equals(stepwise)
