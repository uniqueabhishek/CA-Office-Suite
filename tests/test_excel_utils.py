"""
Tests for core.excel_utils - file discovery, reading and name sanitising.
"""

import os

import pandas as pd

from core.excel_utils import list_excel_files_in_folder, read_all_sheets, read_file_to_df, safe_name


class TestSafeName:
    def test_path_separators_are_replaced(self):
        assert "/" not in safe_name("Q1/Q2")
        assert "\\" not in safe_name("a\\b")

    def test_excel_illegal_characters_are_replaced(self):
        assert safe_name("a:b*c?d[e]") == "a_b_c_d_e_"

    def test_a_plain_name_is_untouched(self):
        assert safe_name("Trial Balance") == "Trial Balance"

    def test_an_empty_name_gets_a_fallback(self):
        assert safe_name("") == "Sheet"
        assert safe_name("   ") == "Sheet"

    def test_non_string_input_is_coerced(self):
        assert safe_name(2024) == "2024"


class TestListExcelFilesInFolder:
    def test_finds_supported_files_recursively(self, tmp_path):
        nested = os.path.join(tmp_path, "sub")
        os.makedirs(nested)
        for name in ("a.xlsx", "b.csv"):
            open(os.path.join(tmp_path, name), "w", encoding="utf-8").close()
        open(os.path.join(nested, "c.xls"), "w", encoding="utf-8").close()

        found = list_excel_files_in_folder(str(tmp_path))
        assert len(found) == 3

    def test_ignores_unsupported_files(self, tmp_path):
        open(os.path.join(tmp_path, "notes.txt"), "w", encoding="utf-8").close()
        assert list_excel_files_in_folder(str(tmp_path)) == []


class TestReadFileToDf:
    def test_reads_a_csv(self, tmp_path):
        p = os.path.join(tmp_path, "a.csv")
        pd.DataFrame({"A": [1, 2]}).to_csv(p, index=False)
        assert list(read_file_to_df(p)["A"]) == ["1", "2"]

    def test_reads_the_first_sheet_of_a_workbook(self, tmp_path):
        p = os.path.join(tmp_path, "a.xlsx")
        with pd.ExcelWriter(p, engine="openpyxl") as writer:
            pd.DataFrame({"First": [1]}).to_excel(writer, sheet_name="One", index=False)
            pd.DataFrame({"Second": [2]}).to_excel(writer, sheet_name="Two", index=False)
        assert list(read_file_to_df(p).columns) == ["First"]

    def test_blanks_are_kept_as_empty_strings(self, tmp_path):
        p = os.path.join(tmp_path, "a.csv")
        with open(p, "w", encoding="utf-8") as fh:
            fh.write("A,B\n1,\n")
        assert read_file_to_df(p).loc[0, "B"] == ""


class TestReadAllSheets:
    def test_returns_every_sheet(self, tmp_path):
        p = os.path.join(tmp_path, "a.xlsx")
        with pd.ExcelWriter(p, engine="openpyxl") as writer:
            pd.DataFrame({"A": [1]}).to_excel(writer, sheet_name="One", index=False)
            pd.DataFrame({"B": [2]}).to_excel(writer, sheet_name="Two", index=False)

        sheets = read_all_sheets(p)
        assert sorted(sheets) == ["One", "Two"]

    def test_a_csv_becomes_a_single_sheet(self, tmp_path):
        p = os.path.join(tmp_path, "a.csv")
        pd.DataFrame({"A": [1]}).to_csv(p, index=False)
        assert list(read_all_sheets(p)) == ["Sheet1"]
