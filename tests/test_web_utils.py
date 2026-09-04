"""
Tests for the Flask bridge helpers in flask_app/utils.py.

The split tests are regression tests: split_files_to_zip previously processed
only file_paths[0] and silently discarded every other upload.
"""

import io
import os
import zipfile

import pandas as pd
import pytest
from openpyxl import load_workbook

import utils


@pytest.fixture
def two_workbooks(tmp_path):
    """Two single-sheet workbooks with distinct sheet names."""
    paths, names = [], []
    for i in ('one', 'two'):
        p = os.path.join(tmp_path, f'{i}.xlsx')
        with pd.ExcelWriter(p, engine='openpyxl') as writer:
            pd.DataFrame({'A': [1]}).to_excel(writer, sheet_name=f'sheet_{i}', index=False)
        paths.append(p)
        names.append(f'{i}.xlsx')
    return paths, names


class TestSplitFilesToZip:

    def test_every_uploaded_file_is_split(self, two_workbooks):
        paths, names = two_workbooks
        buf = utils.split_files_to_zip(paths, names)

        entries = zipfile.ZipFile(buf).namelist()
        assert len(entries) == 2
        assert any('one' in e for e in entries)
        assert any('two' in e for e in entries)

    def test_multi_sheet_file_yields_one_entry_per_sheet(self, tmp_path):
        p = os.path.join(tmp_path, 'multi.xlsx')
        with pd.ExcelWriter(p, engine='openpyxl') as writer:
            pd.DataFrame({'A': [1]}).to_excel(writer, sheet_name='First', index=False)
            pd.DataFrame({'B': [2]}).to_excel(writer, sheet_name='Second', index=False)

        entries = zipfile.ZipFile(utils.split_files_to_zip([p], ['multi.xlsx'])).namelist()
        assert sorted(entries) == ['multi_First.xlsx', 'multi_Second.xlsx']

    def test_entries_are_readable_workbooks(self, two_workbooks):
        paths, names = two_workbooks
        zf = zipfile.ZipFile(utils.split_files_to_zip(paths, names))

        ws = load_workbook(io.BytesIO(zf.read(zf.namelist()[0]))).active
        assert [c.value for c in ws[1]] == ['A']

    def test_colliding_names_are_de_duplicated(self, tmp_path):
        """Two uploads sharing a filename must not overwrite each other."""
        paths = []
        for sub in ('a', 'b'):
            d = os.path.join(tmp_path, sub)
            os.makedirs(d)
            p = os.path.join(d, 'report.xlsx')
            with pd.ExcelWriter(p, engine='openpyxl') as writer:
                pd.DataFrame({'A': [1]}).to_excel(writer, sheet_name='Data', index=False)
            paths.append(p)

        entries = zipfile.ZipFile(
            utils.split_files_to_zip(paths, ['report.xlsx', 'report.xlsx'])
        ).namelist()
        assert len(entries) == len(set(entries)) == 2

    def test_unsafe_sheet_name_is_sanitised(self, tmp_path):
        """A sheet name with a path separator must not create a zip subpath."""
        p = os.path.join(tmp_path, 'odd.csv')
        pd.DataFrame({'A': [1]}).to_csv(p, index=False)

        entries = zipfile.ZipFile(
            utils.split_files_to_zip([p], ['a/b\\c.csv'])
        ).namelist()
        assert all('/' not in e and '\\' not in e for e in entries)


class TestMergeFiles:

    def test_returns_a_workbook_stream(self, two_workbooks, tmp_path):
        paths, _ = two_workbooks
        stream = utils.merge_files(paths)

        ws = load_workbook(stream).active
        assert ws.max_row >= 2

    def test_temp_file_is_cleaned_up(self, two_workbooks, tmp_path):
        paths, _ = two_workbooks
        before = set(os.listdir(os.path.dirname(paths[0])))
        utils.merge_files(paths)
        after = set(os.listdir(os.path.dirname(paths[0])))
        assert before == after


class TestProcessExcelFile:

    def test_applies_transformations_and_names_the_output(self, tmp_path):
        p = os.path.join(tmp_path, 'in.xlsx')
        pd.DataFrame({'Name': ['  alice  ']}).to_excel(p, index=False)

        stream, name = utils.process_excel_file(
            p, 'in.xlsx', {'trim': True, 'text_case': 'upper'}
        )
        assert name == 'in_processed.xlsx'
        assert load_workbook(stream).active['A2'].value == 'ALICE'

    def test_long_sheet_names_are_truncated(self, tmp_path):
        p = os.path.join(tmp_path, 'long.xlsx')
        with pd.ExcelWriter(p, engine='openpyxl') as writer:
            pd.DataFrame({'A': [1]}).to_excel(
                writer, sheet_name='a' * 31, index=False
            )

        stream, _ = utils.process_excel_file(p, 'long.xlsx', {})
        assert all(len(n) <= 31 for n in load_workbook(stream).sheetnames)
