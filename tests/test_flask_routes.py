"""
Tests for the Flask routes: upload validation, temp-file cleanup and the
balance sheet flow.
"""

import io
import os
import time

import app as flask_app_module
import pandas as pd
import pytest


@pytest.fixture
def client(tmp_path):
    """A test client whose upload folder is an empty temp directory."""
    flask_app_module.app.config["TESTING"] = True
    original = flask_app_module.app.config["UPLOAD_FOLDER"]
    flask_app_module.app.config["UPLOAD_FOLDER"] = str(tmp_path)
    with flask_app_module.app.test_client() as c:
        yield c
    flask_app_module.app.config["UPLOAD_FOLDER"] = original


@pytest.fixture
def upload_folder():
    return flask_app_module.app.config["UPLOAD_FOLDER"]


def _xlsx_bytes():
    """A small in-memory .xlsx upload."""
    buf = io.BytesIO()
    pd.DataFrame({"Name": ["  alice  "], "Amount": ["10"]}).to_excel(buf, index=False)
    buf.seek(0)
    return buf


class TestSecretKey:
    def test_a_secret_key_is_configured(self):
        assert flask_app_module.app.secret_key

    def test_the_old_hardcoded_key_is_gone(self):
        assert flask_app_module.app.secret_key != "supersecretkey_change_this_in_prod"

    def test_the_environment_key_is_used(self):
        assert flask_app_module.resolve_secret_key({"SECRET_KEY": "from-env"}, running_directly=False) == "from-env"

    def test_an_imported_app_without_a_key_refuses_to_start(self):
        """gunicorn and flask run import the module; both must supply a key."""
        with pytest.raises(RuntimeError, match="SECRET_KEY"):
            flask_app_module.resolve_secret_key({}, running_directly=False)

    def test_a_direct_run_falls_back_to_the_development_key(self):
        assert flask_app_module.resolve_secret_key({}, running_directly=True) == flask_app_module.DEV_SECRET_KEY


class TestUploadHelpers:
    def test_spreadsheet_extensions_are_accepted(self):
        for name in ("a.xlsx", "B.XLS", "c.csv"):
            assert flask_app_module.allowed_spreadsheet(name)

    def test_other_extensions_are_rejected(self):
        for name in ("a.exe", "b.pdf", "c", ""):
            assert not flask_app_module.allowed_spreadsheet(name)

    def test_discard_upload_ignores_missing_files(self, tmp_path):
        flask_app_module.discard_upload(os.path.join(tmp_path, "not-there.xlsx"), None)

    def test_discard_upload_removes_files(self, tmp_path):
        p = os.path.join(tmp_path, "gone.txt")
        open(p, "w", encoding="utf-8").close()
        flask_app_module.discard_upload(p)
        assert not os.path.exists(p)

    def test_sweep_removes_only_stale_files(self, client, upload_folder):
        # 'client' is requested for its side effect, not its value: it repoints
        # UPLOAD_FOLDER at a temp directory, and fixtures resolve in order, so
        # dropping it would have this test sweep the real uploads folder.
        stale = os.path.join(upload_folder, "stale.xlsx")
        fresh = os.path.join(upload_folder, "fresh.xlsx")
        for p in (stale, fresh):
            open(p, "w", encoding="utf-8").close()
        old = time.time() - flask_app_module.UPLOAD_RETENTION_SECONDS - 60
        os.utime(stale, (old, old))

        flask_app_module.sweep_stale_uploads()

        assert not os.path.exists(stale)
        assert os.path.exists(fresh)


class TestFormatterRoute:
    def test_get_renders(self, client):
        assert client.get("/formatter").status_code == 200

    def test_rejects_an_unsupported_extension(self, client):
        r = client.post(
            "/formatter",
            data={"file": (io.BytesIO(b"not a workbook"), "evil.exe")},
            content_type="multipart/form-data",
        )
        assert r.status_code == 302

    def test_rejected_upload_is_not_written_to_disk(self, client, upload_folder):
        client.post(
            "/formatter",
            data={"file": (io.BytesIO(b"nope"), "evil.exe")},
            content_type="multipart/form-data",
        )
        assert os.listdir(upload_folder) == []

    def test_processes_a_workbook_and_cleans_up(self, client, upload_folder):
        r = client.post(
            "/formatter",
            data={"file": (_xlsx_bytes(), "book.xlsx"), "chk_trim": "on"},
            content_type="multipart/form-data",
        )
        assert r.status_code == 200
        assert r.headers["Content-Disposition"].endswith("book_processed.xlsx")
        assert os.listdir(upload_folder) == [], "upload should be deleted after processing"


class TestMergeRoute:
    def test_get_renders(self, client):
        assert client.get("/merge").status_code == 200

    def test_rejects_an_unsupported_extension(self, client, upload_folder):
        r = client.post(
            "/merge",
            data={"files": [(io.BytesIO(b"x"), "evil.exe")], "operation": "merge"},
            content_type="multipart/form-data",
        )
        assert r.status_code == 302
        assert os.listdir(upload_folder) == []

    def test_merges_and_cleans_up(self, client, upload_folder):
        r = client.post(
            "/merge",
            data={
                "files": [(_xlsx_bytes(), "a.xlsx"), (_xlsx_bytes(), "b.xlsx")],
                "operation": "merge",
                "merged_filename": "combined",
            },
            content_type="multipart/form-data",
        )
        assert r.status_code == 200
        assert r.headers["Content-Disposition"].endswith("combined.xlsx")
        assert os.listdir(upload_folder) == []

    def test_merged_filename_cannot_escape_with_a_path(self, client):
        r = client.post(
            "/merge",
            data={
                "files": [(_xlsx_bytes(), "a.xlsx")],
                "operation": "merge",
                "merged_filename": "../../etc/passwd",
            },
            content_type="multipart/form-data",
        )
        disposition = r.headers["Content-Disposition"]
        assert ".." not in disposition
        assert "/" not in disposition.split("filename=")[-1]

    def test_split_returns_a_zip(self, client, upload_folder):
        r = client.post(
            "/merge",
            data={
                "files": [(_xlsx_bytes(), "a.xlsx"), (_xlsx_bytes(), "b.xlsx")],
                "operation": "split",
            },
            content_type="multipart/form-data",
        )
        assert r.status_code == 200
        assert r.mimetype == "application/zip"
        assert os.listdir(upload_folder) == []


class TestPdfRoutes:
    def test_home_renders(self, client):
        assert client.get("/").status_code == 200

    def test_rejects_a_non_pdf(self, client, upload_folder):
        r = client.post(
            "/",
            data={"file": (io.BytesIO(b"x"), "sheet.xlsx")},
            content_type="multipart/form-data",
        )
        assert r.status_code == 302
        assert os.listdir(upload_folder) == []

    def test_convert_without_a_session_redirects(self, client):
        r = client.post("/convert", data={"selected_tables": ["1-1"]})
        assert r.status_code == 302


class TestBalanceSheetRoutes:
    def test_index_renders(self, client):
        assert client.get("/balance-sheet").status_code == 200

    def test_posting_a_page_advances_to_the_next(self, client):
        r = client.post("/balance-sheet/page1", data={"field": "value"})
        assert r.status_code == 302
        assert "/balance-sheet/page2" in r.headers["Location"]

    def test_repeated_keys_are_kept_as_lists(self, client):
        client.post(
            "/balance-sheet/page1",
            data={"sch3_particulars[]": ["Rent", "Salary"], "sch3_amount[]": ["100", "200"]},
        )
        with client.session_transaction() as s:
            assert s["bs_page1"]["sch3_particulars[]"] == ["Rent", "Salary"]

    def test_generate_returns_a_workbook(self, client):
        client.post("/balance-sheet/page1", data={"trust_name": "Test Trust"})
        r = client.get("/balance-sheet/generate")
        assert r.status_code == 200
        assert r.headers["Content-Disposition"].endswith("Balance_Sheet.xlsx")
