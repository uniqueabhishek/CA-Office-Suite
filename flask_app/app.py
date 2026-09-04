"""
Main Flask Application for CA Office Suite.
This module processes PDF to Excel conversion, Excel formatting,
merging/splitting, and hosts the Balance Sheet generator.
"""

import contextlib
import logging
import os
import sys
import time
import uuid

# The shared core/ and config/ packages live at the repository root, one level
# above this file. Put the root on sys.path before importing anything that
# depends on them (utils, blueprints).
_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

# These imports must follow the sys.path bootstrap above, which ruff reports
# as E402 unless each line says otherwise.
from flask import Flask, flash, redirect, render_template, request, send_file, session, url_for  # noqa: E402
from werkzeug.utils import secure_filename  # noqa: E402

import utils  # noqa: E402
from blueprints.balance_sheet import balance_sheet_bp  # noqa: E402
from config.constants import ALLOWED_UPLOAD_EXTENSIONS, UPLOAD_RETENTION_SECONDS  # noqa: E402

logger = logging.getLogger(__name__)

DEV_SECRET_KEY = "dev-only-insecure-key"


def resolve_secret_key(env, running_directly):
    """
    Return the session signing key, or raise if the deployment has not set one.

    Flask signs the session cookie with this key, and this app trusts the
    session to name a file on disk, so a known key is a real hole. The only
    context that may fall back to a throwaway key is running this file directly
    during development; anything that imports the module (gunicorn, flask run)
    must supply SECRET_KEY. Defaulting the other way is what let the previous
    FLASK_ENV check pass silently in production.

    Args:
        env (dict): Environment mapping, normally os.environ
        running_directly (bool): True when this module is __main__

    Returns:
        str: The signing key

    Raises:
        RuntimeError: No SECRET_KEY set outside of a direct development run
    """
    key = env.get("SECRET_KEY")
    if key:
        return key
    if not running_directly:
        raise RuntimeError(
            "SECRET_KEY environment variable must be set. It is optional only "
            "when running app.py directly for local development."
        )
    logger.warning("SECRET_KEY not set. Using an insecure development key.")
    return DEV_SECRET_KEY


app = Flask(__name__)
app.secret_key = resolve_secret_key(os.environ, __name__ == "__main__")

app.register_blueprint(balance_sheet_bp)
# Anchored to this file, not the working directory, so uploads land in the same
# place whether the app is started from the repo root or from flask_app/.
app.config["UPLOAD_FOLDER"] = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16MB max limit

# Ensure upload folder exists
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)


# ----------------- UPLOAD HELPERS -----------------


def allowed_spreadsheet(filename):
    """Returns True if the filename has a supported spreadsheet extension."""
    return bool(filename) and filename.lower().endswith(ALLOWED_UPLOAD_EXTENSIONS)


def discard_upload(*paths):
    """
    Delete temporary uploads, ignoring files that are already gone.

    Uploads are working files only; leaving them behind grows the upload
    folder without bound.
    """
    for path in paths:
        if not path:
            continue
        with contextlib.suppress(OSError):
            os.remove(path)


def sweep_stale_uploads():
    """
    Remove uploads older than the retention window.

    The PDF flow keeps a file between the upload and the convert request, so it
    cannot be deleted eagerly. This sweep collects anything abandoned partway
    through (session expired, browser closed) on the next upload.
    """
    folder = app.config["UPLOAD_FOLDER"]
    cutoff = time.time() - UPLOAD_RETENTION_SECONDS
    try:
        entries = os.listdir(folder)
    except OSError:
        return
    for name in entries:
        path = os.path.join(folder, name)
        try:
            if os.path.isfile(path) and os.path.getmtime(path) < cutoff:
                os.remove(path)
        except OSError:
            pass


# ----------------- PDF TO EXCEL -----------------


@app.route("/", methods=["GET", "POST"])
def index():  # noqa: PLR0911 - one early return per validation failure reads better than nesting
    """
    Home page (PDF to Excel).
    - GET: Renders the upload form.
    - POST: Handles PDF upload and table extraction.
    """
    if request.method == "POST":
        if "file" not in request.files:
            flash("No file part")
            return redirect(request.url)
        file = request.files["file"]
        # A multipart part can arrive with no filename at all, which is None
        # rather than "" - the old check let that through and then crashed on
        # .lower(). Normalise once so everything below deals with a str.
        upload_name = file.filename or ""
        if not upload_name:
            flash("No selected file")
            return redirect(request.url)
        if upload_name.lower().endswith(".pdf"):
            sweep_stale_uploads()
            filename = secure_filename(upload_name)
            unique_filename = f"{uuid.uuid4()}_{filename}"
            file_path = os.path.join(app.config["UPLOAD_FOLDER"], unique_filename)

            try:
                file.save(file_path)
                tables = utils.extract_tables_from_pdf(file_path)

                if not tables:
                    discard_upload(file_path)
                    flash("No tables found in the PDF. Is it a scanned image?")
                    return redirect(request.url)

                # Kept on disk so /convert can re-read the selected tables.
                session["current_pdf"] = unique_filename
                session["original_filename"] = filename
                return render_template("select_tables.html", tables=tables, filename=filename, active_tab="pdf")
            except Exception as e:
                # Broad exception catch is intentional for top-level error handling
                # We want to catch everything during processing to avoid crashing the server
                # and show a user-friendly flash message instead.
                discard_upload(file_path)
                logger.exception("PDF processing failed for %s", filename)
                flash(f"Error processing PDF: {e!s}")
                return redirect(request.url)
        else:
            flash("Invalid file type.")
            return redirect(request.url)

    # Clear session on GET to ensure fresh UI state
    if request.method == "GET":
        session.pop("current_pdf", None)
        session.pop("original_filename", None)

    return render_template("index.html", active_tab="pdf")


@app.route("/convert", methods=["POST"])
def convert():
    """
    Converts selected PDF tables to Excel.
    """
    current_pdf = session.get("current_pdf")
    original_filename = session.get("original_filename")
    # Both are written together, but a half-populated cookie would otherwise
    # reach os.path.splitext(None) below.
    if not current_pdf or not original_filename:
        flash("Session expired.")
        return redirect(url_for("index"))
    file_path = os.path.join(app.config["UPLOAD_FOLDER"], current_pdf)
    selected_tables = request.form.getlist("selected_tables")
    if not selected_tables:
        flash("Please select at least one table.")
        return redirect(url_for("index"))
    try:
        output = utils.convert_selected_tables_to_excel(file_path, selected_tables)
        base_name = os.path.splitext(original_filename)[0]
        output_filename = f"{base_name}_converted.xlsx"
        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=output_filename,
        )
    except Exception as e:
        logger.exception("Table conversion failed for %s", current_pdf)
        flash(f"Error converting file: {e!s}")
        return redirect(url_for("index"))
    finally:
        # The upload has served its purpose once the workbook is built.
        discard_upload(file_path)
        session.pop("current_pdf", None)
        session.pop("original_filename", None)


# ----------------- EXCEL FORMATTER -----------------


@app.route("/formatter", methods=["GET", "POST"])
def formatter():
    """
    Excel Formatter tool.
    Uploads an Excel file and applies formatting options.
    """
    if request.method == "POST":
        if "file" not in request.files:
            flash("No file uploaded.")
            return redirect(request.url)
        file = request.files["file"]
        upload_name = file.filename or ""
        if not upload_name:
            flash("No file selected.")
            return redirect(request.url)
        if not allowed_spreadsheet(upload_name):
            flash("Invalid file type. Please upload an .xlsx, .xls or .csv file.")
            return redirect(request.url)

        # Save temp file
        sweep_stale_uploads()
        filename = secure_filename(upload_name)
        temp_path = os.path.join(app.config["UPLOAD_FOLDER"], f"fmt_{uuid.uuid4()}_{filename}")
        file.save(temp_path)

        # Collect Options
        options = {
            "numbers": "chk_numbers" in request.form,
            "trim": "chk_trim" in request.form,
            "dates": "chk_dates" in request.form,
            "date_format": request.form.get("date_format", "dd-mm-yyyy"),
            "number_format": "chk_number_format" in request.form,
            "number_format_option": request.form.get("number_format_option", "2_decimals"),
            "currency_symbol": request.form.get("currency_symbol", "₹"),
            "text_case": (request.form.get("text_case_option", "none") if "chk_text_case" in request.form else "none"),
            "remove_dups": "chk_remove_dups" in request.form,
            "apply_autofit": "chk_autofit" in request.form,
            "apply_theme": "chk_theme" in request.form,
        }

        try:
            # Process using Utils bridge
            output_stream, out_name = utils.process_excel_file(temp_path, filename, options)
            return send_file(
                output_stream,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name=out_name,
            )
        except Exception as e:
            logger.exception("Formatting failed for %s", filename)
            flash(f"Error processing file: {e}")
            return redirect(request.url)
        finally:
            discard_upload(temp_path)

    return render_template("formatter.html", active_tab="formatter")


# ----------------- MERGE & SPLIT -----------------


@app.route("/merge", methods=["GET", "POST"])
def merge():
    """
    Excel Merge & Split tool.
    Handles multiple file uploads for merging or splitting operations.
    """
    if request.method == "POST":
        files = request.files.getlist("files")
        # A part with no filename gives None, which would end up in `rejected`
        # and blow up the join below. Normalise before anything reads them.
        upload_names = [f.filename or "" for f in files]
        if not files or not upload_names[0]:
            flash("No files selected.")
            return redirect(request.url)

        rejected = [name for name in upload_names if not allowed_spreadsheet(name)]
        if rejected:
            flash(f"Unsupported file type: {', '.join(rejected)}. Use .xlsx, .xls or .csv.")
            return redirect(request.url)

        operation = request.form.get("operation", "merge")

        # Save all files
        sweep_stale_uploads()
        saved_paths = []
        original_names = []
        for f, upload_name in zip(files, upload_names, strict=True):
            fname = secure_filename(upload_name)
            path = os.path.join(app.config["UPLOAD_FOLDER"], f"merge_{uuid.uuid4()}_{fname}")
            f.save(path)
            saved_paths.append(path)
            original_names.append(fname)

        try:
            if operation == "merge":
                # secure_filename strips any path components a user might submit
                merged_name = secure_filename(request.form.get("merged_filename", "") or "merged_output.xlsx")
                if not merged_name.lower().endswith(".xlsx"):
                    merged_name += ".xlsx"

                output_stream = utils.merge_files(saved_paths)
                return send_file(
                    output_stream,
                    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    as_attachment=True,
                    download_name=merged_name,
                )

            # Split every uploaded file; a single response can only carry one
            # file, so the resulting sheets are returned as a zip archive.
            zip_stream = utils.split_files_to_zip(saved_paths, original_names)
            return send_file(
                zip_stream, mimetype="application/zip", as_attachment=True, download_name="split_files.zip"
            )
        except Exception as e:
            logger.exception("%s failed for %s", operation, original_names)
            flash(f"Error processing: {e}")
            return redirect(request.url)
        finally:
            discard_upload(*saved_paths)

    return render_template("merge.html", active_tab="merge")


if __name__ == "__main__":
    # Under gunicorn the server owns logging configuration, so only the direct
    # development run sets it up here.
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    app.run(debug=os.environ.get("FLASK_DEBUG", "1") == "1")
