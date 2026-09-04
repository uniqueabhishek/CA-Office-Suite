"""
Main Flask Application for CA Office Suite.
This module processes PDF to Excel conversion, Excel formatting,
merging/splitting, and hosts the Balance Sheet generator.
"""
import os
import sys
import uuid
import traceback

# The shared core/ and config/ packages live at the repository root, one level
# above this file. Put the root on sys.path before importing anything that
# depends on them (utils, blueprints).
_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _REPO_ROOT not in sys.path:
    sys.path.insert(0, _REPO_ROOT)

from flask import Flask, render_template, request, send_file, flash, redirect, url_for, session  # noqa: E402
from werkzeug.utils import secure_filename  # noqa: E402

import utils  # noqa: E402
from blueprints.balance_sheet import balance_sheet_bp  # noqa: E402


app = Flask(__name__)
app.secret_key = 'supersecretkey_change_this_in_prod'
app.register_blueprint(balance_sheet_bp)
app.config['UPLOAD_FOLDER'] = os.path.join(os.getcwd(), 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max limit

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)


# ----------------- PDF TO EXCEL -----------------


@app.route('/', methods=['GET', 'POST'])
def index():
    """
    Home page (PDF to Excel).
    - GET: Renders the upload form.
    - POST: Handles PDF upload and table extraction.
    """
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and file.filename.lower().endswith('.pdf'):
            filename = secure_filename(file.filename)
            unique_filename = f"{uuid.uuid4()}_{filename}"
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)

            print(f"DEBUG: Processing upload for {filename}")
            try:
                print(f"DEBUG: Saving to {file_path}")
                file.save(file_path)

                print("DEBUG: Extracting tables...")
                tables = utils.extract_tables_from_pdf(file_path)
                print(f"DEBUG: Extraction complete. Found {len(tables) if tables else 0} tables.")

                if not tables:
                    flash('No tables found in the PDF. Is it a scanned image?')
                    return redirect(request.url)

                session['current_pdf'] = unique_filename
                session['original_filename'] = filename
                print("DEBUG: Rendering template.")
                return render_template(
                    'select_tables.html', tables=tables, filename=filename, active_tab='pdf'
                )
            except Exception as e:  # pylint: disable=broad-except
                # Broad exception catch is intentional for top-level error handling
                # We want to catch everything during processing to avoid crashing the server
                # and show a user-friendly flash message instead.
                print(f"ERROR: {str(e)}")
                traceback.print_exc()
                flash(f'Error processing PDF: {str(e)}')
                return redirect(request.url)
        else:
            flash('Invalid file type.')
            return redirect(request.url)

    # Clear session on GET to ensure fresh UI state
    if request.method == 'GET':
        session.pop('current_pdf', None)
        session.pop('original_filename', None)

    return render_template('index.html', active_tab='pdf')


@app.route('/convert', methods=['POST'])
def convert():
    """
    Converts selected PDF tables to Excel.
    """
    current_pdf = session.get('current_pdf')
    original_filename = session.get('original_filename')
    if not current_pdf:
        flash('Session expired.')
        return redirect(url_for('index'))
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], current_pdf)
    selected_tables = request.form.getlist('selected_tables')
    if not selected_tables:
        flash('Please select at least one table.')
        return redirect(url_for('index'))
    try:
        output = utils.convert_selected_tables_to_excel(file_path, selected_tables)
        base_name = os.path.splitext(original_filename)[0]
        output_filename = f"{base_name}_converted.xlsx"
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=output_filename
        )
    except Exception as e:  # pylint: disable=broad-except
        flash(f'Error converting file: {str(e)}')
        return redirect(url_for('index'))


# ----------------- EXCEL FORMATTER -----------------


@app.route('/formatter', methods=['GET', 'POST'])
def formatter():
    """
    Excel Formatter tool.
    Uploads an Excel file and applies formatting options.
    """
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file uploaded.')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No file selected.')
            return redirect(request.url)

        # Save temp file
        filename = secure_filename(file.filename)
        temp_path = os.path.join(
            app.config['UPLOAD_FOLDER'], f"fmt_{uuid.uuid4()}_{filename}"
        )
        file.save(temp_path)

        # Collect Options
        options = {
            'numbers': 'chk_numbers' in request.form,
            'trim': 'chk_trim' in request.form,
            'dates': 'chk_dates' in request.form,
            'date_format': request.form.get('date_format', 'dd-mm-yyyy'),
            'number_format': 'chk_number_format' in request.form,
            'number_format_option': request.form.get('number_format_option', '2_decimals'),
            'currency_symbol': request.form.get('currency_symbol', '₹'),
            'text_case': (
                request.form.get('text_case_option', 'none')
                if 'chk_text_case' in request.form else 'none'
            ),
            'remove_dups': 'chk_remove_dups' in request.form,
            'apply_autofit': 'chk_autofit' in request.form,
            'apply_theme': 'chk_theme' in request.form
        }

        try:
            # Process using Utils bridge
            output_stream, out_name = utils.process_excel_file(temp_path, filename, options)
            return send_file(
                output_stream,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name=out_name
            )
        except Exception as e:  # pylint: disable=broad-except
            flash(f"Error processing file: {e}")
            return redirect(request.url)

    return render_template('formatter.html', active_tab='formatter')


# ----------------- MERGE & SPLIT -----------------


@app.route('/merge', methods=['GET', 'POST'])
def merge():
    """
    Excel Merge & Split tool.
    Handles multiple file uploads for merging or splitting operations.
    """
    if request.method == 'POST':
        files = request.files.getlist('files')
        if not files or files[0].filename == '':
            flash('No files selected.')
            return redirect(request.url)

        operation = request.form.get('operation', 'merge')

        # Save all files
        saved_paths = []
        original_names = []
        for f in files:
            fname = secure_filename(f.filename)
            path = os.path.join(app.config['UPLOAD_FOLDER'], f"merge_{uuid.uuid4()}_{fname}")
            f.save(path)
            saved_paths.append(path)
            original_names.append(fname)

        try:
            if operation == 'merge':
                merged_name = request.form.get('merged_filename', 'merged_output.xlsx')
                if not merged_name.endswith('.xlsx'):
                    merged_name += '.xlsx'

                output_stream = utils.merge_files(saved_paths)
                return send_file(
                    output_stream,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    as_attachment=True,
                    download_name=merged_name
                )
            else:
                # Split - for web we probably zip the results?
                # For simplicity, let's just split the FIRST file and return a zip, or warn user.
                # Implementing split for multiple files in web is complex due to download limit
                # (one response). We'll use a Zip file.
                zip_stream = utils.split_files_to_zip(saved_paths, original_names)
                return send_file(
                    zip_stream,
                    mimetype='application/zip',
                    as_attachment=True,
                    download_name='split_files.zip'
                )
        except Exception as e:  # pylint: disable=broad-except
            flash(f"Error processing: {e}")
            return redirect(request.url)

    return render_template('merge.html', active_tab='merge')


if __name__ == '__main__':
    app.run(debug=True)
