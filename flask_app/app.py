from flask import Flask, render_template, request, send_file, flash, redirect, url_for, session
from werkzeug.utils import secure_filename
import os
import utils
import uuid

app = Flask(__name__)
app.secret_key = 'supersecretkey_change_this_in_prod'
app.config['UPLOAD_FOLDER'] = os.path.join(os.getcwd(), 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max limit

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)

        file = request.files['file']

        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)

        if file and file.filename.lower().endswith('.pdf'):
            # Save file with unique name to avoid conflicts
            filename = secure_filename(file.filename)
            unique_filename = f"{uuid.uuid4()}_{filename}"
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
            file.save(file_path)

            # Process PDF
            try:
                tables = utils.extract_tables_from_pdf(file_path)

                if not tables:
                    flash('No tables found in the PDF.')
                    return redirect(request.url)

                # Store filename in session to use in next step
                session['current_pdf'] = unique_filename
                session['original_filename'] = filename

                # We can't store the whole dataframe in session (too big),
                # passing the HTML representation to the template is fine.
                return render_template('select_tables.html', tables=tables, filename=filename)

            except Exception as e:
                flash(f'Error processing PDF: {str(e)}')
                return redirect(request.url)

        else:
            flash('Invalid file type. Please upload a PDF.')
            return redirect(request.url)

    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    current_pdf = session.get('current_pdf')
    original_filename = session.get('original_filename')

    if not current_pdf:
        flash('Session expired. Please upload the file again.')
        return redirect(url_for('index'))

    file_path = os.path.join(app.config['UPLOAD_FOLDER'], current_pdf)

    # Get selected tables from form
    # The checkboxes will have names like "table_1-1", "table_1-2" etc.
    selected_tables = request.form.getlist('selected_tables')

    if not selected_tables:
        flash('Please select at least one table.')
        # Ideally we would re-render the select page, but we'd need to re-extract.
        # For simplicity, redirect to index or re-extract here.
        # Let's redirect to index for now to force a fresh flow or handle gracefully.
        return redirect(url_for('index'))

    try:
        output = utils.convert_selected_tables_to_excel(file_path, selected_tables)

        # Create output filename
        base_name = os.path.splitext(original_filename)[0]
        output_filename = f"{base_name}_converted.xlsx"

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=output_filename
        )

    except Exception as e:
        flash(f'Error converting file: {str(e)}')
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
