from flask import Flask, request, render_template_string, send_file
import pandas as pd
import os
import zipfile

app = Flask(__name__)

# Store file temporarily
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# 🔹 Home page
@app.route('/')
def home():
    return '''
    <h2>Upload Sales Excel</h2>
    <form action="/upload" method="post" enctype="multipart/form-data">
        <input type="file" name="file">
        <button type="submit">Upload</button>
    </form>
    '''


# 🔹 Step 1: Upload & read sheets
@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']

    if not file:
        return "No file uploaded"

    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)

    # Read sheet names
    xls = pd.ExcelFile(filepath)
    sheets = xls.sheet_names

    # Show checkboxes
    html = '<h3>Select Sheets (Invoices)</h3>'
    html += '<form action="/process" method="post">'

    html += f'<input type="hidden" name="filepath" value="{filepath}">'

    for sheet in sheets:
        html += f'<input type="checkbox" name="sheets" value="{sheet}">{sheet}<br>'

    html += '<button type="submit">Process</button>'
    html += '</form>'

    return html


# 🔹 Step 2: Process selected sheets
@app.route('/process', methods=['POST'])
def process():
    filepath = request.form['filepath']
    selected_sheets = request.form.getlist('sheets')

    if not selected_sheets:
        return "No sheets selected"

    zip_path = os.path.join(UPLOAD_FOLDER, "output.zip")

    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for sheet in selected_sheets:
            df = pd.read_excel(filepath, sheet_name=sheet)

            output_file = os.path.join(UPLOAD_FOLDER, f"{sheet}.xlsx")
            df.to_excel(output_file, index=False)

            zipf.write(output_file, arcname=f"{sheet}.xlsx")

    return send_file(zip_path, as_attachment=True)


# 🔹 Run app
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
