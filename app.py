from flask import Flask, request, send_file
import pandas as pd
import os
import zipfile
import uuid

# ✅ VERY IMPORTANT (FIRST LINE AFTER IMPORTS)
app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


@app.route('/')
def home():
    return '''
    <h2>Upload Sales Excel</h2>
    <form action="/upload" method="post" enctype="multipart/form-data">
        <input type="file" name="file">
        <button type="submit">Upload</button>
    </form>
    '''


@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']

    if not file:
        return "No file uploaded"

    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)

    xls = pd.ExcelFile(filepath)
    sheets = [s for s in xls.sheet_names if s != "GST Details"]

    html = '<h3>Select Invoice Sheets</h3>'
    html += '<form action="/process" method="post">'
    html += f'<input type="hidden" name="filepath" value="{filepath}">'

    for s in sheets:
        html += f'<input type="checkbox" name="sheets" value="{s}">{s}<br>'

    html += '<button type="submit">Process</button>'
    html += '</form>'

    return html


@app.route('/process', methods=['POST'])
def process():

    filepath = request.form['filepath']
    selected_sheets = request.form.getlist('sheets')

    if not selected_sheets:
        return "No sheets selected"

    xls = pd.ExcelFile(filepath)

    # GST master
    try:
        gst_df = pd.read_excel(xls, sheet_name="GST Details")
        gst_df.iloc[:,0] = gst_df.iloc[:,0].astype(str).str.strip().str.upper()
    except:
        gst_df = None

    zip_path = os.path.join(UPLOAD_FOLDER, "output.zip")

    if os.path.exists(zip_path):
        os.remove(zip_path)

    with zipfile.ZipFile(zip_path, 'w') as zipf:

        for sheet in selected_sheets:
            try:
                df = pd.read_excel(xls, sheet_name=sheet, header=None)

                # safe header
                vch_no = str(df.iloc[10,16]) if df.shape[0] > 10 else sheet
                gstin = str(df.iloc[16,1]).strip().upper() if df.shape[0] > 16 else ""

                party_name = "UNKNOWN"
                if gst_df is not None and gstin:
                    match = gst_df[gst_df.iloc[:,0] == gstin]
                    if not match.empty:
                        party_name = match.iloc[0,1]

                rows = [{
                    "VCH No": vch_no,
                    "GSTIN": gstin,
                    "Party": party_name
                }]

                out_df = pd.DataFrame(rows)

                file_name = f"{sheet}_{str(uuid.uuid4())[:6]}.xlsx"
                file_path = os.path.join(UPLOAD_FOLDER, file_name)

                out_df.to_excel(file_path, index=False)

                zipf.write(file_path, arcname=file_name)

            except Exception as e:
                print("Error:", e)
                continue

    return send_file(zip_path, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
