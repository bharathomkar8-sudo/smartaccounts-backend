from flask import Flask, request, send_file
import pandas as pd
import os
import zipfile
from openpyxl import Workbook

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# ---------- SAFE READ ----------
def safe(df, r, c):
    try:
        return df.iloc[r, c]
    except:
        return ""


def safe_date(val):
    try:
        return pd.to_datetime(val).strftime("%d-%m-%Y")
    except:
        return ""


def safe_float(val):
    try:
        return float(val)
    except:
        return None


# ---------- HOME ----------
@app.route('/')
def home():
    return '''
    <h2>Upload Excel</h2>
    <form action="/upload" method="post" enctype="multipart/form-data">
        <input type="file" name="file" required>
        <button type="submit">Upload</button>
    </form>
    '''


# ---------- UPLOAD ----------
@app.route('/upload', methods=['POST'])
def upload():
    file = request.files.get('file')

    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)

    xls = pd.ExcelFile(filepath)

    html = '<form action="/process" method="post">'
    html += f'<input type="hidden" name="filepath" value="{filepath}">'

    for s in xls.sheet_names:
        html += f'<input type="checkbox" name="sheets" value="{s}" checked> {s}<br>'

    html += '<button type="submit">Process</button></form>'
    return html


# ---------- CREATE EXCEL ----------
def create_excel(file_path, rows):

    headers = [
        "Voucher Type","VCH No / Inv No","Description","VCH Date",
        "Order No","Order Date","Other Ref","POS",
        "Party Name","Address","Party GSTIN",
        "Consignee Name","Con Address",
        "Item Name / Code","Is Item Header"
    ]

    wb = Workbook()
    ws = wb.active

    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)

    for r, row in enumerate(rows, 2):
        for c, h in enumerate(headers, 1):
            ws.cell(row=r, column=c, value=row.get(h, ""))

    wb.save(file_path)


# ---------- PROCESS ----------
@app.route('/process', methods=['POST'])
def process():

    filepath = request.form.get('filepath')
    sheets = request.form.getlist('sheets')

    xls = pd.ExcelFile(filepath)

    zip_path = os.path.join(UPLOAD_FOLDER, "output.zip")

    if os.path.exists(zip_path):
        os.remove(zip_path)

    with zipfile.ZipFile(zip_path, 'w') as zipf:

        for sheet in sheets:

            try:
                df = pd.read_excel(xls, sheet_name=sheet, header=None)

                # HEADER
                vch_type = "Sales E-Invoice"
                vch_no = str(safe(df, 10, 16))
                vch_date = safe_date(safe(df, 11, 16))
                order_no = safe(df, 19, 1)
                order_date = safe_date(safe(df, 20, 1))
                other_ref = safe(df, 12, 16)
                pos = safe(df, 14, 5)

                party_name = "M/S. Bharath"
                gstin = str(safe(df, 16, 1))

                # ADDRESS
                addr1 = safe(df, 10, 0)
                addr2 = safe(df, 11, 0)
                addr3 = safe(df, 12, 0)

                # CONSIGNEE
                con1 = safe(df, 12, 4)
                con2 = safe(df, 13, 4)

                rows = []

                # -------- FIRST ROW --------
                start = 25

                desc = safe(df, start, 1)

                val = safe_float(safe(df, start, 2))
                item_code = val if val and val > 1 else "Header"
                is_header = "Yes" if item_code == "Header" else ""

                rows.append({
                    "Voucher Type": vch_type,
                    "VCH No / Inv No": vch_no,
                    "Description": desc,
                    "VCH Date": vch_date,
                    "Order No": order_no,
                    "Order Date": order_date,
                    "Other Ref": other_ref,
                    "POS": pos,
                    "Party Name": party_name,
                    "Address": addr1,
                    "Party GSTIN": gstin,
                    "Consignee Name": party_name,
                    "Con Address": con1,
                    "Item Name / Code": item_code,
                    "Is Item Header": is_header
                })

                rows.append({"Address": addr2, "Con Address": con2})
                rows.append({"Address": addr3})

                # -------- LOOP --------
                for i in range(start + 1, len(df)):

                    val_stop = str(safe(df, i, 1)).strip()
                    if val_stop.startswith("___"):
                        break

                    desc = safe(df, i, 1)

                    val = safe_float(safe(df, i, 2))
                    item_code = val if val and val > 1 else "Header"
                    is_header = "Yes" if item_code == "Header" else ""

                    rows.append({
                        "Description": desc,
                        "Item Name / Code": item_code,
                        "Is Item Header": is_header
                    })

                # SAVE
                path = os.path.join(UPLOAD_FOLDER, f"{sheet}.xlsx")
                create_excel(path, rows)
                zipf.write(path, arcname=f"{sheet}.xlsx")

            except Exception as e:
                print("ERROR:", sheet, e)

    return send_file(zip_path, as_attachment=True)


# ---------- RUN ----------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
