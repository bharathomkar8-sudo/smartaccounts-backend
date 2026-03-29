from flask import Flask, request, send_file
import pandas as pd
import os
import zipfile
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# ---------------- HOME ----------------
@app.route('/')
def home():
    return '''
    <h2>Smart Accounts - Upload Excel</h2>
    <form action="/upload" method="post" enctype="multipart/form-data">
        <input type="file" name="file" required>
        <button type="submit">Upload</button>
    </form>
    '''


# ---------------- UPLOAD ----------------
@app.route('/upload', methods=['POST'])
def upload():
    file = request.files.get('file')

    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)

    xls = pd.ExcelFile(filepath)
    sheets = [s for s in xls.sheet_names if s != "GST Details"]

    html = '<h3>Select Invoice Sheets</h3>'
    html += '<form action="/process" method="post">'
    html += f'<input type="hidden" name="filepath" value="{filepath}">'

    for s in sheets:
        html += f'<input type="checkbox" name="sheets" value="{s}" checked> {s}<br>'

    html += '<button type="submit">Process</button></form>'
    return html


# ---------------- FORMAT ----------------
def create_excel(file_path, rows):

    headers = [
        "Voucher Type","VCH No / Inv No","Description","VCH Date",
        "Order No","Order Date","Other Ref","POS",
        "Party Name","Address","State","Pincode","Party GSTIN",
        "Consignee Name","Con Address","Consignee State","Consignee Pincode","Con GSTIN",
        "Item Name / Code","Is Item Header"
    ]

    wb = Workbook()
    ws = wb.active

    yellow = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    bold = Font(bold=True)

    # HEADER
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = yellow
        c.font = bold

    # DATA
    for r, row in enumerate(rows, 2):
        for c, h in enumerate(headers, 1):
            ws.cell(row=r, column=c, value=row.get(h, ""))

    wb.save(file_path)


# ---------------- PROCESS ----------------
@app.route('/process', methods=['POST'])
def process():

    filepath = request.form.get('filepath')
    selected_sheets = request.form.getlist('sheets')

    xls = pd.ExcelFile(filepath)

    zip_path = os.path.join(UPLOAD_FOLDER, "output.zip")

    if os.path.exists(zip_path):
        os.remove(zip_path)

    with zipfile.ZipFile(zip_path, 'w') as zipf:

        for sheet in selected_sheets:

            df = pd.read_excel(xls, sheet_name=sheet, header=None)

            # HEADER DATA
            vch_type = "Sales E-Invoice"
            vch_no = str(df.iloc[10, 16])
            vch_date = pd.to_datetime(df.iloc[11, 16]).strftime("%d-%m-%Y")
            order_no = df.iloc[19, 1]
            order_date = pd.to_datetime(df.iloc[20, 1]).strftime("%d-%m-%Y")
            other_ref = df.iloc[12, 16]
            pos = df.iloc[14, 5]

            party_name = "M/S. Bharath"
            gstin = str(df.iloc[16, 1])

            # ADDRESS
            addr1 = df.iloc[10, 0]
            addr2 = df.iloc[11, 0]
            addr3 = df.iloc[12, 0]

            # CONSIGNEE ADDRESS (FIXED)
            con1 = df.iloc[12, 4]
            con2 = df.iloc[13, 4]

            rows = []

            # -------- FIRST ROW (ROW 2) --------
            start = 25  # B26

            desc = df.iloc[start, 1]

            # ITEM CODE LOGIC
            try:
                val = df.iloc[start, 2]
                item_code = val if pd.notna(val) and float(val) > 1 else "Header"
            except:
                item_code = "Header"

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

            # -------- ADDRESS CONTINUE --------
            rows.append({"Address": addr2, "Con Address": con2})
            rows.append({"Address": addr3})

            # -------- DESCRIPTION LOOP --------
            end = len(df)

            for i in range(start + 1, len(df)):
                val = str(df.iloc[i, 1]).strip()
                if val.startswith("___"):
                    end = i
                    break

            for i in range(start + 1, end):

                desc = df.iloc[i, 1]
                if pd.isna(desc):
                    desc = ""

                try:
                    val = df.iloc[i, 2]
                    item_code = val if pd.notna(val) and float(val) > 1 else "Header"
                except:
                    item_code = "Header"

                is_header = "Yes" if item_code == "Header" else ""

                rows.append({
                    "Description": desc,
                    "Item Name / Code": item_code,
                    "Is Item Header": is_header
                })

            # SAVE
            file_path = os.path.join(UPLOAD_FOLDER, f"{sheet}.xlsx")
            create_excel(file_path, rows)
            zipf.write(file_path, arcname=f"{sheet}.xlsx")

    return send_file(zip_path, as_attachment=True)


# ---------------- RUN ----------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
