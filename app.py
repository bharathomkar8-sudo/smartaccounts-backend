from flask import Flask, request, send_file
import pandas as pd
import os
import zipfile
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

# ✅ MUST BE FIRST
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
        html += f'<input type="checkbox" name="sheets" value="{s}" checked> {s}<br>'

    html += '<br><button type="submit">Process</button>'
    html += '</form>'

    return html


# ---------------- FORMAT FUNCTION ----------------
def create_formatted_excel(file_path, rows):

    headers = [
        "Voucher Type","VCH No / Inv No","Description","VCH Date",
        "Order No","Order Date","Other Ref","POS",
        "Party Name","Address","State","Pincode","Party GSTIN",
        "Consignee Name","Con Address","Consignee State","Consignee Pincode","Con GSTIN",
        "Item Name / Code","Is Item Header","width","Height","Qty","Extraudf",
        "Billedqty","Rate","Taxable Value","Dis%","Amount",
        "Sales Ledger","CGST Ledger","CGST Amt","SGST Ledger","SGST Amount",
        "IGST Ledger","IGST Amount","Round off","Invoice Amt","Item header"
    ]

    wb = Workbook()
    ws = wb.active

    yellow = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    bold = Font(bold=True)

    # Header
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = yellow
        cell.font = bold

    # Data
    for r, row in enumerate(rows, 2):
        for c, h in enumerate(headers, 1):
            ws.cell(row=r, column=c, value=row.get(h, ""))

    wb.save(file_path)


# ---------------- PROCESS ----------------
@app.route('/process', methods=['POST'])
def process():

    filepath = request.form.get('filepath')
    selected_sheets = request.form.getlist('sheets')

    if not filepath or not selected_sheets:
        return "Missing input"

    xls = pd.ExcelFile(filepath)

    # GST Master
    try:
        gst_df = pd.read_excel(xls, sheet_name="GST Details")
        gst_df.iloc[:, 0] = gst_df.iloc[:, 0].astype(str).str.strip().str.upper()
    except:
        gst_df = None

    zip_path = os.path.join(UPLOAD_FOLDER, "output.zip")

    if os.path.exists(zip_path):
        os.remove(zip_path)

    with zipfile.ZipFile(zip_path, 'w') as zipf:

        for sheet in selected_sheets:
            try:
                print("Processing:", sheet)

                df = pd.read_excel(xls, sheet_name=sheet, header=None)

                # -------- HEADER --------
                vch_type = "Sales E-Invoice"

                vch_no = str(df.iloc[10, 16]) if df.shape[0] > 10 else sheet
                vch_date = df.iloc[11, 16] if df.shape[0] > 11 else ""
                order_no = df.iloc[19, 1] if df.shape[0] > 19 else ""
                order_date = df.iloc[20, 1] if df.shape[0] > 20 else ""
                other_ref = df.iloc[12, 16] if df.shape[0] > 12 else ""
                pos = df.iloc[14, 5] if df.shape[0] > 14 else ""

                # GST → Party
                gstin = str(df.iloc[16, 1]).strip().upper() if df.shape[0] > 16 else ""
                party_name = "UNKNOWN"

                if gst_df is not None and gstin:
                    match = gst_df[gst_df.iloc[:, 0] == gstin]
                    if not match.empty:
                        party_name = match.iloc[0, 1]

                # Address
                try:
                    address = f"{df.iloc[11,0]} {df.iloc[12,0]} {df.iloc[13,0]}"
                except:
                    address = ""

                state = df.iloc[14, 1] if df.shape[0] > 14 else ""
                pincode = df.iloc[15, 1] if df.shape[0] > 15 else ""

                # -------- ITEMS --------
                rows = []
                start_row = 25

                try:
                    end_row = df[df.apply(
                        lambda r: r.astype(str).str.contains("GST Break", case=False).any(),
                        axis=1
                    )].index[0]
                except:
                    end_row = len(df)

                for i in range(start_row, end_row):
                    try:
                        item = df.iloc[i, 1]
                        qty = df.iloc[i, 5]
                        rate = df.iloc[i, 8]
                        amount = df.iloc[i, 10]
                    except:
                        continue

                    if pd.notna(qty) and qty != 0:
                        rows.append({
                            "Voucher Type": vch_type,
                            "VCH No / Inv No": vch_no,
                            "Description": item,
                            "VCH Date": vch_date,
                            "Order No": order_no,
                            "Order Date": order_date,
                            "Other Ref": other_ref,
                            "POS": pos,
                            "Party Name": party_name,
                            "Address": address,
                            "State": state,
                            "Pincode": pincode,
                            "Party GSTIN": gstin,
                            "Item Name / Code": item,
                            "Qty": qty,
                            "Billedqty": qty,
                            "Rate": rate,
                            "Taxable Value": amount,
                            "Amount": amount,
                            "Sales Ledger": "Sales",
                            "Item header": item
                        })

                # fallback
                if len(rows) == 0:
                    rows.append({
                        "Voucher Type": vch_type,
                        "VCH No / Inv No": vch_no,
                        "Party Name": party_name,
                        "Item header": "No items"
                    })

                # ✅ UNIQUE FILE (no overwrite issue)
                file_name = f"{sheet}_{os.getpid()}.xlsx"
                file_path = os.path.join(UPLOAD_FOLDER, file_name)

                # ✅ USE FORMAT FUNCTION
                create_formatted_excel(file_path, rows)

                zipf.write(file_path, arcname=f"{sheet}.xlsx")

            except Exception as e:
                print("ERROR:", sheet, e)

    return send_file(zip_path, as_attachment=True)


# ---------------- RUN ----------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
