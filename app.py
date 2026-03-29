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

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = yellow
        cell.font = bold

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

    # GST MASTER
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
                df = pd.read_excel(xls, sheet_name=sheet, header=None)

                # -------- HEADER --------
                vch_type = "Sales E-Invoice"

                vch_no = str(df.iloc[10, 16])
                vch_date = pd.to_datetime(df.iloc[11, 16]).strftime("%d-%m-%Y")
                order_no = df.iloc[19, 1]
                order_date = pd.to_datetime(df.iloc[20, 1]).strftime("%d-%m-%Y")
                other_ref = df.iloc[12, 16]
                pos = df.iloc[14, 5]

                gstin = str(df.iloc[16, 1]).strip().upper()

                party_name = "UNKNOWN"
                if gst_df is not None:
                    match = gst_df[gst_df.iloc[:, 0] == gstin]
                    if not match.empty:
                        party_name = match.iloc[0, 1]

                # -------- ADDRESS (3 LINE) --------
                addr1 = df.iloc[10, 0] if df.shape[0] > 10 else ""
                addr2 = df.iloc[11, 0] if df.shape[0] > 11 else ""
                addr3 = df.iloc[12, 0] if df.shape[0] > 12 else ""

                state = df.iloc[14, 1]
                pincode = df.iloc[15, 1]

                # CONSIGNEE ADDRESS
                con_addr1 = df.iloc[12, 4] if df.shape[0] > 12 else ""
                con_addr2 = df.iloc[13, 4] if df.shape[0] > 13 else ""

                # -------- FIND END ("___") --------
                start_row = 25
                end_row = len(df)

                for i in range(start_row, len(df)):
                    val = str(df.iloc[i, 1]).strip()
                    if val.startswith("___"):
                        end_row = i
                        break

                rows = []
                first = True

                for i in range(start_row, end_row):

                    item = df.iloc[i, 1]
                    if pd.isna(item):
                        item = ""

                    try:
                        qty = df.iloc[i, 5]
                        rate = df.iloc[i, 8]
                        amount = df.iloc[i, 10]
                    except:
                        qty, rate, amount = "", "", ""

                    # -------- HEADER FIRST ROW --------
                    if first:

                        # MAIN HEADER ROW
                        row = {
                            "Voucher Type": vch_type,
                            "VCH No / Inv No": vch_no,
                            "VCH Date": vch_date,
                            "Order No": order_no,
                            "Order Date": order_date,
                            "Other Ref": other_ref,
                            "POS": pos,
                            "Party Name": party_name,
                            "Address": addr1,
                            "State": state,
                            "Pincode": pincode,
                            "Party GSTIN": gstin,
                            "Consignee Name": party_name,
                            "Con Address": con_addr1,
                            "Consignee State": pos,
                            "Consignee Pincode": pincode,
                            "Con GSTIN": gstin
                        }
                        rows.append(row)

                        # ADDRESS LINE 2
                        rows.append({"Address": addr2})

                        # ADDRESS LINE 3
                        rows.append({"Address": addr3})

                        # CONSIGNEE ADDRESS LINE 2
                        rows.append({"Con Address": con_addr2})

                        first = False

                    # -------- ITEM ROW --------
                    row = {}

                    row["Description"] = item
                    row["Item Name / Code"] = item
                    row["Item header"] = item

                    row["Qty"] = qty
                    row["Billedqty"] = qty
                    row["Rate"] = rate
                    row["Taxable Value"] = amount
                    row["Amount"] = amount
                    row["Sales Ledger"] = "Sales"

                    rows.append(row)

                # SAVE
                file_name = f"{sheet}_{os.getpid()}.xlsx"
                file_path = os.path.join(UPLOAD_FOLDER, file_name)

                create_formatted_excel(file_path, rows)
                zipf.write(file_path, arcname=f"{sheet}.xlsx")

            except Exception as e:
                print("ERROR:", sheet, e)

    return send_file(zip_path, as_attachment=True)


# ---------------- RUN ----------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
