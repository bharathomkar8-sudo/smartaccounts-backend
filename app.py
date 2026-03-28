from flask import Flask, request, send_file
import pandas as pd
import os
import zipfile
import uuid

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# 🔹 HOME PAGE
@app.route('/')
def home():
    return '''
    <h2>Upload Sales Excel</h2>
    <form action="/upload" method="post" enctype="multipart/form-data">
        <input type="file" name="file">
        <button type="submit">Upload</button>
    </form>
    '''


# 🔹 STEP 1: READ FILE & SHOW SHEETS
@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']
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


# 🔹 STEP 2: PROCESS SELECTED SHEETS
@app.route('/process', methods=['POST'])
def process():

    filepath = request.form['filepath']
    selected_sheets = request.form.getlist('sheets')

    xls = pd.ExcelFile(filepath)

    # 🔹 GST MASTER
    gst_df = pd.read_excel(xls, sheet_name="GST Details")
    gst_df.iloc[:,0] = gst_df.iloc[:,0].astype(str).str.strip().str.upper()

    zip_path = os.path.join(UPLOAD_FOLDER, "output.zip")

    # Remove old zip
    if os.path.exists(zip_path):
        os.remove(zip_path)

    with zipfile.ZipFile(zip_path, 'w') as zipf:

        for sheet in selected_sheets:

            print("Processing:", sheet)

            df = pd.read_excel(xls, sheet_name=sheet, header=None)

            try:
                # 🔹 HEADER
                vch_no = str(df.iloc[10,16])
                vch_date = df.iloc[11,16]

                order_no = df.iloc[19,1]
                order_date = df.iloc[20,1]
                pos = df.iloc[14,5]

                # 🔹 GST → PARTY
                gstin = str(df.iloc[16,1]).strip().upper()
                match = gst_df[gst_df.iloc[:,0] == gstin]
                party_name = match.iloc[0,1] if not match.empty else "UNKNOWN"

                # 🔹 ADDRESS (A12, A13, A14)
                address1 = str(df.iloc[11,0])
                address2 = str(df.iloc[12,0])
                address3 = str(df.iloc[13,0])

                state = df.iloc[14,1]
                pincode = df.iloc[15,1]

                # 🔹 FIND END ROW (GST Break up)
                end_row = df[df.apply(
                    lambda r: r.astype(str).str.contains("GST Break up", case=False).any(),
                    axis=1
                )].index[0]

                start_row = 25  # B26

                rows = []

                # 🔹 LOOP ITEMS
                for i in range(start_row, end_row):

                    desc = df.iloc[i,1]
                    qty = df.iloc[i,5]
                    rate = df.iloc[i,8]
                    amount = df.iloc[i,10]

                    # VALID ITEM
                    if pd.notna(qty) and qty > 0:

                        rows.append({
                            "Voucher Type": "Sales E-Invoice",
                            "VCH No": vch_no,
                            "Date": vch_date,
                            "Party": party_name,
                            "GSTIN": gstin,
                            "Address1": address1,
                            "Address2": address2,
                            "Address3": address3,
                            "State": state,
                            "Pincode": pincode,
                            "Item": desc,
                            "Qty": qty,
                            "Rate": rate,
                            "Amount": amount,
                            "POS": pos,
                            "Order No": order_no,
                            "Order Date": order_date
                        })

                # 🔴 SKIP EMPTY SHEETS
                if len(rows) == 0:
                    print("Skipping empty:", sheet)
                    continue

                out_df = pd.DataFrame(rows)

                # 🔹 UNIQUE FILE NAME
                unique_id = str(uuid.uuid4())[:6]
                file_name = f"{sheet}_{unique_id}.xlsx"
                file_path = os.path.join(UPLOAD_FOLDER, file_name)

                out_df.to_excel(file_path, index=False)

                print("Adding to ZIP:", file_name)

                zipf.write(file_path, arcname=file_name)

            except Exception as e:
                print(f"Error in {sheet}: {e}")
                continue

    return send_file(zip_path, as_attachment=True)


# 🔹 RUN
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
