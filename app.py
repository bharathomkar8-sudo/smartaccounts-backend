from flask import Flask, request, send_file
import pandas as pd
import os
import zipfile
from io import BytesIO

app = Flask(__name__)

@app.route('/')
def home():
    return "Smart Accounts Running"

@app.route('/process', methods=['POST'])
def process():

    file = request.files['file']
    selected_sheets = request.form.getlist('sheets')

    output_files = []

    xls = pd.ExcelFile(file)

    for sheet in selected_sheets:
        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)

            # ✅ Skip invalid sheets
            if df.shape[0] < 30:
                continue

            # =========================
            # HEADER VALUES
            # =========================
            voucher_type = "Sales E-Invoice"
            vch_no = df.iloc[1, 0] if not pd.isna(df.iloc[1, 0]) else ""
            vch_date = df.iloc[1, 3]
            order_no = df.iloc[1, 4]
            order_date = pd.to_datetime(df.iloc[1, 5], errors='coerce')

            other_ref = df.iloc[1, 6]

            party_name = df.iloc[10, 0]
            gstin = df.iloc[10, 1]

            # =========================
            # ADDRESS (3 LINES FIXED)
            # =========================
            addr1 = df.iloc[11, 0]
            addr2 = df.iloc[12, 0]
            addr3 = df.iloc[13, 0]

            # =========================
            # CONSIGNEE ADDRESS FIXED
            # =========================
            con_addr1 = df.iloc[11, 4]
            con_addr2 = df.iloc[12, 4]

            # =========================
            # DESCRIPTION (B26 onwards)
            # =========================
            descriptions = []

            for i in range(25, len(df)):
                val = str(df.iloc[i, 1]).strip()

                if val == "" or val.lower() == "nan":
                    continue

                if "___" in val:
                    break

                descriptions.append(val)

            # =========================
            # OUTPUT DATA
            # =========================
            rows = []

            # HEADER ROW
            rows.append({
                "Voucher Type": voucher_type,
                "VCH No": vch_no,
                "Description": descriptions[0] if descriptions else "",
                "VCH Date": vch_date,
                "Order No": order_no,
                "Order Date": order_date,
                "Other Ref": other_ref,
                "POS": "",
                "Party Name": party_name,
                "Address": addr1,
                "State": "",
                "Pincode": "",
                "Party GSTIN": gstin,
                "Consignee Name": party_name,
                "Con Address": con_addr1,
                "Consignee State": "",
                "Consignee Pincode": "",
                "Con GSTIN": gstin,
                "Item Name / Code": "Header",
                "Is Item Header": "Yes"
            })

            # ADDRESS CONTINUATION
            rows.append({"Address": addr2})
            rows.append({"Address": addr3})

            # CONSIGNEE CONTINUATION (NO SHIFT)
            rows.append({"Con Address": con_addr2})

            # =========================
            # DESCRIPTION LINES
            # =========================
            for desc in descriptions:
                rows.append({
                    "Description": desc,
                    "Item Name / Code": desc if len(desc) > 1 else "Header",
                    "Is Item Header": "Yes" if len(desc) <= 1 else ""
                })

            # =========================
            # SAVE FILE
            # =========================
            out_df = pd.DataFrame(rows)

            filename = f"{sheet}.xlsx"
            out_df.to_excel(filename, index=False)
            output_files.append(filename)

        except Exception as e:
            print("ERROR:", sheet, e)
            continue

    # =========================
    # ZIP DOWNLOAD
    # =========================
    memory_file = BytesIO()
    with zipfile.ZipFile(memory_file, 'w') as zf:
        for f in output_files:
            zf.write(f)
            os.remove(f)

    memory_file.seek(0)

    return send_file(memory_file, download_name='output.zip', as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
