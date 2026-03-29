from flask import Flask, request, send_file, render_template_string
import pandas as pd
import zipfile
from io import BytesIO

app = Flask(__name__)

# =========================
# STEP 1: UPLOAD + SHEET LIST
# =========================
@app.route('/', methods=['GET', 'POST'])
def upload():

    if request.method == 'POST':
        file = request.files['file']

        file_stream = file.read()
        xls = pd.ExcelFile(BytesIO(file_stream))

        sheets = xls.sheet_names

        return render_template_string('''
        <h2>Select Invoice Sheets</h2>

        <form method="POST" action="/process" enctype="multipart/form-data">
            <input type="hidden" name="file_data" value="{{file_data}}">

            {% for s in sheets %}
                <input type="checkbox" name="sheets" value="{{s}}" checked> {{s}}<br>
            {% endfor %}

            <br>
            <button type="submit">Process</button>
        </form>
        ''', sheets=sheets, file_data=file_stream.hex())

    return '''
    <h2>Upload Excel</h2>
    <form method="POST" enctype="multipart/form-data">
        <input type="file" name="file" required>
        <button type="submit">Upload</button>
    </form>
    '''

# =========================
# STEP 2: PROCESS
# =========================
@app.route('/process', methods=['POST'])
def process():

    selected_sheets = request.form.getlist('sheets')

    file_data = bytes.fromhex(request.form['file_data'])
    xls = pd.ExcelFile(BytesIO(file_data))

    output_files = []

    for sheet in selected_sheets:
        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)

            if df.shape[0] < 30:
                continue

            # =========================
            # HEADER
            # =========================
            voucher_type = "Sales E-Invoice"
            vch_no = df.iloc[1, 0]
            vch_date = df.iloc[1, 3]
            order_no = df.iloc[1, 4]
            order_date = pd.to_datetime(df.iloc[1, 5], errors='coerce')
            other_ref = df.iloc[1, 6]

            party_name = df.iloc[10, 0]
            gstin = df.iloc[10, 1]

            # =========================
            # ADDRESS (3 lines)
            # =========================
            addr1 = df.iloc[11, 0]
            addr2 = df.iloc[12, 0]
            addr3 = df.iloc[13, 0]

            # =========================
            # CONSIGNEE ADDRESS
            # =========================
            con_addr1 = df.iloc[11, 4]
            con_addr2 = df.iloc[12, 4]

            # =========================
            # DESCRIPTION (B26 → ___)
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
            # BUILD OUTPUT
            # =========================
            rows = []

            # HEADER ROW
            rows.append({
                "Voucher Type": voucher_type,
                "VCH No / Inv No": vch_no,
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

            # Address continuation
            rows.append({"Address": addr2})
            rows.append({"Address": addr3})

            # Consignee continuation
            rows.append({"Con Address": con_addr2})

            # Description rows
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

            output = BytesIO()
            out_df.to_excel(output, index=False)
            output.seek(0)

            output_files.append((f"{sheet}.xlsx", output))

        except Exception as e:
            print("ERROR:", sheet, e)
            continue

    # =========================
    # ZIP DOWNLOAD
    # =========================
    memory_file = BytesIO()

    with zipfile.ZipFile(memory_file, 'w') as zf:
        for filename, data in output_files:
            zf.writestr(filename, data.getvalue())

    memory_file.seek(0)

    return send_file(memory_file, download_name='output.zip', as_attachment=True)


# =========================
# RUN
# =========================
if __name__ == "__main__":
    app.run(debug=True)
