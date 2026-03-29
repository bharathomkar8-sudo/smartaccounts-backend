from flask import Flask, request, send_file, render_template_string
import pandas as pd
import zipfile
from io import BytesIO

app = Flask(__name__)

uploaded_file = None

@app.route('/', methods=['GET', 'POST'])
def upload():
    global uploaded_file

    if request.method == 'POST':
        file = request.files['file']
        uploaded_file = BytesIO(file.read())

        xls = pd.ExcelFile(uploaded_file)
        sheets = xls.sheet_names

        return render_template_string('''
        <h2>Select Invoice Sheets</h2>
        <form method="POST" action="/process">
            {% for s in sheets %}
                <input type="checkbox" name="sheets" value="{{s}}" checked> {{s}}<br>
            {% endfor %}
            <br>
            <button type="submit">Process</button>
        </form>
        ''', sheets=sheets)

    return '''
    <h2>Upload Excel</h2>
    <form method="POST" enctype="multipart/form-data">
        <input type="file" name="file" required>
        <button type="submit">Upload</button>
    </form>
    '''

@app.route('/process', methods=['POST'])
def process():
    global uploaded_file

    uploaded_file.seek(0)
    xls = pd.ExcelFile(uploaded_file)

    selected_sheets = request.form.getlist('sheets')
    output_files = []

    for sheet in selected_sheets:
        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)

            # =========================
            # HEADER
            # =========================
            vch_no = df.iloc[1, 0]
            vch_date = df.iloc[1, 3]
            order_no = df.iloc[1, 4]
            order_date = df.iloc[1, 5]
            other_ref = df.iloc[1, 6]

            party_name = df.iloc[10, 0]
            gstin = df.iloc[10, 1]

            # =========================
            # ADDRESS (FIXED)
            # =========================
            addr1 = str(df.iloc[11, 0]).strip()
            addr2 = str(df.iloc[12, 0]).strip()
            addr3 = str(df.iloc[13, 0]).strip()

            full_address = f"{addr1}, {addr2}, {addr3}"

            # extract state & pincode
            state = ""
            pincode = ""

            if "–" in addr3:
                parts = addr3.split("–")
                state_part = parts[0].split(",")[-1].strip()
                state = state_part
                pincode = parts[1].strip()

            # =========================
            # CONSIGNEE
            # =========================
            con_addr1 = str(df.iloc[11, 4]).strip()
            con_addr2 = str(df.iloc[12, 4]).strip()
            con_full = f"{con_addr1}, {con_addr2}"

            # =========================
            # DESCRIPTION (FIXED CORE)
            # =========================
            descriptions = []

            for i in range(25, len(df)):  # B26 start
                val = str(df.iloc[i, 1]).strip()

                if val == "" or val.lower() == "nan":
                    continue

                if "gst" in val.lower():
                    break

                descriptions.append(val)

            # =========================
            # BUILD OUTPUT
            # =========================
            rows = []

            # HEADER ROW
            rows.append({
                "Voucher Type": "Sales E-Invoice",
                "VCH No / Inv No": vch_no,
                "Description": descriptions[0] if descriptions else "",
                "VCH Date": vch_date,
                "Order No": order_no,
                "Order Date": order_date,
                "Other Ref": other_ref,
                "POS": "",
                "Party Name": party_name,
                "Address": full_address,
                "State": state,
                "Pincode": pincode,
                "Party GSTIN": gstin,
                "Consignee Name": party_name,
                "Con Address": con_full,
                "Consignee State": state,
                "Consignee Pincode": pincode,
                "Con GSTIN": gstin,
                "Item Name / Code": "Header",
                "Is Item Header": "Yes"
            })

            # ITEM LINES
            for desc in descriptions:
                rows.append({
                    "Description": desc,
                    "Item Name / Code": desc,
                    "Is Item Header": ""
                })

            out_df = pd.DataFrame(rows)

            output = BytesIO()
            out_df.to_excel(output, index=False)
            output.seek(0)

            output_files.append((f"{sheet}.xlsx", output))

        except Exception as e:
            print("ERROR:", e)
            continue

    # ZIP
    memory_file = BytesIO()
    with zipfile.ZipFile(memory_file, 'w') as zf:
        for name, data in output_files:
            zf.writestr(name, data.getvalue())

    memory_file.seek(0)

    return send_file(memory_file, download_name="output.zip", as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
