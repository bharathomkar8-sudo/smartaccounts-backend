from flask import Flask, request, send_file, render_template_string
import pandas as pd
import zipfile
from io import BytesIO

app = Flask(__name__)

uploaded_file = None

# =========================
# UPLOAD PAGE
# =========================
@app.route('/', methods=['GET', 'POST'])
def upload():
    global uploaded_file

    if request.method == 'POST':
        file = request.files['file']
        uploaded_file = BytesIO(file.read())

        xls = pd.ExcelFile(uploaded_file)
        sheets = xls.sheet_names

        return render_template_string('''
        <h2>Select Sheets</h2>
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

# =========================
# PROCESS LOGIC
# =========================
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
            # HEADER DATA
            # =========================
            vch_no = df.iloc[1, 0]
            vch_date = df.iloc[1, 3]
            order_no = df.iloc[1, 4]
            order_date = df.iloc[1, 5]
            other_ref = df.iloc[1, 6]

            party_name = df.iloc[10, 0]
            gstin = df.iloc[10, 1]

            # =========================
            # ADDRESS (EXACT FORMAT)
            # =========================
            addr1 = str(df.iloc[11, 0]).strip()
            addr2 = str(df.iloc[12, 0]).strip()
            addr3 = str(df.iloc[13, 0]).strip()

            # extract state & pincode
            state = ""
            pincode = ""

            if "–" in addr3:
                parts = addr3.split("–")
                if len(parts) == 2:
                    pincode = parts[1].strip()
                    state = parts[0].split(",")[-1].strip()

            # =========================
            # DESCRIPTION (STRICT)
            # =========================
            descriptions = []

            for i in range(25, len(df)):  # B26
                val = str(df.iloc[i, 1]).strip()

                if val == "" or val.lower() == "nan":
                    continue

                # STOP CONDITION (VERY IMPORTANT)
                if "gst break" in val.lower():
                    break

                descriptions.append(val)

            # =========================
            # BUILD OUTPUT
            # =========================
            rows = []

            # FIRST ROW (ONLY FIRST DESCRIPTION)
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
                "Address": addr1,
                "State": state,
                "Pincode": pincode,
                "Party GSTIN": gstin,
                "Consignee Name": party_name,
                "Con Address": addr1,
                "Consignee State": state,
                "Consignee Pincode": pincode,
                "Con GSTIN": gstin,
                "Item Name / Code": "Header",
                "Is Item Header": "Yes"
            })

            # ADDRESS CONTINUE (EXACT FORMAT)
            rows.append({"Address": addr2})
            rows.append({"Address": addr3})

            # DESCRIPTION CONTINUE (NO REPEAT FIRST)
            for desc in descriptions[1:]:

                # 👉 FORMULA LOGIC LIKE YOUR EXCEL
                item_code = desc if len(desc.strip()) > 1 else "Header"
                is_header = "Yes" if item_code == "Header" else ""

                rows.append({
                    "Description": desc,
                    "Item Name / Code": item_code,
                    "Is Item Header": is_header
                })

            # =========================
            # EXPORT
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


if __name__ == "__main__":
    app.run(debug=True)
