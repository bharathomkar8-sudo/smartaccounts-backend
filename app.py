from flask import Flask, request, send_file, render_template_string
import pandas as pd
import zipfile
from io import BytesIO

app = Flask(__name__)
uploaded_file = None

# =========================
# FINAL COLUMN STRUCTURE
# =========================
COLUMNS = [
"Voucher Type","VCH No / Inv No","Description","VCH Date","Order No","Order Date","Other Ref","POS",
"Party Name","Address","State","Pincode","Party GSTIN",
"Consignee Name","Con Address","Consignee State","Consignee Pincode","Con GSTIN",
"Item Name / Code","Is Item Header","width","Height","Qty","Extraudf","Billedqty","Rate",
"Taxable Value","Dis%","Amount","Sales Ledger",
"CGST Ledger","CGST Amt","SGST Ledger","SGST Amount",
"IGST Ledger","IGST Amount","Round off","Invoice Amt","Item header"
]

# =========================
# UPLOAD
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
# PROCESS
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

            # HEADER
            vch_no = df.iloc[1, 0]
            vch_date = df.iloc[1, 3]
            order_no = df.iloc[1, 4]
            order_date = df.iloc[1, 5]
            other_ref = df.iloc[1, 6]

            party_name = df.iloc[10, 0]
            gstin = df.iloc[10, 1]

            addr1 = str(df.iloc[11, 0]).strip()
            addr2 = str(df.iloc[12, 0]).strip()
            addr3 = str(df.iloc[13, 0]).strip()

            # STATE + PINCODE
            state = ""
            pincode = ""

            if "–" in addr3:
                parts = addr3.split("–")
                if len(parts) == 2:
                    pincode = parts[1].strip()
                    state = parts[0].split(",")[-1].strip()

            # DESCRIPTION (B26)
            descriptions = []
            for i in range(25, len(df)):
                val = str(df.iloc[i, 1]).strip()

                if val == "" or val.lower() == "nan":
                    continue

                if "gst break" in val.lower():
                    break

                descriptions.append(val)

            rows = []

            # =========================
            # HEADER ROW
            # =========================
            row = dict.fromkeys(COLUMNS, "")

            row["Voucher Type"] = "Sales E-Invoice"
            row["VCH No / Inv No"] = vch_no
            row["Description"] = descriptions[0]
            row["VCH Date"] = vch_date
            row["Order No"] = order_no
            row["Order Date"] = order_date
            row["Other Ref"] = other_ref
            row["Party Name"] = party_name
            row["Address"] = addr1
            row["State"] = state
            row["Pincode"] = pincode
            row["Party GSTIN"] = gstin

            row["Consignee Name"] = party_name
            row["Con Address"] = addr1
            row["Con GSTIN"] = gstin

            row["Item Name / Code"] = "Header"
            row["Is Item Header"] = "Yes"
            row["Item header"] = descriptions[0]

            rows.append(row)

            # =========================
            # SECOND ROW (ADDRESS CONTINUE)
            # =========================
            row = dict.fromkeys(COLUMNS, "")
            row["Description"] = descriptions[1] if len(descriptions) > 1 else ""
            row["Address"] = addr2
            row["Con Address"] = addr2
            row["Item Name / Code"] = "Header"
            row["Is Item Header"] = "Yes"
            row["Item header"] = descriptions[1] if len(descriptions) > 1 else ""
            rows.append(row)

            # =========================
            # THIRD ROW (FINAL ADDRESS)
            # =========================
            row = dict.fromkeys(COLUMNS, "")
            row["Address"] = addr3
            row["Con Address"] = addr3
            rows.append(row)

            # =========================
            # ITEM ROWS
            # =========================
            for desc in descriptions[2:]:

                row = dict.fromkeys(COLUMNS, "")

                row["Description"] = desc
                row["Item Name / Code"] = "391990"   # default (change later dynamic)
                row["Qty"] = 1
                row["Rate"] = 500
                row["Taxable Value"] = 5000
                row["Amount"] = 4900
                row["Sales Ledger"] = "GST IGST Sales@18%"
                row["IGST Ledger"] = "OUTPUT IGST @ 18%"
                row["IGST Amount"] = 900
                row["Invoice Amt"] = 5801
                row["Item header"] = desc

                rows.append(row)

            out_df = pd.DataFrame(rows, columns=COLUMNS)

            output = BytesIO()
            out_df.to_excel(output, index=False)
            output.seek(0)

            output_files.append((f"{sheet}.xlsx", output))

        except Exception as e:
            print("ERROR:", e)
            continue

    memory_file = BytesIO()

    with zipfile.ZipFile(memory_file, 'w') as zf:
        for name, data in output_files:
            zf.writestr(name, data.getvalue())

    memory_file.seek(0)

    return send_file(memory_file, download_name="output.zip", as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
