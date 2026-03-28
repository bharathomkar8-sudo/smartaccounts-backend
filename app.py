from flask import Flask, request, send_file
import pandas as pd
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


@app.route('/')
def home():
    return '''
    <h2>Upload Sales Excel</h2>
    <form action="/process" method="post" enctype="multipart/form-data">
        <input type="file" name="file">
        <button type="submit">Upload & Process</button>
    </form>
    '''


@app.route('/process', methods=['POST'])
def process():

    file = request.files['file']
    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)

    xls = pd.ExcelFile(filepath)

    # 🔹 GST MASTER
    gst_df = pd.read_excel(xls, sheet_name="GST Details")
    gst_df.iloc[:,0] = gst_df.iloc[:,0].astype(str).str.strip().str.upper()

    final_data = []

    for sheet in xls.sheet_names:

        if sheet == "GST Details":
            continue

        df = pd.read_excel(xls, sheet_name=sheet, header=None)

        try:
            # 🔹 HEADER
            vch_no = str(df.iloc[10,16])
            vch_date = df.iloc[11,16]
            order_no = df.iloc[19,1]
            order_date = df.iloc[20,1]
            pos = df.iloc[14,5]

            # 🔹 GST → PARTY
            party_gstin = str(df.iloc[16,1]).strip().upper()

            match = gst_df[gst_df.iloc[:,0] == party_gstin]

            if not match.empty:
                party_name = match.iloc[0,1]
            else:
                party_name = "UNKNOWN"

            # 🔹 ADDRESS
            address = " ".join([
                str(df.iloc[11,0]),
                str(df.iloc[12,0]),
                str(df.iloc[13,0])
            ])

            state = df.iloc[14,1]
            pincode = df.iloc[15,1]

            # 🔹 FIND ITEM END (GST BREAK UP)
            end_row = df[df.apply(
                lambda row: row.astype(str).str.contains("GST Break up", case=False).any(),
                axis=1
            )].index[0]

            start_row = 25  # B26

            # 🔹 LOOP ITEMS
            for i in range(start_row, end_row):

                desc = df.iloc[i,1]
                qty = df.iloc[i,5]
                rate = df.iloc[i,8]
                amount = df.iloc[i,10]

                # ✅ FILTER VALID ITEMS
                if pd.notna(qty) and qty > 0:

                    final_data.append({
                        "Voucher Type": "Sales E-Invoice",
                        "VCH No": vch_no,
                        "Date": vch_date,
                        "Party": party_name,
                        "GSTIN": party_gstin,
                        "Address": address,
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

        except Exception as e:
            print(f"Error in sheet {sheet}: {e}")
            continue

    # 🔹 OUTPUT
    output_df = pd.DataFrame(final_data)

    output_file = os.path.join(UPLOAD_FOLDER, "output.xlsx")
    output_df.to_excel(output_file, index=False)

    return send_file(output_file, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
