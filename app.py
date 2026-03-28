@app.route('/process', methods=['POST'])
def process():

    filepath = request.form['filepath']
    selected_sheets = request.form.getlist('sheets')

    if not selected_sheets:
        return "No sheets selected"

    xls = pd.ExcelFile(filepath)

    # GST MASTER
    try:
        gst_df = pd.read_excel(xls, sheet_name="GST Details")
        gst_df.iloc[:,0] = gst_df.iloc[:,0].astype(str).str.strip().str.upper()
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

                # ---------------- HEADER ----------------
                vch_type = "Sales E-Invoice"

                try:
                    vch_no = str(df.iloc[10,16])
                except:
                    vch_no = sheet

                try:
                    vch_date = df.iloc[11,16]
                except:
                    vch_date = ""

                try:
                    order_no = df.iloc[19,1]
                    order_date = df.iloc[20,1]
                except:
                    order_no = ""
                    order_date = ""

                try:
                    other_ref = df.iloc[12,16]
                except:
                    other_ref = ""

                try:
                    pos = df.iloc[14,5]
                except:
                    pos = ""

                # GST → PARTY
                try:
                    gstin = str(df.iloc[16,1]).strip().upper()
                except:
                    gstin = ""

                party_name = "UNKNOWN"

                if gst_df is not None and gstin:
                    match = gst_df[gst_df.iloc[:,0] == gstin]
                    if not match.empty:
                        party_name = match.iloc[0,1]

                # ADDRESS
                try:
                    address = f"{df.iloc[11,0]} {df.iloc[12,0]} {df.iloc[13,0]}"
                except:
                    address = ""

                try:
                    state = df.iloc[14,1]
                    pincode = df.iloc[15,1]
                except:
                    state = ""
                    pincode = ""

                # ---------------- ITEMS ----------------
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
                        item = df.iloc[i,1]
                        qty = df.iloc[i,5]
                        rate = df.iloc[i,8]
                        amount = df.iloc[i,10]
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
                            "Rate": rate,
                            "Amount": amount
                        })

                # 🔥 ALWAYS CREATE FILE (even if empty)
                if len(rows) == 0:
                    rows.append({
                        "Voucher Type": vch_type,
                        "VCH No / Inv No": vch_no,
                        "Party Name": party_name,
                        "Note": "No items found"
                    })

                out_df = pd.DataFrame(rows)

                # ✅ ORIGINAL NAME (NO RANDOM)
                file_name = f"{sheet}.xlsx"
                file_path = os.path.join(UPLOAD_FOLDER, file_name)

                out_df.to_excel(file_path, index=False)

                zipf.write(file_path, arcname=file_name)

            except Exception as e:
                print("ERROR IN SHEET:", sheet, e)

                # 🔥 STILL CREATE FILE (NO LOSS)
                error_df = pd.DataFrame([{
                    "Sheet": sheet,
                    "Error": str(e)
                }])

                file_name = f"{sheet}_error.xlsx"
                file_path = os.path.join(UPLOAD_FOLDER, file_name)

                error_df.to_excel(file_path, index=False)
                zipf.write(file_path, arcname=file_name)

    return send_file(zip_path, as_attachment=True)
