@app.route('/process', methods=['POST'])
def process():
    import uuid

    filepath = request.form['filepath']
    selected_sheets = request.form.getlist('sheets')

    if not selected_sheets:
        return "No sheets selected"

    xls = pd.ExcelFile(filepath)

    # 🔹 GST MASTER (safe load)
    try:
        gst_df = pd.read_excel(xls, sheet_name="GST Details")
        gst_df.iloc[:,0] = gst_df.iloc[:,0].astype(str).str.strip().str.upper()
    except:
        gst_df = None

    zip_path = os.path.join(UPLOAD_FOLDER, "output.zip")

    # remove old zip
    if os.path.exists(zip_path):
        os.remove(zip_path)

    with zipfile.ZipFile(zip_path, 'w') as zipf:

        for sheet in selected_sheets:
            try:
                print("Processing:", sheet)

                df = pd.read_excel(xls, sheet_name=sheet, header=None)

                # 🔹 HEADER (SAFE EXTRACTION)
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
                    pos = df.iloc[14,5]
                except:
                    pos = ""

                # 🔹 GST → PARTY
                try:
                    gstin = str(df.iloc[16,1]).strip().upper()
                except:
                    gstin = ""

                party_name = "UNKNOWN"

                if gst_df is not None and gstin:
                    match = gst_df[gst_df.iloc[:,0] == gstin]
                    if not match.empty:
                        party_name = match.iloc[0,1]

                # 🔹 ADDRESS
                try:
                    address1 = str(df.iloc[11,0])
                    address2 = str(df.iloc[12,0])
                    address3 = str(df.iloc[13,0])
                except:
                    address1 = address2 = address3 = ""

                try:
                    state = df.iloc[14,1]
                    pincode = df.iloc[15,1]
                except:
                    state = ""
                    pincode = ""

                # 🔹 FIND END ROW (GST BREAK UP)
                try:
                    end_row = df[df.apply(
                        lambda r: r.astype(str).str.contains("GST Break up", case=False).any(),
                        axis=1
                    )].index[0]
                except:
                    end_row = len(df)

                start_row = 25

                rows = []

                # 🔹 ITEM LOOP (SAFE – NO SKIP OF FILE)
                for i in range(start_row, end_row):

                    try:
                        desc = df.iloc[i,1]
                        qty = df.iloc[i,5]
                        rate = df.iloc[i,8]
                        amount = df.iloc[i,10]
                    except:
                        continue

                    # only valid lines
                    if pd.notna(qty) and qty != 0:

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

                # 🔥 IMPORTANT: ALWAYS CREATE FILE
                if len(rows) == 0:
                    rows.append({
                        "Voucher Type": "Sales E-Invoice",
                        "VCH No": vch_no,
                        "Party": party_name,
                        "Note": "No items detected"
                    })

                out_df = pd.DataFrame(rows)

                # 🔹 UNIQUE FILE NAME
                unique_id = str(uuid.uuid4())[:6]
                file_name = f"{sheet}_{unique_id}.xlsx"
                file_path = os.path.join(UPLOAD_FOLDER, file_name)

                out_df.to_excel(file_path, index=False)

                zipf.write(file_path, arcname=file_name)

                print("Added:", file_name)

            except Exception as e:
                print(f"Error in sheet {sheet}: {e}")
                continue

    return send_file(zip_path, as_attachment=True)
