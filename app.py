@app.route('/process', methods=['POST'])
def process():

    import uuid

    filepath = request.form['filepath']
    selected_sheets = request.form.getlist('sheets')

    xls = pd.ExcelFile(filepath)

    gst_df = pd.read_excel(xls, sheet_name="GST Details")
    gst_df.iloc[:,0] = gst_df.iloc[:,0].astype(str).str.strip().str.upper()

    zip_path = os.path.join(UPLOAD_FOLDER, "output.zip")

    # 🔥 REMOVE OLD ZIP
    if os.path.exists(zip_path):
        os.remove(zip_path)

    with zipfile.ZipFile(zip_path, 'w') as zipf:

        for sheet in selected_sheets:

            print("Processing:", sheet)   # DEBUG

            df = pd.read_excel(xls, sheet_name=sheet, header=None)

            try:
                vch_no = str(df.iloc[10,16])
                vch_date = df.iloc[11,16]

                gstin = str(df.iloc[16,1]).strip().upper()

                match = gst_df[gst_df.iloc[:,0] == gstin]
                party_name = match.iloc[0,1] if not match.empty else "UNKNOWN"

                # 🔥 FIND END ROW
                end_row = df[df.apply(
                    lambda r: r.astype(str).str.contains("GST Break up", case=False).any(),
                    axis=1
                )].index[0]

                start_row = 25

                rows = []

                for i in range(start_row, end_row):

                    qty = df.iloc[i,5]
                    rate = df.iloc[i,8]

                    if pd.notna(qty) and qty > 0:

                        rows.append({
                            "VCH No": vch_no,
                            "Party": party_name,
                            "Qty": qty,
                            "Rate": rate
                        })

                # 🔥 IMPORTANT CHECK
                if len(rows) == 0:
                    print(f"Skipping empty sheet: {sheet}")
                    continue

                out_df = pd.DataFrame(rows)

                # 🔥 UNIQUE FILE NAME
                unique = str(uuid.uuid4())[:6]
                file_name = f"{sheet}_{unique}.xlsx"
                file_path = os.path.join(UPLOAD_FOLDER, file_name)

                out_df.to_excel(file_path, index=False)

                print("Adding to ZIP:", file_name)

                zipf.write(file_path, arcname=file_name)

            except Exception as e:
                print(f"Error in {sheet}: {e}")
                continue

    return send_file(zip_path, as_attachment=True)
