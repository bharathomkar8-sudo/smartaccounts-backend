import pandas as pd
import os
import zipfile

# =========================
# FINAL OUTPUT COLUMNS
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
# CLEAN FUNCTION (convert NaN -> "")
def clean(val):
    if pd.isna(val):
        return ""
    val = str(val).strip()
    if val.lower() == "nan":
        return ""
    return val

# =========================
# DATE FORMAT FUNCTION (force DD-MM-YYYY)
def format_date(val):
    try:
        return pd.to_datetime(val).strftime("%d-%m-%Y")
    except:
        return ""

# =========================
# MAIN PROCESSING FUNCTION
def process_sheet(df, party_df):
    rows = []

    # Read fixed header fields
    voucher_type = "Sales E-Invoice"
    vch_no      = clean(df.iloc[10, 16])          # Q11
    vch_date    = format_date(df.iloc[11, 16])   # Q12 (DD-MM-YYYY)
    order_no    = clean(df.iloc[19, 1])          # B20
    order_date  = format_date(df.iloc[20, 1])    # B21
    other_ref   = clean(df.iloc[12, 16])         # Q13
    pos         = clean(df.iloc[14, 5])          # F15

    state = clean(df.iloc[14, 1])                # B15
    pincode = clean(df.iloc[15, 1])              # B16
    gst = clean(df.iloc[16, 1])                  # B17

    # Party lookup via GSTIN
    party_name = ""
    gst_clean = gst.strip()
    if gst_clean != "":
        match = party_df[party_df.iloc[:,0].astype(str).str.strip() == gst_clean]
        if not match.empty:
            party_name = clean(match.iloc[0,1])
        else:
            vch_no = vch_no + "-ERROR"
    consignee_name = party_name

    # Address lines (A12-A14 and E13-E14)
    address_lines = [
        clean(df.iloc[11, 0]),  # A12
        clean(df.iloc[12, 0]),  # A13
        clean(df.iloc[13, 0])   # A14
    ]
    con_address_lines = [
        clean(df.iloc[12, 4]),  # E13
        clean(df.iloc[13, 4])   # E14
    ]

    blank_count = 0
    # Loop over item rows starting at row 26 (index 25)
    for i in range(25, len(df)):
        desc = clean(df.iloc[i, 1])  # Column B
        print(f"Processing row: {i+1}, Desc: '{desc}'")
        if desc == "":
            blank_count += 1
            if blank_count >= 3:
                break
            else:
                continue
        else:
            blank_count = 0

        if desc.lower() == "end here":
            break

        row = dict.fromkeys(COLUMNS, "")

        # ========== HEADER FIELDS ==========
        row["Voucher Type"] = voucher_type
        row["VCH No / Inv No"] = vch_no
        row["VCH Date"] = vch_date
        row["Order No"] = order_no
        row["Order Date"] = order_date
        row["Other Ref"] = other_ref
        row["POS"] = pos

        row["State"] = state
        row["Pincode"] = pincode
        row["Party GSTIN"] = gst

        row["Party Name"] = party_name
        row["Consignee Name"] = consignee_name

        row["Consignee State"] = pos
        row["Consignee Pincode"] = clean(df.iloc[15, 5])  # F16
        row["Con GSTIN"] = gst

        # ========== DESCRIPTION ==========
        row["Description"] = desc
        row["Item header"] = desc

        # ========== ADDRESS FLOW ==========
        if i - 25 < len(address_lines):
            row["Address"] = address_lines[i - 25]
        if i - 25 < len(con_address_lines):
            row["Con Address"] = con_address_lines[i - 25]

        # ========== ITEM CODE (col C) ==========
        item_val = df.iloc[i, 2]
        if isinstance(item_val, (int, float)) and item_val > 1:
            row["Item Name / Code"] = str(int(item_val))
        else:
            row["Item Name / Code"] = "Header"
            row["Is Item Header"] = "Yes"

        # ========== SAFE FUNCTION ==========
        def safe(val):
            if pd.isna(val):
                return ""
            try:
                if float(val) == 0:
                    return ""
            except:
                pass
            return val

        # ========== OTHER COLUMNS ==========
        row["width"] = safe(df.iloc[i, 3])    # D
        row["Height"] = safe(df.iloc[i, 4])   # E
        row["Qty"] = safe(df.iloc[i, 5])      # F
        row["Extraudf"] = safe(df.iloc[i, 6]) # G
        row["Billedqty"] = safe(df.iloc[i, 6])# G
        row["Rate"] = safe(df.iloc[i, 8])     # I

        # ========== CALCULATIONS ==========
        try:
            qty = float(row["Qty"]) if row["Qty"] != "" else 0
            rate = float(row["Rate"]) if row["Rate"] != "" else 0
        except:
            qty, rate = 0, 0
        taxable = qty * rate

        dis_val = df.iloc[i, 9]  # J
        try:
            dis_percent = float(dis_val) if not pd.isna(dis_val) else ""
        except:
            dis_percent = ""
        if dis_percent != "":
            amount = taxable * (1 - (dis_percent / 100))
        else:
            amount = taxable

        gst_amt = round(amount * 0.18, 2)
        total = amount + gst_amt

        row["Taxable Value"] = taxable if taxable else ""
        row["Dis%"] = dis_percent if dis_percent != "" else ""
        row["Amount"] = amount if amount else ""

        row["Sales Ledger"] = "GST IGST Sales@18%"
        row["IGST Ledger"] = "OUTPUT IGST @ 18%"
        row["IGST Amount"] = gst_amt if gst_amt else ""

        row["Invoice Amt"] = total if total else ""

        rows.append(row)

    print(f"Rows created: {len(rows)}")
    return pd.DataFrame(rows)

# =========================
# MAIN EXECUTION
def run_full_process(file_path):
    excel = pd.ExcelFile(file_path)
    party_df = excel.parse("Customer Name & GST", header=None)
    os.makedirs("output", exist_ok=True)
    created_files = []
    for sheet in excel.sheet_names:
        if sheet == "Customer Name & GST":
            continue
        df = excel.parse(sheet, header=None)
        output_df = process_sheet(df, party_df)
        print(f"Sheet: {sheet}, Rows: {len(output_df)}")
        if len(output_df) > 0:
            out_file = f"output/{sheet}.xlsx"
            output_df.to_excel(out_file, index=False)
            created_files.append(out_file)
    # Create ZIP of all output files
    zip_path = "output.zip"
    with zipfile.ZipFile(zip_path, 'w') as z:
        for file in created_files:
            z.write(file, os.path.basename(file))
    print("✅ Done. ZIP created:", zip_path)
