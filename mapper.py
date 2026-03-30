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

def clean(val):
    if pd.isna(val):
        return ""
    val = str(val).strip()
    if val.lower() == "nan":
        return ""
    return val

def format_date(val):
    try:
        return pd.to_datetime(val).strftime("%d-%m-%Y")
    except:
        return ""

# =========================
# PROCESS FUNCTION (FIXED)
# =========================
def process_sheet(df, party_df):

    rows = []

    voucher_type = "Sales E-Invoice"
    vch_no = clean(df.iloc[10, 16])
    vch_date = format_date(df.iloc[11, 16])
    order_no = clean(df.iloc[19, 1])
    order_date = format_date(df.iloc[20, 1])
    other_ref = clean(df.iloc[12, 16])
    pos = clean(df.iloc[14, 5])

    state = clean(df.iloc[14, 1])
    pincode = clean(df.iloc[15, 1])
    gst = clean(df.iloc[16, 1])

    # PARTY LOOKUP
    party_name = ""
    gst_clean = gst.strip()

    if gst_clean != "":
        match = party_df[party_df.iloc[:,0].astype(str).str.strip() == gst_clean]
        if not match.empty:
            party_name = clean(match.iloc[0,1])
        else:
            vch_no = vch_no + "-ERROR"

    consignee_name = party_name

    address_lines = [
        clean(df.iloc[11, 0]),
        clean(df.iloc[12, 0]),
        clean(df.iloc[13, 0])
    ]

    con_address_lines = [
        clean(df.iloc[12, 4]),
        clean(df.iloc[13, 4])
    ]

    # ✅ FIXED LOOP
    for i in range(25, len(df)):

        desc = clean(df.iloc[i, 1])

        # STOP condition
        if desc == "":
            if i > 30:
                break
            else:
                continue

        if desc.lower() == "end here":
            break

        row = dict.fromkeys(COLUMNS, "")

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
        row["Consignee Pincode"] = clean(df.iloc[15, 5])
        row["Con GSTIN"] = gst

        row["Description"] = desc
        row["Item header"] = desc

        if i - 25 < len(address_lines):
            row["Address"] = address_lines[i - 25]

        if i - 25 < len(con_address_lines):
            row["Con Address"] = con_address_lines[i - 25]

        item_val = df.iloc[i, 2]

        if isinstance(item_val, (int, float)) and item_val > 1:
            row["Item Name / Code"] = str(int(item_val))
        else:
            row["Item Name / Code"] = "Header"

        if row["Item Name / Code"] == "Header":
            row["Is Item Header"] = "Yes"

        def safe(val):
            if pd.isna(val):
                return ""
            try:
                if float(val) == 0:
                    return ""
            except:
                pass
            return val

        row["width"] = safe(df.iloc[i, 3])
        row["Height"] = safe(df.iloc[i, 4])
        row["Qty"] = safe(df.iloc[i, 5])
        row["Extraudf"] = safe(df.iloc[i, 6])
        row["Billedqty"] = safe(df.iloc[i, 6])
        row["Rate"] = safe(df.iloc[i, 8])

        try:
            qty = float(row["Qty"]) if row["Qty"] != "" else 0
            rate = float(row["Rate"]) if row["Rate"] != "" else 0
        except:
            qty, rate = 0, 0

        taxable = qty * rate

        dis_val = df.iloc[i, 9]
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

    return pd.DataFrame(rows)

# =========================
# MAIN EXECUTION
# =========================
def run_full_process(file_path):

    excel = pd.ExcelFile(file_path)

    party_df = excel.parse("Customer Name & GST")

    os.makedirs("output", exist_ok=True)

    created_files = []

    for sheet in excel.sheet_names:

        if sheet == "Customer Name & GST":
            continue

        df = excel.parse(sheet)

        output_df = process_sheet(df, party_df)

        print(f"{sheet} → Rows Generated:", len(output_df))  # ✅ debug

        if not output_df.empty:
            out_file = f"output/{sheet}.xlsx"
            output_df.to_excel(out_file, index=False)
            created_files.append(out_file)

    # ZIP CREATION
    zip_path = "output.zip"
    with zipfile.ZipFile(zip_path, 'w') as z:
        for file in created_files:
            z.write(file, os.path.basename(file))

    print("✅ Done. ZIP created:", zip_path)
