import pandas as pd
import os
import zipfile

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
    return "" if val.lower() == "nan" else val

def format_date(val):
    try:
        return pd.to_datetime(val).strftime("%d-%m-%Y")
    except:
        return ""

def safe(val):
    if pd.isna(val):
        return ""
    try:
        if float(val) == 0:
            return ""
    except:
        pass
    return val


# =========================
# MAIN PROCESS FUNCTION
# =========================
def process_sheet(df, party_df):

    rows = []

    try:
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

    except Exception as e:
        print("❌ Header read error:", e)
        return pd.DataFrame()

    # PARTY LOOKUP
    party_name = ""
    if gst:
        match = party_df[party_df.iloc[:,0].astype(str).str.strip() == gst]
        if not match.empty:
            party_name = clean(match.iloc[0,1])
        else:
            vch_no += "-ERROR"

    consignee_name = party_name

    # =========================
    # LOOP (FIXED + SAFE)
    # =========================
    for i in range(25, len(df)):

        try:
            desc = clean(df.iloc[i, 1])
        except:
            continue

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

        # ITEM
        item_val = df.iloc[i, 2]
        if isinstance(item_val, (int, float)) and item_val > 1:
            row["Item Name / Code"] = str(int(item_val))
        else:
            row["Item Name / Code"] = "Header"
            row["Is Item Header"] = "Yes"

        # VALUES
        row["width"] = safe(df.iloc[i, 3])
        row["Height"] = safe(df.iloc[i, 4])
        row["Qty"] = safe(df.iloc[i, 5])
        row["Extraudf"] = safe(df.iloc[i, 6])
        row["Billedqty"] = safe(df.iloc[i, 6])
        row["Rate"] = safe(df.iloc[i, 8])

        # CALCULATION
        try:
            qty = float(row["Qty"]) if row["Qty"] else 0
            rate = float(row["Rate"]) if row["Rate"] else 0
        except:
            qty, rate = 0, 0

        taxable = qty * rate

        try:
            dis_percent = float(df.iloc[i, 9])
        except:
            dis_percent = ""

        amount = taxable * (1 - dis_percent/100) if dis_percent != "" else taxable
        gst_amt = round(amount * 0.18, 2)
        total = amount + gst_amt

        row["Taxable Value"] = taxable
        row["Dis%"] = dis_percent
        row["Amount"] = amount

        row["Sales Ledger"] = "GST IGST Sales@18%"
        row["IGST Ledger"] = "OUTPUT IGST @ 18%"
        row["IGST Amount"] = gst_amt
        row["Invoice Amt"] = total

        rows.append(row)

    print("👉 Rows created:", len(rows))
    return pd.DataFrame(rows)


# =========================
# RUN FUNCTION
# =========================
def run_full_process(file_path):

    excel = pd.ExcelFile(file_path)

    if "Customer Name & GST" not in excel.sheet_names:
        print("❌ Sheet 'Customer Name & GST' missing")
        return

    party_df = excel.parse("Customer Name & GST")

    os.makedirs("output", exist_ok=True)

    created_files = []

    for sheet in excel.sheet_names:

        if sheet == "Customer Name & GST":
            continue

        print(f"\nProcessing Sheet: {sheet}")

        df = excel.parse(sheet)

        output_df = process_sheet(df, party_df)

        if len(output_df) == 0:
            print("⚠️ No data → skipping file")
            continue

        out_file = f"output/{sheet}.xlsx"
        output_df.to_excel(out_file, index=False)
        created_files.append(out_file)

        print("✅ Excel created:", out_file)

    # ZIP
    if created_files:
        with zipfile.ZipFile("output.zip", 'w') as z:
            for file in created_files:
                z.write(file, os.path.basename(file))
        print("\n✅ ZIP CREATED")
    else:
        print("\n❌ No files to zip")
