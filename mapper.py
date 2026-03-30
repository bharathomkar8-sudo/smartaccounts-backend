import pandas as pd
import os

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
    return str(val).strip()

def process_sheet(df, gst_df):

    print("👉 Sheet Shape:", df.shape)

    rows = []

    try:
        vch_no = clean(df.iloc[10,16])
        print("VCH:", vch_no)
    except Exception as e:
        print("❌ Error reading VCH:", e)
        return pd.DataFrame()

    try:
        pos = clean(df.iloc[14,5])
        consignee_state = clean(df.iloc[14,5])   # F15
        consignee_pincode = clean(df.iloc[15,5]) # F16
    except Exception as e:
        print("❌ Consignee error:", e)
        return pd.DataFrame()

    try:
        gst = clean(df.iloc[16,1])
    except:
        gst = ""

    party_name = ""
    try:
        match = gst_df[gst_df.iloc[:,0] == gst]
        if not match.empty:
            party_name = match.iloc[0,1]
        else:
            party_name = clean(df.iloc[10,0])
    except:
        party_name = clean(df.iloc[10,0])

    # LOOP
    for i in range(25, len(df)):

        desc = df.iloc[i,1]

        if pd.isna(desc):
            continue

        desc = clean(desc)

        if desc == "":
            continue

        if desc.lower() == "end here":
            break

        print(f"Row {i} → {desc}")

        row = dict.fromkeys(COLUMNS, "")

        row["Description"] = desc
        row["Party Name"] = party_name

        # Consignee fix
        row["Consignee State"] = consignee_state
        row["Consignee Pincode"] = consignee_pincode

        # Qty & Rate
        qty = df.iloc[i,5]
        rate = df.iloc[i,8]

        try:
            qty = float(qty)
        except:
            qty = 0

        try:
            rate = float(rate)
        except:
            rate = 0

        taxable = round(qty * rate, 2)

        print("   Qty:", qty, "Rate:", rate, "Taxable:", taxable)

        row["Qty"] = qty if qty else ""
        row["Rate"] = rate if rate else ""
        row["Taxable Value"] = taxable if taxable else ""
        row["Amount"] = taxable if taxable else ""

        rows.append(row)

    print("✅ Total Rows Created:", len(rows))

    return pd.DataFrame(rows)


# =========================
# RUN
# =========================
if __name__ == "__main__":

    input_file = "input.xlsx"
    gst_file = "gst.xlsx"

    print("Starting...")

    if not os.path.exists(input_file):
        print("❌ input.xlsx NOT FOUND")
        exit()

    if not os.path.exists(gst_file):
        print("❌ gst.xlsx NOT FOUND")
        exit()

    xls = pd.ExcelFile(input_file)
    gst_df = pd.read_excel(gst_file)

    os.makedirs("output", exist_ok=True)

    for sheet in xls.sheet_names:
        print("\n======================")
        print("Processing:", sheet)

        df = pd.read_excel(xls, sheet_name=sheet)

        output_df = process_sheet(df, gst_df)

        if not output_df.empty:
            path = f"output/{sheet}.xlsx"
            output_df.to_excel(path, index=False)
            print("✅ Saved:", path)
        else:
            print("⚠️ No output generated")

    print("\nDONE")
