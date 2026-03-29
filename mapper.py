import pandas as pd

COLUMNS = [
"Voucher Type","VCH No / Inv No","Description","VCH Date","Order No","Order Date","Other Ref","POS",
"Party Name","Address","State","Pincode","Party GSTIN",
"Consignee Name","Con Address","Consignee State","Consignee Pincode","Con GSTIN",
"Item Name / Code","Is Item Header","width","Height","Qty","Extraudf","Billedqty","Rate",
"Taxable Value","Dis%","Amount","Sales Ledger",
"CGST Ledger","CGST Amt","SGST Ledger","SGST Amount",
"IGST Ledger","IGST Amount","Round off","Invoice Amt","Item header"
]

def process_sheet(df):

    # =========================
    # HEADER
    # =========================
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

    # =========================
    # DESCRIPTION
    # =========================
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
    row["Party GSTIN"] = gstin
    row["Consignee Name"] = party_name
    row["Con Address"] = addr1
    row["Con GSTIN"] = gstin
    row["Item Name / Code"] = "Header"
    row["Is Item Header"] = "Yes"
    row["Item header"] = descriptions[0]
    rows.append(row)

    # =========================
    # ADDRESS CONTINUE
    # =========================
    row = dict.fromkeys(COLUMNS, "")
    row["Address"] = addr2
    row["Con Address"] = addr2
    rows.append(row)

    row = dict.fromkeys(COLUMNS, "")
    row["Address"] = addr3
    row["Con Address"] = addr3
    rows.append(row)

    # =========================
    # ITEM ROWS
    # =========================
    for desc in descriptions[1:]:

        row = dict.fromkeys(COLUMNS, "")

        row["Description"] = desc
        row["Item Name / Code"] = desc
        row["Item header"] = desc

        rows.append(row)

    return pd.DataFrame(rows, columns=COLUMNS)
