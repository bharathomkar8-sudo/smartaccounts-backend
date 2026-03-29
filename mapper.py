import pandas as pd

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
# MAIN FUNCTION (IMPORTANT)
# =========================
def process_sheet(df):

    rows = []

    # =========================
    # HEADER DATA
    # =========================
    voucher = str(df.iloc[1, 0])
    inv_no = str(df.iloc[1, 1])
    main_desc = str(df.iloc[1, 2])

    party_name = str(df.iloc[1, 8])
    gst = str(df.iloc[1, 12])

    # ADDRESS (FIXED - NO DUPLICATE)
    addr1 = str(df.iloc[1, 9])
    addr2 = str(df.iloc[2, 9]) if pd.notna(df.iloc[2, 9]) else ""
    addr3 = str(df.iloc[3, 9]) if pd.notna(df.iloc[3, 9]) else ""

    full_address = ", ".join([x for x in [addr1, addr2, addr3] if x and x != "nan"])

    state = str(df.iloc[1, 10])
    pincode = str(df.iloc[1, 11])

    # =========================
    # MAIN HEADER ROW
    # =========================
    header = dict.fromkeys(COLUMNS, "")

    header["Voucher Type"] = voucher
    header["VCH No / Inv No"] = inv_no
    header["Description"] = main_desc

    header["Party Name"] = party_name
    header["Address"] = full_address
    header["State"] = state
    header["Pincode"] = pincode
    header["Party GSTIN"] = gst

    header["Consignee Name"] = party_name
    header["Con Address"] = full_address
    header["Con GSTIN"] = gst

    header["Item Name / Code"] = "Header"
    header["Is Item Header"] = "Yes"
    header["Item header"] = main_desc

    rows.append(header)

    # =========================
    # DESCRIPTION (B26 onwards)
    # =========================
    descriptions = []

    for i in range(5, len(df)):
        val = df.iloc[i, 2]
        if pd.isna(val):
            break
        descriptions.append(str(val))

    # =========================
    # LOOP ITEMS
    # =========================
    for desc in descriptions:

        row = dict.fromkeys(COLUMNS, "")

        # HEADER TYPE LINE
        if len(desc.strip()) < 6:
            row["Description"] = desc
            row["Item Name / Code"] = "Header"
            row["Is Item Header"] = "Yes"
            row["Item header"] = desc
            rows.append(row)
            continue

        # NORMAL ITEM
        qty = 10
        rate = 500

        taxable = qty * rate
        gst_amt = round(taxable * 0.18, 2)
        total = taxable + gst_amt

        row["Description"] = desc
        row["Item Name / Code"] = "391990"

        row["Qty"] = qty
        row["Billedqty"] = qty
        row["Rate"] = rate

        row["Taxable Value"] = taxable
        row["Amount"] = taxable

        row["Sales Ledger"] = "GST IGST Sales@18%"
        row["IGST Ledger"] = "OUTPUT IGST @ 18%"
        row["IGST Amount"] = gst_amt

        row["Invoice Amt"] = total
        row["Item header"] = desc

        rows.append(row)

    return pd.DataFrame(rows)
