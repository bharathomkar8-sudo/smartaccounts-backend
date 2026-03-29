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

def process_sheet(df):

    rows = []

    # =========================
    # HEADER MAPPING (AS PER YOUR EXCEL)
    # =========================
    header = dict.fromkeys(COLUMNS, "")

    header["Voucher Type"] = "Sales E-Invoice"
    header["VCH No / Inv No"] = str(df.iloc[10, 16])   # Q11

    header["VCH Date"] = str(df.iloc[11, 16])          # Q12
    header["Order No"] = str(df.iloc[19, 1])           # B20
    header["Order Date"] = str(df.iloc[20, 1])         # B21
    header["Other Ref"] = str(df.iloc[12, 16])         # Q13
    header["POS"] = str(df.iloc[14, 5])                # F15

    # ADDRESS (A12:A14 → combined)
    addr1 = str(df.iloc[11, 0])
    addr2 = str(df.iloc[12, 0])
    addr3 = str(df.iloc[13, 0])
    full_address = ", ".join([x for x in [addr1, addr2, addr3] if x and x != "nan"])

    header["Address"] = full_address

    header["State"] = str(df.iloc[14, 1])              # B15
    header["Pincode"] = str(df.iloc[15, 1])            # B16
    header["Party GSTIN"] = str(df.iloc[16, 1])        # B17

    # CONSIGNEE
    con1 = str(df.iloc[12, 4])
    con2 = str(df.iloc[13, 4])
    con_address = ", ".join([x for x in [con1, con2] if x and x != "nan"])

    header["Consignee Name"] = header["Party Name"]
    header["Con Address"] = con_address
    header["Consignee State"] = str(df.iloc[14, 5])
    header["Consignee Pincode"] = str(df.iloc[15, 5])
    header["Con GSTIN"] = str(df.iloc[16, 1])

    header["Item Name / Code"] = "Header"
    header["Is Item Header"] = "Yes"
    header["Item header"] = str(df.iloc[25, 2])

    rows.append(header)

    # =========================
    # ITEM MAPPING (START B26)
    # =========================
    for i in range(25, len(df)):

        desc = df.iloc[i, 2]

        if pd.isna(desc):
            break

        desc = str(desc).strip()

        row = dict.fromkeys(COLUMNS, "")

        # =========================
        # ITEM NAME / HEADER LOGIC
        # =========================
        try:
            item_code = df.iloc[i, 2]
            if isinstance(item_code, (int, float)) and item_code > 1:
                row["Item Name / Code"] = str(int(item_code))
            else:
                row["Item Name / Code"] = "Header"
        except:
            row["Item Name / Code"] = "Header"

        if row["Item Name / Code"] == "Header":
            row["Is Item Header"] = "Yes"

        # =========================
        # BASIC
        # =========================
        row["Description"] = desc
        row["Item header"] = desc

        # =========================
        # VALUE MAPPING (AS PER YOUR FORMULA)
        # =========================
        def safe(val):
            return "" if pd.isna(val) or val == 0 else val

        row["width"] = safe(df.iloc[i, 3])     # D
        row["Height"] = safe(df.iloc[i, 4])    # E
        row["Qty"] = safe(df.iloc[i, 5])       # F
        row["Extraudf"] = safe(df.iloc[i, 6])  # G
        row["Billedqty"] = safe(df.iloc[i, 6]) # G
        row["Rate"] = safe(df.iloc[i, 8])      # I

        # =========================
        # CALCULATION
        # =========================
        try:
            qty = float(row["Qty"]) if row["Qty"] != "" else 0
            rate = float(row["Rate"]) if row["Rate"] != "" else 0
        except:
            qty, rate = 0, 0

        taxable = qty * rate
        gst = round(taxable * 0.18, 2)
        total = taxable + gst

        row["Taxable Value"] = taxable if taxable else ""
        row["Amount"] = taxable if taxable else ""

        row["Sales Ledger"] = "GST IGST Sales@18%"
        row["IGST Ledger"] = "OUTPUT IGST @ 18%"
        row["IGST Amount"] = gst if gst else ""

        row["Invoice Amt"] = total if total else ""

        rows.append(row)

    return pd.DataFrame(rows)
