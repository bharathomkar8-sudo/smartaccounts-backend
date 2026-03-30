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
# MAIN FUNCTION
# =========================
def process_sheet(df):

    rows = []

    # =========================
    # HEADER VALUES (FIXED CELLS)
    # =========================
    voucher_type = "Sales E-Invoice"
    vch_no = str(df.iloc[10, 16])     # Q11
    vch_date = df.iloc[11, 16]        # Q12
    order_no = str(df.iloc[19, 1])    # B20
    order_date = df.iloc[20, 1]       # B21
    other_ref = str(df.iloc[12, 16])  # Q13
    pos = str(df.iloc[14, 5])         # F15

    state = str(df.iloc[14, 1])       # B15
    pincode = str(df.iloc[15, 1])     # B16
    gst = str(df.iloc[16, 1])         # B17

    # ADDRESS (LINE BY LINE — NOT COMBINED)
    address_lines = [
        str(df.iloc[11, 0]),  # A12
        str(df.iloc[12, 0]),  # A13
        str(df.iloc[13, 0])   # A14
    ]

    # CONSIGNEE ADDRESS (LINE BY LINE)
    con_address_lines = [
        str(df.iloc[12, 4]),  # E13
        str(df.iloc[13, 4])   # E14
    ]

    # =========================
    # LOOP FROM ROW 26
    # =========================
    for i in range(25, len(df)):

        desc = df.iloc[i, 1]   # B column

        # STOP CONDITION
        if pd.isna(desc):
            continue

        desc = str(desc).strip()

        if desc.lower() == "end here":
            break

        row = dict.fromkeys(COLUMNS, "")

        # =========================
        # BASIC HEADER FILL
        # =========================
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

        row["Consignee State"] = pos
        row["Consignee Pincode"] = str(df.iloc[15, 5])  # F16
        row["Con GSTIN"] = gst

        # =========================
        # DESCRIPTION
        # =========================
        row["Description"] = desc
        row["Item header"] = desc

        # =========================
        # ADDRESS FLOW (LINE BY LINE)
        # =========================
        if i - 25 < len(address_lines):
            row["Address"] = address_lines[i - 25]

        if i - 25 < len(con_address_lines):
            row["Con Address"] = con_address_lines[i - 25]

        # =========================
        # ITEM CODE LOGIC (COLUMN C)
        # =========================
        item_val = df.iloc[i, 2]

        if isinstance(item_val, (int, float)) and item_val > 1:
            row["Item Name / Code"] = str(int(item_val))
        else:
            row["Item Name / Code"] = "Header"

        # IS HEADER
        if row["Item Name / Code"] == "Header":
            row["Is Item Header"] = "Yes"

        # =========================
        # SAFE VALUE FUNCTION
        # =========================
        def safe(val):
            if pd.isna(val) or val == 0:
                return ""
            return val

        # =========================
        # COLUMN MAPPING
        # =========================
        row["width"] = safe(df.iloc[i, 3])      # D
        row["Height"] = safe(df.iloc[i, 4])     # E
        row["Qty"] = safe(df.iloc[i, 5])        # F
        row["Extraudf"] = safe(df.iloc[i, 6])   # G
        row["Billedqty"] = safe(df.iloc[i, 6])  # G
        row["Rate"] = safe(df.iloc[i, 8])       # I

        # =========================
        # CALCULATION
        # =========================
        try:
            qty = float(row["Qty"]) if row["Qty"] != "" else 0
            rate = float(row["Rate"]) if row["Rate"] != "" else 0
        except:
            qty, rate = 0, 0

        taxable = qty * rate
        gst_amt = round(taxable * 0.18, 2)
        total = taxable + gst_amt

        row["Taxable Value"] = taxable if taxable else ""
        row["Amount"] = taxable if taxable else ""

        row["Sales Ledger"] = "GST IGST Sales@18%"
        row["IGST Ledger"] = "OUTPUT IGST @ 18%"
        row["IGST Amount"] = gst_amt if gst_amt else ""

        row["Invoice Amt"] = total if total else ""

        rows.append(row)

    return pd.DataFrame(rows)
