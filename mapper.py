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
# CLEAN FUNCTION
# =========================
def clean(val):
    if pd.isna(val):
        return ""
    val = str(val).strip()
    if val.lower() == "nan":
        return ""
    return val

# =========================
# DATE FORMAT FUNCTION
# =========================
def format_date(val):
    try:
        return pd.to_datetime(val).strftime("%d-%m-%Y")
    except:
        return ""

# =========================
# MAIN FUNCTION
# =========================
def process_sheet(df, party_df):   # ✅ added party_df

    rows = []

    # =========================
    # PARTY LOOKUP DICTIONARY (NEW)
    # =========================
    party_map = {}

    for i in range(1, len(party_df)):
        gst_key = clean(party_df.iloc[i, 0])
        name_val = clean(party_df.iloc[i, 1])

        if gst_key != "":
            party_map[gst_key] = name_val

    # =========================
    # HEADER VALUES
    # =========================
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

    # =========================
    # PARTY NAME LOGIC (NEW)
    # =========================
    gst_key = gst.strip()

    if gst_key in party_map:
        party_name = party_map[gst_key]
    else:
        party_name = "NOT FOUND"

    # =========================
    # ADDRESS
    # =========================
    address_lines = [
        clean(df.iloc[11, 0]),
        clean(df.iloc[12, 0]),
        clean(df.iloc[13, 0])
    ]

    con_address_lines = [
        clean(df.iloc[12, 4]),
        clean(df.iloc[13, 4])
    ]

    # =========================
    # LOOP
    # =========================
    for i in range(25, len(df)):

        desc = df.iloc[i, 1]

        if pd.isna(desc):
            continue

        desc = clean(desc)

        if desc == "":
            continue

        if desc.lower() == "end here":
            break

        row = dict.fromkeys(COLUMNS, "")

        # =========================
        # HEADER FILL
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

        # ✅ NEW PARTY + CONSIGNEE
        row["Party Name"] = party_name
        row["Consignee Name"] = party_name

        row["Consignee State"] = pos
        row["Consignee Pincode"] = clean(df.iloc[15, 5])
        row["Con GSTIN"] = gst

        # =========================
        # DESCRIPTION
        # =========================
        row["Description"] = desc
        row["Item header"] = desc

        # =========================
        # ADDRESS FLOW
        # =========================
        if i - 25 < len(address_lines):
            row["Address"] = address_lines[i - 25]

        if i - 25 < len(con_address_lines):
            row["Con Address"] = con_address_lines[i - 25]

        # =========================
        # ITEM CODE
        # =========================
        item_val = df.iloc[i, 2]

        if isinstance(item_val, (int, float)) and item_val > 1:
            row["Item Name / Code"] = str(int(item_val))
        else:
            row["Item Name / Code"] = "Header"

        if row["Item Name / Code"] == "Header":
            row["Is Item Header"] = "Yes"

        # =========================
        # SAFE
        # =========================
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
        # COLUMN MAPPING
        # =========================
        row["width"] = safe(df.iloc[i, 3])
        row["Height"] = safe(df.iloc[i, 4])
        row["Qty"] = safe(df.iloc[i, 5])
        row["Extraudf"] = safe(df.iloc[i, 6])
        row["Billedqty"] = safe(df.iloc[i, 6])
        row["Rate"] = safe(df.iloc[i, 8])

        row["Dis%"] = safe(df.iloc[i, 9])

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
