import pandas as pd

# =========================
# GLOBAL GST DATA
# =========================
party_df = None

# =========================
# LOAD EXCEL + GST
# =========================
def load_excel(file):
    global party_df
    excel_file = pd.ExcelFile(file)

    # Load GST sheet
    if "GST" in excel_file.sheet_names:
        party_df = excel_file.parse("GST")
    else:
        party_df = None

    return excel_file

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
# DATE FORMAT
# =========================
def format_date(val):
    try:
        return pd.to_datetime(val).strftime("%d-%m-%Y")
    except:
        return ""

# =========================
# PROCESS SINGLE SHEET
# =========================
def process_sheet(df):

    global party_df

    # 🚫 Skip small sheets
    if len(df) < 20:
        return None

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
    except:
        return None

    # =========================
    # PARTY NAME (VLOOKUP STYLE)
    # =========================
    party_name = ""

    if party_df is not None:
        match = party_df[
            party_df.iloc[:, 0].astype(str).str.strip() == str(gst).strip()
        ]
        if not match.empty:
            party_name = match.iloc[0, 1]

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

    rows = []

    # =========================
    # LOOP ITEMS
    # =========================
    for i in range(25, len(df)):

        desc = df.iloc[i, 1]

        if pd.isna(desc):
            continue

        desc = clean(desc)

        if desc == "" or desc.lower() == "end here":
            continue

        row = {}

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
        row["Dis%"] = safe(df.iloc[i, 9])

        try:
            qty = float(row["Qty"]) if row["Qty"] else 0
            rate = float(row["Rate"]) if row["Rate"] else 0
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


# =========================
# PROCESS ALL SHEETS
# =========================
def process_file(file):

    excel_file = load_excel(file)
    final_data = []

    for sheet in excel_file.sheet_names:

        # 🚫 Skip GST & unwanted sheets
        if sheet.strip().lower() in ["gst", "customer name & gst", "sheet3"]:
            continue

        df = excel_file.parse(sheet)

        if len(df) < 20:
            continue

        result = process_sheet(df)

        if result is not None and not result.empty:
            final_data.append(result)

    if final_data:
        return pd.concat(final_data, ignore_index=True)
    else:
        return pd.DataFrame()
