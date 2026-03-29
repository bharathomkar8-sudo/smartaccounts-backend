import pandas as pd

# =========================
# EXACT COLUMN STRUCTURE
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
# MAIN MAPPING FUNCTION
# =========================
def process_sheet(df):

    # =========================
    # HEADER DATA
    # =========================
    vch_no = df.iloc[1, 0]
    vch_date = df.iloc[1, 3]
    order_no = df.iloc[1, 4]
    order_date = df.iloc[1, 5]
    other_ref = df.iloc[1, 6]

    party_name = df.iloc[10, 0]
    gstin = df.iloc[10, 1]

    # ADDRESS (STRICT ROW FORMAT)
    addr1 = str(df.iloc[11, 0]).strip()
    addr2 = str(df.iloc[12, 0]).strip()
    addr3 = str(df.iloc[13, 0]).strip()

    # STATE & PINCODE
    state = ""
    pincode = ""

    if "–" in addr3:
        parts = addr3.split("–")
        if len(parts) == 2:
            pincode = parts[1].strip()
            state = parts[0].split(",")[-1].strip()

    # =========================
    # DESCRIPTION (B26 → STOP GST)
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
    # ROW 1 (MAIN HEADER)
    # =========================
    row = dict.fromkeys(COLUMNS, "")

    row["Voucher Type"] = "Sales E-Invoice"
    row["VCH No / Inv No"] = vch_no
    row["Description"] = descriptions[0] if descriptions else ""
    row["VCH Date"] = vch_date
    row["Order No"] = order_no
    row["Order Date"] = order_date
    row["Other Ref"] = other_ref
    row["Party Name"] = party_name

    row["Address"] = addr1
    row["State"] = state
    row["Pincode"] = pincode
    row["Party GSTIN"] = gstin

    row["Consignee Name"] = party_name
    row["Con Address"] = addr1
    row["Consignee State"] = state
    row["Consignee Pincode"] = pincode
    row["Con GSTIN"] = gstin

    row["Item Name / Code"] = "Header"
    row["Is Item Header"] = "Yes"
    row["Item header"] = descriptions[0] if descriptions else ""

    rows.append(row)

    # =========================
    # ROW 2 (ADDRESS LINE 2)
    # =========================
    row = dict.fromkeys(COLUMNS, "")
    row["Description"] = descriptions[1] if len(descriptions) > 1 else ""
    row["Address"] = addr2
    row["Con Address"] = addr2
    row["Item Name / Code"] = "Header"
    row["Is Item Header"] = "Yes"
    row["Item header"] = descriptions[1] if len(descriptions) > 1 else ""
    rows.append(row)

    # =========================
    # ROW 3 (ADDRESS LINE 3)
    # =========================
    row = dict.fromkeys(COLUMNS, "")
    row["Address"] = addr3
    row["Con Address"] = addr3
    rows.append(row)

  # =========================
# ITEM ROWS (FINAL LOGIC)
# =========================
for desc in descriptions[2:]:

    row = dict.fromkeys(COLUMNS, "")

    # BASIC
    row["Description"] = desc
    row["Item Name / Code"] = "391990"   # TODO: later dynamic
    row["Item header"] = desc

    # DEFAULT INPUT (can later read from Excel)
    qty = 10
    rate = 500

    taxable = qty * rate
    gst = round(taxable * 0.18, 2)
    total = taxable + gst

    # FILL VALUES
    row["Qty"] = qty
    row["Billedqty"] = qty
    row["Rate"] = rate

    row["Taxable Value"] = taxable
    row["Amount"] = taxable

    # GST LOGIC (IGST DEFAULT)
    row["Sales Ledger"] = "GST IGST Sales@18%"
    row["IGST Ledger"] = "OUTPUT IGST @ 18%"
    row["IGST Amount"] = gst

    row["Invoice Amt"] = total

    rows.append(row)
    return pd.DataFrame(rows, columns=COLUMNS)
