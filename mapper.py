import pandas as pd

def process_sheet(df, gst_df):

    # Skip small sheets
    if len(df) < 20:
        return None

    # =========================
    # SAFE HEADER READ
    # =========================
    try:
        voucher_type = "Sales E-Invoice"
        vch_no = str(df.iloc[10, 16])
        vch_date = df.iloc[11, 16]
        gst = str(df.iloc[16, 1])
    except:
        return None

    # =========================
    # PARTY NAME (VLOOKUP)
    # =========================
    party_name = ""
    gst_error = ""

    try:
        match = gst_df[gst_df.iloc[:, 0] == gst]

        if not match.empty:
            party_name = match.iloc[0, 1]
        else:
            party_name = ""
            gst_error = "GST NOT FOUND"

    except Exception as e:
        print("LOOKUP ERROR:", e)
        gst_error = "GST NOT FOUND"

    print("GST:", gst)
    print("PARTY NAME:", party_name)

    # =========================
    # LOOP ITEMS
    # =========================
    rows = []

    for i in range(25, len(df)):

        desc = df.iloc[i, 1]

        if pd.isna(desc):
            continue

        desc = str(desc)

        if desc == "":
            continue

        row = {}

        # =========================
        # BASIC FIELDS
        # =========================
        row["Voucher Type"] = voucher_type

        # 👉 If GST missing → modify invoice no
        if gst_error:
            row["VCH No / Inv No"] = str(vch_no) + " - GST NOT FOUND"
        else:
            row["VCH No / Inv No"] = vch_no

        row["VCH Date"] = vch_date
        row["Party GSTIN"] = gst

        # =========================
        # PARTY + CONSIGNEE
        # =========================
        row["Party Name"] = party_name
        row["Consignee Name"] = party_name   # ✅ SAME AS PARTY

        # =========================
        # DESCRIPTION
        # =========================
        row["Description"] = desc

        # =========================
        # ERROR COLUMN
        # =========================
        row["Error"] = gst_error

        rows.append(row)

    return pd.DataFrame(rows)
