import pandas as pd

# =========================
# PROCESS SHEET
# =========================
def process_sheet(df, gst_df):

    # Skip small sheets
    if len(df) < 20:
        return None

    # =========================
    # READ GST
    # =========================
    try:
        gst = str(df.iloc[16, 1])   # B17
    except:
        return None

    # =========================
    # PARTY NAME LOGIC
    # =========================
    party_name = ""

    try:
        match = gst_df[gst_df.iloc[:, 0] == gst]

        if not match.empty:
            # ✅ GST FOUND
            party_name = match.iloc[0, 1]
        else:
            # ❌ GST NOT FOUND → A11
            party_name = str(df.iloc[10, 0])  # A11

    except:
        party_name = str(df.iloc[10, 0])

    # =========================
    # DEBUG (optional)
    # =========================
    print("GST:", gst)
    print("PARTY NAME:", party_name)

    # =========================
    # BUILD OUTPUT
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

        # BASIC FIELDS
        row["Party GSTIN"] = gst
        row["Party Name"] = party_name

        # ✅ CONSIGNEE SAME AS PARTY
        row["Consignee Name"] = party_name

        row["Description"] = desc

        rows.append(row)

    return pd.DataFrame(rows)
