import pandas as pd

# =========================
# PROCESS SHEET
# =========================
def process_sheet(df):

    # Skip small sheets
    if len(df) < 20:
        return None

    # =========================
    # SAFE READ GST
    # =========================
    try:
        gst = str(df.iloc[16, 1])   # B17
    except:
        return None

    # =========================
    # PARTY NAME TEST (VLOOKUP STYLE)
    # =========================
    party_name = ""

    try:
        gst_df = pd.read_excel("testing.xlsx", sheet_name="GST")

        match = gst_df[gst_df.iloc[:, 0] == gst]

        if not match.empty:
            party_name = match.iloc[0, 1]
        else:
            party_name = ""

    except Exception as e:
        print("ERROR IN PARTY LOOKUP:", e)
        party_name = ""

    # =========================
    # DEBUG PRINT
    # =========================
    print("GST:", gst)
    print("PARTY NAME:", party_name)

    # =========================
    # MINIMUM OUTPUT (TEST)
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

        row["Party GSTIN"] = gst
        row["Party Name"] = party_name
        row["Description"] = desc

        rows.append(row)

    return pd.DataFrame(rows)
