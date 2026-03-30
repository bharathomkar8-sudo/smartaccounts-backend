import pandas as pd

def process_sheet(df, gst_df):

    if len(df) < 20:
        return None

    # =========================
    # GST READ
    # =========================
    try:
        gst = str(df.iloc[16, 1])
    except:
        return None

    # =========================
    # PARTY NAME (PURE VLOOKUP)
    # =========================
    party_name = ""

    try:
        match = gst_df[gst_df.iloc[:, 0] == gst]

        if not match.empty:
            party_name = match.iloc[0, 1]

    except Exception as e:
        print("LOOKUP ERROR:", e)

    print("GST:", gst)
    print("PARTY NAME:", party_name)

    # =========================
    # OUTPUT
    # =========================
    rows = []

    for i in range(25, len(df)):
        desc = df.iloc[i, 1]

        if pd.isna(desc):
            continue

        row = {
            "Party GSTIN": gst,
            "Party Name": party_name,
            "Description": str(desc)
        }

        rows.append(row)

    return pd.DataFrame(rows)
