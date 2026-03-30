import pandas as pd

party_df = None

# =========================
# LOAD EXCEL
# =========================
def load_excel(file):
    global party_df
    excel_file = pd.ExcelFile(file)

    if "GST" in excel_file.sheet_names:
        party_df = excel_file.parse("GST")
    else:
        party_df = None

    return excel_file

# =========================
# CLEAN
# =========================
def clean(val):
    if pd.isna(val):
        return ""
    val = str(val).strip()
    if val.lower() == "nan":
        return ""
    return val

def format_date(val):
    try:
        return pd.to_datetime(val).strftime("%d-%m-%Y")
    except:
        return ""

# =========================
# PROCESS ONE SHEET
# =========================
def process_sheet(df):

    global party_df

    if len(df) < 20:
        return None

    try:
        gst = clean(df.iloc[16, 1])  # B17
    except:
        return None

    # =========================
    # PARTY NAME (FIXED)
    # =========================
    party_name = ""

    if party_df is not None:

        # 🔥 DEBUG (remove later)
        print("GST FROM MAIN:", repr(gst))
        print("GST SHEET SAMPLE:", party_df.iloc[:5, 0].tolist())

        match = party_df[
            party_df.iloc[:, 0].astype(str).str.strip() == str(gst).strip()
        ]

        if not match.empty:
            party_name = match.iloc[0, 1]

    # =========================
    # BASIC HEADER (SAFE)
    # =========================
    try:
        vch_no = clean(df.iloc[10, 16])
        vch_date = format_date(df.iloc[11, 16])
    except:
        return None

    rows = []

    for i in range(25, len(df)):

        desc = df.iloc[i, 1]

        if pd.isna(desc):
            continue

        desc = clean(desc)

        if desc == "":
            continue

        row = {}

        row["VCH No / Inv No"] = vch_no
        row["VCH Date"] = vch_date
        row["Party GSTIN"] = gst
        row["Party Name"] = party_name
        row["Description"] = desc

        rows.append(row)

    return pd.DataFrame(rows)

# =========================
# PROCESS FILE
# =========================
def process_file(file):

    excel_file = load_excel(file)
    final_data = []

    for sheet in excel_file.sheet_names:

        # 🚫 SKIP GST SHEET
        if sheet.strip().lower() == "gst":
            continue

        df = excel_file.parse(sheet)

        if len(df) < 20:
            continue

        result = process_sheet(df)

        if result is not None and not result.empty:
            final_data.append(result)

    # ✅ IMPORTANT FIX
    if final_data:
        return pd.concat(final_data, ignore_index=True)
    else:
        return pd.DataFrame()   # ← NEVER return None
