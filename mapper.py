import pandas as pd

def clean(val):
    if pd.isna(val):
        return ""
    return str(val).strip()

# ✅ FIX: accept gst_df also (even if not used now)
def process_sheet(df, gst_df=None):

    rows = []

    try:
        # FIXED CELLS
        consignee_state = clean(df.iloc[14, 5])   # F15
        consignee_pincode = clean(df.iloc[15, 5]) # F16
    except:
        print("❌ Error reading F15/F16")
        return pd.DataFrame()

    print("Consignee State:", consignee_state)
    print("Consignee Pincode:", consignee_pincode)

    # LOOP
    for i in range(25, len(df)):

        try:
            desc = df.iloc[i, 1]
        except:
            continue

        if pd.isna(desc):
            continue

        desc = clean(desc)

        if desc == "":
            continue

        if desc.lower() == "end here":
            break

        # READ VALUES
        billedqty = df.iloc[i, 6]
        rate = df.iloc[i, 8]
        dis = df.iloc[i, 9]

        # SAFE CONVERT
        try:
            billedqty = float(billedqty)
        except:
            billedqty = 0

        try:
            rate = float(rate)
        except:
            rate = 0

        try:
            dis = float(dis)
        except:
            dis = 0

        # ✅ CORE LOGIC (YOUR REQUIREMENT)
        taxable = round(billedqty * rate, 2)
        amount = round(taxable - (taxable * dis / 100), 2)

        print(f"Row {i} → Taxable:{taxable}, Amount:{amount}")

        row = {
            "Description": desc,
            "Consignee State": consignee_state,
            "Consignee Pincode": consignee_pincode,
            "Billedqty": billedqty,
            "Rate": rate,
            "Dis%": dis,
            "Taxable Value": taxable,
            "Amount": amount
        }

        rows.append(row)

    print("✅ Rows created:", len(rows))

    return pd.DataFrame(rows)
