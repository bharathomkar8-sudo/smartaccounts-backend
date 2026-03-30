import pandas as pd

def clean(val):
    if pd.isna(val):
        return ""
    return str(val).strip()

# =========================
# GET VALUE NEXT TO LABEL
# =========================
consignee_state = get_value_next_to_label_right(df, 14, "state")
consignee_pincode = get_value_next_to_label_right(df, 15, "pincode")
# =========================
# MAIN FUNCTION
# =========================
def process_sheet(df, gst_df=None):

    rows = []

    # ✅ CONSIGNEE FIX (FINAL)
    consignee_state = get_value_next_to_label(df, 14, "state")
    consignee_pincode = get_value_next_to_label(df, 15, "pincode")

    print("Consignee State:", consignee_state)
    print("Consignee Pincode:", consignee_pincode)

    # =========================
    # LOOP
    # =========================
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

        # ✅ CALCULATION
        taxable = round(billedqty * rate, 2)
        amount = round(taxable - (taxable * dis / 100), 2)

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
