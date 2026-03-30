import pandas as pd

def clean(val):
    if pd.isna(val):
        return ""
    return str(val).strip()

# =========================
# GET EXACT COLUMN BY EXCEL LETTER
# =========================
def get_excel_col(df, col_letter):
    col_index = ord(col_letter.upper()) - ord('A')
    return df.iloc[:, col_index]

# =========================
# MAIN FUNCTION
# =========================
def process_sheet(df, gst_df=None):

    rows = []

    try:
        # ✅ FORCE COLUMN F
        col_F = get_excel_col(df, 'F')

        consignee_state = clean(col_F.iloc[14])   # F15
        consignee_pincode = clean(col_F.iloc[15]) # F16

    except Exception as e:
        print("❌ Consignee read error:", e)
        consignee_state = ""
        consignee_pincode = ""

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

        # SAFE
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

        # ✅ CALC
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
