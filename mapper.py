import pandas as pd
import os

def clean(val):
    if pd.isna(val):
        return ""
    return str(val).strip()

def process_sheet(df):

    rows = []

    # ✅ FIXED CELLS
    consignee_state = clean(df.iloc[14, 5])   # F15
    consignee_pincode = clean(df.iloc[15, 5]) # F16

    print("Consignee State:", consignee_state)
    print("Consignee Pincode:", consignee_pincode)

    # LOOP
    for i in range(25, len(df)):

        desc = df.iloc[i, 1]

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

        # ✅ CORE LOGIC
        taxable = round(billedqty * rate, 2)
        amount = round(taxable - (taxable * dis / 100), 2)

        print(f"Row {i} → BQty:{billedqty}, Rate:{rate}, Dis:{dis}")
        print(f"   Taxable:{taxable}, Amount:{amount}")

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

    return pd.DataFrame(rows)


# =========================
# RUN TEST
# =========================
if __name__ == "__main__":

    file = "input.xlsx"

    if not os.path.exists(file):
        print("❌ input.xlsx not found")
        exit()

    xls = pd.ExcelFile(file)

    os.makedirs("output", exist_ok=True)

    for sheet in xls.sheet_names:
        print("\nProcessing:", sheet)

        df = pd.read_excel(xls, sheet_name=sheet)

        out = process_sheet(df)

        if not out.empty:
            path = f"output/{sheet}.xlsx"
            out.to_excel(path, index=False)
            print("✅ Saved:", path)
        else:
            print("⚠️ No data")

    print("\nDONE")
