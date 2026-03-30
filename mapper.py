# =========================
# PARTY NAME TEST (VLOOKUP STYLE ONLY)
# =========================

party_name = ""

try:
    # Load GST sheet directly (temporary test)
    gst_df = pd.read_excel("testing.xlsx", sheet_name="GST")

    # Exact VLOOKUP match
    match = gst_df[gst_df.iloc[:, 0] == gst]

    if not match.empty:
        party_name = match.iloc[0, 1]
    else:
        party_name = ""

except Exception as e:
    print("ERROR IN PARTY LOOKUP:", e)
    party_name = ""

print("GST:", gst)
print("PARTY NAME:", party_name)
