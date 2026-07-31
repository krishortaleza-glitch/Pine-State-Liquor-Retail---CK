import pandas as pd

# -----------------------------
# Load Files
# -----------------------------
vendor = pd.read_csv("EGStoreCostFile.csv", dtype=str)
master = pd.read_excel("August 2026 Master Price List (2)(2).xlsx", dtype=str)

# -----------------------------
# Clean Function
# -----------------------------
def clean_uid(x):
    if pd.isna(x):
        return ""

    x = str(x).strip().upper()

    if x.endswith(".0"):
        x = x[:-2]

    x = x.lstrip("0")

    return x

vendor["CleanUPC"] = vendor["vendorProductUID"].apply(clean_uid)
master["CleanUPC"] = master["UPC"].apply(clean_uid)

# -----------------------------
# Validation
# -----------------------------
zero_matches = 0
one_match = 0
multiple_matches = 0

examples = []

for upc in vendor["CleanUPC"]:

    matches = master[
        master["CleanUPC"].str.contains(upc, regex=False, na=False)
    ]

    if len(matches) == 0:
        zero_matches += 1

    elif len(matches) == 1:
        one_match += 1

    else:
        multiple_matches += 1

        if len(examples) < 20:
            examples.append(
                {
                    "VendorUPC": upc,
                    "Matches": len(matches),
                    "MasterUPCs": matches["CleanUPC"].tolist()
                }
            )

print(f"Zero Matches      : {zero_matches:,}")
print(f"One Match         : {one_match:,}")
print(f"Multiple Matches  : {multiple_matches:,}")

print("\nExamples of Multiple Matches")
for e in examples:
    print(e)
