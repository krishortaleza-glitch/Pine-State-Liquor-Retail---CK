import streamlit as st
import pandas as pd
from io import BytesIO

# ==========================================================
# Page Configuration
# ==========================================================

st.set_page_config(
    page_title="Pine State Liquor Retail Builder",
    layout="wide"
)

st.title("Pine State Liquor Retail Builder")
st.markdown(
    """
Generate **Standard** and **Promo** retail files from:

- Raw Vendor Store Cost File
- Pine State Liquor Master Price List
"""
)

st.divider()

# ==========================================================
# File Uploads
# ==========================================================

vendor_file = st.file_uploader(
    "Raw Vendor Store Cost File",
    type=["csv", "xlsx"]
)

master_file = st.file_uploader(
    "Master Price List",
    type=["xlsx"]
)

# ==========================================================
# Helper Functions
# ==========================================================

def read_vendor(uploaded_file):
    """Read Vendor Store Cost File."""

    if uploaded_file.name.lower().endswith(".csv"):
        return pd.read_csv(
            uploaded_file,
            low_memory=False
        )

    return pd.read_excel(uploaded_file)


def clean_uid(series):
    """
    Normalize Product IDs.

    Examples

    0006270
    06270
    6270
    6270.0

    become

    6270
    """

    return (
        series
        .fillna("")
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
        .str.upper()
        .str.lstrip("0")
        .replace("", pd.NA)
    )


def validate_columns(df, required_columns, file_name):
    """
    Validate required columns.
    """

    missing = [
        col
        for col in required_columns
        if col not in df.columns
    ]

    if missing:

        st.error(
            f"""
**{file_name}** is missing the following required column(s):

- """ + "\n- ".join(missing)
        )

        st.stop()


# ==========================================================
# Required Columns
# ==========================================================

VENDOR_COLUMNS = [
    "StoreID",
    "retailProductUID",
    "retailProductName",
    "group",
    "vendorProductUID"
]

MASTER_COLUMNS = [
    "Item .",
    "Retail Price",
    "Sales Price"
]

# ==========================================================
# Main
# ==========================================================

if vendor_file and master_file:

    progress = st.progress(
        0,
        text="Reading files..."
    )

    vendor = read_vendor(vendor_file)

    progress.progress(
        20,
        text="Reading Master Price List..."
    )

    master = pd.read_excel(master_file)

    progress.progress(
        40,
        text="Validating files..."
    )

    validate_columns(
        vendor,
        VENDOR_COLUMNS,
        "Vendor Store Cost File"
    )

    validate_columns(
        master,
        MASTER_COLUMNS,
        "Master Price List"
    )

    progress.progress(
        50,
        text="Preparing data..."
    )

        # ======================================================
    # PART 2 - Processing Logic
    # ======================================================

    progress.progress(
        60,
        text="Cleaning data..."
    )

    # -----------------------------
    # Normalize IDs
    # -----------------------------

    vendor["vendorProductUID"] = clean_uid(vendor["vendorProductUID"])
    vendor["retailProductUID"] = clean_uid(vendor["retailProductUID"])

    master["Item ."] = clean_uid(master["Item ."])

    # -----------------------------
    # Convert prices to numeric
    # -----------------------------

    master["Retail Price"] = pd.to_numeric(
        master["Retail Price"],
        errors="coerce"
    )

    master["Sales Price"] = pd.to_numeric(
        master["Sales Price"],
        errors="coerce"
    )

    # -----------------------------
    # Promo:
    # 0 means NO promo
    # Convert to blank
    # -----------------------------

    master["Sales Price"] = master["Sales Price"].mask(
        master["Sales Price"].fillna(0) == 0
    )

    # -----------------------------
    # Remove duplicate Item IDs
    # -----------------------------

    master = master.drop_duplicates(
        subset="Item .",
        keep="first"
    )

    progress.progress(
        70,
        text="Creating lookup dictionaries..."
    )

    retail_lookup = dict(
        zip(
            master["Item ."],
            master["Retail Price"]
        )
    )

    promo_lookup = dict(
        zip(
            master["Item ."],
            master["Sales Price"]
        )
    )

    # -----------------------------
    # Pack Type
    # -----------------------------

    vendor["Pack Type"] = (
        vendor["group"]
        .fillna("")
        .astype(str)
        .str.lower()
        .apply(
            lambda x: "Each"
            if "single" in x
            else "Pack"
        )
    )

    progress.progress(
        80,
        text="Building output tables..."
    )

    # ======================================================
    # Standard Output
    # ======================================================

    standard = pd.DataFrame({
        "StoreID": vendor["StoreID"],
        "RetailUID": vendor["retailProductUID"],
        "Retail": vendor["vendorProductUID"].map(retail_lookup),
        "Pack Type": vendor["Pack Type"],
        "retailProductName": vendor["retailProductName"],
        "group": vendor["group"]
    })

    # ======================================================
    # Promo Output
    # ======================================================

    promo = pd.DataFrame({
        "StoreID": vendor["StoreID"],
        "RetailUID": vendor["retailProductUID"],
        "Retail": vendor["vendorProductUID"].map(promo_lookup),
        "Pack Type": vendor["Pack Type"],
        "retailProductName": vendor["retailProductName"],
        "group": vendor["group"]
    })

    # Safety check
  # Convert promo retail to numeric
    promo["Retail"] = pd.to_numeric(
        promo["Retail"],
        errors="coerce"
    )

    # Treat 0 as blank
    promo.loc[promo["Retail"] == 0, "Retail"] = pd.NA
    
    # Remove rows with no promo retail
    promo = promo[promo["Retail"].notna()].reset_index(drop=True)

    # ======================================================
    # Missing Standard Retails ONLY
    # ======================================================

    missing = vendor.loc[
        standard["Retail"].isna(),
        [
            "StoreID",
            "retailProductUID",
            "vendorProductUID",
            "retailProductName",
            "group"
        ]
    ].copy()

    missing.rename(
        columns={
            "retailProductUID": "RetailUID",
            "vendorProductUID": "VendorProductUID"
        },
        inplace=True
    )

    # ======================================================
    # PART 3 - Export Workbook
    # ======================================================

    progress.progress(
        90,
        text="Generating Excel workbook..."
    )

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        standard.to_excel(
            writer,
            sheet_name="Standard Output",
            index=False
        )

        promo.to_excel(
            writer,
            sheet_name="Promo Output",
            index=False
        )

        missing.to_excel(
            writer,
            sheet_name="Missing Retails",
            index=False
        )

    output.seek(0)

    progress.progress(
        100,
        text="Done!"
    )

    st.success("✅ Pine State Liquor output has been generated successfully.")

    # ======================================================
    # Summary Metrics
    # ======================================================

    total_records = len(vendor)

    standard_matches = standard["Retail"].notna().sum()
    promo_matches = promo["Retail"].notna().sum()

    missing_standard = standard["Retail"].isna().sum()

    # Promo missing ignores intentional blanks (0 values converted to NA)
    promo_lookup_count = master["Sales Price"].notna().sum()
    # Products that don't exist in the master price list
    vendor_keys = set(vendor["vendorProductUID"].dropna())
    master_keys = set(master["Item ."].dropna())

    missing_promo = len(vendor_keys - master_keys)

    col1, col2, col3, col4, col5 = st.columns(5)

    col1.metric(
        "Vendor Records",
        f"{total_records:,}"
    )

    col2.metric(
        "Standard Matches",
        f"{standard_matches:,}"
    )

    col3.metric(
        "Promo Matches",
        f"{promo_matches:,}"
    )

    col4.metric(
        "Missing Standard",
        f"{missing_standard:,}"
    )

    col5.metric(
        "Missing Promo",
        f"{missing_promo:,}"
    )

    st.divider()

    st.download_button(
        label="📥 Download PineState_Liquor_Output.xlsx",
        data=output,
        file_name="PineState_Liquor_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
