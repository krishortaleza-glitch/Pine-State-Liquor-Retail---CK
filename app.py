import streamlit as st
import pandas as pd
from io import BytesIO
import time

st.set_page_config(page_title="Pine State Liquor Retail Builder", page_icon="🍺", layout="wide")

st.title("🍺 Pine State Liquor Retail Builder")
st.caption("Generate Standard and Promo retail output files from the Vendor Store Cost File and Pine State Master Price List.")

vendor_file = st.file_uploader("Upload Raw Vendor Store Cost File", type=["csv","xlsx"])
master_file = st.file_uploader("Upload Pine State Liquor Master Price List", type=["xlsx"])

def read_vendor(file):
    if file.name.lower().endswith(".csv"):
        return pd.read_csv(file, low_memory=False)
    return pd.read_excel(file)

if vendor_file and master_file:

    progress = st.progress(0, text="Reading Vendor File...")
    vendor = read_vendor(vendor_file)

    progress.progress(20, text="Reading Master Price List...")
    master = pd.read_excel(master_file)

    progress.progress(40, text="Cleaning data...")

    vendor["vendorProductUID"] = vendor["vendorProductUID"].astype(str).str.strip().str.upper()
    master["Item ."] = master["Item ."].astype(str).str.strip().str.upper()

    retail_lookup = dict(zip(master["Item ."], master["Retail Price"]))
    promo_lookup = dict(zip(master["Item ."], master["Sales Price"]))

    vendor["Pack Type"] = vendor["group"].fillna("").astype(str).str.lower().apply(
        lambda x: "Each" if "single" in x else "Pack"
    )

    progress.progress(60, text="Matching retail prices...")

    standard = pd.DataFrame({
        "StoreID": vendor["StoreID"],
        "RetailUID": vendor["retailProductUID"],
        "Retail": vendor["vendorProductUID"].map(retail_lookup),
        "Pack Type": vendor["Pack Type"],
        "retailProductName": vendor["retailProductName"],
        "group": vendor["group"]
    })

    promo = pd.DataFrame({
        "StoreID": vendor["StoreID"],
        "RetailUID": vendor["retailProductUID"],
        "Retail": vendor["vendorProductUID"].map(promo_lookup),
        "Pack Type": vendor["Pack Type"],
        "retailProductName": vendor["retailProductName"],
        "group": vendor["group"]
    })

    missing = vendor.loc[
        standard["Retail"].isna() | promo["Retail"].isna(),
        ["StoreID","retailProductUID","vendorProductUID","retailProductName","group"]
    ].copy()

    missing.rename(columns={"retailProductUID":"RetailUID","vendorProductUID":"VendorProductUID"}, inplace=True)
    missing["Missing Standard"] = missing["VendorProductUID"].map(retail_lookup).isna()
    missing["Missing Promo"] = missing["VendorProductUID"].map(promo_lookup).isna()

    progress.progress(80, text="Generating Excel...")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        standard.to_excel(writer, sheet_name="Standard Output", index=False)
        promo.to_excel(writer, sheet_name="Promo Output", index=False)
        missing.to_excel(writer, sheet_name="Missing Retails", index=False)

    output.seek(0)

    progress.progress(100, text="Done!")

    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Vendor Records", len(vendor))
    c2.metric("Standard Matches", int(standard["Retail"].notna().sum()))
    c3.metric("Promo Matches", int(promo["Retail"].notna().sum()))
    c4.metric("Missing Standard", int(standard["Retail"].isna().sum()))
    c5.metric("Missing Promo", int(promo["Retail"].isna().sum()))

    st.download_button(
        "📥 Download PineState_Liquor_Output.xlsx",
        data=output,
        file_name="PineState_Liquor_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.subheader("Standard Output Preview")
    st.dataframe(standard.head(20), use_container_width=True)

    st.subheader("Promo Output Preview")
    st.dataframe(promo.head(20), use_container_width=True)

    st.subheader("Missing Retails Preview")
    st.dataframe(missing.head(20), use_container_width=True)
