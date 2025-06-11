import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from st_aggrid import AgGrid, GridOptionsBuilder
import streamlit as st
import logging, warnings
import zipfile, os
from io import BytesIO
from datetime import datetime

warnings.filterwarnings("ignore")
st.set_page_config(page_title="📊 Monthly Sales Comparison Dashboard", layout="wide")
st.title("📈 Multi-Month Sales Spike/Drop Detector")

uploaded_zip = st.file_uploader("Upload a ZIP file containing monthly Excel sheets", type="zip")

if uploaded_zip:
    with zipfile.ZipFile(uploaded_zip) as z:
        # Extract Excel files to memory
        dfs = {}
        for name in z.namelist():
            if name.endswith(".xlsx"):
                df = pd.read_excel(z.open(name))
                if {"Product_Name", "Quantity_Sold", "Sales_Value"}.issubset(df.columns):
                    # Try to extract date from filename
                    try:
                        parts = name.replace(".xlsx", "").split("_")
                        month_year = parts[-2] + " " + parts[-1]  # e.g., may 2025
                        date = pd.to_datetime(month_year, format="%B %Y")
                        df["Date"] = date
                        dfs[date] = df
                    except:
                        st.warning(f"⚠️ Could not parse date from file: {name}")
                else:
                    st.warning(f"❌ Skipping file '{name}' — missing required columns.")

        if len(dfs) < 2:
            st.error("❗ Upload at least two valid monthly Excel sheets to compare.")
        else:
            # Sort by date and get latest two
            all_dates = sorted(dfs.keys())
            current_df = dfs[all_dates[-1]]
            prev_df = dfs[all_dates[-2]]

            merged = pd.merge(
                current_df,
                prev_df,
                on="Product_Name",
                suffixes=("_curr", "_prev"),
                how="outer"
            ).fillna(0)

            merged["Growth_Quantity_%"] = ((merged["Quantity_Sold_curr"] - merged["Quantity_Sold_prev"]) /
                                            merged["Quantity_Sold_prev"].replace(0, np.nan)) * 100
            merged["Growth_Value_%"] = ((merged["Sales_Value_curr"] - merged["Sales_Value_prev"]) /
                                         merged["Sales_Value_prev"].replace(0, np.nan)) * 100

            def label_growth(g):
                if g > 10:
                    return "📈 Spike"
                elif g < -10:
                    return "📉 Drop"
                else:
                    return "✅ Stable"

            merged["Alert"] = merged["Growth_Quantity_%"].apply(label_growth)

            st.subheader(f"🧾 Comparison: {all_dates[-2].strftime('%B %Y')} ➡ {all_dates[-1].strftime('%B %Y')}")

            gb = GridOptionsBuilder.from_dataframe(merged)
            gb.configure_pagination()
            gb.configure_default_column(filterable=True, sortable=True, resizable=True)
            gb.configure_side_bar()
            AgGrid(merged, gridOptions=gb.build(), theme='material')

            st.download_button("📥 Download Comparison CSV", data=merged.to_csv(index=False), file_name="monthly_comparison.csv")

            st.subheader("📊 Growth Trend Visualization")
            selected = st.selectbox("Select a Product", merged["Product_Name"])

            if selected:
                sel = merged[merged["Product_Name"] == selected]
                fig, ax = plt.subplots(figsize=(8, 4))
                ax.bar(["Previous"], sel["Quantity_Sold_prev"].values, label="Prev Qty")
                ax.bar(["Current"], sel["Quantity_Sold_curr"].values, label="Curr Qty")
                ax.set_title(f"Quantity Comparison: {selected}")
                ax.legend()
                st.pyplot(fig)
else:
    st.info("📤 Please upload a ZIP file with at least two Excel sheets. Filenames should end with '_may_2025.xlsx', '_june_2025.xlsx', etc.")
