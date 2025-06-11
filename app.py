import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from st_aggrid import AgGrid, GridOptionsBuilder
import streamlit as st
import logging, warnings
import zipfile, os
from io import BytesIO
from datetime import datetime
from sklearn.linear_model import LinearRegression

warnings.filterwarnings("ignore")
st.set_page_config(page_title="📊 Monthly Sales Dashboard", layout="wide")
st.title("📈 Multi-Month Sales Dashboard with Forecasting")

uploaded_zip = st.file_uploader("Upload a ZIP file containing monthly Excel sheets", type="zip")

if uploaded_zip:
    with zipfile.ZipFile(uploaded_zip) as z:
        dfs = {}
        for name in z.namelist():
            if name.endswith(".xlsx"):
                df = pd.read_excel(z.open(name))
                if {"Product_Name", "Quantity_Sold", "Sales_Value"}.issubset(df.columns):
                    try:
                        parts = name.replace(".xlsx", "").split("_")
                        month_year = parts[-2] + " " + parts[-1]
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
            all_dates = sorted(dfs.keys())
            combined_df = pd.concat([df.assign(Month=dt.strftime("%B %Y")) for dt, df in dfs.items()])

            st.subheader("🔍 Filter Data")
            month_options = [dt.strftime("%B %Y") for dt in all_dates]
            selected_month = st.selectbox("Select Month to View Sales Data", month_options, index=len(month_options)-1)
            filtered_data = combined_df[combined_df["Month"] == selected_month]

            product_filter = st.text_input("Search Product Name (optional)")
            if product_filter:
                filtered_data = filtered_data[filtered_data["Product_Name"].str.contains(product_filter, case=False)]

            st.subheader("📄 Monthly Sales Data")
            gb_all = GridOptionsBuilder.from_dataframe(filtered_data)
            gb_all.configure_pagination()
            gb_all.configure_default_column(filterable=True, sortable=True, resizable=True)
            AgGrid(filtered_data, gridOptions=gb_all.build(), theme='material')

            total_cost = filtered_data["Sales_Value"].sum()
            st.markdown(f"### 💰 Total Sales Value: ₹{total_cost:,.2f}")

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

            st.subheader(f"📊 Comparison: {all_dates[-2].strftime('%B %Y')} ➡ {all_dates[-1].strftime('%B %Y')}")

            gb = GridOptionsBuilder.from_dataframe(merged)
            gb.configure_pagination()
            gb.configure_default_column(filterable=True, sortable=True, resizable=True)
            gb.configure_side_bar()
            AgGrid(merged, gridOptions=gb.build(), theme='material')

            st.download_button("📥 Download Comparison CSV", data=merged.to_csv(index=False), file_name="monthly_comparison.csv")

            st.subheader("📈 Forecast Next 30 Days (Quantity)")
            history = combined_df.groupby(["Date", "Product_Name"])["Quantity_Sold"].sum().reset_index()
            history["Date_Ordinal"] = history["Date"].map(datetime.toordinal)

            selected_product = st.selectbox("Choose Product for Forecasting", sorted(history["Product_Name"].unique()))
            prod_data = history[history["Product_Name"] == selected_product].sort_values("Date")

            model = LinearRegression()
            X = prod_data[["Date_Ordinal"]]
            y = prod_data["Quantity_Sold"]
            model.fit(X, y)

            future_dates = pd.date_range(start=prod_data["Date"].max() + pd.DateOffset(days=1), periods=30)
            future_ordinals = future_dates.map(datetime.toordinal).values.reshape(-1, 1)
            forecast = model.predict(future_ordinals)

            fig, ax = plt.subplots(figsize=(10, 4))
            ax.plot(prod_data["Date"], y, label="Historical")
            ax.plot(future_dates, forecast, label="Forecast (30d)")
            ax.set_title(f"30-Day Forecast: {selected_product}")
            ax.set_ylabel("Quantity Sold")
            ax.legend()
            st.pyplot(fig)

            # Predict total sales value for next month
            recent_month = history[history["Product_Name"] == selected_product]["Date"].max()
            future_month_date = recent_month + pd.DateOffset(months=1)
            month_ord = future_month_date.toordinal()
            predicted_qty = model.predict(np.array([[month_ord]]))[0]

            st.markdown(f"### 🔮 Predicted Quantity for {selected_product} in {future_month_date.strftime('%B %Y')}: `{int(predicted_qty):,}` units")
else:
    st.info("📤 Please upload a ZIP file with at least two Excel sheets. Filenames should end with '_may_2025.xlsx', '_june_2025.xlsx', etc.")
