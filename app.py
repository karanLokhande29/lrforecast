import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from st_aggrid import AgGrid, GridOptionsBuilder
import streamlit as st
import logging, warnings
import zipfile
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
                        st.warning(f"⚠️ Could not parse date from file: {name}. Use format like 'sales_may_2025.xlsx'")
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

            st.download_button("📤 Download Current Month Data", data=filtered_data.to_csv(index=False), file_name=f"sales_data_{selected_month.replace(' ', '_')}.csv")

            total_cost = filtered_data["Sales_Value"].sum()
            st.markdown(f"### 💰 Total Sales Value: ₹{total_cost:,.2f}")

            # Comparison between last 2 months
            current_df = dfs[all_dates[-1]].copy()
            prev_df = dfs[all_dates[-2]].copy()

            current_df = current_df.rename(columns={"Quantity_Sold": "Quantity_Sold_curr", "Sales_Value": "Sales_Value_curr"})
            prev_df = prev_df.rename(columns={"Quantity_Sold": "Quantity_Sold_prev", "Sales_Value": "Sales_Value_prev"})

            merged = pd.merge(
                current_df[["Product_Name", "Quantity_Sold_curr", "Sales_Value_curr"]],
                prev_df[["Product_Name", "Quantity_Sold_prev", "Sales_Value_prev"]],
                on="Product_Name",
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

            # 📊 Product-wise Monthly Trend
            st.subheader("📊 Product-wise Monthly Trend")
            selected_prod = st.selectbox("Select Product for Trendline", sorted(combined_df["Product_Name"].unique()))
            trend_data = combined_df[combined_df["Product_Name"] == selected_prod].groupby("Date")[["Quantity_Sold", "Sales_Value"]].sum().reset_index()

            fig, ax = plt.subplots(figsize=(10, 4))
            ax.plot(trend_data["Date"], trend_data["Quantity_Sold"], marker="o", label="Quantity Sold")
            ax.plot(trend_data["Date"], trend_data["Sales_Value"], marker="x", label="Sales Value (₹)")
            ax.set_title(f"Trendline for {selected_prod}")
            ax.set_ylabel("Amount")
            ax.legend()
            st.pyplot(fig)

            # 🔮 Forecasting for selected product
            st.subheader("🔮 Forecast Next 30 Days (Selected Product)")
            history = combined_df.groupby(["Date", "Product_Name"]).agg({"Quantity_Sold": "sum", "Sales_Value": "sum"}).reset_index()
            history["Date_Ordinal"] = history["Date"].map(datetime.toordinal)

            selected_forecast_prod = st.selectbox("Product for Forecast", sorted(history["Product_Name"].unique()))
            hist = history[history["Product_Name"] == selected_forecast_prod].sort_values("Date")
            model = LinearRegression()
            model.fit(hist[["Date_Ordinal"]], hist["Quantity_Sold"])
            future_dates = pd.date_range(start=hist["Date"].max() + pd.Timedelta(days=1), periods=30)
            forecast_qty = model.predict(future_dates.map(datetime.toordinal).values.reshape(-1, 1))

            fig, ax = plt.subplots(figsize=(10, 4))
            ax.plot(hist["Date"], hist["Quantity_Sold"], label="Historical")
            ax.plot(future_dates, forecast_qty, label="Forecast")
            ax.set_title(f"30-Day Forecast: {selected_forecast_prod}")
            ax.set_ylabel("Quantity")
            ax.legend()
            st.pyplot(fig)

            forecast_df = pd.DataFrame({"Date": future_dates, "Predicted_Quantity": forecast_qty})
            st.download_button("📥 Download Forecast (Selected Product)", forecast_df.to_csv(index=False), file_name=f"{selected_forecast_prod}_forecast.csv")

            # 📦 Forecast summary for all products
            st.subheader("📦 Forecast Summary for All Products")
            all_forecasts = []
            for prod in sorted(history["Product_Name"].unique()):
                hist_prod = history[history["Product_Name"] == prod]
                if len(hist_prod) >= 2:
                    model_qty = LinearRegression()
                    model_val = LinearRegression()
                    model_qty.fit(hist_prod[["Date_Ordinal"]], hist_prod["Quantity_Sold"])
                    model_val.fit(hist_prod[["Date_Ordinal"]], hist_prod["Sales_Value"])
                    target_date = hist_prod["Date"].max() + pd.DateOffset(months=1)
                    ordinal = target_date.toordinal()
                    qty_pred = model_qty.predict(np.array([[ordinal]]))[0]
                    val_pred = model_val.predict(np.array([[ordinal]]))[0]
                    all_forecasts.append({
                        "Product_Name": prod,
                        "Forecasted_30d_Quantity": round(qty_pred),
                        "Forecasted_Sales_Value": round(val_pred, 2)
                    })

            forecast_summary_df = pd.DataFrame(all_forecasts)
            AgGrid(forecast_summary_df)
            st.download_button("📥 Download Forecast Summary (All Products)", data=forecast_summary_df.to_csv(index=False), file_name="forecast_summary_all_products.csv")
else:
    st.info("📤 Please upload a ZIP file with Excel sheets named like 'sales_april_2025.xlsx', 'sales_may_2025.xlsx', etc.")
