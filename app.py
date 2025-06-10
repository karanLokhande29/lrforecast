import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.linear_model import LinearRegression
from st_aggrid import AgGrid, GridOptionsBuilder
import streamlit as st
import logging, warnings

warnings.filterwarnings("ignore")
st.set_page_config(page_title="Sales Forecast Dashboard (Linear Regression)", layout="wide")
st.title("📊 Factory-Wise Sales Forecast Dashboard (LR Model)")

uploaded_file = st.file_uploader("Upload your factory sales CSV", type="csv")

if uploaded_file:
    df = pd.read_csv(uploaded_file, parse_dates=["Date"])
    df.rename(columns={"Date": "ds", "Quantity_Sold": "y"}, inplace=True)

    df["ds_ordinal"] = df["ds"].map(pd.Timestamp.toordinal)

    latest_date = df["ds"].max().date()
    st.markdown(f"**🗓️ Latest Date in Dataset:** `{latest_date}`")

    summary = []

    for (product, factory), group in df.groupby(["Product_Name", "Factory"]):
        group = group.sort_values("ds")
        group["ds_ordinal"] = group["ds"].map(pd.Timestamp.toordinal)

        model = LinearRegression()
        X = group[["ds_ordinal"]]
        y = group["y"]
        model.fit(X, y)

        future_dates = pd.date_range(start=group["ds"].max() + pd.Timedelta(days=1), periods=30)
        future_ordinals = future_dates.map(pd.Timestamp.toordinal).values.reshape(-1, 1)
        y_pred = model.predict(future_ordinals)

        avg_hist = y.mean()
        avg_pred = y_pred.mean()
        growth = ((avg_pred - avg_hist) / avg_hist * 100) if avg_hist else 0

        alert = "✅ Stable"
        if growth > 10: alert = "📈 Spike"
        elif growth < -10: alert = "📉 Drop"

        summary.append({
            "Product": product,
            "Factory": factory,
            "First_Date": group["ds"].min().date(),
            "Last_Date": group["ds"].max().date(),
            "Last_Sale": group['y'].iloc[-1],
            "Avg_Historical_Sales": round(avg_hist, 2),
            "Predicted_Avg_Sales": round(avg_pred, 2),
            "Total_Forecast_30d": round(y_pred.sum()),
            "Growth_%": round(growth, 2),
            "Alert": alert
        })

    summary_df = pd.DataFrame(summary)

    st.subheader("📋 Forecast Summary Table")
    with st.expander("🔍 Advanced Filters", expanded=True):
        alert_filter = st.multiselect("⚠️ Alert Type", summary_df["Alert"].unique(), summary_df["Alert"].unique())
        factory_filter = st.multiselect("🏭 Factory", summary_df["Factory"].unique(), summary_df["Factory"].unique())
        product_search = st.text_input("🔎 Product Name")
        growth_min, growth_max = st.slider("📈 Growth % Range", float(summary_df["Growth_%"].min()), float(summary_df["Growth_%"].max()), (float(summary_df["Growth_%"].min()), float(summary_df["Growth_%"].max())))

        filtered_df = summary_df[
            (summary_df["Alert"].isin(alert_filter)) &
            (summary_df["Factory"].isin(factory_filter)) &
            (summary_df["Growth_%"].between(growth_min, growth_max))
        ]
        if product_search:
            filtered_df = filtered_df[filtered_df["Product"].str.contains(product_search, case=False)]

    gb = GridOptionsBuilder.from_dataframe(filtered_df)
    gb.configure_pagination()
    gb.configure_default_column(filterable=True, sortable=True, resizable=True)
    gb.configure_side_bar()
    AgGrid(filtered_df, gridOptions=gb.build(), theme='material')

    st.download_button("📥 Download CSV", data=filtered_df.to_csv(index=False), file_name="summary.csv")

    st.subheader("📈 Forecast Visualization")
    selectable = filtered_df["Product"] + " - " + filtered_df["Factory"]
    selected = st.selectbox("Choose Product-Factory", selectable)
    if selected:
        prod, fact = selected.split(" - ")
        group = df[(df["Product_Name"] == prod) & (df["Factory"] == fact)].sort_values("ds")
        group["ds_ordinal"] = group["ds"].map(pd.Timestamp.toordinal)

        model = LinearRegression()
        X = group[["ds_ordinal"]]
        y = group["y"]
        model.fit(X, y)

        future_dates = pd.date_range(start=group["ds"].max() + pd.Timedelta(days=1), periods=30)
        future_ordinals = future_dates.map(pd.Timestamp.toordinal).values.reshape(-1, 1)
        y_pred = model.predict(future_ordinals)

        plt.figure(figsize=(10, 4))
        plt.plot(group["ds"], y, label="Historical")
        plt.plot(future_dates, y_pred, label="Forecast")
        plt.title(f"Forecast: {prod} - {fact}")
        plt.xlabel("Date")
        plt.ylabel("Quantity Sold")
        plt.legend()
        st.pyplot(plt)
else:
    st.info("📥 Upload a CSV file to begin.")
