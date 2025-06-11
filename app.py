import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from st_aggrid import AgGrid, GridOptionsBuilder
import streamlit as st
import logging, warnings
import calendar

warnings.filterwarnings("ignore")
st.set_page_config(page_title="📊 Monthly Sales Comparison Dashboard", layout="wide")
st.title("📈 Monthly Sales Spike/Drop Detector")

uploaded_file = st.file_uploader("Upload your monthly sales Excel file", type="xlsx")

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Ensure proper column names exist
    required_cols = {"Product_Name", "Quantity_Sold", "Sales_Value"}
    if not required_cols.issubset(df.columns):
        st.error(f"❌ Uploaded file must contain these columns: {required_cols}")
    else:
        # User input for current and previous month
        current_month = st.text_input("📅 Enter Current Month (e.g., May 2025)", "May 2025")

        try:
            # Parse user input
            selected_date = pd.to_datetime(current_month, format="%B %Y")
            prev_date = selected_date - pd.DateOffset(months=1)

            # Add artificial Date column
            df["Date"] = selected_date

            # Simulate previous month data (for demo or real merge in production use)
            hist_path = f"sales_data_{prev_date.strftime('%Y_%m')}.csv"
            try:
                prev_df = pd.read_csv(hist_path)
            except FileNotFoundError:
                st.warning(f"⚠️ Previous month file '{hist_path}' not found. Upload it in advance to enable comparison.")
                prev_df = pd.DataFrame(columns=["Product_Name", "Quantity_Sold", "Sales_Value"])

            # Add previous date column for consistency
            prev_df["Date"] = prev_date

            # Merge and compare
            merged = pd.merge(
                df,
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

            # Display dashboard
            st.subheader("🧾 Monthly Summary Table")

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

        except ValueError:
            st.error("❌ Please enter the current month in the format 'Month YYYY', e.g., 'May 2025'")
else:
    st.info("📤 Please upload your Excel sheet with sales data in the correct format: Product_Name, Quantity_Sold, Sales_Value")
