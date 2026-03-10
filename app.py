import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import io

# Helper function to convert dataframe to excel
@st.cache_data
def convert_df_to_excel(df, include_index=False):
    """Converts a dataframe to an Excel file in memory."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=include_index, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

st.set_page_config(page_title="Data Analytics Dashboard", layout="wide")

st.title("📊 Excel Data Analytics Dashboard")

# ---------------------------
# FILE UPLOAD
# ---------------------------
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("✅ File uploaded successfully!")

    # ---------------------------
    # DATA PREVIEW
    # ---------------------------
    st.subheader("Dataset Preview")
    df_head = df.head()
    st.dataframe(df_head)
    st.download_button(
        label="📥 Download as Excel",
        data=convert_df_to_excel(df_head),
        file_name="data_preview.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.subheader("Dataset Shape")
    st.write(f"Rows: {df.shape[0]} | Columns: {df.shape[1]}")

    st.subheader("Column Details")
    column_details_df = df.dtypes.astype(str).reset_index()
    column_details_df.columns = ["Column Name", "Data Type"]
    st.dataframe(column_details_df)
    st.download_button(
        label="📥 Download as Excel",
        data=convert_df_to_excel(column_details_df),
        file_name="column_details.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ---------------------------
    # CREATE TABS
    # ---------------------------
    tab1, tab2, tab3 = st.tabs(["📈 Analysis", "📊 Visualizations", "🔥 Advanced"])

    # ===========================
    # TAB 1 – ANALYSIS
    # ===========================
    with tab1:
        st.header("Data Analysis Section")

        # KPI Cards
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("Total Records", df.shape[0])
        kpi2.metric("Total Columns", df.shape[1])
        kpi3.metric("Missing Values", df.isna().sum().sum())
        kpi4.metric("Numeric Columns", df.select_dtypes(include=np.number).shape[1])

        # Summary Statistics
        st.subheader("Summary Statistics")
        summary_df = df.describe()
        st.write(summary_df)
        st.download_button(
            label="📥 Download as Excel",
            data=convert_df_to_excel(summary_df, include_index=True),
            file_name="summary_statistics.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Groupby Analysis
        if "Category" in df.columns and "Sales" in df.columns:
            st.subheader("Sales by Category")
            st.bar_chart(df.groupby("Category")["Sales"].sum())

        # Top 5 / Bottom 5
        st.subheader("Top 5 Sales Records")
        top_5_df = df.nlargest(5, "Sales")
        st.dataframe(top_5_df)
        st.download_button(
            label="📥 Download as Excel",
            data=convert_df_to_excel(top_5_df),
            file_name="top_5_sales.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.subheader("Bottom 5 Sales Records")
        bottom_5_df = df.nsmallest(5, "Sales")
        st.dataframe(bottom_5_df)
        st.download_button(
            label="📥 Download as Excel",
            data=convert_df_to_excel(bottom_5_df),
            file_name="bottom_5_sales.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # ===========================
    # TAB 2 – VISUALIZATIONS
    # ===========================
    with tab2:
        st.header("Visualization Section")
        sns.set_style("whitegrid")

        # KPI Cards on Top
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("Total Sales", df["Sales"].sum())
        kpi2.metric("Average Sales", round(df["Sales"].mean(), 2))
        kpi3.metric("Max Sale", df["Sales"].max())
        kpi4.metric("Min Sale", df["Sales"].min())

        st.markdown("---")

        # Arrange 6 charts in grid (3 rows × 2 columns)
        charts = []
        charts.append(("Sales by Category & Region", "bar", {"x": "Sales", "y": "Category", "hue": "Region"}))
        charts.append(("Sales Distribution by Region", "pie", {"group_col": "Region", "value_col": "Sales"}))
        charts.append(("Sales Histogram", "hist", {"col": "Sales"}))
        charts.append(("Sales by Category Boxplot", "box", {"x": "Sales", "y": "Category"}))
        charts.append(("Total Sales by Category", "bar_category", {"x": "Category", "y": "Sales"}))
        charts.append(("Sales vs Profit Scatter", "scatter", {"x": "Sales", "y": "Profit", "hue": "Region"}))

        for i in range(0, len(charts), 2):
            col1, col2 = st.columns(2)
            for j, col in enumerate([col1, col2]):
                if i + j >= len(charts):
                    continue
                title, chart_type, params = charts[i + j]
                with col:
                    st.write(f"### {title}")
                    fig, ax = plt.subplots(figsize=(8, 5))
                    if chart_type == "bar":
                        sns.barplot(x=params["x"], y=params["y"], hue=params["hue"], data=df, palette="Set2", ax=ax)
                    elif chart_type == "pie":
                        group = df.groupby(params["group_col"])[params["value_col"]].sum()
                        ax.pie(group, labels=group.index, autopct="%1.1f%%", startangle=90, colors=sns.color_palette("pastel"))
                        ax.axis("equal")
                    elif chart_type == "hist":
                        sns.histplot(df[params["col"]], bins=20, kde=True, color="skyblue", ax=ax)
                        ax.set_xlabel(params["col"])
                        ax.set_ylabel("Frequency")
                    elif chart_type == "box":
                        sns.boxplot(x=params["x"], y=params["y"], data=df, palette="Set2", ax=ax)
                    elif chart_type == "bar_category":
                        cat_sales = df.groupby(params["x"])[params["y"]].sum().sort_values(ascending=False)
                        sns.barplot(x=cat_sales.index, y=cat_sales.values, palette="viridis", ax=ax)
                        ax.set_xlabel(params["x"])
                        ax.set_ylabel("Total Sales")
                        ax.tick_params(axis='x', rotation=45)
                    elif chart_type == "scatter" and "Profit" in df.columns:
                        sns.scatterplot(x=params["x"], y=params["y"], hue=params["hue"], data=df, palette="Set1", ax=ax)
                    fig.tight_layout()
                    st.pyplot(fig)

    # ===========================
    # TAB 3 – ADVANCED
    # ===========================
    with tab3:
        st.header("Advanced Analysis")

        # Filtering Section
        st.subheader("Filter Data")
        region_filter = st.multiselect("Select Region", options=df["Region"].unique(), default=df["Region"].unique())
        filtered_df = df[df["Region"].isin(region_filter)]

        min_sales = int(df["Sales"].min())
        max_sales = int(df["Sales"].max())
        sales_range = st.slider("Select Sales Range", min_sales, max_sales, (min_sales, max_sales))
        filtered_df = filtered_df[(filtered_df["Sales"] >= sales_range[0]) & (filtered_df["Sales"] <= sales_range[1])]

        st.dataframe(filtered_df)

        st.download_button(
            label="📥 Download Filtered Data as Excel",
            data=convert_df_to_excel(filtered_df),
            file_name="filtered_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

       # Correlation Heatmap
        st.subheader("Correlation Heatmap")

        numeric_df = filtered_df.select_dtypes(include=np.number)

        if not numeric_df.empty:

            col1, col2, col3 = st.columns([1,2,1])  

            with col2:
                fig, ax = plt.subplots(figsize=(5,4))  

                sns.heatmap(
                    numeric_df.corr(),
                    annot=True,
                    cmap="coolwarm",
                    fmt=".2f",
                    linewidths=0.5,
                    square=True,
                    cbar_kws={"shrink": .7},   
                    ax=ax
                )

                ax.set_title("Correlation Matrix", fontsize=12)

                plt.tight_layout()
                st.pyplot(fig, use_container_width=False)

else:
    st.info("Please upload an Excel file to begin.")