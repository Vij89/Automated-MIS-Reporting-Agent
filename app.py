
import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from datetime import datetime

# Optional Groq integration
from groq import Groq


# ============================================================
# 1. PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Automated MIS Reporting Agent",
    page_icon="📊",
    layout="wide"
)


# ============================================================
# 2. GROQ CONFIGURATION
# ============================================================

client = None

try:
    if "GROQ_API_KEY" in st.secrets:
        client = Groq(
            api_key=st.secrets["GROQ_API_KEY"]
        )
except Exception:
    client = None


# ============================================================
# 3. APPLICATION HEADER
# ============================================================

st.title("📊 Automated MIS Reporting Agent")

st.markdown("""
Welcome to your AI-powered Business Reporting System.

Upload business data, analyse KPIs, explore business
performance and generate automated management reports.
""")

st.divider()


# ============================================================
# 4. SIDEBAR NAVIGATION
# ============================================================

st.sidebar.title("MIS Navigation")

menu = st.sidebar.radio(
    "Select Module",
    [
        "Home",
        "Data Upload",
        "Dashboard",
        "AI Insights",
        "Reports"
    ]
)


# ============================================================
# 5. HOME PAGE
# ============================================================

if menu == "Home":

    st.subheader("Welcome to MIS Dashboard")

    col1, col2 = st.columns(2)

    col1.info("📁 Upload and analyse business data")
    col2.info("📈 Generate business performance reports")

    st.markdown("### Project Features")

    st.write("""
    - Excel and CSV data upload
    - Data cleaning and validation
    - Business KPI calculations
    - Interactive sales dashboard
    - Regional and product analysis
    - AI-powered management insights
    - Automated Excel report generation
    """)

    st.success("Application is running successfully!")


# ============================================================
# 6. DATA UPLOAD AND VALIDATION
# ============================================================

elif menu == "Data Upload":

    st.subheader("📁 Data Upload & Validation")

    uploaded_file = st.file_uploader(
        "Upload Sales Excel or CSV",
        type=["csv", "xlsx"]
    )

    if uploaded_file is not None:

        try:

            # Read uploaded file
            if uploaded_file.name.endswith(".csv"):

                df = pd.read_csv(uploaded_file)

            else:

                df = pd.read_excel(uploaded_file)

            # Clean column names
            df.columns = df.columns.str.strip()

            # Remove duplicate records
            duplicates = df.duplicated().sum()

            df = df.drop_duplicates().copy()

            # Convert date column
            if "Order_Date" in df.columns:

                df["Order_Date"] = pd.to_datetime(
                    df["Order_Date"],
                    errors="coerce"
                )

            # Validate required columns
            required_columns = [
                "Order_ID",
                "Order_Date",
                "Region",
                "Product",
                "Product_Category",
                "Revenue",
                "Profit",
                "Quantity"
            ]

            missing_columns = [
                col for col in required_columns
                if col not in df.columns
            ]

            if missing_columns:

                st.error(
                    f"Missing required columns: {missing_columns}"
                )

            else:

                # Store cleaned dataset
                st.session_state["sales_data"] = df

                missing_values = df.isnull().sum().sum()

                st.success(
                    "File uploaded and processed successfully!"
                )

                # Dataset Preview
                st.subheader("Dataset Preview")

                st.dataframe(
                    df.head(10),
                    use_container_width=True,
                    hide_index=True
                )

                # Data Quality Report
                st.subheader("Data Quality Report")

                col1, col2, col3 = st.columns(3)

                col1.metric(
                    "Total Records",
                    f"{len(df):,}"
                )

                col2.metric(
                    "Duplicates Removed",
                    f"{duplicates:,}"
                )

                col3.metric(
                    "Missing Values",
                    f"{missing_values:,}"
                )

                # Dataset Information
                st.subheader("Dataset Information")

                col1, col2, col3 = st.columns(3)

                col1.metric(
                    "Total Columns",
                    df.shape[1]
                )

                col2.metric(
                    "Numeric Columns",
                    len(
                        df.select_dtypes(
                            include="number"
                        ).columns
                    )
                )

                col3.metric(
                    "Text Columns",
                    len(
                        df.select_dtypes(
                            include="object"
                        ).columns
                    )
                )

                # Column Structure
                st.subheader("Column Structure")

                schema_df = pd.DataFrame({
                    "Column Name": df.columns,
                    "Data Type": df.dtypes.astype(str).values,
                    "Missing Values": df.isnull().sum().values
                })

                st.dataframe(
                    schema_df,
                    use_container_width=True,
                    hide_index=True
                )

                # Statistical Summary
                with st.expander("View Statistical Summary"):

                    st.dataframe(
                        df.describe(),
                        use_container_width=True
                    )

        except Exception as e:

            st.error(f"File processing error: {e}")


# ============================================================
# 7. BUSINESS DASHBOARD
# ============================================================


# ============================================================
# 7. PROFESSIONAL BUSINESS DASHBOARD
# ============================================================

elif menu == "Dashboard":

    # ---------------- CUSTOM CSS ----------------

    st.markdown("""
    <style>

    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }

    .dashboard-header {
        font-size: 30px;
        font-weight: 700;
        color: #172B4D;
        margin-bottom: 4px;
    }

    .dashboard-subtitle {
        color: #718096;
        font-size: 14px;
        margin-bottom: 25px;
    }

    .kpi-card {
        background: #FFFFFF;
        border: 1px solid #E5EAF2;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04);
        min-height: 125px;
    }

    .kpi-label {
        color: #718096;
        font-size: 13px;
        font-weight: 600;
        margin-bottom: 10px;
    }

    .kpi-value {
        color: #172B4D;
        font-size: 27px;
        font-weight: 700;
    }

    .section-heading {
        color: #172B4D;
        font-size: 19px;
        font-weight: 700;
        margin-top: 12px;
        margin-bottom: 12px;
    }

    </style>
    """, unsafe_allow_html=True)

    # ---------------- HEADER ----------------

    st.markdown(
        '<div class="dashboard-header">Business Performance Overview</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="dashboard-subtitle">'
        'Monitor revenue, profitability and regional performance'
        '</div>',
        unsafe_allow_html=True
    )

    if "sales_data" not in st.session_state:

        st.warning("Please upload your dataset first.")

    else:

        df = st.session_state["sales_data"].copy()

        # ---------------- FILTERS ----------------

        st.markdown("### Dashboard Filters")

        f1, f2 = st.columns(2)

        regions = sorted(
            df["Region"].dropna().unique().tolist()
        )

        selected_regions = f1.multiselect(
            "Select Region",
            regions,
            default=regions
        )

        filtered_df = df[
            df["Region"].isin(selected_regions)
        ].copy()

        if "Order_Date" in filtered_df.columns:

            filtered_df["Order_Date"] = pd.to_datetime(
                filtered_df["Order_Date"],
                errors="coerce"
            )

            valid_dates = filtered_df["Order_Date"].dropna()

            if not valid_dates.empty:

                date_range = f2.date_input(
                    "Select Date Range",
                    value=(
                        valid_dates.min().date(),
                        valid_dates.max().date()
                    )
                )

                if isinstance(date_range, (tuple, list)) and len(date_range) == 2:

                    start_date, end_date = date_range

                    filtered_df = filtered_df[
                        filtered_df["Order_Date"].dt.date.between(
                            start_date,
                            end_date
                        )
                    ]

        if filtered_df.empty:

            st.warning("No records found for the selected filters.")

        else:

            # ---------------- KPI CALCULATIONS ----------------

            total_orders = filtered_df["Order_ID"].nunique()

            total_revenue = filtered_df["Revenue"].sum()

            total_profit = filtered_df["Profit"].sum()

            profit_margin = (
                total_profit / total_revenue * 100
                if total_revenue != 0 else 0
            )

            # ---------------- KPI CARDS ----------------

            st.markdown("### Key Performance Indicators")

            c1, c2, c3, c4 = st.columns(4)

            kpis = [
                ("TOTAL ORDERS", f"{total_orders:,}", "📦"),
                ("TOTAL REVENUE", f"₹{total_revenue:,.0f}", "💰"),
                ("TOTAL PROFIT", f"₹{total_profit:,.0f}", "📈"),
                ("PROFIT MARGIN", f"{profit_margin:.2f}%", "🎯")
            ]

            for col, (label, value, emoji) in zip(
                [c1, c2, c3, c4], kpis
            ):

                with col:

                    st.markdown(
                        f"""
                        <div class="kpi-card">
                            <div class="kpi-label">
                                {emoji} {label}
                            </div>
                            <div class="kpi-value">
                                {value}
                            </div>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

            st.markdown("<br>", unsafe_allow_html=True)

            # ---------------- CHART THEME ----------------

            chart_template = "plotly_white"

            # ---------------- REVENUE BY REGION ----------------

            st.markdown(
                '<div class="section-heading">Regional Revenue Performance</div>',
                unsafe_allow_html=True
            )

            region_data = (
                filtered_df.groupby("Region", as_index=False)
                .agg(
                    Revenue=("Revenue", "sum"),
                    Profit=("Profit", "sum")
                )
            )

            fig1 = px.bar(
                region_data,
                x="Region",
                y="Revenue",
                color="Region",
                text_auto=".2s",
                template=chart_template,
                title="Revenue by Region"
            )

            fig1.update_layout(
                height=380,
                showlegend=False,
                margin=dict(l=20, r=20, t=60, b=20),
                xaxis_title="Region",
                yaxis_title="Revenue",
                font=dict(family="Arial", size=12)
            )

            fig1.update_traces(
                marker_line_width=0,
                textposition="outside"
            )

            st.plotly_chart(
                fig1,
                use_container_width=True
            )

            # ---------------- CATEGORY DONUT ----------------

            st.markdown(
                '<div class="section-heading">Product Category Analysis</div>',
                unsafe_allow_html=True
            )

            category_data = (
                filtered_df.groupby(
                    "Product_Category",
                    as_index=False
                )["Revenue"].sum()
            )

            fig2 = px.pie(
                category_data,
                names="Product_Category",
                values="Revenue",
                hole=0.55,
                template=chart_template,
                title="Revenue Distribution by Category"
            )

            fig2.update_traces(
                textposition="inside",
                textinfo="percent+label"
            )

            fig2.update_layout(
                height=380,
                margin=dict(l=20, r=20, t=60, b=20)
            )

            st.plotly_chart(
                fig2,
                use_container_width=True
            )

            # ---------------- MONTHLY REVENUE TREND ----------------

            st.markdown(
                '<div class="section-heading">Monthly Revenue Trend</div>',
                unsafe_allow_html=True
            )

            monthly_df = filtered_df.dropna(
                subset=["Order_Date"]
            ).copy()

            monthly_df["Month"] = (
                monthly_df["Order_Date"]
                .dt.to_period("M")
                .dt.to_timestamp()
            )

            monthly_data = (
                monthly_df.groupby("Month", as_index=False)
                ["Revenue"].sum()
            )

            monthly_data = monthly_data.sort_values("Month")

            fig3 = px.area(
                monthly_data,
                x="Month",
                y="Revenue",
                markers=True,
                template=chart_template,
                title="Monthly Revenue Growth"
            )

            fig3.update_layout(
                height=380,
                margin=dict(l=20, r=20, t=60, b=20),
                xaxis_title="Month",
                yaxis_title="Revenue"
            )

            st.plotly_chart(
                fig3,
                use_container_width=True
            )

            # ---------------- PROFIT BY REGION ----------------

            st.markdown(
                '<div class="section-heading">Profitability Analysis</div>',
                unsafe_allow_html=True
            )

            profit_data = (
                filtered_df.groupby("Region", as_index=False)
                ["Profit"].sum()
            )

            fig4 = px.bar(
                profit_data,
                x="Region",
                y="Profit",
                color="Profit",
                template=chart_template,
                title="Profit by Region"
            )

            fig4.update_layout(
                height=350,
                margin=dict(l=20, r=20, t=60, b=20),
                xaxis_title="Region",
                yaxis_title="Profit"
            )

            st.plotly_chart(
                fig4,
                use_container_width=True
            )

            # ---------------- DATA TABLE ----------------

            with st.expander("View Detailed Sales Data"):

                st.dataframe(
                    filtered_df,
                    use_container_width=True,
                    hide_index=True
                )

            st.caption(
                f"Dashboard last refreshed: "
                f"{datetime.now().strftime('%d %b %Y, %I:%M %p')}"
            )

# ============================================================
# 8. AI-POWERED BUSINESS INSIGHTS
# ============================================================

elif menu == "AI Insights":

    st.subheader("🤖 AI-Powered Business Insights")

    if "sales_data" not in st.session_state:

        st.warning("Please upload your dataset first.")

    else:

        df = st.session_state["sales_data"].copy()

        # Business KPI Calculations
        total_revenue = df["Revenue"].sum()

        total_profit = df["Profit"].sum()

        total_orders = df["Order_ID"].nunique()

        profit_margin = (
            total_profit / total_revenue * 100
            if total_revenue != 0 else 0
        )

        # Regional Analysis
        region_data = (
            df.groupby("Region")["Revenue"]
            .sum()
            .sort_values(ascending=False)
        )

        # Product Analysis
        product_data = (
            df.groupby("Product")["Revenue"]
            .sum()
            .sort_values(ascending=False)
        )

        # Best Performing Region and Product
        best_region = (
            region_data.idxmax()
            if not region_data.empty else "N/A"
        )

        best_product = (
            product_data.idxmax()
            if not product_data.empty else "N/A"
        )

        # Business Summary
        st.subheader("Business Performance Summary")

        col1, col2, col3 = st.columns(3)

        col1.metric(
            "Total Revenue",
            f"₹{total_revenue:,.0f}"
        )

        col2.metric(
            "Total Profit",
            f"₹{total_profit:,.0f}"
        )

        col3.metric(
            "Total Orders",
            f"{total_orders:,}"
        )

        st.divider()

        # Management Insights
        st.subheader("Management Insights")

        st.success(
            f"🏆 Best Performing Region: {best_region}"
        )

        st.info(
            f"📦 Top Revenue-Generating Product: {best_product}"
        )

        if profit_margin >= 20:

            st.success(
                "Profitability is healthy based on a 20% benchmark."
            )

        else:

            st.warning(
                "Profit margin is below the 20% benchmark. Review costs and discounts."
            )

        st.divider()

        # Regional Revenue Analysis
        st.subheader("Regional Revenue Analysis")

        st.dataframe(
            region_data.reset_index(),
            use_container_width=True,
            hide_index=True
        )

        # AI Management Report
        st.subheader("Generate AI Management Report")

        if client is None:

            st.warning(
                "Groq API key is not configured. "
                "Add GROQ_API_KEY to .streamlit/secrets.toml "
                "to enable AI insights."
            )

        else:

            if st.button(
                "✨ Generate AI Insights",
                type="primary"
            ):

                business_data = f"""
                Total Orders: {total_orders}
                Total Revenue: ₹{total_revenue:,.2f}
                Total Profit: ₹{total_profit:,.2f}
                Profit Margin: {profit_margin:.2f}%

                Revenue by Region:
                {region_data.to_string()}

                Revenue by Product:
                {product_data.head(5).to_string()}
                """

                prompt = f"""
                Analyse the following verified sales business data.

                {business_data}

                Prepare a professional management report covering:

                1. Overall Business Performance
                2. Revenue and Profit Analysis
                3. Regional Performance
                4. Top Performing Products
                5. Business Risks and Areas of Concern
                6. Three Practical Recommendations

                Use professional business language.
                Use only the supplied figures.
                Do not invent statistics.
                """

                try:

                    with st.spinner(
                        "AI is analysing your business data..."
                    ):

                        response = client.chat.completions.create(
                            model="llama-3.3-70b-versatile",
                            messages=[
                                {
                                    "role": "system",
                                    "content": (
                                        "You are an experienced "
                                        "Business Intelligence "
                                        "and MIS Analyst."
                                    )
                                },
                                {
                                    "role": "user",
                                    "content": prompt
                                }
                            ],
                            temperature=0.3,
                            max_completion_tokens=1500
                        )

                        ai_report = (
                            response.choices[0].message.content
                        )

                    st.success(
                        "AI Management Report Generated!"
                    )

                    st.markdown(ai_report)

                    st.download_button(
                        label="📥 Download AI Report",
                        data=ai_report,
                        file_name="AI_Management_Report.txt",
                        mime="text/plain"
                    )

                except Exception as e:

                    st.error(
                        f"AI generation failed: {e}"
                    )


# ============================================================
# 9. AUTOMATED MIS REPORT GENERATION
# ============================================================

elif menu == "Reports":

    st.subheader("📊 Automated MIS Report Generation")

    if "sales_data" not in st.session_state:

        st.warning(
            "Please upload your dataset first from Data Upload."
        )

    else:

        df = st.session_state["sales_data"].copy()

        st.success(
            "Dataset ready for report generation!"
        )

        # ----------------------------------------------------
        # KPI CALCULATIONS
        # ----------------------------------------------------

        total_orders = df["Order_ID"].nunique()

        total_revenue = df["Revenue"].sum()

        total_profit = df["Profit"].sum()

        profit_margin = (
            total_profit / total_revenue * 100
            if total_revenue != 0 else 0
        )

        # ----------------------------------------------------
        # KPI SUMMARY
        # ----------------------------------------------------

        kpi_summary = pd.DataFrame({
            "Metric": [
                "Total Orders",
                "Total Revenue",
                "Total Profit",
                "Profit Margin"
            ],
            "Value": [
                total_orders,
                total_revenue,
                total_profit,
                round(profit_margin, 2)
            ]
        })

        # ----------------------------------------------------
        # REGIONAL ANALYSIS
        # ----------------------------------------------------

        region_summary = (
            df.groupby("Region", as_index=False)
            .agg(
                Total_Revenue=("Revenue", "sum"),
                Total_Profit=("Profit", "sum"),
                Total_Orders=("Order_ID", "nunique")
            )
        )

        # ----------------------------------------------------
        # CATEGORY ANALYSIS
        # ----------------------------------------------------

        category_summary = (
            df.groupby("Product_Category", as_index=False)
            .agg(
                Total_Revenue=("Revenue", "sum"),
                Total_Profit=("Profit", "sum"),
                Total_Quantity=("Quantity", "sum")
            )
        )

        # ----------------------------------------------------
        # REPORT PREVIEW
        # ----------------------------------------------------

        st.subheader("Management Report Preview")

        col1, col2, col3, col4 = st.columns(4)

        col1.metric(
            "Total Orders",
            f"{total_orders:,}"
        )

        col2.metric(
            "Total Revenue",
            f"₹{total_revenue:,.0f}"
        )

        col3.metric(
            "Total Profit",
            f"₹{total_profit:,.0f}"
        )

        col4.metric(
            "Profit Margin",
            f"{profit_margin:.2f}%"
        )

        st.divider()

        # KPI Summary Table
        st.subheader("KPI Summary")

        st.dataframe(
            kpi_summary,
            use_container_width=True,
            hide_index=True
        )

        # Regional Summary Table
        st.subheader("Regional Performance")

        st.dataframe(
            region_summary,
            use_container_width=True,
            hide_index=True
        )

        # Category Summary Table
        st.subheader("Product Category Performance")

        st.dataframe(
            category_summary,
            use_container_width=True,
            hide_index=True
        )

        st.divider()

        # ----------------------------------------------------
        # EXCEL REPORT GENERATION
        # ----------------------------------------------------

        st.subheader("📥 Download Excel Management Report")

        excel_buffer = BytesIO()

        with pd.ExcelWriter(
            excel_buffer,
            engine="openpyxl"
        ) as writer:

            kpi_summary.to_excel(
                writer,
                sheet_name="KPI Summary",
                index=False
            )

            region_summary.to_excel(
                writer,
                sheet_name="Regional Analysis",
                index=False
            )

            category_summary.to_excel(
                writer,
                sheet_name="Category Analysis",
                index=False
            )

            df.to_excel(
                writer,
                sheet_name="Cleaned Data",
                index=False
            )

        excel_data = excel_buffer.getvalue()

        st.download_button(
            label="⬇️ Download MIS Excel Report",
            data=excel_data,
            file_name="Automated_MIS_Report.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True
        )

        st.success(
            "Your Excel management report is ready!"
        )

        st.info(
            "PDF generation has been disabled. "
            "All report data is available in the Excel workbook."
        )
