import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from datetime import datetime

try:
    from groq import Groq
except ImportError:
    Groq = None

# ============================================================
# 1. PAGE CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="MIS Intelligence | Business Reporting",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# 2. THEME / GLOBAL STYLING
# ============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}
.stApp { background: #F4F7FB; color: #17243A; }
.block-container { padding: 1.7rem 2.2rem 2.5rem; max-width: 1600px; }
[data-testid="stSidebar"] {
    background: #14233B;
    border-right: 1px solid #223653;
}
[data-testid="stSidebar"] * { color: #EAF1FC !important; }
[data-testid="stSidebar"] [data-testid="stRadio"] label {
    background: transparent;
    border-radius: 9px;
    padding: 8px 10px;
}
[data-testid="stSidebar"] [data-testid="stRadio"] label:hover {
    background: #243B5D;
}
[data-testid="stSidebar"] hr { border-color: #344966; }
.brand {
    padding: 8px 4px 22px;
    border-bottom: 1px solid #344966;
    margin-bottom: 22px;
}
.brand-mark {
    display: inline-block; background: #2F80ED; color: white;
    border-radius: 10px; padding: 9px 12px; font-size: 20px;
    margin-bottom: 12px;
}
.brand-name { font-size: 19px; font-weight: 800; color: #FFFFFF; }
.brand-sub { font-size: 11px; color: #A9BAD1; letter-spacing: 1.3px; }
.sidebar-note {
    background: #1D3453; border: 1px solid #304968;
    border-radius: 10px; padding: 12px; color: #C7D5E8;
    font-size: 12px; line-height: 1.6; margin-top: 24px;
}
.page-eyebrow {
    color: #2F80ED; font-weight: 800; font-size: 11px;
    letter-spacing: 1.6px; text-transform: uppercase; margin-bottom: 6px;
}
.page-title { color: #17243A; font-size: 30px; font-weight: 800; margin: 0; }
.page-subtitle { color: #718096; font-size: 14px; margin-top: 5px; }
.top-line { border-bottom: 1px solid #E2E8F0; margin: 18px 0 24px; }
.section-title {
    color: #1D2D45; font-size: 17px; font-weight: 800;
    margin: 12px 0 12px;
}
.section-caption { color: #8492A6; font-size: 12px; margin-top: -8px; margin-bottom: 12px; }
.kpi {
    background: #FFFFFF; border: 1px solid #E5EAF2;
    border-radius: 13px; padding: 17px 18px 16px;
    box-shadow: 0 3px 12px rgba(31, 55, 85, .045);
    min-height: 122px;
}
.kpi-top { display:flex; justify-content:space-between; align-items:center; }
.kpi-label { color:#75849A; font-size:11px; font-weight:800; letter-spacing:.7px; }
.kpi-icon {
    background:#EAF2FF; color:#2F80ED; border-radius:9px;
    padding:6px 8px; font-size:16px;
}
.kpi-value { color:#17243A; font-size:25px; font-weight:800; margin-top:13px; }
.kpi-foot { color:#91A0B3; font-size:11px; margin-top:3px; }
.panel {
    background:#FFFFFF; border:1px solid #E5EAF2; border-radius:13px;
    padding:15px 16px; box-shadow:0 3px 12px rgba(31,55,85,.035);
}
div[data-testid="stMetric"] {
    background:#FFFFFF; border:1px solid #E5EAF2;
    border-radius:12px; padding:13px 15px;
}
.stButton > button, .stDownloadButton > button {
    border-radius:9px; font-weight:700; min-height:42px;
}
.stButton > button[kind="primary"] {
    background:#2F80ED; border-color:#2F80ED;
}
div[data-testid="stDataFrame"] {
    border:1px solid #E5EAF2; border-radius:10px; overflow:hidden;
}
div[data-testid="stFileUploader"] {
    background:#FFFFFF; border:1px dashed #A9BEDA;
    border-radius:12px; padding:12px;
}
div[data-testid="stAlert"] { border-radius:10px; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 3. GROQ CONFIGURATION
# ============================================================
client = None
try:
    if Groq is not None and "GROQ_API_KEY" in st.secrets:
        client = Groq(api_key=st.secrets["GROQ_API_KEY"])
except Exception:
    client = None

# ============================================================
# 4. SHARED HELPERS
# ============================================================
def page_header(eyebrow, title, subtitle):
    st.markdown(
        f'<div class="page-eyebrow">{eyebrow}</div>'
        f'<div class="page-title">{title}</div>'
        f'<div class="page-subtitle">{subtitle}</div>'
        '<div class="top-line"></div>',
        unsafe_allow_html=True
    )

def section_title(title, caption=None):
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)
    if caption:
        st.markdown(f'<div class="section-caption">{caption}</div>', unsafe_allow_html=True)

def kpi_card(col, label, value, icon, foot):
    with col:
        st.markdown(f"""
        <div class="kpi">
            <div class="kpi-top">
                <div class="kpi-label">{label}</div>
                <div class="kpi-icon">{icon}</div>
            </div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-foot">{foot}</div>
        </div>
        """, unsafe_allow_html=True)

def money(value):
    return f"₹{value:,.0f}"

def load_excel_report(kpi_summary, region_summary, category_summary, df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        kpi_summary.to_excel(writer, sheet_name="KPI Summary", index=False)
        region_summary.to_excel(writer, sheet_name="Regional Analysis", index=False)
        category_summary.to_excel(writer, sheet_name="Category Analysis", index=False)
        df.to_excel(writer, sheet_name="Cleaned Data", index=False)
    return buffer.getvalue()

# ============================================================
# 5. SIDEBAR NAVIGATION
# ============================================================
with st.sidebar:
    st.markdown("""
    <div class="brand">
        <div class="brand-mark">▦</div>
        <div class="brand-name">MIS Intelligence</div>
        <div class="brand-sub">BUSINESS REPORTING SUITE</div>
    </div>
    """, unsafe_allow_html=True)

    menu = st.radio(
        "WORKSPACE",
        ["Home", "Data Upload", "Dashboard", "AI Insights", "Reports"],
        label_visibility="visible"
    )

    st.markdown("---")
    if "sales_data" in st.session_state:
        row_count = len(st.session_state["sales_data"])
        st.markdown(
            f'<div class="sidebar-note"><b>DATASET CONNECTED</b><br>'
            f'{row_count:,} records available<br>Ready for analysis</div>',
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            '<div class="sidebar-note"><b>NO DATASET LOADED</b><br>'
            'Upload an Excel or CSV file to begin your analysis.</div>',
            unsafe_allow_html=True
        )
    st.caption("MIS Intelligence • Portfolio Project")

# ============================================================
# 6. HOME
# ============================================================
if menu == "Home":
    page_header(
        "WORKSPACE OVERVIEW",
        "Business Reporting, Simplified.",
        "An AI-assisted MIS workspace to transform raw business data into clear performance insights."
    )

    st.markdown("""
    <div style="background:linear-gradient(115deg,#172F50,#245B91);
                border-radius:16px;padding:28px 30px;color:white;margin-bottom:22px;">
        <div style="font-size:12px;letter-spacing:1.5px;font-weight:700;color:#BFD9F8;">
            AUTOMATED MIS REPORTING AGENT
        </div>
        <div style="font-size:27px;font-weight:800;margin-top:9px;">
            From business data to better decisions.
        </div>
        <div style="font-size:14px;color:#D7E5F5;margin-top:8px;max-width:720px;">
            Upload your sales data, monitor KPIs, explore regional and product performance,
            and generate management-ready reports from one workspace.
        </div>
    </div>
    """, unsafe_allow_html=True)

    section_title("Your reporting workflow", "A simple four-step process")
    steps = st.columns(4)
    workflow = [
        ("01", "Upload data", "Import Excel or CSV business data.", "📁"),
        ("02", "Validate", "Review records, duplicates and missing values.", "✓"),
        ("03", "Analyse", "Explore KPIs, trends and business performance.", "📊"),
        ("04", "Report", "Generate AI commentary and Excel reports.", "📑")
    ]
    for col, (number, title, desc, icon) in zip(steps, workflow):
        with col:
            st.markdown(f"""
            <div class="panel" style="min-height:155px;">
                <div style="font-size:23px;">{icon}</div>
                <div style="color:#2F80ED;font-size:11px;font-weight:800;margin-top:12px;">STEP {number}</div>
                <div style="font-size:16px;font-weight:800;color:#1D2D45;margin-top:4px;">{title}</div>
                <div style="font-size:12px;color:#8492A6;margin-top:5px;">{desc}</div>
            </div>
            """, unsafe_allow_html=True)

    st.write("")
    section_title("Workspace status")
    if "sales_data" in st.session_state:
        df = st.session_state["sales_data"]
        a, b, c = st.columns(3)
        a.metric("Dataset status", "Connected")
        b.metric("Available records", f"{len(df):,}")
        c.metric("Available fields", f"{df.shape[1]:,}")
        st.success("Your dataset is ready. Open Dashboard to explore business performance.")
    else:
        st.info("No dataset loaded yet. Open **Data Upload** from the sidebar to get started.")

# ============================================================
# 7. DATA UPLOAD AND VALIDATION
# ============================================================
elif menu == "Data Upload":
    page_header(
        "DATA MANAGEMENT",
        "Upload & Validate Data",
        "Import your business dataset and review its quality before analysis."
    )

    uploaded_file = st.file_uploader("Choose an Excel or CSV file", type=["csv", "xlsx"])

    if uploaded_file is not None:
        try:
            if uploaded_file.name.lower().endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)

            df.columns = df.columns.astype(str).str.strip()
            duplicates = int(df.duplicated().sum())
            df = df.drop_duplicates().copy()

            if "Order_Date" in df.columns:
                df["Order_Date"] = pd.to_datetime(df["Order_Date"], errors="coerce")

            required_columns = [
                "Order_ID", "Order_Date", "Region", "Product",
                "Product_Category", "Revenue", "Profit", "Quantity"
            ]
            missing_columns = [col for col in required_columns if col not in df.columns]

            if missing_columns:
                st.error(f"Missing required columns: {', '.join(missing_columns)}")
                st.caption("Required fields: " + ", ".join(required_columns))
            else:
                for col in ["Revenue", "Profit", "Quantity"]:
                    df[col] = pd.to_numeric(df[col], errors="coerce")

                st.session_state["sales_data"] = df
                missing_values = int(df.isna().sum().sum())

                st.success("File uploaded, cleaned and validated successfully.")
                section_title("Data quality overview")
                q1, q2, q3, q4 = st.columns(4)
                q1.metric("Clean records", f"{len(df):,}")
                q2.metric("Duplicates removed", f"{duplicates:,}")
                q3.metric("Missing values", f"{missing_values:,}")
                q4.metric("Total columns", f"{df.shape[1]:,}")

                section_title("Dataset preview", "First 10 records after duplicate removal")
                st.dataframe(df.head(10), use_container_width=True, hide_index=True)

                left, right = st.columns(2)
                with left:
                    section_title("Column structure")
                    schema_df = pd.DataFrame({
                        "Column": df.columns,
                        "Data type": df.dtypes.astype(str).values,
                        "Missing": df.isna().sum().values
                    })
                    st.dataframe(schema_df, use_container_width=True, hide_index=True)
                with right:
                    section_title("Numeric summary")
                    numeric_cols = df.select_dtypes(include="number").columns
                    if len(numeric_cols):
                        st.dataframe(df[numeric_cols].describe().T, use_container_width=True)
                    else:
                        st.info("No numeric columns available.")
        except Exception as e:
            st.error(f"File processing error: {e}")

    elif "sales_data" in st.session_state:
        st.info("A dataset is already loaded in this session.")
        st.metric("Loaded records", f"{len(st.session_state['sales_data']):,}")

# ============================================================
# 8. BUSINESS DASHBOARD
# ============================================================
elif menu == "Dashboard":
    page_header(
        "PERFORMANCE ANALYTICS",
        "Business Performance",
        "Track revenue, profitability and sales trends across your business."
    )

    if "sales_data" not in st.session_state:
        st.warning("Please upload your dataset first from Data Upload.")
    else:
        df = st.session_state["sales_data"].copy()
        if df.empty:
            st.warning("The uploaded dataset contains no records.")
        else:
            # Filter panel
            section_title("Filters", "Use the controls to focus the dashboard")
            filter_box = st.container()
            with filter_box:
                f1, f2, f3 = st.columns([1.2, 1.2, 1])
                regions = sorted(df["Region"].dropna().unique().tolist())
                selected_regions = f1.multiselect("Region", regions, default=regions)
                categories = sorted(df["Product_Category"].dropna().unique().tolist())
                selected_categories = f2.multiselect("Product category", categories, default=categories)
                products = sorted(df["Product"].dropna().unique().tolist())
                selected_products = f3.multiselect("Product", products, default=products)

            filtered_df = df[
                df["Region"].isin(selected_regions)
                & df["Product_Category"].isin(selected_categories)
                & df["Product"].isin(selected_products)
            ].copy()

            if "Order_Date" in filtered_df.columns:
                filtered_df["Order_Date"] = pd.to_datetime(filtered_df["Order_Date"], errors="coerce")
                valid_dates = filtered_df["Order_Date"].dropna()
                if not valid_dates.empty:
                    date_col1, date_col2 = st.columns([1, 3])
                    with date_col1:
                        date_range = st.date_input(
                            "Order date range",
                            value=(valid_dates.min().date(), valid_dates.max().date())
                        )
                    if isinstance(date_range, (tuple, list)) and len(date_range) == 2:
                        start_date, end_date = date_range
                        filtered_df = filtered_df[
                            filtered_df["Order_Date"].dt.date.between(start_date, end_date)
                        ]

            if filtered_df.empty:
                st.warning("No records match the selected filters. Adjust your filters.")
            else:
                total_orders = filtered_df["Order_ID"].nunique()
                total_revenue = filtered_df["Revenue"].sum()
                total_profit = filtered_df["Profit"].sum()
                total_quantity = filtered_df["Quantity"].sum()
                profit_margin = (total_profit / total_revenue * 100) if total_revenue else 0

                section_title("Key performance indicators", "A snapshot of the selected data")
                k1, k2, k3, k4 = st.columns(4)
                kpi_card(k1, "TOTAL REVENUE", money(total_revenue), "₹", "Revenue from selected records")
                kpi_card(k2, "TOTAL PROFIT", money(total_profit), "↗", "Profit from selected records")
                kpi_card(k3, "TOTAL ORDERS", f"{total_orders:,}", "▣", "Unique order IDs")
                kpi_card(k4, "PROFIT MARGIN", f"{profit_margin:.1f}%", "％", "Profit as a share of revenue")

                st.write("")
                # Revenue trend and category mix
                section_title("Revenue overview", "Understand when revenue was generated and which categories contributed")
                chart_left, chart_right = st.columns([1.65, 1])
                with chart_left:
                    monthly_df = filtered_df.dropna(subset=["Order_Date"]).copy()
                    if not monthly_df.empty:
                        monthly_df["Month"] = monthly_df["Order_Date"].dt.to_period("M").dt.to_timestamp()
                        monthly_data = monthly_df.groupby("Month", as_index=False)["Revenue"].sum().sort_values("Month")
                        fig = px.area(
                            monthly_data, x="Month", y="Revenue",
                            markers=True, template="plotly_white",
                            color_discrete_sequence=["#2F80ED"]
                        )
                        fig.update_traces(line_width=3, fill="tozeroy", hovertemplate="%{x|%b %Y}<br>Revenue: ₹%{y:,.0f}<extra></extra>")
                        fig.update_layout(
                            height=330, margin=dict(l=10, r=10, t=15, b=10),
                            xaxis_title="", yaxis_title="Revenue",
                            hovermode="x unified", paper_bgcolor="white",
                            plot_bgcolor="white"
                        )
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info("No valid dates are available for the trend chart.")
                with chart_right:
                    category_data = filtered_df.groupby("Product_Category", as_index=False)["Revenue"].sum()
                    fig = px.pie(
                        category_data, names="Product_Category", values="Revenue",
                        hole=.64, template="plotly_white",
                        color_discrete_sequence=["#2F80ED", "#59B4D9", "#6D7FE8", "#A6B8D4", "#37A58A"]
                    )
                    fig.update_traces(textposition="outside", textinfo="percent+label", hole=.64)
                    fig.update_layout(
                        height=330, margin=dict(l=5, r=5, t=15, b=10),
                        showlegend=False, paper_bgcolor="white"
                    )
                    st.plotly_chart(fig, use_container_width=True)

                # Regional and product performance
                section_title("Performance breakdown", "Compare regions and leading products")
                chart_a, chart_b = st.columns(2)
                with chart_a:
                    region_data = filtered_df.groupby("Region", as_index=False).agg(
                        Revenue=("Revenue", "sum"), Profit=("Profit", "sum")
                    ).sort_values("Revenue", ascending=True)
                    fig = px.bar(
                        region_data, x="Revenue", y="Region", orientation="h",
                        text_auto=".2s", template="plotly_white",
                        color_discrete_sequence=["#2F80ED"]
                    )
                    fig.update_traces(textposition="outside", marker_line_width=0)
                    fig.update_layout(height=330, margin=dict(l=5, r=30, t=15, b=10),
                                      xaxis_title="Revenue", yaxis_title="", showlegend=False)
                    st.plotly_chart(fig, use_container_width=True)
                with chart_b:
                    product_data = filtered_df.groupby("Product", as_index=False)["Revenue"].sum()
                    product_data = product_data.nlargest(8, "Revenue").sort_values("Revenue")
                    fig = px.bar(
                        product_data, x="Revenue", y="Product", orientation="h",
                        text_auto=".2s", template="plotly_white",
                        color_discrete_sequence=["#37A58A"]
                    )
                    fig.update_traces(textposition="outside", marker_line_width=0)
                    fig.update_layout(height=330, margin=dict(l=5, r=30, t=15, b=10),
                                      xaxis_title="Revenue", yaxis_title="", showlegend=False)
                    st.plotly_chart(fig, use_container_width=True)

                # Profit view
                section_title("Profitability by region")
                profit_data = filtered_df.groupby("Region", as_index=False)["Profit"].sum().sort_values("Profit")
                fig = px.bar(
                    profit_data, x="Region", y="Profit", text_auto=".2s",
                    template="plotly_white",
                    color="Profit", color_continuous_scale=["#E58A8A", "#F0B85A", "#37A58A"]
                )
                fig.update_layout(height=300, margin=dict(l=5, r=5, t=15, b=10),
                                  xaxis_title="", yaxis_title="Profit", coloraxis_showscale=False)
                st.plotly_chart(fig, use_container_width=True)

                with st.expander("View detailed sales records"):
                    st.dataframe(filtered_df, use_container_width=True, hide_index=True)

                st.caption(
                    f"Showing {len(filtered_df):,} records • Last refreshed "
                    f"{datetime.now().strftime('%d %b %Y, %I:%M %p')}"
                )

# ============================================================
# 9. AI-POWERED BUSINESS INSIGHTS
# ============================================================
elif menu == "AI Insights":
    page_header(
        "AI BUSINESS INTELLIGENCE",
        "Management Insights",
        "Translate verified business metrics into practical management observations."
    )

    if "sales_data" not in st.session_state:
        st.warning("Please upload your dataset first.")
    else:
        df = st.session_state["sales_data"].copy()
        total_revenue = df["Revenue"].sum()
        total_profit = df["Profit"].sum()
        total_orders = df["Order_ID"].nunique()
        profit_margin = (total_profit / total_revenue * 100) if total_revenue else 0

        region_data = df.groupby("Region")["Revenue"].sum().sort_values(ascending=False)
        product_data = df.groupby("Product")["Revenue"].sum().sort_values(ascending=False)
        best_region = region_data.idxmax() if not region_data.empty else "N/A"
        best_product = product_data.idxmax() if not product_data.empty else "N/A"

        a, b, c, d = st.columns(4)
        a.metric("Total revenue", money(total_revenue))
        b.metric("Total profit", money(total_profit))
        c.metric("Total orders", f"{total_orders:,}")
        d.metric("Profit margin", f"{profit_margin:.2f}%")

        section_title("Automated business observations")
        obs1, obs2 = st.columns(2)
        with obs1:
            st.markdown(f"""
            <div class="panel" style="min-height:120px;">
                <div style="font-size:11px;color:#8492A6;font-weight:800;">TOP REGION BY REVENUE</div>
                <div style="font-size:22px;color:#1D2D45;font-weight:800;margin-top:10px;">{best_region}</div>
                <div style="font-size:12px;color:#8492A6;">Highest total revenue among the regions in the dataset.</div>
            </div>
            """, unsafe_allow_html=True)
        with obs2:
            st.markdown(f"""
            <div class="panel" style="min-height:120px;">
                <div style="font-size:11px;color:#8492A6;font-weight:800;">TOP PRODUCT BY REVENUE</div>
                <div style="font-size:22px;color:#1D2D45;font-weight:800;margin-top:10px;">{best_product}</div>
                <div style="font-size:12px;color:#8492A6;">Highest total revenue among the products in the dataset.</div>
            </div>
            """, unsafe_allow_html=True)

        if profit_margin >= 20:
            st.success("Profit margin is at or above the 20% reference benchmark.")
        else:
            st.warning("Profit margin is below the 20% reference benchmark. Review cost and discount drivers.")

        section_title("Regional revenue analysis")
        region_table = region_data.rename("Revenue").reset_index()
        st.dataframe(region_table, use_container_width=True, hide_index=True)

        section_title("Generate AI management report")
        if client is None:
            st.info("Groq is not configured. Add GROQ_API_KEY to .streamlit/secrets.toml to enable generated insights.")
        else:
            if st.button("✨ Generate AI insights", type="primary"):
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

Use professional business language. Use only the supplied figures.
Do not invent statistics. Clearly distinguish observations from recommendations.
"""
                try:
                    with st.spinner("AI is analysing the supplied business metrics..."):
                        response = client.chat.completions.create(
                            model="llama-3.3-70b-versatile",
                            messages=[
                                {"role": "system", "content": "You are an experienced Business Intelligence and MIS Analyst."},
                                {"role": "user", "content": prompt}
                            ],
                            temperature=0.3,
                            max_completion_tokens=1500
                        )
                        ai_report = response.choices[0].message.content

                    st.session_state["ai_report"] = ai_report
                    st.success("AI management report generated.")
                except Exception as e:
                    st.error(f"AI generation failed: {e}")

            if "ai_report" in st.session_state:
                st.markdown('<div class="panel">', unsafe_allow_html=True)
                st.markdown(st.session_state["ai_report"])
                st.markdown('</div>', unsafe_allow_html=True)
                st.download_button(
                    "Download AI management report",
                    data=st.session_state["ai_report"],
                    file_name="AI_Management_Report.txt",
                    mime="text/plain"
                )

# ============================================================
# 10. AUTOMATED MIS REPORTS
# ============================================================
elif menu == "Reports":
    page_header(
        "REPORTING CENTRE",
        "MIS Report Generator",
        "Review management summaries and export a multi-sheet Excel report."
    )

    if "sales_data" not in st.session_state:
        st.warning("Please upload your dataset first from Data Upload.")
    else:
        df = st.session_state["sales_data"].copy()

        total_orders = df["Order_ID"].nunique()
        total_revenue = df["Revenue"].sum()
        total_profit = df["Profit"].sum()
        profit_margin = (total_profit / total_revenue * 100) if total_revenue else 0

        kpi_summary = pd.DataFrame({
            "Metric": ["Total Orders", "Total Revenue", "Total Profit", "Profit Margin"],
            "Value": [total_orders, total_revenue, total_profit, round(profit_margin, 2)]
        })
        region_summary = df.groupby("Region", as_index=False).agg(
            Total_Revenue=("Revenue", "sum"),
            Total_Profit=("Profit", "sum"),
            Total_Orders=("Order_ID", "nunique")
        ).sort_values("Total_Revenue", ascending=False)
        category_summary = df.groupby("Product_Category", as_index=False).agg(
            Total_Revenue=("Revenue", "sum"),
            Total_Profit=("Profit", "sum"),
            Total_Quantity=("Quantity", "sum")
        ).sort_values("Total_Revenue", ascending=False)

        r1, r2, r3, r4 = st.columns(4)
        r1.metric("Total orders", f"{total_orders:,}")
        r2.metric("Total revenue", money(total_revenue))
        r3.metric("Total profit", money(total_profit))
        r4.metric("Profit margin", f"{profit_margin:.2f}%")

        section_title("Regional performance")
        st.dataframe(region_summary, use_container_width=True, hide_index=True)

        section_title("Product category performance")
        st.dataframe(category_summary, use_container_width=True, hide_index=True)

        section_title("Export management report", "The workbook includes KPI, regional, category and cleaned-data sheets.")
        excel_data = load_excel_report(kpi_summary, region_summary, category_summary, df)
        st.download_button(
            "⬇️ Download MIS Excel report",
            data=excel_data,
            file_name="Automated_MIS_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
        st.caption("PDF export is currently disabled. All report data is included in the Excel workbook.")
