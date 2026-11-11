# Automated MIS Reporting Agent

An AI-powered Management Information System (MIS) Reporting Agent built with Python, Streamlit, Pandas, Plotly, and Groq LLM. The application automates sales data analysis, KPI monitoring, interactive dashboard generation, AI-driven business insights, and Excel report generation.

## Project Overview

The Automated MIS Reporting Agent is designed to simplify business reporting by transforming raw sales data into meaningful management insights.

Users can upload Excel or CSV sales datasets, validate data quality, explore business performance through interactive dashboards, generate AI-assisted insights, and export analytical reports.

## Key Features

- **Data Upload:** Import Excel and CSV datasets.
- **Data Validation:** Check required columns and identify data quality issues.
- **Data Processing:** Process and analyse sales information using Pandas.
- **Interactive Dashboard:** Monitor revenue, profit, orders, and profit margins.
- **Business Analytics:** Explore monthly sales trends, regional performance, and product category analysis.
- **AI-Powered Insights:** Generate management summaries and business recommendations using Groq LLM.
- **Dynamic Filtering:** Analyse performance by region, product category, product, and date.
- **Excel Reporting:** Export analytical summaries and reports.

## Technology Stack

| Technology | Purpose |
|---|---|
| Python | Core programming |
| Pandas | Data processing and analysis |
| NumPy | Numerical operations |
| Streamlit | Web application interface |
| Plotly | Interactive visualizations |
| Groq LLM | AI-powered business insights |
| OpenPyXL | Excel report generation |

## Application Modules

1. **Home:** Application overview and navigation.
2. **Data Upload:** Upload and validate sales datasets.
3. **Dashboard:** Interactive KPI cards, charts, and business performance analysis.
4. **AI Insights:** AI-generated business insights and management summaries.
5. **Reports:** Analytical summaries and downloadable Excel reports.

## Project Structure

```text
Automated-MIS-Reporting-Agent/
│
├── app.py
├── requirements.txt
├── README.md
├── .gitignore
│
├── screenshots/
│   ├── dashboard.png
│   ├── data-upload.png
│   ├── ai-insights.png
│   └── reports.png
│
├── data/
└── outputs/
```

## Installation and Setup

### 1. Clone the repository

```bash
git clone https://github.com/YOUR-USERNAME/Automated-MIS-Reporting-Agent.git
```

### 2. Navigate to the project directory

```bash
cd Automated-MIS-Reporting-Agent
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Configure Groq API Key

Create a file named `secrets.toml` inside the `.streamlit` directory:

```toml
GROQ_API_KEY = "your_groq_api_key"
```

Never upload your API key or `secrets.toml` file to GitHub.

### 5. Run the application

```bash
python -m streamlit run app.py
```

## Sample Dataset

The project uses a sales dataset containing fields such as:

- Order ID
- Order Date
- Region
- Salesperson
- Product Category
- Product
- Quantity
- Unit Price
- Revenue
- Cost
- Profit
- Order Status
- Payment Status

## Business Value

This project demonstrates how Python-based automation and AI-assisted analytics can support:

- Faster MIS reporting.
- Improved visibility into business performance.
- KPI-based decision-making.
- Identification of sales and profit trends.
- Automated management reporting.
- Data-driven business insights.

## Future Enhancements

- Support for multiple business dataset types.
- Intelligent column mapping.
- Automated dataset classification.
- Advanced data quality checks.
- Natural-language data querying.
- Anomaly detection and forecasting.
- Cloud deployment and scheduled reporting.

## Author

**Vijay Tiwari**

Project developed as a practical application of Python, Business Intelligence, Data Analytics, and Generative AI.

---

**Note:** This version currently focuses on sales MIS reporting with the expected dataset structure. Broader multi-dataset support is planned as a future enhancement.
