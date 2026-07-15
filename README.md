# 📊 BSH Corporate Pricing Strategy & BI Dashboard

![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
![Data Engineering](https://img.shields.io/badge/Data%20Engineering-ETL%20Pipeline-orange.svg)
![Business Intelligence](https://img.shields.io/badge/Business%20Intelligence-Dashboarding-success.svg)
![Statistical Analysis](https://img.shields.io/badge/Analysis-Welch's%20T--Test-purple.svg)

## 📌 Executive Summary
This project is an end-to-end **Data Engineering and Business Intelligence (BI)** solution designed to analyze the internal pricing strategy and potential market cannibalization between two sister brands: **Bosch** and **Siemens** (under the BSH Group). 

Rather than just scraping data, this system builds a comprehensive data cube, applies statistical hypothesis testing, and generates a fully automated, executive-ready Excel BI Dashboard with advanced charts and variance analysis.

## 💼 The Business Case (Cannibalization & JND)
In consumer durables, the **Just Noticeable Difference (JND)** threshold for cross-price elasticity is scientifically accepted around **7.5%**. 
* If the price gap between a premium brand (Siemens) and a mass-market brand (Bosch) in the same segment falls below 7.5%, the brands risk **cannibalizing** each other's sales.
* This tool automatically flags these risk segments and proves whether the pricing differences are statistically significant or just random data fluctuations.

## 🚀 Core Features

1. **Automated ETL Pipeline:** * Dynamically scrapes live pricing and stock data from BSH web architectures (SPA and Static).
   * Parses domain dynamically and stores historical data in a relational **SQLite 3** database.
2. **Statistical Engine:** * Uses `SciPy` to perform **Welch's T-Tests** on category prices to determine if brand positioning is statistically significant (p < 0.05).
   * Calculates standard deviation (σ) to measure portfolio depth and price elasticity.
3. **Automated BI Dashboard:** * Utilizes `OpenPyXL` to generate a multi-sheet Excel report.
   * Auto-generates Bar/Line charts (Positioning, Premium %, Portfolio Depth, P-Value).
   * Includes automated "Manager Insights" written in professional advisory jargon.

## 🛠️ Technology Stack
* **Web Scraping:** `requests`, `BeautifulSoup4`
* **Data Manipulation:** `pandas`, `numpy`
* **Statistical Modeling:** `scipy`
* **Database:** `sqlite3`
* **Reporting & Visualization:** `openpyxl`

📂 Project Architecture
bsh-price-tracker/
│
├── data/
│   └── bsh_fiyat_veritabani.db       # Auto-generated SQLite Database
│
├── reports/
│   └── BSH_Kurumsal_BI_Raporu.xlsx   # Auto-generated BI Dashboard
│
├── src/
│   ├── __init__.py
│   ├── data_collector.py             # Web scraping logic & dynamic routing
│   ├── database_manager.py           # SQL queries and table creation
│   └── excel_report.py               # Pandas pivots, SciPy stats, and OpenPyXL charts
│
├── main.py                           # Application entry point & Category configs
└── README.md                         # Project documentation

## 📊 Power BI Executive Dashboard
![BSH Pricing Dashboard](bsh_powerbi_report.png)



https://github.com/user-attachments/assets/4a6a70e2-75cb-4e56-8974-4532cf785321




