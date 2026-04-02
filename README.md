# FMCG Supplier Performance Dashboard (Excel & Power BI)

## 📌 Project Overview
This is an end-to-end data analytics portfolio project that simulates a realistic Fast-Moving Consumer Goods (FMCG) supply chain. The project encompasses custom synthetic data generation in Python, automated KPI aggregation, an interactive Excel dashboard, and a complete Power BI implementation complete with DAX measures.

## ⚙️ Project Architecture & Workflow

### 1. Data Generation (Python)
To ensure the data mirrors a real-world Kaggle-style dataset, a Python script was developed to simulate **3,635 purchase orders** across **8 FMCG suppliers** over a **2-year period**. 

The dataset includes realistic anomalies and business constraints:
* **Realistic Supplier Mix:** Built-in vendor profiles (e.g., *FastFoil* is expensive but reliable at 96% On-Time Delivery; *EcoPalm* is cheap but problematic with 70% OTD and a 68 quality score).
* **Seasonal Price Drift:** Unit costs follow a `sin()` curve simulation (e.g., palm oil prices peak in the summer).
* **Random Noise:** Applied ±12% cost variance per order and ±2.5 days of lead time variance to simulate real-world supply chain friction.

### 2. KPI Aggregation & Logic
Before visualization, the core business logic was established to evaluate supplier health. The following metrics were calculated per supplier:
* Total Spend & Avg Unit Cost
* On-Time Delivery (OTD) Rate
* Quality Score & Defect Rate
* **Weighted Performance Score:** A custom metric heavily weighting reliability (40% OTD + 40% Quality + 20% Inverse Lead Time). 

### 3. Excel Workbook Implementation
The generated data flows into a heavily formatted, 3-sheet Excel workbook designed for immediate business use:
* **`Raw_Data`:** Houses all 3,635 rows. Features alternating row color-coding, frozen headers, and strictly formatted cost/date columns.
* **`KPI_Scorecard`:** A supplier scorecard utilizing conditional formatting (Green/Amber/Red thresholds for OTD and overall score). Includes a summary totals row using `SUM/AVERAGE` and a monthly trend table that feeds the dashboard visualizations.
* **`Dashboard`:** A presentation-ready sheet featuring 5 KPI cards connected via cross-sheet formulas, alongside 6 embedded dynamic charts (Bar, Column, Line, and Dual-Axis).

### 4. Power BI Integration
The raw data is also pulled into a complete Power BI dashboard (`.pbix`). The Power BI model features:
* A structured Data Model.
* Advanced DAX measures for dynamic time-intelligence and calculating the Weighted Performance Score.
* Native visuals optimized to highlight supply chain bottlenecks and vendor comparisons.

## 📂 Repository Structure
```text
├── build_dashboard.py           # Python script for realistic data generation & Excel formatting
├── FMCG_Supplier_Dashboard.xlsx # Final Excel Workbook (Raw Data, Scorecard, Dashboard)
└── FMCG_Supplier_Dashboard.pbix # Power BI Dashboard with DAX measures and visuals
```

## 🚀 How to Use
Explore the Data: Run build_dashboard.py to generate a fresh dataset and automatically build the Excel workbook.

View the Excel Dashboard: Open FMCG_Supplier_Dashboard.xlsx to view the Raw Data, Scorecard, and Dashboard sheets. Try adjusting the raw data to see the KPI cards and charts update automatically.

Explore the Power BI Model: Open FMCG_Supplier_Dashboard.pbix to interact with the dashboard, explore the data model, and review the underlying DAX measures.

## 💡 Key Learnings & Skills Demonstrated
Python: Pandas, Numpy, Synthetic Data Generation, Simulation Logic, automated Excel formatting.

Excel: Advanced Formulas (XLOOKUP, SUMIFS), Cross-sheet referencing, Conditional Formatting, Data Visualization.

Power BI: DAX logic, Data Modeling, Dashboard UX/UI Design.

Domain Knowledge: Supply Chain Analytics, Vendor Management, FMCG Operations.
