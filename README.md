# 💰 Automated Financial Portfolio & Expense Intelligence Dashboard

A beginner-friendly Python project that reads raw financial transactions, analyzes spending patterns, and generates:
- A **styled Excel report** with multiple sheets (Raw Data, Monthly Summary, Category Breakdown, KPIs, Outliers)
- An **HTML dashboard** with KPI cards, monthly trends, and spending insights

![Python](https://img.shields.io/badge/Python-3.8+-blue?logo=python)
![Pandas](https://img.shields.io/badge/Pandas-Data%20Analysis-green)
![NumPy](https://img.shields.io/badge/NumPy-Scientific%20Computing-orange)

---

## 📸 Dashboard Preview

After running the script, open `output/dashboard.html` in any browser to see:
- KPI summary cards (Income, Expenses, Savings, Savings Rate)
- Monthly income vs expense table
- Category-wise spending breakdown
- Outlier detection (unusually high expenses)

---

## 🛠️ Tech Stack

| Technology | Purpose |
|------------|---------|
| **Python OOP** | Clean code structure with classes: `Transaction`, `FinancialAnalyzer`, `ExcelExporter`, `HTMLReportGenerator` |
| **Pandas** | Data loading, cleaning, grouping, and aggregation |
| **NumPy** | Statistical calculations (std dev, percentiles, savings rate) |
| **openpyxl** | Excel file generation with multiple sheets |
| **HTML/CSS** | Responsive dashboard with grid layout and styled tables |

---

## 📁 Project Structure

```
expense-dashboard/
├── main.py                  # Main application (run this)
├── requirements.txt         # Python dependencies
├── README.md
├── data/
│   └── transactions.csv     # Sample transaction data (edit with your own!)
└── output/
    ├── financial_report.xlsx # Generated Excel report
    └── dashboard.html        # Generated HTML dashboard
```

---

## 🚀 How to Run

### 1. Clone the repository
```bash
git clone https://github.com/YOUR_USERNAME/expense-dashboard.git
cd expense-dashboard
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the script
```bash
python main.py
```

### 4. View the output
- Open `output/dashboard.html` in your browser
- Open `output/financial_report.xlsx` in Excel / Google Sheets

---

## 📊 Features

- ✅ **Data Cleaning** – Handles missing values, formats dates, normalizes categories
- ✅ **Monthly Summary** – Income vs Expense vs Savings per month
- ✅ **Category Breakdown** – Total, average, and count per spending category
- ✅ **KPI Calculation** – Savings rate, spending volatility, largest expense
- ✅ **Outlier Detection** – Flags top 10% unusual expenses using NumPy percentile
- ✅ **Excel Export** – Multi-sheet workbook with clean formatting
- ✅ **HTML Dashboard** – Responsive cards and tables, no JavaScript needed

---

## 🎯 How to Use Your Own Data

Edit `data/transactions.csv` with your own transactions. Keep this format:

```
Date,Category,Amount,Type,Description
2024-01-05,Salary,50000,Income,Monthly salary
2024-01-07,Groceries,2300,Expense,Big Bazaar weekly grocery
```

**Required columns:** `Date`, `Category`, `Amount`, `Type` (Income/Expense), `Description`

---

## 📝 Concepts Demonstrated

This project is built to showcase beginner-to-intermediate Python skills:
- **OOP** – Classes with properties, methods, encapsulation
- **File I/O** – Reading CSV/Excel, writing HTML and Excel files
- **Data Analysis** – groupby, aggregation, pivot operations
- **NumPy** – Vectorized operations, percentile, standard deviation
- **HTML/CSS** – Semantic HTML, CSS Grid, responsive design

---

## 📄 License

This project is open source and available under the [MIT License](LICENSE).

---

> Built with Python, Pandas, NumPy & ❤️
