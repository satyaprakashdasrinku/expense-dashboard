"""
Automated Financial Portfolio & Expense Intelligence Dashboard
A beginner-level Python project using OOP, Pandas, NumPy, Excel, and HTML.
"""

import pandas as pd
import numpy as np
from datetime import datetime
import os


# ======================== OOP CLASSES ========================

class Transaction:
    """Represents a single financial transaction."""

    def __init__(self, date, category, amount, description, txn_type):
        self.date = pd.to_datetime(date)
        self.category = category
        self.amount = float(amount)
        self.description = description
        self.txn_type = txn_type  # 'Income' or 'Expense'

    def __repr__(self):
        return f"Transaction({self.date.date()}, {self.category}, ₹{self.amount})"


class FinancialAnalyzer:
    """Loads, cleans, and analyzes transaction data."""

    def __init__(self, filepath):
        self.filepath = filepath
        self.df = None
        self.transactions = []

    def load_data(self):
        """Read data from Excel or CSV file."""
        if self.filepath.endswith('.xlsx'):
            self.df = pd.read_excel(self.filepath)
        else:
            self.df = pd.read_csv(self.filepath)
        print(f"✅ Loaded {len(self.df)} transactions from {self.filepath}")

    def clean_data(self):
        """Clean and prepare the data."""
        self.df['Date'] = pd.to_datetime(self.df['Date'])
        self.df['Amount'] = pd.to_numeric(self.df['Amount'], errors='coerce').fillna(0)
        self.df['Category'] = self.df['Category'].str.strip().str.title()
        self.df['Type'] = self.df['Type'].str.strip().str.title()
        self.df['Description'] = self.df['Description'].fillna('N/A')
        self.df['Month'] = self.df['Date'].dt.strftime('%Y-%m')

        # Create Transaction objects
        for _, row in self.df.iterrows():
            txn = Transaction(row['Date'], row['Category'],
                              row['Amount'], row['Description'], row['Type'])
            self.transactions.append(txn)

        print(f"✅ Cleaned data. {len(self.transactions)} valid transactions.")

    def get_monthly_summary(self):
        """Get income vs expense grouped by month."""
        summary = self.df.groupby(['Month', 'Type'])['Amount'].sum().unstack(fill_value=0)
        if 'Income' not in summary.columns:
            summary['Income'] = 0
        if 'Expense' not in summary.columns:
            summary['Expense'] = 0
        summary['Savings'] = summary['Income'] - summary['Expense']
        summary['Savings Rate (%)'] = np.where(
            summary['Income'] > 0,
            np.round((summary['Savings'] / summary['Income']) * 100, 1),
            0
        )
        return summary

    def get_category_breakdown(self):
        """Get total spending per category."""
        expenses = self.df[self.df['Type'] == 'Expense']
        breakdown = expenses.groupby('Category')['Amount'].agg(['sum', 'mean', 'count'])
        breakdown.columns = ['Total Spent', 'Avg per Transaction', 'Num Transactions']
        breakdown = breakdown.sort_values('Total Spent', ascending=False)
        return breakdown.round(2)

    def calculate_kpis(self):
        """Calculate key financial indicators using NumPy."""
        expenses = self.df[self.df['Type'] == 'Expense']['Amount'].values
        incomes = self.df[self.df['Type'] == 'Income']['Amount'].values

        kpis = {
            'total_income': float(np.sum(incomes)),
            'total_expense': float(np.sum(expenses)),
            'net_savings': float(np.sum(incomes) - np.sum(expenses)),
            'avg_daily_expense': float(np.mean(expenses)) if len(expenses) > 0 else 0,
            'max_expense': float(np.max(expenses)) if len(expenses) > 0 else 0,
            'expense_std_dev': float(np.std(expenses)) if len(expenses) > 0 else 0,
            'num_transactions': len(self.df),
        }

        kpis['savings_rate'] = round(
            (kpis['net_savings'] / kpis['total_income'] * 100), 1
        ) if kpis['total_income'] > 0 else 0

        # Flag: spending is "volatile" if std dev > 50% of mean
        kpis['spending_volatile'] = kpis['expense_std_dev'] > (0.5 * kpis['avg_daily_expense'])

        return kpis

    def detect_outliers(self):
        """Find unusually high expenses using NumPy percentile."""
        expenses = self.df[self.df['Type'] == 'Expense'].copy()
        if len(expenses) == 0:
            return pd.DataFrame()
        threshold = np.percentile(expenses['Amount'].values, 90)
        outliers = expenses[expenses['Amount'] > threshold]
        return outliers[['Date', 'Category', 'Amount', 'Description']]


class ExcelExporter:
    """Exports analysis results to a formatted Excel workbook."""

    def __init__(self, output_path):
        self.output_path = output_path

    def export(self, raw_df, monthly_summary, category_breakdown, kpis, outliers):
        """Write all data to a multi-sheet Excel file."""
        with pd.ExcelWriter(self.output_path, engine='openpyxl') as writer:

            # Sheet 1: Raw Data
            raw_df.to_excel(writer, sheet_name='Raw Data', index=False)

            # Sheet 2: Monthly Summary
            monthly_summary.to_excel(writer, sheet_name='Monthly Summary')

            # Sheet 3: Category Breakdown
            category_breakdown.to_excel(writer, sheet_name='Category Breakdown')

            # Sheet 4: KPIs
            kpi_df = pd.DataFrame([kpis]).T
            kpi_df.columns = ['Value']
            kpi_df.to_excel(writer, sheet_name='KPIs')

            # Sheet 5: Outlier Expenses
            if len(outliers) > 0:
                outliers.to_excel(writer, sheet_name='Outlier Expenses', index=False)

        print(f"✅ Excel report saved to: {self.output_path}")


class HTMLReportGenerator:
    """Generates a styled HTML dashboard."""

    def __init__(self, output_path):
        self.output_path = output_path

    def generate(self, monthly_summary, category_breakdown, kpis, outliers):
        """Build and save the HTML dashboard."""

        def fmt(val):
            return f"₹{val:,.2f}"

        # Build category table rows
        cat_rows = ""
        for cat, row in category_breakdown.iterrows():
            cat_rows += f"""
            <tr>
                <td>{cat}</td>
                <td>{fmt(row['Total Spent'])}</td>
                <td>{fmt(row['Avg per Transaction'])}</td>
                <td>{int(row['Num Transactions'])}</td>
            </tr>"""

        # Build monthly table rows
        month_rows = ""
        for month, row in monthly_summary.iterrows():
            color = "green" if row.get('Savings', 0) >= 0 else "red"
            month_rows += f"""
            <tr>
                <td>{month}</td>
                <td>{fmt(row.get('Income', 0))}</td>
                <td>{fmt(row.get('Expense', 0))}</td>
                <td style="color:{color};font-weight:bold;">{fmt(row.get('Savings', 0))}</td>
                <td>{row.get('Savings Rate (%)', 0)}%</td>
            </tr>"""

        # Build outlier rows
        outlier_rows = ""
        if len(outliers) > 0:
            for _, row in outliers.iterrows():
                outlier_rows += f"""
                <tr>
                    <td>{row['Date'].strftime('%Y-%m-%d') if hasattr(row['Date'], 'strftime') else row['Date']}</td>
                    <td>{row['Category']}</td>
                    <td style="color:red;font-weight:bold;">{fmt(row['Amount'])}</td>
                    <td>{row['Description']}</td>
                </tr>"""

        status_color = "green" if kpis['net_savings'] >= 0 else "red"
        volatile_text = "⚠️ High" if kpis['spending_volatile'] else "✅ Stable"

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Expense Intelligence Dashboard</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #f0f2f5; color: #333; padding: 20px; }}
        h1 {{ text-align: center; color: #2c3e50; margin-bottom: 5px; }}
        .subtitle {{ text-align: center; color: #7f8c8d; margin-bottom: 30px; }}
        .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 30px; }}
        .kpi-card {{ background: white; border-radius: 10px; padding: 20px; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }}
        .kpi-card h3 {{ color: #7f8c8d; font-size: 14px; margin-bottom: 8px; }}
        .kpi-card .value {{ font-size: 24px; font-weight: bold; }}
        .section {{ background: white; border-radius: 10px; padding: 20px; margin-bottom: 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }}
        .section h2 {{ color: #2c3e50; margin-bottom: 15px; border-bottom: 2px solid #3498db; padding-bottom: 8px; }}
        table {{ width: 100%; border-collapse: collapse; }}
        th {{ background: #3498db; color: white; padding: 10px; text-align: left; }}
        td {{ padding: 10px; border-bottom: 1px solid #eee; }}
        tr:hover {{ background: #f8f9fa; }}
        .footer {{ text-align: center; color: #95a5a6; margin-top: 30px; font-size: 13px; }}
    </style>
</head>
<body>
    <h1>💰 Expense Intelligence Dashboard</h1>
    <p class="subtitle">Auto-generated on {datetime.now().strftime('%d %B %Y, %I:%M %p')}</p>

    <!-- KPI Cards -->
    <div class="kpi-grid">
        <div class="kpi-card">
            <h3>Total Income</h3>
            <div class="value" style="color:#27ae60;">{fmt(kpis['total_income'])}</div>
        </div>
        <div class="kpi-card">
            <h3>Total Expenses</h3>
            <div class="value" style="color:#e74c3c;">{fmt(kpis['total_expense'])}</div>
        </div>
        <div class="kpi-card">
            <h3>Net Savings</h3>
            <div class="value" style="color:{status_color};">{fmt(kpis['net_savings'])}</div>
        </div>
        <div class="kpi-card">
            <h3>Savings Rate</h3>
            <div class="value">{kpis['savings_rate']}%</div>
        </div>
        <div class="kpi-card">
            <h3>Largest Expense</h3>
            <div class="value" style="color:#e74c3c;">{fmt(kpis['max_expense'])}</div>
        </div>
        <div class="kpi-card">
            <h3>Spending Volatility</h3>
            <div class="value">{volatile_text}</div>
        </div>
    </div>

    <!-- Monthly Summary -->
    <div class="section">
        <h2>📅 Monthly Summary</h2>
        <table>
            <tr><th>Month</th><th>Income</th><th>Expenses</th><th>Savings</th><th>Savings Rate</th></tr>
            {month_rows}
        </table>
    </div>

    <!-- Category Breakdown -->
    <div class="section">
        <h2>📊 Spending by Category</h2>
        <table>
            <tr><th>Category</th><th>Total Spent</th><th>Avg / Transaction</th><th>Count</th></tr>
            {cat_rows}
        </table>
    </div>

    <!-- Outlier Expenses -->
    <div class="section">
        <h2>🚨 Unusual High Expenses (Top 10%)</h2>
        <table>
            <tr><th>Date</th><th>Category</th><th>Amount</th><th>Description</th></tr>
            {outlier_rows if outlier_rows else '<tr><td colspan="4" style="text-align:center;">No outliers detected 🎉</td></tr>'}
        </table>
    </div>

    <p class="footer">Built with Python, Pandas, NumPy &amp; ❤️ | Expense Intelligence Dashboard</p>
</body>
</html>"""

        with open(self.output_path, 'w', encoding='utf-8') as f:
            f.write(html)
        print(f"✅ HTML dashboard saved to: {self.output_path}")


# ======================== MAIN EXECUTION ========================

def main():
    # --- Paths ---
    base_dir = os.path.dirname(os.path.abspath(__file__))
    data_file = os.path.join(base_dir, 'data', 'transactions.csv')
    excel_output = os.path.join(base_dir, 'output', 'financial_report.xlsx')
    html_output = os.path.join(base_dir, 'output', 'dashboard.html')

    # --- Step 1: Load & Clean ---
    analyzer = FinancialAnalyzer(data_file)
    analyzer.load_data()
    analyzer.clean_data()

    # --- Step 2: Analyze ---
    monthly_summary = analyzer.get_monthly_summary()
    category_breakdown = analyzer.get_category_breakdown()
    kpis = analyzer.calculate_kpis()
    outliers = analyzer.detect_outliers()

    print("\n📊 KEY PERFORMANCE INDICATORS:")
    for key, val in kpis.items():
        print(f"   {key:>25s} : {val}")

    # --- Step 3: Export to Excel ---
    excel_exporter = ExcelExporter(excel_output)
    excel_exporter.export(analyzer.df, monthly_summary, category_breakdown, kpis, outliers)

    # --- Step 4: Generate HTML Dashboard ---
    html_gen = HTMLReportGenerator(html_output)
    html_gen.generate(monthly_summary, category_breakdown, kpis, outliers)

    print("\n🎉 All done! Open the files in the 'output/' folder.")


if __name__ == '__main__':
    main()
