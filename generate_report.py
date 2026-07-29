import pandas as pd

SRC = '/Users/jonathan/Downloads/Sample - Superstore - cleaned.csv'
OUT = '/Users/jonathan/Downloads/sales_report.xlsx'

df = pd.read_csv(SRC, parse_dates=['Order Date', 'Ship Date'])
df['Month'] = df['Order Date'].dt.to_period('M').astype(str)

# 1. Sales by Region & Category
region_category = df.groupby(['Region', 'Category'])['Sales'].sum().reset_index()

# 2. Performance by Product, Region, Segment, Time
perf = df.groupby(['Product Name', 'Region', 'Segment', 'Month']).agg(
    Total_Sales=('Sales', 'sum'), Total_Profit=('Profit', 'sum')).reset_index()

# 3. Monthly summary
monthly_summary = df.groupby('Month').agg(Sales=('Sales', 'sum'), Profit=('Profit', 'sum')).reset_index()

# 4. Underperformers (bottom 10% by profit)
underperform = df.groupby(['Category', 'Region'])['Profit'].sum().reset_index()
underperform = underperform[underperform['Profit'] < underperform['Profit'].quantile(0.1)]

# 5. Discount impact on profit
discount_impact = df.groupby('Discount %').agg(
    Avg_Profit=('Profit', 'mean'), Orders=('Order ID', 'count')).reset_index()

# Export report
with pd.ExcelWriter(OUT) as writer:
    region_category.to_excel(writer, sheet_name='Region_Category', index=False)
    perf.to_excel(writer, sheet_name='Product_Performance', index=False)
    monthly_summary.to_excel(writer, sheet_name='Monthly_Trend', index=False)
    underperform.to_excel(writer, sheet_name='Underperformers', index=False)
    discount_impact.to_excel(writer, sheet_name='Discount_Impact', index=False)

print(f"Report written to {OUT}")
