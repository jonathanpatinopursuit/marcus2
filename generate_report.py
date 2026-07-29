import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage

SRC = '/Users/jonathan/Downloads/Sample - Superstore - cleaned.csv'
OUT = '/Users/jonathan/Downloads/sales_report.xlsx'
CHART_PNG = '/Users/jonathan/Downloads/top_products_chart.png'
REGION_CHART_PNG = '/Users/jonathan/Downloads/region_category_chart.png'
MONTHLY_CHART_PNG = '/Users/jonathan/Downloads/monthly_trend_chart.png'

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

# 6. Region x Category pivot (wide format, complements Region_Category)
region_pivot = df.pivot_table(values='Sales', index='Region', columns='Category', aggfunc='sum', fill_value=0)

# 7. Top 10 products by sales
top_products = df.groupby('Product Name')['Sales'].sum().sort_values(ascending=False).head(10).reset_index()

# Export report
with pd.ExcelWriter(OUT, engine='openpyxl') as writer:
    region_category.to_excel(writer, sheet_name='Region_Category', index=False)
    perf.to_excel(writer, sheet_name='Product_Performance', index=False)
    monthly_summary.to_excel(writer, sheet_name='Monthly_Trend', index=False)
    underperform.to_excel(writer, sheet_name='Underperformers', index=False)
    discount_impact.to_excel(writer, sheet_name='Discount_Impact', index=False)
    region_pivot.to_excel(writer, sheet_name='Region_by_Category')
    top_products.to_excel(writer, sheet_name='Top_Products', index=False)

# Add bar chart image for Top Products
plt.figure(figsize=(10, 6))
plt.bar(top_products['Product Name'], top_products['Sales'])
plt.xticks(rotation=45, ha='right', fontsize=8)
plt.title('Top 10 Products by Sales')
plt.tight_layout()
plt.savefig(CHART_PNG, dpi=150)
plt.close()

# Region by Category chart (stacked bar)
plt.figure(figsize=(8, 4))
region_pivot.plot(kind='bar', stacked=True, ax=plt.gca())
plt.title('Sales by Region and Category')
plt.tight_layout()
plt.savefig(REGION_CHART_PNG, dpi=150)
plt.close()

# Monthly Trend chart (line)
plt.figure(figsize=(8, 4))
plt.plot(monthly_summary['Month'], monthly_summary['Sales'], marker='o')
plt.xticks(rotation=45, ha='right', fontsize=7)
plt.title('Monthly Sales Trend')
plt.tight_layout()
plt.savefig(MONTHLY_CHART_PNG, dpi=150)
plt.close()

wb = load_workbook(OUT)

ws = wb['Top_Products']
ws.add_image(XLImage(CHART_PNG), 'E2')

ws2 = wb['Region_by_Category']
ws2.add_image(XLImage(REGION_CHART_PNG), 'H2')

ws3 = wb['Monthly_Trend']
ws3.add_image(XLImage(MONTHLY_CHART_PNG), 'D2')

wb.save(OUT)

print(f"Report written to {OUT}")
