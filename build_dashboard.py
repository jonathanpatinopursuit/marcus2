import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils.dataframe import dataframe_to_rows

SRC = '/Users/jonathan/Downloads/Sample - Superstore - cleaned.csv'
OUT = '/Users/jonathan/Downloads/sales_dashboard.xlsx'
C1_PNG = '/Users/jonathan/Downloads/dash_region_category.png'
C2_PNG = '/Users/jonathan/Downloads/dash_top_products.png'
C3_PNG = '/Users/jonathan/Downloads/dash_monthly_trend.png'

df = pd.read_csv(SRC, parse_dates=['Order Date', 'Ship Date'])
df['Month'] = df['Order Date'].dt.to_period('M').astype(str)

region_pivot = df.pivot_table(values='Sales', index='Region', columns='Category', aggfunc='sum', fill_value=0)
top_products = df.groupby('Product Name')['Sales'].sum().sort_values(ascending=False).head(10).reset_index()
monthly_summary = df.groupby('Month').agg(Sales=('Sales', 'sum'), Profit=('Profit', 'sum')).reset_index()

wb = Workbook()
ws = wb.active
ws.title = "Sales Dashboard"

# Write Region x Category table (top-left)
ws.append(["Sales by Region & Category"])
for r in dataframe_to_rows(region_pivot.reset_index(), index=False, header=True):
    ws.append(r)

# Chart 1: Region x Category
plt.figure(figsize=(6, 3))
region_pivot.plot(kind='bar', stacked=True, ax=plt.gca())
plt.title('Sales by Region and Category')
plt.tight_layout()
plt.savefig(C1_PNG, dpi=150)
plt.close()
ws.add_image(XLImage(C1_PNG), 'H2')

# Chart 2: Top Products
plt.figure(figsize=(9, 5))
plt.bar(top_products['Product Name'], top_products['Sales'])
plt.xticks(rotation=45, ha='right', fontsize=8)
plt.title('Top 10 Products by Sales')
plt.tight_layout()
plt.savefig(C2_PNG, dpi=150)
plt.close()
ws.add_image(XLImage(C2_PNG), 'H26')

# Chart 3: Monthly Trend
plt.figure(figsize=(9, 4))
plt.plot(monthly_summary['Month'], monthly_summary['Sales'], marker='o')
plt.xticks(rotation=45, ha='right', fontsize=7)
plt.title('Monthly Sales Trend')
plt.tight_layout()
plt.savefig(C3_PNG, dpi=150)
plt.close()
ws.add_image(XLImage(C3_PNG), 'H66')

wb.save(OUT)
print(f"Dashboard written to {OUT}")
