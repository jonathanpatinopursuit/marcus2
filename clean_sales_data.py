import pandas as pd

# Load dataset
SRC = 'data/Sample - Superstore.csv'
OUT = 'output/Sample - Superstore - cleaned.csv'
df = pd.read_csv(SRC, encoding='latin1')

# 1. Remove exact duplicate rows
df = df.drop_duplicates()

# 2. Convert date columns to datetime (tolerate mixed/bad formats instead of raising)
df['Order Date'] = pd.to_datetime(df['Order Date'], format='%m/%d/%Y', errors='coerce')
df['Ship Date'] = pd.to_datetime(df['Ship Date'], format='%m/%d/%Y', errors='coerce')

# 3. Format discount as percentage for readability
df['Discount %'] = (df['Discount'] * 100).round(1)

# 4. Flag negative profit orders (not an error—just loss-making sales)
df['Loss_Flag'] = df['Profit'] < 0

# 5. Check for missing values and report which columns/rows are affected
missing = df.isnull().sum()
missing = missing[missing > 0]
if not missing.empty:
    print("Missing values by column:")
    print(missing)
    bad_dates = df['Order Date'].isnull() | df['Ship Date'].isnull()
    if bad_dates.any():
        print(f"\n{bad_dates.sum()} row(s) had unparseable dates and were dropped.")
        df = df[~bad_dates]
else:
    print("No missing values found.")

# 6. Save cleaned dataset
df.to_csv(OUT, index=False)
print(f"Cleaned dataset: {df.shape[0]} rows, {df.shape[1]} columns")
