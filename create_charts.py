
import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path

# Load the original data and prepare it for analysis
df = pd.read_excel(
    "data/Data Cleaning Start.xlsx",
    sheet_name="Sheet1 (2)",
    header=1
)

# Remove empty columns and duplicate rows
df = df.dropna(axis=1, how="all")
df = df.drop_duplicates().copy()

# Standardize payment method names
df["Payment"] = (
    df["Payment"]
    .astype("string")
    .str.strip()
    .str.title()
)

# Count payment methods, including missing values
payment_counts = df["Payment"].fillna("N/A").value_counts()

# Create and save the chart
chart_folder = Path("charts")
chart_folder.mkdir(parents=True, exist_ok=True)

plt.figure(figsize=(8, 5))
payment_counts.plot(kind="bar")

plt.title("Payment Method Distribution")
plt.xlabel("Payment Method")
plt.ylabel("Number of Records")
plt.xticks(rotation=0)
plt.tight_layout()

chart_path = chart_folder / "payment_method_distribution.png"
plt.savefig(chart_path, dpi=300, bbox_inches="tight")
plt.close()

print("Payment method chart created successfully!")
print("Saved to:", chart_path.resolve())
print("\nPayment method counts:")
print(payment_counts.to_string())

# Prepare department and region columns
split_columns = df["Department"].str.split(
    "_", n=1, expand=True
)

df["Department"] = split_columns[0].str.strip()
df["Region"] = split_columns[1].str.strip()

# Convert revenue to numeric values
df["Revenue"] = pd.to_numeric(
    df["Revenue"],
    errors="coerce"
)

# Calculate total revenue for each department
revenue_by_department = df.groupby(
    "Department"
)["Revenue"].sum(min_count=1).sort_values(ascending=False)

# Create and save the revenue chart
plt.figure(figsize=(9, 5))
revenue_by_department.plot(kind="bar")

plt.title("Total Revenue by Department")
plt.xlabel("Department")
plt.ylabel("Total Revenue")
plt.xticks(rotation=30, ha="right")
plt.tight_layout()

revenue_chart_path = chart_folder / "revenue_by_department.png"
plt.savefig(revenue_chart_path, dpi=300, bbox_inches="tight")
plt.close()

print("\nRevenue-by-department chart created successfully!")
print("Saved to:", revenue_chart_path.resolve())
print("\nTotal revenue by department:")
print(revenue_by_department.to_string())

# Convert profit to numeric values
df["Profit"] = pd.to_numeric(
    df["Profit"],
    errors="coerce"
)

# Calculate total profit for each department
profit_by_department = df.groupby(
    "Department"
)["Profit"].sum(min_count=1).sort_values(ascending=False)

# Create and save the profit chart
plt.figure(figsize=(9, 5))
profit_by_department.plot(kind="bar")

plt.title("Total Profit by Department")
plt.xlabel("Department")
plt.ylabel("Total Profit")
plt.xticks(rotation=30, ha="right")
plt.tight_layout()

profit_chart_path = chart_folder / "profit_by_department.png"
plt.savefig(profit_chart_path, dpi=300, bbox_inches="tight")
plt.close()

print("\nProfit-by-department chart created successfully!")
print("Saved to:", profit_chart_path.resolve())
print("\nTotal profit by department:")
print(profit_by_department.to_string())
