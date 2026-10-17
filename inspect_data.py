
import pandas as pd

file_path = "data/Data Cleaning Start.xlsx"

df = pd.read_excel(
    file_path,
    sheet_name="Sheet1 (2)",
    header=1
)

print("\n--- CORRECTED COLUMN NAMES ---")
print(df.columns.tolist())

print("\n--- FIRST 10 ROWS ---")
print(df.head(10).to_string(index=False))

print("\n--- DATASET SHAPE ---")
print(df.shape)

print("\n--- MISSING VALUES ---")
print(df.isnull().sum())

print("\n--- DUPLICATE ROWS ---")
print(df.duplicated().sum())

print("\n--- FIRST COLUMN CHECK ---")
print(df.iloc[:, 0].to_string(index=False))

print("\n--- EMPTY COLUMN CHECK ---")
print(df.isnull().all())

# Remove columns that contain only missing values
df = df.dropna(axis=1, how="all")

print("\n--- DATASET AFTER REMOVING EMPTY COLUMNS ---")
print(df.head(5).to_string(index=False))

print("\n--- UPDATED COLUMN NAMES ---")
print(df.columns.tolist())

print("\n--- UPDATED DATASET SHAPE ---")
print(df.shape)

print("\n--- MISSING VALUE SUMMARY ---")

missing_summary = pd.DataFrame({
    "Missing Count": df.isnull().sum(),
    "Missing Percentage": (df.isnull().mean() * 100).round(2)
})

print(missing_summary)

print("\n--- ROWS WITH MISSING VALUES ---")

missing_rows = df[df.isnull().any(axis=1)]

print(missing_rows.to_string(index=False))

print("\n--- MISSING VALUES BY ROW ---")
print(df.isnull().sum(axis=1))

print("\n--- COMPLETE ROWS WITH MISSING VALUES ---")

missing_rows = df[df.isnull().any(axis=1)]

for index, row in missing_rows.iterrows():
    print(f"\nRow index: {index}")
    print(row.to_string())

print("\n--- PROFIT MARGIN FORMULA CHECK ---")

df["Calculated Margin"] = df["Profit"] / df["Revenue"]

comparison = df[["Profit", "Revenue", "Profit Margin", "Calculated Margin"]]

print(comparison.head(10).to_string(index=False))

print("\n--- FULL DATASET FORMULA VALIDATION ---")

valid_rows = df[
    df["Revenue"].notna() &
    df["Profit"].notna() &
    df["Revenue"].ne(0)
].copy()

valid_rows["Calculated Margin"] = (
    valid_rows["Profit"] / valid_rows["Revenue"]
)

valid_rows["Margin Matches"] = (
    (valid_rows["Profit Margin"] - valid_rows["Calculated Margin"])
    .abs() < 0.000001
)

print("Rows checked:", len(valid_rows))
print("Matching margins:", valid_rows["Margin Matches"].sum())
print("Non-matching margins:", (~valid_rows["Margin Matches"]).sum())

print("\n--- NON-MATCHING ROWS ---")
print(
    valid_rows.loc[
        ~valid_rows["Margin Matches"],
        ["Client", "Revenue", "Profit", "Profit Margin", "Calculated Margin"]
    ].to_string(index=False)
)

print("\n--- ROWS WITH MISSING REVENUE ---")

print(
    df.loc[
        df["Revenue"].isna(),
        ["Date", "Client", "Department", "Payment", "Revenue", "Profit", "Profit Margin"]
    ].to_string(index=True)
)

print("\n--- DUPLICATE RECORDS ---")

duplicate_rows = df[df.duplicated(keep=False)]

print(duplicate_rows.to_string(index=False))

print("\nTotal rows involved in duplicates:", len(duplicate_rows))
print("Rows that would be removed:", df.duplicated().sum())

print("\n--- REMOVE DUPLICATE ROWS ---")

rows_before = len(df)

df = df.drop_duplicates().copy()

rows_after = len(df)

print("Rows before removing duplicates:", rows_before)
print("Rows after removing duplicates:", rows_after)
print("Duplicate rows removed:", rows_before - rows_after)

print("\n--- DATASET SHAPE AFTER REMOVING DUPLICATES ---")
print(df.shape)

print("\n--- CLEAN CLIENT NAMES ---")

df["Client"] = (
    df["Client"]
    .str.replace(r"\s*\([^)]*\)", "", regex=True)
    .str.strip()
)

print(df["Client"].head(10).to_string(index=False))

print("\n--- CLEAN CONTACT NAMES ---")

df["Contact"] = (
    df["Contact"]
    .str.strip()
    .str.replace(r"\s+", " ", regex=True)
    .str.title()
)

print(df["Contact"].head(10).to_string(index=False))

print("\n--- SPLIT DEPARTMENT AND REGION ---")

split_columns = df["Department"].str.split("_", n=1, expand=True)

df["Department"] = split_columns[0].str.strip()
df["Region"] = split_columns[1].str.strip()

print(df[["Department", "Region"]].head(10).to_string(index=False))

print("\n--- UPDATED COLUMNS ---")
print(df.columns.tolist())

print("\n--- CLEAN PAYMENT METHODS ---")

df["Payment"] = (
    df["Payment"]
    .astype("string")
    .str.strip()
    .str.title()
)

print(df["Payment"].value_counts(dropna=False))

print("\n--- NUMERIC DATA TYPE CHECK ---")

print(df[["Revenue", "Profit", "Profit Margin"]].dtypes)

print("\n--- NUMERIC COLUMN SUMMARY ---")
print(df[["Revenue", "Profit", "Profit Margin"]].describe())

print("\n--- REMOVE TEMPORARY VALIDATION COLUMN ---")

if "Calculated Margin" in df.columns:
    df = df.drop(columns=["Calculated Margin"])

print("Final columns:", df.columns.tolist())
print("Current dataset shape:", df.shape)
