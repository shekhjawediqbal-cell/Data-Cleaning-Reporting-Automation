# Data Cleaning & Reporting Automation

## Project Overview

This project automates data cleaning, data quality validation, Excel report generation, and visual analysis using Python.

The project uses an Excel dataset containing client information, contact names, departments, payment methods, revenue, profit, and profit margin.

## Objectives

- Clean inconsistent text and standardize values.
- Remove empty columns and duplicate records.
- Separate department and region information.
- Identify and preserve missing values.
- Validate profit margin calculations.
- Generate a formatted Excel output.
- Create a data-cleaning report and visual summaries.

## Technologies Used

- Python
- Pandas
- NumPy
- Matplotlib
- OpenPyXL
- Visual Studio Code

## Data Cleaning Operations

1. Removed completely empty columns.
2. Removed duplicate rows.
3. Cleaned client names by removing parenthetical information.
4. Standardized contact names and whitespace.
5. Split the combined department field into Department and Region.
6. Standardized payment method names.
7. Checked missing values and numerical data types.
8. Validated profit margins against Profit / Revenue.
9. Exported missing values as `N/A` in the final Excel file.
10. Applied professional Excel formatting.

## Project Results

- Original records: 31
- Final records: 28
- Duplicate rows removed: 3
- Final columns: 9
- Profit margin mismatches: 0 among 26 eligible rows

Missing financial values were preserved rather than replaced with invented values.

## Visualizations

The project generates these charts:

- `charts/payment_method_distribution.png`
- `charts/revenue_by_department.png`
- `charts/profit_by_department.png`

## Output Files

- `output/cleaned_data.xlsx` — cleaned and formatted dataset.
- `reports/data_cleaning_report.txt` — automated data quality report.
- `charts/` — generated visualizations.

## How to Run

1. Open the project folder in Visual Studio Code.
2. Activate the project's virtual environment.
3. Install the required packages if they are not already installed.
4. Run the cleaning script:

   `python clean_data.py`

5. Run the chart generation script:

   `python create_charts.py`

The generated files are saved in the output, reports, and charts folders.

## Conclusion

This project demonstrates how Python can automate repetitive Excel data-cleaning tasks, validate financial data, generate structured reports, and produce visual summaries for business analysis.
