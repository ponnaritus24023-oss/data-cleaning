# ============================================================
# DATA CLEANING & REPORTING AUTOMATION
# ============================================================

import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


# ============================================================
# 1. CONFIGURATION
# ============================================================

INPUT_FILE = "raw_data.csv"

OUTPUT_FOLDER = "automated_report"

CLEANED_FILE = os.path.join(
    OUTPUT_FOLDER,
    "cleaned_data.csv"
)

REPORT_FILE = os.path.join(
    OUTPUT_FOLDER,
    "automated_report.xlsx"
)

CHART_FOLDER = os.path.join(
    OUTPUT_FOLDER,
    "charts"
)


# Create output folders
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(CHART_FOLDER, exist_ok=True)


# ============================================================
# 2. LOAD DATA
# ============================================================

print("=" * 60)
print("DATA CLEANING & REPORTING AUTOMATION")
print("=" * 60)

print("\nLoading dataset...")

df = pd.read_csv(INPUT_FILE)

print("Dataset loaded successfully.")

print("\nOriginal dataset shape:")
print(df.shape)

print("\nFirst 5 rows:")
print(df.head())


# ============================================================
# 3. STANDARDIZE COLUMN NAMES
# ============================================================

print("\nStandardizing column names...")

df.columns = (
    df.columns
    .str.strip()
    .str.lower()
    .str.replace(" ", "_")
)

print("Columns:")
print(df.columns.tolist())


# ============================================================
# 4. CHECK DATA QUALITY BEFORE CLEANING
# ============================================================

print("\n" + "=" * 60)
print("DATA QUALITY CHECK BEFORE CLEANING")
print("=" * 60)

print("\nMissing values:")
print(df.isnull().sum())

print("\nDuplicate rows:")
print(df.duplicated().sum())

print("\nData types:")
print(df.dtypes)


# Store original statistics
original_rows = len(df)
original_columns = len(df.columns)
original_duplicates = df.duplicated().sum()
original_missing = df.isnull().sum().sum()


# ============================================================
# 5. REMOVE DUPLICATE ROWS
# ============================================================

print("\nRemoving duplicate rows...")

before_duplicates = len(df)

df = df.drop_duplicates()

after_duplicates = len(df)

duplicates_removed = (
    before_duplicates - after_duplicates
)

print(
    f"Duplicates removed: {duplicates_removed}"
)


# ============================================================
# 6. REMOVE EXTRA SPACES FROM TEXT
# ============================================================

print("\nCleaning text fields...")

text_columns = df.select_dtypes(
    include=["object"]
).columns

for column in text_columns:

    df[column] = (
        df[column]
        .astype(str)
        .str.strip()
    )


# ============================================================
# 7. STANDARDIZE TEXT VALUES
# ============================================================

# Example: Gender
if "gender" in df.columns:

    df["gender"] = (
        df["gender"]
        .str.lower()
        .replace({
            "m": "male",
            "f": "female",
            "man": "male",
            "woman": "female"
        })
    )

# Example: City
if "city" in df.columns:

    df["city"] = (
        df["city"]
        .str.title()
    )


# ============================================================
# 8. CONVERT NUMERIC COLUMNS
# ============================================================

numeric_columns = [
    "age",
    "sales",
    "quantity",
    "price"
]

for column in numeric_columns:

    if column in df.columns:

        df[column] = pd.to_numeric(
            df[column],
            errors="coerce"
        )


# ============================================================
# 9. CONVERT DATE COLUMNS
# ============================================================

date_columns = [
    "date",
    "order_date",
    "transaction_date"
]

for column in date_columns:

    if column in df.columns:

        df[column] = pd.to_datetime(
            df[column],
            errors="coerce"
        )


# ============================================================
# 10. HANDLE INVALID AGE VALUES
# ============================================================

if "age" in df.columns:

    # Ages below 0 or above 100 are considered invalid
    invalid_age = (
        (df["age"] < 0) |
        (df["age"] > 100)
    )

    df.loc[invalid_age, "age"] = np.nan


# ============================================================
# 11. HANDLE NEGATIVE SALES
# ============================================================

if "sales" in df.columns:

    negative_sales = df["sales"] < 0

    df.loc[
        negative_sales,
        "sales"
    ] = np.nan


# ============================================================
# 12. HANDLE MISSING VALUES
# ============================================================

print("\nHandling missing values...")

for column in df.columns:

    if pd.api.types.is_numeric_dtype(
        df[column]
    ):

        # Fill numeric missing values
        # with median
        median_value = df[column].median()

        df[column] = df[column].fillna(
            median_value
        )

    else:

        # Fill text missing values
        # with most common value
        mode_value = df[column].mode()

        if len(mode_value) > 0:

            df[column] = df[column].fillna(
                mode_value[0]
            )

        else:

            df[column] = df[column].fillna(
                "Unknown"
            )


# ============================================================
# 13. FINAL DATA QUALITY CHECK
# ============================================================

print("\n" + "=" * 60)
print("DATA QUALITY CHECK AFTER CLEANING")
print("=" * 60)

print("\nMissing values:")
print(df.isnull().sum())

print("\nDuplicate rows:")
print(df.duplicated().sum())

print("\nCleaned dataset shape:")
print(df.shape)


# ============================================================
# 14. CREATE DERIVED COLUMNS
# ============================================================

# Create Year and Month from Date
if "date" in df.columns:

    df["year"] = df["date"].dt.year

    df["month"] = df["date"].dt.month

    df["month_name"] = (
        df["date"].dt.month_name()
    )


# ============================================================
# 15. GENERATE SUMMARY STATISTICS
# ============================================================

print("\nGenerating summary statistics...")

summary_statistics = df.describe(
    include="all"
).transpose()

print(summary_statistics)


# ============================================================
# 16. CREATE SALES SUMMARY
# ============================================================

sales_summary = pd.DataFrame()

if "sales" in df.columns:

    total_sales = df["sales"].sum()

    average_sales = df["sales"].mean()

    minimum_sales = df["sales"].min()

    maximum_sales = df["sales"].max()

    sales_summary = pd.DataFrame({
        "Metric": [
            "Total Sales",
            "Average Sales",
            "Minimum Sales",
            "Maximum Sales"
        ],
        "Value": [
            total_sales,
            average_sales,
            minimum_sales,
            maximum_sales
        ]
    })

    print("\nSales Summary:")
    print(sales_summary)


# ============================================================
# 17. CITY-WISE SALES REPORT
# ============================================================

city_sales = pd.DataFrame()

if (
    "city" in df.columns
    and "sales" in df.columns
):

    city_sales = (
        df.groupby("city")["sales"]
        .sum()
        .reset_index()
        .sort_values(
            "sales",
            ascending=False
        )
    )

    print("\nCity-wise Sales:")
    print(city_sales)


# ============================================================
# 18. GENDER-WISE REPORT
# ============================================================

gender_summary = pd.DataFrame()

if (
    "gender" in df.columns
    and "sales" in df.columns
):

    gender_summary = (
        df.groupby("gender")["sales"]
        .agg([
            "count",
            "sum",
            "mean"
        ])
        .reset_index()
    )

    print("\nGender-wise Sales:")
    print(gender_summary)


# ============================================================
# 19. MONTHLY SALES REPORT
# ============================================================

monthly_sales = pd.DataFrame()

if (
    "date" in df.columns
    and "sales" in df.columns
):

    monthly_sales = (
        df.groupby(
            df["date"].dt.to_period("M")
        )["sales"]
        .sum()
        .reset_index()
    )

    monthly_sales["date"] = (
        monthly_sales["date"]
        .astype(str)
    )

    print("\nMonthly Sales:")
    print(monthly_sales)


# ============================================================
# 20. CREATE CHART - SALES BY CITY
# ============================================================

if not city_sales.empty:

    plt.figure(figsize=(10, 6))

    plt.bar(
        city_sales["city"],
        city_sales["sales"],
        color="steelblue"
    )

    plt.title("Sales by City")

    plt.xlabel("City")

    plt.ylabel("Total Sales")

    plt.xticks(rotation=45)

    plt.tight_layout()

    city_chart = os.path.join(
        CHART_FOLDER,
        "sales_by_city.png"
    )

    plt.savefig(city_chart)

    plt.close()

    print(
        f"\nCity chart saved: {city_chart}"
    )


# ============================================================
# 21. CREATE CHART - MONTHLY SALES
# ============================================================

if not monthly_sales.empty:

    plt.figure(figsize=(12, 6))

    plt.plot(
        monthly_sales["date"],
        monthly_sales["sales"],
        marker="o",
        color="green"
    )

    plt.title("Monthly Sales Trend")

    plt.xlabel("Month")

    plt.ylabel("Sales")

    plt.xticks(rotation=45)

    plt.grid(True)

    plt.tight_layout()

    monthly_chart = os.path.join(
        CHART_FOLDER,
        "monthly_sales_trend.png"
    )

    plt.savefig(monthly_chart)

    plt.close()

    print(
        f"Monthly chart saved: {monthly_chart}"
    )


# ============================================================
# 22. CREATE CHART - GENDER SALES
# ============================================================

if not gender_summary.empty:

    plt.figure(figsize=(8, 5))

    plt.bar(
        gender_summary["gender"],
        gender_summary["sum"],
        color=[
            "orange",
            "purple"
        ]
    )

    plt.title("Sales by Gender")

    plt.xlabel("Gender")

    plt.ylabel("Total Sales")

    plt.tight_layout()

    gender_chart = os.path.join(
        CHART_FOLDER,
        "sales_by_gender.png"
    )

    plt.savefig(gender_chart)

    plt.close()

    print(
        f"Gender chart saved: {gender_chart}"
    )


# ============================================================
# 23. CREATE DATA QUALITY REPORT
# ============================================================

final_missing = df.isnull().sum().sum()

final_duplicates = df.duplicated().sum()

cleaning_report = pd.DataFrame({

    "Metric": [
        "Original Rows",
        "Original Columns",
        "Original Missing Values",
        "Original Duplicate Rows",
        "Duplicates Removed",
        "Final Rows",
        "Final Columns",
        "Final Missing Values",
        "Final Duplicate Rows"
    ],

    "Value": [
        original_rows,
        original_columns,
        original_missing,
        original_duplicates,
        duplicates_removed,
        len(df),
        len(df.columns),
        final_missing,
        final_duplicates
    ]
})

print("\nData Cleaning Report:")
print(cleaning_report)


# ============================================================
# 24. SAVE CLEANED DATA
# ============================================================

df.to_csv(
    CLEANED_FILE,
    index=False
)

print(
    f"\nCleaned data saved to: {CLEANED_FILE}"
)


# ============================================================
# 25. GENERATE AUTOMATED EXCEL REPORT
# ============================================================

print("\nGenerating Excel report...")

with pd.ExcelWriter(
    REPORT_FILE,
    engine="openpyxl"
) as writer:

    # Cleaned dataset
    df.to_excel(
        writer,
        sheet_name="Cleaned Data",
        index=False
    )

    # Data quality report
    cleaning_report.to_excel(
        writer,
        sheet_name="Data Quality",
        index=False
    )

    # Summary statistics
    summary_statistics.to_excel(
        writer,
        sheet_name="Statistics"
    )

    # Sales summary
    if not sales_summary.empty:

        sales_summary.to_excel(
            writer,
            sheet_name="Sales Summary",
            index=False
        )

    # City sales
    if not city_sales.empty:

        city_sales.to_excel(
            writer,
            sheet_name="City Sales",
            index=False
        )

    # Gender summary
    if not gender_summary.empty:

        gender_summary.to_excel(
            writer,
            sheet_name="Gender Summary",
            index=False
        )

    # Monthly sales
    if not monthly_sales.empty:

        monthly_sales.to_excel(
            writer,
            sheet_name="Monthly Sales",
            index=False
        )


# ============================================================
# 26. FINAL MESSAGE
# ============================================================

print("\n" + "=" * 60)
print("AUTOMATION COMPLETED SUCCESSFULLY")
print("=" * 60)

print(f"""
Output files:

1. Cleaned dataset:
   {CLEANED_FILE}

2. Excel report:
   {REPORT_FILE}

3. Charts:
   {CHART_FOLDER}

The workflow has automatically:
- Cleaned column names
- Removed duplicates
- Handled missing values
- Standardized text
- Converted numeric fields
- Converted date fields
- Removed invalid values
- Generated summary statistics
- Generated sales reports
- Generated visualizations
- Exported the cleaned dataset
- Created an automated Excel report
""")
