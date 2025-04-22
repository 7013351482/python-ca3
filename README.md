import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# Load the Excel file with error handling
file_path = "data set abhishek eu.xlsx"
try:
    df = pd.read_excel(file_path)
except FileNotFoundError:
    print(f"Error: File '{file_path}' not found. Please ensure it is in the correct directory.")
    exit(1)
except Exception as e:
    print(f"Error loading file: {e}")
    exit(1)

# Validate and rename columns
expected_columns = [
    'Store', 'Customer ID', 'Date', 'Region', 'State', 'City',
    'Category', 'Units Sold', 'Price per Unit', 'Total Sales',
    'Profit', 'Discount', 'Sales Channel'
]
if len(df.columns) != len(expected_columns):
    print(f"Error: Expected {len(expected_columns)} columns, but found {len(df.columns)}.")
    exit(1)
df.columns = expected_columns

# Convert 'Date' column to datetime with error handling
try:
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    if df['Date'].isna().any():
        print("Warning: Some dates could not be parsed and were set to NaT.")
except Exception as e:
    print(f"Error converting 'Date' column: {e}")
    exit(1)

# Add Month column for grouping
df['Month'] = df['Date'].dt.to_period('M').astype(str)

# Check for missing values
if df.isna().any().any():
    print("Warning: Dataset contains missing values. Filling numeric columns with 0 and categorical with 'Unknown'.")
    df.fillna({'Total Sales': 0, 'Units Sold': 0, 'Price per Unit': 0, 'Profit': 0, 'Discount': 0,
               'Store': 'Unknown', 'Customer ID': 'Unknown', 'Region': 'Unknown', 'State': 'Unknown',
               'City': 'Unknown', 'Category': 'Unknown', 'Sales Channel': 'Unknown'}, inplace=True)

# Set seaborn style for better aesthetics
sns.set_style("whitegrid")

# 1. Monthly Sales Trend
def plot_monthly_sales():
    plt.figure(figsize=(10, 6))
    monthly_sales = df.groupby('Month')['Total Sales'].sum().reset_index()
    sns.lineplot(data=monthly_sales, x='Month', y='Total Sales', marker='o')
    plt.title("Monthly Sales Trend", fontsize=14)
    plt.xlabel("Month", fontsize=12)
    plt.ylabel("Total Sales", fontsize=12)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig("monthly_sales_trend.png")
    plt.show()

# 2. Sales Channel Breakdown
def plot_sales_channel():
    plt.figure(figsize=(8, 5))
    sns.countplot(data=df, x='Sales Channel', palette='viridis')
    plt.title("Sales Channel Distribution", fontsize=14)
    plt.xlabel("Sales Channel", fontsize=12)
    plt.ylabel("Count", fontsize=12)
    plt.tight_layout()
    plt.savefig("sales_channel_distribution.png")
    plt.show()

# 3. Top Product Categories
def plot_top_categories():
    plt.figure(figsize=(10, 6))
    category_sales = df.groupby('Category')['Total Sales'].sum().sort_values(ascending=False).head(10)
    sns.barplot(x=category_sales.values, y=category_sales.index, palette='magma')
    plt.title("Top 10 Product Categories by Sales", fontsize=14)
    plt.xlabel("Total Sales", fontsize=12)
    plt.ylabel("Category", fontsize=12)
    plt.tight_layout()
    plt.savefig("top_categories.png")
    plt.show()

# 4. Sales by Region
def plot_sales_by_region():
    plt.figure(figsize=(10, 6))
    region_sales = df.groupby('Region')['Total Sales'].sum().sort_values(ascending=False)
    sns.barplot(x=region_sales.values, y=region_sales.index, palette='coolwarm')
    plt.title("Sales by Region", fontsize=14)
    plt.xlabel("Total Sales", fontsize=12)
    plt.ylabel("Region", fontsize=12)
    plt.tight_layout()
    plt.savefig("sales_by_region.png")
    plt.show()

# 5. Discount vs Units Sold
def plot_discount_vs_units():
    plt.figure(figsize=(8, 6))
    sns.scatterplot(data=df, x='Discount', y='Units Sold', alpha=0.6, hue='Sales Channel', size='Total Sales')
    plt.title("Discount vs Units Sold", fontsize=14)
    plt.xlabel("Discount", fontsize=12)
    plt.ylabel("Units Sold", fontsize=12)
    plt.tight_layout()
    plt.savefig("discount_vs_units.png")
    plt.show()

# Call the functions to show the plots
try:
    plot_monthly_sales()
    plot_sales_channel()
    plot_top_categories()
    plot_sales_by_region()
    plot_discount_vs_units()
except Exception as e:
    print(f"Error generating plots: {e}")
