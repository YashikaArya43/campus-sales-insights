import pandas as pd
import sqlite3
import matplotlib.pyplot as plt
import seaborn as sns

# =============================================================================
# PHASE 1: DATA LOADING & EXPLORATION
# =============================================================================
print("=" * 50)
print("📊 PHASE 1: DATA LOADING & EXPLORATION")
print("=" * 50)

try:
    # Step 1: Load the Excel file
    df = pd.read_excel("Sales_Data_Pipeline.xlsx")
    print("✅ Excel file loaded successfully!")
    
    # Step 2: Preview the first 5 rows
    print("\n🔍 First 5 rows:")
    print(df.head())
    
    # Step 3: Shape of the dataset
    print(f"\n📐 Shape (Rows, Columns): {df.shape}")
    
    # Step 4: List of column names
    print("\n🧾 Column Names:")
    print(df.columns.tolist())
    
    # Step 5: Data types of each column
    print("\n🔠 Data Types:")
    print(df.dtypes)
    
    # Step 6: Check for missing values
    print("\n❓ Missing Values:")
    print(df.isnull().sum())
    
except FileNotFoundError:
    print("❌ Error: Sales_Data_Pipeline.xlsx not found!")
    print("Please ensure the file is in the same directory or provide the full path.")
    exit()
except Exception as e:
    print(f"❌ Error loading data: {e}")
    exit()

# =============================================================================
# PHASE 2: DATA CLEANING & TRANSFORMATION
# =============================================================================
print("\n" + "=" * 50)
print("🧹 PHASE 2: DATA CLEANING & TRANSFORMATION")
print("=" * 50)

try:
    # Convert 'Date' to datetime
    df['Date'] = pd.to_datetime(df['Date'])
    print("✅ Date column converted to datetime")
    
    # Add new time-related features
    df['Month'] = df['Date'].dt.strftime('%b')         # Jan, Feb, etc.
    df['Month_Num'] = df['Date'].dt.month
    df['Quarter'] = df['Date'].dt.quarter
    df['Year'] = df['Date'].dt.year
    print("✅ Time-related features added")
    
    # Standardize text formatting
    df['Product'] = df['Product'].str.title().str.strip()
    df['Category'] = df['Category'].str.title().str.strip()
    df['Region'] = df['Region'].str.title().str.strip()
    print("✅ Text formatting standardized")
    
    # Sort by date
    df = df.sort_values(by='Date')
    print("✅ Data sorted by date")
    
    # Check if any missing values remain
    print("\n🔍 Missing values after cleaning:")
    print(df.isnull().sum())
    
    # Show basic statistics
    print("\n📊 Revenue Statistics:")
    print(df['Revenue'].describe())
    
except Exception as e:
    print(f"❌ Error during data cleaning: {e}")
    exit()

# =============================================================================
# PHASE 3: DATA STORAGE & VISUALIZATION
# =============================================================================
print("\n" + "=" * 50)
print("💾 PHASE 3: DATA STORAGE & VISUALIZATION")
print("=" * 50)

try:
    # Store in SQLite database
    print("📂 Connecting to SQLite database...")
    conn = sqlite3.connect("sales_data.db")
    
    # Store the cleaned data
    df.to_sql("sales", conn, if_exists="replace", index=False)
    print("✅ Data successfully stored in SQLite database!")
    
    # Verify the data was stored correctly
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM sales")
    row_count = cursor.fetchone()[0]
    print(f"✅ Verified: {row_count} rows stored in database")
    
    # Show sample data from database
    cursor.execute("SELECT * FROM sales LIMIT 3")
    sample_data = cursor.fetchall()
    print("\n🔍 Sample data from database:")
    for row in sample_data:
        print(row)
    
    # Get column names from database
    cursor.execute("PRAGMA table_info(sales)")
    columns_info = cursor.fetchall()
    print(f"\n🧾 Database columns: {[col[1] for col in columns_info]}")
    
    conn.close()
    print("✅ Database connection closed")
    
except Exception as e:
    print(f"❌ Error during database operations: {e}")
    if 'conn' in locals():
        conn.close()

try:
    # Create and display boxplot
    print("\n📊 Creating revenue boxplot...")
    plt.figure(figsize=(10, 6))
    sns.boxplot(data=df, y='Revenue')
    plt.title("Revenue Distribution - Boxplot")
    plt.ylabel("Revenue ($)")
    
    # Add some statistics to the plot
    mean_revenue = df['Revenue'].mean()
    median_revenue = df['Revenue'].median()
    plt.axhline(y=mean_revenue, color='red', linestyle='--', alpha=0.7, label=f'Mean: ${mean_revenue:,.2f}')
    plt.axhline(y=median_revenue, color='green', linestyle='--', alpha=0.7, label=f'Median: ${median_revenue:,.2f}')
    plt.legend()
    
    plt.tight_layout()
    plt.show()
    print("✅ Boxplot displayed successfully")
    
except Exception as e:
    print(f"❌ Error creating visualization: {e}")

try:
    # Export cleaned DataFrame to Excel
    output_filename = "Cleaned_Sales_Data.xlsx"
    df.to_excel(output_filename, index=False)
    print(f"✅ Cleaned data exported to '{output_filename}'")
    
except Exception as e:
    print(f"❌ Error exporting to Excel: {e}")

# =============================================================================
# FINAL SUMMARY
# =============================================================================
print("\n" + "=" * 50)
print("📋 FINAL SUMMARY")
print("=" * 50)
print(f"✅ Original dataset shape: {df.shape}")
print(f"✅ Date range: {df['Date'].min()} to {df['Date'].max()}")
print(f"✅ Total revenue: ${df['Revenue'].sum():,.2f}")
print(f"✅ Average revenue per transaction: ${df['Revenue'].mean():,.2f}")
print(f"✅ Unique products: {df['Product'].nunique()}")
print(f"✅ Unique categories: {df['Category'].nunique()}")
print(f"✅ Unique regions: {df['Region'].nunique()}")
print("\n🎉 Data pipeline completed successfully!")
print("📢 This is the last line - Pipeline finished!")