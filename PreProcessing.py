import pandas as pd

# Path to the ODS file
file_path = 'DailySheet.ods'

# Load the sheets
sheets_data = pd.read_excel(file_path, engine='odf', sheet_name=None, skiprows=1)

# Extract sheets into DataFrames
expenses_df = sheets_data['Expenses']
stock_df = sheets_data['Stock']
sales_df = sheets_data['Sales']

# Clean data and ensure columns are correctly named
expenses_df.columns = expenses_df.columns.str.strip()  # Trim spaces from column names
stock_df.columns = stock_df.columns.str.strip()
sales_df.columns = sales_df.columns.str.strip()

# Calculate total expenses from "Amount Paid" column (if it exists)
total_expenses = expenses_df['Amount Paid'].sum() if 'Amount Paid' in expenses_df.columns else 0

# Calculate total revenue from "Cost Added" column in the Stock sheet
total_revenue = sales_df['Amount'].sum() if 'Amount' in sales_df.columns else 0

# Calculate profit
profit = total_revenue - total_expenses

# Print the result
print(f"Total Revenue: {total_revenue}")
print(f"Total Expenses: {total_expenses}")
print(f"Profit: {profit}")


#####################################################################################################################
#Write back to master file

# File path of the existing ODS file
file_path = 'Overview.ods'

# Profit calculation results
profit_data = {
    "Metric": ["Total Revenue", "Total Expenses", "Profit"],
    "Value": [total_revenue, total_expenses, profit]
}

# Create a DataFrame for profit
profit_df = pd.DataFrame(profit_data)

# Load existing ODS file
sheets_data = pd.read_excel(file_path, engine='odf', sheet_name=None)

# Add a new sheet for profit
sheets_data['Profit Summary'] = profit_df

# Save back to the ODS file
with pd.ExcelWriter(file_path, engine='odf') as writer:
    for sheet_name, df in sheets_data.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Profit summary added to the file: {file_path}")