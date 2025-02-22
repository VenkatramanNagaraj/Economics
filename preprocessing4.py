import os
import pandas as pd
from datetime import datetime, timedelta

# Paths
daily_sheets_folder = "Daily_Sheets"
overview_file_path = "Overview.ods"
os.makedirs(daily_sheets_folder, exist_ok=True)

# Get today's and next day's date
current_date = datetime.now().strftime('%Y-%m-%d')
next_date = (datetime.now() + timedelta(days=1)).strftime('%Y-%m-%d')

# Load the latest DailySheet file
daily_sheets = sorted([f for f in os.listdir(daily_sheets_folder) if f.startswith("DailySheet") and f.endswith(".ods")])

if daily_sheets:
    latest_sheet = os.path.join(daily_sheets_folder, daily_sheets[-1])  # Most recent file
    sheets_data = pd.read_excel(latest_sheet, engine='odf', sheet_name=None)

    # Extract Expenses and Sales sheets
    expenses_df = sheets_data.get('Expenses', pd.DataFrame())
    sales_df = sheets_data.get('Sales', pd.DataFrame())

    # Clean column names
    expenses_df.columns = expenses_df.columns.str.strip()
    sales_df.columns = sales_df.columns.str.strip()

    # Remove "Date" column if it exists
    if 'Date' in expenses_df.columns:
        expenses_df.drop(columns=['Date'], inplace=True)
    if 'Date' in sales_df.columns:
        sales_df.drop(columns=['Date'], inplace=True)

    # Calculate total expenses, revenue, and outstanding balance
    total_expenses = expenses_df['Amount'].sum() if 'Amount' in expenses_df.columns else 0
    total_revenue = sales_df['Amount'].sum() if 'Amount' in sales_df.columns else 0
    total_outstanding_balance = expenses_df['Outstanding Balance'].sum() if 'Outstanding Balance' in expenses_df.columns else 0
    profit = total_revenue - total_expenses

    # Print results
    print(f"Total Revenue: {total_revenue}")
    print(f"Total Expenses: {total_expenses}")
    print(f"Profit: {profit}")
    print(f"Outstanding Balance: {total_outstanding_balance}")

    # Prepare data for writing to Overview sheet
    summary_data = {
        "Metric": ["Total Revenue", "Total Expenses", "Profit", "Outstanding Balance"],
        current_date: [total_revenue, total_expenses, profit, total_outstanding_balance],
    }

    summary_df = pd.DataFrame(summary_data)

    # Load and update the Overview sheet
    overview_data = pd.read_excel(overview_file_path, engine='odf', sheet_name='Overview')

    for index, row in summary_df.iterrows():
        metric = row['Metric']
        value = row[current_date]
        if current_date in overview_data.columns:
            overview_data.loc[overview_data['Metric'] == metric, current_date] = value
        else:
            overview_data[current_date] = None
            overview_data.loc[overview_data['Metric'] == metric, current_date] = value

    # Save the updated Overview sheet
    with pd.ExcelWriter(overview_file_path, engine='odf') as writer:
        overview_data.to_excel(writer, sheet_name='Overview', index=False)

    print(f"Overview sheet updated for {current_date}.")

    # ----------------- CREATE NEXT DAY'S DAILY SHEET -----------------

    new_file_path = os.path.join(daily_sheets_folder, f"DailySheet_{next_date}.ods")

    # Copy structure of the previous Expenses sheet and retain outstanding balance
    new_expenses_df = expenses_df.copy()
    if 'Outstanding Balance' in new_expenses_df.columns:
        new_expenses_df['Outstanding Balance'] = expenses_df['Outstanding Balance']
    new_expenses_df['Amount'] = None  # Clear Amount values in Expenses sheet

    # Copy structure of Sales sheet and empty the "Amount" column
    new_sales_df = sales_df.copy()
    if 'Amount' in new_sales_df.columns:
        new_sales_df['Amount'] = None  # Clear Amount values in Sales sheet

    # Save the new DailySheet
    with pd.ExcelWriter(new_file_path, engine='odf') as writer:
        new_expenses_df.to_excel(writer, sheet_name='Expenses', index=False)
        new_sales_df.to_excel(writer, sheet_name='Sales', index=False)

    print(f"New daily sheet created: {new_file_path}")

else:
    print("No existing DailySheet found. Please create an initial file.")
