#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Dec 26 18:11:18 2024

@author: venkat
"""

import pandas as pd

# File path for the new ODS file
new_file_path = 'financial_summary.ods'

# Sample data for additional information
# Adjust these with real calculations or data as required
sales_data = {
    "Item": ["Fruits", "Vegetables", "Spices", "Pulses"],
    "Quantity Sold": [50, 100, 30, 20],
    "Revenue": [500, 1000, 300, 200]
}
sales_df = pd.DataFrame(sales_data)

personal_expenses_data = {
    "Category": ["Groceries", "Rent", "Utilities", "Entertainment"],
    "Amount": [300, 1000, 200, 150]
}
personal_expenses_df = pd.DataFrame(personal_expenses_data)

savings_data = {
    "Month": ["January", "February", "March", "April"],
    "Savings": [500, 600, 700, 800]
}
savings_df = pd.DataFrame(savings_data)

# Profit calculation data (example)
profit_data = {
    "Metric": ["Total Revenue", "Total Expenses", "Profit"],
    "Value": [total_revenue, total_expenses, profit]
}
profit_df = pd.DataFrame(profit_data)

# Combine all data into a dictionary for saving
sheets_to_save = {
    "Expenses": expenses_df,
    "Stock": stock_df,
    "Sales": sales_df,
    "Profit Summary": profit_df,
    "Savings": savings_df,
    "Personal Expenses": personal_expenses_df
}

# Save all sheets into a new ODS file
with pd.ExcelWriter(new_file_path, engine='odf') as writer:
    for sheet_name, df in sheets_to_save.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"New financial summary file created: {new_file_path}")
