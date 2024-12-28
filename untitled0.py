#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Dec 25 20:31:00 2024

@author: venkat
"""

import pandas as pd
from odf.opendocument import load
from odf.table import Table, TableRow, TableCell

# Load the ODF file
doc = load("DailySheet.ods")

# Initialize an empty list to store the extracted data
data = []

# Find all tables in the ODF file
for table in doc.getElementsByType(Table):
    for row in table.getElementsByType(TableRow):
        cells = row.getElementsByType(TableCell)
        
        # Assuming columns: Expense, Category, Shop, Amount Paid (adjust based on your file's structure)
        if len(cells) == 4:  # Check if the row has 4 columns
            expense = cells[0].firstChild.nodeValue if cells[0].firstChild else ''
            category = cells[1].firstChild.nodeValue if cells[1].firstChild else ''
            shop = cells[2].firstChild.nodeValue if cells[2].firstChild else ''
            amount_paid = cells[3].firstChild.nodeValue if cells[3].firstChild else ''
            
            data.append([expense, category, shop, amount_paid])

# Create the dataframe
df = pd.DataFrame(data, columns=["Expense", "Category", "Shop", "Amount Paid"])

# Display the dataframe
print(df)
