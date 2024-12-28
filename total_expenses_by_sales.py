import pandas as pd

def calculate_expenses_by_shop(file_path, sheet_name=None):
    """
    Load data from an ODS file (skipping the first row) and calculate 
    total expenses grouped by shop.

    Args:
        file_path (str): Path to the ODS file.
        sheet_name (str, optional): Name of the sheet (default is the first sheet).

    Returns:
        pd.DataFrame: A DataFrame with total expenses grouped by shop.
    """
    try:
        # Read the ODS file and skip the first row
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='odf', skiprows=1)

        # Ensure required columns exist
        required_columns = {'Shop', 'Amount'}
        if not required_columns.issubset(df.columns):
            raise ValueError(f"Missing required columns. Ensure the file has columns: {required_columns}")

        # Group by shop and calculate the sum of Amount
        grouped_data = df.groupby(['Shop'], as_index=False).agg({'Amount': 'sum'})

        return grouped_data

    except Exception as e:
        print(f"Error reading or processing the ODS file: {e}")
        return None

def main():
    # Path to your ODS file
    file_path = 'DailySheet.ods'  # Replace with your ODS file path
    sheet_name = 'Expenses'  # Specify the sheet name if needed

    # Calculate and display expenses grouped by shop
    result = calculate_expenses_by_shop(file_path, sheet_name=sheet_name)
    if result is not None:
        print("Total Expenses Grouped by shop:")
        print(result)

        # Optionally, save the results to a ods file
        result.to_ods('total_expenses_by_shop.ods', index=False)
        print("\nResults saved to 'total_expenses_by_shop.ods'.")

if __name__ == "__main__":
    main()
