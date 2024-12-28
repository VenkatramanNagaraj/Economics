import pandas as pd

def calculate_expenses_by_category(file_path, sheet_name=None):
    """
    Load data from an ODS file (skipping the first row) and calculate 
    total expenses grouped by category.

    Args:
        file_path (str): Path to the ODS file.
        sheet_name (str, optional): Name of the sheet (default is the first sheet).

    Returns:
        pd.DataFrame: A DataFrame with total expenses grouped by category.
    """
    try:
        # Read the ODS file and skip the first row
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='odf', skiprows=1)

        # Ensure required columns exist
        required_columns = {'Category', 'Amount'}
        if not required_columns.issubset(df.columns):
            raise ValueError(f"Missing required columns. Ensure the file has columns: {required_columns}")

        # Group by Category and calculate the sum of Amount
        grouped_data = df.groupby(['Category'], as_index=False).agg({'Amount': 'sum'})

        return grouped_data

    except Exception as e:
        print(f"Error reading or processing the ODS file: {e}")
        return None

def main():
    # Path to your ODS file
    file_path = 'DailySheet.ods'  # Replace with your ODS file path
    sheet_name = 'Expenses'  # Specify the sheet name if needed

    # Calculate and display expenses grouped by category
    result = calculate_expenses_by_category(file_path, sheet_name=sheet_name)
    if result is not None:
        print("Total Expenses Grouped by Category:")
        print(result)

        # Optionally, save the results to a CSV file
        result.to_csv('total_expenses_by_category.csv', index=False)
        print("\nResults saved to 'total_expenses_by_category.csv'.")

if __name__ == "__main__":
    main()
