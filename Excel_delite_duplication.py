import os
import pandas as pd
from tkinter import Tk, filedialog

def remove_duplicates_from_excel(file_path):
    # Check if the file exists
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return
    
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(file_path)

    # Print the original DataFrame
    print("Original DataFrame:")
    print(df)

    # Print the column names in the DataFrame
    print("\nColumn names in the DataFrame:", df.columns)

    # Make sure the specified column ('kolomoetss@list.ru') is in the DataFrame
    column_name = 'kolomoetss@list.ru'  # Replace with the actual column name
    if column_name not in df.columns:
        print(f"Error: '{column_name}' column not found in the DataFrame.")
        return

    # Identify and print the rows that will be removed
    duplicate_rows = df[df.duplicated(subset=column_name, keep='first')]
    print("\nRows that will be removed:")
    print(duplicate_rows)

    # Remove duplicates based on the specified column
    df_no_duplicates = df.drop_duplicates(subset=column_name, keep='first')

    # Save the modified DataFrame back to the Excel file
    output_file_path = os.path.splitext(file_path)[0] + "_no_duplicates.xlsx"
    df_no_duplicates.to_excel(output_file_path, index=False)

    print(f"\nDuplicates removed. Result saved to: {output_file_path}")

def select_excel_file():
    root = Tk()
    root.withdraw()  # Hide the main window

    # Display the file dialog to select the Excel file
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx;*.xls")]
    )

    return file_path

# Example usage
if __name__ == "__main__":
    # Use the file dialog to select the Excel file
    excel_file_path = select_excel_file()

    if excel_file_path:
        remove_duplicates_from_excel(excel_file_path)
    else:
        print("File selection canceled.")
