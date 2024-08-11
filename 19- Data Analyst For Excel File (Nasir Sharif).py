import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def load_excel_file():
    # Open a file dialog to pick an Excel file
    Tk().withdraw()  # Hide the root Tk window
    file_path = askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    
    if not file_path:
        print("No file selected.")
        return None
    return file_path

def search_in_dataframe(df, search_term):
    if df is not None:
        print("\nSearch Results:")
        for column in df.columns:
            if df[column].dtype == 'object':  # Search only in string/object columns
                matches = df[df[column].str.contains(search_term, case=False, na=False)]
                if not matches.empty:
                    print(f"\nIn column '{column}':")
                    print(matches)

# Main Program
file_path = load_excel_file()
if file_path:
    try:
        # Load the Excel file into a DataFrame
        df = pd.read_excel(file_path, sheet_name=None)  # Load all sheets

        # Search term
        search_term = input("Enter the keyword or phrase to search for: ")

        # Search across all sheets
        for sheet_name, data in df.items():
            print(f"\nSearching in sheet: {sheet_name}")
            search_in_dataframe(data, search_term)
    
    except Exception as e:
        print(f"Error loading or processing the file: {e}")
