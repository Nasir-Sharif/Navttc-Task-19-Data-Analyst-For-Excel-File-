# Navttc-Task-19-Data-Analyst-For-Excel-File

# Excel Search Script

## Description

This Python script allows you to search for a specific keyword or phrase within all string columns across all sheets of an Excel file. The script provides a simple GUI to select the Excel file and displays the search results directly in the console.

## Code

```python
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
```

## Steps

1. **Import Required Libraries:**
   - Import the `pandas` library for working with Excel files.
   - Import the necessary modules from `tkinter` for creating a file dialog.

2. **Define the `load_excel_file` Function:**
   - This function opens a file dialog allowing the user to select an Excel file (`.xlsx` or `.xls`).
   - If no file is selected, the function returns `None`.

3. **Define the `search_in_dataframe` Function:**
   - This function searches for the specified keyword or phrase in all string columns of the DataFrame.
   - It prints the results, including the column name and matching rows.

4. **Main Program:**
   - The script first loads the Excel file using the `load_excel_file` function.
   - It then loads all sheets into a DataFrame.
   - The user is prompted to enter the search term, which is then searched across all sheets.

5. **Search Results:**
   - The script prints the sheet name, column name, and matching rows where the search term was found.

## How to Run

1. **Install Python and Required Libraries:**
   - Ensure Python is installed on your system. You can download it from [python.org](https://www.python.org/downloads/).
   - Install the required libraries using pip:
     ```bash
     pip install pandas tkinter
     ```

2. **Save the Script:**
   - Save the provided Python code into a file named `19- Data Analyst For Excel File (Nasir Sharif).py`.

3. **Execute the Script:**
   - Open a terminal or command prompt.
   - Navigate to the directory where `19- Data Analyst For Excel File (Nasir Sharif).py` is saved.
   - Run the script using the following command:
     ```bash
     python 19- Data Analyst For Excel File (Nasir Sharif).py
     ```

4. **Provide Input:**
   - When prompted, select the Excel file you want to search in.
   - Enter the keyword or phrase you want to search for.

## Example

- **File:** If you select an Excel file `example.xlsx`, the script will search across all sheets in that file.
- **Search Term:** If you input `"example"`, the script will search for this word in each string column and display the matching rows.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contact

For any questions or feedback, please contact Nasir Sharif at nasirsharifqasoori786@gmail.com
