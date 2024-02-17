import pandas as pd
from openpyxl import load_workbook

# Load the Excel file
excel_file = "input1.xlsx"  # Replace "your_file.xlsx" with the path to your Excel file
wb = load_workbook(excel_file)

# Select the first sheet (you can choose a specific sheet if needed)
sheet = wb.active

# Get the max row and max column count from the sheet
max_row = sheet.max_row
max_col = sheet.max_column

# Initialize an empty list to hold the data
data = []

# Iterate over rows in the sheet
for row in sheet.iter_rows(min_row=2, max_row=max_row, min_col=1, max_col=max_col, values_only=True):
    # Create a list to hold processed row data
    processed_row = []
    for cell_value in row:
        if cell_value is None:
            # If cell is None, propagate value from the top-left cell of the merged region
            cell_value = processed_row[-1] if processed_row else None
        processed_row.append(cell_value)
    data.append(processed_row)

# Convert the list of rows into a pandas DataFrame
df = pd.DataFrame(data)

# Now you have your DataFrame with merged cells handled properly
print(df)