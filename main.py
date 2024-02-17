import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

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
for row in sheet.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col, values_only=True):
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

# Handle vertical merged cells
for merged_cell in sheet.merged_cells.ranges:
    # Check if the merged cell range is vertical
    if merged_cell.min_row != merged_cell.max_row:
        # Get the top-left cell value
        top_left_value = sheet.cell(row=merged_cell.min_row, column=merged_cell.min_col).value
        # Assign the top-left cell value to all cells within the merged range vertically
        for row_index in range(merged_cell.min_row + 1, merged_cell.max_row + 1):
            df.iloc[row_index - 1, merged_cell.min_col - 1] = top_left_value

# Now you have your DataFrame with both horizontal and vertical merged cells handled properly

# Drop the first row from the DataFrame
df = df.drop([0,1,2,4])

# Reset the index of the DataFrame after dropping the row
df = df.reset_index(drop=True)
# print(df)

date_obj = df.iloc[0, 1]  # Assuming the date is in the second column

# Format the datetime object as "07-Dec-2023"
formatted_date = date_obj.strftime("%d-%b-%Y")

# Now the formatted date is stored in the variable formatted_date
# print(formatted_date)  

df = df.drop(0)

# Reset the index of the DataFrame after dropping the row
df = df.reset_index(drop=True)
df = df.drop(df.columns[[2,3,4]], axis=1)
df.columns = range(len(df.columns))

df[[0, 1]] = df[[1, 0]]

print(df)