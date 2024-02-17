import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from datetime import datetime

# Load the Excel file
excel_file = "input1.xlsx"  # Replace "input1.xlsx" with the path to your Excel file
wb = load_workbook(excel_file)

# Select the first sheet (you can choose a specific sheet if needed)
sheet = wb.active

# Get the max row and max column count from the sheet
max_row = sheet.max_row
max_col = sheet.max_column

# Initialize an empty list to hold the data
data = []

# Iterate over rows in the sheet
for row in sheet.iter_rows(
    min_row=1, max_row=max_row, min_col=1, max_col=max_col, values_only=True
):
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
        top_left_value = sheet.cell(
            row=merged_cell.min_row, column=merged_cell.min_col
        ).value
        # Assign the top-left cell value to all cells within the merged range vertically
        for row_index in range(merged_cell.min_row + 1, merged_cell.max_row + 1):
            df.iloc[row_index - 1, merged_cell.min_col - 1] = top_left_value

# Now you have your DataFrame with both horizontal and vertical merged cells handled properly

# Drop the unnecessary rows and columns
df = df.drop([0, 1, 2, 3, 4])
df = df.drop(df.columns[[2, 3, 4]], axis=1)

# Reset the index of the DataFrame after dropping rows and columns
df = df.reset_index(drop=True)
df.columns = range(len(df.columns))

# Swap columns 0 and 1
df[[0, 1]] = df[[1, 0]]

# Insert a blank row after each group of subject code
current_subject_code = None
new_rows = []
for index, row in df.iterrows():
    if row[1] != current_subject_code:
        if current_subject_code is not None:
            new_rows.append(pd.Series([None] * len(df.columns), index=df.columns))
            new_rows[-1][0] = "Total"
            new_rows[-1][1] = None
        current_subject_code = row[1]
    new_rows.append(row)


# Create a new DataFrame with inserted blank rows
df_with_blank_rows = pd.DataFrame(new_rows)

# Reset the index of the DataFrame
df_with_blank_rows = df_with_blank_rows.reset_index(drop=True)

# Manually set the column names
df_with_blank_rows.columns = df.columns

df_with_blank_rows = df_with_blank_rows.drop(1)

df_with_blank_rows = df_with_blank_rows.reset_index(drop=True)

print(df_with_blank_rows)

output_excel_file = "output1.xlsx"

# Write the DataFrame with inserted blank rows to an Excel file
df_with_blank_rows.to_excel(output_excel_file, index=False)

excel_file = "output1.xlsx"  # Replace "input1.xlsx" with the path to your Excel file
wb1 = load_workbook(excel_file)

# Select the first sheet (you can choose a specific sheet if needed)
ws = wb1.active

for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=2):
    # Check if the first cell is "Total" and the second cell is empty
    if row[0].value == "Total" and not row[1].value:
        print(f"Merging cells in row {row[0].row}")
        # Merge cells for "Total"
        ws.merge_cells(
            start_row=row[0].row, start_column=1, end_row=row[0].row, end_column=2
        )

# Create an Alignment object for centering
alignment = Alignment(horizontal="center", vertical="center")

# Iterate through all cells and set the alignment
for row in ws.iter_rows():
    for cell in row:
        cell.alignment = alignment

group_sum = 0
start_row = 1

for row_num in range(3, ws.max_row):
    if ws[f"C{row_num}"].value is not None:
        group_sum += int(
            ws[f"C{row_num}"].value
        )  # Add the value in the third cell to the group sum
    elif ws[f"C{row_num}"].value is None:
        print(group_sum)
        ws[f"C{row_num}"] = group_sum
        group_sum = 0


wb1.save(excel_file)

print(f"Excel file '{output_excel_file}' created successfully.")
