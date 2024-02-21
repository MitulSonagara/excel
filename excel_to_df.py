import pandas as pd
from openpyxl import load_workbook


def abc(path):
    wb = load_workbook(path)
    sheet = wb.active
    max_row = sheet.max_row
    max_col = sheet.max_column
    data = []

    for row in sheet.iter_rows(
        min_row=1, max_row=max_row, min_col=1, max_col=max_col, values_only=True
    ):
        processed_row = []
        for cell_value in row:
            if cell_value is None:
                cell_value = processed_row[-1] if processed_row else None
            processed_row.append(cell_value)
        data.append(processed_row)

    df = pd.DataFrame(data)

    for merged_cell in sheet.merged_cells.ranges:
        if merged_cell.min_row != merged_cell.max_row:
            top_left_value = sheet.cell(
                row=merged_cell.min_row, column=merged_cell.min_col
            ).value
            for row_index in range(merged_cell.min_row + 1, merged_cell.max_row + 1):
                df.iloc[row_index - 1, merged_cell.min_col - 1] = top_left_value

    return df


df = abc("emergency.xlsx")
df = df.drop(0)
df = df.drop(df.columns[[0, 1, 2, 6, 7, 8]], axis=1)
df = df.reset_index(drop=True)
df.columns = range(len(df.columns))
df[[0, 2]] = df[[2, 0]]
cols = ["block no", "subject code", "number"]
df.columns = cols

grouped_df = (
    df.groupby(["block no", "subject code"])
    .agg({"number": lambda x: "\n".join(map(str, x))})
    .reset_index()
)

# Adding the 'total' column to grouped_df
grouped_df["total"] = df.groupby(["block no", "subject code"]).size().values

# Correcting the column assignments
grouped_df[["number", "total"]] = grouped_df[["total", "number"]]
cols = ["block no", "subject code", "total","number"]
grouped_df.columns = cols


print(grouped_df)
