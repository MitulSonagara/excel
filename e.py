import pandas as pd

# First DataFrame
data1 = {
    "block_no": [1, 1, 2, 2, 3],
    0: [1234567, 1234567, 2345678, 2345678, 3456789],  # Using integers as column index
    "enrollment_number": [101, 102, 201, 202, 301],
}

# Second DataFrame
data2 = {
    "block_no": [1, 2, 3],
    0: [1234567, 2345678, 4567890],  # Using integers as column index
    "enrollment_number": [101, 102, 103],
}

df1 = pd.DataFrame(data1)
df2 = pd.DataFrame(data2)

# Grouping df1 by 'subject_code' (using column index)
groups_df1 = df1.groupby(df1.columns[1])  # Change index as per your data

# Flag to check if any row from df2 is inserted into df1
inserted_flag = False

# Iterate over rows of df2
for index, row in df2.iterrows():
    subject_code = row[df2.columns[1]]  # Change index as per your data

    # Find the corresponding group in df1
    if subject_code in groups_df1.groups:
        group_index = groups_df1.groups[subject_code]

        # Find the last index of this group
        last_index = max(group_index)

        # Insert after the last occurrence of this group
        df1 = pd.concat(
            [df1.iloc[: last_index + 1], row.to_frame().T, df1.iloc[last_index + 1 :]]
        ).reset_index(drop=True)

        # Set the flag to True since a row from df2 is inserted into df1
        inserted_flag = True
    else:
        # If subject_code is not found in df1, append the row to the end
        df1 = df1._append(row, ignore_index=True)
        inserted_flag = True

# If no row from df2 is inserted into df1, it means the 'subject_code' from df2 is completely different
# In this case, add all rows from df2 to the end of df1
if not inserted_flag:
    df1 = pd.concat([df1, df2], ignore_index=True)

print(df1)
