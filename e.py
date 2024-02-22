import pandas as pd

# Given DataFrame df1
df1 = pd.DataFrame(
    {
        0: range(1, 34),
        1: [3150711] * 6
        + [3150910] * 5
        + [3151107] * 5
        + [3151606] * 3
        + [3151607] * 3
        + [3151608] * 3
        + [3151609] * 3
        + [3151610] * 3
        + [3152407] * 2,
        2: [30] * 33,
        3: [None] * 33,
        4: [None] * 33,
        5: [None] * 33,
        6: [None] * 33,
        7: [None] * 33,
        8: [None] * 33,
    }
)

# Given DataFrame df2
df2 = pd.DataFrame(
    {
        0: [11, 19, 2, 25, 28, 31, 35, 35, 35, 35, 4],
        1: [
            3150910,
            3151606,
            3161012,
            3150610,
            3151909,
            3151705,
            3150501,
            3150711,
            3151107,
            3152407,
            3161613,
        ],
        2: [1] * 11,
        3: [None] * 11,
        4: [None] * 11,
        5: [None] * 11,
        6: [None] * 11,
        7: [None] * 11,
        8: [
            "200170109105",
            "210170116048",
            "190170111146",
            "210170106024",
            "210170119026",
            "210170117032",
            "210170105046",
            "220173107016\n210170107525\n210170107101",
            "200170111125\n210170111029\n210170111054\n210170111035\n210170111056\n210170111088\n210170111092",
            "210170124029",
            "200170116015",
        ],
    }
)

current_subject_code = None
new_rows = []
for index, row in df1.iterrows():
    if row[1] != current_subject_code:
        if current_subject_code is not None:
            for index, row1 in df2.iterrows():
                subject = row1[1]
                if current_subject_code == subject:
                    new_rows.append(row1)
                    df2 = df2.drop(index)
                    df2 = df2.reset_index(drop=True)

        current_subject_code = row[1]
    new_rows.append(row)


# Create a new DataFrame with inserted blank rows
df_with_blank_rows = pd.DataFrame(new_rows)
# Reset the index of the DataFrame
df_with_blank_rows = df_with_blank_rows.reset_index(drop=True)


df_with_blank_rows = pd.concat([df_with_blank_rows,df2],ignore_index=True)

print(df_with_blank_rows)
