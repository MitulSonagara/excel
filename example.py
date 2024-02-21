import pandas as pd

# Sample DataFrame
data = {
    "block_no": [1, 1, 2, 2, 3],
    "subject_code": ["Math", "Math", "Science", "Science", "English"],
    "enrollment_number": [101, 102, 201, 202, 301],
}

df = pd.DataFrame(data)

# Grouping by 'block_no' and 'subject_code' and aggregating 'enrollment_number' using a lambda function
grouped_df = (
    df.groupby(["block_no", "subject_code"])
    .agg({"enrollment_number": lambda x: "\n".join(map(str, x))})
    .reset_index()
)

# Adding a new column 'total' to represent the total number of rows merged within each group
grouped_df["total"] = df.groupby(["block_no", "subject_code"]).size().values

print(grouped_df)
