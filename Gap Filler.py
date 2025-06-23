import pandas as pd

# Load original Excel file
df = pd.read_excel("C:/Users/Welcome/Desktop/AV1.xlsx", engine='openpyxl')

# Get minimum and maximum values from the column 'Rows'
min_val = df['Numbers'].min()
max_val = df['Numbers'].max()

# Create a full range DataFrame
full_range = pd.DataFrame({'Numbers': range(min_val, max_val + 1)})

# Merge with the original data to detect missing numbers
merged = full_range.merge(df, on='Numbers', how='left', indicator=True)

# For rows that were missing in original data, clear the 'Rows' value
merged.loc[merged['_merge'] == 'left_only', 'Numbers'] = ''

# Drop the indicator column
merged.drop(columns=['_merge'], inplace=True)

# Save to new Excel file
merged.to_excel("C:/Users/Welcome/Desktop/AV2.xlsx", index=False)

print("✅ File created! Missing numbers now have blank cells in the 'Rows' column.")
