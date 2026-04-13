import pandas as pd

# Define a simple dataframe
data = {
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Age': [25, 30, 35],
    'Date': pd.to_datetime(['2023-01-01', '2023-02-01', '2023-03-01'])
}
df = pd.DataFrame(data)

# Create an Excel file
with pd.ExcelWriter('test_file.xlsx') as writer:
    df.to_excel(writer, sheet_name='Sheet1', index=False)

print("Created test_file.xlsx")
