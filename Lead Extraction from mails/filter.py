import pandas as pd

# Load the Excel file
file_path = 'New_Signup_on_Axonator.xlsx'
df = pd.read_excel(file_path)

# Display the original DataFrame
print("Original DataFrame:")
print(df)

# # Drop rows where any cell contains 'Aarohi Kulkarni'
df = df[~df.apply(lambda row: row.astype(str).str.contains('Aarohi Kulkarni').any(), axis=1)]
df = df[~df.apply(lambda row: row.astype(str).str.lower().str.contains('aarohi kulkarni').any(), axis=1)]
df = df[~df.apply(lambda row: row.astype(str).str.lower().str.contains('@axonator.com').any(), axis=1)]
df = df[~df.apply(lambda row: row.astype(str).str.lower().str.contains('white snow').any(), axis=1)]
df = df[~df.apply(lambda row: row.astype(str).str.lower().str.contains('whitesnow').any(), axis=1)]
df = df[~df.apply(lambda row: row.astype(str).str.contains('Whitesnow').any(), axis=1)]
df = df[~df.apply(lambda row: row.astype(str).str.contains('Jayesh Kitukale').any(), axis=1)]
df['Body'] = df['Body'].str.replace(r'\n', '<br/>')
# Apply replacements to each row in the 'Body' column
df['Body'] = df['Body'].str.replace(r'<span>', '<br/>')

# Keep only the rows where any cell contains 'Lead is looking'
# df = df[df.apply(lambda row: row.astype(str).str.contains('Lead is looking').any(), axis=1)]



# Save the result back to an Excel file
df.to_excel('try.xlsx', index=False)
print("\nFile Saved:")
