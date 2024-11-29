import pandas as pd
import re

# Load the Excel file
file_path = 'try.xlsx'
df = pd.read_excel(file_path)

# Display column names to confirm the 'Body' column exists
print("Column names in the DataFrame:")
print(df.columns)

# Check if 'Body' column exists
if 'Body' not in df.columns:
    raise KeyError("The 'Body' column does not exist in the DataFrame.")

# Function to extract the required fields from the Body text
def extract_info(Body):
    Body = Body.replace('\n', '<br/>')  # Replace newlines with spaces
    Body = Body.replace('<span>', '<br/>')  # Replace newlines with spaces
    user_name = re.search(r'User Name:\s*(.*?)(<br/>|$)', Body)
    email = re.search(r'Email:\s*(.*?)(<br/>|$)', Body)
    organization = re.search(r'Organization:\s*(.*?)(<br/>|$)', Body)
    industry = re.search(r'Industry:\s*(.*?)(<br/>|$)', Body)
    lead_looking_for = re.search(r'Lead is looking for:\s*(.*?)(<br/>|$)', Body)

    return {
        'User Name': user_name.group(1).strip() if user_name else None,
        'Email': email.group(1).strip() if email else None,
        'Organization': organization.group(1).strip() if organization else None,
        'Industry': industry.group(1).strip() if industry else None,
        'Lead is looking for': lead_looking_for.group(1).strip() if lead_looking_for else None
    }

# Apply the function to the 'Body' column
# Apply replacements to each row in the 'Body' column
df['Body'] = df['Body'].str.replace(r'\n', '<br/>')
# Apply replacements to each row in the 'Body' column
df['Body'] = df['Body'].str.replace(r'<span>', '<br/>')
extracted_data = df['Body'].apply(lambda x: extract_info(x) if isinstance(x, str) else None)

# Create a new DataFrame from the extracted data
new_df = pd.DataFrame(extracted_data.tolist())

# Save the new DataFrame to a new Excel file
new_file_path = 'extracted_data.xlsx'
new_df.to_excel(new_file_path, index=False)

print(f"Extracted data has been saved to {new_file_path}")
