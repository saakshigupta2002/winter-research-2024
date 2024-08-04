import pandas as pd

# File path
file_path = 'NEM Generation Information May 2024.xlsx'

# Sheet name
sheet_name = 'ExistingGeneration&NewDevs'

# Read the Excel file
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Function to find the first non-empty row
def find_first_data_row(df):
    for index, row in df.iterrows():
        if not row.isna().all():
            return index
    return 0

# Find the first row with actual data
first_data_row = find_first_data_row(df)

# Extract columns based on their position
region = df.iloc[first_data_row:, 0]
site_name = df.iloc[first_data_row:, 2]
technology_type = df.iloc[first_data_row:, 4]
nameplate_capacity = df.iloc[first_data_row:, 12]
unit_status = df.iloc[first_data_row:, 14]

# Create a new DataFrame with the extracted columns
new_df = pd.DataFrame({
    'Region': region,
    'Site Name': site_name,
    'Technology Type': technology_type,
    'Nameplate Capacity (MW)': nameplate_capacity,
    'Unit Status': unit_status
})

# Reset the index to start from 0
new_df = new_df.reset_index(drop=True)

# Display the first few rows of the extracted data
print("\nExtracted data:")
print(new_df.head())

# Save the extracted data to a new Excel file
output_file = 'extracted.xlsx'
new_df.to_excel(output_file, index=False)
print(f"\nData has been extracted and saved to '{output_file}'")
