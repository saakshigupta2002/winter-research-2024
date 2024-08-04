import pandas as pd
import re

# File path
file_path = 'Generation_Information_NSW_20131104.xlsx'

# Sheet name
sheet_name = 'Existing S & SS Generation'

# Read the Excel file, skipping the first row
df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)

# Remove any unnamed columns
df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

# Extract region from filename
def extract_region(filename):
    states = ['NSW', 'QLD', 'SA', 'TAS', 'VIC']
    for state in states:
        if state in filename:
            return state
    return 'Unknown'

region = extract_region(file_path)

# Initialize variables
committed_encountered = False
rows_to_keep = []

# Process rows
for index, row in df.iterrows():
    site_name = row.get('Power Station', '')
    
    if pd.isna(site_name) or site_name == 'Total':
        continue
    
    if site_name == 'Committed':
        committed_encountered = True
        continue
    
    new_row = row.copy()
    new_row['Region'] = region
    if committed_encountered:
        new_row['Unit Status'] = 'Committed'
    else:
        new_row['Unit Status'] = 'In Service'
    
    rows_to_keep.append(new_row)

# Create new DataFrame with processed rows
new_df = pd.DataFrame(rows_to_keep)

# Rename columns to match our desired output
column_mapping = {
    'Power Station': 'Site Name',
    'Plant Type': 'Technology Type',
    'Unit Numbers and Nameplate Capacity (MW)': 'Nameplate Capacity'
}
new_df = new_df.rename(columns=column_mapping)

# Select and order the required columns
required_columns = ['Region', 'Site Name', 'Technology Type', 'Nameplate Capacity', 'Unit Status']
new_df = new_df[required_columns]

# Display the first few rows of the extracted data
print("\nExtracted data:")
print(new_df.head())

# Save the extracted data to a new Excel file
output_file = 'extracted2.xlsx'
new_df.to_excel(output_file, index=False)
print(f"\nData has been extracted and saved to '{output_file}'")

# Print total number of rows extracted
print(f"Total rows extracted: {len(new_df)}")
