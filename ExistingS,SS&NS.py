import pandas as pd
import re

# File path
file_path = 'Generation_Information_NSW_20131104.xlsx'

# Function to extract region from filename
def extract_region(filename):
    states = ['NSW', 'QLD', 'SA', 'TAS', 'VIC']
    for state in states:
        if state in filename:
            return state
    return 'Unknown'

# Function to process a single sheet
def process_sheet(df, region, is_scheduled):
    committed_encountered = False
    rows_to_keep = []

    for index, row in df.iterrows():
        site_name = row.get('Power Station', '')
        
        if pd.isna(site_name) or site_name == 'Total':
            continue
        
        if site_name == 'Committed':
            committed_encountered = True
            continue
        
        new_row = row.copy()
        new_row['Region'] = region
        if is_scheduled:
            if committed_encountered:
                new_row['Unit Status'] = 'Committed'
            else:
                new_row['Unit Status'] = 'In Service'
        else:
            new_row['Unit Status'] = 'In Service'
        
        rows_to_keep.append(new_row)

    return pd.DataFrame(rows_to_keep)

# Extract region from filename
region = extract_region(file_path)

# Process Existing S & SS Generation sheet
df_scheduled = pd.read_excel(file_path, sheet_name='Existing S & SS Generation', header=1)
df_scheduled = df_scheduled.loc[:, ~df_scheduled.columns.str.contains('^Unnamed')]
new_df_scheduled = process_sheet(df_scheduled, region, True)

# Process Non-Scheduled Generation sheet
df_non_scheduled = pd.read_excel(file_path, sheet_name='Non-Scheduled Generation', header=1)
df_non_scheduled = df_non_scheduled.loc[:, ~df_non_scheduled.columns.str.contains('^Unnamed')]
new_df_non_scheduled = process_sheet(df_non_scheduled, region, False)

# Function to standardize column names
def standardize_columns(df):
    column_mapping = {
        'Power Station': 'Site Name',
        'Plant Type': 'Technology Type',
        'Technology Type': 'Technology Type',
        'Unit Numbers and Nameplate Capacity (MW)': 'Nameplate Capacity',
        'Nameplate Capacity (MW)': 'Nameplate Capacity'
    }
    df = df.rename(columns=column_mapping)
    return df

# Standardize column names for both dataframes
new_df_scheduled = standardize_columns(new_df_scheduled)
new_df_non_scheduled = standardize_columns(new_df_non_scheduled)

# Combine the two DataFrames
combined_df = pd.concat([new_df_scheduled, new_df_non_scheduled], ignore_index=True)

# Merge Technology Type columns if both exist
if 'Technology Type' in combined_df.columns and 'Technology Type_y' in combined_df.columns:
    combined_df['Technology Type'] = combined_df['Technology Type'].fillna(combined_df['Technology Type_y'])
    combined_df = combined_df.drop('Technology Type_y', axis=1)

# Merge Nameplate Capacity columns if both exist
if 'Nameplate Capacity' in combined_df.columns and 'Nameplate Capacity_y' in combined_df.columns:
    combined_df['Nameplate Capacity'] = combined_df['Nameplate Capacity'].fillna(combined_df['Nameplate Capacity_y'])
    combined_df = combined_df.drop('Nameplate Capacity_y', axis=1)

# Select and order the required columns
required_columns = ['Region', 'Site Name', 'Technology Type', 'Nameplate Capacity', 'Unit Status']
combined_df = combined_df[required_columns]

# Display the first few rows of the extracted data
print("\nExtracted data:")
print(combined_df.head())

# Save the extracted data to a new Excel file
output_file = 'extracted2.xlsx'
combined_df.to_excel(output_file, index=False)
print(f"\nData has been extracted and saved to '{output_file}'")

# Print total number of rows extracted
print(f"Total rows extracted: {len(combined_df)}")
print(f"Rows from Scheduled Generation: {len(new_df_scheduled)}")
print(f"Rows from Non-Scheduled Generation: {len(new_df_non_scheduled)}")
