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
def process_sheet(df, region, sheet_type):
    print(f"\nProcessing {sheet_type} sheet:")
    print(f"Original shape: {df.shape}")
    print(f"Columns: {df.columns.tolist()}")
    print(df.head())

    rows_to_keep = []

    for index, row in df.iterrows():
        site_name = row.get('Power Station') or row.get('Project', '')
        
        if pd.isna(site_name) or site_name == 'Total':
            continue
        
        new_row = row.copy()
        new_row['Region'] = region
        
        if sheet_type == 'new_developments':
            new_row['Unit Status'] = row.get('Unit Status', 'Unknown')
        elif sheet_type == 'scheduled':
            if site_name == 'Committed':
                continue  # Skip the 'Committed' row itself
            new_row['Unit Status'] = 'Committed' if 'Committed' in df.iloc[:index+1]['Power Station'].values else 'In Service'
        else:  # non-scheduled
            new_row['Unit Status'] = 'In Service'
        
        rows_to_keep.append(new_row)

    processed_df = pd.DataFrame(rows_to_keep)
    print(f"\nProcessed {sheet_type} sheet:")
    print(f"Processed shape: {processed_df.shape}")
    print(f"Columns: {processed_df.columns.tolist()}")
    print(processed_df.head())
    
    return processed_df

# Extract region from filename
region = extract_region(file_path)

# Process Existing S & SS Generation sheet
df_scheduled = pd.read_excel(file_path, sheet_name='Existing S & SS Generation', header=1)
df_scheduled = df_scheduled.loc[:, ~df_scheduled.columns.str.contains('^Unnamed')]
new_df_scheduled = process_sheet(df_scheduled, region, 'scheduled')

# Process Non-Scheduled Generation sheet
df_non_scheduled = pd.read_excel(file_path, sheet_name='Non-Scheduled Generation', header=1)
df_non_scheduled = df_non_scheduled.loc[:, ~df_non_scheduled.columns.str.contains('^Unnamed')]
new_df_non_scheduled = process_sheet(df_non_scheduled, region, 'non_scheduled')

# Process New Developments sheet
df_new_developments = pd.read_excel(file_path, sheet_name='New Developments', header=1)
df_new_developments = df_new_developments.loc[:, ~df_new_developments.columns.str.contains('^Unnamed')]
new_df_new_developments = process_sheet(df_new_developments, region, 'new_developments')

# Function to standardize column names
def standardize_columns(df, sheet_type):
    column_mapping = {
        'Power Station': 'Site Name',
        'Project': 'Site Name',
        'Plant Type': 'Technology Type',
        'Technology Type': 'Technology Type',
        'Generation Type': 'Technology Type',
        'Unit Numbers and Nameplate Capacity (MW)': 'Nameplate Capacity',
        'Nameplate Capacity (MW)': 'Nameplate Capacity'
    }
    df = df.rename(columns=column_mapping)
    
    # Ensure all required columns are present
    required_columns = ['Region', 'Site Name', 'Technology Type', 'Nameplate Capacity', 'Unit Status']
    for col in required_columns:
        if col not in df.columns:
            df[col] = ''
    
    return df[required_columns]

# Standardize column names for all dataframes
new_df_scheduled = standardize_columns(new_df_scheduled, 'scheduled')
new_df_non_scheduled = standardize_columns(new_df_non_scheduled, 'non_scheduled')
new_df_new_developments = standardize_columns(new_df_new_developments, 'new_developments')

# Combine the three DataFrames
combined_df = pd.concat([new_df_scheduled, new_df_non_scheduled, new_df_new_developments], ignore_index=True)

# Display the first few rows of the extracted data
print("\nFinal combined data:")
print(f"Shape: {combined_df.shape}")
print(f"Columns: {combined_df.columns.tolist()}")
print(combined_df.head())

# Save the extracted data to a new Excel file
output_file = 'extracted2.xlsx'
combined_df.to_excel(output_file, index=False)
print(f"\nData has been extracted and saved to '{output_file}'")

# Print total number of rows extracted
print(f"Total rows extracted: {len(combined_df)}")
print(f"Rows from Scheduled Generation: {len(new_df_scheduled)}")
print(f"Rows from Non-Scheduled Generation: {len(new_df_non_scheduled)}")
print(f"Rows from New Developments: {len(new_df_new_developments)}")
