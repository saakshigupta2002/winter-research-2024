import pandas as pd
import re

# File path
file_path = 'Generation_Information_NSW_20131104.xlsx'

def extract_region(filename):
    states = ['NSW', 'QLD', 'SA', 'TAS', 'VIC']
    return next((state for state in states if state in filename), 'Unknown')

def translate_region(region):
    region_mapping = {
        'NSW': 'NSW1',
        'QLD': 'QLD1',
        'SA': 'SA1',
        'TAS': 'TAS1',
        'VIC': 'VIC1'
    }
    return region_mapping.get(region, region)

def find_first_data_row(df):
    for index, row in df.iterrows():
        if not row.isna().all():
            return index
    return 0

def extract_single_sheet(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    first_data_row = find_first_data_row(df)
    
    new_df = pd.DataFrame({
        'Region': df.iloc[first_data_row:, 0],
        'Site Name': df.iloc[first_data_row:, 2],
        'Technology Type': df.iloc[first_data_row:, 4],
        'Nameplate Capacity': df.iloc[first_data_row:, 12],
        'Unit Status': df.iloc[first_data_row:, 14]
    })
    
    new_df = new_df.reset_index(drop=True)
    new_df['Region'] = new_df['Region'].apply(translate_region)
    
    return new_df

def find_sheet_name(excel_file, possible_names):
    with pd.ExcelFile(excel_file) as xls:
        return next((name for name in possible_names if name in xls.sheet_names), None)

def is_note_or_statement(text):
    if not isinstance(text, str):
        return False
    return text.startswith('Note:') or text.startswith('*') or ':' in text

def extract_max_capacity(capacity_str):
    if not capacity_str or not isinstance(capacity_str, (str, int, float)):
        return ''
    if isinstance(capacity_str, (int, float)):
        return capacity_str
    numbers = re.findall(r'\d+(?:\.\d+)?', str(capacity_str))
    return max(map(float, numbers)) if numbers else capacity_str

def translate_unit_status(status, sheet_type):
    if sheet_type == 'new_developments':
        if status == 'Pub An':
            return 'Publicly Announced'
        elif status == 'Com':
            return 'Committed'
    return status

def process_sheet(df, region, sheet_type):
    print(f"\nProcessing {sheet_type} sheet:")
    print(f"Original shape: {df.shape}")
    print(f"Columns: {df.columns.tolist()}")
    print(df.head())

    rows_to_keep = []
    committed_encountered = False
    service_status_column = next((col for col in df.columns if 'Service Status' in col), None)

    for index, row in df.iterrows():
        site_name = row.get('Power Station') or row.get('Project', '') or row.get('  Project', '')
        
        if pd.isna(site_name) or site_name == 'Total' or is_note_or_statement(site_name):
            continue
        
        if site_name == 'Committed':
            committed_encountered = True
            continue
        
        nameplate_capacity = (row.get('Unit Number and Nameplate Capacity (MW)') or 
                              row.get('Unit Numbers and Nameplate Capacity (MW)') or 
                              row.get('Nameplate Capacity (MW)', ''))
        
        unit_status = (row.get(service_status_column) if service_status_column else
                       (row.get('Unit Status', 'Unknown') if sheet_type == 'new_developments' 
                        else ('Committed' if committed_encountered else 'In Service')))
        
        unit_status = translate_unit_status(unit_status, sheet_type)
        
        new_row = {
            'Region': translate_region(region),
            'Site Name': site_name,
            'Technology Type': row.get('Plant Type') or row.get('Technology Type') or row.get('Generation Type', ''),
            'Nameplate Capacity': extract_max_capacity(nameplate_capacity),
            'Unit Status': unit_status
        }
        
        rows_to_keep.append(new_row)

    processed_df = pd.DataFrame(rows_to_keep)
    print(f"\nProcessed {sheet_type} sheet:")
    print(f"Processed shape: {processed_df.shape}")
    print(f"Columns: {processed_df.columns.tolist()}")
    print(processed_df.head())
    
    return processed_df

# Check if ExistingGeneration&NewDevs sheet exists
single_sheet_name = 'ExistingGeneration&NewDevs'
if single_sheet_name in pd.ExcelFile(file_path).sheet_names:
    print(f"Found {single_sheet_name} sheet. Processing single sheet.")
    combined_df = extract_single_sheet(file_path, single_sheet_name)
else:
    print(f"{single_sheet_name} sheet not found. Processing multiple sheets.")
    # Extract region from filename
    region = extract_region(file_path)

    # Process sheets
    df_scheduled = pd.read_excel(file_path, sheet_name='Existing S & SS Generation', header=1)
    df_scheduled = df_scheduled.loc[:, ~df_scheduled.columns.str.contains('^Unnamed')]
    new_df_scheduled = process_sheet(df_scheduled, region, 'scheduled')

    non_scheduled_sheet_name = find_sheet_name(file_path, ['Non-Scheduled Generation', 'Existing NS Generation'])
    if non_scheduled_sheet_name:
        df_non_scheduled = pd.read_excel(file_path, sheet_name=non_scheduled_sheet_name, header=1)
        df_non_scheduled = df_non_scheduled.loc[:, ~df_non_scheduled.columns.str.contains('^Unnamed')]
        new_df_non_scheduled = process_sheet(df_non_scheduled, region, 'non_scheduled')
    else:
        print("Warning: Non-Scheduled Generation sheet not found")
        new_df_non_scheduled = pd.DataFrame()

    df_new_developments = pd.read_excel(file_path, sheet_name='New Developments', header=1)
    df_new_developments = df_new_developments.loc[:, ~df_new_developments.columns.str.contains('^Unnamed')]
    new_df_new_developments = process_sheet(df_new_developments, region, 'new_developments')

    wind_sheet_name = 'Existing Wind Generation'
    if wind_sheet_name in pd.ExcelFile(file_path).sheet_names:
        df_wind = pd.read_excel(file_path, sheet_name=wind_sheet_name, header=1)
        df_wind = df_wind.loc[:, ~df_wind.columns.str.contains('^Unnamed')]
        new_df_wind = process_sheet(df_wind, region, 'wind')
    else:
        print("Warning: Existing Wind Generation sheet not found")
        new_df_wind = pd.DataFrame()

    # Combine all DataFrames
    all_dfs = [new_df_scheduled, new_df_non_scheduled, new_df_new_developments]
    if not new_df_wind.empty:
        all_dfs.append(new_df_wind)
    combined_df = pd.concat(all_dfs, ignore_index=True)

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
