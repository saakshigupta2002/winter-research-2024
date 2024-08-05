import pandas as pd
import re
import os
from pathlib import Path
from datetime import datetime
from collections import defaultdict

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

def extract_single_sheet(file_path, sheet_name, status_column_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    first_data_row = find_first_data_row(df)
    
    new_df = pd.DataFrame({
        'Region': df.iloc[first_data_row:, 0],
        'Site Name': df.iloc[first_data_row:, 2],
        'Technology Type': df.iloc[first_data_row:, 4],
        'Nameplate Capacity': df.iloc[first_data_row:, 12],
        status_column_name: df.iloc[first_data_row:, 14]
    })
    
    new_df = new_df.reset_index(drop=True)
    new_df['Region'] = new_df['Region'].apply(translate_region)
    
    # Remove rows with more than two missing entries
    new_df = new_df.dropna(thresh=3)
    
    return new_df

def find_sheet_name(excel_file, possible_names):
    with pd.ExcelFile(excel_file) as xls:
        return next((name for name in possible_names if name in xls.sheet_names), None)

def is_note_or_statement(text):
    if not isinstance(text, str):
        return False
    return text.startswith('Note:') or text.startswith('*') or text.startswith('a.') or ':' in text

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

def process_sheet(df, region, sheet_type, status_column_name):
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
                              row.get('Nameplate Capacity (MW)', '') or
                              row.get('Nameplate Capacity (MW)a', '') or
                              row.get('Nameplate Capacity (MW)^a', ''))
        
        technology_type = row.get('Plant Type') or row.get('Technology Type') or row.get('Generation Type', '')
        
        unit_status = (row.get(service_status_column) if service_status_column else
                       (row.get('Unit Status', 'Unknown') if sheet_type == 'new_developments' 
                        else ('Committed' if committed_encountered else 'In Service')))
        
        unit_status = translate_unit_status(unit_status, sheet_type)
        
        # Count non-empty entries
        non_empty_count = sum(1 for v in [technology_type, nameplate_capacity, unit_status] if v)
        
        # Include the row if at least two of the essential columns have a value
        if non_empty_count >= 2:
            new_row = {
                'Region': translate_region(region),
                'Site Name': site_name,
                'Technology Type': technology_type,
                'Nameplate Capacity': extract_max_capacity(nameplate_capacity),
                status_column_name: unit_status
            }
            rows_to_keep.append(new_row)

    processed_df = pd.DataFrame(rows_to_keep)
    print(f"\nProcessed {sheet_type} sheet:")
    print(f"Processed shape: {processed_df.shape}")
    print(f"Columns: {processed_df.columns.tolist()}")
    print(processed_df.head())
    
    return processed_df

def process_file(file_path, status_column_name):
    print(f"\nProcessing file: {file_path}")
    region = extract_region(file_path)

    # Check if ExistingGeneration&NewDevs sheet exists
    single_sheet_name = 'ExistingGeneration&NewDevs'
    if single_sheet_name in pd.ExcelFile(file_path).sheet_names:
        print(f"Found {single_sheet_name} sheet. Processing single sheet.")
        return extract_single_sheet(file_path, single_sheet_name, status_column_name)
    
    print(f"{single_sheet_name} sheet not found. Processing multiple sheets.")
    
    # Process sheets
    df_scheduled = pd.read_excel(file_path, sheet_name='Existing S & SS Generation', header=1)
    df_scheduled = df_scheduled.loc[:, ~df_scheduled.columns.str.contains('^Unnamed')]
    new_df_scheduled = process_sheet(df_scheduled, region, 'scheduled', status_column_name)

    non_scheduled_sheet_name = find_sheet_name(file_path, ['Non-Scheduled Generation', 'Existing NS Generation'])
    if non_scheduled_sheet_name:
        df_non_scheduled = pd.read_excel(file_path, sheet_name=non_scheduled_sheet_name, header=1)
        df_non_scheduled = df_non_scheduled.loc[:, ~df_non_scheduled.columns.str.contains('^Unnamed')]
        new_df_non_scheduled = process_sheet(df_non_scheduled, region, 'non_scheduled', status_column_name)
    else:
        print("Warning: Non-Scheduled Generation sheet not found")
        new_df_non_scheduled = pd.DataFrame()

    df_new_developments = pd.read_excel(file_path, sheet_name='New Developments', header=1)
    df_new_developments = df_new_developments.loc[:, ~df_new_developments.columns.str.contains('^Unnamed')]
    new_df_new_developments = process_sheet(df_new_developments, region, 'new_developments', status_column_name)

    wind_sheet_name = 'Existing Wind Generation'
    if wind_sheet_name in pd.ExcelFile(file_path).sheet_names:
        df_wind = pd.read_excel(file_path, sheet_name=wind_sheet_name, header=1)
        df_wind = df_wind.loc[:, ~df_wind.columns.str.contains('^Unnamed')]
        new_df_wind = process_sheet(df_wind, region, 'wind', status_column_name)
    else:
        print("Warning: Existing Wind Generation sheet not found")
        new_df_wind = pd.DataFrame()

    # Combine all DataFrames
    all_dfs = [new_df_scheduled, new_df_non_scheduled, new_df_new_developments]
    if not new_df_wind.empty:
        all_dfs.append(new_df_wind)
    return pd.concat(all_dfs, ignore_index=True)

def merge_data(existing_df, new_df, status_column_name):
    for _, row in new_df.iterrows():
        site_name = row['Site Name']
        existing_row = existing_df[existing_df['Site Name'] == site_name]
        
        if not existing_row.empty:
            # Update existing row
            idx = existing_row.index[0]
            existing_df.at[idx, 'Technology Type'] = row['Technology Type']
            existing_df.at[idx, 'Nameplate Capacity'] = row['Nameplate Capacity']
            existing_df.at[idx, status_column_name] = row[status_column_name]
        else:
            # Add new row
            new_row = row.to_dict()
            existing_df = existing_df.append(new_row, ignore_index=True)
    
    return existing_df

def extract_date_from_filename(filename):
    # Extract day, month, and year from filename
    match = re.search(r'(\d{1,2})?\s*(\w+)\s*(\d{4})', filename)
    if match:
        day, month, year = match.groups()
        try:
            if day:
                date = datetime.strptime(f"{day} {month} {year}", "%d %b %Y")
            else:
                date = datetime.strptime(f"{month} {year}", "%b %Y")
            return date.strftime("%d %B %Y") if day else date.strftime("%B %Y"), (month, year)
        except ValueError:
            try:
                if day:
                    date = datetime.strptime(f"{day} {month} {year}", "%d %B %Y")
                else:
                    date = datetime.strptime(f"{month} {year}", "%B %Y")
                return date.strftime("%d %B %Y") if day else date.strftime("%B %Y"), (month, year)
            except ValueError:
                print(f"Warning: Could not parse date from filename: {filename}")
                return None, None
    return None, None

def main(input_folder):
    combined_df = pd.DataFrame(columns=['Region', 'Site Name', 'Technology Type', 'Nameplate Capacity'])
    month_year_counter = defaultdict(int)
    
    # Process files in the main folder
    for item in os.listdir(input_folder):
        item_path = os.path.join(input_folder, item)
        
        if os.path.isfile(item_path) and item.endswith('.xlsx'):
            status_column_name, month_year = extract_date_from_filename(item)
            if not status_column_name:
                status_column_name = os.path.splitext(item)[0]  # Use filename without extension if date not found
            else:
                if month_year:
                    month_year_counter[month_year] += 1
                    if month_year_counter[month_year] > 1:
                        status_column_name = f"{month_year_counter[month_year]} {status_column_name}"
            
            try:
                processed_data = process_file(item_path, status_column_name)
                
                if status_column_name not in combined_df.columns:
                    combined_df[status_column_name] = ''
                
                combined_df = merge_data(combined_df, processed_data, status_column_name)
            except Exception as e:
                print(f"Error processing file {item}: {str(e)}")
        
        elif os.path.isdir(item_path):
            # Process files in subfolders
            for root, _, files in os.walk(item_path):
                for file in files:
                    if file.endswith('.xlsx'):
                        file_path = os.path.join(root, file)
                        status_column_name = os.path.basename(os.path.dirname(file_path))
                        
                        try:
                            processed_data = process_file(file_path, status_column_name)
                            
                            if status_column_name not in combined_df.columns:
                                combined_df[status_column_name] = ''
                            
                            combined_df = merge_data(combined_df, processed_data, status_column_name)
                        except Exception as e:
                            print(f"Error processing file {file}: {str(e)}")
    
    if combined_df.empty:
        print("No data processed. Check your input folder and file types.")
        return

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

if __name__ == "__main__":
    input_folder = input("Enter the path to the folder containing Excel files and subfolders: ")
    main(input_folder)
