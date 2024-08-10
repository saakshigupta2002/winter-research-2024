import pandas as pd
from dateutil import parser
from datetime import datetime
import re

def normalize_date(date_str):
    try:
        # Handle the specific case of "2 22 February 2022"
        if '22 February 2022' in date_str:
            return datetime(2022, 2, 22)
        
        # Handle month and year format
        if len(date_str.split()) == 2:
            try:
                return datetime.strptime(date_str, "%B %Y").replace(day=1)
            except ValueError:
                pass  # If it's not a month name, continue to other parsing methods
        
        # Handle other date formats
        date = parser.parse(date_str, dayfirst=False, yearfirst=False)
        
        # If day is not specified, set it to 1
        if date.day == 1 and len(date_str.split()) == 2:
            return date.replace(day=1)
        
        return date
    except:
        print(f"Warning: Unable to parse date '{date_str}'")
        return None

def normalize_capacity(value):
    if pd.isna(value):
        return None
    
    if isinstance(value, (int, float)):
        return round(value, 2)
    
    if isinstance(value, str):
        value = value.strip().lower()
        
        if value in ['tba', 'tbc']:
            return value.upper()
        
        if value in ['zero capacity', 'none specified']:
            return 0
        
        if value == '':
            return None
        
        # Check for range
        range_match = re.match(r'(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)', value)
        if range_match:
            start, end = map(float, range_match.groups())
            return round((start + end) / 2, 2)
        
        # Try to convert to float
        try:
            return round(float(re.sub(r'[^\d.]', '', value)), 2)
        except ValueError:
            print(f"Warning: Unable to normalize capacity '{value}'")
            return None

def infer_technology_type(row):
    if pd.isna(row['Technology Type']) or row['Technology Type'] == '':
        site_name = row['Site Name'].lower()
        if 'wind' in site_name:
            return 'Wind'
        elif 'solar' in site_name:
            return 'Solar'
        elif 'storage' in site_name:
            return 'Storage'
    return row['Technology Type']

# Read the Excel file
df = pd.read_excel('extracted2.xlsx')

# Normalize Nameplate Capacity
df['Nameplate Capacity'] = df['Nameplate Capacity'].apply(normalize_capacity)

# Remove rows where Nameplate Capacity is None (invalid or empty), but keep 'TBA', 'TBC', and 0 (zero capacity)
df = df[df['Nameplate Capacity'].notna() | (df['Nameplate Capacity'].isin(['TBA', 'TBC', 0]))]

# Infer Technology Type from Site Name if missing
df['Technology Type'] = df.apply(infer_technology_type, axis=1)

# List of date columns (excluding 'Region', 'Site Name', 'Technology Type', and 'Nameplate Capacity')
non_date_columns = ['Region', 'Site Name', 'Technology Type', 'Nameplate Capacity']
date_columns = [col for col in df.columns if col not in non_date_columns]

# Dictionary to store new column names and their corresponding dates
new_column_names = {}
column_dates = {}

# Normalize dates in column names
for col in date_columns:
    date = normalize_date(col)
    if date:
        new_name = date.strftime('%d-%m-%Y')
        new_column_names[col] = new_name
        column_dates[new_name] = date
    else:
        new_column_names[col] = col
        column_dates[col] = datetime.max
    print(f"Original: {col}, Normalized: {new_column_names[col]}")

# Rename the columns
df = df.rename(columns=new_column_names)

# Sort date columns
sorted_date_columns = sorted(new_column_names.values(), key=lambda x: column_dates[x])

# Reorder columns
final_column_order = non_date_columns + sorted_date_columns
df = df[final_column_order]

# Write the result to a new Excel file
df.to_excel('extracted_new.xlsx', index=False)

print("Date and Nameplate Capacity normalization completed. Technology Type inferred where missing. Invalid entries removed, zero capacity entries kept. Output saved to 'extracted_new.xlsx'.")
print(f"Number of rows in original file: {len(pd.read_excel('extracted2.xlsx'))}")
print(f"Number of rows in new file: {len(df)}")
print(f"Number of Technology Types inferred: {sum(df['Technology Type'].isin(['Wind', 'Solar', 'Storage'])) - sum(pd.read_excel('extracted2.xlsx')['Technology Type'].isin(['Wind', 'Solar', 'Storage']))}")
