import pandas as pd
from dateutil import parser
from datetime import datetime

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

# Read the Excel file
df = pd.read_excel('extracted2.xlsx')

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

print("Date normalization and sorting completed. Output saved to 'extracted_new.xlsx'.")
