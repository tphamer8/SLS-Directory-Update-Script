import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import pandas as pd
import re


# Adding New Hires
def add_new_hires(spreadsheet, path):
    new_hires = spreadsheet.worksheet('New Hires')
    print("Adding New Hires to Sheet:")

    # Read existing names from the first column
    existing_names = set(name[0] for name in new_hires.get_all_values()[1:])  # skip header
    df = pd.read_excel(path)

    for _, row in df.iterrows():
        if isinstance(row['Category'], str) and 'New Hires' in row['Category']:
            name = reformat_new_hire_names(row['Name'])
            title = row['Business Title']
            start_date = ''
            if pd.notna(row['Projected Start Date']):
                try:
                    date_val = pd.to_datetime(row['Projected Start Date'], errors='coerce')
                    if pd.notna(date_val):
                        start_date = date_val.strftime('%-m/%-d/%Y')  # or use '%#m/%#d/%Y' on Windows
                except Exception as e:
                    print(f"⚠️ Error parsing date for {row['Name']}: {e}")

            if name in existing_names:
                print(f"⏭️  Skipping duplicate: {name}")
                continue
            
            status = ''
            supervisor = reformat_names(row['Supervisor'])
            predecessor = row['Predecessor']

            values = [name, title, start_date, status, supervisor, predecessor]
            safe_values = [str(v) if pd.notna(v) else '' for v in values]
            new_hires.append_row(safe_values)
            if (safe_values[0] != ''):
                print(f"✔️ Added: {safe_values[0]}, {safe_values[1]}, {safe_values[2]}")

            
# Reformat New Hire names
def reformat_new_hire_names(name):
    if not isinstance(name, str):
        return ''
    return " ".join(re.sub(r"\(.*?\)", "", name).strip().split(", ")[::-1])

# Reformat Termination names
def reformat_names(name):
    if not isinstance(name, str):
        return ''
    return " ".join(name.split(", ")[::-1])


# Adding terminations that aren't duplicates
def add_terminations(spreadsheet, path):
    terminations = spreadsheet.worksheet('Terminations')
    print("Adding Terminations to Sheet:")

    # Read existing names from the first column
    existing_names = set(name[0] for name in terminations.get_all_values()[1:])  # skip header

    df = pd.read_excel(path)

    for _, row in df.iterrows():
        if row['Category'] == 'Departure':
            name = reformat_names(row['Name'])

            if name in existing_names:
                print(f"⏭️  Skipping duplicate: {name}")
                continue

            department = row['Department']
            end_date = row['Estimated End Date'].strftime('%-m/%-d/%Y') if not pd.isna(row['Estimated End Date']) else ''

            values = [name, department, end_date]
            safe_values = [str(v) if pd.notna(v) and v == v else '' for v in values]  # v == v excludes NaN
            terminations.append_row(safe_values)
            print(f"✔️ Added: {safe_values[0]}, {safe_values[1]}, {safe_values[2]}")


# Testing by reading column
def read_excel(path):
    print('Reading File')
    df = pd.read_excel(path)
    for index, row in df.iterrows():
        print(row['Category'])

# Authenticate Google Sheets
def authenticate_google_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(creds)
    return client

# Open the spreadsheet
def open_spreadsheet(client, spreadsheet_name):
    spreadsheet = client.open(spreadsheet_name)
    return spreadsheet

def main():
    client = authenticate_google_sheet()
    spreadsheet_name = 'Directory Updates (Excel Loader)'  # Replace with your spreadsheet name
    spreadsheet = open_spreadsheet(client, spreadsheet_name)
    
    path = input("Enter the path of the xlsx report file: ").strip()

    if not os.path.isfile(path):
        print(f"❌ File not found: {path}")
        exit(1)

    add_terminations(spreadsheet, path)
    add_new_hires(spreadsheet, path)

if __name__ == "__main__":
    main()