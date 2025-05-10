import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import pandas as pd




# Reformat Termination names
def reformat_names(name):
    first_last = first_last = " ".join(name.split(", ")[::-1])
    return first_last

# Adding terminations that aren't duplicates
def add_terminations(spreadsheet, path):
    terminations = spreadsheet.worksheet('Terminations')
    print("Adding Terminations to Sheet:")

    # Read existing names from the first column (assumes "Name" is in column A)
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

            terminations.append_row([name, department, end_date])
            print(f"✔️ Added: {name}, {department}, {end_date}")

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
    spreadsheet_name = 'Directory Updates (CSV Loader)'  # Replace with your spreadsheet name
    spreadsheet = open_spreadsheet(client, spreadsheet_name)
    
    path = input("Enter the path of the xlsx report file: ").strip()

    if not os.path.isfile(path):
        print(f"❌ File not found: {path}")
        exit(1)

    # read_excel(path) 
    add_terminations(spreadsheet, path)

if __name__ == "__main__":
    main()