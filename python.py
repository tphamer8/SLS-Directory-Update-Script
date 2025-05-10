import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import pandas as pd


def add_terminations(path):
    df = pd.read_excel(path)
    for index, row in df.iterrows():
        print(row['column_name'])  # replace with your column name

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

# Main
if __name__ == '__main__':
    client = authenticate_google_sheet()
    spreadsheet_name = 'Directory Updates (CSV Loader)'  # Replace with your spreadsheet name
    spreadsheet = open_spreadsheet(client, spreadsheet_name)

    path = input("Enter the path of the xlsx report file: ").strip()

    if not os.path.isfile(path):
        print(f"❌ File not found: {path}")
        exit(1)
    add_terminations(path)