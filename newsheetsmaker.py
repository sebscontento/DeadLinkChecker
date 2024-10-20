import gspread
from google.oauth2.service_account import Credentials
import csv

# Define the scope and authorize the API client
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_file("/Users/contento/dev/linksProject/deadlinks-439220-6fbc62e6f02c.json", scopes=scope)  # Update the path
client = gspread.authorize(creds)

# Open the Google Sheet by its name
spreadsheet_name = "My Dead Links Report"
sheet = client.create(spreadsheet_name)  # Create new spreadsheet
worksheet = sheet.get_worksheet(0)  # Open the first sheet in the spreadsheet

# Import CSV data
csv_file = "dead_links_report.csv"

with open(csv_file, mode='r') as file:
    reader = csv.reader(file)
    rows = list(reader)

# Update the sheet with CSV data
worksheet.update(rows)

print(f"Data successfully imported to Google Sheets: {spreadsheet_name}")

