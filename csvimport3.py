# Author: Sebastian Holbøll 
# Date: 2024-10-20
# Version: 3.0
# Description: This script imports a CSV file to a Google Sheet.

import gspread
from google.oauth2.service_account import Credentials
import csv
import os
import glob
import pandas as pd

# Define the scope and authorize the API client
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_file("/Users/contento/dev/linksProject/deadlinks-439220-6fbc62e6f02c.json", scopes=scope)  # Update the path
client = gspread.authorize(creds)

# Open the Google Sheet by its name
spreadsheet_name = "My Dead Links Report"

# Create new spreadsheet in the root folder for testing
sheet = client.create(spreadsheet_name)

# Share the spreadsheet with your Google account
email_to_share = "contentocameras@gmail.com"  # Replace with your email
sheet.share(email_to_share, perm_type='user', role='writer')  # Grant editor access

# Pattern to match the CSV files
pattern = "dead_links_report_*.csv"

# Find all CSV files that match the pattern
csv_files = glob.glob(pattern)

if csv_files:
    # Sort files based on the date and timestamp in the filename
    latest_file = max(csv_files, key=lambda x: os.path.splitext(x)[0].rsplit('_', 2)[-1])
    
    print(f"Importing data from: {latest_file}")

    # Import CSV data
    df = pd.read_csv(latest_file)
    
    # Convert DataFrame to a list of lists (rows)
    rows = df.values.tolist()
    
    # Open the first sheet in the spreadsheet
    worksheet = sheet.get_worksheet(0)  
    
    # Update the sheet with CSV data
    worksheet.update([df.columns.values.tolist()] + rows)  # Include header

    print(f"Data successfully imported to Google Sheets: {spreadsheet_name}")
    print(f"Spreadsheet URL: {sheet.url}")
else:
    print("No CSV files found.")
