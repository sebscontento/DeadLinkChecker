import gspread
from google.oauth2.service_account import Credentials
import csv

# Define the scope and authorize the API client
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_file("/Users/contento/dev/linksProject/deadlinks-439220-6fbc62e6f02c.json", scopes=scope)
client = gspread.authorize(creds)

# Open the Google Sheet by its name
spreadsheet_name = "My Dead Links Report"
sheet = client.open(spreadsheet_name)  # Open the existing spreadsheet
worksheet = sheet.get_worksheet(0)  # Open the first sheet in the spreadsheet

# Import CSV data
csv_file = "dead_links_report.csv"

with open(csv_file, mode='r') as file:
    reader = csv.reader(file)
    rows = list(reader)

# Add headers if needed
if not rows[0][0].lower() == 'dead links':
    rows.insert(0, ['Link Number', 'Dead Links', 'Corrected?'])

# Update the sheet with CSV data
worksheet.update('A1', rows)

# Number the entries in the Link Number column
link_numbers = [[i] for i in range(1, len(rows))]  # Numbering from 1
worksheet.update('A2:A', link_numbers)  # Update link numbers in column A

# Prepare empty cells for checkboxes in the Corrected? column
checkboxes = [['FALSE'] for _ in range(1, len(rows))]  # Initialize checkboxes as FALSE
worksheet.update('C2:C', checkboxes)  # Add empty cells for checkboxes

# Set conditional formatting rules
formatting_rules = [
    {
        'addConditionalFormatRule': {
            'rule': {
                'booleanRule': {
                    'condition': {
                        'type': 'BOOLEAN',
                        'booleanCondition': {
                            'type': 'TRUE'
                        }
                    },
                    'format': {
                        'backgroundColor': {
                            'red': 0.0,
                            'green': 1.0,
                            'blue': 0.0
                        }
                    }
                },
                'range': {
                    'sheetId': worksheet.id,
                    'startRowIndex': 1,  # Starting from the second row
                    'endRowIndex': len(rows),
                    'startColumnIndex': 2,  # Column C
                    'endColumnIndex': 3,
                }
            }
        }
    },
    {
        'addConditionalFormatRule': {
            'rule': {
                'booleanRule': {
                    'condition': {
                        'type': 'BOOLEAN',
                        'booleanCondition': {
                            'type': 'FALSE'
                        }
                    },
                    'format': {
                        'backgroundColor': {
                            'red': 1.0,
                            'green': 0.0,
                            'blue': 0.0
                        }
                    }
                },
                'range': {
                    'sheetId': worksheet.id,
                    'startRowIndex': 1,  # Starting from the second row
                    'endRowIndex': len(rows),
                    'startColumnIndex': 2,  # Column C
                    'endColumnIndex': 3,
                }
            }
        }
    }
]

# Apply the conditional formatting rules
worksheet.batch_update(formatting_rules)

print(f"Data successfully imported to Google Sheets: {spreadsheet_name}")
print(f"Spreadsheet URL: {sheet.url}")

