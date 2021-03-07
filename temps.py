"""
Simple API Data Gathering - API Basics - (3 of 8)

de YT Chris Vaden.

"""

"""
Import JSON data to an Excel spreadsheet using Python

from YT Jie Jenn.
"""

from pprint import pprint
import os
import json
import win32com.client as win32 # pip install pywin32

"""
Step 1.1 Read the JSON file
"""
json_data = json.loads(open('data.json').read())
pprint(json_data)

"""
Step 1.2 Examing the data and flatten the records into a 2D layout
"""
rows = []

for record in json_data:
    id = record['_id']
    is_active = record['isActive']
    email = record['email']
    balance = record['balance']
    first_name = record['name']['first']
    last_name = record['name']['last']
    tags = ','.join(record['tags'])
    friends = '; '.join(['Id: {0}, name: {1}'.format(friend['id'], friend['name']) for friend in record['friends']]).strip()
    rows.append([id, is_active, email, balance, first_name, last_name, tags, friends])

"""
Step 2. Inserting Records to an Excel Spreadsheet
"""
ExcelApp = win32.Dispatch('Excel.Application')
ExcelApp.Visible = True

wb = ExcelApp.Workbooks.Add()
ws = wb.Worksheets(1)

header_labels = ('id', 'is active', 'email', 'balance', 'first name', 'last name', 'tags', 'friends')

# insert header labels
for indx, val in enumerate(header_labels):
    ws.Cells(1, indx + 1).Value = val

# insert Records
row_tracker = 2
column_size = len(header_labels)

for row in rows:
    ws.Range(
        ws.Cells(row_tracker, 1),
        ws.Cells(row_tracker, column_size)
    ).value = row
    row_tracker += 1

wb.SaveAs(os.path.join(os.getcwd(), 'Json output.xlsx'), 51)
wb.Close()
ExcelApp.Quit()
ExcelApp = None