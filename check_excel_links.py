#!/usr/bin/env python
import sys
import openpyxl
import requests
from os.path import splitext

# setup and variables
# use an offset to skip header row
offset = 2

# column numbers for URL, status code, and content type.
url_col = 2
status_col = 5
content_type_col = 6

# check to see if a filename was passed in; if not, print a message & exit
if len(sys.argv) > 1:
    file_name = sys.argv[1]
else:
    print("This script requires the name of the Excel XLSX file to be processed.")
    exit(1)

# create the output filename based on the passed in filename
outfile = f"{splitext(file_name)[0]}_linkchecked.xlsx"

# open the file and get the active sheet
wb = openpyxl.load_workbook(filename=file_name)
sheet = wb.active

# Progress reporting
print(f"Processing {sheet.max_row} rows in file {file_name}.")
print(f"Will save records to XSLX file {outfile}.")

# Write in headers for additional columns
sheet.cell(row=1, column=status_col).value = 'STATUS CODE'
sheet.cell(row=1, column=content_type_col).value = 'CONTENT TYPE'

# calculate padding for processed rows
padding = len(str(sheet.max_row))

# Loop through the rows
for index, row in enumerate(sheet.iter_rows(min_row=offset)):
    # Every hundred rows, print total rows processed and percentage done.
    if index > 1 and index % 100 == 0:
        print(f"{index:{padding}} rows processed.", end=' ')
        print("({(index*100)/sheet.max_row:.0f}%)")
    try:
        # if URL, get headers: status code, mimetype
        # TODO: pass URL column as an argument rather than hard-code
        cv = row[url_col].value
        req = requests.head(cv, allow_redirects=True)
        sheet.cell(row=index + offset, column=status_col).value = req.status_code
        # take just the mime type, drop encoding
        content_type = req.headers['content-type'].split(';')[0]
        sheet.cell(row=index + offset, column=content_type_col).value = content_type
    # record if there's no URL, then continue
    except requests.exceptions.MissingSchema:
        sheet.cell(row=index + offset, column=status_col).value = "No valid URL."
        continue

try:
    # try to save the output file
    wb.save(outfile)
except IOError as err:
    print(f"Could not save records to {outfile}: {str(err)}")
    exit(1)

# Done. print a message and exit.
print(f"Finished. {sheet.max_row} rows processed, saved to file {outfile}.")
