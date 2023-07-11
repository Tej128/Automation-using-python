import openpyxl
from openpyxl import Workbook
from openpyxl.styles import *
from openpyxl.worksheet.table import Table, TableStyleInfo
import os

path = "~/Desktop/Python-udemy/employees.txt"
expanded_path = os.path.expanduser(path)
text_file = open(expanded_path)
records = []

text_file.seek(0)

for record in text_file.readlines():
    records.append(record.rstrip("\n").split(";"))

workbook = Workbook()
fpath = "~/Desktop/Python-udemy/staff.xlsx"
file_path = os.path.expanduser(fpath)
workbook.save(file_path)

sheet = workbook["Sheet"]
sheet.title = "Employees"

for row in records:
    sheet.append(row)

table = Table(displayName="Table", ref="A1:G11")

style = TableStyleInfo(
    name="TableStyleMedium9", showRowStripes=True, showColumnStripes=True
)

table.tableStyleInfo = style

sheet.add_table(table)

font = Font(color=Color(rgb="FF0000"), bold=True, italic=True)

for cell_no in range(2, 12):
    if int(sheet["G%s" % cell_no].value) > 5500:
        sheet["G%s" % cell_no].font = font

workbook.save(file_path)

# Closing the text file
text_file.close()

# Closing the workbook
workbook.close()
