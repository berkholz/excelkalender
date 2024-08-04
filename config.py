

############################ IMPORTS #####################
from openpyxl import Workbook
import datetime
import calendar
from calendar import monthrange
from openpyxl.styles import PatternFill,Font, Color, Border, Side
############################ VARIABLES #####################
excel_file_name = "kalender.xlsx"

users = ['Marcel','Antje','Johannes','Irma','Arthur']

cal = calendar.Calendar()
#actual_year = datetime.date.today().year
actual_year = 2023


def add_header(worksheet):
    worksheet['A1'] = actual_year

def append_month_header(worksheet):
    # Rows can also be appended
    worksheet.append(["Name", ])

############################ MAIN #####################
wb = Workbook()

# grab the active worksheet
ws = wb.active
ws.title = str(actual_year)

add_header(ws)

row_count = 3
# iterating over the month
for month in range(1,13):
    # get count of days for the actual month in for loop
    _, last_day_of_month = monthrange(actual_year, month)
    # save all days of month in row
    actual_cell = ws.cell(row=row_count,column=1,value=calendar.month_name[month])
    actual_cell.font = Font(bold=True)
    actual_cell.border = Border(bottom=Side(border_style="thin", color="000000"))
    #actual_cell.fill = PatternFill(start_color="8a2be2", end_color="8a2be2", fill_type="solid")

    for column in range(1,last_day_of_month + 1):
        actual_cell = ws.cell(row=row_count, column=column+1, value=column)
        actual_cell.font = Font(bold=True)
        actual_cell.border = Border(bottom=Side(border_style="thin", color="000000"))
        # col = str(chr(64 + actual_cell))
        # ws.column_dimensions[col].width = 20

    actual_cell = ws.cell(row=row_count,column=33,value=calendar.month_name[month])
    actual_cell.font = Font(bold=True)
    actual_cell.border = Border(bottom=Side(border_style="thin", color="000000"))
    row_count = row_count + 1

    for user in users:
        ws.cell(row=row_count,column=1,value=user)
        ws.cell(row=row_count,column=33,value=user)
        row_count = row_count + 1

    # let two rows empty for spacing
    row_count = row_count + 2



# Save the file
wb.save(excel_file_name)