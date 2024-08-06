

# packages to install on ubuntu:
# - python3-openpyxl
# - python3-bs4
# - pyyaml

############################ IMPORTS #####################
from openpyxl import Workbook
import datetime
import calendar
from calendar import monthrange
from openpyxl.styles import PatternFill,Font, Color, Border, Side, Alignment
import random
import sys
import yaml

############################ VARIABLES #####################
excel_file_name = "kalender.xlsx"

# row to start from, row_count must be greater equal 1
row_count = 1

# dict with users to setup
users = ['Marcel','Antje','Johannes','Irma','Arthur']
user_colors = dict()

colors = ['00FF0000','0000FF00','000000FF','00FFFF00','00FF00FF','0000FFFF','00008000','00008080','00800080','009999FF','00FFFFCC','00FF8080','00CCCCFF','0099CC00','00FFCC00','00FF9900','00FF6600','00993300']

configuration_file = None

cal = calendar.Calendar()
actual_year = datetime.date.today().year



############################ FUNCTIONS #####################
def load_configuration():
    global configuration_file
    with open('config.yml', 'r') as file:
        configuration_file = yaml.safe_load(file)

def check_actual_year():
    global actual_year
    if c.year_for_calendar != '':
        actual_year = c.year_for_calendar

def print_list(list):
    print("[", end='')
    for element in list:
        print("{},".format(element), end='')
    print("]")

def associate_colors_to_users():
    global user_colors
    for user in users:
        key = random.sample(colors, 1)
        user_colors.update({ user : key[0] })
        colors.remove(key[0])

def add_header(worksheet):
    global row_count
    worksheet.merge_cells('A1:AG1')
    worksheet.row_dimensions[1].height = 30
    cell = ws.cell(row=1,column=1,value=actual_year)
    cell.font = Font(name='Arial',size=16,bold=True)
    cell.alignment = Alignment(horizontal='center',vertical='center')
    cell.border = Border(bottom=Side(border_style="thin", color="000000"))
    #cell.fill = PatternFill(start_color="CFCFCF", end_color="0F0F0F", fill_type="solid")
    row_count = 3

def set_header_colomn_style (cell):
    cell.font = Font(name='Arial', bold=True)
    cell.alignment = Alignment(horizontal='center')
    cell.border = Border(bottom=Side(border_style="thin", color="000000"))
    cell.fill = PatternFill(start_color='CFCFCF', end_color='CFCFCF', fill_type="solid")

def set_user_cell_style(cell, usercolor):
    cell.font = Font(name='Arial', bold=False)
    cell.alignment = Alignment(horizontal='left')
    #cell.border = Border(bottom=Side(border_style="thin", color="000000"))
    cell.fill = PatternFill(start_color=usercolor, end_color=usercolor, fill_type="solid")

def set_column_width(worksheet):
    # set cell width of month name
    column_range = ['A','AG']
    for column in column_range:
        worksheet.column_dimensions[column].width = 25

    # set cell width of days of month
    column_range = ['B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF']
    for column in column_range:
        worksheet.column_dimensions[column].width = 5

def append_month(year, range_list):
    global row_count
    for month in range_list:
        _, last_day_of_month = monthrange(year, month)
        actual_cell = ws.cell(row=row_count,column=1,value=calendar.month_name[month])
        set_header_colomn_style(actual_cell)

        column_position = 0
        for column in range(1,32):
            if column_position < last_day_of_month:
                actual_cell = ws.cell(row=row_count, column=column+1, value=column)
                set_header_colomn_style(actual_cell)
            else:
                actual_cell = ws.cell(row=row_count, column=column+1, value='')
                set_header_colomn_style(actual_cell)
            column_position += 1
        actual_cell = ws.cell(row=row_count,column=33,value=calendar.month_name[12])
        set_header_colomn_style(actual_cell)
        row_count = row_count + 1
        for user in users:
            actual_cell = ws.cell(row=row_count,column=1,value=user)
            set_user_cell_style(actual_cell, user_colors[user])
            actual_cell = ws.cell(row=row_count,column=33,value=user)
            set_user_cell_style(actual_cell, user_colors[user])
            row_count = row_count + 1
        # let two rows empty for spacing
        row_count = row_count + 2

############################ MAIN #####################
load_configuration()

associate_colors_to_users()

wb = Workbook()

# grab the active worksheet
ws = wb.active
ws.title = str(actual_year)

add_header(ws)
set_column_width(ws)



# generating month
append_month(actual_year - 1,list([12]))
append_month(actual_year,range(1,13))
append_month(actual_year + 1,list([1]))

# Save the file
wb.save(excel_file_name)
