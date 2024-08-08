
# packages to install on ubuntu:
# - python3-openpyxl
# - python3-bs4
# - pyyaml

############################ IMPORTS #####################
import sys

# import for creating excel files
from openpyxl import Workbook
import datetime
import calendar
from calendar import monthrange
from openpyxl.styles import PatternFill,Font, Color, Border, Side, Alignment

# import for pickung randomized elements from list colors[]
import random

# import for triggering REST API for holidays retrieval
import requests

# import for reading yaml configuration files
import yaml

# import for printing date in localized format
import locale

############################ VARIABLES #####################
excel_file_name = None

# row to start from, row_count must be greater equal 1
row_count = 1

colors = list()

cal_config = None
config_name = 'config.yml'

cal = calendar.Calendar()

actual_year = None

holidays = list()

wb = Workbook()

############################ FUNCTIONS #####################
def load_config():
    global cal_config
    with open(config_name, 'r') as file:
        cal_config = yaml.safe_load(file)

def check_config_locale():
    global cal_config
    if not "locale" in cal_config:
        cal_config.update({ "locale" : "de_DE" })
        print("No locale found in confiuguration file {}, using default locale {}".format(config_name,cal_config['locale']))

def check_config_colors():
    global cal_config
    global colors
    # check if colors are specified in configurations file, otherwise use default
    if "available_colors" in cal_config:
        colors = cal_config['available_colors'].copy()
    else:
        colors = ['00FF0000','0000FF00','000000FF','00FFFF00','00FF00FF','0000FFFF','00008000','00008080','00800080','009999FF','00FFFFCC','00FF8080','00CCCCFF','0099CC00','00FFCC00','00FF9900','00FF6600','00993300']
        print("Using default colors: {}".format(colors))

def check_config_excel_file_name():
    global cal_config
    global excel_file_name
    # check if output file name is given via configuration file, otherwise use >kalender.xlsx<
    if "excel_file_name" in cal_config:
        excel_file_name = cal_config['excel_file_name']
    else:
        excel_file_name = "kalender.xlsx"
        print("Using default file name for Excel output: {}".format(excel_file_name))

def check_config_year():
    global cal_config
    global actual_year
    # check if actual year is given via configuration file, otherwise take datetime.date.today().year as actual_year
    if "year_for_calendar" in cal_config:
        actual_year = cal_config['year_for_calendar']
    else:
        actual_year = datetime.date.today().year
        print("No option \"year_for_calendar\" found in configuration file {}, using default: {}".format("config.yml", actual_year))

def check_config():
    check_config_locale()
    check_config_colors()
    check_config_excel_file_name()
    check_config_year()

def print_list(list):
    print("[", end='')
    for element in list:
        print("{},".format(element), end='')
    print("]")

def associate_colors_to_users_if_not_set():
    for user in cal_config['users']:
        if not "color" in user:
            key = random.sample(colors, 1)
            user.update( { "color" : key[0] } )
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

def append_month(year, month_range_list):
    global row_count
    for month in month_range_list:

        # add month name to begin of row
        _, last_day_of_month = monthrange(year, month)
        actual_cell = ws.cell(row=row_count,column=1,value=calendar.month_name[month])
        set_header_colomn_style(actual_cell)

        # add days to same row of month name
        column_position = 0
        for column in range(1,32):
            if column_position < last_day_of_month:
                actual_cell = ws.cell(row=row_count, column=column+1, value=column)
                set_header_colomn_style(actual_cell)
            else:
                actual_cell = ws.cell(row=row_count, column=column+1, value='')
                set_header_colomn_style(actual_cell)
            column_position += 1
        # add month name to end of row
        actual_cell = ws.cell(row=row_count,column=33,value=calendar.month_name[month])
        set_header_colomn_style(actual_cell)
        row_count = row_count + 1

        # add users as separate rows
        for user in cal_config['users']:
            if user['name'] == cal_config['oh_api_show_name']:
                actual_cell = ws.cell(row=row_count,column=1,value=user['name'])
                set_user_cell_style(actual_cell, user["color"])
                for col in range(1,last_day_of_month + 1):
                    actual_cell = ws.cell(row=row_count,column=col+1,value='')
                    if is_date_in_holidays(datetime.datetime(year,month,col)):
                        set_user_cell_style(actual_cell, user["color"])
                actual_cell = ws.cell(row=row_count,column=33,value=user['name'])
                set_user_cell_style(actual_cell, user["color"])
            else:
                actual_cell = ws.cell(row=row_count,column=1,value=user['name'])
                set_user_cell_style(actual_cell, user["color"])
                actual_cell = ws.cell(row=row_count,column=33,value=user['name'])
                set_user_cell_style(actual_cell, user["color"])
            row_count = row_count + 1
        # let two rows empty for spacing
        row_count = row_count + 2


def get_public_holidays():
    global holidays
    global cal_config
    url = cal_config['oh_api_base_url'] + "PublicHolidays?countryIsoCode=" + cal_config['oh_api_country_iso_code'] + "&languageIsoCode=" + cal_config['oh_api_language_iso_code'] + "&validFrom=" + str(actual_year - 1) + "-12-01" + "&validTo=" + str(actual_year + 1) + "-01-31" + "&subdivisionCode=" + cal_config['oh_api_subdivision_code']

    response = requests.get(url)
    response_json = response.json()

    for holiday in response_json:
        startDate = holiday['startDate'].split("-")
        endDate = holiday['endDate'].split('-')
        conv_startDate = datetime.datetime(int(startDate[0]),int(startDate[1]),int(startDate[2]))
        conv_endDate = datetime.datetime(int(endDate[0]),int(endDate[1]),int(endDate[2]))
        holidays.append((conv_startDate,conv_endDate))

def get_school_holidays():
    global holidays
    url = cal_config['oh_api_base_url'] + "SchoolHolidays?countryIsoCode=" + cal_config['oh_api_country_iso_code'] + "&languageIsoCode=" + cal_config['oh_api_language_iso_code'] + "&validFrom=" + str(actual_year - 1) + "-12-01" + "&validTo=" + str(actual_year + 1) + "-01-31" + "&subdivisionCode=" + cal_config['oh_api_subdivision_code']

    response = requests.get(url)
    response_json = response.json()

    for holiday in response_json:
        startDate = holiday['startDate'].split("-")
        endDate = holiday['endDate'].split('-')
        conv_startDate = datetime.datetime(int(startDate[0]),int(startDate[1]),int(startDate[2]))
        conv_endDate = datetime.datetime(int(endDate[0]),int(endDate[1]),int(endDate[2]))
        holidays.append((conv_startDate,conv_endDate))

def add_holidays_as_user():
    new_user = dict({ "name" : cal_config['oh_api_show_name'], "color" : cal_config['oh_api_show_color']})
    cal_config['users'].append(new_user)

def print_holidays():
    for tupel in holidays:
        print(tupel)

def is_date_in_holidays(date):
    global holidays
    return is_date_in_list_tupels(date, holidays)

def is_date_in_list_tupels(date, list_of_date_tupels):
    for start, end in list_of_date_tupels:
        if date >= start and date <= end:
            return True
    return False

############################ MAIN #####################

load_config()

check_config()

locale.setlocale(locale.LC_ALL, cal_config['locale'])

get_school_holidays()

# print_list(holidays)

add_holidays_as_user()

# associate colors to users
associate_colors_to_users_if_not_set()

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
