from openpyxl import Workbook
from openpyxl import load_workbook

import os
import datetime

file_path = "C:\\Users\\user\\PycharmProjects\\PythonXL-Finance\\workbook\\FinanceTracker.xlsx"

date_column = "B"
day_column = "C"
time_column = "D"
category_column = "E"
expense_column = "F"


def create_worksheet(workbook, name: str, position: int, colour: str):
    worksheet = workbook.create_sheet(name, position)
    worksheet.sheet_properties.tabColor = colour
    return worksheet


def enter_data(sheet, cell, data):
    sheet[cell] = data
    return


def create_file(path):
    file_check = os.path.isfile(path)

    if file_check:
        print("File exists.")
        return
    else:
        data_book = Workbook()

        worksheet = create_worksheet(data_book, "Personal Finance", 0, "05CEBD")
        enter_data(worksheet, f"{date_column}2", "Date")
        enter_data(worksheet, f"{day_column}2", "Day")
        enter_data(worksheet, f"{time_column}2", "Time")
        enter_data(worksheet, f"{category_column}2", "Category")
        enter_data(worksheet, f"{expense_column}2", "Expense")

        data_book.save(path)
        return


def open_file(path):
    os.system(f"start EXCEL.EXE {path}")
    return


def get_date_time():
    date_time = datetime.datetime.now()
    date = date_time.strftime("%d/%m/%Y")
    day = date_time.strftime("%A")
    time = date_time.strftime(("%H:%M:%S %p"))
    return date, day, time


def user_input_data(path, sheet, category, amount):
    wb = load_workbook(path)
    ws = wb[sheet]
    max_row = str(ws.max_row + 1)

    date, day, time = get_date_time()

    enter_data(ws, f"{date_column}{max_row}", date)  # Entering date
    enter_data(ws, f"{day_column}{max_row}", day)  # Entering day
    enter_data(ws, f"{time_column}{max_row}", time)  # Entering time
    enter_data(ws, f"{category_column}{max_row}", category)  # Entering category
    enter_data(ws, f"{expense_column}{max_row}", amount)  # Entering expense

    wb.save(path)
    return


def auto_adjust_column(path, sheet):
    wb = load_workbook(path)
    ws = wb[sheet]
    for column in ws.columns:
        max_length = 0
        column_name = column[0].column_letter

        for row in column:
            try:
                if len(str(row.value)) > max_length:
                    max_length = len(row.value)
            except:
                pass
        new_width = (max_length + 3) * 1.3
        ws.column_dimensions[column_name].width = new_width
    wb.save(file_path)
    return


create_file(file_path)
user_input_data(file_path, "Personal Finance", "Food", "25.00")
auto_adjust_column(file_path, "Personal Finance")
open_file(file_path)

