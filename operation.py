from openpyxl import Workbook
from openpyxl.chart import (
    PieChart,
    ProjectedPieChart,
    Reference
)
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

        worksheet_1 = create_worksheet(data_book, "Personal Finance", 0, "05CEBD")
        enter_data(worksheet_1, f"{date_column}2", "Date")
        enter_data(worksheet_1, f"{day_column}2", "Day")
        enter_data(worksheet_1, f"{time_column}2", "Time")
        enter_data(worksheet_1, f"{category_column}2", "Category")
        enter_data(worksheet_1, f"{expense_column}2", "Expense")

        create_worksheet(data_book, "Expenses by Category", 1, "05CEBD")
        create_worksheet(data_book, "Monthly Expenses", 2, "05CEBD")

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


def load_book(path, sheet):
    wb = load_workbook(path)
    ws = wb[sheet]
    max_row = str(ws.max_row+1)
    return wb, ws, max_row


def user_input_data(path, sheet, category, amount):
    wb, ws, max_row = load_book(path, sheet)

    date, day, time = get_date_time()

    enter_data(ws, f"{date_column}{max_row}", date)  # Entering date
    enter_data(ws, f"{day_column}{max_row}", day)  # Entering day
    enter_data(ws, f"{time_column}{max_row}", time)  # Entering time
    enter_data(ws, f"{category_column}{max_row}", category)  # Entering category
    enter_data(ws, f"{expense_column}{max_row}", amount)  # Entering expense

    wb.save(path)
    return


def filter_data():
    wb, ws, max_row = load_book(file_path, "Personal Finance")
    filtered_list = []
    ignored_words = [None, 'Category']

    for row in ws.iter_rows(max_col=5, min_col=5):
        for cell in row:
            if cell.value not in filtered_list and cell.value not in ignored_words:
                filtered_list.append(cell.value)

    return filtered_list


def create_pie_chart():
    wb, ws, max_row = load_book(file_path, "Personal Finance")

    data = filter_data()
    chart_data = {}

    for item in data:
        item_data = []
        for row in ws.iter_rows(min_row=3, min_col=5, max_col=6):
            if row[0].value == item:
                item_data.append(float(row[1].value))
        total_number = sum(i for i in item_data)
        chart_data[item] = total_number

    chart_worksheet = wb["Expenses by Category"]
    print(chart_data)

    r = 3
    for item in chart_data:
        chart_worksheet.cell(row=r, column=2).value = item
        chart_worksheet.cell(row=r, column=3).value = chart_data[item]
        r+=1

    Category_Expenses = PieChart()
    labels = Reference(chart_worksheet, min_col=2, min_row=3, max_row=r-1)
    data = Reference(chart_worksheet, min_col=3, min_row=2, max_row=r-1)
    Category_Expenses.add_data(data, titles_from_data=True)
    Category_Expenses.set_categories(labels)
    Category_Expenses.title = "Expenses by Category"

    chart_worksheet.add_chart(Category_Expenses, "F1")

    wb.save(file_path)

def create_graph(path, sheet):
    wb, ws, max_row = load_book(path, sheet)




    pass




def auto_adjust_column(path):
    wb = load_workbook(path)
    sheet_list = wb.sheetnames
    for sheet in sheet_list:
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
user_input_data(file_path, "Personal Finance", "Pet", "30.00")
user_input_data(file_path, "Personal Finance", "Rent", "125.00")
user_input_data(file_path, "Personal Finance", "Junk", "5.00")
create_pie_chart()

auto_adjust_column(file_path)
open_file(file_path)