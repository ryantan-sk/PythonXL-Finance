import tkinter as tk
from tkinter import ttk

import operation


def initialize_file():
    operation.create_file()
    operation.delete_sheet("Sheet")


def add_data():
    Category = Category_Entry.get()
    Expense = Amount_Spent_Entry.get()
    operation.user_input_data("Personal Finance", Category, float(Expense))


def open_file():
    operation.delete_sheet("Category Chart")
    operation.delete_sheet("Time Chart")

    operation.create_pie_chart()
    operation.create_bar_chart()

    operation.auto_adjust_column()
    operation.open_file()


# styling
field_font = ("Times New Roman", "20")
entry_font = ("Times New Roman", "20")
button_font = ("Times New Roman", "20")

app = tk.Tk()
app.geometry("700x700")
app.title("Personal finance tracker")

initialize_file()


Category_Field = ttk.Label(app, text="Category: ", font=field_font)
Category_Field.grid(row=0, column=0, sticky="W", pady=2)

Amount_Spent_Field = ttk.Label(app, text="Amount spent (MYR): ", font=field_font)
Amount_Spent_Field.grid(row=1, column=0 , sticky="W", pady=2)

Category_Entry = ttk.Entry(app, width=15, font=entry_font)
Category_Entry.grid(row=0, column=1, pady=2)

Amount_Spent_Entry = ttk.Entry(app, width=15, font=entry_font)
Amount_Spent_Entry.grid(row=1, column=1, pady=2)

Category_Entry.insert(0, "Food")
Amount_Spent_Entry.insert(0, "100")

add_data = tk.Button(app, text="Enter expense", command=add_data, font=button_font)
add_data.grid(row=2, column=0, pady=2)

open_excel_file = tk.Button(app, text="Open", command=open_file, font=button_font)
open_excel_file.grid(row=2, column=1, pady=2)

quit_button = tk.Button(app, text="Quit", command=app.destroy, font=button_font)
quit_button.grid(row=2, column=2, pady=2)

alerts = tk.Label(app, text="Important messages", font=field_font)
alerts.grid(row=3, column=0, padx=10)

app.mainloop()