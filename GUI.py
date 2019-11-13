import tkinter as tk
from tkinter import ttk


from operation import (
    initialize_file,
    user_input_data,
    open_function)

field_font = ("Times New Roman", "20")
entry_font = ("Times New Roman", "20")
button_font = ("Times New Roman", "20")

class MainUI:

    def __init__(self, root):
        self.root = root

        root.geometry("700x700")
        root.title("Personal finance tracker")
        initialize_file()

        self.Category_Field = ttk.Label(root, text="Category: ", font=field_font)
        self.Category_Field.grid(row=0, column=0, sticky="W", pady=2)

        self.Amount_Spent_Field = ttk.Label(root, text="Amount spent (MYR): ", font=field_font)
        self.Amount_Spent_Field.grid(row=1, column=0, sticky="W", pady=2)

        self.Category_Entry = ttk.Entry(root, width=15, font=entry_font)
        self.Category_Entry.grid(row=0, column=1, pady=2)

        self.Amount_Spent_Entry = ttk.Entry(root, width=15, font=entry_font)
        self.Amount_Spent_Entry.grid(row=1, column=1, pady=2)

        self.Category_Entry.insert(0, "Food")
        self.Amount_Spent_Entry.insert(0, "100")

        self.add_data = tk.Button(root, text="Enter expense", command=self.add_data, font=button_font)
        self.add_data.grid(row=2, column=0, pady=2)

        self.open_excel_file = tk.Button(root, text="Open", command=self.open_button, font=button_font)
        self.open_excel_file.grid(row=2, column=1, pady=2)

        self.quit_button = tk.Button(root, text="Quit", command=root.destroy, font=button_font)
        self.quit_button.grid(row=2, column=2, pady=2)

        self.alerts = tk.Label(root, text="Important messages", font=field_font)
        self.alerts.grid(row=3, column=1, padx=5, pady=5)

    def pop_up(self, string):
        popup = tk.Tk()
        popup.wm_title("Alert!")
        label = ttk.Label(popup, text=string)
        label.pack(side="top", fill="x", pady=10)
        close_button = ttk.Button(popup, text="Close", command=popup.destroy)
        close_button.pack()
        popup.mainloop()

    def error_handler(self, error):
        if isinstance(error, ValueError):
            string = "Invalid entry. Only letters are valid for category and numbers for expenses."
            self.pop_up(string)

        elif isinstance(error, PermissionError):
            string = "Please close the excel file and try again."
            self.pop_up(string)

        else:
            string = "Unknown error. Please restart the program."
            self.pop_up(string)
        return

    def add_data(self):
        category = self.Category_Entry.get()
        expense = self.Amount_Spent_Entry.get()

        try:
            amount = float(expense)
            if category.isalpha():
                user_input_data("Personal Finance", category.title(), amount)
                self.alerts.configure(text=f"Entry saved!"
                                      f"\nCategory: {category}"
                                      f"\nExpense: {expense}")
            else:
                raise ValueError
        except (ValueError, PermissionError) as error:
            self.error_handler(error)
        return

    def open_button(self):
        try:
            open_function()
        except (ValueError, PermissionError) as error:
            self.error_handler(error)
        return