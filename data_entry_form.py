from pathlib import Path
from tkinter import END, Entry, Label, Tk, Button, Frame

from openpyxl import Workbook, load_workbook

EXCEL_FILE = Path(__file__).with_name("data.xlsx")

name_entry = None
age_entry = None
email_entry = None
phone_entry = None
status_label = None
wb = None
ws = None


def init_excel():
    global wb, ws

    if EXCEL_FILE.exists():
        wb = load_workbook(EXCEL_FILE)
    else:
        wb = Workbook()

    ws = wb.active
    headers = ["Name", "Age", "Email", "Phone"]

    if ws.max_row == 0 or ws.cell(row=1, column=1).value != headers[0]:
        for i, header in enumerate(headers, start=1):
            ws.cell(row=1, column=i).value = header

    wb.save(EXCEL_FILE)
    return wb, ws


def add_data():
    name = name_entry.get().strip()
    age = age_entry.get().strip()
    email = email_entry.get().strip()
    phone = phone_entry.get().strip()

    if name and age and email and phone:
        ws.append([name, age, email, phone])
        wb.save(EXCEL_FILE)
        name_entry.delete(0, END)
        age_entry.delete(0, END)
        email_entry.delete(0, END)
        phone_entry.delete(0, END)
        status_label.config(text="Data added successfully!", fg="green")
    else:
        status_label.config(text="Please fill in all fields.", fg="red")


def clear_form():
    name_entry.delete(0, END)
    age_entry.delete(0, END)
    email_entry.delete(0, END)
    phone_entry.delete(0, END)
    status_label.config(text="Form cleared.", fg="blue")


def next_field(event):
    event.widget.tk_focusNext().focus()


def build_gui():
    global name_entry, age_entry, email_entry, phone_entry, status_label

    root = Tk()
    root.title("Data Entry Form")
    root.geometry("420x260")
    root.columnconfigure(1, weight=1)

    labels = ["Name", "Age", "Email", "Phone"]
    entries = []

    for i, label_text in enumerate(labels):
        Label(root, text=label_text).grid(row=i, column=0, padx=10, pady=5, sticky="e")
        entry = Entry(root)
        entry.grid(row=i, column=1, padx=10, pady=5, sticky="ew")
        entry.bind("<Return>", next_field)
        entries.append(entry)

    name_entry, age_entry, email_entry, phone_entry = entries

    button_frame = Frame(root)
    button_frame.grid(row=4, column=0, columnspan=2, pady=10)
    Button(button_frame, text="Add", width=12, command=add_data).pack(side="left", padx=5)
    Button(button_frame, text="Clear", width=12, command=clear_form).pack(side="left", padx=5)

    status_label = Label(root, text="", anchor="w")
    status_label.grid(row=5, column=0, columnspan=2, padx=10, sticky="w")

    root.mainloop()


if __name__ == "__main__":
    init_excel()
    build_gui()