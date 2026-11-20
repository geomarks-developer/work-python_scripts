import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from pathlib import Path


folder = Path(__file__).resolve().parent.parent.parent

FILE = folder / "Report Engine Report List.xlsx"


# --------------------------------------------------
# Functions
# --------------------------------------------------

def load_sheet():

    global df

    df = pd.read_excel(
        FILE,
        sheet_name=sheet_var.get()
    )

    refresh_tree()


def refresh_tree():

    for item in tree.get_children():
        tree.delete(item)

    for idx, row in df.iterrows():

        tree.insert(
            "",
            "end",
            iid=str(idx),
            values=(
                row["Active"],
                row["ReportName"],
                row["EmailTo"]
            )
        )


def load_selected(event):

    selected = tree.selection()

    if not selected:
        return

    row_idx = int(selected[0])

    active_var.set(bool(df.loc[row_idx, "Active"]))
    report_var.set(df.loc[row_idx, "ReportName"])
    email_var.set(df.loc[row_idx, "EmailTo"])

    sql_text.delete("1.0", tk.END)
    sql_text.insert(
        "1.0",
        str(df.loc[row_idx, "SQLQuery"])
    )


def update_record():

    selected = tree.selection()

    if not selected:
        messagebox.showwarning(
            "Warning",
            "Select a report first"
        )
        return

    row_idx = int(selected[0])

    df.loc[row_idx, "Active"] = active_var.get()
    df.loc[row_idx, "ReportName"] = report_var.get()
    df.loc[row_idx, "EmailTo"] = email_var.get()
    df.loc[row_idx, "SQLQuery"] = sql_text.get(
        "1.0",
        tk.END
    ).strip()

    refresh_tree()

    messagebox.showinfo(
        "Success",
        "Record updated"
    )


def add_record():

    global df

    new_row = {
        "Active": active_var.get(),
        "ReportName": report_var.get(),
        "EmailTo": email_var.get(),
        "SQLQuery": sql_text.get(
            "1.0",
            tk.END
        ).strip()
    }

    df = pd.concat(
        [df, pd.DataFrame([new_row])],
        ignore_index=True
    )

    refresh_tree()


def delete_record():

    global df

    selected = tree.selection()

    if not selected:
        return

    row_idx = int(selected[0])

    df = df.drop(row_idx).reset_index(drop=True)

    refresh_tree()


def save_excel():

    with pd.ExcelWriter(
        FILE,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name=sheet_var.get(),
            index=False
        )

    messagebox.showinfo(
        "Success",
        "Excel file saved"
    )

# --------------------------------------------------
# Window
# --------------------------------------------------

root = tk.Tk()
root.title("Report Engine Maintenance")
root.geometry("1200x700")

# --------------------------------------------------
# Top Bar
# --------------------------------------------------

top_frame = ttk.Frame(root)
top_frame.pack(fill="x", padx=10, pady=10)

ttk.Label(
    top_frame,
    text="Sheet:"
).pack(side="left")

sheet_var = tk.StringVar(value="AdHoc")

sheet_combo = ttk.Combobox(
    top_frame,
    textvariable=sheet_var,
    values=["Daily", "Weekly", "Monthly", "AdHoc"],
    width=20
)

sheet_combo.pack(side="left", padx=5)

ttk.Button(
    top_frame,
    text="Load",
    command=load_sheet
).pack(side="left")

# --------------------------------------------------
# Treeview
# --------------------------------------------------

tree_frame = ttk.Frame(root)
tree_frame.pack(
    fill="both",
    expand=True,
    padx=10,
    pady=10
)

tree = ttk.Treeview(
    tree_frame,
    columns=("Active", "ReportName", "EmailTo"),
    show="headings"
)

tree.heading("Active", text="Active")
tree.heading("ReportName", text="Report Name")
tree.heading("EmailTo", text="Email To")

tree.column("Active", width=80)
tree.column("ReportName", width=300)
tree.column("EmailTo", width=350)

tree.pack(fill="both", expand=True)

tree.bind(
    "<<TreeviewSelect>>",
    load_selected
)

# --------------------------------------------------
# Edit Form
# --------------------------------------------------

form = ttk.Frame(root)
form.pack(fill="x", padx=10, pady=10)

active_var = tk.BooleanVar()
report_var = tk.StringVar()
email_var = tk.StringVar()

ttk.Checkbutton(
    form,
    text="Active",
    variable=active_var
).grid(row=0, column=0, sticky="w")

ttk.Label(
    form,
    text="Report Name"
).grid(row=1, column=0, sticky="w")

ttk.Entry(
    form,
    textvariable=report_var,
    width=80
).grid(row=1, column=1, sticky="ew")

ttk.Label(
    form,
    text="Email To"
).grid(row=2, column=0, sticky="w")

ttk.Entry(
    form,
    textvariable=email_var,
    width=80
).grid(row=2, column=1, sticky="ew")

ttk.Label(
    form,
    text="SQL Query"
).grid(row=3, column=0, sticky="nw")

sql_text = tk.Text(
    form,
    height=10,
    width=100
)

sql_text.grid(
    row=3,
    column=1,
    pady=5,
    sticky="ew"
)

# --------------------------------------------------
# Buttons
# --------------------------------------------------

button_frame = ttk.Frame(root)
button_frame.pack(fill="x", padx=10, pady=10)

ttk.Button(
    button_frame,
    text="Add",
    command=add_record
).pack(side="left", padx=5)

ttk.Button(
    button_frame,
    text="Update",
    command=update_record
).pack(side="left", padx=5)

ttk.Button(
    button_frame,
    text="Delete",
    command=delete_record
).pack(side="left", padx=5)

ttk.Button(
    button_frame,
    text="Save to Excel",
    command=save_excel
).pack(side="left", padx=5)

# --------------------------------------------------
# Initial Load
# --------------------------------------------------

load_sheet()

root.mainloop()