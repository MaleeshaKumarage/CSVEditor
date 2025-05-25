import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import csv
import os

try:
    import openpyxl
except ImportError:
    openpyxl = None

CONDITIONS = [
    ("", ""),  # No filter
    ("==", "Equals"),
    (">", "Greater than"),
    ("<", "Less than"),
    (">=", "Greater or equal"),
    ("<=", "Less or equal"),
    ("contains", "Contains"),
    ("not contains", "Not contains"),
]

def select_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("CSV or Excel files", "*.csv *.xlsx")],
        title="Select a CSV or Excel file"
    )
    if file_path:
        ext = os.path.splitext(file_path)[1].lower()
        if ext not in [".csv", ".xlsx"]:
            messagebox.showerror("Error", "Please select a CSV (.csv) or Excel (.xlsx) file.")
            return
        if ext == ".xlsx" and openpyxl is None:
            messagebox.showerror("Error", "openpyxl is required for Excel files. Please install it with 'pip install openpyxl'.")
            return
        file_var.set(file_path)
        load_headers(file_path)
        reset_column_selection()
        update_count_label()  # Update count when file is loaded

def load_headers(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".csv":
        with open(file_path, 'r', newline='', encoding='utf-8') as infile:
            reader = csv.reader(infile)
            headers = next(reader)
    elif ext == ".xlsx":
        wb = openpyxl.load_workbook(file_path, read_only=True)
        ws = wb.active
        headers = [cell.value for cell in next(ws.iter_rows(max_row=1))]
    else:
        headers = []
    all_columns.clear()
    all_columns.extend(headers)
    column_select['values'] = all_columns
    column_select.set('')
    # Remove any previously added columns/conditions
    reset_column_selection()

def reset_column_selection():
    for widget in selected_columns_frame.winfo_children():
        widget.destroy()
    selected_columns.clear()
    filter_widgets.clear()
    update_count_label()

def add_column():
    col = column_select.get()
    if not col or col in selected_columns:
        return
    selected_columns.append(col)
    idx = len(selected_columns) - 1

    row = idx
    tk.Label(selected_columns_frame, text=col).grid(row=row, column=0, padx=2, pady=2)
    cond_var = tk.StringVar(value="")
    cond_menu = ttk.Combobox(selected_columns_frame, textvariable=cond_var, width=12, state="readonly")
    cond_menu['values'] = [label for code, label in CONDITIONS]
    cond_menu.current(0)
    cond_menu.grid(row=row, column=1, padx=2, pady=2)
    entry = tk.Entry(selected_columns_frame, width=14)
    entry.grid(row=row, column=2, padx=2, pady=2)
    remove_btn = tk.Button(selected_columns_frame, text="Remove", command=lambda: remove_column(col))
    remove_btn.grid(row=row, column=3, padx=2, pady=2)
    filter_widgets.append((col, cond_var, entry))

    # Bind events for real-time count update
    cond_menu.bind("<<ComboboxSelected>>", lambda e: update_count_label())
    entry.bind("<KeyRelease>", lambda e: update_count_label())
    update_count_label()

def remove_column(col):
    if col in selected_columns:
        idx = selected_columns.index(col)
        selected_columns.pop(idx)
        filter_widgets.pop(idx)
        # Redraw the selected columns frame
        for widget in selected_columns_frame.winfo_children():
            widget.destroy()
        for i, (col, cond_var, entry) in enumerate(filter_widgets):
            tk.Label(selected_columns_frame, text=col).grid(row=i, column=0, padx=2, pady=2)
            cond_menu = ttk.Combobox(selected_columns_frame, textvariable=cond_var, width=12, state="readonly")
            cond_menu['values'] = [label for code, label in CONDITIONS]
            cond_menu.current(0)
            cond_menu.grid(row=i, column=1, padx=2, pady=2)
            entry.grid(row=i, column=2, padx=2, pady=2)
            remove_btn = tk.Button(selected_columns_frame, text="Remove", command=lambda c=col: remove_column(c))
            remove_btn.grid(row=i, column=3, padx=2, pady=2)
            cond_menu.bind("<<ComboboxSelected>>", lambda e: update_count_label())
            entry.bind("<KeyRelease>", lambda e: update_count_label())
        update_count_label()

def row_matches_filters(row, filters, header):
    for col, cond_var, entry in filters:
        if col not in header:
            return False
        idx = header.index(col)
        value = row[idx]
        cond_label = cond_var.get()
        cond_code = ""
        for code, label in CONDITIONS:
            if label == cond_label:
                cond_code = code
                break
        filter_value = entry.get().strip()
        if not cond_code or not filter_value:
            continue  # No filter for this column
        try:
            if cond_code == "==":
                if str(value) != filter_value:
                    return False
            elif cond_code == ">":
                if float(value) <= float(filter_value):
                    return False
            elif cond_code == "<":
                if float(value) >= float(filter_value):
                    return False
            elif cond_code == ">=":
                if float(value) < float(filter_value):
                    return False
            elif cond_code == "<=":
                if float(value) > float(filter_value):
                    return False
            elif cond_code == "contains":
                if filter_value.lower() not in str(value).lower():
                    return False
            elif cond_code == "not contains":
                if filter_value.lower() in str(value).lower():
                    return False
        except Exception:
            return False
    return True

def get_filtered_rows(input_file, filters, max_records=None):
    ext = os.path.splitext(input_file)[1].lower()
    if ext == ".csv":
        with open(input_file, 'r', newline='', encoding='utf-8') as infile:
            reader = csv.reader(infile)
            header = next(reader)
            filtered_rows = (row for row in reader if row_matches_filters(row, filters, header))
            if max_records is not None:
                limited_rows = [row for _, row in zip(range(max_records), filtered_rows)]
            else:
                limited_rows = list(filtered_rows)
    elif ext == ".xlsx":
        wb = openpyxl.load_workbook(input_file, read_only=True)
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        header = next(rows)
        filtered_rows = (row for row in rows if row_matches_filters(row, filters, header))
        if max_records is not None:
            limited_rows = [row for _, row in zip(range(max_records), filtered_rows)]
        else:
            limited_rows = list(filtered_rows)
    else:
        header = []
        limited_rows = []
    return header, limited_rows

def save_csv(header, rows, output_file):
    with open(output_file, 'w', newline='', encoding='utf-8') as outfile:
        writer = csv.writer(outfile)
        writer.writerow(header)
        writer.writerows(rows)

def save_xlsx(header, rows, output_file):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(header)
    for row in rows:
        ws.append(row)
    wb.save(output_file)

def update_count_label(*args):
    input_file = file_var.get()
    limit = limit_var.get().strip()
    ext = os.path.splitext(input_file)[1].lower()
    if not input_file or not os.path.isfile(input_file) or ext not in [".csv", ".xlsx"]:
        count_var.set("Matching rows: 0")
        return
    if limit:
        try:
            max_records = int(limit)
            if max_records <= 0:
                raise ValueError
        except ValueError:
            count_var.set("Matching rows: 0")
            return
    else:
        max_records = None

    filters = filter_widgets
    try:
        header, filtered_rows = get_filtered_rows(input_file, filters, max_records)
        row_count = len(filtered_rows)
    except Exception:
        row_count = 0
    count_var.set(f"Matching rows: {row_count}")

def process_file():
    input_file = file_var.get()
    limit = limit_var.get().strip()
    ext = os.path.splitext(input_file)[1].lower()
    if not input_file or not os.path.isfile(input_file) or ext not in [".csv", ".xlsx"]:
        messagebox.showerror("Error", "Please select a valid CSV (.csv) or Excel (.xlsx) file.")
        return
    if ext == ".xlsx" and openpyxl is None:
        messagebox.showerror("Error", "openpyxl is required for Excel files. Please install it with 'pip install openpyxl'.")
        return
    if limit:
        try:
            max_records = int(limit)
            if max_records <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid positive integer for the row limit, or leave blank for all.")
            return
    else:
        max_records = None

    filters = filter_widgets

    try:
        header, filtered_rows = get_filtered_rows(input_file, filters, max_records)
        row_count = len(filtered_rows)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while filtering:\n{e}")
        return

    if row_count == 0:
        messagebox.showinfo("No Rows", "No rows match the given filters.")
        return

    confirm = messagebox.askyesno(
        "Confirm Save",
        f"{row_count} row(s) match your filters and will be saved.\n\nDo you want to continue?"
    )
    if not confirm:
        return

    # Set output file type and extension based on input
    if ext == ".csv":
        filetypes = [("CSV files", "*.csv")]
        defaultext = ".csv"
        save_func = save_csv
    else:
        filetypes = [("Excel files", "*.xlsx")]
        defaultext = ".xlsx"
        save_func = save_xlsx

    output_file = filedialog.asksaveasfilename(
        defaultextension=defaultext,
        filetypes=filetypes,
        title="Save filtered file as"
    )
    if not output_file:
        return

    # Export all columns, but filter only on selected columns
    try:
        save_func(header, filtered_rows, output_file)
        messagebox.showinfo("Success", f"Filtered file saved to:\n{output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

# --- Tkinter UI setup ---
root = tk.Tk()
root.title("CSV/Excel Record Limiter")

file_var = tk.StringVar()
limit_var = tk.StringVar(value="100")
all_columns = []
selected_columns = []
filter_widgets = []
count_var = tk.StringVar(value="Matching rows: 0")

tk.Label(root, text="CSV/Excel File:").grid(row=0, column=0, sticky="e")
tk.Entry(root, textvariable=file_var, width=40).grid(row=0, column=1)
tk.Button(root, text="Browse...", command=select_file).grid(row=0, column=2)

tk.Label(root, text="Max Rows (leave blank for all):").grid(row=1, column=0, sticky="e")
limit_entry = tk.Entry(root, textvariable=limit_var, width=15)
limit_entry.grid(row=1, column=1, sticky="w")
limit_entry.bind("<KeyRelease>", lambda e: update_count_label())  # Real-time update

# Column selection dropdown and add button
tk.Label(root, text="Select Column:").grid(row=2, column=0, sticky="e")
column_select = ttk.Combobox(root, values=all_columns, state="readonly", width=30)
column_select.grid(row=2, column=1, sticky="w")
tk.Button(root, text="Add", command=add_column).grid(row=2, column=2, sticky="w")

# Frame for selected columns and their conditions
selected_columns_frame = tk.Frame(root)
selected_columns_frame.grid(row=3, column=0, columnspan=3, pady=10)

count_label = tk.Label(root, textvariable=count_var, fg="blue")
count_label.grid(row=5, column=0, columnspan=3, pady=5)

tk.Button(root, text="Process", command=process_file).grid(row=4, column=1, pady=10)

root.mainloop()
