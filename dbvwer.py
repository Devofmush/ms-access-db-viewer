import os
import pyodbc
import sys
import tkinter as tk
from tkinter import ttk, messagebox

def get_db_path():
    if getattr(sys, 'frozen', False):
        script_dir = os.path.dirname(sys.executable)
    else:
        #if the application is running in a script
        script_dir = os.path.dirname(os.path.abspath(__file__))

    db_path = os.path.join(script_dir, "db.accdb")
    return db_path

#database connection
def connect_to_db():
    db_path = get_db_path()
    try:
        conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db_path + ';')
        return conn
    except Exception as e:
        messagebox.showerror("Database Connection Error", str(e))
        return None

#search
def search_db(event=None):
    search_value = entry.get()
    
    if not search_value:
        selected_item = tree.selection()
        if selected_item:
            item = tree.item(selected_item)
            row_values = item['values']
            search_value = row_values[1]

    if not search_value:
        messagebox.showwarning("Input Error", "Please enter a value to search or select a Document Code")
        return

    conn = connect_to_db()
    if not conn:
        return

    cursor = conn.cursor()
    query = """
        SELECT Id, DocumentCode, ReceptionCode, RankNo, FirstName, LastName, FatherName, 
               IdentifyNo, BirthDate, Address, TelephoneNumber, ReceptionDate, PaymentRate, 
               PatientDescription1, PatientDescription2 
        FROM reception 
        WHERE FirstName LIKE ? 
        OR LastName LIKE ? 
        OR TelephoneNumber LIKE ?
        OR DocumentCode LIKE ?
    """

    search_pattern = f"%{search_value}%"
    cursor.execute(query, (search_pattern, search_pattern, search_pattern, search_pattern))
    results = cursor.fetchall()

    for i in tree.get_children():
        tree.delete(i)

    grouped_results = {}
    for row in results:
        key = (row.FirstName, row.LastName, row.TelephoneNumber)
        if key not in grouped_results:
            grouped_results[key] = []
        grouped_results[key].append(row)

    if grouped_results:
        for key, rows in grouped_results.items():
            first_row = rows[0]
            parent_id = tree.insert('', 'end', values=(
                '', first_row.DocumentCode, '', '', first_row.FirstName, first_row.LastName,
                first_row.FatherName, '', '', '', '', '', '', '', '', ''
            ))
            
            for sub_row in rows:
                tree.insert(parent_id, 'end', values=(
                    sub_row.Id, sub_row.DocumentCode, sub_row.ReceptionCode,
                    sub_row.RankNo, sub_row.FirstName, sub_row.LastName,
                    sub_row.FatherName, sub_row.IdentifyNo, sub_row.BirthDate,
                    sub_row.Address, sub_row.TelephoneNumber, sub_row.ReceptionDate,
                    sub_row.PaymentRate, sub_row.PatientDescription1, sub_row.PatientDescription2
                ))
    else:
        tree.insert('', 'end', values=('', 'No matching records found.', '', '', '', '', '', '', '', '', '', '', '', '', ''))

    conn.close()

def show_row_details(event):
    selected_item = tree.selection()
    if not selected_item:
        return

    parent_id = tree.parent(selected_item)
    if parent_id:
        item = tree.item(selected_item)
    else:
        item = tree.item(selected_item)

    row_values = item['values']

    detail_window = tk.Toplevel(root)
    detail_window.title("Row Details")
    detail_window.state('zoomed')  

    fields = [
        "شماره پرونده", "شماره پذيرش", "نوبت مراجعه", "نام", "نام خانوادگي", 
        "اسم پدر", "شماره شناسايي", "تاريخ تولد", "آدرس", "شماره تلفن", 
        "تاريخ پذيرش", "پرداختي"
    ]

    for i, field in enumerate(fields):
        field_label = tk.Label(detail_window, text=f"{field}:")
        field_label.grid(row=i // 2, column=(i % 2) * 2, padx=10, pady=5, sticky='e')

        value_label = tk.Label(detail_window, text=row_values[i + 1] if (i + 1) < len(row_values) else '')
        value_label.grid(row=i // 2, column=(i % 2) * 2 + 1, padx=10, pady=5, sticky='w')

    for i in range(len(fields)):
        detail_window.grid_columnconfigure(i, weight=1)

    #patient Description
    desc1_label = tk.Label(detail_window, text="توضيحات 1:")
    desc1_label.grid(row=len(fields), column=0, padx=10, pady=5, sticky='e')
    desc1_text = tk.Text(detail_window, height=4, wrap=tk.WORD)
    desc1_text.insert(tk.END, row_values[13])  # PatientDescription1
    desc1_text.config(state=tk.DISABLED)  
    desc1_text.grid(row=len(fields), column=1, padx=10, pady=5, sticky='ew')

    desc2_label = tk.Label(detail_window, text="توضيحات 2:")
    desc2_label.grid(row=len(fields) + 1, column=0, padx=10, pady=5, sticky='e')
    desc2_text = tk.Text(detail_window, height=4, wrap=tk.WORD)
    desc2_text.insert(tk.END, row_values[14])  # PatientDescription2
    desc2_text.config(state=tk.DISABLED)  
    desc2_text.grid(row=len(fields) + 1, column=1, padx=10, pady=5, sticky='ew')

    for col in range(2):
        detail_window.grid_columnconfigure(col, weight=1)

def toggle_collapse(event):
    selected_item = tree.identify_row(event.y)
    if selected_item:
        if tree.item(selected_item, 'open'):
            tree.item(selected_item, open=False)
        else:
            tree.item(selected_item, open=True)

def enable_paste(event):
    entry.event_generate('<<Paste>>')

def adjust_treeview_width():
    total_width = root.winfo_width() - 50
    num_columns = len(columns)
    column_width = total_width // num_columns
    for col in columns:
        tree.column(col, width=column_width)

def copy_selected_document_code(event):
    selected_item = tree.selection()
    if selected_item:
        item = tree.item(selected_item)
        row_values = item['values']
        document_code = row_values[1] if len(row_values) > 1 else row_values[1]
        root.clipboard_clear()
        root.clipboard_append(document_code)
        root.update()
        messagebox.showinfo("Copied", "Document Code copied to clipboard.")

def paste_to_entry(event):
    entry.delete(0, tk.END)
    pasted_text = root.clipboard_get()
    entry.insert(tk.END, pasted_text)

root = tk.Tk()
root.title("Database Search App")
root.state('zoomed')  #maximize main window

entry = tk.Entry(root, width=50)
entry.pack(pady=10)

entry.bind("<Button-3>", enable_paste)

search_button = tk.Button(root, text="Search", command=search_db)
search_button.pack(pady=5)

entry.bind('<Return>', search_db)

frame = tk.Frame(root)
frame.pack(pady=10, fill='both', expand=True)

columns = (
    "Id", "DocumentCode", "ReceptionCode", "RankNo", "FirstName", "LastName", 
    "FatherName", "IdentifyNo", "BirthDate", "Address", "TelephoneNumber", 
    "ReceptionDate", "PaymentRate", "PatientDescription1", "PatientDescription2"
)
tree = ttk.Treeview(frame, columns=columns, show='headings', selectmode="browse")

column_labels = {
    "Id": "ID",
    "DocumentCode": "شماره پرونده",
    "ReceptionCode": "شماره پذيرش",
    "RankNo": "نوبت مراجعه",
    "FirstName": "نام",
    "LastName": "نام خانوادگي",
    "FatherName": "نام پدر",
    "IdentifyNo": "شماره شناسايي",
    "BirthDate": "تاريخ تولد",
    "Address": "آدرس",
    "TelephoneNumber": "شماره تلفن",
    "ReceptionDate": "تاريخ پذيرش",
    "PaymentRate": "پرداختي",
    "PatientDescription1": "توضيحات 1",
    "PatientDescription2": "توضيحات 2",
}

for col in columns:
    tree.heading(col, text=column_labels[col])
    tree.column(col, anchor='center', width=100)

scrollbar_y = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=scrollbar_y.set)
scrollbar_y.pack(side="right", fill="y")

scrollbar_x = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
tree.configure(xscrollcommand=scrollbar_x.set)
scrollbar_x.pack(side="bottom", fill="x")

tree.pack(fill='both', expand=True)

tree.bind("<Double-1>", show_row_details)
tree.bind("<Button-1>", toggle_collapse)
tree.bind("<Control-c>", copy_selected_document_code)

root.bind("<Configure>", lambda e: adjust_treeview_width())

root.mainloop()
