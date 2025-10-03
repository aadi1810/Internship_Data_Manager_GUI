import pandas as pd
import tkinter as tk
import tkinter.font as tkFont
from tkinter import ttk, messagebox, filedialog
import os
from datetime import datetime

# ================== CONFIG ===================
excel_file = "cleaned_data.xlsx"
backup_folder = "backups"
if not os.path.exists(backup_folder):
    os.makedirs(backup_folder)

# Load Excel file
df = pd.read_excel(excel_file)

# ================== MAIN WINDOW ===================
root = tk.Tk()
root.title("📊 Internship & Academic Data Manager")
root.geometry("1000x800")
root.configure(bg="#2c3e50")  # Dark background

# ================== STYLING ===================
style = ttk.Style()
style.theme_use("clam")
style.configure("Treeview", background="#ffffff", foreground="black", rowheight=25, fieldbackground="#f4f4f9")
style.configure("Treeview.Heading", font=("Arial", 11, "bold"))
style.map("Treeview", background=[('selected', '#6fa1f2')], foreground=[('selected', 'white')])

font_heading = tkFont.Font(family="Arial", size=12, weight="bold")
font_normal = tkFont.Font(family="Arial", size=10)
font_title = tkFont.Font(family="Arial", size=16, weight="bold")

button_style = {"bg": "#4CAF50", "fg": "white", "font": font_normal, "width": 15, "relief": "raised"}

# ================== TITLE ===================
title_label = tk.Label(root, text="📊 Internship & Academic Data Manager", bg="#2c3e50", fg="white", font=font_title)
title_label.pack(pady=10)

# ================== SEARCH FRAME ===================
search_frame = tk.Frame(root, bg="#2c3e50")
search_frame.pack(fill="x", padx=10, pady=5)

tk.Label(search_frame, text="Search:", bg="#2c3e50", fg="white", font=font_heading).pack(side="left")
search_var = tk.StringVar()
search_entry = tk.Entry(search_frame, textvariable=search_var, bg="#ecf0f1", fg="black", font=font_normal)
search_entry.pack(side="left", padx=5)

def search():
    query = search_var.get().lower()
    filtered = df[df.apply(lambda row: query in str(row).lower(), axis=1)]
    update_table(filtered)

tk.Button(search_frame, text="Search", command=search, **button_style).pack(side="left", padx=5)

# ================== TABLE ===================
tree = ttk.Treeview(root)
tree["columns"] = list(df.columns)
tree["show"] = "headings"

for col in df.columns:
    tree.heading(col, text=col)
    tree.column(col, width=120)

tree.pack(expand=True, fill="both", padx=10, pady=10)

def update_table(dataframe):
    tree.delete(*tree.get_children())
    for _, row in dataframe.iterrows():
        tree.insert("", "end", values=list(row))

update_table(df)

# ================== DETAILS FRAME ===================
details_frame = tk.Frame(root, bg="#2c3e50")
details_frame.pack(fill="x", padx=10, pady=10)

details_text = tk.Text(details_frame, height=5, bg="#ecf0f1", fg="black", font=font_normal)
details_text.pack(expand=True, fill="both", padx=10, pady=5)

# ================== EDIT FRAME ===================
edit_frame = tk.Frame(root, bg="#2c3e50")
edit_frame.pack(fill="x", padx=10, pady=10)

entries = {}
for i, col in enumerate(df.columns):
    tk.Label(edit_frame, text=col, bg="#2c3e50", fg="white", font=font_normal).grid(row=0, column=i, padx=5)
    var = tk.StringVar()
    entry = tk.Entry(edit_frame, textvariable=var, bg="#ecf0f1", fg="black", font=font_normal)
    entry.grid(row=1, column=i, padx=5)
    entries[col] = var

selected_index = None

def show_details(event):
    global selected_index
    selected = tree.focus()
    values = tree.item(selected, "values")
    if values:
        selected_index = tree.index(selected)
        details_text.delete("1.0", tk.END)
        details_text.insert(tk.END, "\n".join(f"{col}: {val}" for col, val in zip(df.columns, values)))
        for col, val in zip(df.columns, values):
            entries[col].set(val)

tree.bind("<Double-1>", show_details)

# ================== BACKUP FUNCTION ===================
def backup_file():
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"backup_{timestamp}.xlsx"
    backup_path = os.path.join(backup_folder, backup_name)
    df.to_excel(backup_path, index=False)
    print(f"Backup created: {backup_path}")

# ================== RESTORE BACKUP ===================
def restore_backup():
    backup_path = filedialog.askopenfilename(initialdir=backup_folder, title="Select backup file",
                                             filetypes=(("Excel files", ".xlsx"), ("All files", ".*")))
    if backup_path:
        global df
        df = pd.read_excel(backup_path)
        df.to_excel(excel_file, index=False)
        update_table(df)
        messagebox.showinfo("Success", "Backup restored successfully.")

# ================== SAVE FUNCTION ===================
def save_changes():
    global df
    if selected_index is None:
        messagebox.showwarning("Warning", "No row selected to edit.")
        return
    for col in df.columns:
        df.at[selected_index, col] = entries[col].get()
    df.to_excel(excel_file, index=False)
    backup_file()
    update_table(df)
    messagebox.showinfo("Success", "Changes saved and backup created.")

# ================== DELETE FUNCTION ===================
def delete_record():
    global df, selected_index
    if selected_index is None:
        messagebox.showwarning("Warning", "No row selected to delete.")
        return
    df = df.drop(df.index[selected_index]).reset_index(drop=True)
    df.to_excel(excel_file, index=False)
    backup_file()
    update_table(df)
    details_text.delete("1.0", tk.END)
    selected_index = None
    messagebox.showinfo("Success", "Record deleted and backup created.")

# ================== BUTTONS FRAME ===================
button_frame = tk.Frame(root, bg="#2c3e50")
button_frame.pack(fill="x", padx=10, pady=5)

tk.Button(button_frame, text="Save Changes", command=save_changes, **button_style).pack(side="left", padx=5)
tk.Button(button_frame, text="Delete Record", command=delete_record, **button_style).pack(side="left", padx=5)
tk.Button(button_frame, text="Restore Backup", command=restore_backup, bg="#e67e22", fg="white", font=font_normal, width=20).pack(side="left", padx=5)

# ================== TOOLTIP FUNCTION ===================
def create_tooltip(widget, text):
    tooltip = tk.Toplevel(widget)
    tooltip.withdraw()
    tooltip.overrideredirect(True)
    tooltip_label = tk.Label(tooltip, text=text, bg="yellow", font=font_normal)
    tooltip_label.pack()

    def enter(event):
        tooltip.deiconify()
        tooltip.geometry(f"+{event.x_root + 10}+{event.y_root + 10}")

    def leave(event):
        tooltip.withdraw()

    widget.bind("<Enter>", enter)
    widget.bind("<Leave>", leave)

create_tooltip(search_entry, "Type keywords to search the table.")

root.mainloop()
