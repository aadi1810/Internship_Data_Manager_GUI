import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def load_data(filename):
    try:
        data = pd.read_excel(filename)
        data.columns = data.columns.str.strip()  # remove extra spaces
        return data
    except Exception as e:
        messagebox.showerror("Error", f"Error loading file: {e}")
        return None

def clean_data(data):
    data = data.dropna()
    data['Start Date'] = pd.to_datetime(data['Start Date'])
    return data.sort_values(by='Start Date')

def generate_report(data, text_widget):
    text_widget.delete("1.0", tk.END)
    report = "========== Internship & Project Report ==========\n"
    report += f"Total entries: {len(data)}\n"
    report += f"Skills learned so far: {', '.join(data['Skills Learned'].unique())}\n\n"
    report += "Detailed List:\n"
    report += data[['Name', 'Internship/Project', 'Start Date', 'End Date', 'Skills Learned']].to_string(index=False)
    text_widget.insert(tk.END, report)

def save_clean_data(data, filename):
    try:
        data.to_excel(filename, index=False)
        messagebox.showinfo("Success", f"Cleaned data saved to '{filename}'")
    except Exception as e:
        messagebox.showerror("Error", f"Error saving file: {e}")

def browse_file(entry_widget):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    entry_widget.delete(0, tk.END)
    entry_widget.insert(0, file_path)

def run_app(file_entry, text_widget):
    file_name = file_entry.get()
    if not file_name:
        messagebox.showwarning("Warning", "Please select a file")
        return

    cleaned_file_name = "cleaned_data.xlsx"
    data = load_data(file_name)
    if data is not None:
        data = clean_data(data)
        generate_report(data, text_widget)
        save_clean_data(data, cleaned_file_name)

# GUI setup
root = tk.Tk()
root.title("Internship & Academic Data Manager")
root.geometry("800x600")

tk.Label(root, text="Select Excel File:").pack(pady=5)
file_entry = tk.Entry(root, width=60)
file_entry.pack(pady=5)
tk.Button(root, text="Browse", command=lambda: browse_file(file_entry)).pack(pady=5)
tk.Button(root, text="Run", command=lambda: run_app(file_entry, text_widget)).pack(pady=10)

text_widget = tk.Text(root, height=25, width=95)
text_widget.pack(pady=10)

root.mainloop()
