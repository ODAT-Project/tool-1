#Developed by ODAT project
#please see https://odat.info
#please see https://github.com/ODAT-Project
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from bs4 import BeautifulSoup

#parse html to extract main table and save cleaned data
def clean_data(path):
    #Determine folder paths
    base, fname = os.path.split(path)
    processed_dir = os.path.join(base, "processed_data")
    os.makedirs(processed_dir, exist_ok=True)
    name, _ = os.path.splitext(fname)

    try:
        with open(path, "rb") as f:
            content = f.read()
        soup = BeautifulSoup(content, "lxml")
        table = next((t for t in soup.find_all("table") if "Reference Key" in t.get_text()), None)
        if table is None:
            messagebox.showerror("error", f"file {fname} missing 'Reference Key'")
            return
        rows = table.find_all("tr")
        data = [
            [cell.get_text(strip=True) for cell in row.find_all("td")]
            for row in rows if row.find_all("td")
        ]
        df = pd.DataFrame(data)

        #save cleaned excel in processed_data folder
        excel_out = os.path.join(processed_dir, f"{name}_cleaned.xlsx")
        df.to_excel(excel_out, index=False, engine="openpyxl")

        #remove first row and save as csv in same processed_data folder
        drop_first_row(excel_out)
    except Exception as e:
        messagebox.showerror("error", f"failed to process {fname}: {e}")

#remove first row from excel and output csv
def drop_first_row(path):
    df = pd.read_excel(path, skiprows=1)
    df.to_excel(path, index=False, engine="openpyxl")
    file_dir, fname = os.path.split(path)
    name, _ = os.path.splitext(fname)
    csv_out = os.path.join(file_dir, f"{name}_cleaned.csv")
    df.to_csv(csv_out, index=False)

#ask user to select excel files and process them
def select_files():
    file_paths = filedialog.askopenfilenames(title="Select excel file(s)", filetypes=[("Excel files","*.xls *.xlsx")])
    if not file_paths:
        return
    total = len(file_paths)
    window, bar, label = show_progress_window(root, total)
    for count, path in enumerate(file_paths, start=1):
        clean_data(path)
        bar['value'] = count
        label.config(text=f"processing file {count} of {total}")
        window.update_idletasks()
    label.config(text="completed!")
    window.after(1000, window.destroy)
    messagebox.showinfo("done", "run complete and files cleaned!")

#show progress window with bar and label
def show_progress_window(parent, total):
    window = tk.Toplevel(parent)
    window.title("progress")
    window.geometry("300x100")
    label = tk.Label(window, text="processing...")
    label.pack(pady=10)
    bar = ttk.Progressbar(window, orient="horizontal", length=250, mode="determinate", maximum=total)
    bar.pack(pady=10)
    return window, bar, label

#show about dialog
def show_about():
    messagebox.showinfo("About", "HA long format data cleaner\n\ndeveloped by ODAT project")

#setup gui root window
root = tk.Tk()
root.title("HA long format data cleaner")
root.geometry("400x300")
btn = tk.Button(root, text="open excel files", command=select_files)
btn.pack(pady=20)
about_btn = tk.Button(root, text="About", command=show_about)
about_btn.pack()
root.mainloop()
