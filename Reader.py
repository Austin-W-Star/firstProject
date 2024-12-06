import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

def load_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if file_path:
        try:
            
            df = pd.read_excel(file_path)
          
            display_data(df)

            global data_frame
            data_frame = df
            global excel_file_path
            excel_file_path = file_path
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel file: {e}")

def display_data(df):

    for row in tree.get_children():
        tree.delete(row)
    
    tree["columns"] = df.columns.tolist()
    tree["show"] = "headings"
    
    for col in df.columns:
        tree.heading(col, text=col)

    for _, row in df.iterrows():
        tree.insert("", "end", values=row.tolist())

def sort_by_date():
    try:
        date_column = data_frame.columns[0]
        
        data_frame[date_column] = pd.to_datetime(data_frame[date_column], errors='coerce')
        
        data_frame.dropna(subset=[date_column], inplace=True)
       
        sorted_data = data_frame.sort_values(by=date_column)
        
        display_data(sorted_data)
        messagebox.showinfo("Success", f"Data sorted by '{date_column}'")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to sort data: {e}")

def save_file():
    if 'data_frame' in globals():
        try:
            sorted_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx;*.xls")])
            if sorted_file_path:
                data_frame.to_excel(sorted_file_path, index=False)
                messagebox.showinfo("Success", "Sorted data saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Excel file: {e}")
    else:
        messagebox.showwarning("Warning", "No data to save!")

root = tk.Tk()
root.title("Excel Sorter by Date")
root.geometry("800x600")

frame = ttk.Frame(root)
frame.pack(fill="both", expand=True, padx=10, pady=10)

load_button = ttk.Button(frame, text="Load Excel File", command=load_file)
load_button.grid(row=0, column=0, padx=10, pady=5)

sort_button = ttk.Button(frame, text="Sort by Date", command=sort_by_date)
sort_button.grid(row=0, column=1, padx=10, pady=5)

save_button = ttk.Button(frame, text="Save Sorted Data", command=save_file)
save_button.grid(row=0, column=2, padx=10, pady=5)

tree = ttk.Treeview(frame, show="headings")
tree.grid(row=1, column=0, columnspan=3, pady=10, sticky="nsew")

frame.grid_rowconfigure(1, weight=1)
frame.grid_columnconfigure(0, weight=1)
frame.grid_columnconfigure(1, weight=1)
frame.grid_columnconfigure(2, weight=1)

root.mainloop()