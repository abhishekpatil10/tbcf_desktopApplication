import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import threading
from openpyxl import Workbook
from openpyxl.styles import Border, Side

# Create the main window
root = tk.Tk()
root.title("CSV/Excel File Uploader & Calculator")
root.state('zoomed')  # Fullscreen mode
root.resizable(False, False)  # Disable resizing
root.configure(bg="white")  # Light mode

# Title Label
header_label = tk.Label(root, text="CSV/Excel File Uploader & Calculator", font=("Arial", 18, "bold"), bg="white", fg="black")
header_label.pack(pady=10)

# Frame for file selection
file_frame = tk.Frame(root, bg="white")
file_frame.pack(pady=10)
file_label = tk.Label(file_frame, text="Upload a CSV or Excel file", font=("Arial", 14), bg="white", fg="black")
file_label.pack(side=tk.LEFT, padx=10)
button = tk.Button(file_frame, text="Choose File", command=lambda: upload_file())
button.pack(side=tk.LEFT)

# Variable to store the loaded data
data = None
output_data = []  # To store calculation results

# Treeview frame
treeview_frame = tk.Frame(root, bg="white", highlightbackground="black", highlightthickness=3)
treeview_frame.pack(pady=10, fill="both", expand=True)

# Treeview widget to display the table
treeview = ttk.Treeview(treeview_frame, show="headings")
treeview.pack(side=tk.LEFT, fill="both", expand=True)


# Configure Treeview style to show borders
style = ttk.Style()
style.configure("Treeview", background="white", foreground="black", rowheight=25, fieldbackground="white")
style.configure("Treeview.Heading", font=("Arial", 12, "bold"), background="lightgray", foreground="black")
treeview.tag_configure("oddrow", background="white")
treeview.tag_configure("evenrow", background="#f0f0f0")

# Scrollbars
vertical_scrollbar = ttk.Scrollbar(treeview_frame, orient="vertical", command=treeview.yview)
vertical_scrollbar.pack(side="right", fill="y")
treeview.configure(yscrollcommand=vertical_scrollbar.set)


horizontal_scrollbar = ttk.Scrollbar(treeview_frame, orient="horizontal", command=treeview.xview)
horizontal_scrollbar.pack(side="bottom", fill="x")
treeview.configure(xscrollcommand=horizontal_scrollbar.set)

# Button frame
button_frame = tk.Frame(root, bg="white")
button_frame.pack(pady=10)
run_button = tk.Button(button_frame, text="Run Calculation", command=lambda: start_calculation())
download_button = tk.Button(button_frame, text="Download Excel", command=lambda: start_download())
download_unique_button = tk.Button(button_frame, text="Download Unique Calculation", command=lambda: start_download_unique())

# Loading Label
loading_label = tk.Label(root, text="", font=("Arial", 12, "italic"), bg="white", fg="blue")
loading_label.pack(pady=10)

# Function to show a loading message
def show_loading(message):
    loading_label.config(text=message)
    root.update_idletasks()

def hide_loading():
    loading_label.config(text="")

# Function to load and display the file preview in a table
def upload_file():
    global data
    file_path = filedialog.askopenfilename(
        title="Select a CSV or Excel file",
        filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx"), ("Excel files", "*.xls")]
    )

    if file_path:
        try:
            show_loading("Loading file...")
            if file_path.endswith('.csv'):
                data = pd.read_csv(file_path)
            else:
                data = pd.read_excel(file_path, header=0, skipfooter=0)

            data.columns = data.columns.str.strip().str.lower()
            data = data.dropna(axis=1, how='all')
            data = data.fillna("--")

            for row in treeview.get_children():
                treeview.delete(row)

            treeview["columns"] = list(data.columns)
            treeview["show"] = "headings"

            for col in data.columns:
                treeview.heading(col, text=col, anchor="w")
                treeview.column(col, width=150, anchor="w")

            for _, row in data.iterrows():
                treeview.insert("", "end", values=row.tolist())

            button.pack_forget()
            run_button.pack(side=tk.LEFT, padx=10)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load the file: {str(e)}")
        finally:
            hide_loading()

# Function to start calculation in a separate thread
def start_calculation():
    threading.Thread(target=run_calculation).start()

# Function to run calculations on the uploaded data
def run_calculation():
    global output_data
    show_loading("Calculating...")

    try:
        numeric_columns = ['hmnh', 'drs', 'tds', 'net drs amt', 'hmnh percentage', 'drs percentage', 'tds', 'net amount']
        
        # Convert all necessary columns to numeric
        for col in numeric_columns:
            data[col] = pd.to_numeric(data[col], errors='coerce').fillna(0)

        output_data = []
        for _, row in data.iterrows():
            doctor_name = row['performing doctor name']
            sys_net_amt = row['net amount']
            hmnh = row['hmnh'] * row['hmnh percentage'] / 100
            drs = row['drs'] * row['drs percentage'] / 100
            tds_amt = row['tds'] * row['tds'] / 100
            net_drs_amt = row['net drs amt'] * row['drs percentage'] / 100
            
            output_data.append([doctor_name, sys_net_amt, hmnh, drs, tds_amt, net_drs_amt])

        for row in treeview.get_children():
            treeview.delete(row)

        result_columns = ["Doctor Name", "Net Amount", "HMNH", "DRS", "TDS", "Net DRS Amount"]
        treeview["columns"] = result_columns
        treeview["show"] = "headings"

        for col in result_columns:
            treeview.heading(col, text=col, anchor="w")
            treeview.column(col, width=150, anchor="w")

        for row in output_data:
            treeview.insert("", "end", values=row)

        loading_label.config(text="Calculation completed successfully.")
        download_button.pack(side=tk.LEFT, padx=10)
        download_unique_button.pack(side=tk.LEFT, padx=10)


    finally:
        hide_loading()

# Function to start the download process
def start_download():
    threading.Thread(target=download_excel).start()

# Function to download the output data as an Excel file with cell borders
def download_excel():
    show_loading("Downloading Excel...")
    
    try:
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Save as")
        if file_path:
            wb = Workbook()
            ws = wb.active
            ws.append(["Doctor Name", "Net Amount", "HMNH", "DRS", "TDS", "Net DRS Amount"])
            
            for row in output_data:
                ws.append(row)
            
            border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            for row in ws.iter_rows():
                for cell in row:
                    cell.border = border
            
            wb.save(file_path)
            messagebox.showinfo("Success", "File saved successfully!")
    finally:
        hide_loading()


def start_download_unique():
    threading.Thread(target=download_unique_excel).start()

# Function to download unique calculation results
# Function to download unique calculation results
def download_unique_excel():
    show_loading("Downloading Unique Calculation Excel...")
    
    try:
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Save as")
        if file_path:
            unique_df = pd.DataFrame(output_data, columns=["Doctor Name", "Net Amount", "HMNH", "DRS", "TDS", "Net DRS Amount"])
            unique_df = unique_df.groupby("Doctor Name", as_index=False).sum()
            
            wb = Workbook()
            ws = wb.active
            ws.append(["Doctor Name", "Net Amount", "HMNH", "DRS", "TDS", "Net DRS Amount"])
            
            for row in unique_df.itertuples(index=False, name=None):
                ws.append(row)
            
            border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            for row in ws.iter_rows():
                for cell in row:
                    cell.border = border
            
            wb.save(file_path)
            messagebox.showinfo("Success", "Unique Calculation File saved successfully!")
    finally:
        hide_loading()

# Start the main event loop
root.mainloop()
