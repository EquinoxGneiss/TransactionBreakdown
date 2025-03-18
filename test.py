import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def process_transaction_file(input_file, output_dir):
    """
    Processes a transaction Excel file and outputs a new Excel file
    <originalfilename>_processed.xlsx into the specified output_dir.
    """
    # --- Load the Excel ---
    xls = pd.ExcelFile(input_file)
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
    
    # --- Rename columns (assuming your file layout is consistent) ---
    df.columns = ["Account", "Date", "SRC_SYST_TXN_CD", "Transaction Type", "Description", "Amount"]
    df = df.iloc[3:].reset_index(drop=True)  # remove header-like rows

    # --- Convert data types ---
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")

    # --- Create structured breakdown ---
    structured_df = pd.DataFrame()
    structured_df["Date"] = df["Date"]
    structured_df["Transaction Type"] = df["Transaction Type"]
    structured_df["Amount"] = df["Amount"]
    structured_df["Sender Name"] = df["Description"].str.extract(r'B/O: ([^,]*)')
    structured_df["Sender Bank & Location"] = df["Description"].str.extract(r'VIA: ([^/]*)')
    structured_df["Receiver Name"] = df["Description"].str.extract(r'BNF=([^/]*)')
    structured_df["Receiver Bank & Location"] = df["Description"].str.extract(r'A/C: ([^/]*)')
    structured_df["Reference / Purpose"] = df["Description"].str.extract(r'REF: ([^/]*)')
    structured_df["Transaction ID (IMAD, TRN)"] = df["Description"].str.extract(r'(IMAD: [^ ]*|TRN: [^ ]*)')

    # --- Further breakdown of "Receiver Bank & Location" ---
    structured_df["Receiver Bank"] = structured_df["Receiver Bank & Location"].str.extract(r'([^,]+)')
    structured_df["Receiver Account Name"] = structured_df["Receiver Bank & Location"].str.extract(r', (.*)')
    structured_df.drop(columns=["Receiver Bank & Location"], inplace=True)

    # --- Construct the output file path ---
    base_name = os.path.basename(input_file)          # e.g. "input.xlsx"
    file_root, file_ext = os.path.splitext(base_name) # e.g. "input", ".xlsx"
    output_file = os.path.join(output_dir, f"{file_root}_processed.xlsx")

    # --- Save to Excel ---
    structured_df.to_excel(output_file, index=False)
    return output_file

def process_file():
    """
    Reads the selected file paths from the GUI, validates them,
    calls the processing function, and shows a result message.
    """
    input_file = input_file_var.get()
    output_dir = output_dir_var.get()

    # Basic validation
    if not input_file or not os.path.isfile(input_file):
        messagebox.showerror("Error", "Please select a valid input file.")
        return
    if not output_dir or not os.path.isdir(output_dir):
        messagebox.showerror("Error", "Please select a valid output directory.")
        return

    try:
        output_file = process_transaction_file(input_file, output_dir)
        messagebox.showinfo("Success", f"Processed data saved to:\n{output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

def browse_input_file():
    """Open a file dialog to choose an Excel input file."""
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx *.xls")],
        title="Select Input Excel File"
    )
    if file_path:
        input_file_var.set(file_path)

def browse_output_dir():
    """Open a directory dialog to choose where to save the processed file."""
    folder_path = filedialog.askdirectory(title="Select Output Directory")
    if folder_path:
        output_dir_var.set(folder_path)

# ----------------------- Build the Simple Tkinter GUI -----------------------

root = tk.Tk()
root.title("Simple Transaction Processor")

# Two variables to hold user selections
input_file_var = tk.StringVar()
output_dir_var = tk.StringVar()

# Row 1: Input file
tk.Label(root, text="Input Excel File:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
tk.Entry(root, textvariable=input_file_var, width=50).grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="Browse...", command=browse_input_file).grid(row=0, column=2, padx=5, pady=5)

# Row 2: Output directory
tk.Label(root, text="Output Directory:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
tk.Entry(root, textvariable=output_dir_var, width=50).grid(row=1, column=1, padx=5, pady=5)
tk.Button(root, text="Browse...", command=browse_output_dir).grid(row=1, column=2, padx=5, pady=5)

# Row 3: Process button
tk.Button(root, text="Process File", command=process_file).grid(row=2, column=1, padx=5, pady=15)

root.mainloop()
