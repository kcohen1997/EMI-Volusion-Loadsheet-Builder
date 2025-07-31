# --- Import necessary libraries ---
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import chardet
import threading
import os

# --- Global to hold the processed DataFrame ---
processed_df = None

# --- Helper: Shorten long filenames with ellipsis ---
def shorten_filename(path, max_length=50):
    filename = os.path.basename(path)
    if len(filename) <= max_length:
        return filename
    else:
        return "..." + filename[-(max_length - 3):]

# --- Build full title from options (kept in case needed in future) ---
def build_full_title(row):
    base_title = str(row.get('Title', '')).strip()
    options = []
    for opt_key in ['Option1 Value', 'Option2 Value', 'Option3 Value']:
        val = row.get(opt_key)
        if pd.notna(val):
            val_str = str(val).strip()
            if val_str and val_str.lower() != 'default title':
                options.append(val_str)
    return f"{base_title} - {' - '.join(options)}" if options else base_title

# --- Core file processing logic that runs in a background thread ---
def _process_file_worker(file_path):
    global processed_df
    try:
        with open(file_path, 'rb') as f:
            raw_data = f.read()
            encoding = chardet.detect(raw_data)['encoding']

        df = pd.read_csv(file_path, encoding=encoding, low_memory=False)

        productcode_to_title = df.set_index('productcode')['productname'].to_dict()
        df['Parent Title'] = df['ischildofproductcode'].map(productcode_to_title)

        child_product_codes = df['ischildofproductcode'].dropna().unique()
        df = df[~df['productcode'].isin(child_product_codes)]

        if 'productprice' in df.columns:
            df['productprice'] = pd.to_numeric(df['productprice'], errors='coerce') \
                .map(lambda x: f"${x:,.2f}" if pd.notnull(x) else "")

        final_variant_list = df.copy()
        final_column_list = [
            'productcode', 'productname', 'ischildofproductcode', 'Parent Title', 'productprice',
            'length', 'width', 'height', 'productweight',
            'productdescriptionshort', 'photourl', 'Image 2', 'Image 3', 'producturl'
        ]

        for col in final_column_list:
            if col not in final_variant_list.columns:
                final_variant_list[col] = '#N/A'
        final_variant_list = final_variant_list[final_column_list]

        final_variant_list.rename(columns={
            'productcode': 'Part #',
            'productname': 'Full Title',
            'ischildofproductcode': 'Parent #',
            'Parent Title': 'Title',
            'productprice': 'Retail Price',
            'length': 'Length (in)',
            'width': 'Width (in)',
            'height': 'Height (in)',
            'productweight': 'Weight (in)',
            'productdescriptionshort': 'Description',
            'photourl': 'Image 1',
            'producturl': 'Product Link'
        }, inplace=True)

        final_variant_list.fillna("#N/A", inplace=True)
        processed_df = final_variant_list

        root.after(0, lambda: [
            status_label.config(text="Processing complete. You may now save the file."),
            save_button.config(state=tk.NORMAL)
        ])

    except Exception as e:
        root.after(0, lambda: [
            status_label.config(text=f"Error: {e}"),
            messagebox.showerror("Error", f"An error occurred:\n{e}"),
            save_button.config(state=tk.DISABLED)
        ])

# --- File selection and processing trigger ---
def process_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        status_label.config(text="Processing...")
        save_button.config(state=tk.DISABLED)

        short_name = shorten_filename(file_path)
        filename_label.config(text=f"File: {short_name}")

        threading.Thread(target=_process_file_worker, args=(file_path,), daemon=True).start()

# --- Save processed file ---
def save_file():
    global processed_df
    if processed_df is not None:
        output_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialfile="processed_loadsheet.csv",
            filetypes=[("CSV files", "*.csv")]
        )
        if output_path:
            try:
                processed_df.to_csv(output_path, index=False)
                status_label.config(text="File saved successfully.")
                messagebox.showinfo("Success", f"File saved to:\n{output_path}")
            except PermissionError:
                messagebox.showerror("Permission Error", "The file is open in another program (e.g., Excel). Please close it and try again.")
                status_label.config(text="Error saving file.")
            except Exception as e:
                messagebox.showerror("Save Error", f"Failed to save file:\n{e}")
                status_label.config(text="Error saving file.")
        else:
            status_label.config(text="Save cancelled.")
    else:
        messagebox.showwarning("No Data", "No processed data to save.")

# --- GUI setup ---
root = tk.Tk()
root.title("EMI Loadsheet Builder")
root.geometry("400x150")

# --- Process Button ---
process_button = tk.Button(root, text="Select and Process CSV File", command=process_file)
process_button.pack(padx=20, pady=(20, 5))

# --- Filename Label (moved directly under the process button) ---
filename_label = tk.Label(root, text="", fg="gray")
filename_label.pack(pady=(0, 10))

# --- Save Button ---
save_button = tk.Button(root, text="Save Processed File", command=save_file, state=tk.DISABLED)
save_button.pack(padx=20, pady=(0, 10))

# --- Status Label ---
status_label = tk.Label(root, text="", fg="blue")
status_label.pack(pady=(0, 10))

# --- Run the GUI event loop ---
root.mainloop()
