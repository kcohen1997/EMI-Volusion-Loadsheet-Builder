import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import chardet
import threading

# --- Float Validation ---
def validate_float_input(new_value):
    if new_value == "":
        return True
    try:
        float(new_value)
        return True
    except ValueError:
        return False

# --- Text Placeholder Function ---
def add_placeholder(entry, text):
    entry.insert(0, text)
    entry.config(fg="gray")

    def on_focus_in(event):
        if entry.get() == text:
            entry.delete(0, tk.END)
            entry.config(fg="black")

    def on_focus_out(event):
        if entry.get() == "":
            entry.insert(0, text)
            entry.config(fg="gray")

    entry.bind("<FocusIn>", on_focus_in)
    entry.bind("<FocusOut>", on_focus_out)

# --- Background Processing (Background Worker Function) ---
def _process_file_worker(file_path):
    try:
        with open(file_path, 'rb') as f:
            raw_data = f.read()
            result = chardet.detect(raw_data)
            encoding = result['encoding']

        # Step 1: Read csv file. Confirm if required columns are in the file.
        df = pd.read_csv(file_path, encoding=encoding)
        required_columns = ['productcode', 'productname', 'ischildofproductcode', 'productprice',
                            'length', 'width', 'height', 'productweight', 'productdescriptionshort', 'photourl']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
        
        # Step 2: Check to make sure that product price of items are greater than 0
        df['productprice'] = pd.to_numeric(df['productprice'], errors='coerce')
        invalid_prices = df['productprice'].isna().sum()
        if invalid_prices > 0:
            raise ValueError(f"{invalid_prices} row(s) have invalid or missing 'productprice' values.")

        # Step 3: Filter out parent ids
        variant_list = df[required_columns].copy()
        child_product_codes = variant_list['ischildofproductcode'].dropna().unique()
        variant_list = variant_list[~variant_list['productcode'].isin(child_product_codes)]
        variant_list.drop(columns=['ischildofproductcode'], inplace=True)

        # Step 4: Calculate price multipliers
        try:
            jobber_multiplier = float(jobber_price_entry.get()) if jobber_price_entry.get() else 0.85
            dealer_multiplier = float(dealer_price_entry.get()) if dealer_price_entry.get() else 0.75
            oemwd_multiplier = float(oemwd_price_entry.get()) if oemwd_price_entry.get() else 0.675
        except ValueError:
            root.after(0, lambda: [
                status_label.config(text="Error: Invalid multiplier"),
                messagebox.showerror("Input Error", "Please enter valid numeric values for the price multipliers."),
                process_button.config(state=tk.NORMAL)
            ])
            return
        if jobber_price_var.get():
            variant_list['Jobber Price'] = round(variant_list['productprice'] * jobber_multiplier, 2)
        if dealer_price_var.get():
            variant_list['Dealer Price'] = round(variant_list['productprice'] * dealer_multiplier, 2)
        if oemwd_price_var.get():
            variant_list['OEM/WD Price'] = round(variant_list['productprice'] * oemwd_multiplier, 2)
        selected_columns = ['productcode', 'productname', 'productprice', 'length', 'width', 'height',
                            'productweight', 'productdescriptionshort', 'photourl']
        if jobber_price_var.get():
            selected_columns.insert(3, 'Jobber Price')
        if dealer_price_var.get():
            selected_columns.insert(3 if not jobber_price_var.get() else 4, 'Dealer Price')
        if oemwd_price_var.get():
            selected_columns.insert(3 + sum([jobber_price_var.get(), dealer_price_var.get()]), 'OEM/WD Price')

        # Step 4: Rename and reorder columns
        variant_list = variant_list[selected_columns]
        variant_list.rename(columns={
            'productcode': 'Part #',
            'productname': 'Title',
            'productprice': 'Retail Price',
            'length': 'Length (in)',
            'width': 'Width (in)',
            'height': 'Height (in)',
            'productweight': 'Weight (lb)',
            'productdescriptionshort': 'Description',
            'photourl': 'Image'
        }, inplace=True)

        global processed_data
        processed_data = variant_list

        # Step 5: Process data successful message
        root.after(0, lambda: [
            save_button.config(state=tk.NORMAL),
            process_button.config(state=tk.NORMAL),
            status_label.config(text="Done"),
            messagebox.showinfo("Success", f"File processed successfully! You can now save your file.")
        ])

    except Exception as e:
        root.after(0, lambda: [
            status_label.config(text="Error during processing"),
            process_button.config(state=tk.NORMAL),
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        ])

# --- Main Process Trigger (on Button) ---
def process_file():
    process_button.config(state=tk.DISABLED)
    save_button.config(state=tk.DISABLED)
    status_label.config(text="Processing...")

    # Validate multipliers before opening file
    try:
        jm = jobber_price_entry.get()
        dm = dealer_price_entry.get()
        om = oemwd_price_entry.get()

        jobber_multiplier = float(jm) if jm else 0.85
        dealer_multiplier = float(dm) if dm else 0.75
        oemwd_multiplier = float(om) if om else 0.675

        for name, value in [('Jobber', jobber_multiplier), ('Dealer', dealer_multiplier), ('OEM/WD', oemwd_multiplier)]:
            if not (0 < value <= 1):
                raise ValueError(f"{name} multiplier must be between 0 and 1.")

    except ValueError as ve:
        messagebox.showerror("Multiplier Error", f"Invalid input:\n{ve}")
        status_label.config(text="Ready")
        process_button.config(state=tk.NORMAL)
        return

    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if not file_path:
        process_button.config(state=tk.NORMAL)
        status_label.config(text="Ready")
        return

    # If everything is valid, proceed with background processing
    threading.Thread(target=_process_file_worker, args=(file_path,), daemon=True).start()

# --- Save File Function ---
def save_file():
    if 'processed_data' not in globals():
        messagebox.showerror("Error", "No data to save!")
        return

    save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if save_path:
        try:
            processed_data.to_csv(save_path, index=False)
            messagebox.showinfo("Success", "File saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while saving the file:\n{str(e)}")

# --- GUI Setup ---
root = tk.Tk()

# GUI Basics
root.title("Create Volusion Loadsheet")
root.geometry("700x600")

# Add label for downloading product list from site
label = tk.Label(root, text="Step 1: Download product list from Volusion if you haven't already (must be in CSV format)", font=("Helvetica", 10, "bold"))
label.pack(pady=10)

# Add label for apply multipliers (optional)
label = tk.Label(root, text="Step 2: Apply multipliers for additional pricing metrics (optional)", font=("Helvetica", 10, "bold"))
label.pack(pady=10)

# Add Multiplier Text Inputs
vcmd = root.register(validate_float_input)

jobber_label = tk.Label(root, text="Jobber Price Multiplier (default 0.85):") # Jobber multiplier
jobber_label.pack(pady=5)
jobber_price_entry = tk.Entry(root, validate="key", validatecommand=(vcmd, '%P'))
jobber_price_entry.pack(pady=5)
add_placeholder(jobber_price_entry, "0.85")

dealer_label = tk.Label(root, text="Dealer Price Multiplier (default 0.75):") # Dealer multiplier
dealer_label.pack(pady=5)
dealer_price_entry = tk.Entry(root, validate="key", validatecommand=(vcmd, '%P'))
dealer_price_entry.pack(pady=5)
add_placeholder(dealer_price_entry, "0.75")

oemwd_label = tk.Label(root, text="OEM/WD Price Multiplier (default 0.675):") # OEM/WD multiplier
oemwd_label.pack(pady=5)
oemwd_price_entry = tk.Entry(root, validate="key", validatecommand=(vcmd, '%P'))
oemwd_price_entry.pack(pady=5)
add_placeholder(oemwd_price_entry, "0.675")

# Checkbox: should we include jobber/dealer, oem/wd price in loadsheet?
jobber_price_var = tk.BooleanVar(value=True)
jobber_check = tk.Checkbutton(root, text="Include Jobber Price", variable=jobber_price_var)
jobber_check.pack(pady=5)

dealer_price_var = tk.BooleanVar(value=True)
dealer_check = tk.Checkbutton(root, text="Include Dealer Price", variable=dealer_price_var)
dealer_check.pack(pady=5)

oemwd_price_var = tk.BooleanVar(value=True)
oemwd_check = tk.Checkbutton(root, text="Include OEM/WD Price", variable=oemwd_price_var)
oemwd_check.pack(pady=5)

# Add label for processing CSV file
label = tk.Label(root, text="Step 3: Click button below to select and process your CSV file", font=("Helvetica", 10, "bold"))
label.pack(pady=10)

# Select CSV input download
process_button = tk.Button(root, text="Select CSV and Process", command=process_file)
process_button.pack(pady=10)

# Status label
status_label = tk.Label(root, text="", fg="blue")
status_label.pack(pady=10)

# Add label for saving processed file
label = tk.Label(root, text="Step 4: Click button below to save your newly processed file", font=("Helvetica", 10, "bold"))
label.pack(pady=10)

# Button to save newly processed file
save_button = tk.Button(root, text="Save Processed CSV", command=save_file, state=tk.DISABLED)
save_button.pack(pady=10)

root.mainloop()
