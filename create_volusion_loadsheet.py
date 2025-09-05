import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import chardet
import threading
import os
import re
from html import unescape

processed_df = None
category_mapping = {}   # categoryid -> categoryname
parent_mapping = {}     # categoryid -> parentid
product_file_path = None

def shorten_filename(path, max_length=50):
    filename = os.path.basename(path)
    return filename if len(filename) <= max_length else "..." + filename[-(max_length - 3):]

def update_buttons_state():
    if product_file_path and category_mapping:
        process_button.config(state=tk.NORMAL)
    else:
        process_button.config(state=tk.DISABLED)
        save_button.config(state=tk.DISABLED)

# --- Load Category CSV and build mappings ---
def load_category_file():
    global category_mapping, parent_mapping
    file_path = filedialog.askopenfilename(title="Select Category CSV", filetypes=[("CSV files", "*.csv")])
    if not file_path:
        status_label.config(text="Category file load cancelled.")
        return

    progress_bar.pack(fill='x', padx=20, pady=5)
    progress_bar.start()
    status_label.config(text="Loading category file...")

    def worker():
        global category_mapping, parent_mapping
        category_mapping = {}
        parent_mapping = {}

        try:
            with open(file_path, 'rb') as f:
                raw_data = f.read()
                encoding = chardet.detect(raw_data)['encoding']

            cat_df = pd.read_csv(file_path, encoding=encoding, low_memory=False)

            required_cols = ['categoryid', 'categoryname', 'parentid']
            missing_cols = set(required_cols) - set(cat_df.columns)
            if missing_cols:
                root.after(0, lambda: messagebox.showerror(
                    "Category File Error", f"Missing column(s): {', '.join(missing_cols)}"))
                root.after(0, lambda: status_label.config(text="Invalid category file."))
                return

            # Ensure all IDs are strings
            cat_df['categoryid'] = cat_df['categoryid'].astype(str)
            cat_df['parentid'] = cat_df['parentid'].astype(str)

            category_mapping = dict(zip(cat_df['categoryid'], cat_df['categoryname']))
            parent_mapping = dict(zip(cat_df['categoryid'], cat_df['parentid']))

            root.after(0, lambda: category_filename_label.config(text=f"Category File: {shorten_filename(file_path)}"))
            root.after(0, lambda: status_label.config(text="Category file loaded successfully."))
            root.after(0, update_buttons_state)

        except Exception as e:
            root.after(0, lambda: [
                messagebox.showerror("Category File Error", f"Failed to load category file:\n{e}"),
                status_label.config(text="Error loading category file.")
            ])
        finally:
            root.after(0, lambda: [
                progress_bar.stop(),
                progress_bar.pack_forget()
            ])

    threading.Thread(target=worker, daemon=True).start()

# --- Helper to get true depth from parent chain ---
def get_category_depth(cat_id):
    depth = 0
    current_id = cat_id
    visited = set()
    while current_id in parent_mapping and parent_mapping[current_id] not in (None, '', '0') and current_id not in visited:
        visited.add(current_id)
        current_id = parent_mapping[current_id]
        depth += 1
    return depth + 1  # include current category

# --- Get category at specific depth (fallback to depth-1 if not exist) ---
def get_category_by_depth(ids_str, target_depth):
    if not isinstance(ids_str, str):
        return "Other"
    ids = [id_.strip() for id_ in ids_str.split(',') if id_.strip()]
    # First, try exact target depth
    for id_ in ids:
        if id_ in category_mapping:
            depth = get_category_depth(id_)
            if depth == target_depth:
                return category_mapping[id_]
    # Fallback: target depth - 1
    for id_ in ids:
        if id_ in category_mapping:
            depth = get_category_depth(id_)
            if depth == target_depth - 1:
                return category_mapping[id_]
    return "Other"

# --- Process Product CSV ---
def _process_file_worker(file_path):
    global processed_df

    def clean_description(text):
        if pd.isna(text):
            return ""
        text = unescape(text)
        text = re.sub(r'<[^>]+>', '', text)
        text = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', text)
        text = re.sub(r"[^a-zA-Z0-9\s.,;:!?(){}\[\]\-_'\"&/%+°•$@]", "", text)
        text = re.sub(r'\s+', ' ', text)
        return text.strip()

    try:
        with open(file_path, 'rb') as f:
            raw_data = f.read()
            encoding = chardet.detect(raw_data)['encoding']

        df = pd.read_csv(file_path, encoding=encoding, low_memory=False)

        required_cols = ['productcode', 'productname', 'ischildofproductcode']
        missing_cols = set(required_cols) - set(df.columns)
        if missing_cols:
            raise ValueError(f"Missing required column(s): {', '.join(missing_cols)}")

        # Map parent product titles
        productcode_to_title = df.set_index('productcode')['productname'].to_dict()
        df['Parent Title'] = df['ischildofproductcode'].map(productcode_to_title)

        # Remove child products from main list
        child_product_codes = df['ischildofproductcode'].dropna().unique()
        df = df[~df['productcode'].isin(child_product_codes)]

        # Format price
        if 'productprice' in df.columns:
            df['productprice'] = pd.to_numeric(df['productprice'], errors='coerce') \
                .map(lambda x: f"${x:,.2f}" if pd.notnull(x) else "")

        # Assign Category (Depth 3 preferred, fallback = "Other")
        if 'categoryids' in df.columns and category_mapping:
            df['Category'] = df['categoryids'].map(lambda x: get_category_by_depth(x, 3))
        else:
            df['Category'] = "Other"

        # Clean descriptions
        if 'productdescriptionshort' in df.columns:
            df['productdescriptionshort'] = df['productdescriptionshort'].map(clean_description)

        # Prepare final column list
        final_variant_list = df.copy()
        final_column_list = [
            'productcode', 'productname', 'ischildofproductcode', 'Parent Title',
            'productprice', 'length', 'width', 'height', 'productweight',
            'productdescriptionshort', 'photourl', 'producturl', 'Category'
        ]
        for col in final_column_list:
            if col not in final_variant_list.columns:
                final_variant_list[col] = ""

        final_variant_list = final_variant_list[final_column_list]

        # Rename columns
        final_variant_list.rename(columns={
            'productcode': 'Part #',
            'productname': 'Full Title',
            'ischildofproductcode': 'Parent #',
            'productprice': 'Retail Price',
            'length': 'Length (in)',
            'width': 'Width (in)',
            'height': 'Height (in)',
            'productweight': 'Weight (in)',
            'productdescriptionshort': 'Description',
            'photourl': 'Image Link',
            'producturl': 'Product Link'
        }, inplace=True)

        # --- Fill empty fields with proper defaults ---
        for col in final_variant_list.columns:
            if col == "Category":
            # Treat empty, NaN, or whitespace-only as Other
                final_variant_list[col] = final_variant_list[col].apply(lambda x: "Other" if pd.isna(x) or str(x).strip() == "" else str(x).strip())
            else:
             # Treat empty, NaN, or whitespace-only as "#N/A"
                final_variant_list[col] = final_variant_list[col].apply(lambda x: "#N/A" if pd.isna(x) or str(x).strip() == "" else str(x).strip())

        processed_df = final_variant_list

        root.after(0, lambda: [
            status_label.config(text="Processing complete. You may now save the file."),
            save_button.config(state=tk.NORMAL),
            progress_bar.stop(),
            progress_bar.pack_forget()
        ])

    except Exception as e:
        root.after(0, lambda: [
            status_label.config(text=f"Error: {e}"),
            messagebox.showerror("Processing Error", str(e)),
            save_button.config(state=tk.DISABLED),
            progress_bar.stop(),
            progress_bar.pack_forget()
        ])

# --- GUI Functions ---
def select_product_file():
    global product_file_path
    file_path = filedialog.askopenfilename(title="Select Product CSV", filetypes=[("CSV files", "*.csv")])
    if not file_path:
        return

    category_button.config(state=tk.DISABLED)
    process_button.config(state=tk.DISABLED)
    save_button.config(state=tk.DISABLED)

    progress_bar.pack(fill='x', padx=20, pady=5)
    progress_bar.start()
    status_label.config(text="Loading product file...")

    def worker():
        global product_file_path
        try:
            with open(file_path, 'rb') as f:
                raw_data = f.read()
                encoding = chardet.detect(raw_data)['encoding']
            df = pd.read_csv(file_path, encoding=encoding, low_memory=False)

            required_cols = ['productcode', 'productname', 'ischildofproductcode']
            missing_cols = set(required_cols) - set(df.columns)
            if missing_cols:
                root.after(0, lambda: messagebox.showerror(
                    "Invalid Product File", f"Missing column(s): {', '.join(missing_cols)}"))
                root.after(0, lambda: status_label.config(text="Invalid product file."))
                return

            product_file_path = file_path
            root.after(0, lambda: product_filename_label.config(text=f"Product File: {shorten_filename(file_path)}"))
            root.after(0, lambda: status_label.config(text="Product file loaded. Now select category file."))
            root.after(0, lambda: category_button.config(state=tk.NORMAL))
            root.after(0, update_buttons_state)

        except Exception as e:
            root.after(0, lambda: [
                messagebox.showerror("Product File Error", f"Failed to load product file:\n{e}"),
                status_label.config(text="Error loading product file.")
            ])
        finally:
            root.after(0, lambda: [
                progress_bar.stop(),
                progress_bar.pack_forget()
            ])

    threading.Thread(target=worker, daemon=True).start()

def process_files():
    if not product_file_path:
        messagebox.showwarning("No Product File", "Please select the product CSV file first.")
        return
    if not category_mapping:
        messagebox.showwarning("No Category File", "Please select the category CSV file first.")
        return

    status_label.config(text="Processing...")
    save_button.config(state=tk.DISABLED)
    progress_bar.pack(fill='x', padx=20, pady=5)
    progress_bar.start()

    threading.Thread(target=_process_file_worker, args=(product_file_path,), daemon=True).start()

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
                processed_df.to_csv(output_path, index=False, encoding='utf-8-sig')
                status_label.config(text="File saved successfully.")
                messagebox.showinfo("Success", f"File saved to:\n{output_path}")
            except Exception as e:
                messagebox.showerror("Save Error", str(e))
                status_label.config(text="Error saving file.")
        else:
            status_label.config(text="Save cancelled.")
    else:
        messagebox.showwarning("No Data", "No processed data to save.")

# --- GUI Setup ---
root = tk.Tk()
root.title("EMI Loadsheet Builder")
root.geometry("600x400")

product_button = tk.Button(root, text="1. Select Product CSV File", command=select_product_file)
product_button.pack(padx=20, pady=(20, 5))
product_filename_label = tk.Label(root, text="", fg="gray")
product_filename_label.pack(pady=(0, 10))

category_button = tk.Button(root, text="2. Select Category CSV File", command=load_category_file, state=tk.DISABLED)
category_button.pack(padx=20, pady=(5, 10))
category_filename_label = tk.Label(root, text="", fg="gray")
category_filename_label.pack(pady=(0, 10))

process_button = tk.Button(root, text="3. Process Files", command=process_files, state=tk.DISABLED)
process_button.pack(padx=20, pady=(0, 10))

save_button = tk.Button(root, text="4. Save Processed File", command=save_file, state=tk.DISABLED)
save_button.pack(padx=20, pady=(0, 15))

progress_bar = ttk.Progressbar(root, mode='indeterminate')
status_label = tk.Label(root, text="", fg="blue")
status_label.pack(pady=5)

root.mainloop()
