import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import json
import requests
import os
import threading

BATCH_SIZE = 100  # Number of rows per batch
df = None  # Global variable to store dataframe
JSON_FOLDER = r"C:\Users\DELL\Downloads\AVAADA\API_DATA\Batch_JSON_Files"  # New directory

def browse_file():
    global df
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", ".xlsx;.xls")])
    if file_path:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, file_path)
        try:
            df = pd.read_excel(file_path)  # Load data once
            total_batches = (len(df) // BATCH_SIZE) + (1 if len(df) % BATCH_SIZE != 0 else 0)
            count_label.config(text=f"Total Batches: {total_batches}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

def show_loader(loader_text):
    progress_bar.grid()  # Show progress bar
    progress_bar.start(10)
    text_output.insert(tk.END, f"{loader_text}...\n")
    text_output.update_idletasks()
    btn_convert.config(state=tk.DISABLED)
    btn_submit.config(state=tk.DISABLED)
    btn_browse.config(state=tk.DISABLED)  # Disable browse button

def hide_loader():
    progress_bar.stop()
    progress_bar.grid_remove()  # Hide progress bar
    btn_convert.config(state=tk.NORMAL)
    btn_submit.config(state=tk.NORMAL)
    btn_browse.config(state=tk.NORMAL)  # Enable browse button

def convert_to_batches():
    global df
    if df is None:
        messagebox.showerror("Error", "Please select an Excel file first!")
        return

    show_loader("Converting data to batches")

    total_rows = len(df)
    total_batches = (total_rows // BATCH_SIZE) + (1 if total_rows % BATCH_SIZE != 0 else 0)
    
    text_output.delete(1.0, tk.END)

    # Create folder to store JSON files
    os.makedirs(JSON_FOLDER, exist_ok=True)

    # Clear any old batch JSON files
    for file in os.listdir(JSON_FOLDER):
        os.remove(os.path.join(JSON_FOLDER, file))
    
    # Split data into batches and save as JSON files
    for batch_num, start in enumerate(range(0, total_rows, BATCH_SIZE), start=1):
        batch_df = df.iloc[start:start + BATCH_SIZE]

        # Prepare to collect SerialNo across all rows in the batch
        serial_numbers = []
        
        for _, row in batch_df.iterrows():
            serial_numbers.extend([str(x) for x in row["SerialNo"].split(';')])  # Assuming 'SerialNo' are separated by semicolons

        json_data = [{
            "Key": str(batch_df["ProductionOrder"].iloc[0]),
            "OrderState": "1",
            "MaterialCode": str(batch_df["MaterialCode"].iloc[0]),
            "OrderQuantity": str(batch_df["Order Qty."].iloc[0]),
            "Priority": str(5),
            "SerialNo": serial_numbers
        }]

        request_json_file = os.path.join(JSON_FOLDER, f"batch_{batch_num}_request.json")
        try:
            with open(request_json_file, "w") as f:
                json.dump(json_data, f, indent=4)
            print(f"Saved Request JSON: {request_json_file}")
        except Exception as e:
            print(f"Error saving request JSON: {e}")

    hide_loader()

    messagebox.showinfo("Success", "Data has been converted to batches successfully!")
    count_label.config(text=f"Total Batches: {total_batches}")
    text_output.insert(tk.END, f"Total {total_batches} Batches converted to JSON files.\n")
    text_output.update_idletasks()

def submit_data():
    def process_batches():
        username = "SAP"
        password = "SAP"

        batch_files = [f for f in os.listdir(JSON_FOLDER) if f.endswith('_request.json')]

        if not batch_files:
            messagebox.showerror("Error", "No batch JSON files found. Please convert data to batches first!")
            return

        show_loader("Submitting batches to API")

        total_batches = len(batch_files)
        text_output.delete(1.0, tk.END)

        for batch_num, batch_file in enumerate(batch_files, start=1):
            with open(os.path.join(JSON_FOLDER, batch_file), "r") as f:
                json_data = json.load(f)

            # api_url = "https://jcmestest.jinchencorp.com:8850/JCMESAPI/WorkOrder/CreateWorkOrderAndLot"
            api_url = "http://115.244.229.166:8000/JCMESAPI/WorkOrder/CreateWorkOrderAndLot"
            # http://115.244.229.166:8000/JCMESAPI/WorkOrder/CreateWorkOrderAndLot
            headers = {"Content-Type": "application/json"}

            try:
                response = requests.post(api_url, json=json_data, auth=(username, password), headers=headers, timeout=30)
                response_data = response.json()

                response_json_file = os.path.join(JSON_FOLDER, f"batch_{batch_num}_response.json")
                try:
                    with open(response_json_file, "w") as f:
                        json.dump(response_data, f, indent=4)
                    print(f"Saved Response JSON: {response_json_file}")
                except Exception as e:
                    print(f"Error saving response JSON: {e}")

                text_output.insert(tk.END, f"Batch {batch_num}/{total_batches} Processed!\n")
                text_output.insert(tk.END, json.dumps(response_data, indent=4) + "\n\n")
                text_output.update_idletasks()

                count_label.config(text=f"Processed Batch {batch_num}/{total_batches}")
                root.update_idletasks()

            except Exception as e:
                messagebox.showerror("Error", f"API request failed for Batch {batch_num}: {e}")
                hide_loader()
                return

        hide_loader()

        messagebox.showinfo("Success", "All batches processed successfully!")
        count_label.config(text=f"All {total_batches} Batches Processed")

    threading.Thread(target=process_batches, daemon=True).start()

# GUI Setup
root = tk.Tk()
root.title("Excel API Uploader")
root.geometry("500x400")

tk.Label(root, text="Select Excel File:").grid(row=0, column=0, padx=10, pady=10)
entry_file = tk.Entry(root, width=40)
entry_file.grid(row=0, column=1, padx=10, pady=10)

btn_browse = tk.Button(root, text="Browse", command=browse_file)
btn_browse.grid(row=1, column=0, padx=10, pady=10)

count_label = tk.Label(root, text="Total Batches: 0")
count_label.grid(row=2, column=0, padx=10, pady=10, columnspan=2) 

progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="indeterminate")
progress_bar.grid(row=3, column=0, padx=10, pady=5, columnspan=2)
progress_bar.grid_remove()

btn_convert = tk.Button(root, text="Convert to Batches", command=convert_to_batches)
btn_convert.grid(row=4, column=0, padx=10, pady=10)

btn_submit = tk.Button(root, text="Submit (Send 100 Rows at a Time)", command=submit_data)
btn_submit.grid(row=4, column=1, padx=10, pady=10)

tk.Label(root, text="API Response:").grid(row=5, column=0, padx=10, pady=10)

text_output = tk.Text(root, height=10, width=60)
text_output.grid(row=6, column=0, columnspan=2, padx=10, pady=10)

root.mainloop()
