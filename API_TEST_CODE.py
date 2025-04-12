import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os, json, requests, threading

BATCH_SIZE = 100
df = None
file_directory = None  # To store the directory of the selected file

# GUI functions
def browse_file():
    global df, file_directory
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, file_path)
        file_directory = os.path.dirname(file_path)  # Store the directory of the file
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
            total_batches = (len(df) + BATCH_SIZE - 1) // BATCH_SIZE
            count_label.config(text=f"Total Batches: {total_batches}")
            messagebox.showinfo("Success", "File loaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", str(e))

# Logging function
def log(message):
    log_text.insert(tk.END, f"{message}\n")
    log_text.see(tk.END)

# Batch processing functions
def convert_to_batches():
    global df, file_directory
    if df is None:
        messagebox.showerror("Error", "Load Excel file first!")
        return
    if not file_directory:
        messagebox.showerror("Error", "File directory not found!")
        return
    
    json_folder = os.path.join(file_directory, "JSON_Batches")
    os.makedirs(json_folder, exist_ok=True)
    
    for old_file in os.listdir(json_folder):
        os.remove(os.path.join(json_folder, old_file))
    
    total_batches = (len(df) + BATCH_SIZE - 1) // BATCH_SIZE
    
    for batch_num, start in enumerate(range(0, len(df), BATCH_SIZE), 1):
        batch_df = df.iloc[start:start+BATCH_SIZE]
        serial_numbers = batch_df["SerialNo"].astype(str).str.cat(sep=';').split(';')
        batch_json = {
            "Key": str(batch_df["ProductionOrder"].iloc[0]),
            "OrderState": "1",
            "MaterialCode": str(batch_df["MaterialCode"].iloc[0]),
            "OrderQuantity": str(batch_df["Order Qty."].iloc[0]),
            "Priority": "5",
            "SerialNo": serial_numbers
        }
        
        with open(os.path.join(json_folder, f"batch_{batch_num}_request.json"), 'w') as f:
            json.dump(batch_json, f, indent=2)
    
    messagebox.showinfo("Success", f"Converted {total_batches} batches successfully!")
    log(f"Converted {total_batches} batches successfully!")

def submit_batches():
    global file_directory
    if not file_directory:
        messagebox.showerror("Error", "No JSON folder found!")
        return
    
    json_folder = os.path.join(file_directory, "JSON_Batches")
    batch_files = [f for f in os.listdir(json_folder) if f.endswith('_request.json')]
    total_batches = len(batch_files)
    
    if total_batches == 0:
        messagebox.showerror("Error", "No batches found to submit!")
        return
    
    # Hardcoded credentials and URL
    username = "SAP"
    password = "SAP@2024"
    api_url = "https://jcmestest.jinchencorp.com:8850/JCMESAPI/WorkOrder/CreateWorkOrderAndLot"
    
    def api_thread():
        for idx, batch_file in enumerate(sorted(batch_files), 1):
            with open(os.path.join(json_folder, batch_file), 'r') as f:
                batch_data = json.load(f)
            
            try:
                response = requests.post(api_url, json=batch_data, auth=(username, password), timeout=60)
                response.raise_for_status()
                response_json = response.json()
                status = "Success"
            except requests.exceptions.ConnectTimeout:
                response_json = {"error": "Connection timed out"}
                status = "Timeout Error"
            except requests.exceptions.HTTPError as e:
                response_json = {"error": f"HTTP Error: {e.response.status_code}"}
                status = "HTTP Error"
            except requests.exceptions.RequestException as e:
                response_json = {"error": str(e)}
                status = "Failed"
            except json.JSONDecodeError:
                response_json = {"error": "Invalid JSON response"}
                status = "Invalid JSON"
            
            response_file = os.path.join(json_folder, batch_file.replace("request", "response"))
            with open(response_file, 'w') as f:
                json.dump(response_json, f, indent=2)
            
            log(f"Batch {idx}/{total_batches} submitted: {status}")
        
        messagebox.showinfo("Info", "Batch submission completed. Check logs for details.")
    
    threading.Thread(target=api_thread, daemon=True).start()

# GUI Setup
root = tk.Tk()
root.title("Excel to JSON Batch Converter")

frame = tk.Frame(root, padx=20, pady=20)
frame.grid()

entry_file = ttk.Entry(frame, width=60)
entry_file.grid(row=0, column=0, padx=5, pady=5)

btn_browse = ttk.Button(frame, text="Browse", command=browse_file)
btn_browse.grid(row=0, column=1, padx=5, pady=5)

btn_convert = ttk.Button(frame, text="Convert to Batches", command=lambda: threading.Thread(target=convert_to_batches, daemon=True).start())
btn_convert.grid(row=1, column=0, padx=5, pady=10)

btn_submit = ttk.Button(frame, text="Submit Batches", command=submit_batches)
btn_submit.grid(row=1, column=1, padx=5, pady=10)

count_label = ttk.Label(frame, text="Total Batches: 0")
count_label.grid(row=2, column=0, padx=5, pady=5)

log_text = tk.Text(frame, width=80, height=15)
log_text.grid(row=3, column=0, columnspan=2, padx=5, pady=5)

root.mainloop()