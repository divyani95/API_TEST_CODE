# Excel API Uploader

## 📌 Overview
This application allows users to **upload Excel files**, **convert data into JSON batches**, and **submit them to an API** for processing. It is designed for bulk data handling, supporting **batch processing of 100 rows per request**.

## 🚀 Features
- **Upload and process Excel files** (.xlsx, .xls)
- **Convert data into JSON batches** for API submission
- **Save batch request & response JSON files** for record-keeping
- **Progress indicator & multi-threaded processing** to keep the GUI responsive
- **Error handling & validation checks**

## 🛠️ Technologies Used
- **Python** (Core logic)
- **Tkinter** (GUI framework)
- **Pandas** (Excel file processing)
- **Requests** (API communication)
- **Threading** (Asynchronous batch submission)

## 📂 Project Structure
```
📂 Excel API Uploader
 ├── 📜 main.py  # Main script
 ├── 📂 Batch_JSON_Files  # Stores batch request/response JSONs
 ├── 📜 README.md  # Project documentation
 ├── 📜 requirements.txt  # Dependencies
```

## ⚙️ Installation
1. **Clone the repository**
```bash
git clone https://github.com/yourusername/excel-api-uploader.git
cd excel-api-uploader
```

2. **Install dependencies**
```bash
pip install -r requirements.txt
```

3. **Run the application**
```bash
python main.py
```

## 🔹 How to Use
1. **Click 'Browse'** to select an Excel file.
2. **Click 'Convert to Batches'** to create JSON files.
3. **Click 'Submit'** to send the data to the API in batches.
4. **View API responses** in the application.

## 📌 API Endpoint
- **Primary API:** `http://115.244.229.166:8000/JCMESAPI/WorkOrder/CreateWorkOrderAndLot`
- **Alternative API:** `https://jcmestest.jinchencorp.com:8850/JCMESAPI/WorkOrder/CreateWorkOrderAndLot`

## 📝 JSON Format (Example)
```json
[
    {
        "Key": "123456",
        "OrderState": "1",
        "MaterialCode": "MAT001",
        "OrderQuantity": "500",
        "Priority": "5",
        "SerialNo": ["S001", "S002"]
    }
]
```

## 🛠️ Troubleshooting
- **Issue:** GUI not opening
  - ✅ Run `python main.py` in the terminal.
- **Issue:** API request failed
  - ✅ Check internet connection & API URL.
- **Issue:** JSON files not saving
  - ✅ Ensure `Batch_JSON_Files/` exists & has write permissions.

## 📜 License
This project is **open-source**. Feel free to modify and use it!
