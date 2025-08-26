# VBA File Format Converters

This module provides a set of **VBA macros** to batch convert files between common formats (`TXT`, `CSV`, and `XLSX`).  
It is designed to automate repetitive conversion tasks directly from Excel, with logging for execution time and user tracking.  

---

## 🔧 Features

- **Convert TXT ➝ CSV**  
- **Convert CSV ➝ TXT**  
- **Convert XLSX ➝ CSV**  
- **Convert CSV ➝ XLSX**  
- Logs:
  - Execution status  
  - Start time  
  - Time taken  
  - Username  

---

## 📂 Usage

1. Place all the source files in the folder defined in:  

   ```vba
   sPath = Sheets("Dashboard").Range("C20") & "\"
   ```
⚠️ Ensure the folder path cell (Dashboard!C20) ends with a \ character.

2. Run one of the following macros:
- Converttocsv → Converts all .txt files in folder to .csv
- Converttotxt → Converts all .csv files in folder to .txt
- Convertexceltocsv → Converts all .xlsx files in folder to .csv
- Converttoxlsx → Converts all .csv files in folder to .xlsx
3. Results are saved in the same folder as the source files.

---

##📝 Output Logging
Each macro updates the dashboard with the following info:

| Column/Name | Description                         |
| ----------- | ----------------------------------- |
| Status      | Success / Failure                   |
| Start\_Time | Start time of macro                 |
| Time\_Taken | Duration in `HH:MM:SS`              |
| UserName    | Windows username that ran the macro |

---

##⚡ Example
Example folder (Dashboard!C20 = C:\Data\):
- Input:
  - data1.txt
  - data2.txt
- Run: Converttocsv
- Output:
  - data1.csv
  - data2.csv

---

##✅ Notes
- Ensure macros capturetime, captureendtime, and MyShape_Click exist in your workbook.
- Adjust Workbooks.OpenText parameters if your delimiter settings differ.
- Recommended to backup files before running conversions (files are overwritten).
