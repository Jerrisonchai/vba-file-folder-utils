# 📂 VBA File & Folder Utilities

A curated collection of **VBA templates and modules** for automating file and folder operations in Windows & Excel.  
This library is designed for Business Analysts, Developers, and Power Users who frequently handle bulk file tasks such as renaming, exporting, and organizing files.  

---

## ✨ Features
- Count lines of VBA code across modules  
- Convert Excel sheets to TXT/CSV files  
- Count files in folders and auto-rename  
- Create folders for email automation (auto-move after sent)  
- Loop and list items inside a folder  
- Rename all files in bulk  
- Split each worksheet into a separate file  

---

## 📂 Repository Structure
vba-file-folder-utils/
 - README.md <- Main overview (this file)
/templates <- Individual templates/modules
- CodeCounter/
- ConvertToTxtCsv/
- CountAndRenameFolders/
- CreateFoldersEmailing/
- LoopListFolderItems/
- RenameAllFiles/
- SplitSheetsToFiles/

/docs
- USER_GUIDE.md
- DEV_GUIDE.md
- WORKFLOW.md

/tests
- test_cases.md
- performance_benchmark.md


Each **template** has its own subfolder with:  
- `Module.bas` → the core VBA module  
- `README.md` → usage guide for the specific template  
- `sample.xlsm` → demo file (if applicable)  

---

## 🚀 Getting Started

### 1. Import a Template
1. Download the `.bas` file from the template folder (e.g., `CodeCounter.bas`).  
2. In Excel, press `ALT + F11` to open the VBA editor.  
3. Go to `File > Import File...` and select the `.bas`.  

### 2. Run the Macro
1. Press `ALT + F8` in Excel.  
2. Select the desired macro (e.g., `CountLines`) and click **Run**.  

---

## 📘 Templates Included
| Template | Description | Link |
|----------|-------------|------|
| **Code Counter** | Count lines of VBA code across all modules | [View →](./templates/CodeCounter/README.md) |
| **Convert To Txt/Csv** | Convert Excel sheets into `.txt` or `.csv` files | [View →](./templates/ConvertToTxtCsv/README.md) |
| **Count & Rename Folders** | Count files in folders and auto-rename folders | [View →](./templates/CountAndRenameFolders/README.md) |
| **Create Folders for Emailing** | Auto-create folders for email exports, move processed files to “Done” | [View →](./templates/CreateFoldersEmailing/README.md) |
| **Loop & List Folder Items** | Loop through all files/folders and list them in Excel | [View →](./templates/LoopListFolderItems/README.md) |
| **Rename All Files** | Bulk rename files in a directory | [View →](./templates/RenameAllFiles/README.md) |
| **Split Sheets into Files** | Export each worksheet as an individual Excel file | [View →](./templates/SplitSheetsToFiles/README.md) |

---

## 🧪 Tests & Benchmark
- Functional test cases are listed in [`/tests/test_cases.md`](./tests/test_cases.md)  
- Performance benchmarks (e.g., large folders, many files) are tracked in [`/tests/performance_benchmark.md`](./tests/performance_benchmark.md)  

---

## 🤝 Contributing
Contributions are welcome!  
- Fork the repo  
- Add your template under `/templates/`  
- Include a **README.md** and **sample file**  
- Submit a Pull Request  

---
