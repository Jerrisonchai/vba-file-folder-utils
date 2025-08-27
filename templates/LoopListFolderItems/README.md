# VBA: List Folders and Files

This VBA module provides utilities to **list all files and folders** in a given directory. It captures details such as file name, path, creation date, last modified date, and file size, then outputs them into a designated worksheet.

---

## Features

- Loop through all files in a specified folder.
- List files with metadata:
  - File name
  - Full path
  - Date created
  - Date last modified
  - File size (bytes)
- Supports recursive traversal of subfolders.
- Generates summary tables with additional formulas for file type identification and classification.

---

## Macros

### `LoopThroughFiles`
- Clears existing data in **Sheet "Data"**.
- Reads files from the folder path defined in `Sheets("Dashboard").Range("C20")`.
- Outputs:
  - Column A → File Name  
  - Column B → Last Modified Date
- Logs runtime, username, and timestamps.
- Displays completion message.

---

### `Main_List_Folders_and_Files`
- Clears existing content in **Sheet "List"**.
- Calls the recursive function `List_Folders_and_Files`.
- Populates columns with:
  - File name
  - Full path
  - Date created
  - Last modified
  - File size
- Adds formulas:
  - Extract folder path
  - File extension
  - Template detection (`xlsm` flag)
- Autoformats date, time, and size columns.
- Displays completion message.

---

### `List_Folders_and_Files`
Recursive function that:
- Walks through folders and subfolders.
- Outputs file details into the target cell.
- Returns the total number of files processed.

---

## Sheet Setup

- **Dashboard (C20)** → Folder path input
- **Data (A:B)** → Simple file listing (LoopThroughFiles)
- **List (A:I)** → Detailed recursive file/folder listing with metadata

---

## Example Output

| File Name    | Path                   | Date Created       | Last Modified      | Size (bytes) | Folder Path | File Ext | Type     |
| ------------ | ---------------------- | ------------------ | ------------------ | ------------ | ----------- | -------- | -------- |
| Report.xlsm  | C:\Users\Jerr\Report   | 2024-08-01 12:00   | 2024-08-18 15:30   | 14523        | C:\Users    | xlsm     | Template |
| Notes.txt    | C:\Users\Jerr\Docs     | 2024-07-22 09:10   | 2024-07-23 17:45   | 482          | C:\Users    | txt      |          |

---

## Requirements

- Enable **Microsoft Scripting Runtime** (optional, else late binding used).
- Ensure target worksheets exist:
  - `Dashboard`
  - `Data`
  - `List`

---

## Usage

1. Set folder path in **Dashboard!C20**.
2. Run either:
   - `LoopThroughFiles` → simple file list
   - `Main_List_Folders_and_Files` → detailed recursive listing
3. Check **Sheet Data** or **Sheet List** for output.

---

## Notes
- Large directories with thousands of files may take longer to process.
- Use the recursive version (`Main_List_Folders_and_Files`) for detailed insights.

