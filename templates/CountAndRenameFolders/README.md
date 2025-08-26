# 📂 VBA File & Folder Automation Toolkit

This VBA module provides a set of macros to **list, rename, archive, and manage files/folders** directly from Excel.  
It leverages `FileSystemObject (FSO)` to handle directory operations and integrates with a worksheet-based dashboard.

---

## 🚀 Features

### 1. **CountFiles2**
- Reads folder path from `Dashboard!C21`.
- Lists **all subfolders** into the `Data` sheet.
- Counts the number of files in each folder.
- Optionally **renames folders** (based on `Data!D`).
- Tracks execution time, user, and status.

**Output Table (Data sheet):**

| Column | Content          |
| ------ | ---------------- |
| A      | Subfolder Name   |
| B      | Subfolder Path   |
| C      | File Count       |
| D      | New Folder Name  |

---

### 2. **Create_ArchiveFolder_LH**
- Creates a dated **Archive folder** (format: `YYYYMMDD_HHMM`) inside the main directory.
- Copies all files from the current directory to the archive folder.
- Prevents overwriting if the archive already exists.

**Example Structure:**
- 📁 MainFolder
  - 📁 Archive
    - 20250818_1415
      - File1.xlsx
      - File2.pdf


---

### 3. **LoopThroughFiles**
- Lists **all files** from the folder path (`Dashboard!C21`).
- Extracts filename, modified date, full path, and parent folder.

**Output Table (Data sheet):**

| Column | Content             |
| ------ | ------------------- |
| A      | File Name           |
| B      | Last Modified Date  |
| C      | Full Path           |
| D      | Folder              |

---

### 4. **RenameFile**
- Renames files based on mapping in the `Data` sheet.
- Reads old file path from **Column C**, new file path from **Column F**.

**Mapping Table Example:**

| Column C (Old Path)                | Column F (New Path)                  |
| ---------------------------------- | ------------------------------------ |
| `C:\Folder\File1.pdf`              | `C:\Folder\Invoice_1234.pdf`         |
| `C:\Folder\File2.pdf`              | `C:\Folder\Invoice_5678.pdf`         |

---

## 📊 Status Tracking
Macros update status fields automatically in the workbook:
- `[Status]` → Success/Fail
- `[Start_Time]` → Execution start
- `[Time_Taken]` → Time taken
- `[UserName]` → Windows username

---

## ⚡ Dependencies
- Microsoft Excel with VBA enabled.
- FileSystemObject (`Scripting.FileSystemObject`).
- A sheet named `Dashboard` with:
  - `C21` → main folder path input.
- A sheet named `Data` for results.

---

## 🛠 Usage
1. Open VBA Editor (`Alt+F11`).
2. Insert this module into your project.
3. Ensure your workbook has the required sheets (`Dashboard`, `Data`).
4. Run macros from:
   - The `Macros` menu (`Alt+F8`).
   - Assigned buttons/shapes in Excel.

---

## 🔮 Roadmap
- [ ] Add error logging to a sheet instead of silent `On Error Resume Next`.
- [ ] Support batch rename rules (e.g., regex or prefix/suffix).
- [ ] Add file size and type metadata in `LoopThroughFiles`.

---


