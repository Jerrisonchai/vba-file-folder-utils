# 📂 VBA File Management: List and Rename Files

This project provides **Excel VBA macros** to:
1. List all files in a specified folder (with metadata such as name, path, and last modified date).  
2. Rename files (e.g., PDFs) based on values in Excel cells.

---

## 🚀 Features
- Extracts **file metadata** (name, path, modified date, folder name).  
- Tracks **execution time** and **user details**.  
- Renames files directly from Excel based on mapping in the worksheet.  
- Updates **status logs** (`Status`, `Start_Time`, `Time_Taken`, `UserName`).  

---

## 📌 Usage
### 1. **List Files in a Folder**
- Folder path is taken from:  
  `Sheets("Dashboard").Range("C21").Value`  
- Output goes to:  
  `Sheets("Data").Range("A2:D400")`  

```vba
'To get the list of files in a folder
Sub LoopThroughFiles()
    Call capturetime
    Call MyShape_Click
    Dim startTime As Date
    Dim endTime As Date
    Dim timetaken As Date

    startTime = Now()

    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim i As Integer
    Sheets("Data").Range("A2:D400").Clear
    Set oFSO = CreateObject("Scripting.FileSystemObject")

    Set oFolder = oFSO.GetFolder(Sheets("Dashboard").Range("C21").Value)
    i = 1

    For Each oFile In oFolder.Files
        Sheets("Data").Cells(i + 1, 1) = oFile.Name
        Sheets("Data").Cells(i + 1, 2) = oFile.DateLastModified
        Sheets("Data").Cells(i + 1, 3) = oFile.Path
        Sheets("Data").Cells(i + 1, 4) = oFolder
        i = i + 1
    Next oFile

    endTime = Now()
    timetaken = startTime - endTime

    [Status].Value = "Success"
    [Start_Time].Value = startTime
    [Time_Taken].Value = Format(timetaken, "HH:MM:SS")
    [UserName].Value = Environ("UserName")

    Call captureendtime
    MsgBox "All file has been listed in Data"
End Sub

```

### 2. Rename Files
- Reads file paths from column C (Data sheet).
- Renames them using values from column F.
- Loops through rows until row 100.

### ✅ Example Output (Data Sheet)
- File Name	Date Last Modified	File Path	Folder Path
- file1.pdf	2025-08-18 12:35	C:\Users\Docs\file1.pdf	C:\Users\Docs
- file2.xlsx	2025-08-17 09:10	C:\Users\Docs\file2.xlsx	C:\Users\Docs

###⚠️ Notes
- Ensure capturetime, captureendtime, and MyShape_Click helper subs exist.
- Modify the loop limit (Do Until X = 100) depending on how many files you need to rename.
- Excel requires file paths to be accessible — make sure files are not open/locked.
