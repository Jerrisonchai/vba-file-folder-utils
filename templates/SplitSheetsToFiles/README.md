# 📑 SplitSheets Macro

This VBA macro automates the process of splitting all worksheets from multiple Excel workbooks into separate files. Each worksheet will be saved as an individual `.xlsx` file in a newly created **split** folder.

---

## ⚡ Features
- Lets the user pick a folder containing Excel files.  
- Automatically creates a `split` folder inside the chosen path.  
- Loops through each `.xls`, `.xlsx`, `.xlsm` file in the folder.  
- Splits each worksheet into a **new standalone Excel file**.  
- Filenames are constructed as:
  - <OriginalWorkbookName>-<WorksheetName>.xlsx
 

- Preserves performance with `OptimizedMode` (turns off screen updating, events, etc.).  
- Updates status and processing time in predefined named ranges: `[Status]`, `[Start_Time]`, `[Time_Taken]`, `[UserName]`.  
- Displays a message box once the process is complete.  

---

## 🛠 Requirements
- Excel with VBA enabled  
- References: none required (uses late binding for FileSystemObject and FileDialog)  
- A worksheet named **Dashboard** with:  
  - `C20` → default folder path to start from  

---

## 📜 Code
```vba
Option Explicit

Sub SplitSheets()

    Call capturetime
    Call MyShape_Click
    Dim startTime As Date
    Dim endTime As Date
    Dim timetaken As Date

    startTime = Now()

    Dim sFolder As String
    sFolder = Sheets("Dashboard").Range("C20").Value & "\" 'Put your folder in this cell
    Dim sFile As String

    Dim wshO As Worksheet
    Set wshO = ThisWorkbook.Sheets("Dashboard")
    Dim wbkS As Workbook

    OptimizedMode True
    
    Dim FolderPicker As FileDialog
    Dim mypath As String

    Set FolderPicker = Application.FileDialog(msoFileDialogFolderPicker)
    
    With FolderPicker
        .Title = "Please Choose One"
        .InitialFileName = sFolder
        .AllowMultiSelect = False
        .ButtonName = "Confirm"
        If .Show = -1 Then
            mypath = .SelectedItems(1)
        Else
            End
        End If
    End With

    Sheets("Dashboard").Range("C20").Value = mypath
    MkDir mypath & "\split"
    sFolder = Sheets("Dashboard").Range("C20").Value & "\"
    sFile = Dir(sFolder & "*.xls*")

    Application.CopyObjectsWithCells = False
    Do While sFile <> ""
        ' Open source workbook
        Set wbkS = Workbooks.Open(sFolder & sFile)
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim fname As String
        fname = fso.GetBaseName(wbkS.Name)
        
        Dim ws As Worksheet
        Application.DisplayAlerts = False
        For Each ws In wbkS.Sheets
            ws.Copy
            Application.ActiveWorkbook.SaveAs Filename:=mypath & "\split\" & fname & "-" & ws.Name & ".xlsx"
            Application.ActiveWorkbook.Close False
        Next
        Application.DisplayAlerts = True

        ' Close source workbook
        wbkS.Close Savechanges:=False
        ' Get next filename
        sFile = Dir
    Loop
    Application.CopyObjectsWithCells = True

    OptimizedMode False
    
    wshO.Activate
    Set wshO = Nothing
    Set wbkS = Nothing
    Set FolderPicker = Nothing

    endTime = Now()
    timetaken = startTime - endTime

    [Status].Value = "Success"
    [Start_Time].Value = startTime
    [Time_Taken].Value = Format(timetaken, "HH:MM:SS")
    [UserName].Value = Environ("UserName")
    MsgBox "Supplier Data has been splitted"
    Call captureendtime
End Sub
```

## 📊 Output Example
- If your folder contains:
  - Book1.xlsx with sheets Sheet1, Sheet2
  - Book2.xlsx with sheets Summary, Data
- Your split folder will contain:
  - Book1-Sheet1.xlsx
  - Book1-Sheet2.xlsx
  - Book2-Summary.xlsx
  - Book2-Data.xlsx

## 🚀 Usage
- Go to the Dashboard sheet and set C20 with a default path.
- Run SplitSheets macro.
- Select a folder when prompted.
- Wait until process completes — output saved inside a split folder.

## 📝 Notes
- The macro overwrites any existing folder named split.
- Does not copy VBA modules — only worksheets.
- Optimized for large files with OptimizedMode.
