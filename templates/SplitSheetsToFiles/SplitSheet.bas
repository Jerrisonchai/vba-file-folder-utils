Attribute VB_Name = "SplitSheet"
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

    Set wshO = ThisWorkbook.Sheets("Dashboard") ' or use ActiveSheet
    Dim wbkS As Workbook

    OptimizedMode True
    
    Dim FolderPicker As FileDialog
    Dim mypath As String

    Set FolderPicker = Application.FileDialog(msoFileDialogFolderPicker)
    
        With FolderPicker
            .Title = "Please Choose One"
            .InitialFileName = sFolder
            .AllowMultiSelect = False
'            .Filters.Clear
'            .Filters.Add "Custom Excel Files", "*.xlsx, *.csv, *.xls"
            .ButtonName = "Confirm"
                If .Show = -1 Then
                    mypath = .SelectedItems(1)
                    Else
                        End
                End If
        End With
    Sheets("Dashboard").Range("C20").Value = mypath
    MkDir mypath & "\split"
    sFolder = Sheets("Dashboard").Range("C20").Value & "\" 'Put your folder in this cell
    sFile = Dir(sFolder & "*.xls*")
    ' Loop through the files
    Application.CopyObjectsWithCells = False
    Do While sFile <> ""
        ' Open source workbook
'        On Error Resume Next
        Set wbkS = Workbooks.Open(sFolder & sFile)
'        Set wshS = wbkS.Worksheets(1)
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
