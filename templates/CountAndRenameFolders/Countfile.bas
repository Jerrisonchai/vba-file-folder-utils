Attribute VB_Name = "Countfile"
Option Explicit

Sub CountFiles2()
Call capturetime
Call MyShape_Click
Dim startTime As Date
Dim endTime As Date
Dim timetaken As Date

startTime = Now()

Dim objFSO As Object
Dim objFolder As Object
Dim objSubFolder As Object
Dim objFolder2 As Object
Dim objSubFolder2 As Object
Dim i As Integer

'Create an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Get the folder object
Set objFolder = objFSO.GetFolder(Sheets("Dashboard").Range("C21").Value)
Sheets("Data").Range("A2:C40").Clear
i = 1
'loops through each file in the directory and prints their names and path
For Each objSubFolder In objFolder.SubFolders
    'print folder name
    Sheets("Data").Cells(i + 1, 1) = objSubFolder.Name
    'print folder path
    Sheets("Data").Cells(i + 1, 2) = objSubFolder.Path
    i = i + 1
Next objSubFolder

Dim MyFolder As String
Dim MyFile As String
Dim X As Integer
Dim Folder_Access As Object
Dim Folder As Object
Dim File As Object
Dim count As Integer

X = 2

Do Until X = Sheets("Data").Range("A" & Sheets("Data").Rows.count).End(xlUp).Row + 1
    On Error Resume Next
    MyFolder = Sheets("Data").Range("B" & X)
    MyFile = Dir(MyFolder & "*.xlsx")
    Set Folder_Access = CreateObject("Scripting.FileSystemObject")
    Set Folder = Folder_Access.GetFolder(MyFolder)
    count = 0
    For Each File In Folder.Files
        count = count + 1
    Next File
    Sheets("Data").Range("C" & X).Value = count
    X = X + 1

Loop

Dim a As Integer
a = 2

Do Until a = 40
    Dim sFolder_OldName As String
    Dim sFolder_NewName As String
    On Error Resume Next
    sFolder_OldName = Sheets("Data").Range("B" & a)
    sFolder_NewName = Sheets("Data").Range("D" & a)
    Name sFolder_OldName As sFolder_NewName
    a = a + 1

Loop

Set objFSO = Nothing
Set objFolder = Nothing
Set objSubFolder = Nothing
Set objFolder2 = Nothing
Set objSubFolder2 = Nothing

endTime = Now()
timetaken = startTime - endTime

[Status].Value = "Success"
[Start_Time].Value = startTime
[Time_Taken].Value = Format(timetaken, "HH:MM:SS")
[UserName].Value = Environ("UserName")

Call captureendtime
MsgBox "All folder has been renamed"

End Sub


'To create archive folder with current date
Sub Create_ArchiveFolder_LH()
    
    'Variable declaration
    Dim sFolderName As String, sFolder As String
    Dim sFolderPath As String
    Dim oFSO As Object
    
       
    'Main Folder
    sFolder = Sheets("Dashboard").Range("C21").Value & "\Archive\"
    
    'Folder Name
    sFolderName = Format(Now, "YYYYMMDD_HHMM")
    
    'Folder Path
    sFolderPath = sFolder & sFolderName
        
    'Create FSO Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    'Check Specified Folder exists or not
    If oFSO.FolderExists(sFolderPath) Then
        'If folder is available with today's date
        MsgBox "Folder already exists  with today's date!", vbInformation, "VBAF1"
        Exit Sub
    Else
        'Create Folder
        MkDir sFolderPath
    End If
    Dim MyFile1 As String
    Dim myPath1 As String, myPath2 As String

    myPath1 = Sheets("Dashboard").Range("C21").Value & "\"
    myPath2 = sFolderPath & "\"
    MyFile1 = Dir(myPath1 & "*.*")
    Do While MyFile1 <> ""
        FileCopy myPath1 & MyFile1, myPath2 & MyFile1
        MyFile1 = Dir
    Loop
    
End Sub


