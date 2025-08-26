Attribute VB_Name = "Createfolder"
Option Explicit

'To create folder according to cell value
Sub CreateSubFolder() 'PLace the name of the folder in A1
Call capturetime
Call MyShape_Click
Dim startTime As Date
Dim endTime As Date
Dim timetaken As Date

startTime = Now()

Dim X As Integer
    
    X = 2
    
    Do Until X = Sheets("Dashboard").Range("C22").Value
        On Error Resume Next
        MkDir Sheets("Data").Range("E" & X)
        X = X + 1
    Loop

Call FSOMoveFile
Call MovePDFsToAnotherFolder
endTime = Now()
timetaken = startTime - endTime

[Status].Value = "Success"
[Start_Time].Value = startTime
[Time_Taken].Value = Format(timetaken, "HH:MM:SS")
[UserName].Value = Environ("UserName")

Call captureendtime
MsgBox "All file has been moved"
End Sub

'To move existing single file to new folder
Sub FSOMoveFile()
    Dim FSO As New FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")

Dim X As Integer
    
    X = 2
    Do Until X = Sheets("Dashboard").Range("C22").Value
        FSO.MoveFile Sheets("Data").Range("C" & X), Sheets("Data").Range("F" & X)
        X = X + 1
    Loop

End Sub

'To undo move file
Sub FSOReverseMoveFile()
Call capturetime
Call MyShape_Click
Dim startTime As Date
Dim endTime As Date
Dim timetaken As Date

startTime = Now()
    Dim FSO As New FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")

Dim X As Integer
    
    X = 2
    Do Until X = Sheets("Dashboard").Range("C22").Value
        FSO.MoveFile Sheets("Data").Range("F" & X), Sheets("Data").Range("C" & X)
        X = X + 1
    Loop

Call DeleteSubfolders
endTime = Now()
timetaken = startTime - endTime

[Status].Value = "Success"
[Start_Time].Value = startTime
[Time_Taken].Value = Format(timetaken, "HH:MM:SS")
[UserName].Value = Environ("UserName")

Call captureendtime
MsgBox "All file has been returned"
End Sub

'To move all pdfs
Sub MovePDFsToAnotherFolder()

MkDir Sheets("Dashboard").Range("C21") & "\" & "No Email"

Dim FSO As Object, sourcePath As String, destPath As String
Dim Fldr As Object, f As Object, ct As Long
sourcePath = Sheets("Dashboard").Range("C21") & "\"  'Change path and folder name to suit
destPath = Sheets("Dashboard").Range("C21") & "\" & "No Email" & "\" 'Change path and folder name to suit
Set FSO = CreateObject("Scripting.FileSystemObject")
Set Fldr = FSO.GetFolder(sourcePath).Files
For Each f In Fldr
    If f.Name Like "*.xlsx*" Or f.Name Like "*xlsx*" Then
        ct = ct + 1
        FSO.MoveFile Sheets("Dashboard").Range("C21") & "\" & f.Name, Sheets("Dashboard").Range("C21") & "\" & "No Email" & "\" & f.Name
    End If
Next f
If ct > 0 Then
    MsgBox ct & " files have no email address"
Else
    MsgBox "No files have no email address"
End If
End Sub

