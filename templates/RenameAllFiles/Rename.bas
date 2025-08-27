Attribute VB_Name = "Rename"
'New code
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


'To rename each pdf according to cell value
Sub Renamefile()
Call capturetime
Call MyShape_Click
Dim startTime As Date
Dim endTime As Date
Dim timetaken As Date

startTime = Now()

    'Call FilenMyShape.MyShape_Click
    Dim X As Integer
    X = 2
    Do Until X = 100
    On Error Resume Next
        Name Sheets("Data").Range("C" & X) As _
    Sheets("Data").Range("F" & X)
        X = X + 1
    Loop
endTime = Now()
timetaken = startTime - endTime

[Status].Value = "Success"
[Start_Time].Value = startTime
[Time_Taken].Value = Format(timetaken, "HH:MM:SS")
[UserName].Value = Environ("UserName")

Call captureendtime
MsgBox "All file has been renamed"

End Sub


