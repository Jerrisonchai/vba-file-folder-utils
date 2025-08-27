Attribute VB_Name = "LoopFolders"
Sub LoopThroughFiles()
Call capturetime
Call MyShape_Click
Dim startTime As Date
Dim endTime As Date
Dim timetaken As Date
'Dim UserName As String

startTime = Now()
Dim oFSO As Object
Dim oFolder As Object
Dim oFile As Object
Dim i As Integer
Sheets("Data").Range("A2:B4000").Clear
Set oFSO = CreateObject("Scripting.FileSystemObject")

Set oFolder = oFSO.GetFolder(Sheets("Dashboard").Range("C20").Value)
i = 1

For Each oFile In oFolder.Files

    Sheets("Data").Cells(i + 1, 1) = oFile.Name
    Sheets("Data").Cells(i + 1, 2) = oFile.Datelastmodified

    i = i + 1

Next oFile
MsgBox "Files has been listed in Sheet Data"
[Start_Time].Value = startTime
[UserName].Value = Environ("Username")
Call captureendtime
End Sub

