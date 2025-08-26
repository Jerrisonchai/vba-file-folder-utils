Attribute VB_Name = "Emailing"
Option Explicit

'To get original data into Working table sheet
Sub CopytoGroupEmail()
Sheets("GroupEmail").Range("A1:Z400").Clear
Sheets("Data").Range("A1:Z400").Copy
Sheets("GroupEmail").Range("A1:Z400").PasteSpecial xlPasteValues
End Sub

'To remove duplicate row
Sub Remove_Duplicates_Folder()

  Sheets("GroupEmail").Range("A1:Z400").RemoveDuplicates Columns:=9, Header:=xlYes

End Sub

'To find error and delete row
Sub DeleteErrorCountry()
Dim Rng As Range
Dim WorkRng As Range
On Error Resume Next
Set WorkRng = Sheets("Data").Range("I1:I400")
Do
    Set Rng = WorkRng.Find("00 NOT FOUND", LookIn:=xlValues)
    If Not Rng Is Nothing Then
        Rng.EntireRow.Delete
    End If
Loop While Not Rng Is Nothing
Application.ScreenUpdating = True
End Sub

'To send email according to respective customer with attached pdf
Sub SendtoEmail()
If MsgBox("Note01: Have you verified the list?" & vbCrLf & "Note02: Have you opened Outlook app?", vbYesNo) = vbNo Then Exit Sub
Call capturetime
Call MyShape_Click
Dim startTime As Date
Dim endTime As Date
Dim timetaken As Date

startTime = Now()

Call DeleteErrorCountry
Call CopytoGroupEmail
Call Remove_Duplicates_Folder

Dim OutlookApp As Object
Dim OutlookMailItem As Object
Dim myAttachments As Object
Dim attachPath As String
Dim strFolder As String
Dim strEmail As String
'Dim strSubject As String
Dim fsFolder As Object
Dim fsFile As Object
Dim FSO As Object

Dim count, X As Integer
X = 2
Do Until X = Sheets("Dashboard").Range("C23").Value
    strFolder = Sheets("GroupEmail").Range("E" & X).Value
    Set OutlookApp = CreateObject("Outlook.application")
    Set OutlookMailItem = OutlookApp.CreateItem(0)
    Set myAttachments = OutlookMailItem.Attachments
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set fsFolder = FSO.GetFolder(strFolder)
    'attachPath = Range("H" & x).Value
    With OutlookMailItem
        strEmail = Sheets("GroupEmail").Range("j" & X)
        .To = strEmail
        .Subject = Sheets("GroupEmail").Range("k" & X).Value
        .Body = Sheets("GroupEmail").Range("l" & X).Value
        'myAttachments.Add attachPath
        For Each fsFile In fsFolder.Files
            If fsFile.Name Like "*.xlsx" Then
                .Attachments.Add strFolder & "\" & fsFile.Name
            End If
        Next
        .send
    End With
    Set OutlookMailItem = Nothing
    Set OutlookApp = Nothing
    
    X = X + 1
        
Loop

Call Move_Folder
Call DeleteSubfolders
endTime = Now()
timetaken = startTime - endTime

[Status].Value = "Success"
[Start_Time].Value = startTime
[Time_Taken].Value = Format(timetaken, "HH:MM:SS")
[UserName].Value = Environ("UserName")

Call captureendtime
MsgBox "All file has been emailed"

End Sub


'Once done with routine work, move everyting to DONE
Sub Move_Folder()

MkDir Sheets("Dashboard").Range("C24") & "\" & Sheets("Dashboard").Range("C16")
'This example copy all files and subfolders from FromPath to ToPath.
'Note: If ToPath already exist it will overwrite existing files in this folder
'if ToPath not exist it will be made for you.
    Dim FSO As Object
    Dim FromPath As String
    Dim ToPath As String

    FromPath = Sheets("Dashboard").Range("C21").Value
    ToPath = Sheets("Dashboard").Range("C24") & "\" & Sheets("Dashboard").Range("C16")

    Set FSO = CreateObject("scripting.filesystemobject")

    If FSO.FolderExists(FromPath) = False Then
        MsgBox FromPath & " doesn't exist"
        Exit Sub
    End If

    FSO.CopyFolder Source:=FromPath, Destination:=ToPath
    MsgBox "You can find the files and subfolders in " & ToPath

End Sub


Sub DeleteSubfolders()
    
    'Variable declaration
    Dim sFolderPath As String
    Dim FSO As Object
    
     'Define Folder Path
    sFolderPath = Sheets("Dashboard").Range("C21").Value
    
    'Check if slash is added
    If Right(sFolderPath, 1) = "\" Then
        'If added remove it from the specified path
        sFolderPath = Left(sFolderPath, Len(sFolderPath) - 1)
    End If
    
    'Create FSO Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    'Check Specified Folder exists or not
    If FSO.FolderExists(sFolderPath) Then
                        
          'Delete All Subfolders
          FSO.DeleteFolder sFolderPath & "\*.*", True
          
     End If
    
End Sub

