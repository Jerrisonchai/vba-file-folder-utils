Attribute VB_Name = "CodeCounterNew"
Option Explicit
Sub callcodecount()
Call capturetime
Call MyShape_Click
Dim startTime As Date
Dim endTime As Date
Dim timetaken As Date

startTime = Now()

    Dim sFolder As String
    Sheets("Data").Range("A2:F100").Clear
    sFolder = Sheets("Dashboard").Range("C21").Value & "\" 'Put your folder in this cell
    Dim FolderPicker As fileDialog
    Dim mypath As String
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Set FolderPicker = Application.fileDialog(msoFileDialogFolderPicker)
    
        With FolderPicker
            .Title = "Please Choose One Folder"
            .InitialFileName = sFolder
            .AllowMultiSelect = False
            .ButtonName = "Confirm"
                If .Show = -1 Then
                    mypath = .SelectedItems(1)
                    Else
                        End
                End If
        End With
    Sheets("Dashboard").Range("C21").Value = mypath
    sFolder = Sheets("Dashboard").Range("C21").Value & "\"
    
    Dim sFile As String
    Dim wshT As Worksheet
    Dim wshD As Worksheet
    Dim t As Long
    Dim wbkS As Workbook
    'Application.ScreenUpdating = False
    ' Target sheet
    Set wshT = ThisWorkbook.Sheets("Data") ' or use ActiveSheet
    Set wshD = ThisWorkbook.Sheets("Dashboard") ' or use ActiveSheet

    ' Get first Excel filename in the folder
    sFile = Dir(sFolder & "*.xlsm*")
    t = 2
    ' Loop through the files
    Do While sFile <> ""
        ' Open source workbook
        On Error Resume Next
        Application.AskToUpdateLinks = False
        Application.DisplayAlerts = False
        Set wbkS = Workbooks.Open(sFolder & sFile)
        Application.AskToUpdateLinks = True
        Application.DisplayAlerts = True

        Dim CodeLineCount_Var As Object
        Dim CodeLineCount As Double, CodeLineCount_Total As Integer, CodeLineComments As Double
        Dim n As Long, s As String

            Set CodeLineCount_Var = wbkS.VBProject
            CodeLineComments = 0
        
            'counts total lines in modules
            For Each CodeLineCount_Var In CodeLineCount_Var.VBComponents
                CodeLineCount = CodeLineCount + CodeLineCount_Var.CodeModule.CountOfLines
        
                With CodeLineCount_Var.CodeModule
                    For n = 1 To .CountOfLines
        
                        s = .Lines(n, 1)
                        'finds comment line
                        '*********MODIFY CODE HERE to find "Public Sub" or use Right(Trim(s), 1) = ")"
                        '*********Consider getting line number of Right(Trim(s), 1) = ")" and subtracting from next line number with ")"
                        If Left(Trim(s), 1) = "'" Then
                            CodeLineComments = CodeLineComments + 1
                        End If
                    Next n
                End With
            Next CodeLineCount_Var
        
        CodeLineCount_Total = CodeLineCount
        
        wshT.Range("A" & t).Value = sFile
        wshT.Range("C" & t).Value = CodeLineCount_Total
        wshT.Range("D" & t).Value = CodeLineCount_Total - CodeLineComments
        
        t = t + 1
        CodeLineCount = 0
        CodeLineCount_Total = 0
        ' Turn off clipboard
        Application.CutCopyMode = False
        ' Close source workbook
        wbkS.Close Savechanges:=False
        ' Get next filename
        sFile = Dir
    Loop

    'Application.ScreenUpdating = True
    wshD.Activate
    Set wshT = Nothing
    Set wshD = Nothing
    Set wbkS = Nothing
    Application.EnableEvents = True
    Application.ScreenUpdating = True
 MsgBox "CodeCount list has been generated"
endTime = Now()
timetaken = startTime - endTime

[Status].Value = "Success"
[Start_Time].Value = startTime
[Time_Taken].Value = Format(timetaken, "HH:MM:SS")
[UserName].Value = Environ("Username")
 Call captureendtime
End Sub
