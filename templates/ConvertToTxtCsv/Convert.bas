Attribute VB_Name = "Convert"
Sub Converttocsv()
Call capturetime
Call MyShape_Click
Dim startTime As Date, endTime As Date, timetaken As Date
startTime = Now()
Dim sName As String, sPath As String
sPath = Sheets("Dashboard").Range("C20") & "\"  '<== change to reflect your folder.  Make sure it ends with a "\" character
sName = Dir(sPath)
Do While sName <> ""
  If LCase(Right(sName, 3)) = "txt" Then
     Workbooks.OpenText Filename:=sPath & sName, _
        Origin:=437, _
        StartRow:=1, _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=True, _
        Semicolon:=False, _
        Comma:=False, _
        Space:=False, _
        Other:=False, _
        TrailingMinusNumbers:=True
     ActiveWorkbook.SaveAs Filename:=sPath & Left(sName, Len(sName) - 4) & ".csv", _
          FileFormat:=xlCSV, _
          CreateBackup:=False
     ActiveWorkbook.Close Savechanges:=False
  End If
  sName = Dir()
Loop
Call captureendtime
endTime = Now(): timetaken = startTime - endTime: [Status].Value = "Success": [Start_Time].Value = startTime: [Time_Taken].Value = Format(timetaken, "HH:MM:SS"): [UserName].Value = Environ("Username")
End Sub
Sub Converttotxt()
Call capturetime
Call MyShape_Click
Dim startTime As Date, endTime As Date, timetaken As Date
startTime = Now()
Dim sName As String, sPath As String
sPath = Sheets("Dashboard").Range("C20") & "\"  '<== change to reflect your folder.  Make sure it ends with a "\" character
sName = Dir(sPath)
Do While sName <> ""
  If LCase(Right(sName, 3)) = "csv" Then
     Workbooks.OpenText Filename:=sPath & sName, _
        Origin:=437, _
        StartRow:=1, _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=True, _
        Semicolon:=False, _
        Comma:=False, _
        Space:=False, _
        Other:=False, _
        TrailingMinusNumbers:=True
     ActiveWorkbook.SaveAs Filename:=sPath & Left(sName, Len(sName) - 4) & ".txt", _
          FileFormat:=xlCSV, _
          CreateBackup:=False
     ActiveWorkbook.Close Savechanges:=False
  End If
  sName = Dir()
Loop
Call captureendtime
endTime = Now(): timetaken = startTime - endTime: [Status].Value = "Success": [Start_Time].Value = startTime: [Time_Taken].Value = Format(timetaken, "HH:MM:SS"): [UserName].Value = Environ("Username")
End Sub
Sub Convertexceltocsv()
Call capturetime
Call MyShape_Click
Dim startTime As Date, endTime As Date, timetaken As Date
startTime = Now()
Dim sName As String, sPath As String
sPath = Sheets("Dashboard").Range("C20") & "\"  '<== change to reflect your folder.  Make sure it ends with a "\" character
sName = Dir(sPath)
Do While sName <> ""
  If LCase(Right(sName, 4)) = "xlsx" Then
     Workbooks.OpenText Filename:=sPath & sName, _
        Origin:=437, _
        StartRow:=1, _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=True, _
        Semicolon:=False, _
        Comma:=False, _
        Space:=False, _
        Other:=False, _
        TrailingMinusNumbers:=True
     ActiveWorkbook.SaveAs Filename:=sPath & Left(sName, Len(sName) - 4) & ".csv", _
          FileFormat:=xlCSV, _
          CreateBackup:=False
     ActiveWorkbook.Close Savechanges:=False
  End If
  sName = Dir()
Loop
Call captureendtime
endTime = Now(): timetaken = startTime - endTime: [Status].Value = "Success": [Start_Time].Value = startTime: [Time_Taken].Value = Format(timetaken, "HH:MM:SS"): [UserName].Value = Environ("Username")
End Sub
Sub Converttoxlsx()
Call capturetime
Call MyShape_Click
Dim startTime As Date, endTime As Date, timetaken As Date
startTime = Now()
Dim sName As String, sPath As String, dPath As String
sPath = Sheets("Dashboard").Range("C20") & "\"  '<== change to reflect your folder.  Make sure it ends with a "\" character
dPath = Sheets("Dashboard").Range("C20") & "\"
sName = Dir(sPath)
Do While sName <> ""
  If LCase(Right(sName, 3)) = "csv" Then
     Workbooks.OpenText Filename:=sPath & sName, _
        Origin:=437, _
        StartRow:=1, _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=False, _
        Semicolon:=False, _
        Comma:=True, _
        Space:=False, _
        Other:=False, _
        TrailingMinusNumbers:=True
     ActiveWorkbook.SaveAs Filename:=dPath & Left(sName, Len(sName) - 4) & ".xlsx", _
          FileFormat:=51, _
          CreateBackup:=False
     ActiveWorkbook.Close Savechanges:=False
  End If
  sName = Dir()
Loop
Call captureendtime
endTime = Now(): timetaken = startTime - endTime: [Status].Value = "Success": [Start_Time].Value = startTime: [Time_Taken].Value = Format(timetaken, "HH:MM:SS"): [UserName].Value = Environ("Username")
End Sub
