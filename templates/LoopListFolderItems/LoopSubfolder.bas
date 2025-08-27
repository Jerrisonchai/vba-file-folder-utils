Attribute VB_Name = "LoopSubfolder"
Option Explicit

Public Sub Main_List_Folders_and_Files()

    With ActiveSheet
        Sheets("List").Range("A2:O10000").Clear
        List_Folders_and_Files Sheets("Dashboard").Range("C20").Value, Sheets("List").Range("A1")
    End With
    
    Sheets("List").Range("F2").FormulaArray = _
        "=LEFT(RC[-4],MAX((MID(RC[-4],ROW(R1:R256),1)=""\"")*(ROW(R1:R256))))"
    Sheets("List").Range("G2").Value = "=RC[-3]-RC[-4]"
    Sheets("List").Range("H2").Value = "=SUBSTITUTE(REPLACE(RIGHT(RC[-7],5),1,SEARCH(""."",RIGHT(RC[-7],5)),""""),""."","""")"
    Sheets("List").Range("I2").Value = "=IF(RC[-1]=""xlsm"",""Template"","""")"
    
    Dim data_lastrow As Long
    
    data_lastrow = Sheets("List").Cells(Rows.Count, 1).End(xlUp).Row
    Sheets("List").Range("C2:D" & data_lastrow).NumberFormat = "m/d/yyyy h:mm"
    Sheets("List").Range("E2:E" & data_lastrow).NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"
    Sheets("List").Range("F2:I" & data_lastrow).FillDown
    Sheets("List").Range("G2:G" & data_lastrow).NumberFormat = "0.0"
    
MsgBox "Done!"

End Sub


Private Function List_Folders_and_Files(folderPath As String, destCell As Range) As Long

    Dim FSO As Object
    Dim FSfolder As Object, FSsubfolder As Object, FSfile As Object
    Dim folders As Collection, levels As Collection
    Dim subfoldersColl As Collection
    Dim n As Long, c As Long, i As Long
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set folders = New Collection
    Set levels = New Collection
    
    'Add start folder to stack
    
    folders.Add FSO.GetFolder(folderPath)
    levels.Add 0
       
    n = 1

    Do While folders.Count > 0
    
        'Remove next folder from top of stack
        
        Set FSfolder = folders(folders.Count): folders.Remove folders.Count
        c = levels(levels.Count): levels.Remove levels.Count
        
        'Output this folder and its files
        
'        destCell.Offset(n, c).Value = "'" & FSfolder.Name
'        n = n + 1
        c = c + 1
        For Each FSfile In FSfolder.Files
        On Error Resume Next
            destCell.Offset(n, 0).Value = FSfile.Name
            destCell.Offset(n, 1).Value = FSfile.Path
            destCell.Offset(n, 2).Value = FSfile.DateCreated
            destCell.Offset(n, 3).Value = FSfile.Datelastmodified
            destCell.Offset(n, 4).Value = FileLen(FSfile)
            n = n + 1
        Next
               
        'Get collection of subfolders in this folder
        
        Set subfoldersColl = New Collection
        For Each FSsubfolder In FSfolder.SubFolders
            subfoldersColl.Add FSsubfolder
        Next
        
        'Loop through collection in reverse order and put each subfolder on top of stack.  As a result, the subfolders are processed and
        'output in the correct ascending ASCII order
        
        For i = subfoldersColl.Count To 1 Step -1
            If folders.Count = 0 Then
                folders.Add subfoldersColl(i)
                levels.Add c
            Else
                folders.Add subfoldersColl(i), , , folders.Count
                levels.Add c, , , levels.Count
            End If
        Next
        Set subfoldersColl = Nothing
                
    Loop
    
    List_Folders_and_Files = n

End Function
