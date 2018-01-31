Attribute VB_Name = "XXGET_FILES"
Option Explicit

Sub GET_FILES()

Dim objFSO As Object
Dim objFolder As Object
Dim objFile As Object
Dim ws As Worksheet
Dim LR As Long, i As Long
Dim LResult As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False


    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set ws = Sheets("LIST")
     
     'Get the folder object associated with the directory
    Set objFolder = objFSO.GetFolder(ThisWorkbook.Sheets("LIST").Range("D1"))
    ws.Cells(1, 1).Value = objFolder.NAME
     
     'Loop through the Files collection
    For Each objFile In objFolder.Files
        ws.Cells(ws.UsedRange.Rows.COUNT + 1, 1).Value = objFile.NAME
    Next

     'Clean up!
    Set objFolder = Nothing
    Set objFile = Nothing
    Set objFSO = Nothing
    

LR = Range("A" & Rows.COUNT).End(xlUp).Row
For i = LR To 2 Step -1
    If Not Range("A" & i).Value Like "*.txt" Then Rows(i).Delete
     
Next i
'Range(Range("A2"), Range("A2").End(xlDown)).NumberFormat = "@"

Range(Range("A2"), Range("A2").End(xlDown)).Replace What:=".txt", Replacement:="", Lookat:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

For i = LR To 2 Step -1
    If IsNumeric(Range("A" & i).Value) = False Then Rows(i).Delete

Next i

Range("A" & 2).Value = WorksheetFunction.Rept("0", 7 - Len(Range("A" & 2))) & Range("A" & 2)

                     
Application.ScreenUpdating = True
Application.DisplayAlerts = True




End Sub

