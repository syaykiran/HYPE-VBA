VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WS_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
Application.EnableEvents = True
Application.StatusBar = " >> Info.txt  :  Model options and simulation settings (mandatory)."
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
Dim LOOK_RANGE As Range
Dim COMMENT_RANGE As Range
Dim FOUNDED As Range

On Error Resume Next

'Dim Rng1, Rng2, Rng3 As Range
'COMMENT_RANGE_1, COMMENT_RANGE_2, COMMENT_RANGE_3, COMMENT_RANGE_4,
'With ActiveSheet
'    Set Rng1 = .Range("RNG_INFO").Find(what:="basinoutput variable", _
'                lookIn:=xlValues, _
'                Lookat:=xlWhole, _
'                Searchorder:=xlByRows, _
'                Searchdirection:=xlNext)
'End With
'
'If Rng1 Is Nothing Then
'Set COMMENT_RANGE_1 = Worksheets("Info").Range("A1")
'Else
'Set COMMENT_RANGE_1 = Worksheets("Info").Range(Range(Rng1.Address), Range(Rng1.Address).End(xlToRight))
'End If
'
'
'With ActiveSheet
'    Set Rng2 = .Range("RNG_INFO").Find(what:="mapoutput variable", _
'                lookIn:=xlValues, _
'                Lookat:=xlWhole, _
'                Searchorder:=xlByRows, _
'                Searchdirection:=xlNext)
'End With
'
'If Rng2 Is Nothing Then
'Set COMMENT_RANGE_2 = Worksheets("Info").Range("A1")
'Else
'Set COMMENT_RANGE_2 = Worksheets("Info").Range(Range(Rng2.Address), Range(Rng2.Address).End(xlToRight))
'End If
'
'
'With ActiveSheet
'    Set Rng3 = .Range("RNG_INFO").Find(what:="timeoutput variable", _
'                lookIn:=xlValues, _
'                Lookat:=xlWhole, _
'                Searchorder:=xlByRows, _
'                Searchdirection:=xlNext)
'End With
'If Rng3 Is Nothing Then
'Set COMMENT_RANGE_3 = Worksheets("Info").Range("A1")
'Else
'Set COMMENT_RANGE_3 = Worksheets("Info").Range(Range(Rng3.Address), Range(Rng3.Address).End(xlToRight))
'End If
'
'
'Set COMMENT_RANGE_4 = Worksheets("Info").Range("RNG_INFO")
'Set COMMENT_RANGE = Application.Union(COMMENT_RANGE_1, COMMENT_RANGE_2, COMMENT_RANGE_3, COMMENT_RANGE_4)

Set COMMENT_RANGE = ThisWorkbook.Worksheets("Info").Range("RNG_INFO")
Set LOOK_RANGE = ThisWorkbook.Worksheets("COMMENT").Range("DESC_INFO")

If Intersect(Target, COMMENT_RANGE) Is Nothing Or Target.Cells.COUNT > 1 Then Exit Sub
    
    If Target.Value = "bdate" Then
    Target.Offset(0, 1).NumberFormat = "yyyy-mm-dd"
    
    End If
    
    If IsEmpty(Target.Value) = True Or Target.Value = " " Then
        Target.COMMENT.Delete
    Else
        Set FOUNDED = LOOK_RANGE.Find(What:=Target.Value, lookIn:=xlValues, Lookat:=xlWhole, SearchFormat:=False)
    
        If FOUNDED Is Nothing Then
            Target.ClearComments
        
            Else
            With Target
                    .ClearComments
                    .AddComment
                    .COMMENT.Text Text:=CStr(WorksheetFunction.VLookup(Target.Value, LOOK_RANGE, 2, False))
            .COMMENT.Shape.TextFrame.Characters.Font.ColorIndex = 51
            .COMMENT.Shape.TextFrame.Characters.Font.Size = 8
            .COMMENT.Shape.TextFrame.Characters.Font.NAME = "Consola"
           
           If Len(.COMMENT.Text) > 350 Then
               .COMMENT.Shape.Width = 350
               .COMMENT.Shape.Height = 80
               
           ElseIf Len(.COMMENT.Text) > 225 Then
               .COMMENT.Shape.Width = 250
               .COMMENT.Shape.Height = 50
           
           ElseIf Len(.COMMENT.Text) > 100 Then
               .COMMENT.Shape.Width = 150
               .COMMENT.Shape.Height = 50
               
           ElseIf Len(.COMMENT.Text) > 50 Then
               .COMMENT.Shape.Width = 120
               .COMMENT.Shape.Height = 50
           
           ElseIf Len(.COMMENT.Text) > 25 Then
               .COMMENT.Shape.Width = 110
               .COMMENT.Shape.Height = 30
           
           Else
               .COMMENT.Shape.Width = 110
               .COMMENT.Shape.Height = 40
           End If
                           
            .COMMENT.Shape.TextFrame.HorizontalAlignment = xlHAlignCenter
            .COMMENT.Shape.TextFrame.VerticalAlignment = xlVAlignCenter
            .COMMENT.Shape.AutoShapeType = msoShapeRoundedRectangle
            .COMMENT.Shape.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .COMMENT.Shape.Fill.Transparency = 0
            .COMMENT.Visible = False
            End With
        End If
    End If

End Sub

Private Sub maino()

'replace "J2" with the cell you want to insert the drop down list
With ThisWorkbook.Worksheets("Info").Range("RNG_INFO").Validation
.Delete
'replace "=A1:A6" with the range the data is in.
.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
Operator:=xlBetween, Formula1:="=COMMENT!$B$257:$B$679"
.IgnoreBlank = True
.InCellDropdown = True
.InputTitle = ""
.ErrorTitle = ""
.InputMessage = ""
.ErrorMessage = ""
.ShowInput = True
.ShowError = False
End With
End Sub





