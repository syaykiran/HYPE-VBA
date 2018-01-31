VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WS_Par"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub Worksheet_Activate()
Application.EnableEvents = True
Application.StatusBar = " >> Par.txt  :  Model parameters, some is calibrated (mandatory)."
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
Dim COMMENT_RANGE As Range
Dim LOOK_RANGE As Range
Dim FOUNDED As Range

On Error Resume Next

Set COMMENT_RANGE = Range("RNG_PAR")
Set LOOK_RANGE = ThisWorkbook.Worksheets("COMMENT").Range("DESC_PAR")

If Intersect(Target, COMMENT_RANGE) Is Nothing Or Target.Cells.COUNT > 1 Then Exit Sub

    If IsEmpty(Target.Value) = True Or Target.Value = " " Then
        Target.COMMENT.Delete
    Else
        Set FOUNDED = LOOK_RANGE.Find(What:=Target.Value, lookIn:=xlValues, Lookat:=xlWhole, SearchFormat:=False)
    
        If FOUNDED Is Nothing Then
            Target.COMMENT.Delete
        
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
    
     Dim r As Range, c As Range
    Set r = Intersect(Target, Range("a1:a937"))
            For Each c In r
                c.Comments.Delete
        Next c
    

    
    
End Sub


Sub Worksheet_Changesss()

  Dim KOLON_ON As Variant
  Dim SATIR_YIRMI As Variant

KOLON_ON = Worksheets("Par").UsedRange.SpecialCells(xlCellTypeLastCell).Column + 10
SATIR_YIRMI = Worksheets("Par").UsedRange.SpecialCells(xlCellTypeLastCell).Row + 20


    Worksheets("Par").Cells.FormatConditions.Delete
    
    
    With Worksheets("Par").Cells(1, 1).FormatConditions.Add(Type:=xlExpression, Formula1:="=ISNUMBER(FIND(""!"",$A1))")
    .Font.Italic = True
    .Font.Bold = True
    .Font.Color = -16744448
    .StopIfTrue = False
    End With
    
    With Cells(1, 1).FormatConditions.Add(Type:=xlExpression, Formula1:="=MOD(row(),2)=0")
    .Interior.Color = RGB(240, 240, 240)
    .StopIfTrue = False
    End With
    
    Cells.FormatConditions(1).ModifyAppliesToRange Worksheets("Par").Range(Cells(1, 1), Cells(SATIR_YIRMI, KOLON_ON))
    Cells.FormatConditions(2).ModifyAppliesToRange Worksheets("Par").Range(Cells(1, 1), Cells(SATIR_YIRMI, KOLON_ON))
    Worksheets("Par").UsedRange.Borders.LineStyle = xlNone
    ActiveWindow.DisplayGridlines = False





'   Application.ScreenUpdating = True
'   Application.EnableEvents = True
End Sub

Private Sub main()

'replace "J2" with the cell you want to insert the drop down list
With ThisWorkbook.Worksheets("Par").Range("RNG_PAR").Validation
.Delete
'replace "=A1:A6" with the range the data is in.
.Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, _
Operator:=xlBetween, Formula1:="=COMMENT!$B$2:$B$256"
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


