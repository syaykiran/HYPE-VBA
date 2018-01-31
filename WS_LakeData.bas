VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WS_LakeData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
Application.EnableEvents = True
Application.StatusBar = " >> LakeData.txt  :  Properties of specific lakes, including regulated dams (optional)."
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
Dim COMMENT_RANGE As Range
Dim LOOK_RANGE As Range
Dim FOUNDED As Range

On Error Resume Next

Set COMMENT_RANGE = Range("RNG_LAKEDATA")
Set LOOK_RANGE = ThisWorkbook.Worksheets("COMMENT").Range("DESC_LAKEDATA")

If Intersect(Target, COMMENT_RANGE) Is Nothing Or Target.Cells.COUNT > 1 Then Exit Sub

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
                        If Len(.COMMENT.Text) > 100 Then
                            .COMMENT.Shape.Width = 250
                            .COMMENT.Shape.Height = 50
                        ElseIf Len(.COMMENT.Text) > 50 Then
                            .COMMENT.Shape.Width = 120
                            .COMMENT.Shape.Height = 50
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


