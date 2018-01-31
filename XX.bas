Attribute VB_Name = "XX"
Option Explicit

Sub CommentLookukjlq()
Dim commentRange As Range
Dim c As Range
Dim lookRange As Range
 Dim rFound As Range

On Error Resume Next
Worksheets("Par").Select
Set commentRange = Worksheets("Par").Range("A1:A138")
Set lookRange = Worksheets("COMMENT").Range("C2:H356")

Application.ScreenUpdating = False


For Each c In commentRange
    Set rFound = lookRange.Find(What:=c.Value, lookIn:=xlFormulas, Lookat:=xlWhole, SearchFormat:=False)
    
    If rFound Is Nothing Then
        c.ClearComments
          
    ElseIf IsEmpty(c.Value) = True Then
        c.ClearComments
    
    ElseIf c.Value = " " Then
        c.ClearComments
    
    Else
        With c
        .ClearComments
        .AddComment
        .COMMENT.Text Text:=CStr(WorksheetFunction.VLookup(c.Value, lookRange, 6, False))
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
'       .Comment.Shape.TextFrame.AutoSize = True
        .COMMENT.Shape.AutoShapeType = msoShapeRoundedRectangle
        .COMMENT.Shape.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .COMMENT.Shape.Fill.Transparency = 0
        .COMMENT.Visible = False
        End With
    End If
Next c

Application.ScreenUpdating = True
End Sub

'http://www.vbaexpress.com/forum/showthread.php?34005-How-to-wrap-comment-text-in-vba

Sub RemoveIndicatorShapes()
'www.contextures.com/xlcomments03.html

Dim ws As Worksheet
Dim shp As Shape

Set ws = ActiveSheet

For Each shp In ws.Shapes
If Not shp.TopLeftCell.COMMENT Is Nothing Then
  If Left(shp.NAME, 6) = "CmtNum" Then
    shp.Delete
  End If
End If
Next shp

End Sub

'Private Sub Worksheet_Change()
'Dim LOOK_RANGE As Range
'Dim COMMENT_RANGE_1, COMMENT_RANGE_2, COMMENT_RANGE As Range
'Dim FOUNDED As Range
'Dim sCellVal As String
'
'
'On Error Resume Next
'Worksheets("Info").Select
'
'sCellVal = Range("RNG_INFO").Value
'If sCellVal Like "*variable*" Then
'Set COMMENT_RANGE_1 = Range("sCellVal").xlToRight
'
'End If
'
'Set COMMENT_RANGE_2 = Worksheets("Info").Range("RNG_INFO")
'
'Set COMMENT_RANGE = Application.Union(COMMENT_RANGE_1, COMMENT_RANGE_2)
'Set LOOK_RANGE = Worksheets("COMMENT").Range("C299:H614")
'
'If Intersect(Target, Range("RNG_INFO")) Is Nothing Or Target.Cells.COUNT > 1 Then Exit Sub
'
'    If IsEmpty(Target.Value) = True Then
'        Target.COMMENT.Delete
'    ElseIf Target.Value = " " Then
'        Target.COMMENT.Delete
'    Else
'        Set FOUNDED = LOOK_RANGE.Find(What:=Target.Value, lookIn:=xlFormulas, Lookat:=xlWhole, SearchFormat:=False)
'
'       If FOUNDED Is Nothing Then
'            Target.ClearComments
'
'        Else
'            With Target
'            .ClearComments
'            .AddComment
'            .COMMENT.Text Text:=CStr(WorksheetFunction.VLookup(Target.Value, LOOK_RANGE, 6, False))
'            .COMMENT.Shape.TextFrame.Characters.Font.ColorIndex = 51
'            .COMMENT.Shape.TextFrame.Characters.Font.Size = 8
'            .COMMENT.Shape.TextFrame.Characters.Font.NAME = "Consola"
'                If Len(.COMMENT.Text) > 100 Then
'                    .COMMENT.Shape.Width = 250
'                    .COMMENT.Shape.Height = 50
'                ElseIf Len(.COMMENT.Text) > 50 Then
'                    .COMMENT.Shape.Width = 120
'                    .COMMENT.Shape.Height = 50
'                Else
'                    .COMMENT.Shape.Width = 110
'                    .COMMENT.Shape.Height = 40
'                End If
'            .COMMENT.Shape.TextFrame.HorizontalAlignment = xlHAlignCenter
'            .COMMENT.Shape.TextFrame.VerticalAlignment = xlVAlignCenter
'            .COMMENT.Shape.AutoShapeType = msoShapeRoundedRectangle
'            .COMMENT.Shape.Fill.ForeColor.RGB = RGB(255, 255, 255)
'            .COMMENT.Shape.Fill.Transparency = 0
'            .COMMENT.Visible = False
'            End With
'        End If
'    End If
'
'End Sub
'
'
'
