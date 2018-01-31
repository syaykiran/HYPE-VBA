Attribute VB_Name = "DATECONVERT"
Option Explicit
Sub MCR_DATE_CONVERT(control As IRibbonControl)
Dim CELLO As Range

On Error Resume Next
Application.ScreenUpdating = False
'If Selection.Cells.COUNT < 50 Then
    For Each CELLO In Selection
    If Not CELLO.NumberFormat = "yyyymmdd" Then
    CELLO.NumberFormat = "yyyymmdd"
    Else
    CELLO.NumberFormat = "yyyy-mm-dd"
    End If
    Next
'Else
'    Application.StatusBar = "Too many cells have been selected!"
'    Application.Wait Now + TimeValue("00:00:01")
'End If

Application.StatusBar = True
Application.ScreenUpdating = True
End Sub



