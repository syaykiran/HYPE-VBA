Attribute VB_Name = "INACTIVE"
Option Explicit

Sub MCR_ACTIVE_INACTIVE(control As IRibbonControl)
Dim CELLO As Range

On Error Resume Next
Application.ScreenUpdating = False

If Selection.Cells.COUNT < 50 Then
    For Each CELLO In Selection
        CELLO = IIf(Left(CELLO, 2) = "!!", Right(CELLO, Len(CELLO) - 2), IIf(Left(CELLO, 1) = "!", Right(CELLO, Len(CELLO) - 1), "!!" & CELLO))
    Next
Else
    Application.StatusBar = "Too many cells have been selected!"
    Application.Wait Now + TimeValue("00:00:01")
End If

Application.StatusBar = True
Application.ScreenUpdating = True
End Sub

