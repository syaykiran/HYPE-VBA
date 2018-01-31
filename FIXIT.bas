Attribute VB_Name = "FIXIT"
Option Explicit
Sub MCR_FIX(control As IRibbonControl)

Dim INPUT_PATHS As String

INPUT_PATHS = ThisWorkbook.Path & "\INPUT\"

If Dir(INPUT_PATHS, vbDirectory) = "" Then
USF_LOAD_STARTUP.Show vbModal

Else
    USF_FIX.Show vbModal
End If

End Sub


