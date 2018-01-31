Attribute VB_Name = "BACKUP"
Option Explicit
Sub MCR_BACKUP(control As IRibbonControl)

Dim INPUT_PATHS As String

INPUT_PATHS = ThisWorkbook.Path & "\INPUT\"

If Dir(INPUT_PATHS, vbDirectory) = "" Then
USF_LOAD_STARTUP.Show vbModal

Else
    USF_BACKUP.Show vbModal
End If

End Sub
