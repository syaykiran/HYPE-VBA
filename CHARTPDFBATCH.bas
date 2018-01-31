Attribute VB_Name = "CHARTPDFBATCH"

Sub MCR_BATCH_EXPORT(control As IRibbonControl)

Dim INPUT_PATHS As String

INPUT_PATHS = ThisWorkbook.Path & "\INPUT\"

If Dir(INPUT_PATHS, vbDirectory) = "" Then
USF_LOAD_STARTUP.Show vbModal

Else
    USF_BATCH_EXPORT.Show vbModal
End If

End Sub

