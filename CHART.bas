Attribute VB_Name = "CHART"
Sub MCR_CHART(control As IRibbonControl)

Dim INPUT_PATHS As String

INPUT_PATHS = ThisWorkbook.Path & "\INPUT\"

If Dir(INPUT_PATHS, vbDirectory) = "" Then
USF_LOAD_STARTUP.Show vbModal

Else
    USF_GETCHARTS.Show vbModal
End If

End Sub
