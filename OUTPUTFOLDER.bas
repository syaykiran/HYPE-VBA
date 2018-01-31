Attribute VB_Name = "OUTPUTFOLDER"
Option Explicit
Sub MCR_OUTPUT(control As IRibbonControl)

Dim MAIN_FOLDER As String
Dim OUTPUT_PATH As String
Dim INPUT_PATHS As String

INPUT_PATHS = ThisWorkbook.Path & "\INPUT\"

If Dir(INPUT_PATHS, vbDirectory) = "" Then
USF_LOAD_STARTUP.Show vbModal
Else
    MAIN_FOLDER = ThisWorkbook.Path
    OUTPUT_PATH = MAIN_FOLDER & "\" & "OUTPUT"
    Call Shell("explorer.exe " & OUTPUT_PATH, vbNormalFocus)
End If

End Sub

' s.yaykýran 10/11/2016


