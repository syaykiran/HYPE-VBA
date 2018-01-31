Attribute VB_Name = "INPUTFOLDER"
Option Explicit
Sub MCR_INPUT(control As IRibbonControl)

Dim MAIN_FOLDER As String
Dim INPUT_PATH As String
Dim INPUT_PATHS As String

INPUT_PATHS = ThisWorkbook.Path & "\INPUT\"

If Dir(INPUT_PATHS, vbDirectory) = "" Then
USF_LOAD_STARTUP.Show vbModal

Else
    MAIN_FOLDER = ThisWorkbook.Path
    INPUT_PATH = MAIN_FOLDER & "\" & "INPUT"
    Call Shell("explorer.exe " & INPUT_PATH, vbNormalFocus)
End If

End Sub

' s.yaykýran 10/11/2016

