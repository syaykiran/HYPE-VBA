Attribute VB_Name = "RUN"
Option Explicit

Sub MCR_RUN(control As IRibbonControl)

Dim MAIN_FOLDER As String
Dim INPUT_PATH As Variant
Dim MODEL_NAME As Variant
Dim EXE_PATH As Variant
Dim objShell As Object
Dim INPUT_PATHS As String

INPUT_PATHS = ThisWorkbook.Path & "\INPUT\"

If Dir(INPUT_PATHS, vbDirectory) = "" Then
USF_LOAD_STARTUP.Show vbModal
Else
    MAIN_FOLDER = ThisWorkbook.Path
    INPUT_PATH = MAIN_FOLDER & "\" & "INPUT" & "\"
    
    MODEL_NAME = "HYPE"
    EXE_PATH = INPUT_PATH & MODEL_NAME
       
    Set objShell = CreateObject("Shell.Application")
    objShell.ShellExecute EXE_PATH, , INPUT_PATH, , 10     'önemli! yönetici olarak çalýþtýrýr
End If
' s.yaykýran 10/11/2016
End Sub
