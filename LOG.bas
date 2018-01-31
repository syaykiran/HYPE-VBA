Attribute VB_Name = "LOG"
Option Explicit

Sub MCR_LOG(control As IRibbonControl)

Dim LOG_FILE As String
Dim LATEST_ONE As String
Dim LATEST_DATE As Date
Dim LOG_DATE As Date
Dim OBJ As Object
Dim MAIN_FOLDER As String
Dim INPUT_PATH As String
Dim INPUT_PATHS As String

INPUT_PATHS = ThisWorkbook.Path & "\INPUT\"

If Dir(INPUT_PATHS, vbDirectory) = "" Then
USF_LOAD_STARTUP.Show vbModal

Else
    MAIN_FOLDER = ThisWorkbook.Path
    INPUT_PATH = MAIN_FOLDER & "\" & "INPUT" & "\"
            
    LOG_FILE = Dir(INPUT_PATH & "*.log", vbNormal)
    
    If Len(LOG_FILE) = 0 Then
        MsgBox "No files were found...", vbExclamation
        Exit Sub
    End If
        
    Do While Len(LOG_FILE) > 0
        LOG_DATE = FileDateTime(INPUT_PATH & LOG_FILE)
        If LOG_DATE > LATEST_DATE Then
            LATEST_ONE = LOG_FILE
            LATEST_DATE = LOG_DATE
        End If
        LOG_FILE = Dir
    Loop
    
    Set OBJ = CreateObject("Shell.Application")
    OBJ.Open (INPUT_PATH & LATEST_ONE)
End If

'https://stackoverflow.com/questions/18921168/how-can-excel-vba-open-file-using-default-application
End Sub
