Attribute VB_Name = "CHARTPDF"
Option Explicit
Sub MCR_CHART_PDF(control As IRibbonControl)

Dim OUTPUT_PATH As String
Dim PDF_FILE_NAME As String
Dim WRBK As Workbook
Dim WSHT_SERS, WSHT_CHRT, WSHT_LIST As Worksheet
Dim Ret
Dim INPUT_PATHS As String
Dim ZEROPAD As String

INPUT_PATHS = ThisWorkbook.Path & "\INPUT\"

If Dir(INPUT_PATHS, vbDirectory) = "" Then
USF_LOAD_STARTUP.Show vbModal

Else
    On Error GoTo 0
    
    With Application
        .EnableEvents = False
    '    .ScreenUpdating = True
        .DisplayAlerts = False
    '    .StatusBar = True
    End With
    
     
    Set WRBK = ThisWorkbook
    Set WSHT_CHRT = WRBK.Sheets("CHARTS")
        WSHT_CHRT.Visible = True
        WSHT_CHRT.Calculate
        
    Application.StatusBar = " >>  Exporting to PDF... (Please close PDF file if is already open!)"
    
    OUTPUT_PATH = ThisWorkbook.Path & "\OUTPUT"
    ZEROPAD = WorksheetFunction.Rept("0", 7 - Len(WSHT_CHRT.Range("CHART_SUBID"))) & WSHT_CHRT.Range("CHART_SUBID")
    PDF_FILE_NAME = OUTPUT_PATH & "\" & ZEROPAD & ".pdf"
     
'    Ret = IsFileOpen(PDF_FILE_NAME)
    
'    If Ret = True Then
'        MsgBox "PDF file if is already open! Please close it and try again!", vbExclamation, "HYPE VBA"
'        GoTo BYPASS
'    End If
    
            Application.StatusBar = " >>  Exporting to PDF ..."
            '.pdf export
            WSHT_CHRT.Activate
            WSHT_CHRT.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            FileName:=PDF_FILE_NAME, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=False, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=True
    
BYPASS:
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
        .DisplayAlerts = True
        .StatusBar = True
    End With
    
    WSHT_CHRT.Visible = True
End If
End Sub

Function IsFileOpen(FileName As String)
    Dim filenum As Integer, errnum As Integer

    On Error Resume Next   ' Turn error checking off.
    filenum = FreeFile()   ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open FileName For Input Lock Read As #filenum
    Close filenum          ' Close the file.
    errnum = Err           ' Save the error number that occurred.
    On Error GoTo 0        ' Turn error checking back on.

    ' Check to see which error occurred.
    Select Case errnum

        ' No error occurred.
        ' File is NOT already open by another user.
        Case 0
         IsFileOpen = False

        ' Error number for "Permission Denied."
        ' File is already opened by another user.
        Case 70
            IsFileOpen = True

        ' Another error occurred.
        Case Else
'            Error errnum
    End Select

End Function









