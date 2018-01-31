Attribute VB_Name = "XXTEXT_TO_GRAF"
'Option Explicit
'Sub MCR_CHART_PDF_BATCH()
'
'Dim OUTPUT_FILE_NAME As String
'Dim OUTPUT_PATH As String
'Dim PDF_FILE_NAME As String
'Dim OUTPUT_FILE As Variant
'Dim OBS_DATA As Range
'Dim SORC_DAT_DATA, SORC_SIM_DATA, SORC_OBS_DATA, DEST_OBX_DATA, PDF_ARE As Range
'Dim DEST_DAT_DATA, DEST_SIM_DATA, DEST_OBS_DATA, DEST_AVS_DATA, DEST_AVO_DATA As Range
'
'Dim subbasin() As String
'Dim element As Variant
'Dim hypefile As String
'Dim ARRAYS() As String, CELL As Range, i As Integer
'Dim listFile As Range
'Dim lastdatenum As Variant
'Dim lastdate As Variant
'Dim LAST_SATIRAI As Variant
'Dim Ret
'
'Dim WRBK As Workbook
'Dim WSHT_SERS, WSHT_CHRT, WSHT_LIST, WSHT_SYST  As Worksheet
'
'
'
'On Error GoTo 0
'
'With Application
'    .EnableEvents = False
''    .ScreenUpdating = False
'    .DisplayAlerts = False
'    .StatusBar = False
'    .Calculation = xlCalculationManual
'End With
'
'USF_GETWAIT.Show vbModeless
'DoEvents
'
'Call GET_FILES
'
'
'Application.Calculation = xlCalculationAutomatic
'ThisWorkbook.Sheets("LIST").Activate
'Set WRBK = ThisWorkbook
'Set WSHT_SERS = WRBK.Sheets("SERIES")
'Set WSHT_CHRT = WRBK.Sheets("CHARTS")
'Set WSHT_LIST = WRBK.Sheets("LIST")
'Set WSHT_SYST = WRBK.Sheets("SYSTEM")
'
'Application.Calculation = xlCalculationManual
'
'
'i = 0
'ReDim ARRAYS(0)
'For Each CELL In WSHT_LIST.Range("A2", Range("A2").End(xlDown))
' If CELL Is Nothing Or CELL.Value = "" Then GoTo BYPASS
'    ARRAYS(i) = CELL
'    i = i + 1
'    ReDim Preserve ARRAYS(i)
'
'    OUTPUT_PATH = ThisWorkbook.Path & "\OUTPUT"
'    OUTPUT_FILE_NAME = ThisWorkbook.Sheets("LIST").Range("D1") & CELL & ".txt"
'    PDF_FILE_NAME = ThisWorkbook.Sheets("LIST").Range("D1") & CELL & ".pdf"
'
'    Application.StatusBar = " >>  Opening .txt file..."
'    Set OUTPUT_FILE = Workbooks.Open(OUTPUT_FILE_NAME)
'    Set SORC_DAT_DATA = OUTPUT_FILE.Sheets(1).Range("A3", Range("A3").End(xlDown))
'    Set SORC_SIM_DATA = OUTPUT_FILE.Sheets(1).Range("B3", Range("B3").End(xlDown))
'    Set SORC_OBS_DATA = OUTPUT_FILE.Sheets(1).Range("C3", Range("C3").End(xlDown))
'    LAST_SATIRAI = OUTPUT_FILE.Sheets(1).Range("A3").End(xlDown).Row
'
'    OUTPUT_FILE.Sheets(1).Range(Cells(3, 1), Cells(LAST_SATIRAI, 1)).NumberFormat = "dd/mm/yyyy"
'
'
'    WSHT_SERS.Activate
'    WSHT_SERS.Range("A:F").Clear
'    WSHT_SERS.Range(Cells(3, 1), Cells(LAST_SATIRAI, 1)).NumberFormat = "dd/mm/yyyy"
'    Set DEST_DAT_DATA = WSHT_SERS.Range(Cells(3, 1), Cells(LAST_SATIRAI, 1))
'    Set DEST_SIM_DATA = WSHT_SERS.Range(Cells(3, 2), Cells(LAST_SATIRAI, 2))
'    Set DEST_OBS_DATA = WSHT_SERS.Range(Cells(3, 3), Cells(LAST_SATIRAI, 3))
'    Set DEST_OBX_DATA = WSHT_SERS.Range(Cells(3, 4), Cells(LAST_SATIRAI, 4))
'    Set DEST_AVS_DATA = WSHT_SERS.Range(Cells(3, 5), Cells(LAST_SATIRAI, 5))
'    Set DEST_AVO_DATA = WSHT_SERS.Range(Cells(3, 6), Cells(LAST_SATIRAI, 6))
'
'
'    'text to GRAFIK
'    DEST_DAT_DATA.Value = SORC_DAT_DATA.Value
'    DEST_SIM_DATA.Value = SORC_SIM_DATA.Value
'    DEST_OBS_DATA.Value = SORC_OBS_DATA.Value
'    DEST_OBX_DATA.Value = SORC_OBS_DATA.Value
'
'    'missing value to space or blank
'    DEST_DAT_DATA.Replace What:="-9999", Replacement:=" "
'    DEST_SIM_DATA.Replace What:="-9999", Replacement:=" "
'    DEST_OBS_DATA.Replace What:="-9999", Replacement:=" "
'    DEST_OBX_DATA.Replace What:="-9999", Replacement:=""
'    DEST_AVS_DATA = Application.Average(WSHT_SERS.Range("XMODEL"))
'    DEST_AVO_DATA = Application.Average(WSHT_SERS.Range("XGOZLEM"))
'
'    If WorksheetFunction.CountA(WSHT_SERS.Range(Cells(3, 5), Cells(LAST_SATIRAI, 5))) = 0 Then Cells(3, 5).Formula = "=NA()"
'
'    'extra 1 ay daha tarih ekle, grafik sonu için
'    lastdatenum = WSHT_SERS.Range("A" & Rows.COUNT).End(xlUp).Row
'    lastdate = WSHT_SERS.Range("A" & lastdatenum).Value
'    WSHT_SERS.Range("A" & lastdatenum + 1).Value = DateAdd("m", 1, CDate(lastdate))
'
'    Application.StatusBar = " >>  Loading charts ..."
'    OUTPUT_FILE.Close (False)
'
'    WSHT_SYST.Calculate
'    WSHT_CHRT.Calculate
'
'    '.pdf çýktýsý alma
'    WSHT_LIST.Activate
'    WSHT_LIST.Range("ST_NO") = CELL
'
'    If Not Application.CalculationState = xlDone Then
'        DoEvents
'    End If
'
'    Ret = IsFileOpen(PDF_FILE_NAME)
'
'    If Ret = True Then
'        MsgBox "PDF file if is already open! Please close it and try again!", vbExclamation, "HYPE VBA"
'        GoTo BYPASS
'    End If
'
'    WSHT_CHRT.Activate
'
''    If Me.CHBX_EXPDF = True Then
'            Application.StatusBar = " >>  Exporting to PDF ..."
'            '.pdf çýktýsý alma
'            WSHT_CHRT.Activate
'            WSHT_CHRT.Range("CHART_SUBID") = CStr(CELL)
'            WSHT_CHRT.Range("PDF_AREA").ExportAsFixedFormat _
'            Type:=xlTypePDF, _
'            FileName:=PDF_FILE_NAME, _
'            Quality:=xlQualityStandard, _
'            IncludeDocProperties:=False, _
'            IgnorePrintAreas:=False, _
'            OpenAfterPublish:=True
''        End If
'
'BYPASS:
'
'    USF_GETWAIT.HIDE
'    Unload USF_GETCHARTS
'    Unload USF_GETWAIT
'    USF_GETFINISH.Show vbModeless
'
'    WSHT_CHRT.Activate
'
'    Unload USF_GETFINISH
'    Application.StatusBar = " >>  Finished!"
'End If
'
'With Application
'    .EnableEvents = True
'    .ScreenUpdating = True
'    .DisplayAlerts = True
'    .StatusBar = True
'    .Calculation = xlCalculationAutomatic
'End With
'
'Call Shell("explorer.exe " & OUTPUT_PATH, vbNormalFocus)
'
'End Sub
'
'
'
'
'
'
