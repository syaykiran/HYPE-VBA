VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USF_BATCH_EXPORT 
   Caption         =   "Batch Export to PDF"
   ClientHeight    =   9780.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8295.001
   OleObjectBlob   =   "USF_BATCH_EXPORT.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "USF_BATCH_EXPORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

Dim ws1, ws2, ws3 As Worksheet
Dim cPart As Range
Dim ROW_NO, ROW_NO2, SAYI, SAYI2 As Long
Dim BASIN_RAN, BASIN_RAN2 As String

'**********************

'Set ws1 = ThisWorkbook.Worksheets("LIST")
Set ws2 = ThisWorkbook.Worksheets("Info")

With Application
    .EnableEvents = False
    .ScreenUpdating = False
    .DisplayAlerts = False
'   .StatusBar = False
    .Calculation = xlCalculationManual
End With

ws2.Activate

ROW_NO = ws2.Columns(1). _
Find(What:="basinoutput subbasin", SearchOrder:=xlByRows, searchDirection:=xlPrevious).Row

SAYI = ws2.Range(Cells(ROW_NO, 2), Cells(ROW_NO, 2).End(xlToRight)).COUNT
Select Case SAYI
Case Is > 500
    BASIN_RAN = ws2.Range(Cells(ROW_NO, 2), Cells(ROW_NO, 2)).Address

Case 1
    BASIN_RAN = ws2.Range(Cells(ROW_NO, 2), Cells(ROW_NO, 2)).Address

Case Else
    BASIN_RAN = ws2.Range(Cells(ROW_NO, 2), Cells(ROW_NO, 2).End(xlToRight)).Address

End Select

ThisWorkbook.Names("CHR_BAS_SUB").RefersTo = "=Info!" & BASIN_RAN


For Each cPart In ws2.Range("CHR_BAS_SUB")
    Me.LSTBX_SUBID.AddItem cPart.Value
Next cPart




ROW_NO2 = ws2.Columns(1). _
Find(What:="basinoutput variable", SearchOrder:=xlByRows, searchDirection:=xlPrevious).Row

SAYI2 = ws2.Range(Cells(ROW_NO2, 2), Cells(ROW_NO2, 2).End(xlToRight)).COUNT

Select Case SAYI2
Case Is > 500
    BASIN_RAN2 = ws2.Range(Cells(ROW_NO2, 2), Cells(ROW_NO2, 2)).Address

Case 1
    BASIN_RAN2 = ws2.Range(Cells(ROW_NO2, 2), Cells(ROW_NO2, 2)).Address

Case Else
    BASIN_RAN2 = ws2.Range(Cells(ROW_NO2, 2), Cells(ROW_NO2, 2).End(xlToRight)).Address

End Select


ThisWorkbook.Names("CHR_BAS_VAR").RefersTo = "=Info!" & BASIN_RAN2
   
For Each cPart In ws2.Range("CHR_BAS_VAR")
    Me.CB_OBSPAR.AddItem cPart.Value
Next cPart

For Each cPart In ws2.Range("CHR_BAS_VAR")
    Me.CB_SIMPAR.AddItem cPart.Value
Next cPart


'    .List(.ListCount - 1, 1) = cPart.Offset(0, 1).Value

'With Application
'    .EnableEvents = True
'    .ScreenUpdating = True
'    .DisplayAlerts = True
'    .StatusBar = True
''    .Calculation = xlCalculationAutomatic
'End With


End Sub

Private Sub CBUT_GETCHART_Click()

Dim OUTPUT_FILE_NAME As String
Dim OUTPUT_PATH As String
Dim PDF_FILE_NAME As String
Dim OUTPUT_FILE As Variant
Dim OBS_DATA As Range
Dim SORC_DAT_DATA, SORC_SIM_DATA, SORC_OBS_DATA, DEST_OBX_DATA, PDF_ARE As Range
Dim DEST_DAT_DATA, DEST_SIM_DATA, DEST_OBS_DATA, DEST_AVS_DATA, DEST_AVO_DATA As Range
Dim FIRST_CALC_YEAR, LAST_CALC_YEAR, DIFF_CALC As Integer
Dim DEST_SIM_LABEL, DEST_OBS_LABEL, DEST_SIM_UNIT, DEST_OBS_UNIT As Range

Dim CELL As Variant
Dim listFile, CALC_RANGE As Range
Dim lastdatenum As Variant
Dim lastdate As Variant
Dim LAST_SATIRAI As Variant
Dim Ret
Dim SLC_SIM_COL As Variant
Dim SLC_OBS_COL As Variant
Dim i, s, w As Integer

On Error GoTo 0
    
Dim WRBK As Workbook
Dim WSHT_SERS, WSHT_CHRT, WSHT_LIST, WSHT_SYST, WHST_INFO As Worksheet



Set WRBK = ThisWorkbook
Set WSHT_SERS = WRBK.Sheets("SERIES")
Set WSHT_CHRT = WRBK.Sheets("CHARTS")
Set WSHT_SYST = WRBK.Sheets("SYSTEM")

With Application
    .EnableEvents = False
    .ScreenUpdating = False
    .DisplayAlerts = False
'   .StatusBar = False
End With

    WSHT_SERS.Visible = True
    WSHT_SYST.Visible = True
    WSHT_CHRT.Visible = True
    
'If Dir(OUTPUT_FILE_NAME, vbDirectory) = "" Then
''    MsgBox "File doesn't exist in the output folder. First run the model please!" & vbCrLf, vbExclamation, "HYPE VBA"
''    Unload USF_BATCH_EXPORT
'Else
'If CELL Is Nothing Or CELL.Value = "" Then GoTo BYPASS

    ' SELECTED COUNT
    With Me.LSTBX_SUBID
        s = -1
        For i = 0 To .ListCount - 1
            If .Selected(i) = True Then s = s + 1
        Next
        w = -1
    End With
    
    'PROCESS
    
    For i = 0 To Me.LSTBX_SUBID.ListCount - 1
        If Me.LSTBX_SUBID.Selected(i) = True Then
        CELL = Me.LSTBX_SUBID.LIST(i)
         
        CELL = WorksheetFunction.Rept("0", 7 - Len(CELL)) & CELL
        w = w + 1
        
        OUTPUT_PATH = ThisWorkbook.Path & "\OUTPUT"
        OUTPUT_FILE_NAME = OUTPUT_PATH & "\" & CELL & ".txt"
        PDF_FILE_NAME = OUTPUT_PATH & "\" & CELL & ".pdf"
        
        If Dir(OUTPUT_FILE_NAME, vbDirectory) = "" Then GoTo BYPASS

                USF_BATCH_EXPORT.HIDE
                USF_GETWAIT.Show vbModeless
                DoEvents
            '    ThisWorkbook.Sheets("LIST").Activate
                
                Set WRBK = ThisWorkbook
                Set WSHT_SERS = WRBK.Sheets("SERIES")
                Set WSHT_CHRT = WRBK.Sheets("CHARTS")
                Set WHST_INFO = WRBK.Sheets("Info")
                Set WSHT_SYST = WRBK.Sheets("SYSTEM")
                
                Application.Calculation = xlCalculationManual
                
                SLC_SIM_COL = Application.WorksheetFunction.Match(Me.CB_SIMPAR, WHST_INFO.Range("CHR_BAS_VAR"), 0) + 1
                SLC_OBS_COL = Application.WorksheetFunction.Match(Me.CB_OBSPAR, WHST_INFO.Range("CHR_BAS_VAR"), 0) + 1
                
            
                        Debug.Print SLC_SIM_COL
                        Debug.Print SLC_OBS_COL
            
                Application.StatusBar = "(" & w + 1 & "/" & s + 1 & ")   " & ">>  Opening .txt file..."
                Set OUTPUT_FILE = Workbooks.Open(OUTPUT_FILE_NAME)
                
                If OUTPUT_FILE.Sheets(1).Range("A3").Value = "" Then
                    MsgBox "File doesn't exist in the output folder. Firstly, run the model properly please!" & vbCrLf, vbExclamation, "HYPE VBA"
                    OUTPUT_FILE.Close (False)
                    GoTo BYPASS
                End If
                
                    Set SORC_DAT_DATA = OUTPUT_FILE.Sheets(1).Range("A3", Range("A3").End(xlDown))
                    LAST_SATIRAI = OUTPUT_FILE.Sheets(1).Range("A3").End(xlDown).Row
                    Set SORC_SIM_DATA = OUTPUT_FILE.Sheets(1).Range(Cells(3, SLC_SIM_COL), Cells(LAST_SATIRAI, SLC_SIM_COL))
                    Set SORC_OBS_DATA = OUTPUT_FILE.Sheets(1).Range(Cells(3, SLC_OBS_COL), Cells(LAST_SATIRAI, SLC_OBS_COL))
                    
                    Set DEST_SIM_LABEL = OUTPUT_FILE.Sheets(1).Cells(1, SLC_SIM_COL)
                    Set DEST_OBS_LABEL = OUTPUT_FILE.Sheets(1).Cells(1, SLC_OBS_COL)
                    Set DEST_SIM_UNIT = OUTPUT_FILE.Sheets(1).Cells(2, SLC_SIM_COL)
                    Set DEST_OBS_UNIT = OUTPUT_FILE.Sheets(1).Cells(2, SLC_OBS_COL)
                
                            Debug.Print DEST_SIM_LABEL
                            Debug.Print DEST_OBS_LABEL
                            Debug.Print DEST_SIM_UNIT
                            Debug.Print DEST_OBS_UNIT
                            '    FIRST_CALC_YEAR = Year(OUTPUT_FILE.Sheets(1).Range("A3").Value)
                            '    LAST_CALC_YEAR = Year(OUTPUT_FILE.Sheets(1).Cells(LAST_SATIRAI, 1).Value)
                            '    DIFF_CALC = LAST_CALC_YEAR - FIRST_CALC_YEAR + 4
                            '
                            '    WSHT_SYST.Activate
                            '    Set CALC_RANGE = Union(WSHT_SYST.Range(Cells(1, 1), Cells(DIFF_CALC, 88)), WSHT_SYST.Range("ADD_CALC_RNG"))
                            '
                            '    Debug.Print FIRST_CALC_YEAR
                            '    Debug.Print LAST_CALC_YEAR
                            '
                    WSHT_SERS.Activate
                    WSHT_SERS.Range("A:G").Clear
                    
                    Set DEST_DAT_DATA = WSHT_SERS.Range(Cells(3, 1), Cells(LAST_SATIRAI, 1))
                    Set DEST_SIM_DATA = WSHT_SERS.Range(Cells(3, 2), Cells(LAST_SATIRAI, 2))
                    Set DEST_OBS_DATA = WSHT_SERS.Range(Cells(3, 3), Cells(LAST_SATIRAI, 3))
                    Set DEST_OBX_DATA = WSHT_SERS.Range(Cells(3, 4), Cells(LAST_SATIRAI, 4))
                    Set DEST_AVS_DATA = WSHT_SERS.Range(Cells(3, 5), Cells(LAST_SATIRAI, 5))
                    Set DEST_AVO_DATA = WSHT_SERS.Range(Cells(3, 6), Cells(LAST_SATIRAI, 6))
                    
                    
                    '    'tarih formatý ayarlama
                    '    With ActiveSheet.UsedRange.Columns("A").Cells
                    '       .TextToColumns Destination:=.Cells(1), DataType:=xlFixedWidth, FieldInfo:=Array(0, xlYMDFormat)
                    '       .NumberFormat = "yyyy/mm/dd"
                    '    End With
                    
                    'text to GRAFIK
                    Application.StatusBar = "(" & w + 1 & "/" & s + 1 & ")   " & ">>  Copying data from .txt file..."
                    DEST_DAT_DATA.Value = SORC_DAT_DATA.Value
                    DEST_SIM_DATA.Value = SORC_SIM_DATA.Value
                    DEST_OBS_DATA.Value = SORC_OBS_DATA.Value
                    DEST_OBX_DATA.Value = SORC_OBS_DATA.Value
                    
                    WSHT_SYST.Range("CW2").Value = DEST_SIM_LABEL
                    WSHT_SYST.Range("CW3").Value = DEST_OBS_LABEL
                    WSHT_SYST.Range("CY2").Value = DEST_SIM_UNIT
                    WSHT_SYST.Range("CY3").Value = DEST_OBS_UNIT
                    
                
                    
                    'missing value to space or blank
                    Application.StatusBar = " >>  Editing chart data ..."
                    DEST_DAT_DATA.Replace What:="-9999", Replacement:=" "
                    DEST_SIM_DATA.Replace What:="-9999", Replacement:=" "
                    DEST_OBS_DATA.Replace What:="-9999", Replacement:=" "
                    DEST_OBX_DATA.Replace What:="-9999", Replacement:=""
                    
                    WSHT_SERS.Calculate
                    
                    DEST_AVO_DATA.Value = WSHT_SERS.Range("AV_OBS").Value
                    DEST_AVS_DATA.Value = WSHT_SERS.Range("AV_SIM").Value
                    
                    If WorksheetFunction.CountA(WSHT_SERS.Range(Cells(3, 5), Cells(LAST_SATIRAI, 5))) = 0 Then Cells(3, 5).Formula = "=NA()"
                 
                    'extra 1 ay daha tarih ekle, grafik sonu için
                    lastdatenum = WSHT_SERS.Range("A" & Rows.COUNT).End(xlUp).Row
                    lastdate = WSHT_SERS.Range("A" & lastdatenum).Value
                    WSHT_SERS.Range("A" & lastdatenum + 1).Value = DateAdd("m", 1, CDate(lastdate))
                    Application.StatusBar = "(" & w + 1 & "/" & s + 1 & ")   " & ">>  Loading charts ..."
                    OUTPUT_FILE.Close (False)
                    
                    
                '    CALC_RANGE.Calculate
                    WSHT_SYST.Calculate
'                    If Not Application.CalculationState = xlDone Then
'                        DoEvents
'                    End If
                    
'                    If Not Application.CalculationState = xlDone Then
'                        DoEvents
'                    End If
                    
'                    If Not Application.CalculationState = xlDone Then
'                        DoEvents
'                    End If
                                    
                    Ret = IsFileOpen(PDF_FILE_NAME)
                    
                    If Ret = True Then
                        MsgBox "PDF file if is already open! Please close it and try again!", vbExclamation, "HYPE VBA"
                        GoTo BYPASS
                    End If
                    
                    WSHT_CHRT.Activate
                    WSHT_CHRT.Range("CHART_SUBID") = CStr(CELL)
                    WSHT_CHRT.Calculate
                    
                    With WSHT_CHRT.PageSetup
                        .LeftFooter = "&""Cambria""&8&K00-034&B" & " SUB ID: " & "&B" & WSHT_CHRT.Range("CHART_SUBID") & Chr(10) & WSHT_CHRT.Range("HDR_STNAME") & " " & WSHT_CHRT.Range("HDR_STN0")
                        .RightFooter = "&""Cambria""&8&K00-034&D" & Chr(10) & "&T"
                        .CenterFooter = "&""Cambria""&8&K00-034&P/&N" & Chr(10)
                    End With
                            
                    
                        Application.StatusBar = "(" & w + 1 & "/" & s + 1 & ")   " & ">>  Exporting to PDF ..."
                        '.pdf export
                        WSHT_CHRT.Activate
                        WSHT_CHRT.Range("CHART_SUBID") = CStr(CELL)
                        WSHT_CHRT.ExportAsFixedFormat _
                        Type:=xlTypePDF, _
                        FileName:=PDF_FILE_NAME, _
                        Quality:=xlQualityStandard, _
                        IncludeDocProperties:=False, _
                        IgnorePrintAreas:=False, _
                        OpenAfterPublish:=True
        End If
BYPASS:
        Next i
        
    
'    End If
                USF_GETWAIT.HIDE
                Unload USF_BATCH_EXPORT
                Unload USF_GETWAIT
                USF_GETFINISH.Show vbModeless
            
                WSHT_CHRT.Activate
                
                Unload USF_GETFINISH
                Application.StatusBar = "(" & w + 1 & "/" & s + 1 & ")   " & " >>  Finished!"
                
                       
    WSHT_SERS.Visible = xlVeryHidden
    WSHT_SYST.Visible = xlVeryHidden
    WSHT_CHRT.Visible = True


With Application
    .EnableEvents = True
    .ScreenUpdating = True
    .DisplayAlerts = True
'    .StatusBar = True
    .Calculation = xlCalculationManual
End With


End Sub

Private Sub CBUT_CANCEL_Click()
Unload USF_BATCH_EXPORT
Unload USF_GETWAIT
Unload USF_GETFINISH
End Sub











