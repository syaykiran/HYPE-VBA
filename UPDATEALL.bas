Attribute VB_Name = "UPDATEALL"
Option Explicit

Sub MCR_UPDATE_ALL(control As IRibbonControl)

Dim WSHEET_ARRAY As Variant
Dim WSHEET_INDEX As Variant
Dim WSHEET As Worksheet
Dim WBOOK As Workbook
Dim COUNT, i As Integer
Dim MAIN_FOLDER As String
Dim INPUT_PATH As String
Dim INPUT_PATHS As String

On Error Resume Next

INPUT_PATHS = ThisWorkbook.Path & "\INPUT\"

If Dir(INPUT_PATHS, vbDirectory) = "" Then
USF_LOAD_STARTUP.Show vbModal

Else
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
        .DisplayAlerts = False
    '   .StatusBar = False
        .Calculation = xlCalculationManual
    End With
    
ThisWorkbook.Worksheets("010101").Visible = xlVeryHidden

    WSHEET_ARRAY = Array("Filedir", _
                         "Info", _
                         "Par", _
                         "GeoClass", _
                         "GeoData", _
                         "LakeData", _
                         "BranchData", _
                         "CropData", _
                         "ForcKey", _
                         "MgmtData", _
                         "PointSourceData", _
                         "Pobs", _
                         "Tobs", _
                         "Qobs", _
                         "Xobs")
                         
    
    MAIN_FOLDER = ThisWorkbook.Path
    INPUT_PATH = MAIN_FOLDER & "\" & "INPUT" & "\"
    
    'For Each WSHEET In ThisWorkbook.Worksheets
    'If WSHEET.NAME Like "_*" = False Then '  alt çizgi(_) ile baþlayanlarý yazdýrmaz!
    
    
    For Each WSHEET In ThisWorkbook.Worksheets
        For i = 1 To UBound(WSHEET_ARRAY, 1)
            If WSHEET.NAME = WSHEET_ARRAY(i) Then
                Application.StatusBar = "Updating... >> " & WSHEET.NAME
                Set WBOOK = Workbooks.Add
                ThisWorkbook.Worksheets(WSHEET.NAME).Copy WBOOK.Worksheets(1)
                WBOOK.Worksheets(1).SaveAs INPUT_PATHS & WSHEET.NAME & ".txt", xlTextWindows
                WBOOK.Close False
                Set WBOOK = Nothing
            End If
        Next i
    Next WSHEET
    
    'MsgBox "All tabs were updated!", vbInformation, "HYPE_SY"
    'Call Shell("explorer.exe " & INPUT_PATH, vbNormalFocus)
    
    Application.StatusBar = "All tabs were updated!"
    
    ActiveSheet.Select ' çoklu WSHEET seçimi unutulursa, diðerlerinin üzerine yazýlmasýn diye!
End If
With Application
    .EnableEvents = True
    .ScreenUpdating = True
    .DisplayAlerts = True
'    .StatusBar = True
    .Calculation = xlCalculationManual
End With

' s.yaykýran 10/11/2016
End Sub

'Function WorksheetExists2(WorksheetName As String, Optional wb As Workbook) As Boolean
'    If wb Is Nothing Then Set wb = ThisWorkbook
'    With wb
'        On Error Resume Next
'        WorksheetExists2 = (.Sheets(WorksheetName).NAME = WorksheetName)
'        On Error GoTo 0
'    End With
'End Function


