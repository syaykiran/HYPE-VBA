Attribute VB_Name = "UPDATE"
Option Explicit
Sub MCR_UPDATE(control As IRibbonControl)
'sayfayý .txt olarak yazdýr

Dim WSHEET As Worksheet
Dim WBOOK As Workbook
Dim WSHEET_ARRAY As Variant
Dim WSHEET_INDEX As Variant
Dim COUNT, i As Integer
Dim INPUT_PATHS As String
Dim MAIN_FOLDER As String
Dim INPUT_PATH As String


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
    
    ThisWorkbook.Activate

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
    
    
           
            
    For Each WSHEET In ActiveWindow.SelectedSheets
    Application.StatusBar = "Updating... >> " & WSHEET.NAME
        For i = 1 To UBound(WSHEET_ARRAY, 1)
        If WSHEET.NAME = WSHEET_ARRAY(i) Then
              WSHEET.Calculate
              Set WBOOK = Workbooks.Add
              COUNT = WBOOK.Worksheets.COUNT
              WSHEET.Copy WBOOK.Worksheets(1)
              WBOOK.Worksheets(1).SaveAs INPUT_PATHS & WSHEET.NAME & ".txt", xlTextWindows 'filedir adresine kaydet
              WBOOK.Close False
              Set WBOOK = Nothing
            End If
        Next i
    Next WSHEET
    
    Application.StatusBar = "Selected tabs were updated!"
    'Call Shell("explorer.exe " & INPUT_PATH, vbNormalFocus)
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
