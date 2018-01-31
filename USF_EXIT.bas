VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USF_EXIT 
   Caption         =   "Save and Update"
   ClientHeight    =   2970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9390.001
   OleObjectBlob   =   "USF_EXIT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USF_EXIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
Unload Me

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


    With Application
        .EnableEvents = False
        .ScreenUpdating = False
        .DisplayAlerts = False
    '   .StatusBar = False
        .Calculation = xlCalculationManual
    End With
    
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
                WBOOK.Worksheets(1).SaveAs INPUT_PATH & WSHEET.NAME & ".txt", xlTextWindows
                WBOOK.Close False
                Set WBOOK = Nothing
            End If
        Next i
    Next WSHEET
    
    'MsgBox "All tabs were updated!", vbInformation, "HYPE_SY"
    'Call Shell("explorer.exe " & INPUT_PATH, vbNormalFocus)
    
    ActiveSheet.Select ' çoklu WSHEET seçimi unutulursa, diðerlerinin üzerine yazýlmasýn diye!

With Application
    .EnableEvents = True
    .ScreenUpdating = True
    .DisplayAlerts = True
    .StatusBar = False
    .StatusBar = True
    .Calculation = xlCalculationManual
End With



Application.StatusBar = "Saved and Updated!"

Application.DisplayAlerts = False
ThisWorkbook.Save
Application.DisplayAlerts = True

Application.Quit

End Sub


Private Sub CommandButton3_Click()
Unload Me
Application.DisplayAlerts = False
ThisWorkbook.Save
Application.DisplayAlerts = True
Application.Quit
End Sub


Private Sub CommandButton2_Click()
Unload Me
ThisWorkbook.Saved = True
Application.Quit
End Sub

