VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XX_WORKBOOK_HYPE_VBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

Dim INPUT_PATHS As String
On Error Resume Next
INPUT_PATHS = ThisWorkbook.Path & "\INPUT\"
Debug.Print INPUT_PATHS

With Application
    .EnableEvents = True
    .ScreenUpdating = True
    .DisplayAlerts = True
    .StatusBar = True
    .Calculation = xlCalculationManual
End With

If Dir(INPUT_PATHS, vbDirectory) = "" Then
    
    ThisWorkbook.Worksheets("010101").Visible = True
    ThisWorkbook.Worksheets("010101").Select
    USF_LOAD_STARTUP.Show
End If

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

Dim INPUT_PATHS As String

INPUT_PATHS = ThisWorkbook.Path & "\INPUT\"

If Dir(INPUT_PATHS, vbDirectory) = "" Then

Else
    USF_EXIT.Show
End If
End Sub
