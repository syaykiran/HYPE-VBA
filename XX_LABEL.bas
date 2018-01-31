VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XX_LABEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
Application.EnableEvents = True
Application.StatusBar = "LABEL  >> Information about stations table. The information according to the SUBID value in this list is transferred to charts label and footnotes of PDF files."
End Sub

