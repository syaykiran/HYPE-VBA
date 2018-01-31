VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XX_COMMENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Activate()
Application.EnableEvents = True
Application.StatusBar = "COMMENT  >> HYPE Model parameters and their descriptions. It can be edited by user if requested. Or it can be  updated according to new release versions."
End Sub

