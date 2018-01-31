VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WS_Filedir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
Application.EnableEvents = True
Application.StatusBar = " >> Filedir.txt  :  Provides location of info.txt (optional)."
End Sub

