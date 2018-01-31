VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XX_CHARTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
Application.EnableEvents = True
Application.StatusBar = "CHARTS  >> Simulated and observed data can be compared with model results as time series, annual averages, monthly averages, long time averages and R2 related charts."
End Sub


