Attribute VB_Name = "Module2"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
    End With
    Selection.Font.Size = 10
    Selection.Font.Size = 9
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = "&""Cambria,Bold""&8SUB&K00-034 ID: 0200097" & Chr(10) & "SUB ID: "
        .CenterFooter = ""
        .RightFooter = "&9&K00-034&P / &N"
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
    Application.WindowState = xlMinimized
    Application.WindowState = xlNormal
End Sub
