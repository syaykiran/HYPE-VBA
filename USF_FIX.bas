VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USF_FIX 
   Caption         =   "Fix (Under Development)"
   ClientHeight    =   2280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6255
   OleObjectBlob   =   "USF_FIX.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USF_FIX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CBUT_NO_Click()
Unload Me
End Sub

Private Sub CommandButton1_Click()

Dim WSHEET_ARRAY As Variant
Dim WSHEET_INDEX As Variant
Dim ACT_RANGE As Range
Dim COMMENT_RANGE As Range
Dim LOOK_RANGE As Range
Dim FOUNDED As Range
Dim Target As Range
Dim START_TIME As Double
Dim MIN_ELAPSED
 
    
USF_FIX.HIDE
USF_FIX_GETWAIT.Show vbModeless
DoEvents
    
    
'KRONOMETRE BASLAT
START_TIME = Timer
Debug.Print START_TIME

On Error Resume Next

USF_FIX_GETWAIT.Show vbModeless

DoEvents
With Application
    .EnableEvents = False
    .ScreenUpdating = False
    .DisplayAlerts = False
    .StatusBar = False
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
                   "MgmtData", _
                   "PointSourceData")
    

For Each WSHEET_INDEX In WSHEET_ARRAY
Debug.Print WSHEET_INDEX
    Worksheets(WSHEET_INDEX).Select
    
    Set ACT_RANGE = Worksheets(WSHEET_INDEX).Range("A1").CurrentRegion
    Set ACT_RANGE = ACT_RANGE.Resize(ACT_RANGE.Rows.COUNT + 20, ACT_RANGE.Columns.COUNT + 10)
    
'****************** FORMAT CONDITIONAL *******************************

    Worksheets(WSHEET_INDEX).Cells.FormatConditions.Delete
    
    With Cells(1, 1).FormatConditions.Add(Type:=xlExpression, Formula1:="=ISNUMBER(FIND(""!"",$A1))")
    .Font.Italic = True
    .Font.Bold = True
    .Font.Color = -16744448
    .StopIfTrue = False
    End With
    
    With Cells(1, 1).FormatConditions.Add(Type:=xlExpression, Formula1:="=MOD(row(),2)=1")
    .Interior.Color = RGB(240, 240, 240)
    .StopIfTrue = False
    End With
    
    Cells.FormatConditions(1).ModifyAppliesToRange ACT_RANGE
    Cells.FormatConditions(2).ModifyAppliesToRange ACT_RANGE
    ACT_RANGE.Borders.LineStyle = xlNone
    ActiveWindow.DisplayGridlines = False
    
    If Worksheets(WSHEET_INDEX).NAME = "GeoClass" Then
    
        
            Worksheets("GeoClass").Select
            Worksheets("GeoClass").Range("RNG_GEOCLASS_LANDUSE").FormatConditions.Delete
            Worksheets("GeoClass").Range("RNG_GEOCLASS_SOIL").FormatConditions.Delete
            Worksheets("GeoClass").Range("RNG_GEOCLASS_CROP").FormatConditions.Delete
        
        
        
            With Cells(6, 2).FormatConditions.Add(xlCellValue, xlEqual, "1")
            .Font.Bold = True
            .Interior.Color = RGB(255, 255, 102)
            .Font.Color = vbBlack
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
            With Cells(6, 2).FormatConditions.Add(xlCellValue, xlEqual, "2")
            .Font.Bold = True
            .Interior.Color = RGB(112, 173, 71)
            .Font.Color = vbBlack
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
            With Cells(6, 2).FormatConditions.Add(xlCellValue, xlEqual, "3")
            .Font.Bold = True
            .Interior.Color = RGB(198, 224, 180)
            .Font.Color = vbBlack
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
            With Cells(6, 2).FormatConditions.Add(xlCellValue, xlEqual, "4")
            .Font.Bold = True
            .Interior.Color = RGB(142, 169, 219)
            .Font.Color = vbBlack
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
            With Cells(6, 2).FormatConditions.Add(xlCellValue, xlEqual, "5")
            .Font.Bold = True
            .Interior.Color = RGB(204, 0, 102)
            .Font.Color = vbBlack
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
            With Cells(6, 2).FormatConditions.Add(xlCellValue, xlEqual, "6")
            .Font.Bold = True
            .Interior.Color = RGB(255, 240, 244)
            .Font.Color = vbBlack
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
            With Cells(6, 2).FormatConditions.Add(xlCellValue, xlEqual, "7")
            .Font.Bold = True
            .Interior.Color = RGB(0, 176, 240)
            .Font.Color = vbBlack
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
        
        
            With Cells(6, 3).FormatConditions.Add(xlCellValue, xlEqual, "1")
            .Font.Bold = True
            .Interior.Color = RGB(198, 89, 17)
            .Font.Color = vbBlack
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
            With Cells(6, 3).FormatConditions.Add(xlCellValue, xlEqual, "2")
            .Font.Bold = True
            .Interior.Color = RGB(244, 176, 132)
            .Font.Color = vbBlack
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
            With Cells(6, 3).FormatConditions.Add(xlCellValue, xlEqual, "3")
            .Font.Bold = True
            .Interior.Color = RGB(248, 203, 173)
            .Font.Color = vbBlack
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
            With Cells(6, 3).FormatConditions.Add(xlCellValue, xlEqual, "4")
            .Font.Bold = True
            .Interior.Color = RGB(255, 192, 0)
            .Font.Color = vbBlack
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
            With Cells(6, 3).FormatConditions.Add(xlCellValue, xlEqual, "5")
            .Font.Bold = True
            .Interior.Color = RGB(191, 143, 0)
            .Font.Color = vbBlack
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
            With Cells(6, 3).FormatConditions.Add(xlCellValue, xlEqual, "6")
            .Font.Bold = True
            .Interior.Color = RGB(204, 204, 0)
            .Font.Color = vbBlack
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
            With Cells(6, 3).FormatConditions.Add(xlCellValue, xlEqual, "7")
            .Font.Bold = True
            .Interior.Color = RGB(235, 234, 234)
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
        
        
        
            With Cells(6, 4).FormatConditions.Add(Type:=xlExpression, Formula1:="=ISNUMBER(FIND(""0"",$D6))")
            .Font.Bold = True
            .Interior.Color = RGB(240, 240, 240)
            .Font.Color = vbBlack
            .HorizontalAlignment = xlCenter
           .Borders.LineStyle = xlContinuous
           .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
            With Cells(6, 4).FormatConditions.Add(xlCellValue, xlEqual, "1")
            .Font.Bold = True
            .Interior.Color = RGB(230, 230, 230)
            .Font.Color = vbBlack
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
            With Cells(6, 4).FormatConditions.Add(xlCellValue, xlEqual, "2")
            .Font.Bold = True
            .Interior.Color = RGB(220, 220, 220)
            .Font.Color = vbBlack
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
            With Cells(6, 4).FormatConditions.Add(xlCellValue, xlEqual, "3")
            .Font.Bold = True
            .Interior.Color = RGB(200, 200, 200)
            .Font.Color = vbBlack
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
            With Cells(6, 4).FormatConditions.Add(xlCellValue, xlEqual, "4")
            .Font.Bold = True
            .Interior.Color = RGB(190, 190, 190)
            .Font.Color = vbBlack
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
            With Cells(6, 4).FormatConditions.Add(xlCellValue, xlEqual, "6")
            .Font.Bold = True
            .Interior.Color = RGB(175, 175, 175)
            .Font.Color = vbBlack
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
            With Cells(6, 4).FormatConditions.Add(xlCellValue, xlEqual, "7")
            .Font.Bold = True
            .Interior.Color = RGB(160, 160, 160)
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
            With Cells(6, 4).FormatConditions.Add(xlCellValue, xlEqual, "8")
            .Font.Bold = True
            .Interior.Color = RGB(150, 150, 150)
            .HorizontalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(192, 192, 192)
            .StopIfTrue = False
            End With
        
        
            Cells.FormatConditions(3).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_LANDUSE_DEF")
            Cells.FormatConditions(4).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_LANDUSE_DEF")
            Cells.FormatConditions(5).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_LANDUSE_DEF")
            Cells.FormatConditions(6).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_LANDUSE_DEF")
            Cells.FormatConditions(7).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_LANDUSE_DEF")
            Cells.FormatConditions(8).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_LANDUSE_DEF")
            Cells.FormatConditions(9).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_LANDUSE_DEF")
            
            Cells.FormatConditions(10).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_SOIL_DEF")
            Cells.FormatConditions(11).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_SOIL_DEF")
            Cells.FormatConditions(12).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_SOIL_DEF")
            Cells.FormatConditions(13).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_SOIL_DEF")
            Cells.FormatConditions(14).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_SOIL_DEF")
            Cells.FormatConditions(15).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_SOIL_DEF")
            Cells.FormatConditions(16).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_SOIL_DEF")
            
            Cells.FormatConditions(17).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_CROP_DEF")
            Cells.FormatConditions(18).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_CROP_DEF")
            Cells.FormatConditions(19).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_CROP_DEF")
            Cells.FormatConditions(20).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_CROP_DEF")
            Cells.FormatConditions(21).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_CROP_DEF")
            Cells.FormatConditions(22).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_CROP_DEF")
            Cells.FormatConditions(23).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_CROP_DEF")
            Cells.FormatConditions(24).ModifyAppliesToRange Worksheets("GeoClass").Range("RNG_GEOCLASS_CROP_DEF")

    
    End If

'****************** COMMENTS *****************************************
        
    Select Case WSHEET_INDEX
    
        Case Is = "Filedir"
        Set COMMENT_RANGE = Worksheets("Filedir").Range("A1")
        
        Case Is = "Info"
        Set COMMENT_RANGE = ThisWorkbook.Worksheets("Info").UsedRange
        
        Case Is = "Par"
        Set COMMENT_RANGE = ThisWorkbook.Worksheets("Par").Range("RNG_PAR")
        
        Case Is = "GeoClass"
        Set COMMENT_RANGE = ThisWorkbook.Worksheets("GeoClass").Range("RNG_GEOCLASS")
            
        Case Is = "GeoData"
        Set COMMENT_RANGE = ThisWorkbook.Worksheets("GeoData").Range("RNG_GEODATA")
            
        Case Is = "LakeData"
        Set COMMENT_RANGE = ThisWorkbook.Worksheets("LakeData").Range("RNG_LAKEDATA")
            
        Case Is = "BranchData"
        Set COMMENT_RANGE = ThisWorkbook.Worksheets("BranchData").Range("RNG_BRANCHDATA")
        
        Case Is = "CropData"
        Set COMMENT_RANGE = ThisWorkbook.Worksheets("CropData").Range("RNG_CROPDATA")
            
        Case Is = "ForcKey"
        Set COMMENT_RANGE = ThisWorkbook.Worksheets("ForcKey").Range("RNG_FORCKEY")
            
        Case Is = "MgmtData"
        Set COMMENT_RANGE = ThisWorkbook.Worksheets("MgmtData").Range("RNG_MGMTDATA")
            
        Case Is = "PointSourceData"
        Set COMMENT_RANGE = ThisWorkbook.Worksheets("PointSourceData").Range("RNG_POINTSOURCEDATA")
        
        Case Else
        Set COMMENT_RANGE = ACT_RANGE
    
    End Select
    
    Set LOOK_RANGE = Worksheets("COMMENT").Range("DESC_ALL")
    Cells.ClearComments
    
    For Each Target In COMMENT_RANGE
    Set FOUNDED = LOOK_RANGE.Find(What:=Target.Value, lookIn:=xlValues, Lookat:=xlWhole, SearchFormat:=False)
            
           If Not FOUNDED Is Nothing Then
                With Target
                .ClearComments
                .AddComment
                .COMMENT.Text Text:=CStr(WorksheetFunction.VLookup(Target.Value, LOOK_RANGE, 2, False))
                .COMMENT.Shape.TextFrame.Characters.Font.ColorIndex = 51
                .COMMENT.Shape.TextFrame.Characters.Font.Size = 8
                .COMMENT.Shape.TextFrame.Characters.Font.NAME = "Consola"
                
                If Len(.COMMENT.Text) > 350 Then
                    .COMMENT.Shape.Width = 350
                    .COMMENT.Shape.Height = 80
                    
                ElseIf Len(.COMMENT.Text) > 225 Then
                    .COMMENT.Shape.Width = 250
                    .COMMENT.Shape.Height = 50
                
                ElseIf Len(.COMMENT.Text) > 100 Then
                    .COMMENT.Shape.Width = 150
                    .COMMENT.Shape.Height = 50
                    
                ElseIf Len(.COMMENT.Text) > 50 Then
                    .COMMENT.Shape.Width = 120
                    .COMMENT.Shape.Height = 50
                
                ElseIf Len(.COMMENT.Text) > 25 Then
                    .COMMENT.Shape.Width = 110
                    .COMMENT.Shape.Height = 30
                
                Else
                    .COMMENT.Shape.Width = 110
                    .COMMENT.Shape.Height = 40
                End If
                
                .COMMENT.Shape.TextFrame.HorizontalAlignment = xlHAlignCenter
                .COMMENT.Shape.TextFrame.VerticalAlignment = xlVAlignCenter
                .COMMENT.Shape.AutoShapeType = msoShapeRoundedRectangle
                .COMMENT.Shape.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .COMMENT.Shape.Fill.Transparency = 0
                .COMMENT.Visible = False
                End With
            End If

    Next Target
 

Next WSHEET_INDEX



USF_FIX_GETWAIT.HIDE
Unload USF_FIX_GETWAIT
Unload USF_FIX

'Worksheets(1).Select

With Application
    .EnableEvents = True
    .ScreenUpdating = True
    .DisplayAlerts = True
    .StatusBar = True
    .Calculation = xlCalculationManual
End With

'MIN_ELAPSED = Format((Timer - START_TIME) / 86400, "hh:mm:ss")
'Debug.Print Timer
'
'MsgBox "Dosyalar Tamamlandi!" & vbCrLf & _
'       "Islem süresi: " & MIN_ELAPSED







''https://www.mrexcel.com/forum/excel-questions/662266-excel-vba-add-comment-change-cell-value.html
'
'Private Sub Project_BeforeSave(ByVal pj As Project)
''If ActiveProject.NAME = "HYPE_VBA_ver_1_5" Then
'Application.DisplayAlerts = False
'End Sub


    
End Sub

