Attribute VB_Name = "XXHIDE"
Sub MCR_HIDE()

Dim WSHEET_ARRAY As Variant
Dim WSHEET_INDEX As Variant
Dim WSHEET As Worksheet
Dim i As Integer

On Error Resume Next
ThisWorkbook.Worksheets("010101").Visible = True
ThisWorkbook.Worksheets("010101").Visible.Activate

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
                                                 
                            


                            For Each WSHEET In ThisWorkbook.Worksheets
                                For i = 0 To UBound(WSHEET_ARRAY, 1)
                                    If InStr(WSHEET.NAME, WSHEET_ARRAY(i)) > 0 Then
                                    WSHEET.Visible = xlSheetVeryHidden
                                    End If
                                Next i
                            Next WSHEET
                            
                            
  WSHEET_ARRAY = Array("Filedir", _
                                                 "LABEL", _
                                                 "COMMENT", _
                                                 "CHARTS", _
                                                 "LIST", "SERIES", "SYSTEM")
                                              
                                                 

                            For Each WSHEET In ThisWorkbook.Worksheets
                                For i = 0 To UBound(WSHEET_ARRAY, 1)
                                    If InStr(WSHEET.NAME, WSHEET_ARRAY(i)) > 0 Then
                                    WSHEET.Visible = xlSheetVeryHidden
                                    End If
                                Next i
                            Next WSHEET
                            
                ThisWorkbook.Worksheets("LABEL").Visible = xlSheetVeryHidden
                ThisWorkbook.Worksheets("COMMENT").Visible = xlSheetVeryHidden
                ThisWorkbook.Worksheets("CHARTS").Visible = xlSheetVeryHidden

                                                

' + exe yi geri taþý diðerlerini de yedekle
End Sub
