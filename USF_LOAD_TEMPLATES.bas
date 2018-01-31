VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USF_LOAD_TEMPLATES 
   Caption         =   "Load"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7335
   OleObjectBlob   =   "USF_LOAD_TEMPLATES.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USF_LOAD_TEMPLATES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMBX_LOAD_TEMPLATES_Change()

  Select Case Me.CMBX_LOAD_TEMPLATES
    Case Is = "CPET HYPE Model"
         Me.CPET.Visible = True
         Me.DEMO1.Visible = False
         Me.DEMO2.Visible = False
         Me.DEMO3.Visible = False
         Me.Frame1.Caption = "CPET HYPE Model"
        
    Case Is = "HYPE Demo 1"
         Me.CPET.Visible = False
         Me.DEMO1.Visible = True
         Me.DEMO2.Visible = False
         Me.DEMO3.Visible = False
         Me.Frame1.Caption = "HYPE Demo 1"
         
    Case Is = "HYPE Demo 2"
         Me.CPET.Visible = False
         Me.DEMO1.Visible = False
         Me.DEMO2.Visible = True
         Me.DEMO3.Visible = False
         Me.Frame1.Caption = "HYPE Demo 2"
        
    Case Is = "HYPE Demo 3"
         Me.CPET.Visible = False
         Me.DEMO1.Visible = False
         Me.DEMO2.Visible = False
         Me.DEMO3.Visible = True
         Me.Frame1.Caption = "HYPE Demo 3"
        
      Case Else
         Me.CPET.Visible = False
         Me.DEMO1.Visible = False
         Me.DEMO2.Visible = False
         Me.DEMO3.Visible = False
         Me.Frame1.Caption = "Description"
         
End Select

End Sub

Private Sub CMND_LOAD_TEMPLATES_Click()
Dim MAIN_FOLDER As Variant
Dim MAIN_FOLDERS As Variant
Dim INPUT_PATH As Variant
Dim INPUT_PATHS As String
Dim OUTPUT_PATH As Variant
Dim OUTPUT_PATHS As Variant
Dim BACKUP_PATH As Variant
Dim MODEL_NAME As Variant
Dim OLD_PATH As Variant
Dim EXE_PATH As Variant
Dim WSHEET As Worksheet
Dim WBOOK As Workbook
Dim WSHEET_ARRAY As Variant
Dim COUNT, i As Integer
Dim FSO As Object




Select Case Me.CMBX_LOAD_TEMPLATES
    Case Is = "CPET HYPE Model"
    
        Unload Me
        With Application
            .EnableEvents = False
            .ScreenUpdating = False
            .DisplayAlerts = False
        '   .StatusBar = False
            .Calculation = xlCalculationManual
        End With
        
        
        On Error Resume Next
        
        MAIN_FOLDER = ThisWorkbook.Path
        MAIN_FOLDERS = MAIN_FOLDER & "\"
        OUTPUT_PATH = MAIN_FOLDERS & "OUTPUT"
        INPUT_PATH = MAIN_FOLDERS & "INPUT"
        BACKUP_PATH = MAIN_FOLDERS & "BACKUP"
        INPUT_PATHS = INPUT_PATH & "\"
        OUTPUT_PATHS = OUTPUT_PATH & "\"
        MODEL_NAME = "HYPE.exe"
        EXE_PATH = INPUT_PATHS & MODEL_NAME
        OLD_PATH = MAIN_FOLDERS & MODEL_NAME
        
        Set FSO = CreateObject("Scripting.Filesystemobject")
        
        If Dir(INPUT_PATHS, vbDirectory) = "" Then
        
            If Dir(OLD_PATH, vbDirectory) = "" Then
            Unload Me
            MsgBox "Cannot find [HYPE.exe]!" & vbCrLf & _
             vbCrLf & _
            "Please make sure that [HYPE VBA.xlsb] and [HYPE.exe] are in the same folder and the name of the exe file is correct." & vbCrLf & _
             vbCrLf & _
            "And please try again!", _
            vbCrLf & _
            vbExclamation, "HYPE VBA"
            
            USF_LOAD.HIDE
            Exit Sub
        
            Else
                USF_LOAD_TEMPLATES.HIDE
                USF_LOAD_GETWAIT.Show vbModeless
                
                Shell ("cmd /c md " & Chr(34) & INPUT_PATH & Chr(34)) 'input klasörü oluþtur. mkdir kullanýlmadý, bu komut hem çoklu dosya oluþturabilir, hem error vermez dosya varsa, avantajlý!
                Shell ("cmd /c md " & Chr(34) & OUTPUT_PATH & Chr(34)) 'output klasörü oluþtur
                Shell ("cmd /c md " & Chr(34) & BACKUP_PATH & Chr(34)) 'yedek klasörü oluþtur
                Application.StatusBar = "Creat..."
                 
                Do
                    If Not Dir(INPUT_PATH, vbDirectory) = "" Then Exit Do
                    DoEvents
                    Application.Wait Now + TimeValue("0:00:01")
                Loop
            
                Call FSO.MoveFile(OLD_PATH, EXE_PATH)
        
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
                        WSHEET.Visible = xlSheetVisible
                            Application.StatusBar = "Processing... >> " & WSHEET_ARRAY(i)
                            Set WBOOK = Workbooks.Add
                            ThisWorkbook.Worksheets(WSHEET.NAME).Copy WBOOK.Worksheets(1)
                            WBOOK.Worksheets(1).SaveAs INPUT_PATHS & WSHEET.NAME & ".txt", xlTextWindows
                            WBOOK.Close False
                            Set WBOOK = Nothing
                        End If
                    Next i
                Next WSHEET
                               
                                                 

                ThisWorkbook.Worksheets("LABEL").Visible = True
                ThisWorkbook.Worksheets("COMMENT").Visible = True
                ThisWorkbook.Worksheets("CHARTS").Visible = True
                
            
'                Worksheets("Info").Range("B2").Value = "=UI_MODELDIR"
'                Worksheets("Info").Range("B3").Value = "=UI_RESULTDIR"

                Worksheets("Info").Range("UI_MODELDIR").Value = INPUT_PATHS
                Worksheets("Info").Range("UI_RESULTDIR").Value = OUTPUT_PATHS


                USF_LOAD_GETWAIT.HIDE
                Unload Me
                                             
                MsgBox "The template was loaded successfully!", vbInformation, "HYPE VBA"
                Call Shell("explorer.exe " & MAIN_FOLDER, vbNormalFocus) 'klasörü aç
            
            End If
            
        Else
            MsgBox "You have already a template!", vbExclamation, "HYPE VBA"
            Unload Me
            Exit Sub
            
        End If
        
        ActiveSheet.Select ' çoklu WSHEET seçimi unutulursa, diðerlerinin üzerine yazýlmasýn diye!
        Unload Me
        
        
        With Application
            .EnableEvents = True
            .ScreenUpdating = True
            .DisplayAlerts = True
            .StatusBar = True
            .Calculation = xlCalculationManual
        End With
        
        If Not ThisWorkbook.Saved Then
            If MsgBox("Do you want to save this Excel file?", vbYesNo, "Save?") = vbYes Then
                ThisWorkbook.Save
                ThisWorkbook.Worksheets("010101").Visible = xlVeryHidden
            End If
        End If
    
    Case Is = "HYPE Demo 1"
    ' UNDER CONSTRACTION
    Case Is = "HYPE Demo 2"
    ' UNDER CONSTRACTION
    Case Is = "HYPE Demo 3"
    ' UNDER CONSTRACTION
    Case Else
     ' UNDER CONSTRACTION
End Select



' s.yaykýran 10/11/2016
' s.yaykýran 20/11/16
' s.yaykýran 4/1/18
End Sub



Private Sub UserForm_Initialize()


With CMBX_LOAD_TEMPLATES
    .AddItem "CPET HYPE Model"
    .AddItem "HYPE Demo 1"
    .AddItem "HYPE Demo 2"
    .AddItem "HYPE Demo 3"
End With
         Me.CPET.Visible = False
         Me.DEMO1.Visible = False
         Me.DEMO2.Visible = False
         Me.DEMO3.Visible = False

  

  
End Sub
