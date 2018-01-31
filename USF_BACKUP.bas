VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USF_BACKUP 
   Caption         =   "Backup"
   ClientHeight    =   2685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5445
   OleObjectBlob   =   "USF_BACKUP.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USF_BACKUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()

Dim MAIN_FOLDER As Variant
Dim MAIN_FOLDER2 As Variant
Dim INPUT_PATH As Variant
Dim OUTPUT_PATH As Variant
Dim NAME As Variant
Dim NAME_DATE As String
Dim NAME_DATE2 As String
Dim BACKUP_PATH As String
Dim BACKUP_PATHS As String
Dim BACKUP_PATH_XLSX As String
Dim INPUT_PATHS As String
Dim FSO As Object

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.StatusBar = True

Set FSO = CreateObject("scripting.filesystemobject")

Unload USF_BACKUP
NAME = Me.BACKUP_NAME
    
'Application.InputBox(prompt:="Enter backup name", Title:="Backup Name?", Default:="Backup")

If NAME <> False Then
   
    Application.StatusBar = "Backing up files..."
    
    MAIN_FOLDER = ThisWorkbook.Path
    INPUT_PATH = MAIN_FOLDER & "\" & "INPUT"
    OUTPUT_PATH = MAIN_FOLDER & "\" & "OUTPUT"
    BACKUP_PATH = MAIN_FOLDER & "\" & "BACKUP\" & Format(Now, "yyyymmdd_hhmm") & "_" & NAME & "\"
    BACKUP_PATHS = MAIN_FOLDER & "\" & "BACKUP\" & Format(Now, "yyyymmdd_hhmm") & "_" & NAME
    BACKUP_PATH_XLSX = BACKUP_PATH & ThisWorkbook.NAME
    Shell ("cmd /c md " & Chr(34) & BACKUP_PATH & Chr(34))
    
    Application.Wait Now + TimeValue("00:00:01") 'beklemezse dosyayý görmez!
    FSO.CopyFolder INPUT_PATH, BACKUP_PATH, True 'kopyalama olayý
    FSO.CopyFolder OUTPUT_PATH, BACKUP_PATH, True 'kopyalama olayý
    
    ActiveWorkbook.SaveCopyAs BACKUP_PATH_XLSX 'excel kopyalama olayý
    
    MsgBox "All files have been backed up successfully!", vbInformation, "HYPE VBA"
        
    Call Shell("explorer.exe " & BACKUP_PATH, vbNormalFocus)
    
Else
    Application.StatusBar = "Cancelled!"
    Application.Wait Now + TimeValue("00:00:01")
End If

Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.StatusBar = True

End Sub
' s.yaykýran 20/11/16

' ilgili linkler:
' input outout dosya aç makrosu
' http://www.get-digital-help.com/2015/01/27/save-selected-sheets-to-a-pdf-file/
' https://www.experts-exchange.com/questions/27556108/Excel-Macro-to-save-tab-delimited-text-from-multiple-sheets-as-separate-files.html#answer37503113
' http://www.tek-tips.com/faqs.cfm?fid=6756
' http://www.mrexcel.com/forum/excel-questions/70500-visual-basic-applications-remove-last-character-string.html
' http://gethowstuff.com/vba-macro-save-each-sheet-workbook-to-seperate-csv-files/
' http://www.markwithall.com/programming/2014/04/18/save-all-excel-worksheets-as-tab-separated-text.html
' http://www.taltech.com/support/entry/opening_and_closing_an_application_from_vba
' güncelle makrosu
' https://www.reddit.com/r/excel/comments/3r87w2/my_save_to_txt_macro_is_changing_the_filename/
' klasör oluþtur makrosu
' http://superuser.com/questions/799666/creating-folders-and-sub-folders-with-a-vba-macro
' model koþtur makrosu
' http://ss64.com/vb/shellexecute.html
' http://stackoverflow.com/questions/18602979/how-to-give-a-time-delay-of-less-than-one-second-in-excel-vba



Private Sub CommandButton2_Click()
Unload USF_BACKUP
Application.StatusBar = "Cancelled!"
End Sub


