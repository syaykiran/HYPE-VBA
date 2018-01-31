Attribute VB_Name = "XEXPORTGITHUB"
Private Sub XEXPORTTOGITHUB()

Dim wb As Workbook
Dim FolderName As String
Dim GetFileExtension As String
wbpath = ActiveWorkbook.Path
FolderName = "E:\HYPE_VBA_CODE\HYPE-VBA\"

    
    For Each VBComp In ActiveWorkbook.VBProject.VBComponents
    

        
        On Error Resume Next
        Err.Clear
                 Select Case VBComp.Type
            Case vbext_ct_ClassModule
                GetFileExtension = ".cls"
            Case vbext_ct_Document
                GetFileExtension = ".cls"
            Case vbext_ct_MSForm
                GetFileExtension = ".frm"
            Case vbext_ct_StdModule
                GetFileExtension = ".bas"
            Case Else
                GetFileExtension = ".bas"
        End Select
        
        VBComp.Export FolderName & VBComp.NAME & GetFileExtension

    Next
    
End Sub

 
