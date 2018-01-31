VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} USF_LOAD 
   Caption         =   "Help"
   ClientHeight    =   10185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12765
   OleObjectBlob   =   "USF_LOAD.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "USF_LOAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton3_Click()
Unload Me
USF_LOAD_TEMPLATES.Show vbModal


End Sub

Private Sub CommandButton4_Click()
Unload Me
End Sub

