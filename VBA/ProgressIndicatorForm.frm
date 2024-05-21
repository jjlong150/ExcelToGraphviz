VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressIndicatorForm 
   Caption         =   "Progress Indicator"
   ClientHeight    =   1230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4320
   OleObjectBlob   =   "ProgressIndicatorForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressIndicatorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Activate()
    ' Reset the progress indicator to 0%
    UpdateProgressIndicator 0
End Sub

Private Sub UserForm_Initialize()
    Me.Top = Application.Top + ((Application.height - ProgressIndicatorForm.height) / 2)
    Me.Left = Application.Left + ((Application.Width - ProgressIndicatorForm.Width) / 2)
End Sub
