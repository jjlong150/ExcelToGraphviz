VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CellValueEditForm 
   Caption         =   "Edit Cell Value"
   ClientHeight    =   8640.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15765
   OleObjectBlob   =   "CellValueEditForm.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "CellValueEditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@IgnoreModule HungarianNotation
'@Folder("Relationship Visualizer.Forms.CellValueEdit")

Option Explicit

Private Sub btnSave_Click()
    ' Save the updated value back to the original cell
    ActiveCell.value = txtMultiline.Text
    Unload Me ' Close the userform
End Sub

Private Sub btnCancel_Click()
    ' Close the userform without saving
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' Translate the controls to the local language
    CellValueEditForm.caption = GetLabel("CellValueEditFormCaption")
    btnSave.caption = GetLabel("CellValueEditFormSaveButton")
    btnCancel.caption = GetLabel("CellValueEditFormCancelButton")
    
    ' Initialize the textbox with the cell value
    txtMultiline.Text = ActiveCell.value
    txtMultiline.SetFocus
    txtMultiline.SelStart = 0
End Sub
