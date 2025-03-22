VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DotSourceForm 
   Caption         =   "dot"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10665
   OleObjectBlob   =   "DotSourceForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "DotSourceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@IgnoreModule FunctionReturnValueDiscarded
'@Folder("Relationship Visualizer.Forms.DotSource")

Option Explicit

Private Sub CopyButton_Click()
#If Not Mac Then
    ClipBoard_SetData (DotSourceForm.dotMultiline.value)
#End If
End Sub

Private Sub FontSizeDecrease_Click()
    If DotSourceForm.dotMultiline.font.Size > 8 Then
        DotSourceForm.dotMultiline.font.Size = DotSourceForm.dotMultiline.font.Size - 2
    End If
    
    If DotSourceForm.dotMultiline.font.Size < 8 Then
        DotSourceForm.FontSizeDecrease.enabled = False
    End If
End Sub

Private Sub FontSizeIncrease_Click()
    DotSourceForm.dotMultiline.font.Size = DotSourceForm.dotMultiline.font.Size + 2
    
    If DotSourceForm.dotMultiline.font.Size > 6 Then
        DotSourceForm.FontSizeDecrease.enabled = True
    End If
End Sub

Private Sub UserForm_Activate()
    DotSourceForm.dotMultiline.WordWrap = False
    DotSourceForm.dotMultiline.ScrollBars = fmScrollBarsBoth
    
#If Mac Then
    DotSourceForm.CopyButton.visible = False
#Else
    DotSourceForm.CopyButton.visible = True
#End If
End Sub


Private Sub wordWrapToggle_Click()
    If DotSourceForm.wordWrapToggle.value = True Then
        DotSourceForm.dotMultiline.WordWrap = True
        DotSourceForm.dotMultiline.ScrollBars = fmScrollBarsVertical
    Else
        DotSourceForm.dotMultiline.WordWrap = False
        DotSourceForm.dotMultiline.ScrollBars = fmScrollBarsBoth
    End If
End Sub

