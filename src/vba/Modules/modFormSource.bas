Attribute VB_Name = "modFormSource"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Forms.DotSource")

Option Explicit

Public Sub ClearSourceForm()
    DotSourceForm.dotMultiline.Text = vbNullString
End Sub

Public Sub ShowSourceForm()
    DotSourceForm.CopyButton.caption = GetLabel("sourceFormCopy")
    DotSourceForm.wordWrapToggle.caption = GetLabel("sourceFormWrapText")
    DotSourceForm.show
    ClearSourceForm
End Sub

Public Sub HideSourceForm()
    Unload DotSourceForm
End Sub

Public Sub DisplaySourceInForm(ByVal dotSource As String)
    If Not DotSourceForm.visible Then Exit Sub
    
    Dim popupSource As String
    popupSource = dotSource
    replace popupSource, vbLf, Chr$(10)
    DotSourceForm.dotMultiline.Text = popupSource
End Sub

