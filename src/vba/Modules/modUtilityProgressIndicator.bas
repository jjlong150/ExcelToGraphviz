Attribute VB_Name = "modUtilityProgressIndicator"
' Copyright (c) 2015-2023 Jeffrey J. Long. All rights reserved
'@Folder("Utility.ProgressIndicator")

Option Explicit

Public Sub UpdateProgressIndicator(ByVal pctCompl As Long)
    If ProgressIndicatorForm.visible Then
        ProgressIndicatorForm.Text.caption = pctCompl & "%"
        ProgressIndicatorForm.Bar.Width = pctCompl * 2
        ProgressIndicatorForm.Repaint
    End If
End Sub

Public Sub ShowProgressIndicator(ByVal title As String)
    ProgressIndicatorForm.caption = title
    ProgressIndicatorForm.show vbModeless
End Sub

Public Sub HideProgressIndicator()
    Unload ProgressIndicatorForm
End Sub
