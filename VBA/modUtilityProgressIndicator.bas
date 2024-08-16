Attribute VB_Name = "modUtilityProgressIndicator"
' Copyright (c) 2015-2022 Jeffrey J. Long. All rights reserved

Option Explicit

Public Sub UpdateProgressIndicator(pctCompl As Long)
    If ProgressIndicatorForm.visible Then
        ProgressIndicatorForm.Text.caption = pctCompl & "%"
        ProgressIndicatorForm.Bar.Width = pctCompl * 2
        ProgressIndicatorForm.Repaint
    End If
End Sub

Public Sub ShowProgressIndicator(title As String)
    ProgressIndicatorForm.caption = title
    ProgressIndicatorForm.show vbModeless
    OptimizeCode_Begin
End Sub

Public Sub HideProgressIndicator()
    OptimizeCode_End
    Unload ProgressIndicatorForm
End Sub
