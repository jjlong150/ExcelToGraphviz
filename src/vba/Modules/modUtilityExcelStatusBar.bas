Attribute VB_Name = "modUtilityExcelStatusBar"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Utility.Excel")

Option Explicit

Public Sub UpdateStatusBar(ByVal statusMessage As String)
    Application.StatusBar = statusMessage
    DoEvents
End Sub

Public Sub UpdateStatusBarForNSeconds(ByVal statusMessage As String, ByVal seconds As Long)
    Application.StatusBar = statusMessage
    DoEvents
    Application.OnTime Now + TimeSerial(0, 0, seconds), "ClearStatusBar"
End Sub

Public Sub ClearStatusBar()
    Application.StatusBar = False
    DoEvents
End Sub

