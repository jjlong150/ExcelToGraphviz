VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StylesSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.Styles")

Option Explicit

Private Sub Worksheet_Activate()
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabStyles"
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo ErrorHandler

    ' Detect changes in the settings dropdown lists
    If Not Intersect(Target, StylesSheet.Range("E1:AZ1")) Is Nothing Then
        Application.EnableEvents = False
        RefreshRibbon
    End If
    
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    If Err.number > 0 Then
        MsgBox "Sheet: StylesSheet - " & Err.number & " " & Err.Description, vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
        Err.Clear
        Resume Next
    End If
End Sub

