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
    InvalidateRibbonControl RIBBON_CTL_STYLES_EDIT
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabStyles"
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler

    ' Detect changes in the settings dropdown lists
    If Not Intersect(Target, StylesSheet.Range("E1:AZ1")) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
    End If

CleanExit:
    Application.enableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Sheet: StylesSheet - " & Err.number & " " & Err.Description, vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
    Err.Clear
    Resume CleanExit
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target Is Nothing Or Target.Cells.count <> 1 Then Exit Sub

    InvalidateRibbonControl RIBBON_CTL_STYLES_EDIT

    ' Safely remove existing Edit button
    On Error Resume Next
    Me.buttons("FormatEditButton").Delete
    On Error GoTo 0

    Dim formatCol As Long
    formatCol = GetSettingColNum(SETTINGS_STYLES_COL_FORMAT)

    If Intersect(Target, Me.columns(formatCol)) Is Nothing Then Exit Sub

    Dim row As Long
    row = Target.row

    Dim styleType As String
    styleType = StylesSheet.Cells(row, GetSettingColNum(SETTINGS_STYLES_COL_OBJECT_TYPE)).value

    If Not (styleType = TYPE_NODE Or styleType = TYPE_EDGE Or styleType = TYPE_SUBGRAPH_OPEN) Then Exit Sub
    If Me.ProtectContents Then Exit Sub

    ' Try to add a Form Control button
    On Error Resume Next
    Dim btn As button
    Set btn = Me.buttons.Add(Left:=Target.Left + Target.Width + 5, _
                             Top:=Target.Top + 1, _
                             Width:=20, _
                             height:=20)
    If Not btn Is Nothing Then
        With btn
            .name = "FormatEditButton"
            .caption = "..."
            .OnAction = "RestoreStyleDesigner"
        End With
    End If
    On Error GoTo 0
End Sub



