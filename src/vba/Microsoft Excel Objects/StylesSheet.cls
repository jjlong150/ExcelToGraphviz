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
#If Mac Then
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabStyles"
#Else
    ActivateTabStyles
#End If
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
    EmitMessage "Sheet: StylesSheet - " & Err.number & " " & Err.Description
    Err.Clear
    Resume CleanExit
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells.CountLarge > 1 Or Target.Cells.count = 0 Then Exit Sub
    
    On Error Resume Next
    InvalidateRibbonControl RIBBON_CTL_STYLES_EDIT
    On Error GoTo 0
    
    ' Remove old button/icon
    On Error Resume Next
    Me.buttons("FormatEditButton").Delete
    Me.Shapes("FormatEditButton").Delete
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
    
    ' Add clickable pencil icon with subtle button look
    Dim shp As shape
    Set shp = Me.Shapes.AddShape(msoShapeRoundedRectangle, _
                Target.left + Target.Width + 5, _
                Target.top + 1, 20, 20)
    With shp
        .name = "FormatEditButton"
        .Fill.ForeColor.RGB = RGB(245, 245, 245)      ' very light gray background
        .Fill.Transparency = 0
        .Line.Weight = 0.75
        .Line.ForeColor.RGB = RGB(180, 180, 180)      ' soft border
        .TextFrame2.TextRange.Characters.Text = ChrW(&H270E)  ' pencil
        .TextFrame2.TextRange.Font.name = "Segoe UI Symbol"
        .TextFrame2.TextRange.Font.Size = 16
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(60, 60, 60)  ' dark gray pencil
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
        .Placement = xlMoveAndSize
        .OnAction = "RestoreStyleDesigner"
    End With
End Sub

