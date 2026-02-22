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
        Application.EnableEvents = False
        RefreshRibbon
    End If

CleanExit:
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    EmitMessage "Sheet: StylesSheet - " & Err.number & " " & Err.Description
    Err.Clear
    Resume CleanExit
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Exit if protected worksheet
    If Me.ProtectContents Then Exit Sub
    
    ' Enable or disable the edit button in the ribbon
    On Error Resume Next
    InvalidateRibbonControl RIBBON_CTL_STYLES_EDIT
    On Error GoTo 0
    
    ' Get button configs and use for both creation and cleanup
    Dim configs() As ButtonConfig
    configs = GetFloatingButtonConfigs()

    ' Cleanup old buttons, and create new ones
    RemoveFloatingButtons Me, GetButtonNamesFromConfigs(configs)
    UpdateFloatingButtonsOnSheet Me, Target, configs
End Sub

' Define button configurations specific to this worksheet
Private Function GetFloatingButtonConfigs() As ButtonConfig()
    Dim configs() As ButtonConfig
    ReDim configs(0 To 2)   ' ? change upper bound when adding more

    With configs(0)
        .ColumnNumber = GetSettingColNum(SETTINGS_STYLES_COL_FORMAT)
        .ButtonName = "StyleRefreshButton"
        .ActionMacro = "PreviewStyleForCurrentRow"
        .IconText = ChrW(&H27F3)
        .VOffset = 0
        .HOffset = 5
        .ValidationFunc = "IsStylesRowActive"
    End With
    
    With configs(1)
        .ColumnNumber = GetSettingColNum(SETTINGS_STYLES_COL_FORMAT)
        .ButtonName = "FormatEditButton"
        .ActionMacro = "RestoreStyleDesigner"
        .IconText = ChrW(&H270E)
        .VOffset = 22
        .HOffset = 5
        .ValidationFunc = "IsStylesRowActive"
    End With

    With configs(2)
        .ColumnNumber = GetSettingColNum(SETTINGS_STYLES_COL_STYLE)
        .ButtonName = "StyleRefreshButton"
        .ActionMacro = "PreviewStyleForCurrentRow"
        .IconText = ChrW(&H27F3)
        .VOffset = 0
        .HOffset = 5
        .ValidationFunc = "IsStylesRowActive"
    End With

    GetFloatingButtonConfigs = configs
End Function


