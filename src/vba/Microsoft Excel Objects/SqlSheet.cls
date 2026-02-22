VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved

'@IgnoreModule ParameterNotUsed
'@Folder("Relationship Visualizer.Sheets.SQL")

Option Explicit

Private Sub Worksheet_Activate()
#If Mac Then
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabSql"
#Else
    ActivateTabSql
#End If
    InvalidateRibbonControl RIBBON_CTL_SQL_EDIT_CELL
    InvalidateRibbonControl RIBBON_CTL_SQL_CLIPBOARD
End Sub

Private Sub Worksheet_Deactivate()
    InvalidateRibbonControl RIBBON_CTL_SQL_EDIT_CELL
    InvalidateRibbonControl RIBBON_CTL_SQL_CLIPBOARD
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    InvalidateRibbonControl RIBBON_CTL_SQL_EDIT_CELL
    InvalidateRibbonControl RIBBON_CTL_SQL_CLIPBOARD
    RefreshRibbon
    On Error GoTo 0

    ' Get configs once ? reuse for both creation and cleanup
    Dim configs() As ButtonConfig
    configs = GetFloatingButtonConfigs()

    ' Cleanup using the shared utility
    RemoveFloatingButtons Me, GetButtonNamesFromConfigs(configs)

    ' Create buttons using the shared utility
    UpdateFloatingButtonsOnSheet Me, Target, configs
End Sub

' Define button configurations specific to this worksheet
Private Function GetFloatingButtonConfigs() As ButtonConfig()
    Dim configs() As ButtonConfig
    ReDim configs(0 To 2)   ' ? change upper bound when adding more

    With configs(0)
        .ColumnNumber = GetSettingColNum(SETTINGS_SQL_COL_SQL_STATEMENT)
        .ButtonName = "SqlEditButton"
        .ActionMacro = "ShowSQLEditForm"
        .IconText = ChrW(&H270E)
        .VOffset = 0
        .HOffset = 5
        .ValidationFunc = ""
    End With

    With configs(1)
        .ColumnNumber = GetSettingColNum(SETTINGS_SQL_COL_SQL_STATEMENT)
        .ButtonName = "SqlRunButton"
        .ActionMacro = "RunOneSqlStatement"
        .IconText = ChrW(&H25B6)
        .VOffset = 22
        .HOffset = 5
        .ValidationFunc = "IsSqlRowActive"
    End With

    With configs(2)
        .ColumnNumber = GetSettingColNum(SETTINGS_SQL_COL_STATUS)
        .ButtonName = "SqlStatusButton"
        .ActionMacro = "ShowSQLEditForm"
        .IconText = ChrW(&H2315)
        .VOffset = 0
        .HOffset = 5
        .ValidationFunc = ""
    End With

    GetFloatingButtonConfigs = configs
End Function


    
