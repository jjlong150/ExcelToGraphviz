Attribute VB_Name = "modRibbonTabExchange"
' Copyright (c) 2015-2022 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Ribbon.Tabs")
'@IgnoreModule ProcedureNotUsed

Option Explicit

' ===========================================================================
' Callbacks for importJson

'@Ignore ParameterNotUsed
Public Sub importJson_onAction(ByVal control As IRibbonControl)
    ImportData
End Sub

'@Ignore ParameterNotUsed
Public Sub importJson_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub

' ===========================================================================
' Callbacks for exportJson

'@Ignore ParameterNotUsed
Public Sub exportJson_onAction(ByVal control As IRibbonControl)
    ExportData
End Sub

'@Ignore ParameterNotUsed
Public Sub exportJson_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub

'@Ignore ParameterNotUsed
Public Sub exchangeData_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_EXCHANGE_DATA_WORKSHEET).Value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

'@Ignore ParameterNotUsed
Public Sub exchangeData_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_DATA_WORKSHEET)
End Sub

'@Ignore ParameterNotUsed
Public Sub exchangeStyles_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_EXCHANGE_STYLES_WORKSHEET).Value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

'@Ignore ParameterNotUsed
Public Sub exchangeStyles_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_STYLES_WORKSHEET)
End Sub

'@Ignore ParameterNotUsed
Public Sub exchangeSql_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_EXCHANGE_SQL_WORKSHEET).Value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

'@Ignore ParameterNotUsed
Public Sub exchangeSql_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_SQL_WORKSHEET)
End Sub

'@Ignore ParameterNotUsed
Public Sub exchangeGraphOptions_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_EXCHANGE_GRAPH_OPTIONS).Value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

'@Ignore ParameterNotUsed
Public Sub exchangeGraphOptions_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_GRAPH_OPTIONS)
End Sub

'@Ignore ParameterNotUsed
Public Sub exchangeWorksheetLayouts_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_EXCHANGE_WORKSHEET_LAYOUTS).Value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

'@Ignore ParameterNotUsed
Public Sub exchangeWorksheetLayouts_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_WORKSHEET_LAYOUTS)
End Sub

'@Ignore ParameterNotUsed
Public Sub exchangeMetadata_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_EXCHANGE_METADATA).Value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

'@Ignore ParameterNotUsed
Public Sub exchangeMetadata_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_METADATA)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportDataRowNumber_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_DATA_EXPORT_ROW).Value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportDataRowNumber_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_DATA_EXPORT_ROW)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportDataRowHeight_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_DATA_EXPORT_HEIGHT).Value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportDataRowHeight_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_DATA_EXPORT_HEIGHT)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportDataRowVisible_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_DATA_EXPORT_VISIBLE).Value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportDataRowVisible_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_DATA_EXPORT_VISIBLE)
End Sub

'@Ignore ParameterNotUsed
Public Sub importDataRowAppend_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_DATA_IMPORT_ACTION).Value = IMPORT_APPEND

    InvalidateRibbonControl RIBBON_CTL_IMPORT_DATA_ROW_APPEND
    InvalidateRibbonControl RIBBON_CTL_IMPORT_DATA_ROW_REPLACE
End Sub

'@Ignore ParameterNotUsed
Public Sub importDataRowAppend_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_EXCHANGE_DATA_IMPORT_ACTION, IMPORT_APPEND)
End Sub

'@Ignore ParameterNotUsed
Public Sub importDataRowReplace_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_DATA_IMPORT_ACTION).Value = IMPORT_REPLACE

    InvalidateRibbonControl RIBBON_CTL_IMPORT_DATA_ROW_APPEND
    InvalidateRibbonControl RIBBON_CTL_IMPORT_DATA_ROW_REPLACE
End Sub

'@Ignore ParameterNotUsed
Public Sub importDataRowReplace_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_EXCHANGE_DATA_IMPORT_ACTION, IMPORT_REPLACE)
End Sub


' ---------- styles ---------------

'@Ignore ParameterNotUsed
Public Sub exportStylesRowNumber_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_STYLES_EXPORT_ROW).Value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportStylesRowNumber_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_STYLES_EXPORT_ROW)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportStylesRowHeight_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_STYLES_EXPORT_HEIGHT).Value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportStylesRowHeight_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_STYLES_EXPORT_HEIGHT)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportStylesRowVisible_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_STYLES_EXPORT_VISIBLE).Value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportStylesRowVisible_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_STYLES_EXPORT_VISIBLE)
End Sub

'@Ignore ParameterNotUsed
Public Sub importStylesRowAppend_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_STYLES_IMPORT_ACTION).Value = IMPORT_APPEND

    InvalidateRibbonControl RIBBON_CTL_IMPORT_STYLES_ROW_APPEND
    InvalidateRibbonControl RIBBON_CTL_IMPORT_STYLES_ROW_REPLACE
End Sub

'@Ignore ParameterNotUsed
Public Sub importStylesRowAppend_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_EXCHANGE_STYLES_IMPORT_ACTION, IMPORT_APPEND)
End Sub

'@Ignore ParameterNotUsed
Public Sub importStylesRowReplace_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_STYLES_IMPORT_ACTION).Value = IMPORT_REPLACE

    InvalidateRibbonControl RIBBON_CTL_IMPORT_STYLES_ROW_APPEND
    InvalidateRibbonControl RIBBON_CTL_IMPORT_STYLES_ROW_REPLACE
End Sub

'@Ignore ParameterNotUsed
Public Sub importStylesRowReplace_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_EXCHANGE_STYLES_IMPORT_ACTION, IMPORT_REPLACE)
End Sub

' ---------- sql ---------------

'@Ignore ParameterNotUsed
Public Sub exportSqlRowNumber_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_SQL_EXPORT_ROW).Value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportSqlRowNumber_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_SQL_EXPORT_ROW)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportSqlRowHeight_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_SQL_EXPORT_HEIGHT).Value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportSqlRowHeight_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_SQL_EXPORT_HEIGHT)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportSqlRowVisible_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_SQL_EXPORT_VISIBLE).Value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportSqlRowVisible_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_SQL_EXPORT_VISIBLE)
End Sub

'@Ignore ParameterNotUsed
Public Sub importSqlRowAppend_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_SQL_IMPORT_ACTION).Value = IMPORT_APPEND

    InvalidateRibbonControl RIBBON_CTL_IMPORT_SQL_ROW_APPEND
    InvalidateRibbonControl RIBBON_CTL_IMPORT_SQL_ROW_REPLACE
End Sub

'@Ignore ParameterNotUsed
Public Sub importSqlRowAppend_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_EXCHANGE_SQL_IMPORT_ACTION, IMPORT_APPEND)
End Sub

'@Ignore ParameterNotUsed
Public Sub importSqlRowReplace_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_SQL_IMPORT_ACTION).Value = IMPORT_REPLACE

    InvalidateRibbonControl RIBBON_CTL_IMPORT_SQL_ROW_APPEND
    InvalidateRibbonControl RIBBON_CTL_IMPORT_SQL_ROW_REPLACE
End Sub

'@Ignore ParameterNotUsed
Public Sub importSqlRowReplace_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_EXCHANGE_SQL_IMPORT_ACTION, IMPORT_REPLACE)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportOptionsData_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub

'@Ignore ParameterNotUsed
Public Sub importOptionsData_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub

'@Ignore ParameterNotUsed
Public Sub exportOptionsStyles_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub

'@Ignore ParameterNotUsed
Public Sub importOptionsStyles_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub

'@Ignore ParameterNotUsed
Public Sub exportOptionsSql_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub

'@Ignore ParameterNotUsed
Public Sub importOptionsSql_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub











' ---------- svg ---------------

'@Ignore ParameterNotUsed
Public Sub exchangeSvg_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_EXCHANGE_SVG_WORKSHEET).Value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

'@Ignore ParameterNotUsed
Public Sub exchangeSvg_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_SVG_WORKSHEET)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportSvgRowNumber_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_SVG_EXPORT_ROW).Value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportSvgRowNumber_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_SVG_EXPORT_ROW)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportSvgRowHeight_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_SVG_EXPORT_HEIGHT).Value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportSvgRowHeight_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_SVG_EXPORT_HEIGHT)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportSvgRowVisible_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_SVG_EXPORT_VISIBLE).Value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

'@Ignore ParameterNotUsed
Public Sub exportSvgRowVisible_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_SVG_EXPORT_VISIBLE)
End Sub

'@Ignore ParameterNotUsed
Public Sub importSvgRowAppend_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_SVG_IMPORT_ACTION).Value = IMPORT_APPEND

    InvalidateRibbonControl RIBBON_CTL_IMPORT_SVG_ROW_APPEND
    InvalidateRibbonControl RIBBON_CTL_IMPORT_SVG_ROW_REPLACE
End Sub

'@Ignore ParameterNotUsed
Public Sub importSvgRowAppend_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_EXCHANGE_SVG_IMPORT_ACTION, IMPORT_APPEND)
End Sub

'@Ignore ParameterNotUsed
Public Sub importSvgRowReplace_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_SVG_IMPORT_ACTION).Value = IMPORT_REPLACE

    InvalidateRibbonControl RIBBON_CTL_IMPORT_SVG_ROW_APPEND
    InvalidateRibbonControl RIBBON_CTL_IMPORT_SVG_ROW_REPLACE
End Sub

'@Ignore ParameterNotUsed
Public Sub importSvgRowReplace_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_EXCHANGE_SVG_IMPORT_ACTION, IMPORT_REPLACE)
End Sub

'@Ignore ParameterNotUsed
Public Sub importOptionsSvg_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub

' ===========================================================================
' Callbacks for Help

'@Ignore ParameterNotUsed
Public Sub toolsHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLExchangeTab").Value, NewWindow:=True
End Sub


