Attribute VB_Name = "modRibbonTabExchange"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modRibbonTabExchange
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Excel UI / Ribbon
'
' ROLE:
'   Callback bridge for the "Exchange" Ribbon Tab, coordinating JSON import/
'   export and granular include/exclude settings for Data, Styles, SQL, SVG,
'   Metadata, Layouts, and Graph Options.
'
' RESPONSIBILITIES:
'   - Dispatch IRibbonControl callbacks for Exchange tab controls.
'   - Manage include/exclude toggles across multiple worksheet categories.
'   - Persist import/export settings via SETTINGS_* named ranges.
'   - Support Append/Replace import modes with Ribbon invalidation.
'
' INTERACTIONS:
'   - Ribbon XML: CustomUI.xml, CustomUI14.xml.
'   - Named Ranges: SETTINGS_TOOLS_EXCHANGE_*, SETTINGS_EXCHANGE_*.
'   - Modules: ImportData, ExportData, SettingsSheet utilities.
'
' CROSS-PLATFORM NOTES:
'   - Fully supported on Windows and macOS.
'
' ERROR HANDLING:
'   - Localized; Ribbon hydration remains stable.
'
' RELATED WIKI PAGES:
'   - JSON Import/Export
'   - Data Exchange Architecture
' =============================================================================

Option Explicit

' ===========================================================================
' Callbacks for importJson

Private Sub importJson_onAction(ByVal control As IRibbonControl)
    ImportData
End Sub

Private Sub importJson_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub

' ===========================================================================
' Callbacks for exportJson

Private Sub exportJson_onAction(ByVal control As IRibbonControl)
    ExportData
End Sub

Private Sub exportJson_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub

Private Sub exchangeData_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_EXCHANGE_DATA_WORKSHEET).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

Private Sub exchangeData_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_DATA_WORKSHEET)
End Sub

Private Sub exchangeStyles_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_EXCHANGE_STYLES_WORKSHEET).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

Private Sub exchangeStyles_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_STYLES_WORKSHEET)
End Sub

Private Sub exchangeSql_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_EXCHANGE_SQL_WORKSHEET).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

Private Sub exchangeSql_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_SQL_WORKSHEET)
End Sub

Private Sub exchangeGraphOptions_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_EXCHANGE_GRAPH_OPTIONS).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

Private Sub exchangeGraphOptions_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_GRAPH_OPTIONS)
End Sub

Private Sub exchangeWorksheetLayouts_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_EXCHANGE_WORKSHEET_LAYOUTS).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

Private Sub exchangeWorksheetLayouts_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_WORKSHEET_LAYOUTS)
End Sub

Private Sub exchangeMetadata_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_EXCHANGE_METADATA).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

Private Sub exchangeMetadata_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_METADATA)
End Sub

Private Sub exportDataRowNumber_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_DATA_EXPORT_ROW).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

Private Sub exportDataRowNumber_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_DATA_EXPORT_ROW)
End Sub

Private Sub exportDataRowHeight_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_DATA_EXPORT_HEIGHT).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

Private Sub exportDataRowHeight_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_DATA_EXPORT_HEIGHT)
End Sub

Private Sub exportDataRowVisible_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_DATA_EXPORT_VISIBLE).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

Private Sub exportDataRowVisible_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_DATA_EXPORT_VISIBLE)
End Sub

Private Sub importDataRowAppend_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_DATA_IMPORT_ACTION).value = IMPORT_APPEND

    InvalidateRibbonControl RIBBON_CTL_IMPORT_DATA_ROW_APPEND
    InvalidateRibbonControl RIBBON_CTL_IMPORT_DATA_ROW_REPLACE
End Sub

Private Sub importDataRowAppend_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_EXCHANGE_DATA_IMPORT_ACTION, IMPORT_APPEND)
End Sub

Private Sub importDataRowReplace_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_DATA_IMPORT_ACTION).value = IMPORT_REPLACE

    InvalidateRibbonControl RIBBON_CTL_IMPORT_DATA_ROW_APPEND
    InvalidateRibbonControl RIBBON_CTL_IMPORT_DATA_ROW_REPLACE
End Sub

Private Sub importDataRowReplace_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_EXCHANGE_DATA_IMPORT_ACTION, IMPORT_REPLACE)
End Sub

' ---------- styles ---------------

Private Sub exportStylesRowNumber_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_STYLES_EXPORT_ROW).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

Private Sub exportStylesRowNumber_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_STYLES_EXPORT_ROW)
End Sub

Private Sub exportStylesRowHeight_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_STYLES_EXPORT_HEIGHT).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

Private Sub exportStylesRowHeight_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_STYLES_EXPORT_HEIGHT)
End Sub

Private Sub exportStylesRowVisible_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_STYLES_EXPORT_VISIBLE).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

Private Sub exportStylesRowVisible_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_STYLES_EXPORT_VISIBLE)
End Sub

Private Sub importStylesRowAppend_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_STYLES_IMPORT_ACTION).value = IMPORT_APPEND

    InvalidateRibbonControl RIBBON_CTL_IMPORT_STYLES_ROW_APPEND
    InvalidateRibbonControl RIBBON_CTL_IMPORT_STYLES_ROW_REPLACE
End Sub

Private Sub importStylesRowAppend_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_EXCHANGE_STYLES_IMPORT_ACTION, IMPORT_APPEND)
End Sub

Private Sub importStylesRowReplace_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_STYLES_IMPORT_ACTION).value = IMPORT_REPLACE

    InvalidateRibbonControl RIBBON_CTL_IMPORT_STYLES_ROW_APPEND
    InvalidateRibbonControl RIBBON_CTL_IMPORT_STYLES_ROW_REPLACE
End Sub

Private Sub importStylesRowReplace_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_EXCHANGE_STYLES_IMPORT_ACTION, IMPORT_REPLACE)
End Sub

' ---------- sql ---------------

Private Sub exportSqlRowNumber_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_SQL_EXPORT_ROW).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

Private Sub exportSqlRowNumber_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_SQL_EXPORT_ROW)
End Sub

Private Sub exportSqlRowHeight_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_SQL_EXPORT_HEIGHT).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

Private Sub exportSqlRowHeight_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_SQL_EXPORT_HEIGHT)
End Sub

Private Sub exportSqlRowVisible_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_SQL_EXPORT_VISIBLE).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

Private Sub exportSqlRowVisible_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_SQL_EXPORT_VISIBLE)
End Sub

Private Sub importSqlRowAppend_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_SQL_IMPORT_ACTION).value = IMPORT_APPEND

    InvalidateRibbonControl RIBBON_CTL_IMPORT_SQL_ROW_APPEND
    InvalidateRibbonControl RIBBON_CTL_IMPORT_SQL_ROW_REPLACE
End Sub

Private Sub importSqlRowAppend_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_EXCHANGE_SQL_IMPORT_ACTION, IMPORT_APPEND)
End Sub

Private Sub importSqlRowReplace_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_SQL_IMPORT_ACTION).value = IMPORT_REPLACE

    InvalidateRibbonControl RIBBON_CTL_IMPORT_SQL_ROW_APPEND
    InvalidateRibbonControl RIBBON_CTL_IMPORT_SQL_ROW_REPLACE
End Sub

Private Sub importSqlRowReplace_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_EXCHANGE_SQL_IMPORT_ACTION, IMPORT_REPLACE)
End Sub

Private Sub exportOptionsData_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub

Private Sub importOptionsData_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub

Private Sub exportOptionsStyles_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub

Private Sub importOptionsStyles_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub

Private Sub exportOptionsSql_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub

Private Sub importOptionsSql_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub

' ---------- svg ---------------

Private Sub exchangeSvg_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_EXCHANGE_SVG_WORKSHEET).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

Private Sub exchangeSvg_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_EXCHANGE_SVG_WORKSHEET)
End Sub

Private Sub exportSvgRowNumber_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_SVG_EXPORT_ROW).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

Private Sub exportSvgRowNumber_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_SVG_EXPORT_ROW)
End Sub

Private Sub exportSvgRowHeight_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_SVG_EXPORT_HEIGHT).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

Private Sub exportSvgRowHeight_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_SVG_EXPORT_HEIGHT)
End Sub

Private Sub exportSvgRowVisible_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_SVG_EXPORT_VISIBLE).value = Toggle(pressed, TOGGLE_INCLUDE, TOGGLE_EXCLUDE)
End Sub

Private Sub exportSvgRowVisible_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_EXCHANGE_SVG_EXPORT_VISIBLE)
End Sub

Private Sub importSvgRowAppend_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_SVG_IMPORT_ACTION).value = IMPORT_APPEND

    InvalidateRibbonControl RIBBON_CTL_IMPORT_SVG_ROW_APPEND
    InvalidateRibbonControl RIBBON_CTL_IMPORT_SVG_ROW_REPLACE
End Sub

Private Sub importSvgRowAppend_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_EXCHANGE_SVG_IMPORT_ACTION, IMPORT_APPEND)
End Sub

Private Sub importSvgRowReplace_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_EXCHANGE_SVG_IMPORT_ACTION).value = IMPORT_REPLACE

    InvalidateRibbonControl RIBBON_CTL_IMPORT_SVG_ROW_APPEND
    InvalidateRibbonControl RIBBON_CTL_IMPORT_SVG_ROW_REPLACE
End Sub

Private Sub importSvgRowReplace_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = getPressed(SettingsSheet.name, SETTINGS_EXCHANGE_SVG_IMPORT_ACTION, IMPORT_REPLACE)
End Sub

Private Sub importOptionsSvg_getEnabled(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = True
End Sub

' ===========================================================================
' Callbacks for Help

Private Sub toolsHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLExchangeTab").value, NewWindow:=True
End Sub


