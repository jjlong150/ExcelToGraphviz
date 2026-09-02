Attribute VB_Name = "modRibbonTabStyles"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modRibbonTabStyles
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Excel UI / Ribbon Callbacks
'
' ROLE:
'   Ribbon callback module for the Styles Tab. Provides worksheet-level style
'   previewing, suffix/affix configuration, concatenation formatting, and
'   integration with the Style Designer. Acts as the UI controller for managing
'   and inspecting style definitions stored on the Styles worksheet.
'
' RESPONSIBILITIES:
'   - Trigger style previews:
'       o PreviewStyleForCurrentRow
'       o GenerateStylesPreviewAll
'       o ClearStylesPreview
'   - Manage style affix (open/close) and concatenation format settings via
'     named ranges on the Settings worksheet.
'   - Provide "Edit Style" enablement logic based on the active row's object
'     type (Node, Edge, Subgraph).
'   - Launch the Style Designer (RestoreStyleDesigner) from Ribbon actions.
'   - Navigate to Styles-tab help content using workbook-configured URLs.
'   - Synchronize Ribbon state through getText/getEnabled callbacks.
'
' INTERACTIONS:
'   - Named Ranges:
'       SETTINGS_STYLES_AFFIX_OPEN,
'       SETTINGS_STYLES_AFFIX_CLOSE,
'       SETTINGS_STYLES_CONCAT_FORMAT,
'       SETTINGS_STYLES_COL_OBJECT_TYPE,
'       HelpURLStylesTab.
'   - Worksheets: StylesSheet, SettingsSheet, DataSheet.
'   - Modules: ClearStylesPreview, PreviewStyleForCurrentRow,
'              GenerateStylesPreviewAll, RestoreStyleDesigner.
'
' CROSS-PLATFORM NOTES:
'   - Fully supported on Windows and macOS.
'   - All actions rely on worksheet operations and hyperlink navigation.
'
' ERROR HANDLING:
'   - Validates active sheet, selection type, and row count before enabling
'     "Edit Style."
'   - All callbacks follow IRibbonControl signature requirements.
'
' RELATED WIKI PAGES:
'   - Styles & the Style Gallery
'   - Style Designer Ribbon Tab
'   - Worksheet Architecture & Named Ranges
'   - Working with the Data Worksheet
' =============================================================================

Option Explicit

Private Sub stylesClear_onAction(ByVal control As IRibbonControl)
    ClearStylesPreview
End Sub

Private Sub stylesPreview_onAction(ByVal control As IRibbonControl)
    PreviewStyleForCurrentRow
End Sub

Private Sub stylesPreviewAll_onAction(ByVal control As IRibbonControl)
    StylesSheet.Activate
    ClearStylesPreview
    GenerateStylesPreviewAll
    ClearStatusBar
End Sub

' ===========================================================================
' Callbacks for stylesConcat

Private Sub stylesConcat_onChange(ByVal control As IRibbonControl, ByVal Text As String)
    SettingsSheet.Range(SETTINGS_STYLES_CONCAT_FORMAT).value = Text
End Sub

Private Sub stylesConcat_getText(ByVal control As IRibbonControl, ByRef Text As Variant)
    Text = Trim$(SettingsSheet.Range(SETTINGS_STYLES_CONCAT_FORMAT))
End Sub

' ===========================================================================
' Callbacks for stylesSuffixBegin

Private Sub stylesSuffixBegin_onChange(ByVal control As IRibbonControl, ByVal Text As String)
    SettingsSheet.Range(SETTINGS_STYLES_AFFIX_OPEN).value = Text
End Sub

Private Sub stylesSuffixBegin_getText(ByVal control As IRibbonControl, ByRef Text As Variant)
    Text = Trim$(SettingsSheet.Range(SETTINGS_STYLES_AFFIX_OPEN))
End Sub

' ===========================================================================
' Callbacks for stylesSuffixEnd

Private Sub stylesSuffixEnd_onChange(ByVal control As IRibbonControl, ByVal Text As String)
    SettingsSheet.Range(SETTINGS_STYLES_AFFIX_CLOSE).value = Text
End Sub

Private Sub stylesSuffixEnd_getText(ByVal control As IRibbonControl, ByRef Text As Variant)
    Text = Trim$(SettingsSheet.Range(SETTINGS_STYLES_AFFIX_CLOSE))
End Sub

' ===========================================================================
' Callbacks for Help

Private Sub stylesHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLStylesTab").value, NewWindow:=True
End Sub

' ===========================================================================
' Callbacks for stylesEdit

Private Sub stylesEdit_onAction(ByVal control As IRibbonControl)
    RestoreStyleDesigner
End Sub

Private Sub stylesEdit_getEnabled(ByVal control As IRibbonControl, ByRef Enabled As Variant)
    Enabled = False

    If ActiveSheet.name <> StylesSheet.name Then Exit Sub
    If Not TypeOf Selection Is Range Then Exit Sub
    If Selection.rows.Count <> 1 Then Exit Sub

    Dim row As Long
    row = Selection.row

    Dim typeCol As Long
    typeCol = GetSettingColNum(SETTINGS_STYLES_COL_OBJECT_TYPE)

    Dim styleType As String
    styleType = StylesSheet.Cells(row, typeCol).value

    If Not (styleType = TYPE_NODE Or styleType = TYPE_EDGE Or styleType = TYPE_SUBGRAPH_OPEN) Then Exit Sub

    Enabled = True
End Sub
