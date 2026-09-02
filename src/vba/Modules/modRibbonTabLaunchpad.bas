Attribute VB_Name = "modRibbonTabLaunchpad"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modRibbonTabLaunchpad
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Excel UI / Ribbon
'
' ROLE:
'   Centralized visibility controller for all auxiliary worksheets (Help,
'   Settings, Source, SQL, Console, SVG, Lists, Diagnostics, Styles, About,
'   Exchange, Translations).
'
' RESPONSIBILITIES:
'   - Toggle worksheet visibility via SETTINGS_* named ranges.
'   - Activate appropriate sheet when shown; return to DataSheet when hidden.
'   - Manage macOS-specific console availability.
'   - Handle localization controls (language, verbose mode, translation sheets).
'
' INTERACTIONS:
'   - Ribbon XML: CustomUI.xml, CustomUI14.xml.
'   - Worksheets: DataSheet, SettingsSheet, SourceSheet, SqlSheet, etc.
'
' CROSS-PLATFORM NOTES:
'   - macOS console availability depends on AppleScript version.
'
' ERROR HANDLING:
'   - Localized; Ribbon hydration remains stable.
'
' RELATED WIKI PAGES:
'   - Launchpad Tab
'   - Worksheet Architecture
' =============================================================================

Option Explicit

' ===========================================================================
' Callbacks for helpAttributes

Private Sub helpAttributes_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_HELP_ATTRIBUTES).value = Toggle(pressed, TOGGLE_SHOW, TOGGLE_HIDE)
    If SettingsSheet.Range(SETTINGS_HELP_ATTRIBUTES).value = TOGGLE_SHOW Then
        HelpAttributesSheet.visible = True
        HelpAttributesSheet.Activate
    Else
        HelpAttributesSheet.visible = False
        DataSheet.Activate
    End If
    RefreshRibbon
End Sub

Private Sub helpAttributes_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_HELP_ATTRIBUTES)
End Sub

' ===========================================================================
' Callbacks for helpColors

Private Sub helpColors_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_HELP_COLORS).value = Toggle(pressed, TOGGLE_SHOW, TOGGLE_HIDE)
    If SettingsSheet.Range(SETTINGS_HELP_COLORS).value = TOGGLE_SHOW Then
        HelpColorsSheet.visible = True
        HelpColorsSheet.Activate
    Else
        HelpColorsSheet.visible = False
        DataSheet.Activate
    End If
    RefreshRibbon
End Sub

Private Sub helpColors_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_HELP_COLORS)
End Sub

' ===========================================================================
' Callbacks for helpShapes

Private Sub helpShapes_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_HELP_SHAPES).value = Toggle(pressed, TOGGLE_SHOW, TOGGLE_HIDE)
    If SettingsSheet.Range(SETTINGS_HELP_SHAPES).value = TOGGLE_SHOW Then
        HelpShapesSheet.visible = True
        HelpShapesSheet.Activate
    Else
        HelpShapesSheet.visible = False
        DataSheet.Activate
    End If
    RefreshRibbon
End Sub

Private Sub helpShapes_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_HELP_SHAPES)
End Sub

' ===========================================================================
' Callbacks for toggleSettings

Private Sub toggleSettings_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SETTINGS).value = Toggle(pressed, TOGGLE_SHOW, TOGGLE_HIDE)
    If SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SETTINGS).value = TOGGLE_SHOW Then
        SettingsSheet.visible = True
        SettingsSheet.Activate
        TabSelectCmdLineOptions
    Else
        SettingsSheet.visible = False
        DataSheet.Activate
    End If
    RefreshRibbon
End Sub

Private Sub toggleSettings_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_SETTINGS)
End Sub

' ===========================================================================
' Callbacks for toggleSource

Private Sub toggleSource_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SOURCE).value = Toggle(pressed, TOGGLE_SHOW, TOGGLE_HIDE)
    If SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SOURCE).value = TOGGLE_SHOW Then
        Application.enableEvents = False
        ClearSource
        ShowSource CreateGraphSource()
        Application.enableEvents = True
        SourceSheet.visible = True
        SourceSheet.Activate
    Else
        SourceSheet.visible = False
        ClearSourceWorksheet
        DataSheet.Activate
    End If
    RefreshRibbon
End Sub

Private Sub toggleSource_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_SOURCE)
End Sub

' ===========================================================================
' Callbacks for toggleSql

Private Sub toggleSql_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SQL).value = Toggle(pressed, TOGGLE_SHOW, TOGGLE_HIDE)
    If SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SQL).value = TOGGLE_SHOW Then
        SqlSheet.visible = True
        SqlSheet.Activate
     Else
        SqlSheet.visible = False
        DataSheet.Activate
     End If
    RefreshRibbon
End Sub

Private Sub toggleSql_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_SQL)
End Sub

' ===========================================================================
' Callbacks for toggleConsole

Public Function enableConsole() As Boolean
#If Mac Then
    enableConsole = False
    
    Dim applescriptVersion As String
    applescriptVersion = Trim$(RunAppleScriptTask("getVersion", vbNullString))
    If applescriptVersion <> vbNullString Then
        If CInt(applescriptVersion) >= 2 Then
            enableConsole = True
        End If
    End If
#Else
    enableConsole = True
#End If
End Function

Private Sub toggleConsole_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_CONSOLE).value = Toggle(pressed, TOGGLE_SHOW, TOGGLE_HIDE)
    If SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_CONSOLE).value = TOGGLE_SHOW Then
        ConsoleSheet.visible = True
        ConsoleSheet.Activate
    Else
        ConsoleSheet.visible = False
        DataSheet.Activate
    End If
    RefreshRibbon
End Sub

Private Sub toggleConsole_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_CONSOLE)
End Sub

Private Sub toggleConsole_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = enableConsole()
End Sub

' ===========================================================================
' Callbacks for toggleSvg

Private Sub toggleSvg_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SVG).value = Toggle(pressed, TOGGLE_SHOW, TOGGLE_HIDE)
    If SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SVG).value = TOGGLE_SHOW Then
        SvgSheet.visible = True
        SvgSheet.Activate
    Else
        SvgSheet.visible = False
        DataSheet.Activate
    End If
    RefreshRibbon
End Sub

Private Sub toggleSvg_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_SVG)
End Sub

' ===========================================================================
' Callbacks for toggleLists

Private Sub toggleLists_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_LISTS).value = Toggle(pressed, TOGGLE_SHOW, TOGGLE_HIDE)
    If SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_LISTS).value = TOGGLE_SHOW Then
        ListsSheet.visible = True
        ListsSheet.Activate
    Else
        ListsSheet.visible = False
        DataSheet.Activate
    End If
    RefreshRibbon
End Sub

Private Sub toggleLists_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_LISTS)
End Sub

' ===========================================================================
' Callbacks for toggleDiagnostics

Private Sub toggleDiagnostics_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_DIAGNOSTICS).value = Toggle(pressed, TOGGLE_SHOW, TOGGLE_HIDE)
    If SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_DIAGNOSTICS).value = TOGGLE_SHOW Then
        DiagnosticsSheet.visible = True
        DiagnosticsSheet.Activate
    Else
        DiagnosticsSheet.visible = False
        DataSheet.Activate
    End If
    RefreshRibbon
End Sub

Private Sub toggleDiagnostics_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_DIAGNOSTICS)
End Sub

' ===========================================================================
' Callbacks for toggleStyleDesigner

Private Sub toggleStyleDesigner_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_STYLE_DESIGNER).value = Toggle(pressed, TOGGLE_SHOW, TOGGLE_HIDE)
    If SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_STYLE_DESIGNER).value = TOGGLE_SHOW Then
        StyleDesignerSheet.visible = True
        StyleDesignerSheet.Activate
    Else
        StyleDesignerSheet.visible = False
        DataSheet.Activate
    End If
    RefreshRibbon
End Sub

Private Sub toggleStyleDesigner_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_STYLE_DESIGNER)
End Sub

' ===========================================================================
' Callbacks for toggleStyles

Private Sub toggleStyles_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_STYLES).value = Toggle(pressed, TOGGLE_SHOW, TOGGLE_HIDE)
    If SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_STYLES).value = TOGGLE_SHOW Then
        StylesSheet.visible = True
        StylesSheet.Activate
    Else
        StylesSheet.visible = False
        DataSheet.Activate
    End If
    RefreshRibbon
End Sub

Private Sub toggleStyles_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_STYLES)
End Sub

' ===========================================================================
' Callbacks for toggleAbout

Private Sub toggleAbout_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_ABOUT).value = Toggle(pressed, TOGGLE_SHOW, TOGGLE_HIDE)
    If SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_ABOUT).value = TOGGLE_SHOW Then
        AboutSheet.visible = True
        AboutSheet.Activate
    Else
        AboutSheet.visible = False
        DataSheet.Activate
    End If
    RefreshRibbon
End Sub

Private Sub toggleAbout_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_ABOUT)
End Sub

' ===========================================================================
' Callbacks for toggleExchange

Private Sub toggleExchange_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TABS_TOGGLE_EXCHANGE).value = Toggle(pressed, TOGGLE_SHOW, TOGGLE_HIDE)
    If SettingsSheet.Range(SETTINGS_TABS_TOGGLE_EXCHANGE).value = TOGGLE_SHOW Then
        Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabExchange"
    End If
    RefreshRibbon
End Sub

Private Sub toggleExchange_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TABS_TOGGLE_EXCHANGE)
End Sub

' ===========================================================================
' Callbacks for toggleGraphviz

Private Sub toggleGraphviz_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TABS_TOGGLE_GRAPHVIZ).value = Toggle(pressed, TOGGLE_SHOW, TOGGLE_HIDE)
    If SettingsSheet.Range(SETTINGS_TABS_TOGGLE_GRAPHVIZ).value = TOGGLE_SHOW Then
        Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabGraphviz"
    End If
    RefreshRibbon
End Sub

Private Sub toggleGraphviz_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TABS_TOGGLE_GRAPHVIZ)
End Sub

' ===========================================================================
' Callbacks for Help

Private Sub worksheetsHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLWorksheetsTab").value, NewWindow:=True
End Sub

Private Sub language_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SettingsSheet.Range(SETTINGS_LANGUAGE).value = replace(controlId, "language", "locale")
    Localize
    RefreshRibbon
End Sub

Private Sub language_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = replace(SettingsSheet.Range(SETTINGS_LANGUAGE).value, "locale", "language")
End Sub

Private Sub language_getVisible(ByVal control As IRibbonControl, ByRef makeVisible As Variant)
    Dim workbookSheet As Variant
    Dim languageCount As Long
    languageCount = 0
    
    makeVisible = True
    
    ' Enumerate the worksheets and count the number which begin with "locale_"
    For Each workbookSheet In ThisWorkbook.Sheets
        If StartsWith(workbookSheet.name, RIBBON_LOCALE_PREFIX) Then
            languageCount = languageCount + 1
        End If
    Next
    
    ' Only make the controls visible if multiple languages have been provided
    If languageCount <= 1 Then
        makeVisible = False
    End If
End Sub

' ===========================================================================
' Callbacks for languageVerbose

Private Sub languageVerbose_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SetVerbose (pressed)
    RefreshRibbon
End Sub

Private Sub languageVerbose_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetVerbose()
End Sub

' ===========================================================================
' Callbacks for nodeMetric

Private Sub nodeMetric_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    If pressed Then
        StyleDesignerSheet.Range(DESIGNER_NODE_METRIC).value = TOGGLE_YES
    Else
        StyleDesignerSheet.Range(DESIGNER_NODE_METRIC).value = TOGGLE_NO
    End If
    StyleDesignerSheet.Range("NodeHeight,NodeWidth").ClearContents
    InvalidateRibbonControl RIBBON_CTL_NODE_HEIGHT
    InvalidateRibbonControl RIBBON_CTL_NODE_HEIGHT_METRIC
    InvalidateRibbonControl RIBBON_CTL_NODE_WIDTH
    InvalidateRibbonControl RIBBON_CTL_NODE_WIDTH_METRIC
    InvalidateRibbonControl RIBBON_CTL_CLUSTER_MARGIN
    InvalidateRibbonControl RIBBON_CTL_CLUSTER_MARGIN_MM
End Sub

Private Sub nodeMetric_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    If StyleDesignerSheet.Range(DESIGNER_NODE_METRIC).value = vbNullString Then
        returnedVal = False
    Else
        returnedVal = GetCellBoolean(StyleDesignerSheet.name, DESIGNER_NODE_METRIC)
    End If
End Sub

' ===========================================================================
' Callbacks for toggleSettings

Private Sub translations_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    Dim id As String
    id = replace(control.id, "-", "_")  ' Excel won't let you have a hyphen in a cell name
    SettingsSheet.Range(id).value = Toggle(pressed, TOGGLE_SHOW, TOGGLE_HIDE)
    If SettingsSheet.Range(id).value = TOGGLE_SHOW Then
        ActiveWorkbook.Sheets.[_Default](control.id).visible = True
        ActiveWorkbook.Sheets.[_Default](control.id).Activate
    Else
        ActiveWorkbook.Sheets.[_Default](control.id).visible = False
    End If
End Sub

Private Sub translations_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    Dim id As String
    id = replace(control.id, "-", "_")  ' Excel won't let you have a hyphen in a cell name
    pressed = GetSettingBoolean(id)
End Sub

