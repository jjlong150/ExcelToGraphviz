Attribute VB_Name = "modRibbonTabLaunchpad"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Ribbon.Tabs")
'@IgnoreModule ParameterNotUsed, UseMeaningfulName, UnassignedVariableUsage, ProcedureNotUsed

Option Explicit

' ===========================================================================
' Callbacks for helpAttributes

'@Ignore ParameterNotUsed
Public Sub helpAttributes_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
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

'@Ignore ParameterNotUsed
Public Sub helpAttributes_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_HELP_ATTRIBUTES)
End Sub

' ===========================================================================
' Callbacks for helpColors

'@Ignore ParameterNotUsed
Public Sub helpColors_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
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

'@Ignore ParameterNotUsed
Public Sub helpColors_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_HELP_COLORS)
End Sub

' ===========================================================================
' Callbacks for helpShapes

'@Ignore ParameterNotUsed
Public Sub helpShapes_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
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

'@Ignore ParameterNotUsed
Public Sub helpShapes_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_HELP_SHAPES)
End Sub

' ===========================================================================
' Callbacks for toggleSettings

'@Ignore ParameterNotUsed
Public Sub toggleSettings_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SETTINGS).value = Toggle(pressed, TOGGLE_SHOW, TOGGLE_HIDE)
    If SettingsSheet.Range(SETTINGS_TOOLS_TOGGLE_SETTINGS).value = TOGGLE_SHOW Then
        SettingsSheet.visible = True
        SettingsSheet.Activate
    Else
        SettingsSheet.visible = False
        DataSheet.Activate
    End If
    RefreshRibbon
End Sub

'@Ignore ParameterNotUsed
Public Sub toggleSettings_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_SETTINGS)
End Sub

' ===========================================================================
' Callbacks for toggleSource

'@Ignore ParameterNotUsed
Public Sub toggleSource_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
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

'@Ignore ParameterNotUsed
Public Sub toggleSource_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_SOURCE)
End Sub

' ===========================================================================
' Callbacks for toggleSql

'@Ignore ParameterNotUsed
Public Sub toggleSql_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
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

'@Ignore ParameterNotUsed
Public Sub toggleSql_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
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

'@Ignore ParameterNotUsed
Public Sub toggleConsole_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
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

'@Ignore ParameterNotUsed
Public Sub toggleConsole_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_CONSOLE)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Public Sub toggleConsole_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = enableConsole()
End Sub

' ===========================================================================
' Callbacks for toggleSvg

'@Ignore ParameterNotUsed
Public Sub toggleSvg_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
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

'@Ignore ParameterNotUsed
Public Sub toggleSvg_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_SVG)
End Sub

' ===========================================================================
' Callbacks for toggleLists

'@Ignore ParameterNotUsed
Public Sub toggleLists_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
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

'@Ignore ParameterNotUsed
Public Sub toggleLists_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_LISTS)
End Sub

' ===========================================================================
' Callbacks for toggleDiagnostics

'@Ignore ParameterNotUsed
Public Sub toggleDiagnostics_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
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

'@Ignore ParameterNotUsed
Public Sub toggleDiagnostics_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_DIAGNOSTICS)
End Sub

' ===========================================================================
' Callbacks for toggleStyleDesigner

'@Ignore ParameterNotUsed
Public Sub toggleStyleDesigner_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
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

'@Ignore ParameterNotUsed
Public Sub toggleStyleDesigner_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_STYLE_DESIGNER)
End Sub

' ===========================================================================
' Callbacks for toggleStyles

'@Ignore ParameterNotUsed
Public Sub toggleStyles_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
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

'@Ignore ParameterNotUsed
Public Sub toggleStyles_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_STYLES)
End Sub

' ===========================================================================
' Callbacks for toggleAbout

'@Ignore ParameterNotUsed
Public Sub toggleAbout_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
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

'@Ignore ParameterNotUsed
Public Sub toggleAbout_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_ABOUT)
End Sub

' ===========================================================================
' Callbacks for toggleExchange

'@Ignore ParameterNotUsed
Public Sub toggleExchange_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SettingsSheet.Range(SETTINGS_TABS_TOGGLE_EXCHANGE).value = Toggle(pressed, TOGGLE_SHOW, TOGGLE_HIDE)
    If SettingsSheet.Range(SETTINGS_TABS_TOGGLE_EXCHANGE).value = TOGGLE_SHOW Then
        Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabExchange"
    End If
    RefreshRibbon
End Sub


'@Ignore ParameterNotUsed
Public Sub toggleExchange_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = GetSettingBoolean(SETTINGS_TABS_TOGGLE_EXCHANGE)
End Sub

' ===========================================================================
' Callbacks for Help

'@Ignore ParameterNotUsed
Public Sub worksheetsHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLWorksheetsTab").value, NewWindow:=True
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub language_getItemLabel(ByVal control As IRibbonControl, ByVal index As Long, ByRef label As Variant)
    label = GetLabel(control.ID)
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub language_onAction(ByVal control As IRibbonControl, ByVal controlId As String, ByVal index As Long)
    SettingsSheet.Range(SETTINGS_LANGUAGE).value = replace(controlId, "language", "locale")
    Localize
    RefreshRibbon
End Sub

Public Sub language_GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemId As Variant)
    itemId = replace(SettingsSheet.Range(SETTINGS_LANGUAGE).value, "locale", "language")
End Sub

Public Sub language_getVisible(ByVal control As IRibbonControl, ByRef makeVisible As Variant)
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

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub languageVerbose_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    SetVerbose (pressed)
    RefreshRibbon
End Sub

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub languageVerbose_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetVerbose()
End Sub

' ===========================================================================
' Callbacks for Help

'@Ignore ParameterNotUsed
Public Sub localeHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLLocaleTab").value, NewWindow:=True
End Sub

' ===========================================================================
' Callbacks for nodeMetric

'@Ignore ProcedureNotUsed, ParameterNotUsed
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

'@Ignore ProcedureNotUsed, ParameterNotUsed
Private Sub nodeMetric_getPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    If StyleDesignerSheet.Range(DESIGNER_NODE_METRIC).value = vbNullString Then
        returnedVal = False
    Else
        returnedVal = GetCellBoolean(StyleDesignerSheet.name, DESIGNER_NODE_METRIC)
    End If
End Sub

' ===========================================================================
' Callbacks for toggleSettings

'@Ignore ParameterNotUsed
Public Sub translations_onAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    Dim ID As String
    ID = replace(control.ID, "-", "_")  ' Excel won't let you have a hyphen in a cell name
    SettingsSheet.Range(ID).value = Toggle(pressed, TOGGLE_SHOW, TOGGLE_HIDE)
    If SettingsSheet.Range(ID).value = TOGGLE_SHOW Then
        ActiveWorkbook.Sheets.[_Default](control.ID).visible = True
        ActiveWorkbook.Sheets.[_Default](control.ID).Activate
    Else
        ActiveWorkbook.Sheets.[_Default](control.ID).visible = False
    End If
End Sub

'@Ignore ParameterNotUsed
Public Sub translations_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    Dim ID As String
    ID = replace(control.ID, "-", "_")  ' Excel won't let you have a hyphen in a cell name
    pressed = GetSettingBoolean(ID)
End Sub

