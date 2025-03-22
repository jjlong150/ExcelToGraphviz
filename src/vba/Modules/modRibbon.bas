Attribute VB_Name = "modRibbon"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Ribbon.Loader")
'@IgnoreModule UnreachableCase, ProcedureNotUsed

Option Explicit

Private internalMyRibbon As IRibbonUI

' Get/Let to encapsulate myRibbon
Public Static Property Get myRibbon() As IRibbonUI
    Set myRibbon = internalMyRibbon
End Property

Public Static Property Let myRibbon(ByVal ribbon As IRibbonUI)
    Set internalMyRibbon = ribbon
End Property

' Load the ribbon
Public Sub ribbon_onLoad(ByVal ribbon As IRibbonUI)
    '@Ignore ValueRequired
    myRibbon = ribbon
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ribbon_activateTab"
End Sub

Public Sub ribbon_activateTab()
    On Error GoTo ErrorHandler

    ' Defer initialization of 'settings' worksheet until workbook
    ' startup is complete
    TabSelectGraphOptions
    
    ' Show the appropriate tab for the worksheet displayed
    Select Case ActiveSheet.name
        Case DataSheet.name:            ActivateTabGraphviz
        Case AboutSheet.name:           ActivateTabAbout
        Case ConsoleSheet.name:         ActivateTabConsole
        Case DiagnosticsSheet.name:     ActivateTabDiagnostics
        Case GraphSheet.name:           ActivateTabGraphviz
        Case HelpAttributesSheet.name:  ActivateTabLaunchpad
        Case HelpColorsSheet.name:      ActivateTabLaunchpad
        Case HelpShapesSheet.name:      ActivateTabLaunchpad
        Case LocaleDeDeSheet.name:      ActivateTabLaunchpad
        Case LocaleEnGbSheet.name:      ActivateTabLaunchpad
        Case LocaleEnUsSheet.name:      ActivateTabLaunchpad
        Case LocaleFrFrSheet.name:      ActivateTabLaunchpad
        Case LocaleItItSheet.name:      ActivateTabLaunchpad
        Case LocalePlPlSheet.name:      ActivateTabLaunchpad
        Case SettingsSheet.name:        ActivateTabLaunchpad
        Case StyleDesignerSheet.name:   ActivateTabStyleDesigner
        Case StylesSheet.name:          ActivateTabStyles
        Case SourceSheet.name:          ActivateTabSource
        Case SqlSheet.name:             ActivateTabSql
        Case SvgSheet.name:             ActivateTabSvg
        Case Else:                      ActivateTabGraphviz
    End Select
             
    Exit Sub

ErrorHandler:
    If Err.number > 0 Then
        Err.Clear
        Resume Next
    End If
End Sub

Public Sub ribbon_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case control.ID
        Case RIBBON_TAB_STYLE_DESIGNER
            visible = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_STYLE_DESIGNER)
        Case RIBBON_TAB_STYLES
            visible = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_STYLES)
        Case RIBBON_TAB_ABOUT
            visible = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_ABOUT)
        Case RIBBON_TAB_CONSOLE
            visible = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_CONSOLE)
        Case RIBBON_TAB_DIAGNOSTICS
            visible = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_DIAGNOSTICS)
        Case RIBBON_TAB_EXCHANGE
            visible = GetSettingBoolean(SETTINGS_TABS_TOGGLE_EXCHANGE)
        Case RIBBON_TAB_SOURCE
            visible = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_SOURCE)
        Case RIBBON_TAB_SQL
            visible = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_SQL)
        Case RIBBON_TAB_SVG
            visible = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_SVG)
        Case Else
            visible = True
    End Select
End Sub

Public Sub RefreshRibbon()
    On Error GoTo ErrorHandler
    If myRibbon Is Nothing Then
        ' This message cannot be localized due to error state.
        MsgBox "Error refreshing the ribbon. Save and reopen this file."
    Else
        myRibbon.Invalidate
        If Err.number <> 0 Then
            ' This message cannot be localized due to error state.
            MsgBox "Lost the Ribbon object. Save this file, close worksbook, and reopen."
        End If
    End If

    Exit Sub

ErrorHandler:
    If Err.number > 0 Then
        Err.Clear
        Resume Next
    End If
End Sub

Public Sub InvalidateRibbonControl(ByVal controlName As String)
    On Error GoTo ErrorHandler
    If myRibbon Is Nothing Then
        ' This message cannot be localized due to error state.
        UpdateStatusBar replace("Error updating the ribbon for control named '{controlName}'. Save and reopen this file.", "{controlName}", controlName)
    Else
        myRibbon.InvalidateControl controlName
    End If
ErrorHandler:
    If Err.number > 0 Then
        Err.Clear
        Resume Next
    End If
End Sub

Public Sub ActivateTab(ByVal tabName As String)
    On Error GoTo ErrorHandler
    If myRibbon Is Nothing Then
        ' This message cannot be localized due to error state.
        UpdateStatusBar replace("Error activating a ribbon tab named '{tabName}'. Save and reopen this file.", "{tabName}", tabName)
    Else
        myRibbon.ActivateTab tabName
    End If
ErrorHandler:
    If Err.number > 0 Then
        Err.Clear
        Resume Next
    End If
End Sub

' ===========================================================================
' Ribbon Callbacks for prefixed ribbon buttons

Public Sub button_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetSettingBoolean(control.ID & BUTTON_SUFFIX_VISIBLE)
End Sub

Public Sub button_getLabel(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = SettingsSheet.Range(control.ID & BUTTON_SUFFIX_TEXT).value
End Sub

Public Sub button_getScreentip(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = SettingsSheet.Range(control.ID & BUTTON_SUFFIX_SCREENTIP).value
End Sub

Public Sub button_getSupertip(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = SettingsSheet.Range(control.ID & BUTTON_SUFFIX_SUPERTIP).value
End Sub

' ===========================================================================
' Methods for activating tabs asynchronously

Public Sub ActivateTabSql()
    ActivateTab RIBBON_TAB_SQL
End Sub

Public Sub ActivateTabSource()
    ActivateTab RIBBON_TAB_SOURCE
End Sub

Public Sub ActivateTabConsole()
    ActivateTab RIBBON_TAB_CONSOLE
End Sub

Public Sub ActivateTabSvg()
    ActivateTab RIBBON_TAB_SVG
End Sub

Public Sub ActivateTabDiagnostics()
    ActivateTab RIBBON_TAB_DIAGNOSTICS
End Sub

Public Sub ActivateTabStyleDesigner()
    ActivateTab RIBBON_TAB_STYLE_DESIGNER
End Sub

Public Sub ActivateTabStyles()
    ActivateTab RIBBON_TAB_STYLES
End Sub

Public Sub ActivateTabAbout()
    ActivateTab RIBBON_TAB_ABOUT
End Sub

Public Sub ActivateTabExchange()
    ActivateTab RIBBON_TAB_EXCHANGE
End Sub

Public Sub ActivateTabLaunchpad()
    ActivateTab RIBBON_TAB_WORKSHEETS
End Sub

Public Sub ActivateTabGraphviz()
    ActivateTab RIBBON_TAB_GRAPHVIZ
End Sub

