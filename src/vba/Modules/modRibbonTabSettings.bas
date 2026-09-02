Attribute VB_Name = "modRibbonTabSettings"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modRibbonTabSettings
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Excel UI / Ribbon Callbacks
'
' ROLE:
'   Ribbon callback module for the Settings Tab. Manages navigation between
'   settings sub-pages, synchronizes toggle-button state, and provides platform-
'   specific visibility logic. Acts as the UI controller for workbook-wide
'   configuration panels.
'
' RESPONSIBILITIES:
'   - Maintain the currently selected settings toggleButton (CurrentSettingsID).
'   - Manage the full set of Settings Tab toggleButtons and invalidate them
'     collectively to enforce radio-button behavior.
'   - Route user selections to the appropriate Settings sub-page:
'       o Command Line Options
'       o Graph Options
'       o Console Tab
'       o Data Tab
'       o Graphviz Tab
'       o Help URLs
'       o Source Tab
'       o Exchange Tab
'       o Extensions Tab
'       o Launchpad Tab
'       o Worksheet-specific settings (Data, Styles, Source, SQL, Knowledge)
'       o SQL Keywords
'       o SVG Tab
'   - Provide macOS-specific visibility rules for controls that differ by platform.
'   - Handle Settings Tab help navigation via workbook-configured URLs.
'
' INTERACTIONS:
'   - Named Ranges: HelpURLSettingsTab, worksheet/tab selector ranges.
'   - Modules: TabSelect* procedures (Settings navigation handlers),
'     modUtilityRibbon, modUtilitySettings.
'   - Sheets: SettingsSheet.
'
' CROSS-PLATFORM NOTES:
'   - macOS visibility toggles hide or show platform-specific controls.
'
' ERROR HANDLING:
'   - Uses Ribbon invalidation to maintain consistent toggle state across
'     all settings categories.
'
' RELATED WIKI PAGES:
'   - Settings Tab
'   - Workbook Configuration Architecture
'   - Platform-Specific UI Behavior (Windows/macOS)
' =============================================================================

Option Explicit

' Stores the ID of the currently selected settings toggleButton
Public CurrentSettingsID As String
Public SettingsButtons As Variant

Private Sub InitSettingsButtons()
    SettingsButtons = Array( _
        "settingsCommandOptions", _
        "settingsGraphOptions", _
        "settingsConsoleTab", _
        "settingsDataWorksheet", _
        "settingsDataTab", _
        "settingsGraphvizTab", _
        "settingsExchangeTab", _
        "settingsExtensionsTab", _
        "settingsHelpUrls", _
        "settingsKnowledgeWorksheet", _
        "settingsLaunchpadTab", _
        "settingsSourceWorksheet", _
        "settingsSourceTab", _
        "settingsSqlWorksheet", _
        "settingsSqlTab", _
        "settingsSqlKeywords", _
        "settingsStylesWorksheet", _
        "settingsStylesTab", _
        "settingsSVGTab" _
    )
End Sub

Private Sub InvalidateSettingsButtons()

    ' Lazy initialization
    If IsEmpty(SettingsButtons) Then InitSettingsButtons
    
    Dim id As Variant
    For Each id In SettingsButtons
        InvalidateRibbonControl id
    Next id
End Sub

' ============================================================
'  settings_onAction
' ============================================================
Private Sub settings_onAction(control As IRibbonControl, pressed As Boolean)

    SettingsSheet.Activate
    
    If pressed Then
        ' Update the selected control ID
        CurrentSettingsID = control.id

        ' Refresh all toggleButtons
        InvalidateSettingsButtons
        
        ' Invoke your action
        Select Case control.id

            Case "settingsConsoleTab"
                TabSelectConsoleTab
                
            Case "settingsCommandOptions"
                TabSelectCmdLineOptions

            Case "settingsGraphOptions"
                TabSelectGraphOptions

            Case "settingsDataTab"
                TabSelectDataTab

            Case "settingsGraphvizTab"
                TabSelectGraphvizTab

            Case "settingsHelpUrls"
                TabSelectHelpUrls
            
            Case "settingsSourceTab"
                TabSelectSourceTab

            Case "settingsExchangeTab"
                TabSelectExchangeTab

            Case "settingsExtensionsTab"
                TabSelectExtensionsTab

            Case "settingsDataWorksheet"
                TabSelectDataWorksheet

            Case "settingsLaunchpadTab"
                TabSelectLaunchpadTab

            Case "settingsSourceWorksheet"
                TabSelectSourceWorksheet

            Case "settingsStylesWorksheet"
                TabSelectStylesWorksheet

            Case "settingsStylesTab"
                TabSelectStylesTab

            Case "settingsSqlWorksheet"
                TabSelectSqlWorksheet

            Case "settingsSqlTab"
                TabSelectSqlTab

            Case "settingsSqlKeywords"
                TabSelectSqlKeywords

            Case "settingsSVGTab"
                TabSelectSVGTab

        End Select
    End If

End Sub

Private Sub macHide_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
#If Mac Then
    returnedVal = False
#Else
    returnedVal = True
#End If
End Sub

Private Sub macShow_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
#If Mac Then
    returnedVal = True
#Else
    returnedVal = False
#End If
End Sub

' ============================================================
'  settings_getPressed
' ============================================================
Private Sub settings_getPressed(ByVal control As IRibbonControl, ByRef pressed As Variant)
    pressed = (control.id = CurrentSettingsID)
End Sub

' ===========================================================================
' Callbacks for Help

Private Sub settingsHelp_onAction(ByVal control As IRibbonControl)
    ActiveWorkbook.FollowHyperlink Address:=SettingsSheet.Range("HelpURLSettingsTab").value, NewWindow:=True
End Sub


