VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SettingsSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
' PROJECT:          Excel to Graphviz
' WORKSHEET MODULE: SettingsSheet
' COPYRIGHT:        Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:            Relationship Visualizer / Sheets / Settings
'
' ROLE:
'   Administrative control surface for all global configuration. Hosts the
'   Named Ranges, extension metadata, file-output settings, and UI-definition
'   values consumed by the rendering engine, Ribbon callbacks, SQL subsystem,
'   and Style Designer.
'
' DESCRIPTION:
'   This worksheet module governs the interactive behavior of the Settings
'   sheet—the authoritative storage location for every GetSetting call. It
'   manages Ribbon synchronization, monitors changes to key configuration
'   ranges, and provides the metadata that powers user-defined extensions
'   (SourceWeb, ExtCode, ExtWeb). The module also enforces robust event
'   hygiene, ensuring that Ribbon refreshes never leave Excel in an unstable
'   state.
'
' KEY RESPONSIBILITIES:
'   - CONTEXT-AWARE NAVIGATION:
'       • Activates the Launchpad Ribbon tab when the sheet is selected
'       • Ensures administrative tools and worksheet navigation are visible
'
'   - RIBBON SYNCHRONIZATION:
'       • Monitors Named Ranges tied to output paths, filenames, indentation,
'         and extension definitions
'       • Calls RefreshRibbon to invalidate and redraw the Ribbon UI
'
'   - EXTENSION ARCHITECTURE:
'       • Hosts metadata for user-defined web links (SourceWeb1–6)
'       • Hosts metadata for user-defined macro triggers (ExtCode1–6)
'       • Hosts metadata for user-defined external web actions (ExtWeb1–6)
'       • Enables no-code extensibility of the Ribbon
'
'   - NAMED RANGE AUTHORITY:
'       • Stores all configuration values retrieved by GetSetting*
'       • Provides the physical backing for Settings-driven worksheet layouts
'
'   - ERROR-HARDENED EVENT MANAGEMENT:
'       • Uses a recovery block to guarantee EnableEvents is restored
'       • Logs unexpected failures to the Console for diagnostics
'
' USAGE:
'   - Fires automatically when users modify any configuration value
'   - Ensures Ribbon state always reflects the latest Settings
'   - Provides the authoritative configuration source for the entire system
'
' RELATED WIKI PAGES:
'   - Settings Worksheet Architecture
'   - Extension Buttons (SourceWeb / ExtCode / ExtWeb)
'   - Ribbon Refresh & IRibbonUI.Invalidate
'   - Named Range API & Configuration Retrieval
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: Worksheet_Activate
' PURPOSE:
'   Synchronizes the Ribbon UI with the global settings workspace.
'
' TECHNICAL WORKFLOW:
'   1. UI NAVIGATION: Triggers 'ActivateTabLaunchpad' to provide immediate
'      access to the 'Worksheets' and 'Settings' ribbon groups.
'   2. PLATFORM ADAPTATION:
'      - macOS: Uses 'Application.OnTime' (1s delay) to prevent UI race
'        conditions in the Mac sandbox.
'      - Windows: Activates the tab immediately for zero-latency navigation.
' ==========================================================================
Private Sub Worksheet_Activate()
#If Mac Then
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabLaunchpad"
#Else
    ActivateTabLaunchpad
#End If
End Sub

' ==========================================================================
' PROCEDURE: Worksheet_Change
' PURPOSE:
'   Monitors key configuration cells and synchronizes the Ribbon UI.
'
' TECHNICAL WORKFLOW:
'   1. CHANGE DETECTION: Uses 'Intersect' to identify if the edited cell
'      belongs to specific Named Ranges (Output Paths, File Names, or
'      Custom Extension Button definitions).
'   2. UI REFRESH: If a match is found, it calls 'RefreshRibbon', which
'      triggers the IRibbonUI.Invalidate method to redraw the interface.
'   3. EVENT SUPPRESSION: Disables 'Application.EnableEvents' during the
'      refresh to prevent recursive loops or performance lag.
'   4. ERROR RECOVERY: Implements a 'Whoa' error trap to ensure that
'      unexpected issues log a message to the Console while always
'      restoring the 'EnableEvents' state.
'
' USAGE:
'   - Allows users to rename or toggle custom "Extension" buttons and
'     see the results in the Ribbon without restarting the workbook.
' ==========================================================================
Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo Whoa

    ' Detect changes in the settings dropdown lists
    If Not Intersect(Target, SettingsSheet.Range(SETTINGS_OUTPUT_DIRECTORY)) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
        
    ElseIf Not Intersect(Target, SettingsSheet.Range(SETTINGS_FILE_NAME)) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range(SETTINGS_SOURCE_INDENT)) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
        
    ElseIf Not Intersect(Target, SettingsSheet.Range("SourceWeb1Text,SourceWeb1Visible,SourceWeb1URL,SourceWeb1Screentip,SourceWeb1Supertip")) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("SourceWeb2Text,SourceWeb2Visible,SourceWeb2URL,SourceWeb2Screentip,SourceWeb2Supertip")) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("SourceWeb3Text,SourceWeb3Visible,SourceWeb3URL,SourceWeb3Screentip,SourceWeb3Supertip")) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("SourceWeb4Text,SourceWeb4Visible,SourceWeb4URL,SourceWeb4Screentip,SourceWeb4Supertip")) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("SourceWeb5Text,SourceWeb5Visible,SourceWeb5URL,SourceWeb5Screentip,SourceWeb5Supertip")) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon

    ElseIf Not Intersect(Target, SettingsSheet.Range("SourceWeb6Text,SourceWeb6Visible,SourceWeb6URL,SourceWeb6Screentip,SourceWeb6Supertip")) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon

    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtCode1Text,ExtCode1Visible,ExtCode1Sub,ExtCode1Screentip,ExtCode1Supertip")) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtCode2Text,ExtCode2Visible,ExtCode2Sub,ExtCode2Screentip,ExtCode2Supertip")) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtCode3Text,ExtCode3Visible,ExtCode3Sub,ExtCode3Screentip,ExtCode3Supertip")) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtCode4Text,ExtCode4Visible,ExtCode4Sub,ExtCode4Screentip,ExtCode4Supertip")) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtCode5Text,ExtCode5Visible,ExtCode5Sub,ExtCode5Screentip,ExtCode5Supertip")) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtCode6Text,ExtCode6Visible,ExtCode6Sub,ExtCode6Screentip,ExtCode6Supertip")) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtWeb1Text,ExtWeb1Visible,ExtWeb1URL,ExtWeb1Screentip,ExtWeb1Supertip")) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtWeb2Text,ExtWeb2Visible,ExtWeb2URL,ExtWeb2Screentip,ExtWeb2Supertip")) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtWeb3Text,ExtWeb3Visible,ExtWeb3URL,ExtWeb3Screentip,ExtWeb3Supertip")) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtWeb4Text,ExtWeb4Visible,ExtWeb4URL,ExtWeb4Screentip,ExtWeb4Supertip")) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtWeb5Text,ExtWeb5Visible,ExtWeb5URL,ExtWeb5Screentip,ExtWeb5Supertip")) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
    
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtWeb6Text,ExtWeb6Visible,ExtWeb6URL,ExtWeb6Screentip,ExtWeb6Supertip")) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
   
    ElseIf Not Intersect(Target, SettingsSheet.Range("ExtTabName,ExttabGroupNameCode,ExtTabGroupNameWeb")) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
    End If
    
LetsContinue:
    Application.enableEvents = True
    Exit Sub
Whoa:
    EmitMessage err.number & " " & err.Description
    
    Resume LetsContinue

End Sub

