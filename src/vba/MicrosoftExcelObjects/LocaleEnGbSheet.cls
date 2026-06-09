VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocaleEnGbSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
' PROJECT:          Excel to Graphviz
' WORKSHEET MODULE: LocaleEnGbSheet
' COPYRIGHT:        Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:            Relationship Visualizer / Sheets / Localization
'
' ROLE:
'   UK English language resource sheet containing key-value translation pairs
'   consumed by the localization engine. Provides the authoritative source
'   for UI labels, messages, tooltips, and Ribbon text when UK English is the
'   active application language.
'
' DESCRIPTION:
'   This worksheet module manages the interactive behavior of the UK English
'   localization sheet. When activated, it switches the Ribbon to the
'   Launchpad tab for worksheet-level utilities and uses macOS-safe deferred
'   activation to avoid Ribbon timing issues. The sheet itself serves as a
'   structured lookup table for modLocalize, supplying translated strings
'   for GetLabel, GetMessage, and other localization calls.
'
' KEY RESPONSIBILITIES:
'   - CONTEXT-AWARE NAVIGATION:
'       o Activates the Launchpad Ribbon tab when the sheet is selected
'       o Ensures worksheet navigation and settings tools are accessible
'
'   - TRANSLATION AUTHORITY:
'       o Hosts the UK English key-value pairs used by modLocalize
'       o Supplies translated UI labels, messages, and tooltips
'       o Supports dynamic language switching across the application
'
'   - PLATFORM-SAFE ACTIVATION:
'       o Uses Application.OnTime on macOS to avoid Ribbon race conditions
'       o Performs immediate activation on Windows for seamless transitions
'
' USAGE:
'   - Fires automatically when navigating to the UK English localization sheet
'   - Provides the authoritative translation source for the UK English UI
'
' RELATED WIKI PAGES:
'   - Localization Framework
'   - Language Packs & Translation Keys
'   - Worksheet Navigation & Launchpad
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: Worksheet_Activate
' PURPOSE:
'   Synchronizes the Ribbon UI with the administrative localization workspace.
'
' TECHNICAL WORKFLOW:
'   1. UI NAVIGATION: Triggers 'ActivateTabLaunchpad' to ensure access
'      to global 'Worksheets' and 'Settings' controls.
'   2. PLATFORM ADAPTATION:
'      - macOS: Uses 'Application.OnTime' (1s delay) for stable UI activation.
'      - Windows: Activates the tab immediately for a seamless transition.
' ==========================================================================
Private Sub Worksheet_Activate()
#If Mac Then
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabLaunchpad"
#Else
    ActivateTabLaunchpad
#End If
End Sub

