VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListsSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
' PROJECT:          Excel to Graphviz
' WORKSHEET MODULE: ListsSheet
' COPYRIGHT:        Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:            Relationship Visualizer / Sheets / Internal
'
' ROLE:
'   Hidden worksheet hosting the static and dynamic lists used throughout the
'   application for Settings dropdowns, Data Validation, and platform-specific
'   configuration (e.g., macOS font enumeration).
'
' DESCRIPTION:
'   This worksheet module manages the interactive behavior of the Lists sheet,
'   which acts as the authoritative source for system-level dropdown lists.
'   When activated, it switches the Ribbon to the Launchpad tab for worksheet
'   navigation and administrative tools. The sheet itself stores lists such as
'   ExcludedFonts, macOS Fonts, FontSizes, and other validation sources used by
'   the Settings sheet, Style Designer, and Data worksheet.
'
' KEY RESPONSIBILITIES:
'   - CONTEXT-AWARE NAVIGATION:
'       o Activates the Launchpad Ribbon tab when the sheet is selected
'       o Ensures worksheet-level utilities and settings tools are accessible
'
'   - CONFIGURATION AUTHORITY:
'       o Hosts static and dynamic lists for Settings dropdowns
'       o Stores macOS font lists used by GetMacFontList and fallback logic
'       o Provides validation sources for Style Designer and Data sheet controls
'
'   - PLATFORM-SAFE ACTIVATION:
'       o Uses Application.OnTime on macOS to avoid Ribbon race conditions
'       o Performs immediate activation on Windows for seamless transitions
'
' USAGE:
'   - Fires automatically when navigating to the Lists sheet
'   - Provides the underlying data sources for dropdowns across the workbook
'
' RELATED WIKI PAGES:
'   - Settings Worksheet Architecture
'   - Font Enumeration (Windows & macOS)
'   - Data Validation & List Management
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: Worksheet_Activate
' PURPOSE:
'   Synchronizes the Ribbon UI when navigating the system list repository.
'
' TECHNICAL WORKFLOW:
'   1. UI NAVIGATION: Triggers 'ActivateTabLaunchpad' to provide access
'      to the global 'Worksheets' and 'Settings' controls.
'   2. PLATFORM ADAPTATION: Handles Mac-specific timing via 'Application.OnTime'.
' ==========================================================================
Private Sub Worksheet_Activate()
#If Mac Then
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabLaunchpad"
#Else
    ActivateTabLaunchpad
#End If
End Sub

