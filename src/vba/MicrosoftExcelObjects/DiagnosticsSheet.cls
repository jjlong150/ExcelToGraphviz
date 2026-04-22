VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DiagnosticsSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
' PROJECT:          Excel to Graphviz
' WORKSHEET MODULE: DiagnosticsSheet
' COPYRIGHT:        Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:            Relationship Visualizer / Sheets / Diagnostics
'
' ROLE:
'   Event-driven UI host for the "diagnostics" worksheet. Generates real-time
'   environment reports, activates the Diagnostics Ribbon tab, and ensures
'   system-sensitive information is cleared when the user navigates away.
'
' DESCRIPTION:
'   This worksheet module provides the interactive layer for the Diagnostics
'   subsystem. When activated, it triggers a full environment audit—capturing
'   OS details, Excel version, provider state, file paths, and other runtime
'   metadata. It also manages Ribbon activation and uses macOS-safe deferred
'   scheduling to avoid UI race conditions. Upon deactivation, it clears all
'   diagnostic content to prevent stale or sensitive system data from being
'   saved in the workbook.
'
' KEY RESPONSIBILITIES:
'   - REAL-TIME ENVIRONMENT REPORTING:
'       • Calls ReportDiagnostics on activation to populate system metadata
'       • Surfaces provider state, file paths, and environment fingerprints
'
'   - CONTEXT-AWARE NAVIGATION:
'       • Activates the Diagnostics Ribbon tab when the sheet is selected
'       • Ensures diagnostic tools are immediately accessible
'
'   - PLATFORM-SAFE ACTIVATION:
'       • Uses Application.OnTime on macOS to avoid Ribbon race conditions
'       • Performs immediate activation on Windows for seamless transitions
'
'   - PRIVACY & DATA HYGIENE:
'       • Calls ClearDiagnostics on deactivation to remove system-specific data
'       • Prevents environment fingerprints from persisting in saved files
'
' USAGE:
'   - Fires automatically when the user selects the Diagnostics worksheet
'   - Ensures environment data is always fresh and never persisted unnecessarily
'
' RELATED WIKI PAGES:
'   - Diagnostics & Environment Reporting
'   - ADO Provider Detection
'   - Image Cache Management
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: Worksheet_Activate
' PURPOSE:
'   Triggers a real-time system audit and synchronizes the Ribbon UI.
'
' TECHNICAL WORKFLOW:
'   1. DATA GENERATION: Immediately calls 'ReportDiagnostics' to populate
'      the sheet with current performance metrics and environment paths.
'   2. UI NAVIGATION: Switches the Ribbon focus to the 'Diagnostics' tab.
'   3. PLATFORM ADAPTATION:
'      - macOS: Uses 'Application.OnTime' (1s delay) for stable UI activation.
'      - Windows: Executes activation instantly for a seamless transition.
' ==========================================================================
Private Sub Worksheet_Activate()
    ReportDiagnostics
#If Mac Then
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabDiagnostics"
#Else
    ActivateTabDiagnostics
#End If
End Sub

' ==========================================================================
' PROCEDURE: Worksheet_Deactivate
' PURPOSE:
'   Ensures system-sensitive information is purged from the workbook.
'
' TECHNICAL WORKFLOW:
'   1. DATA HYGIENE: Invokes 'ClearDiagnostics' to wipe the worksheet content.
'   2. PRIVACY PROTECTION: Prevents system-specific data (like local file
'      paths or memory states) from being saved permanently in the file.
' ==========================================================================
Private Sub Worksheet_Deactivate()
    ClearDiagnostics
End Sub

