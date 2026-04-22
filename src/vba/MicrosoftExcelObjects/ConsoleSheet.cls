VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConsoleSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
' PROJECT:          Excel to Graphviz
' WORKSHEET MODULE: ConsoleSheet
' COPYRIGHT:        Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:            Relationship Visualizer / Sheets / Diagnostics
'
' ROLE:
'   Event-driven UI host for the "console" worksheet. Provides contextual Ribbon
'   activation, platform-safe initialization, and a stable destination for
'   Graphviz, SQL, and application-level diagnostic output.
'
' DESCRIPTION:
'   This worksheet module manages the interactive behavior of the Console
'   environment. When activated, it switches the Ribbon to the Console tab,
'   ensures macOS-safe deferred activation, and prepares the sheet to receive
'   diagnostic messages. All logging logic is handled elsewhere; this module
'   focuses exclusively on UI context, activation behavior, and integration
'   with the logging pipeline.
'
' KEY RESPONSIBILITIES:
'   - CONTEXT-AWARE NAVIGATION:
'       • Activates the Console Ribbon tab when the sheet is selected
'       • Surfaces log-management tools (Clear Log, Export, Verbose Mode)
'
'   - LOGGING TARGET:
'       • Serves as the destination for EmitMessage, EmitDiagnostic, and
'         Graphviz console output captured during rendering
'
'   - PLATFORM-SAFE ACTIVATION:
'       • Uses Application.OnTime on macOS to avoid Ribbon race conditions
'       • Performs immediate activation on Windows for seamless transitions
'
'   - UI STATE MANAGEMENT:
'       • Ensures Console-specific Ribbon controls remain synchronized with
'         the active worksheet context
'
' USAGE:
'   - Fires automatically when the user selects the Console worksheet
'   - Provides a stable, always-visible destination for diagnostic output
'
' RELATED WIKI PAGES:
'   - Logging & Diagnostics Pipeline
'   - ExecuteAndCapture Architecture
'   - Console Ribbon Tools
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: Worksheet_Activate
' PURPOSE:
'   Synchronizes the Ribbon UI with the 'Console' diagnostic worksheet.
'
' TECHNICAL WORKFLOW:
'   1. UI NAVIGATION: Triggers 'ActivateTabConsole' to switch the Ribbon
'      focus to the log management and diagnostic tools.
'   2. PLATFORM ADAPTATION:
'      - macOS: Employs 'Application.OnTime' to introduce a one-second delay,
'        ensuring UI stability within the Mac sandbox environment.
'      - Windows: Executes the tab switch instantly for a responsive feel.
'
' USAGE:
'   - Fires automatically when the user selects the 'Console' tab.
'   - Facilitates rapid troubleshooting by bringing 'Clear Log' and 'Export'
'     actions to the forefront.
' ==========================================================================
Private Sub Worksheet_Activate()
#If Mac Then
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabConsole"
#Else
    ActivateTabConsole
#End If
End Sub

