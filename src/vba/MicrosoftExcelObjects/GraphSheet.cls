VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraphSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
' PROJECT:          Excel to Graphviz
' WORKSHEET MODULE: GraphSheet
' COPYRIGHT:        Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:            Relationship Visualizer / Sheets / Output
'
' ROLE:
'   Dedicated viewing surface for rendered diagrams. Manages Ribbon context,
'   status-bar hygiene, and platform-safe activation to provide a clean,
'   distraction-free environment for inspecting generated graphs.
'
' DESCRIPTION:
'   This worksheet module governs the interactive behavior of the Graph sheet.
'   When activated, it switches the Ribbon to the Graphviz tab, clears any
'   lingering status messages, and prepares the workspace for visual review.
'   It does not participate in rendering logic itself; instead, it provides
'   the UI context for diagrams produced by the rendering engine.
'
' KEY RESPONSIBILITIES:
'   - CONTEXT-SENSITIVE UI:
'       o Activates the Graphviz Ribbon tab when the sheet is selected
'       o Ensures zoom, export, and redraw tools are immediately available
'
'   - UI HYGIENE:
'       o Clears the StatusBar on entry and exit to remove stale messages
'       o Maintains a clean viewing environment for rendered diagrams
'
'   - PLATFORM-SAFE ACTIVATION:
'       o Uses Application.OnTime on macOS to avoid Ribbon race conditions
'       o Performs immediate activation on Windows for seamless transitions
'
' USAGE:
'   - Fires automatically when the user navigates to the Graph worksheet
'   - Provides the optimized workspace for inspecting rendered diagrams
'
' RELATED WIKI PAGES:
'   - Graph Viewing & Export Tools
'   - Rendering Pipeline Overview
'   - Ribbon Navigation Architecture
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: Worksheet_Activate
' PURPOSE:
'   Optimizes the workspace for graph viewing and Ribbon tool access.
'
' TECHNICAL WORKFLOW:
'   1. UI NAVIGATION: Triggers 'ActivateTabGraphviz' to ensure the user has
'      immediate access to scaling and export controls.
'   2. PLATFORM ADAPTATION: Handles Mac-specific timing via 'Application.OnTime'.
'   3. STATUS CLEANUP: Calls 'ClearStatusBar' to remove lingering processing
'      messages (e.g., "Rendering...").
' ==========================================================================
Private Sub Worksheet_Activate()
#If Mac Then
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabGraphviz"
#Else
    ActivateTabGraphviz
#End If
    ClearStatusBar
End Sub

' ==========================================================================
' PROCEDURE: Worksheet_Deactivate
' PURPOSE:
'   Resets the application peripheral UI when moving away from the graph.
'
' TECHNICAL WORKFLOW:
'   1. STATUS RESET: Calls 'ClearStatusBar' to ensure the next workspace
'      starts with a clean status area.
' ==========================================================================
Private Sub Worksheet_Deactivate()
    ClearStatusBar
End Sub
