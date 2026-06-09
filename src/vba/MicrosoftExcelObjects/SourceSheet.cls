VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SourceSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
' PROJECT:          Excel to Graphviz
' WORKSHEET MODULE: SourceSheet
' COPYRIGHT:        Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:            Relationship Visualizer / Sheets / Source
'
' ROLE:
'   Interactive editing surface for generated DOT source. Provides contextual
'   Ribbon activation, automated line-number maintenance, and a stable entry
'   point for manual DOT refinement, previewing, and clipboard export.
'
' DESCRIPTION:
'   This worksheet module governs the behavior of the DOT Source Viewer sheet.
'   When activated, it switches the Ribbon to the Source tab, resets the view
'   to the top of the code block, and ensures a consistent editing context.
'   Any manual edits trigger automatic line-number regeneration, wrapped in
'   performance-optimized update blocks to avoid flicker and recalculation lag.
'
' KEY RESPONSIBILITIES:
'   - CONTEXT-AWARE NAVIGATION:
'       o Activates the Source Ribbon tab when the sheet is selected
'       o Surfaces Copy, Preview, Wrap, and Export controls
'
'   - AUTOMATED LINE INDEXING:
'       o Calls UpdateSourceWorksheetLineNumbers on Worksheet_Change
'       o Maintains the left-hand gutter for readability and debugging
'
'   - PERFORMANCE OPTIMIZATION:
'       o Wraps updates in OptimizeCode_Begin / OptimizeCode_End
'       o Prevents screen flicker and unnecessary recalculation
'
'   - UX HYGIENE:
'       o Resets selection to A1 on activation for a consistent starting point
'
' USAGE:
'   - Fires automatically when navigating to the Source worksheet
'   - Provides the primary environment for inspecting and editing DOT code
'
' RELATED WIKI PAGES:
'   - DOT Source Viewer & Console Architecture
'   - Line-Numbering Subsystem
'   - Manual DOT Overrides & Preview Pipeline
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: Worksheet_Activate
' PURPOSE:
'   Synchronizes the Ribbon UI and resets the code viewing area.
'
' TECHNICAL WORKFLOW:
'   1. UI NAVIGATION: Triggers 'ActivateTabSource' to surface DOT-specific
'      utilities (Copy, Preview, Wrap).
'   2. PLATFORM ADAPTATION: Handles Mac-specific timing via 'Application.OnTime'.
'   3. CURSOR RESET: Selects cell 'A1' to ensure the user starts at the
'      beginning of the code block.
' ==========================================================================
Private Sub Worksheet_Activate()
#If Mac Then
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabSource"
#Else
    ActivateTabSource
#End If
    ActiveSheet.Range("A1").Select
End Sub

' ==========================================================================
' PROCEDURE: Worksheet_Change
' PURPOSE:
'   Maintains the integrity of the line-numbering gutter.
'
' TECHNICAL WORKFLOW:
'   1. STATE LOCKDOWN: Disables screen updates/events via 'OptimizeCode_Begin'.
'   2. INDEX REFRESH: Calls 'UpdateSourceWorksheetLineNumbers' to re-calculate
'      the sequential numbering in the left-hand gutter.
'   3. STATE RESTORATION: Re-enables Excel interaction via 'OptimizeCode_End'.
' ==========================================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    OptimizeCode_Begin
    UpdateSourceWorksheetLineNumbers
    OptimizeCode_End
End Sub

