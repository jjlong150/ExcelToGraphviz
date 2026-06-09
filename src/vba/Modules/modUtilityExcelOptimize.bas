Attribute VB_Name = "modUtilityExcelOptimize"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modUtilityExcelOptimize
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Utility / Excel Performance
'
' ROLE:
'   Lightweight performance-tuning helpers for temporarily disabling expensive
'   Excel behaviors during bulk operations, ensuring faster and more predictable
'   execution of worksheet-driven workflows.
'
' RESPONSIBILITIES:
'   - OptimizeCode_Begin:
'       o Disable ScreenUpdating to prevent UI redraws
'       o Disable events to avoid triggering Worksheet_Change/SelectionChange
'   - OptimizeCode_End:
'       o Restore events and screen updating to their prior state
'
' ARCHITECTURAL NOTES:
'   - Designed for short-lived performance windows around high-volume cell
'     operations, SQL output, or shape creation.
'   - Does not modify Calculation mode; callers retain full control over
'     calculation semantics.
'   - Safe for both Windows and macOS.
'
' USAGE:
'   - Wrap bulk operations:
'         OptimizeCode_Begin
'         ' ... heavy work ...
'         OptimizeCode_End
'
' RELATED WIKI PAGES:
'   - Performance Guidelines
'   - Worksheet Automation Patterns
' =============================================================================

Option Explicit

' Public routines to turn code optimizations on and off
Public Sub OptimizeCode_Begin()
    Application.ScreenUpdating = False
    Application.enableEvents = False
End Sub

Public Sub OptimizeCode_End()
    Application.enableEvents = True
    Application.ScreenUpdating = True
End Sub



