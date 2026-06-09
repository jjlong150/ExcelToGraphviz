Attribute VB_Name = "modUtilityExcelStatusBar"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modUtilityStatusBar
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Excel UI / Feedback
'
' ROLE:
'   Minimal status-bar messaging helpers for providing transient, low-overhead
'   user feedback during long-running or multi-step operations.
'
' RESPONSIBILITIES:
'   - UpdateStatusBar:
'       o Set Application.StatusBar to a caller-supplied message
'       o Yield control via DoEvents to ensure immediate UI update
'   - UpdateStatusBarForNSeconds:
'       o Display a message for a fixed duration using Application.OnTime
'       o Automatically schedule ClearStatusBar
'   - ClearStatusBar:
'       o Restore Excel's native status bar behavior
'
' ARCHITECTURAL NOTES:
'   - Uses Application.StatusBar = False to return control to Excel.
'   - DoEvents ensures prompt repainting even during heavy operations.
'   - OnTime scheduling avoids blocking the caller's execution flow.
'   - Consumed by SQL engine, file operations, Ribbon callbacks, and
'     long-running workflows requiring lightweight progress cues.
'
' USAGE:
'   - Ideal for progress messages, transient notifications, and
'     non-modal user feedback during automation.
'
' RELATED WIKI PAGES:
'   - UI Feedback & Status Bar Conventions
'   - Long-Running Operation Patterns
' =============================================================================

Option Explicit

Public Sub UpdateStatusBar(ByVal statusMessage As String)
    Application.StatusBar = statusMessage
    DoEvents
End Sub

Public Sub UpdateStatusBarForNSeconds(ByVal statusMessage As String, ByVal seconds As Long)
    Application.StatusBar = statusMessage
    DoEvents
    Application.OnTime Now + TimeSerial(0, 0, seconds), "ClearStatusBar"
End Sub

Public Sub ClearStatusBar()
    Application.StatusBar = False
    DoEvents
End Sub

