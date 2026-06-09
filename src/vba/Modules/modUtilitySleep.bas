Attribute VB_Name = "modUtilitySleep"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modUtilitySleep
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Utility / Timing & Execution Control
'
' ROLE:
'   Provide a lightweight, API-free millisecond sleep routine for Windows,
'   allowing controlled pacing of loops, UI updates, and COM-sensitive
'   operations without introducing external dependencies.
'
' RESPONSIBILITIES:
'   - SleepMilliseconds:
'       o Implement a busy-wait loop using Timer + DoEvents
'       o Avoid Win32 Sleep API to maintain macro-security compatibility
'       o Provide predictable millisecond-scale delays for throttling
'
' ARCHITECTURAL NOTES:
'   - Windows-only implementation (Timer resolution differs on macOS).
'   - DoEvents prevents Excel from appearing frozen during the delay.
'   - Useful for COM-reentrancy mitigation, animation pacing, and controlled
'     retry loops in file or process polling.
'
' VERSION NOTES:
'   - Introduced in Version 10.0.0 as part of the ADO SQL hardening changes
'
' USAGE:
'   - Ideal for micro-delays in recursive routines, UI pacing, or throttled
'     polling loops where API calls are undesirable.
'
' RELATED WIKI PAGES:
'   - Timing & Delay Patterns
'   - COM Reentrancy & Safe Looping
' =============================================================================

Option Explicit

' Simple sleep using Timer + DoEvents (no API).
Public Sub SleepMilliseconds(ByVal ms As Long)
#If Win32 Or Win64 Then
    Dim t As Single
    t = Timer
    Do While (Timer - t) * 1000 < ms
        DoEvents
    Loop
#End If
End Sub


