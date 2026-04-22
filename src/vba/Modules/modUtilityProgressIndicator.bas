Attribute VB_Name = "modUtilityProgressIndicator"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modUtilityProgressIndicator
' COPYRIGHT: Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Excel UI / Feedback
'
' ROLE:
'   Lightweight wrapper around the ProgressIndicatorForm userform, providing
'   simple, non-modal progress reporting during long-running operations.
'
' RESPONSIBILITIES:
'   - UpdateProgressIndicator:
'       • Update percentage text and bar width
'       • Repaint form to ensure immediate visual feedback
'   - ShowProgressIndicator:
'       • Display modeless progress form with caller-supplied title
'   - HideProgressIndicator:
'       • Unload the form and release UI resources
'
' ARCHITECTURAL NOTES:
'   - Modeless display allows workflows to continue executing without blocking.
'   - Bar width uses a fixed multiplier (pctCompl * 2) aligned with form layout.
'   - Defensive visibility check prevents updates when the form is not shown.
'   - Consumed by SQL engine, file operations, and any multi-step processes
'     requiring user-visible progress cues.
'
' VERSION NOTES:
'   - Introduced in Version 6.1.0
'   - Ceased being used in Version 10.3.0 (deprecated)
'
' USAGE:
'   - Wrap long operations:
'         ShowProgressIndicator "Exporting…"
'         UpdateProgressIndicator pct
'         HideProgressIndicator
'
' RELATED WIKI PAGES:
'   - UI Feedback & Progress Indicators
'   - Long-Running Operation Patterns
' =============================================================================

Option Explicit

Public Sub UpdateProgressIndicator(ByVal pctCompl As Long)
    If ProgressIndicatorForm.visible Then
        ProgressIndicatorForm.Text.caption = pctCompl & "%"
        ProgressIndicatorForm.Bar.Width = pctCompl * 2
        ProgressIndicatorForm.Repaint
    End If
End Sub

Public Sub ShowProgressIndicator(ByVal title As String)
    ProgressIndicatorForm.caption = title
    ProgressIndicatorForm.show vbModeless
End Sub

Public Sub HideProgressIndicator()
    Unload ProgressIndicatorForm
End Sub
