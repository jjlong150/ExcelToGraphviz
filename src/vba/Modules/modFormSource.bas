Attribute VB_Name = "modFormSource"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    DotFormSource
' COPYRIGHT: Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:     UI / Forms Subsystem
'
' ROLE:
'   Controller for the DOT Source Viewer form. Manages lifecycle, localization,
'   and real-time population of the multiline DOT preview surface.
'
' RESPONSIBILITIES:
'   - Form lifecycle management:
'       • show/hide the DOT Source Viewer
'       • clear/reset the multiline text surface
'   - Localization:
'       • apply translated captions to Copy and Word-Wrap controls
'   - Source presentation:
'       • inject generated DOT source into the form when visible
'       • normalize line endings for consistent display
'
' ARCHITECTURAL NOTES:
'   - Lightweight controller invoked by modRibbonTabSource and modCreateGraph.
'   - Uses DotSourceForm as the UI surface; no external dependencies.
'   - Designed for non-modal, always-on-top inspection during graph generation.
'
' USAGE:
'   - Called by Ribbon actions ("View Source", "Copy Source").
'   - Used during debugging, validation, and advanced authoring workflows.
'
' RELATED WIKI PAGES:
'   - DOT Source Viewer
'   - Source Worksheet & Source Form Architecture
' =============================================================================

Option Explicit

Public Sub ClearSourceForm()
    DotSourceForm.dotMultiline.Text = vbNullString
End Sub

Public Sub ShowSourceForm()
    DotSourceForm.CopyButton.caption = GetLabel("sourceFormCopy")
    DotSourceForm.wordWrapToggle.caption = GetLabel("sourceFormWrapText")
    DotSourceForm.show
    ClearSourceForm
End Sub

Public Sub HideSourceForm()
    Unload DotSourceForm
End Sub

Public Sub DisplaySourceInForm(ByVal dotSource As String)
    If Not DotSourceForm.visible Then Exit Sub
    
    Dim popupSource As String
    popupSource = dotSource
    replace popupSource, vbLf, Chr$(10)
    DotSourceForm.dotMultiline.Text = popupSource
End Sub

