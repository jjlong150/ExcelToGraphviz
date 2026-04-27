VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StyleDesignerSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
' PROJECT:          Excel to Graphviz
' WORKSHEET MODULE: StyleDesignerSheet
' COPYRIGHT:        Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:            Relationship Visualizer / Sheets / Style Designer
'
' ROLE:
'   Interactive design surface for authoring and previewing Graphviz node,
'   edge, and cluster styles. Provides a reactive rendering engine, dual
'   editing pathways (UI lists and raw DOT format strings), and contextual
'   Ribbon activation for visual-design workflows.
'
' DESCRIPTION:
'   This worksheet module governs the behavior of the "style designer" sheet.
'   When activated, it switches the Ribbon to the Style Designer tab and
'   immediately renders a preview to align the UI with the current worksheet
'   state. Any edits to labels, format strings, or image attributes trigger
'   targeted preview updates. The module also synchronizes image-related
'   controls (scale, position) and uses event-hardened logic to prevent
'   recursive rendering loops.
'
' KEY RESPONSIBILITIES:
'   - CONTEXT-AWARE RIBBON ACTIVATION:
'       • Activates the Style Designer Ribbon tab on entry
'       • Surfaces color, font, attribute, and image-handling tools
'
'   - REAL-TIME PREVIEW ENGINE:
'       • Calls RenderPreview on activation
'       • Routes Worksheet_Change events to:
'           – RenderPreviewFromLists (label/UI-driven changes)
'           – RenderPreviewFromFormatString (raw DOT edits)
'           – RefreshControlsImageScale / RefreshControlsImagePosition
'
'   - BI-DIRECTIONAL EDITING:
'       • Supports UI-friendly attribute lists
'       • Supports direct DOT format-string editing with live feedback
'
'   - IMAGE-AWARE CONTROL SYNC:
'       • Automatically refreshes scale/position controls when image names
'         change, ensuring UI consistency with rendered output
'
'   - PLATFORM-SAFE ACTIVATION:
'       • Uses Application.OnTime on macOS to avoid Ribbon race conditions
'       • Performs immediate activation on Windows for seamless transitions
'
'   - ERROR-HARDENED EVENT MANAGEMENT:
'       • Disables Application.EnableEvents during updates
'       • Uses a recovery block to guarantee event-state restoration
'
' USAGE:
'   - Fires automatically when navigating to the Style Designer sheet
'   - Provides the primary environment for designing reusable Graphviz styles
'
' RELATED WIKI PAGES:
'   - Style Designer Overview
'   - Format-String Architecture
'   - Live Preview Rendering Pipeline
'   - Image Handling & Scale/Position Controls
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: Worksheet_Activate
' PURPOSE:
'   Synchronizes the Ribbon UI and initializes the visual design session.
'
' TECHNICAL WORKFLOW:
'   1. UI NAVIGATION: Triggers 'ActivateTabStyleDesigner' to surface
'      specialized color, font, and attribute pickers.
'   2. PLATFORM ADAPTATION: Handles Mac-specific timing via 'Application.OnTime'.
'   3. INITIAL RENDERING: Calls 'RenderPreview' to ensure the preview image
'      matches the current worksheet state upon entry.
' ==========================================================================
Private Sub Worksheet_Activate()
#If Mac Then
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabStyleDesigner"
#Else
    ActivateTabStyleDesigner
#End If
    RenderPreview
End Sub

' ==========================================================================
' PROCEDURE: Worksheet_Change
' PURPOSE:
'   Drives the reactive live-preview engine based on user input.
'
' TECHNICAL WORKFLOW:
'   1. ATTRIBUTE DETECTION: Monitors changes in labels, format strings,
'      and image properties using 'Intersect'.
'   2. PREVIEW ROUTING:
'      - Label Changes -> 'RenderPreviewFromLists'
'      - Format String Changes -> 'RenderPreviewFromFormatString'
'      - Image Changes -> Refreshes scaling and position controls.
'   3. EVENT LOCKDOWN: Disables 'Application.EnableEvents' during updates
'      to prevent recursive rendering loops.
'   4. ERROR RECOVERY: Uses a recovery block to log failures while
'      guaranteeing the restoration of Excel's event state.
' ==========================================================================
Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo Whoa

    ' Detect changes in the NodeFormatString text box
    If Not Intersect(Target, StyleDesignerSheet.Range("LabelText,XlabelText,TailLabelText,HeadLabelText")) Is Nothing Then
        Application.enableEvents = False
        RenderPreviewFromLists
        
    ElseIf Not Intersect(Target, StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING)) Is Nothing Then
        Application.enableEvents = False
        If StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING).value = vbNullString Then
            ClearStyleRibbon
        Else
            RenderPreviewFromFormatString
            StyleDesignerSheet.Range(DESIGNER_FORMAT_STRING).Select
        End If
        
    ElseIf Not Intersect(Target, StyleDesignerSheet.Range(DESIGNER_ADD_CAPTION)) Is Nothing Then
        Application.enableEvents = False
        RenderPreview

    ElseIf Not Intersect(Target, StyleDesignerSheet.Range(DESIGNER_NODE_IMAGE_NAME)) Is Nothing Then
        Application.enableEvents = False
        RefreshControlsImageScale
        RefreshControlsImagePosition
    End If
    
LetsContinue:
    Application.enableEvents = True
    Exit Sub
Whoa:
    EmitMessage "StyleDesignerSheet " & err.number & " " & err.Description
    
    Resume LetsContinue

End Sub

