VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SvgSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
' PROJECT:          Excel to Graphviz
' WORKSHEET MODULE: SvgSheet
' COPYRIGHT:        Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:            Relationship Visualizer / Sheets / SVG
'
' ROLE:
'   Event-driven UI host for the "svg" worksheet. Manages contextual Ribbon
'   activation, floating edit controls, and selection-sensitive behaviors
'   that support the SVG post-processing workflow.
'
' DESCRIPTION:
'   This worksheet module provides the interactive layer for editing SVG
'   transformation rules. It reacts to user navigation, injects a floating
'   "Edit" button when the Replace column is selected, launches the large
'   multi-line editor, and keeps Ribbon controls synchronized with the
'   current cursor context. All XML transformation logic is delegated to
'   modWorksheetSVG; this module focuses exclusively on user interaction.
'
' KEY RESPONSIBILITIES:
'   - CONTEXTUAL NAVIGATION:
'       • Activates the SVG Ribbon tab when the sheet is selected
'       • Invalidates Edit/Clipboard controls to reflect cursor position
'
'   - INTERACTIVE EDITING:
'       • Injects a floating pencil-icon button beside Replace cells
'       • Launches CellValueEditForm for multi-line CSS/JS/XML editing
'
'   - UI STATE MANAGEMENT:
'       • Cleans up injected shapes on selection changes
'       • Handles macOS Ribbon activation via deferred OnTime scheduling
'
'   - UX RESPONSIVENESS:
'       • Ensures smooth interaction when navigating or editing rules
'
' RELATED WIKI PAGES:
'   - SVG Post-Processing
'   - Editing Replacement Rules
'   - JavaScript Animation Injection
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: Worksheet_Activate
' PURPOSE:
'   Synchronizes the Ribbon UI and initializes the SVG post-processing workspace.
'
' TECHNICAL WORKFLOW:
'   1. UI NAVIGATION: Triggers 'ActivateTabSvg' to surface SVG-specific
'      utilities (Apply Changes, Export, SVG Editor).
'   2. STATE REFRESH: Specifically invalidates 'Edit Cell' and 'Clipboard'
'      ribbon controls to sync with the active selection.
' ==========================================================================
Private Sub Worksheet_Activate()
#If Mac Then
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabSvg"
#Else
    ActivateTabSvg
#End If
    InvalidateRibbonControl RIBBON_CTL_SVG_EDIT_CELL
    InvalidateRibbonControl RIBBON_CTL_SVG_CLIPBOARD
End Sub

' ==========================================================================
' PROCEDURE: Worksheet_Deactivate
' PURPOSE:
'   Synchronizes the Ribbon UI state upon exiting the SVG workspace.
'
' TECHNICAL WORKFLOW:
'   1. UI INVALIDATION: Specifically refreshes the 'Edit Cell' and 'Clipboard'
'      Ribbon controls.
'   2. STATE TRANSITION: Ensures the Ribbon focus is reset, preventing
'      accidental triggering of SVG-specific editing forms on standard
'      worksheet ranges.
'
' USAGE:
'   - Fires automatically whenever the user navigates away from the SVG sheet.
'   - Maintains UI hygiene by resetting the context for developer tools.
' ==========================================================================
Private Sub Worksheet_Deactivate()
    InvalidateRibbonControl RIBBON_CTL_SVG_EDIT_CELL
    InvalidateRibbonControl RIBBON_CTL_SVG_CLIPBOARD
End Sub

' ==========================================================================
' PROCEDURE: Worksheet_SelectionChange
' PURPOSE:
'   Manages the "Edit-on-Demand" interactive UI for SVG XML injection.
'
' TECHNICAL WORKFLOW:
'   1. UI CLEANUP: Automatically deletes any existing 'SVGEditButton' to
'      ensure the interface remains fresh.
'   2. COLUMN VALIDATION: Checks if the selection is in the 'Replace'
'      column where complex code injections typically occur.
'   3. SHAPE INJECTION: Programmatically draws a rounded-rectangle button
'      with a Unicode pencil icon next to the active cell.
'   4. LOGIC BINDING: Links the button to 'ShowSVGEditForm' for handling
'      large blocks of XML or CSS.
' ==========================================================================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Not TypeOf Target Is Range Then Exit Sub

    ' Remove any existing Edit button
    On Error Resume Next
    Me.buttons("SVGEditButton").Delete
    Me.Shapes("SVGEditButton").Delete
    On Error GoTo 0

    If Target.Cells.count = 1 Then
        If Not Intersect(Target, Me.columns(svgLayoutColumn.replaceColumn)) Is Nothing Then
            If Not Me.ProtectContents Then
                On Error Resume Next
                ' Add clickable pencil icon with subtle button look
                Dim shp As shape
                Set shp = Me.Shapes.AddShape(msoShapeRoundedRectangle, _
                            Target.Left + Target.Width + 5, _
                            Target.top + 1, 20, 20)
                With shp
                    .name = "SVGEditButton"
                    .Fill.ForeColor.RGB = RGB(245, 245, 245)      ' very light gray background
                    .Fill.Transparency = 0
                    .line.Weight = 0.75
                    .line.ForeColor.RGB = RGB(180, 180, 180)      ' soft border
                    .TextFrame2.TextRange.Characters.Text = ChrW(&H270E)  ' pencil
                    .TextFrame2.TextRange.Font.name = "Segoe UI Symbol"
                    .TextFrame2.TextRange.Font.Size = 16
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(60, 60, 60)  ' dark gray pencil
                    .TextFrame2.VerticalAnchor = msoAnchorMiddle
                    .TextFrame2.HorizontalAnchor = msoAnchorCenter
                    .Placement = xlMoveAndSize
                    .OnAction = "ShowSVGEditForm"
                End With
                On Error GoTo 0
            End If
        End If
    End If

    InvalidateRibbonControl RIBBON_CTL_SVG_EDIT_CELL
    InvalidateRibbonControl RIBBON_CTL_SVG_CLIPBOARD
    RefreshRibbon
End Sub
