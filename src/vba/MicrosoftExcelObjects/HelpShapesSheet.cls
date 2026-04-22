VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HelpShapesSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
' PROJECT:          Excel to Graphviz
' WORKSHEET MODULE: HelpShapesSheet
' COPYRIGHT:        Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:            Relationship Visualizer / Sheets / Reference
'
' ROLE:
'   Visual reference surface for all supported Graphviz node shapes. Provides
'   version-specific notes, compatibility guidance, and a companion lookup
'   resource for the Style Designer's shape-selection controls.
'
' DESCRIPTION:
'   This worksheet module manages the interactive behavior of the Shapes
'   reference sheet. When activated, it switches the Ribbon to the Launchpad
'   tab for worksheet-level utilities and uses macOS-safe deferred activation
'   to avoid Ribbon timing issues. The sheet itself serves as a visual guide
'   for users exploring Graphviz node shapes, including those requiring newer
'   Graphviz versions or special attribute combinations. Clicking on any shape
'   image transfers the shape's name to the Style Designer's shape field.
'
' KEY RESPONSIBILITIES:
'   - CONTEXT-AWARE NAVIGATION:
'       • Activates the Launchpad Ribbon tab when the sheet is selected
'       • Ensures worksheet navigation and settings tools are accessible
'
'   - SHAPE REFERENCE SUPPORT:
'       • Displays the catalog of Graphviz node shapes
'       • Provides compatibility notes for version-specific shapes
'       • Serves as a companion to the Style Designer's shape controls
'
'   - PLATFORM-SAFE ACTIVATION:
'       • Uses Application.OnTime on macOS to avoid Ribbon race conditions
'       • Performs immediate activation on Windows for seamless transitions
'
' USAGE:
'   - Fires automatically when the user navigates to the Shapes reference sheet
'   - Provides a visual guide for selecting Graphviz node shapes
'
' RELATED WIKI PAGES:
'   - Graphviz Node Shapes
'   - Style Designer: Shape Controls
'   - Worksheet Navigation & Launchpad
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: Worksheet_Activate
' PURPOSE:
'   Synchronizes the Ribbon UI with the visual shapes reference.
'
' TECHNICAL WORKFLOW:
'   1. UI NAVIGATION: Triggers 'ActivateTabLaunchpad' to provide access
'      to the global 'Worksheets' and 'Settings' controls.
'   2. PLATFORM ADAPTATION:
'      - macOS: Uses 'Application.OnTime' (1s delay) for stable UI activation.
'      - Windows: Activates the tab immediately for a seamless transition.
' ==========================================================================
Private Sub Worksheet_Activate()
#If Mac Then
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabLaunchpad"
#Else
    ActivateTabLaunchpad
#End If
End Sub

