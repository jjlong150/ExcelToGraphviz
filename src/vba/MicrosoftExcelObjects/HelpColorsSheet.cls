VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HelpColorsSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
' PROJECT:          Excel to Graphviz
' WORKSHEET MODULE: HelpColorsSheet
' COPYRIGHT:        Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:            Relationship Visualizer / Sheets / Reference
'
' ROLE:
'   Visual reference surface for all supported Graphviz color schemes.
'   Provides interactive hyperlink integration that allows users to apply
'   color schemes directly to the Style Designer with a single click. Provides
'   the catalog of color values used to create Style Designer image galleries.
'
' DESCRIPTION:
'   This worksheet module manages the interactive behavior of the "HELP - colors"
'   reference sheet. When activated, it switches the Ribbon to the Launchpad
'   tab for worksheet-level utilities and uses macOS-safe deferred activation
'   to avoid Ribbon timing issues. It also captures hyperlink clicks to
'   instantly apply a selected color scheme to the Style Designer, enabling
'   rapid experimentation with Graphviz palettes.
'
' KEY RESPONSIBILITIES:
'   - CONTEXT-AWARE NAVIGATION:
'       • Activates the Launchpad Ribbon tab when the sheet is selected
'       • Ensures worksheet navigation and settings tools are accessible
'
'   - INTERACTIVE COLOR INJECTION:
'       • Captures hyperlink clicks via Worksheet_FollowHyperlink
'       • Passes the selected scheme to SetStyleDesignerColorScheme
'       • Enables "one-click" palette application for rapid styling
'
'   - PLATFORM-SAFE ACTIVATION:
'       • Uses Application.OnTime on macOS to avoid Ribbon race conditions
'       • Performs immediate activation on Windows for seamless transitions
'
' USAGE:
'   - Fires automatically when the user navigates to the Colors reference sheet
'   - Provides a fast, interactive way to apply Graphviz color schemes
'   - Referenced by Style Designer when constructing images to populate image
'     galleries for color-related attribute choices.
'
' RELATED WIKI PAGES:
'   - Graphviz Color Schemes
'   - Style Designer: Color & Palette Controls
'   - Worksheet Navigation & Launchpad
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: Worksheet_Activate
' PURPOSE:
'   Synchronizes the Ribbon UI and prepares the color reference workspace.
'
' TECHNICAL WORKFLOW:
'   1. UI NAVIGATION: Triggers 'ActivateTabLaunchpad' to provide access
'      to global worksheet navigation and settings.
'   2. PLATFORM ADAPTATION: Handles Mac-specific timing via 'Application.OnTime'.
' ==========================================================================
Private Sub Worksheet_Activate()
#If Mac Then
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabLaunchpad"
#Else
    ActivateTabLaunchpad
#End If
End Sub

' ==========================================================================
' PROCEDURE: Worksheet_FollowHyperlink
' PURPOSE:
'   Acts as a "Color Picker" for the Style Designer colorscheme.
'
' TECHNICAL WORKFLOW:
'   1. EVENT CAPTURE: Detects when a user clicks a colorscheme name hyperlink.
'   2. ATTRIBUTE INJECTION: Passes 'Target.TextToDisplay' (the color name)
'      to 'SetStyleDesignerColorScheme' in 'modStyleDesigner'.
'   3. SILENT RECOVERY: Uses a localized error trap to bypass failures
'      if the Style Designer sheet is protected or invalid.
' ==========================================================================
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    On Error GoTo ErrorHandler
    SetStyleDesignerColorScheme (Target.TextToDisplay)
    Exit Sub

ErrorHandler:
    err.Clear
    Resume Next
End Sub


