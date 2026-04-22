VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HelpAttributesSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
' PROJECT:          Excel to Graphviz
' WORKSHEET MODULE: HelpAttributesSheet
' COPYRIGHT:        Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:            Relationship Visualizer / Sheets / Reference
'
' ROLE:
'   Technical reference worksheet containing the authoritative list of
'   Graphviz attributes, types, defaults, and usage notes. Provides a stable
'   lookup surface for users exploring advanced styling and DOT-level options.
'
' DESCRIPTION:
'   This worksheet module manages the interactive behavior of the "HELP - attributes"
'   reference sheet. When activated, it switches the Ribbon to the Launchpad
'   tab, ensuring users have access to worksheet navigation and administrative
'   tools. It also supports macOS-safe deferred activation to avoid Ribbon
'   timing issues. The sheet itself serves as the primary reference location
'   for users investigating Graphviz-native attributes and "Extra Style
'   Attributes" supported by the Style Designer.
'
' KEY RESPONSIBILITIES:
'   - CONTEXT-AWARE NAVIGATION:
'       • Activates the Launchpad Ribbon tab when the sheet is selected
'       • Ensures worksheet-level tools are immediately accessible
'
'   - PLATFORM-SAFE ACTIVATION:
'       • Uses Application.OnTime on macOS to avoid Ribbon race conditions
'       • Performs immediate activation on Windows for seamless transitions
'
'   - REFERENCE SUPPORT:
'       • Serves as the lookup table for Graphviz attribute definitions
'       • Supports Style Designer workflows that rely on attribute metadata
'
' USAGE:
'   - Fires automatically when the user navigates to the Attributes reference
'     sheet, ensuring the Ribbon is aligned with worksheet-level tools
'
' RELATED WIKI PAGES:
'   - Graphviz Attribute Reference
'   - Style Designer: Extra Attributes
'   - Worksheet Navigation & Launchpad
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: Worksheet_Activate
' PURPOSE:
'   Synchronizes the Ribbon UI with the technical reference workspace.
'
' TECHNICAL WORKFLOW:
'   1. UI NAVIGATION: Triggers 'ActivateTabLaunchpad' to provide access
'      to the global 'Worksheets' and 'Settings' controls.
'   2. PLATFORM ADAPTATION:
'      - macOS: Uses 'Application.OnTime' (1s delay) to ensure the
'        Ribbon handle is initialized properly.
'      - Windows: Activates the tab immediately for zero-latency navigation.
' ==========================================================================
Private Sub Worksheet_Activate()
#If Mac Then
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabLaunchpad"
#Else
    ActivateTabLaunchpad
#End If
End Sub

