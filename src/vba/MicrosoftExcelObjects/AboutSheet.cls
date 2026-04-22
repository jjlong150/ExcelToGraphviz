VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AboutSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
' PROJECT:          Excel to Graphviz
' WORKSHEET MODULE: AboutSheet
' COPYRIGHT:        Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:            Relationship Visualizer / Sheets / Info
'
' ROLE:
'   Event-driven UI host for the informational "info" worksheet.
'   Manages Ribbon activation, localized legal text, and branding visibility
'   whenever the user navigates to the sheet.
'
' DESCRIPTION:
'   This worksheet module provides the interactive layer for the application's
'   informational and licensing center. When activated, it ensures the Ribbon
'   switches to the About tab, applies macOS-safe deferred activation when
'   required, and serves as the target for localized licensing text injected
'   by the localization subsystem. Its purpose is strictly UI-centric: it
'   coordinates context, branding, and legal visibility rather than performing
'   any business logic.
'
' KEY RESPONSIBILITIES:
'   - CONTEXT-AWARE NAVIGATION:
'       • Activates the About Ribbon tab when the sheet is selected
'       • Ensures licensing, credits, and version information are visible
'
'   - LOCALIZATION TARGET:
'       • Receives translated legal text via modLocalize.LocalizeWorksheetAbout
'       • Updates credits and licensing paragraphs based on active language
'
'   - PLATFORM-SAFE RIBBON ACTIVATION:
'       • Uses Application.OnTime on macOS to avoid Ribbon race conditions
'       • Performs immediate activation on Windows for seamless transitions
'
'   - UI STATE MANAGEMENT:
'       • Ensures About-specific Ribbon controls are synchronized with focus
'
' USAGE:
'   - Fires automatically when the user selects the Info/About worksheet
'   - Ensures branding, licensing, and contributor information remain accessible
'
' RELATED WIKI PAGES:
'   - Application Branding & Credits
'   - Localization Framework
'   - Ribbon Navigation Architecture
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: Worksheet_Activate
' PURPOSE:
'   Synchronizes the Ribbon UI with the 'About' worksheet.
'
' TECHNICAL WORKFLOW:
'   1. UI NAVIGATION: Triggers 'ActivateTabAbout' to switch the Ribbon
'      focus to the informational and legal controls.
'   2. PLATFORM ADAPTATION:
'      - macOS: Utilizes 'Application.OnTime' to introduce a one-second delay,
'        ensuring the Ribbon handle is stable within the Mac sandbox.
'      - Windows: Executes the tab switch immediately for seamless transition.
'
' USAGE:
'   - Fires automatically when the user selects the 'About' tab.
'   - Ensures the 'Relationship Visualizer' branding and licensing tools
'     are front-and-center.
' ==========================================================================
Private Sub Worksheet_Activate()
#If Mac Then
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabAbout"
#Else
    ActivateTabAbout
#End If
End Sub

