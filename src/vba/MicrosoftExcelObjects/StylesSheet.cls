VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StylesSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
' PROJECT:          Excel to Graphviz
' WORKSHEET MODULE: StylesSheet
' COPYRIGHT:        Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:            Relationship Visualizer / Sheets / Styles
'
' ROLE:
'   Authoritative repository for all Node, Edge, and Cluster styles. Provides
'   an interactive Style Gallery with live previews, floating action buttons,
'   and seamless integration with the Style Designer for GUI-based refinement.
'
' DESCRIPTION:
'   This worksheet module governs the behavior of the Styles Gallery. When
'   activated, it switches the Ribbon to the Styles tab and refreshes the
'   Edit-Style control to match the current selection. As the user navigates
'   the sheet, a floating-button engine injects Preview and Edit controls next
'   to the active style row, enabling instant testing and round-trip editing.
'   Header edits trigger Ribbon refreshes to maintain global consistency.
'
' KEY RESPONSIBILITIES:
'   - CONTEXT-AWARE RIBBON ACTIVATION:
'       • Activates the Styles Ribbon tab on entry
'       • Invalidates the Edit-Style control to match the active row
'
'   - INTERACTIVE PREVIEW ENGINE:
'       • Uses Worksheet_SelectionChange to purge and recreate floating
'         Preview/Edit buttons
'       • Calls PreviewStyleForCurrentRow for instant thumbnail rendering
'
'   - DESIGNER INTEGRATION:
'       • Provides a direct pathway to RestoreStyleDesigner for GUI-based
'         editing of existing style strings
'
'   - REACTIVE METADATA MODEL:
'       • Monitors header edits (E1:AZ1) and triggers RefreshRibbon to keep
'         galleries, pickers, and Style Designer controls synchronized
'
'   - DATA HYGIENE:
'       • Uses IsStylesRowActive to suppress UI noise on header/comment rows
'
'   - PLATFORM-SAFE ACTIVATION:
'       • Uses Application.OnTime on macOS to avoid Ribbon race conditions
'       • Performs immediate activation on Windows for seamless transitions
'
' USAGE:
'   - Fires automatically when navigating to the Styles worksheet
'   - Provides the primary environment for browsing, testing, and refining
'     reusable Graphviz styles
'
' RELATED WIKI PAGES:
'   - Style Gallery Overview
'   - Style Designer Integration
'   - Floating UI Engine
'   - Style Metadata & Format-String Architecture
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: Worksheet_Activate
' PURPOSE:
'   Synchronizes the Ribbon UI with the Style Gallery environment.
'
' TECHNICAL WORKFLOW:
'   1. UI REFRESH: Immediately invalidates the 'Edit Style' ribbon control
'      to match the current selection.
'   2. TAB SELECTION: Switches to the 'Styles' Ribbon tab (with Mac delay).
' ==========================================================================
Private Sub Worksheet_Activate()
    InvalidateRibbonControl RIBBON_CTL_STYLES_EDIT
#If Mac Then
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ActivateTabStyles"
#Else
    ActivateTabStyles
#End If
End Sub

' ==========================================================================
' PROCEDURE: Worksheet_Change
' PURPOSE:
'   Synchronizes the Ribbon UI when style metadata or headers are modified.
'
' TECHNICAL WORKFLOW:
'   1. INTERSECT CHECK: Specifically watches the horizontal header range
'      (E1:AZ1) for edits.
'   2. UI REDRAW: Disables events and calls 'RefreshRibbon' to update
'      the visual state of the application.
' ==========================================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler

    ' Detect changes in the view switch headings
    If Not Intersect(Target, StylesSheet.Range("E1:AZ1")) Is Nothing Then
        Application.enableEvents = False
        RefreshRibbon
    End If

    If IsStylesRowActive(Target.row) Then
        PreviewStyleForRow Target.row
    End If
    
CleanExit:
    Application.enableEvents = True
    Exit Sub

ErrorHandler:
    EmitMessage "Sheet: StylesSheet - " & err.number & " " & err.Description
    err.Clear
    Resume CleanExit
End Sub

' ==========================================================================
' PROCEDURE: Worksheet_SelectionChange
' PURPOSE:
'   Orchestrates the interactive "Follow-the-Cursor" style tools.
'
' TECHNICAL WORKFLOW:
'   1. PROTECTION CHECK: Exits if the worksheet is locked.
'   2. UI INVALIDATION: Refreshes the Ribbon 'Edit' button state.
'   3. FLOATING UI:
'      - Clears previous on-sheet buttons.
'      - Re-injects 'Preview' and 'Edit' buttons next to the target format cell.
' ==========================================================================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Exit if protected worksheet
    If Me.ProtectContents Then Exit Sub
    
    ' Enable or disable the edit button in the ribbon
    On Error Resume Next
    InvalidateRibbonControl RIBBON_CTL_STYLES_EDIT
    On Error GoTo 0
    
    ' Get button configs and use for both creation and cleanup
    Dim configs() As ButtonConfig
    configs = GetFloatingButtonConfigs()

    ' Cleanup old buttons, and create new ones
    RemoveFloatingButtons Me, GetButtonNamesFromConfigs(configs)
    UpdateFloatingButtonsOnSheet Me, Target, configs
End Sub

' ==========================================================================
' FUNCTION: GetFloatingButtonConfigs
' PURPOSE:
'   Defines the interactive on-sheet controls for the Styles Gallery.
'
' TECHNICAL WORKFLOW:
'   1. REGISTRY INITIALIZATION: Allocates an array of 'ButtonConfig' UDTs
'      to store the visual and functional metadata for three controls.
'   2. ICONOGRAPHY & MAPPING:
'      - Refresh: Triggers 'PreviewStyleForCurrentRow' to render a
'        temporary thumbnail of the style.
'      - Edit: Triggers 'RestoreStyleDesigner' to load the current
'        style string back into the Visual Style Designer.
'   3. TARGETED POSITIONING:
'      - Pins buttons to the 'Format' column for raw code edits.
'      - Pins a secondary 'Refresh' button to the 'Style Name' column for
'        easier access during organization.
'   4. LOGIC VALIDATION: Binds every button to 'IsStylesRowActive' to ensure
'      controls only appear on functional style rows (skipping headers/comments).
' ==========================================================================
Private Function GetFloatingButtonConfigs() As ButtonConfig()
    Dim configs() As ButtonConfig
    ReDim configs(0 To 2)   ' change upper bound when adding more

    With configs(0)
        .ColumnNumber = GetSettingColNum(SETTINGS_STYLES_COL_FORMAT)
        .ButtonName = "StyleRefreshButton"
        .ActionMacro = "PreviewStyleForCurrentRow"
        .IconText = ChrW(&H27F3)
        .VOffset = 0
        .HOffset = 5
        .ValidationFunc = "IsStylesRowActive"
    End With
    
    With configs(1)
        .ColumnNumber = GetSettingColNum(SETTINGS_STYLES_COL_FORMAT)
        .ButtonName = "FormatEditButton"
        .ActionMacro = "RestoreStyleDesigner"
        .IconText = ChrW(&H270E)
        .VOffset = 22
        .HOffset = 5
        .ValidationFunc = "IsStylesRowActive"
    End With

    With configs(2)
        .ColumnNumber = GetSettingColNum(SETTINGS_STYLES_COL_STYLE)
        .ButtonName = "StyleRefreshButton"
        .ActionMacro = "PreviewStyleForCurrentRow"
        .IconText = ChrW(&H27F3)
        .VOffset = 0
        .HOffset = 5
        .ValidationFunc = "IsStylesRowActive"
    End With

    GetFloatingButtonConfigs = configs
End Function


