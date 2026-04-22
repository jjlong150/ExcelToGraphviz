Attribute VB_Name = "modRibbon"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modRibbon
' COPYRIGHT: Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Excel UI / Ribbon
'
' ROLE:
'   Central bootstrapper and controller for the custom Office Ribbon. Manages
'   lifecycle, visibility, localization, and context-sensitive tab activation
'   across all Relationship Visualizer workflows.
'
' RESPONSIBILITIES:
'   - Ribbon lifecycle management:
'       • Capture and cache the IRibbonUI handle on load
'       • Preload font and color image assets for instant dropdown rendering
'       • Defer initial tab activation to avoid Excel UI race conditions
'   - Context-sensitive routing:
'       • Auto-select the correct Ribbon tab based on the active worksheet
'       • Provide named wrappers for tab activation (Graphviz, SQL, Styles, etc.)
'   - Dynamic localization:
'       • Bind getLabel/getScreentip/getSupertip to worksheet-driven values
'       • Refresh Ribbon state when language or settings change
'   - Visibility and platform gating:
'       • Respect user toggles for optional tabs (Console, Diagnostics, Styles)
'       • Enforce macOS restrictions (SQL tab hidden on Mac)
'   - Partial and full invalidation:
'       • Invalidate entire Ribbon or individual controls for efficient updates
'
' ARCHITECTURAL NOTES:
'   - Ribbon XML is data-driven: Named Ranges on the Settings sheet act as
'     the authoritative source for labels, screentips, and visibility flags.
'   - IRibbonUI handle is stored statically to survive cross-module calls.
'   - Designed to remain resilient after VBA resets, workbook reloads, and
'     cross-platform execution (Windows/macOS).
'   - Integrates with localization, settings, and worksheet activation events.
'
' USAGE:
'   - Called automatically by Ribbon XML (onLoad, getVisible, getLabel, etc.).
'   - Invoked by worksheet activation events to maintain UI context.
'   - Used by settings/localization modules to refresh the Ribbon after changes.
'
' RELATED WIKI PAGES:
'   - Ribbon XML Architecture
'   - Worksheet-Driven UI Model
'   - Context-Sensitive Tab Routing
' =============================================================================

Option Explicit

' The private storage variable for the Ribbon handle
Private internalMyRibbon As IRibbonUI

''
' RIBBON ACCESSORS: Standardized Get/Let properties to manage the global
' IRibbonUI object.
'
' TECHNICAL WORKFLOW:
'   1. ENCAPSULATION: Hides the raw 'internalMyRibbon' variable to ensure
'      the pointer is only modified through controlled property calls.
'   2. GLOBAL REACH: Allows any module in the project to call 'myRibbon.Invalidate'
'      to trigger UI updates when data or settings change.
'
Public Static Property Get myRibbon() As IRibbonUI
    Set myRibbon = internalMyRibbon
End Property

Public Static Property Let myRibbon(ByVal ribbon As IRibbonUI)
    Set internalMyRibbon = ribbon
End Property

' ==========================================================================
' PROCEDURE: ribbon_onLoad
' PURPOSE:
'   The primary callback triggered by Excel when the custom UI is loaded.
'
' TECHNICAL WORKFLOW:
'   1. HANDLE CAPTURE: Assigns the 'IRibbonUI' object to the global 'myRibbon'
'      property for project-wide UI control.
'   2. ASSET PRE-LOADING: Triggers the initialization of the Font and Color
'      image caches to ensure the Ribbon dropdowns are populated instantly.
'   3. DEFERRED ACTIVATION: Schedules a one-second delayed call to
'      'ribbon_activateTab' using 'Application.OnTime'.
'
' USAGE:
'   - Fires once per session when the Relationship Visualizer is opened.
'   - The delay prevents "UI Race Conditions" where Excel tries to select
'     a tab before the Ribbon is fully rendered.
' ==========================================================================
Public Sub ribbon_onLoad(ByVal ribbon As IRibbonUI)
    '@Ignore ValueRequired
    myRibbon = ribbon
    LoadFontImageCache
    ColorLoadImageCache
    Application.OnTime Now + TimeValue(ONE_SECOND_DELAY), "ribbon_activateTab"
End Sub

' ==========================================================================
' PROCEDURE: ribbon_activateTab
' PURPOSE:
'   Synchronizes the active Ribbon tab with the user's current worksheet.
'
' TECHNICAL WORKFLOW:
'   1. LAZY INITIALIZATION: Triggers 'TabSelectGraphOptions' to ensure
'      core settings are loaded before the UI is presented.
'   2. CONTEXTUAL ROUTING: Uses a 'Select Case' on the 'ActiveSheet.name'
'      to determine the target tab:
'      - Worksheets (Data, Graph) -> ActivateTabGraphviz
'      - Development (SQL)        -> ActivateTabSql
'      - Styling (Styles, StyleDesigner) -> Specialized Styling Tabs
'      - Support (Settings, Locale)      -> ActivateTabLaunchpad
'   3. ERROR RESILIENCE: Implements a localized handler to silently
'      bypass UI activation failures (e.g., if the Ribbon handle is lost).
'
' USAGE:
'   - Called during 'ribbon_onLoad' (via OnTime) for initial startup.
'   - Frequently called by sheet activation events to maintain UI context.
' ==========================================================================
Public Sub ribbon_activateTab()
    On Error GoTo ErrorHandler

    ' Defer initialization of 'settings' worksheet until workbook
    ' startup is complete
    TabSelectGraphOptions
    
    ' Show the appropriate tab for the worksheet displayed
    Select Case ActiveSheet.name
        Case DataSheet.name:            ActivateTabGraphviz
        Case AboutSheet.name:           ActivateTabAbout
        Case ConsoleSheet.name:         ActivateTabConsole
        Case DiagnosticsSheet.name:     ActivateTabDiagnostics
        Case GraphSheet.name:           ActivateTabGraphviz
        Case HelpAttributesSheet.name:  ActivateTabLaunchpad
        Case HelpColorsSheet.name:      ActivateTabLaunchpad
        Case HelpShapesSheet.name:      ActivateTabLaunchpad
        Case LocaleDeDeSheet.name:      ActivateTabLaunchpad
        Case LocaleEnGbSheet.name:      ActivateTabLaunchpad
        Case LocaleEnUsSheet.name:      ActivateTabLaunchpad
        Case LocaleFrFrSheet.name:      ActivateTabLaunchpad
        Case LocaleItItSheet.name:      ActivateTabLaunchpad
        Case LocalePlPlSheet.name:      ActivateTabLaunchpad
        Case SettingsSheet.name:        ActivateTabLaunchpad
        Case StyleDesignerSheet.name:   ActivateTabStyleDesigner
        Case StylesSheet.name:          ActivateTabStyles
        Case SourceSheet.name:          ActivateTabSource
        Case SqlSheet.name:             ActivateTabSql
        Case SvgSheet.name:             ActivateTabSvg
        Case Else:                      ActivateTabGraphviz
    End Select
             
    Exit Sub

ErrorHandler:
    If Err.number > 0 Then
        Err.Clear
        Resume Next
    End If
End Sub

' ==========================================================================
' PROCEDURE: ribbon_getVisible
' PURPOSE:
'   A standard Ribbon callback that determines if a tab or control should
'   be displayed to the user.
'
' TECHNICAL WORKFLOW:
'   1. PREFERENCE LOOKUP: Consults the 'Settings' worksheet (via GetSettingBoolean)
'      to see which features the user has opted to enable.
'   2. PLATFORM ENFORCEMENT: specifically handles the 'SQL' tab:
'      - macOS: Forces visibility to False (SQL/ADO is Windows-only).
'      - Windows: Respects the user's visibility setting.
'   3. MODULAR UI: Allows the user to declutter their workspace by hiding
'      specialized tabs like 'Style Designer', 'Console', or 'Diagnostics'.
'   4. DEFAULT STATE: Any control not explicitly handled defaults to 'Visible = True'.
'
' USAGE:
'   - Triggered whenever 'myRibbon.Invalidate' is called.
'   - Essential for maintaining a clean, cross-platform interface.
' ==========================================================================
Public Sub ribbon_getVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Select Case control.id
        Case RIBBON_TAB_STYLE_DESIGNER
            visible = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_STYLE_DESIGNER)
        Case RIBBON_TAB_STYLES
            visible = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_STYLES)
        Case RIBBON_TAB_ABOUT
            visible = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_ABOUT)
        Case RIBBON_TAB_CONSOLE
            visible = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_CONSOLE)
        Case RIBBON_TAB_DIAGNOSTICS
            visible = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_DIAGNOSTICS)
        Case RIBBON_TAB_EXCHANGE
            visible = GetSettingBoolean(SETTINGS_TABS_TOGGLE_EXCHANGE)
        Case RIBBON_TAB_SOURCE
            visible = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_SOURCE)
        Case RIBBON_TAB_SQL
#If Mac Then
            visible = False
#Else
            visible = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_SQL)
#End If
        Case RIBBON_TAB_SVG
            visible = GetSettingBoolean(SETTINGS_TOOLS_TOGGLE_SVG)
        Case Else
            visible = True
    End Select
End Sub

' ==========================================================================
' PROCEDURE: RefreshRibbon
' PURPOSE:
'   Forces a full recalculation and redraw of all custom Ribbon controls.
'
' TECHNICAL WORKFLOW:
'   1. POINTER VALIDATION: Checks if the 'myRibbon' object still exists in
'      memory. If the VBA project has reset, it prompts the user to reload.
'   2. INVALIDATION: Calls the 'Invalidate' method on the IRibbonUI handle,
'      triggering every 'getVisible', 'getLabel', and 'getEnabled' callback.
'   3. ERROR RECOVERY: Catches scenarios where the object exists but is
'      unresponsive, providing a clear path for user remediation.
'
' USAGE:
'   - Called after updating settings (like changing languages or themes).
'   - Essential for "un-stucking" the UI if controls become unresponsive.
' ==========================================================================
Public Sub RefreshRibbon()
    On Error GoTo ErrorHandler
    If myRibbon Is Nothing Then
        ' This message cannot be localized due to error state.
        EmitMessage "Error refreshing the ribbon. Save and reopen this file."
    Else
        myRibbon.Invalidate
        If Err.number <> 0 Then
            ' This message cannot be localized due to error state.
            EmitMessage "Lost the Ribbon object. Save this file, close worksbook, and reopen."
        End If
    End If

    Exit Sub

ErrorHandler:
    If Err.number > 0 Then
        Err.Clear
        Resume Next
    End If
End Sub

' ==========================================================================
' PROCEDURE: InvalidateRibbonControl
' PURPOSE:
'   Triggers a redraw for a specific Ribbon control identified by its ID.
'
' TECHNICAL WORKFLOW:
'   1. POINTER VALIDATION: Verifies the 'myRibbon' object is still alive.
'      If the handle is lost, it provides feedback via the StatusBar.
'   2. PARTIAL INVALIDATION: Calls 'InvalidateControl', forcing Excel to
'      re-run the callbacks (getLabel, getEnabled, etc.) for only the
'      requested 'controlName'.
'   3. ERROR RESILIENCE: Silently traps errors to ensure that a UI glitch
'      does not interrupt the primary execution logic (like a SQL batch).
'
' USAGE:
'   - Used to enable/disable specific buttons (e.g., 'Reset Pool') after
'     a database operation completes.
'   - Ideal for high-frequency UI updates where a full refresh would be slow.
' ==========================================================================
Public Sub InvalidateRibbonControl(ByVal controlName As String)
    On Error GoTo ErrorHandler
    If myRibbon Is Nothing Then
        ' This message cannot be localized due to error state.
        UpdateStatusBar replace("Error updating the ribbon for control named '{controlName}'. Save and reopen this file.", "{controlName}", controlName)
    Else
        myRibbon.InvalidateControl controlName
    End If
ErrorHandler:
    If Err.number > 0 Then
        Err.Clear
        Resume Next
    End If
End Sub

' ==========================================================================
' PROCEDURE: ActivateTab
' PURPOSE:
'   Forces the Excel Ribbon to switch focus to a specific custom tab.
'
' TECHNICAL WORKFLOW:
'   1. POINTER VALIDATION: Checks the 'myRibbon' handle. If lost (VBA reset),
'      it routes a diagnostic message to the StatusBar instead of crashing.
'   2. FOCUS SHIFT: Invokes the 'ActivateTab' method using the unique string
'      ID of the target tab (e.g., "tabGraphviz").
'   3. ERROR ISOLATION: Employs a silent error trap to ensure that a UI
'      focus failure never halts the underlying data processing logic.
'
' USAGE:
'   - The core engine behind 'ribbon_activateTab'.
'   - Automates the "Context-Sensitive" UI experience as users navigate
'     between SQL, Data, and Styling worksheets.
' ==========================================================================
Public Sub ActivateTab(ByVal tabName As String)
    On Error GoTo ErrorHandler
    If myRibbon Is Nothing Then
        ' This message cannot be localized due to error state.
        UpdateStatusBar replace("Error activating a ribbon tab named '{tabName}'. Save and reopen this file.", "{tabName}", tabName)
    Else
        myRibbon.ActivateTab tabName
    End If
ErrorHandler:
    If Err.number > 0 Then
        Err.Clear
        Resume Next
    End If
End Sub

' ==========================================================================
' SECTION: DYNAMIC BUTTON CALLBACKS
' PURPOSE:
'   Ribbon Callbacks for prefixed ribbon buttons.
'   Synchronizes Ribbon button properties with the 'Settings' worksheet.
'
' TECHNICAL WORKFLOW:
'   1. SUFFIX BINDING: Uses standardized constants (VISIBLE, TEXT, SCREENTIP,
'      SUPERTIP) to construct the target Named Range for each control.
'   2. LIVE DATA RETRIEVAL:
'      - 'button_getLabel': Fetches the button's display name.
'      - 'button_getVisible': Controls visibility based on boolean toggles.
'      - 'button_getScreentip/Supertip': Populates hover-text for UI guidance.
'   3. WORKSHEET-DRIVEN UI: Allows the 'Settings' sheet to act as a
'      translation table; updating a cell on the sheet immediately updates
'      the UI upon the next 'RefreshRibbon' call.
' ==========================================================================

Public Sub button_getVisible(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = GetSettingBoolean(control.id & BUTTON_SUFFIX_VISIBLE)
End Sub

Public Sub button_getLabel(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = SettingsSheet.Range(control.id & BUTTON_SUFFIX_TEXT).value
End Sub

Public Sub button_getScreentip(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = SettingsSheet.Range(control.id & BUTTON_SUFFIX_SCREENTIP).value
End Sub

Public Sub button_getSupertip(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = SettingsSheet.Range(control.id & BUTTON_SUFFIX_SUPERTIP).value
End Sub

' ==========================================================================
' SECTION: TAB ACTIVATION WRAPPERS
' PURPOSE:
'   Provides a clean, named interface for switching between custom Ribbon tabs.
'
' TECHNICAL WORKFLOW:
'   1. ABSTRACTION: Wraps the low-level 'ActivateTab' call and its string-based
'      ID (e.g., RIBBON_TAB_SQL) into a dedicated procedure.
'   2. ASYNCHRONOUS SUPPORT: Designed to be called safely during sheet
'      activation events or 'Application.OnTime' initializations.
'
' USAGE:
'   - Primary consumers are the 'Worksheet_Activate' events and the
'     'ribbon_activateTab' router.
'   - Ensures that the Ribbon always reflects the active workspace
'     (e.g., selecting the SQL sheet automatically calls 'ActivateTabSql').
' ==========================================================================

Public Sub ActivateTabSql()
    ActivateTab RIBBON_TAB_SQL
End Sub

Public Sub ActivateTabSource()
    ActivateTab RIBBON_TAB_SOURCE
End Sub

Public Sub ActivateTabConsole()
    ActivateTab RIBBON_TAB_CONSOLE
End Sub

Public Sub ActivateTabSvg()
    ActivateTab RIBBON_TAB_SVG
End Sub

Public Sub ActivateTabDiagnostics()
    ActivateTab RIBBON_TAB_DIAGNOSTICS
End Sub

Public Sub ActivateTabStyleDesigner()
    ActivateTab RIBBON_TAB_STYLE_DESIGNER
End Sub

Public Sub ActivateTabStyles()
    ActivateTab RIBBON_TAB_STYLES
End Sub

Public Sub ActivateTabAbout()
    ActivateTab RIBBON_TAB_ABOUT
End Sub

Public Sub ActivateTabExchange()
    ActivateTab RIBBON_TAB_EXCHANGE
End Sub

Public Sub ActivateTabLaunchpad()
    ActivateTab RIBBON_TAB_WORKSHEETS
End Sub

Public Sub ActivateTabGraphviz()
    ActivateTab RIBBON_TAB_GRAPHVIZ
End Sub

