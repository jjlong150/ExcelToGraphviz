Attribute VB_Name = "modUtilityGraphviz"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modUtilityGraphviz
' COPYRIGHT: Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Bootstrap / Graphviz Presence
'
' ROLE:
'   Provide a single, centralized alert routine for notifying users when a
'   Graphviz engine cannot be located on the system. Ensures consistent,
'   localized messaging across all Data-sheet-driven workflows.
'
' RESPONSIBILITIES:
'   - AlertGraphvizNotFound:
'       • Emit a localized message indicating that the requested Graphviz
'         engine (dot, neato, fdp, etc.) was not found
'       • Perform platform-specific handling (Windows implemented; macOS
'         placeholder pending port)
'
' ARCHITECTURAL NOTES:
'   - Windows implementation uses EmitMessage + GetMessage/GetLabel to ensure
'     consistent localization and UI behavior.
'   - macOS branch currently stubbed with TODO for future parity.
'   - Consumed by Data-sheet helpers, SQL engine, and any workflow that invokes
'     external Graphviz executables.
'
' USAGE:
'   - Called upon workbook startup to establish Graphviz installation occurred.
'
' RELATED WIKI PAGES:
'   - Graphviz Installation & PATH Requirements
'   - Error Messaging & Localization
' =============================================================================

Option Explicit

Public Sub AlertGraphvizNotFound(ByVal graphEngine As String)
#If Mac Then
    'TODO Port
#Else
    EmitMessage replace(GetMessage("msgboxGraphvizNotFound"), "{graphEngine}", graphEngine)
#End If
End Sub


