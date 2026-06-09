Attribute VB_Name = "modAppleScript"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modAppleScript
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Utility / macOS Interop
'
' ROLE:
'   Sandbox-safe gateway for executing AppleScript handlers from VBA. Provides
'   a unified, defensive wrapper around AppleScriptTask to enable file dialogs,
'   color pickers, font enumeration, path resolution, and Graphviz execution
'   on macOS-restoring feature parity with Windows.
'
' RESPONSIBILITIES:
'   - Route all macOS system calls through a single, registered AppleScript
'     file (ExcelToGraphviz.applescript) stored in the user-script sandbox.
'   - Provide RunAppleScriptTask as a hardened wrapper around AppleScriptTask,
'     returning vbNullString on failure to prevent VBA runtime errors.
'   - Support higher-level features that depend on AppleScript:
'       o Font enumeration (Style Designer)
'       o RGB color picker (Style Designer)
'       o File dialogs (Source Save, SQL datasource selection)
'       o Graphviz execution and path resolution
'       o Console availability checks
'
' INTERACTIONS:
'   - Ribbon Tabs: Launchpad (console visibility), Style Designer (color picker,
'                  font previews), Source (Save As), SQL (datasource selection).
'   - Utility Modules: modUtilityFileSystem (path handling), modUtilityString.
'   - External Script: ExcelToGraphviz.applescript (all handlers).
'
' CROSS-PLATFORM NOTES:
'   - macOS-only subsystem; excluded entirely on Windows via #If Mac.
'   - Script versioning (v1-v3+) determines feature availability (e.g., RGB
'     picker introduced in script v3).
'   - All failures return vbNullString to avoid Excel for Mac crash scenarios.
'
' ERROR HANDLING:
'   - Fully defensive: AppleScript exceptions, missing handlers, and sandbox
'     permission failures are trapped and returned as vbNullString.
'   - Calling modules must interpret vbNullString as a soft failure.
'
' RELATED WIKI PAGES:
'   - macOS Architecture & Sandbox Model
'   - AppleScript Integration
'   - Cross-Platform Parity (Windows vs macOS)
' =============================================================================

Option Explicit

#If Mac Then

Private Const APPLE_SCRIPT_FILE = "ExcelToGraphviz.applescript"

' ==========================================================================
' FUNCTION: RunAppleScriptTask
' PURPOSE:
'   Executes a specific handler within the external AppleScript file.
'   If an AppleScript command returns a non-zero value, an error is thrown.
'   This code is to ensure any use of AppleScriptTask is wrapped with error
'   handling, and all AppleScript tasks have been written within a single
'   script file.
'
' TECHNICAL WORKFLOW:
'   1. SANDBOX COMMUNICATION: Invokes 'AppleScriptTask', the modern Office
'      for Mac API for executing out-of-process scripts.
'   2. SCRIPT ROUTING: Targets 'ExcelToGraphviz.applescript', passing the
'      requested 'scriptHandler' and its arguments.
'   3. ERROR ISOLATION: Employs a 'GoTo taskError' trap to catch system-level
'      AppleScript failures (like missing files or permission denials).
'   4. SILENT FAIL-SAFE: Returns 'vbNullString' on failure, allowing the
'      calling function to handle the error gracefully without a popup.
' ==========================================================================
Public Function RunAppleScriptTask(ByVal scriptHandler As String, ByVal scriptParameterString As String)

On Error GoTo taskError
    
    RunAppleScriptTask = AppleScriptTask(APPLE_SCRIPT_FILE, scriptHandler, scriptParameterString)
    Exit Function
    
taskError:
    RunAppleScriptTask = vbNullString
End Function


#End If

