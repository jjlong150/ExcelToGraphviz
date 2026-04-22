Attribute VB_Name = "modWorksheetDiagnostics"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modWorksheetDiagnostics
' COPYRIGHT: Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER: Relationship Visualizer / Sheets / Diagnostics
'
' ROLE:
'   Provide the system-health auditing engine for the Relationship Visualizer.
'   Performs cross-platform discovery of environment state, Graphviz
'   connectivity, file-system readiness, and AppleScript bridge availability.
'   Populates the Diagnostics worksheet via the DIAGNOSTICS_ Named Range API.
'
' RESPONSIBILITIES:
'   - System auditing:
'       • ReportDiagnostics: collect OS, Excel, workbook, username, and
'         Graphviz version data
'       • Validate Temp, Font-cache, Color-cache, and image-path directories
'       • Populate DIAGNOSTICS_ fields with environment state
'
'   - macOS bridge validation:
'       • Confirm AppleScript sandbox folder
'       • Detect ExcelToGraphviz.applescript and query its version
'       • Clear Mac-specific fields on Windows
'
'   - Cache hygiene:
'       • DeleteFolderContents: purge directory contents cross-platform
'       • ClearFontImageFolder / ClearColorsImageFolder: reset Style Designer
'         preview caches
'
'   - Graphviz connectivity:
'       • GetGraphvizVersion: run "dot -V" via ExecuteAndCapture (Windows) or
'         RunAppleScriptTask (macOS)
'
'   - Worksheet maintenance:
'       • ClearDiagnostics: wipe stale diagnostic values prior to a new audit
'
' ARCHITECTURAL NOTES:
'   - Integrates tightly with DIAGNOSTICS_ Named Range API for structured
'     worksheet population.
'   - Uses OptimizeCode_Begin/End for UI-safe batch updates.
'   - Cross-platform logic ensures consistent behavior across Windows and macOS,
'     including delimiter differences and sandbox constraints.
'   - Diagnostics worksheet acts as a first-class troubleshooting surface for
'     environment validation and support workflows.
'
' USAGE:
'   - Ideal for setup verification, environment troubleshooting, and
'     pre-execution health checks before graph generation.
'
' RELATED WIKI PAGES:
'   - Diagnostics Worksheet Architecture
'   - Graphviz Class & Process Execution
'   - Style Designer Cache Management
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: ReportDiagnostics
'
' PURPOSE:
'   Collects and populates system-wide health and environment data into the
'   'Diagnostics' worksheet to assist in troubleshooting and setup verification.
'
' TECHNICAL WORKFLOW:
'   1. UI OPTIMIZATION: Disables events/screen updating via 'OptimizeCode_Begin'
'      and sets the 'xlWait' cursor.
'   2. SYSTEM INVENTORY: Records the Workbook name, OS version, and Excel
'      build number directly into the named-range API slots.
'   3. ENGINE VERIFICATION: Invokes 'GetGraphvizVersion' to confirm binary
'      connectivity and records local/system usernames.
'   4. DIRECTORY AUDIT: Validates the existence of critical paths:
'      - System Temp directory.
'      - Style Designer Font and Color image caches.
'      - External image paths defined by environment variables.
'   5. MAC SANDBOX VALIDATION (#If Mac):
'      - Hard-codes the required AppleScript safety path.
'      - Verifies the presence of 'ExcelToGraphviz.applescript'.
'      - Queries the script version via 'RunAppleScriptTask' to ensure
'        the bridge is functional.
'   6. CROSS-PLATFORM RESET (#Else): Clears Mac-specific diagnostic
'      fields when running on Windows.
'
' TECHNICAL NOTES:
'   - Layer: Settings & Diagnostics.
'   - Contract: Relies heavily on the 'DIAGNOSTICS_' Named Range API.
'   - DeepWiki Context: Documents the "Diagnostics Worksheet" page
'     referenced in the wiki.json.
' ==========================================================================
Public Sub ReportDiagnostics()
    ' Show the hourglass cursor
    Application.Cursor = xlWait
    DoEvents

    ' Turn off screen updating and events
    OptimizeCode_Begin
    
    ' Current Workbook File Name
    DiagnosticsSheet.Range(DIAGNOSTICS_WORKBOOK_NAME).value = ThisWorkbook.name
    
    ' Operating System
    DiagnosticsSheet.Range(DIAGNOSTICS_APPLICATION_OPERATING_SYSTEM).value = Application.OperatingSystem
    
    ' Excel version and build number
    DiagnosticsSheet.Range(DIAGNOSTICS_APPLICATION_VERSION).value = Application.version & Application.Build
    
     ' Graphviz version number
    DiagnosticsSheet.Range(DIAGNOSTICS_GRAPHVIZ_VERSION).value = GetGraphvizVersion
   
    ' User name as seen by Excel Application
    DiagnosticsSheet.Range(DIAGNOSTICS_APPLICATION_USER_NAME).value = Application.username
    
    ' User name as returned by OS
    DiagnosticsSheet.Range(DIAGNOSTICS_USERNAME).value = GetUsername()

    ' Temp file directory
    DiagnosticsSheet.Range(DIAGNOSTICS_TEMP_DIRECTORY).value = GetTempDirectory()
        If DirectoryExists(GetTempDirectory()) Then
        DiagnosticsSheet.Range(DIAGNOSTICS_TEMP_DIRECTORY_EXISTS).value = 1
    Else
        DiagnosticsSheet.Range(DIAGNOSTICS_TEMP_DIRECTORY_EXISTS).value = 0
    End If

    ' Style Designer Image Cache Directory of font preview images
    DiagnosticsSheet.Range(DIAGNOSTICS_FONT_IMAGE_DIR).value = GetFontImageDir()
    If DirectoryExists(GetFontImageDir()) Then
        DiagnosticsSheet.Range(DIAGNOSTICS_FONT_IMAGE_DIR_EXISTS).value = 1
    Else
        DiagnosticsSheet.Range(DIAGNOSTICS_FONT_IMAGE_DIR_EXISTS).value = 0
    End If
    
    ' Style Designer Image Cache Directory of color scheme preview images
    DiagnosticsSheet.Range(DIAGNOSTICS_COLOR_IMAGE_DIR).value = GetColorImageDir()
    If DirectoryExists(GetColorImageDir()) Then
        DiagnosticsSheet.Range(DIAGNOSTICS_COLOR_IMAGE_DIR_EXISTS).value = 1
    Else
        DiagnosticsSheet.Range(DIAGNOSTICS_COLOR_IMAGE_DIR_EXISTS).value = 0
    End If
    
    ' Name of the environment variable which can be defined to point to a folder of images
    DiagnosticsSheet.Range(DIAGNOSTICS_IMAGE_ENV_VARIABLE_NAME).value = "ExcelToGraphvizImages"
    
    ' The folder of images pointed to by the environment variable
    DiagnosticsSheet.Range(DIAGNOSTICS_EXCELTOGRAPHVIZ_IMAGE_DIRECTORY).value = GetExcelToGraphvizImageDirectory()
    If DirectoryExists(GetExcelToGraphvizImageDirectory()) Then
        DiagnosticsSheet.Range(DIAGNOSTICS_EXCELTOGRAPHVIZ_IMAGE_DIRECTORY_EXISTS).value = 1
    Else
        DiagnosticsSheet.Range(DIAGNOSTICS_EXCELTOGRAPHVIZ_IMAGE_DIRECTORY_EXISTS).value = 0
    End If
    
    ' The directory paths to be searched for images when creating a graph
    DiagnosticsSheet.Range(DIAGNOSTICS_IMAGE_PATH).value = GetImagePath()
    If DirectoryExists(GetImagePath()) Then
        DiagnosticsSheet.Range(DIAGNOSTICS_IMAGE_PATH_EXISTS).value = 1
    Else
        DiagnosticsSheet.Range(DIAGNOSTICS_IMAGE_PATH_EXISTS).value = 0
    End If
              
#If Mac Then
    ' Security sandbox where applescript files must reside to be executed by AppleScriptTask command
    DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FOLDER).value = "/Users/" & GetUsername & "/Library/Application Scripts/com.microsoft.Excel"
    If DirectoryExists(DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FOLDER).value) Then
        DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FOLDER_EXISTS).value = 1
    Else
        DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FOLDER_EXISTS).value = 0
    End If

    ' Name of file containing the AppleScriptTask commands needed by the Excel version of Excel to Graphviz
    DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FILE).value = "ExcelToGraphviz.applescript"
    
    ' Was the file of AppleScriptTask commands found in the sandbox directory?
    Dim applescriptfile As String
    applescriptfile = DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FOLDER).value & "/" & DiagnosticsSheet.Range("Diagnostics.AppleScriptFile").value
    If FileExists(applescriptfile) Then
        DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FILE_EXISTS).value = 1
        ' Version of the AppleScriptTask commands
        DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_VERSION).value = RunAppleScriptTask("getVersion", vbNullString)
    Else
        DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FILE_EXISTS).value = 0
        DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_VERSION).value = vbNullString
    End If

#Else
    DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FOLDER).value = vbNullString
    DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FOLDER_EXISTS).value = 0
    DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FILE).value = vbNullString
    DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FILE_EXISTS).value = 0
    DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_VERSION).value = vbNullString
#End If
    
    ' Turn on screen updating and events
    OptimizeCode_End
    
    ' Reset the cursor back to the default
    Application.Cursor = xlDefault
End Sub

' ==========================================================================
' PROCEDURE: ClearDiagnostics
'
' PURPOSE:
'   Purges stale environment data and health check results from the
'   'Diagnostics' worksheet to prepare for a fresh system audit.
'
' TECHNICAL WORKFLOW:
'   1. SYSTEM CLEARANCE: Wipes the primary system info block (Workbook,
'      OS, Excel version, Usernames).
'   2. PATH CLEARANCE: Wipes the directory validation results (Temp, Font,
'      and Color cache status).
'
' TECHNICAL NOTES:
'   - Layer: Settings & Diagnostics.
'   - Warning: Unlike other modules, this uses hard-coded cell addresses
'     (D4:D15, D19:D21) rather than the Named Range API. Use caution if
'     manually restructuring the Diagnostics sheet.
' ==========================================================================
Public Sub ClearDiagnostics()
    DiagnosticsSheet.Range("D4:D15").ClearContents
    DiagnosticsSheet.Range("D19:D21").ClearContents
End Sub

' ==========================================================================
' PROCEDURE: DeleteFolderContents
'
' PURPOSE:
'   Purges all files within a specified directory to maintain disk hygiene
'   and clear cached assets.
'
' TECHNICAL WORKFLOW:
'   1. MAC EXECUTION (#If Mac): Uses the native 'Kill' command with a
'      wildcard (*) to remove files. Wrapped in 'On Error Resume Next' to
'      gracefully handle empty directories or locked files.
'   2. WINDOWS EXECUTION (#Else): Instantiates the 'Scripting.FileSystemObject'
'      to perform a bulk 'DeleteFile' operation. The 'True' parameter
'      forces the deletion of read-only files.
'
' TECHNICAL NOTES:
'   - Platform: Cross-Platform (Native Mac Kill / Windows FSO).
'   - DeepWiki Context: Crucial for managing the "Style Designer Image Cache."
' ==========================================================================
Private Sub DeleteFolderContents(ByVal folder As String)
#If Mac Then
    On Error Resume Next
    Kill folder & "/*"
    On Error GoTo 0
#Else
    Dim fileSystemObject As Object
    Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
    If Not fileSystemObject Is Nothing Then
        fileSystemObject.DeleteFile folder & "\*.*", True
        Set fileSystemObject = Nothing
    End If
#End If

End Sub

' ==========================================================================
' PROCEDURE: ClearFontImageFolder
' PURPOSE:
'   Clears the temporary font preview images generated by the Style Designer.
' TECHNICAL WORKFLOW:
'   1. Resolves the cache path via 'GetFontImageDir'.
'   2. Invokes 'DeleteFolderContents' to purge the directory.
' ==========================================================================
Public Sub ClearFontImageFolder()
    Dim folder As String
    folder = GetFontImageDir()
    DeleteFolderContents folder
End Sub

' ==========================================================================
' PROCEDURE: ClearColorsImageFolder
' PURPOSE:
'   Clears the temporary color swatch images generated by the Style Designer.
' TECHNICAL WORKFLOW:
'   1. Resolves the cache path via 'GetColorImageDir'.
'   2. Invokes 'DeleteFolderContents' to purge the directory.
' ==========================================================================
Public Sub ClearColorsImageFolder()
    Dim folder As String
    folder = GetColorImageDir()
    DeleteFolderContents folder
End Sub

' ==========================================================================
' FUNCTION: GetGraphvizVersion
'
' PURPOSE:
'   Queries the external Graphviz binary to retrieve its version signature,
'   serving as a primary "heartbeat" check for the rendering engine.
'
' TECHNICAL WORKFLOW:
'   1. MAC EXECUTION (#If Mac): Invokes 'RunAppleScriptTask' with the
'      "runDot" command and "-V" flag to bypass macOS shell limitations.
'   2. WINDOWS EXECUTION (#Else):
'      - Calls 'ExecuteAndCapture' to run the "dot -V" CLI command.
'      - Note: Graphviz traditionally outputs version info to 'stderr'.
'   3. CLEANUP: Strips 'vbNewLine' from the result to provide a clean,
'      single-line version string (e.g., "dot - graphviz version 12.0.0").
'
' TECHNICAL NOTES:
'   - Platform: Cross-Platform (AppleScript vs. Windows Pipe).
'   - DeepWiki Context: Foundational check for the "Graphviz Class & Process Execution" page.
' ==========================================================================
Public Function GetGraphvizVersion() As String
#If Mac Then
    GetGraphvizVersion = RunAppleScriptTask("runDot", "-V")
#Else
    Dim stdOut As String
    Dim stdErr As String
    ExecuteAndCapture "dot -V", stdOut, stdErr

    GetGraphvizVersion = replace(stdErr, vbNewLine, vbNullString)
#End If
End Function

' ==========================================================================
' PROCEDURE: TestGetGraphvizVersion
' PURPOSE:
'   Developer utility to print the Graphviz version string to the
'   Immediate Window, wrapped in pipes (|) to identify leading/trailing whitespace.
' ==========================================================================
Public Sub TestGetGraphvizVersion()
    Debug.Print "|" & GetGraphvizVersion() & "|"
End Sub

