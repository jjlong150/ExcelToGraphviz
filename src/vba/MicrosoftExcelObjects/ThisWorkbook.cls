VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
' PROJECT:         Excel to Graphviz
' WORKBOOK MODULE: ThisWorkbook
' COPYRIGHT:       Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:           Relationship Visualizer / Application Lifecycle
'
' ROLE:
'   Central lifecycle controller for the Relationship Visualizer. Coordinates
'   environment initialization, cross-platform dependency validation, resource
'   provisioning, and teardown of system-level components such as ADO
'   connection pools and Ribbon-image caches.
'
' DESCRIPTION:
'   This workbook module orchestrates the startup and shutdown phases of the
'   application. On open, it caches environment variables, validates Graphviz
'   and AppleScript dependencies, prepares UI asset directories, and restores
'   worksheet visibility states. On close, it performs a controlled teardown of
'   Windows-only ADO connection pools to prevent orphaned sessions. The module
'   ensures that both Windows and macOS follow a unified initialization path
'   while respecting platform-specific constraints.
'
' KEY RESPONSIBILITIES:
'   - ENVIRONMENT BOOTSTRAPPING:
'       o Caches temporary directory and username for high-frequency access
'       o Minimizes repeated system calls during rendering and SQL operations
'
'   - GRAPHVIZ PIPELINE VALIDATION:
'       o Windows: auto-detects dot.exe, normalizes paths, validates existence
'       o macOS: verifies AppleScriptTask bridge version and Excel version parity
'
'   - RESOURCE MANAGEMENT:
'       o Creates cache directories for Ribbon color/font images
'       o Tears down ADO connection pools on close (Windows only)
'
'   - WORKSPACE SYNCHRONIZATION:
'       o Restores worksheet visibility based on saved settings
'       o Ensures UI contract consistency across sessions
'
'   - PLATFORM UNIFICATION:
'       o Maintains a single codebase for Windows (32/64-bit) and macOS
'       o Uses conditional compilation for platform-specific logic
'
' USAGE:
'   - Fires automatically on workbook open and close
'   - Ensures the application starts in a valid, fully prepared state
'   - Guarantees clean shutdown of platform-specific resources
'
' RELATED WIKI PAGES:
'   - Application Lifecycle & Initialization
'   - Graphviz Pipeline Validation
'   - AppleScriptTask Bridge (macOS)
'   - ADO Connection Pool Architecture
'   - Worksheet Visibility Contract
' =============================================================================


Option Explicit

' ==========================================================================
' PROCEDURE: Workbook_BeforeClose
' PURPOSE:
'   Triggers the graceful teardown of persistent database handles.
'
' TECHNICAL WORKFLOW:
'   1. PLATFORM CHECK: Uses conditional compilation (#If Not Mac) as the
'      ADO connection pool is a Windows-specific feature.
'   2. RESOURCE RECLAMATION: Calls 'CleanupConnectionPool' to explicitly
'      close and destroy all cached ADO connection objects.
'
' USAGE:
'   - Fires automatically when the user clicks the 'X' or closes the file.
'   - Prevents "orphaned" sessions on the data source (SQL Server, Access, etc.).
' ==========================================================================
Private Sub Workbook_BeforeClose(Cancel As Boolean)
#If Not Mac Then
    ' Release any connections which may have been cached in the connection pool.
    CleanupConnectionPool
#End If
End Sub

' ==========================================================================
' PROCEDURE: Workbook_Open
' PURPOSE:
'   Orchestrates the environment setup and dependency verification.
'
' TECHNICAL WORKFLOW:
'   1. ENVIRONMENT CACHING: Sets the temporary directory and retrieves the
'      username once to avoid expensive system calls during runtime.
'   2. WINDOWS VALIDATION (#If Not Mac):
'      - Auto-detects 'dot.exe' on the system PATH if no manual path is set.
'      - Normalizes and validates the Graphviz path to prevent rendering failures.
'   3. MACOS VALIDATION (#If Mac):
'      - Verifies the 'ExcelToGraphviz.applescript' version to ensure API parity.
'      - Checks Excel version compatibility (16.40+) to ensure VBA stability.
'   4. RESOURCE PREPARATION: Creates cache directories for Font and Color
'      ribbon icons to optimize UI performance.
'   5. UI SYNCHRONIZATION: Enforces worksheet visibility based on saved settings.
' ==========================================================================
Private Sub Workbook_Open()
    Application.Calculation = xlCalculationAutomatic
    
    ' Establish the temporary directory once so we don't make frequent system level calls to retrieve
    ' environment variable values which slow things down on Windows PCs.
    SetTempDirectory
    
    ' Cache the username (retrieving it causes it to be cached).
    If GetUsername() = vbNullString Then
        EmitMessage GetMessage("msgboxUnableToGetUsername")
    End If

    ' Check that Graphviz is installed
#If Not Mac Then
    Dim gvPath As String
    gvPath = Trim$(SettingsSheet.Range(SETTINGS_GV_PATH))
    If gvPath = vbNullString Then
        ' Locate the graphing program on the path
        Dim engine As String
        engine = FindFileOnPath("dot.exe")
        If engine = vbNullString Then
            AlertGraphvizNotFound "dot"
        End If
    Else
        ' Fix path if it does not end with a path separator
        If Not EndsWith(gvPath, Application.pathSeparator) Then
            gvPath = gvPath & Application.pathSeparator
            SettingsSheet.Range(SETTINGS_GV_PATH).value = gvPath
        End If
        
        ' See if the full path to dot.exe exists
        gvPath = gvPath & "dot.exe"
        If Not FileExists(gvPath) Then
            EmitMessage replace(GetMessage("msgboxInvalidDotPath"), "{gvPath}", gvPath)
        End If
    End If
#End If

#If Mac Then
    ' Establish the version of "ExcelToGraphviz.applescript" which this spreadsheet requires.
    Dim currentVersion As Long
    currentVersion = 3
    
    ' Check that AppleScriptTask scripts are present, and retrieve the version number
    Dim scriptVersion As String
    scriptVersion = RunAppleScriptTask("getVersion", vbNullString)
    
    If scriptVersion = vbNullString Then
        EmitMessage GetMessage("msgboxScriptsNotFound")
        SetScriptVersion 0
    Else
        ' Save the version number for future reference. AppleScriptTask has slow performance, so
        ' we want to minimize how often we call it.
        Dim yourVersion As Long
        yourVersion = CInt(scriptVersion)
        SetScriptVersion yourVersion
        
        ' Give a warning if an old version of "ExcelToGraphviz.applescript" is installed.
        If yourVersion < currentVersion Then
            Dim msg As String
            msg = GetMessage("msgboxVersionMismatch")
            msg = replace(msg, "{currentVersion}", CStr(currentVersion))
            msg = replace(msg, "{yourVersion}", CStr(yourVersion))
            EmitMessage msg
        End If
    End If

    ' Check that you are running an version of Excel with VBA support
    Dim requiredMajor As Long: requiredMajor = 16
    Dim requiredMinor As Long: requiredMinor = 40
    
    Dim actualVersion As String
    Dim actualMajor As Long, actualMinor As Long
    Dim versionParts As Variant
    
    actualVersion = Application.version
    versionParts = split(actualVersion, ".")
    
    If UBound(versionParts) >= 1 Then
        actualMajor = CLng(versionParts(0))
        actualMinor = CLng(versionParts(1))
        
        If (actualMajor < requiredMajor) Or _
           (actualMajor = requiredMajor And actualMinor < requiredMinor) Then
            msg = GetMessage("msgboxIncompatibleExcelVersion")
            msg = replace(msg, "{requiredExcelVersion}", requiredMajor & "." & requiredMinor)
            msg = replace(msg, "{yourExcelVersion}", actualVersion)
            EmitMessage msg
        End If
    End If
#End If

    ' Create directories where html, json, and images of colors and fonts can be cached
    ' This directory is created when the spreadsheet is launched so we don't
    ' waste cycles attempting to create it for every ribbon image of a color or font.
    CreateColorImageDir
    CreateFontImageDir
    CreateHtmlTempDir
    
    ' Version 11 introduced a semaphore file to signal when gallery images are current
    ClearStaleImageCaches
    
    ' Sync the worksheet visibility with the saved values (in case someone hides
    ' worksheets manually).
    ShowOrHideWorksheets
    
    ' Settings are cached using lazy intialization. This call indicates the settings
    ' need to be loaded upon first reference.
    InvalidateSettings
End Sub

' ==========================================================================
' ROUTINE: ClearStaleImageCaches
'
' PURPOSE:
'   Ensures that Style Designer image galleries (color palettes and font
'   previews) remain synchronized with the current workbook version. Detects
'   stale or missing cache folders via a lightweight semaphore file
'   (version.properties) and performs a targeted cleanup when necessary.
'   Guarantees that UI previews always reflect the latest assets without
'   requiring manual user intervention.
'
' FUNCTIONAL WORKFLOW:
'   1. SEMAPHORE RESOLUTION:
'        - Computes the full path to version.properties inside each
'          image-cache directory (colors and fonts).
'        - Uses Application.pathSeparator for cross-platform correctness.
'
'   2. CACHE VALIDATION:
'        - If version.properties is missing:
'             o Treats the cache as stale or uninitialized.
'             o Clears the corresponding image folder.
'
'   3. SEMAPHORE REPAIR:
'        - Writes a fresh version.properties file containing:
'             version=11.0
'          establishing the cache version expected by the current workbook.
'
'   4. OUTPUT:
'        - Leaves both image-cache directories in a known-good state with
'          synchronized version markers.
'
' TECHNICAL NOTES:
'   - version.properties acts as a simple, fast, cross-platform semaphore
'     indicating that the cache was built by the current workbook version.
'   - Clearing only when the semaphore is missing avoids unnecessary
'     regeneration and improves performance during Style Designer operations.
'   - DeepWiki Context: Implements the cache-validation rules described in
'     the "Style Designer", "Image Assets", and "Cache Synchronization"
'     documentation.
' ==========================================================================
Private Sub ClearStaleImageCaches()
    Dim semaphoreFile As String
    Dim sep As String
    sep = Application.pathSeparator

    ' Style Designer colorscheme galleries
    semaphoreFile = GetColorImageDir & sep & "version.properties"
    If Not FileExists(semaphoreFile) Then
        ClearColorsImageFolder
        WriteTextToFile "version=11.0", semaphoreFile
    End If

    ' Style Designer font preview image gallery
    semaphoreFile = GetFontImageDir & sep & "version.properties"
    If Not FileExists(semaphoreFile) Then
        ClearFontImageFolder
        WriteTextToFile "version=11.0", semaphoreFile
    End If
End Sub

