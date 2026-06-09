Attribute VB_Name = "modUtilityADODBDiagnosticLogger"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modUtilityADODBDiagnosticLogger
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Utility / ADO SQL / Message Logging
'
' ROLE:
'   Forensic diagnostic logger for the SQL engine. Captures structured ADO
'   failure reports, environment fingerprints, and execution metadata to a
'   persistent log file in the workbook directory. Provides deep telemetry
'   for troubleshooting SQL, ADO provider issues, and environment-specific
'   regressions.
'
' RESPONSIBILITIES:
'   - Toggle and expose global logging state for Ribbon and SQL subsystems.
'   - Write structured diagnostic entries (timestamp, error number, category,
'     attempt count, SQL text) to a local log file.
'   - Generate environment fingerprints including OS, Excel version, bitness,
'     macro security, trusted-location status, OneDrive/SharePoint detection,
'     and provider information.
'   - Support the SQL engine, connection pool, and ExecuteAndCapture by
'     recording failures that would otherwise be silent.
'
' ARCHITECTURAL NOTES:
'   - Windows-only subsystem (ADO and provider diagnostics are not available
'     on macOS).
'   - Fully defensive: all logging operations use On Error Resume Next to
'     avoid cascading failures.
'   - Log file creation and writes are resilient to missing directories,
'     locked files, and restricted environments.
'   - Integrates with modUtilityADODBConnectionPool, SQL engine modules,
'     and Ribbon diagnostics controls.
'
' USAGE:
'   - Called by SQL execution pipeline, connection pool, and Ribbon toggles.
'   - Provides developers with actionable telemetry for debugging SQL failures.
'
' RELATED WIKI PAGES:
'   - SQL Engine & Connection Pooling
'   - Diagnostics & Environment Fingerprinting
'   - Troubleshooting ADO Provider Failures
' =============================================================================

Option Explicit

Private loggingEnabled As Boolean
Private Const LOG_FILE_NAME As String = "Relationship Visualizer ADO Log.txt"

' ===========================
' Public entry point for logging
' ===========================

''
' PROCEDURE: SetLoggingEnabled
' PURPOSE:
'   Globally toggles the diagnostic logging engine.
'
' TECHNICAL WORKFLOW:
'   1. STATE UPDATE: Sets the private 'loggingEnabled' boolean.
'   2. PERSISTENCE: This state is typically linked to a Ribbon toggle
'      button or a setting on the 'Settings' worksheet.
'
' USAGE:
'   - Called by the 'Console' ribbon tab to enable or disable the
'     capture of deep technical ADO logs.
'
Public Sub SetLoggingEnabled(ByVal Enabled As Boolean)
    loggingEnabled = Enabled
End Sub

''
' FUNCTION: IsLoggingEnabled
' PURPOSE:
'   Retrieves the current activation status of the ADO diagnostic logger.
'
' TECHNICAL WORKFLOW:
'   1. READ STATE: Accesses the private 'loggingEnabled' boolean variable.
'   2. UI FEEDBACK: Returns the state to the Ribbon callback to determine
'      if the 'Logging' toggle button should appear as pressed (Selected).
'
' USAGE:
'   - Used by 'modRibbon.bas' to synchronize the visual state of the
'     ribbon controls with the underlying engine state.
'
Public Function IsLoggingEnabled() As Boolean
    IsLoggingEnabled = loggingEnabled
End Function

''
' THE PRIMARY LOGGER: Writes a structured diagnostic entry to disk.
' 1. Logic: Only executes if 'loggingEnabled' is True.
' 2. Metadata: Records timestamp, attempt count, error numbers, and categories.
' 3. Fingerprinting: Optionally appends a full system audit (GetEnvironmentFingerprint).
' @param message [String]: The high-level error description.
' @param sql [Optional String]: The SQL query being executed at the time of failure.
' @param includeFingerprint [Boolean]: If True, appends hardware/software specs.
'
Public Sub LogDiagnostic(ByVal message As String, _
                            Optional ByVal errorNumber As Long = 0, _
                            Optional ByVal errorCategory As String = vbNullString, _
                            Optional ByVal attempt As Long = 0, _
                            Optional ByVal sql As String = vbNullString, _
                            Optional ByVal includeFingerprint As Boolean = False)
    If Not loggingEnabled Then Exit Sub
    On Error Resume Next
    
    Dim fso As Object
    Dim ts As Object
    Dim logPath As String
    
    logPath = GetLogFilePath()
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(logPath, 8, True) ' ForAppending
    
    ts.WriteLine String(80, "-")
    ts.WriteLine vbCrLf & format$(Now, "yyyy-mm-dd hh:nn:ss") & "   : " & message
    
    If attempt > 0 Then
        ts.WriteLine "  Attempt Number      : " & attempt
    End If
    
    If errorNumber <> 0 Then
        ts.WriteLine "  Error Number        : " & errorNumber
    End If
    
    If Trim(errorCategory) <> vbNullString Then
        ts.WriteLine "  Error Category      : " & errorCategory
    End If
    
    If sql <> vbNullString Then
        ts.WriteLine "  SQL Statement       : " & vbCrLf
        ts.WriteLine sql & vbCrLf
    End If
    
    If includeFingerprint Then
        ts.WriteLine GetEnvironmentFingerprint()
    End If
    
    ts.Close
End Sub

''
' THE FINGERPRINT ENGINE: Performs a deep-dive audit of the host environment.
' Captures critical debugging data:
' - Hardware: Processor count and architecture.
' - Software: Application version, Build number, and VBA version (VBA7 vs legacy).
' - Security: Macro security level (High/Low) and Trusted Location status.
' - Context: Detects OneDrive/SharePoint paths which often cause ADO locks.
' @returns String: A multi-line technical report for the log file.
'
Private Function GetEnvironmentFingerprint() As String
    Dim s As String
    's = vbCrLf & "Environment" & vbCrLf
    
    ' User + Machine
    s = s & "  User Name           : " & Environ$("USERNAME") & vbCrLf
    s = s & "  Machine Name        : " & Environ$("COMPUTERNAME") & vbCrLf
    
    ' Hardware and OS
    s = s & "  OS Version          : " & Environ$("OS") & vbCrLf
    s = s & "  Processor Count     : " & Environ$("NUMBER_OF_PROCESSORS") & vbCrLf
    s = s & "  Processor Arch.     : " & Environ$("PROCESSOR_ARCHITECTURE") & vbCrLf
    
    ' Excel + VBA
    s = s & "  Application Name    : " & Application.name & vbCrLf
    s = s & "  Applicaton Version  : " & Application.version & vbCrLf
    s = s & "  Applicaton OS       : " & Application.OperatingSystem & vbCrLf
    
    ' Is MS Office 32 or 64 bit?
    s = s & "  32/64 bit           : " & GetOfficeBitness() & vbCrLf
    
#If VBA7 Then
    s = s & "  VBA Version         : VBA7" & vbCrLf
#Else
    s = s & "  VBA7                : VBA6 or earlier" & vbCrLf
#End If

    ' Security + Trust
    s = s & "  Macro Security      : " & GetMacroSecurityMode() & vbCrLf
    s = s & "  Trusted Location    : " & IsWorkbookInTrustedLocation(ThisWorkbook) & vbCrLf
    
    ' OS
    s = s & "  Locale              : " & Application.International(xlCountrySetting) & vbCrLf
    s = s & "  Time Zone           : " & format$(Now, "zzz") & vbCrLf
    
    ' Workbook context
    On Error Resume Next
    s = s & "  Workbook Path       : " & ThisWorkbook.FullName & vbCrLf
    s = s & "  AutoRecover Enabled : " & Application.AutoRecover.Enabled & vbCrLf
    s = s & "  On OneDrive         : " & IsWorkbookOnOneDrive(ThisWorkbook) & vbCrLf
    
    ' Provider info
    s = s & "  Provider            : " & DetectBestExcelProvider() & vbCrLf
    
    GetEnvironmentFingerprint = s
End Function

''
' FUNCTION: IsWorkbookOnOneDrive
' PURPOSE:
'   Determines if the active workbook is hosted on a cloud-synchronized path.
'
' TECHNICAL WORKFLOW:
'   1. PATH CAPTURE: Retrieves the workbook's physical file path.
'   2. HEURISTIC SCAN: Performs a case-insensitive search for "onedrive"
'      or "sharepoint" substrings within the path string.
'   3. STATE REPORT: Returns True if a cloud-based directory structure is detected.
'
' USAGE:
'   - Crucial for the 'Environment Fingerprint' report to diagnose ADO
'     connection failures caused by URI-based cloud paths vs. local paths.
'
Private Function IsWorkbookOnOneDrive(wb As Workbook) As Boolean
    On Error Resume Next
    Dim p As String
    p = LCase$(wb.path)
    
    If InStr(p, "onedrive") > 0 Then
        IsWorkbookOnOneDrive = True
    ElseIf InStr(p, "sharepoint") > 0 Then
        IsWorkbookOnOneDrive = True
    Else
        IsWorkbookOnOneDrive = False
    End If
End Function

' ==========================================================================
' SECTION: ENVIRONMENT RESOLUTION HELPERS
' ==========================================================================

''
' PATH RESOLVER: Determines the physical location for the diagnostic log file.
' 1. Logic: Prioritizes the directory of the current workbook.
' 2. Fallback: Uses 'CurDir$' if the workbook hasn't been saved yet.
' 3. Cross-Platform: Employs 'Application.pathSeparator' for Win/Mac compatibility.
'
Private Function GetLogFilePath() As String
    On Error Resume Next
    Dim basePath As String
    basePath = ThisWorkbook.path
    If Len(basePath) = 0 Then basePath = CurDir$
    GetLogFilePath = basePath & Application.pathSeparator & LOG_FILE_NAME
End Function

''
' BITNESS PROBE: Identifies if the host Excel application is 32-bit or 64-bit.
' Critical for ADO troubleshooting, as database drivers (ACE/Jet) must
' match the bitness of the Office installation.
'
Private Function GetOfficeBitness() As String
#If Win64 Then
    GetOfficeBitness = "64-bit"
#Else
    GetOfficeBitness = "32-bit"
#End If
End Function

''
' FUNCTION: GetExcelBuildNumber
' PURPOSE:
'   Retrieves the specific build number of the host Excel application.
'
' TECHNICAL WORKFLOW:
'   1. SYSTEM QUERY: Accesses the 'Application.Build' property.
'   2. ERROR HANDLING: Uses 'On Error Resume Next' to prevent failures
'      on legacy versions of Excel where this property might be restricted.
'
' USAGE:
'   - Incorporated into the 'Environment Fingerprint' report.
'   - Allows developers to identify version-specific regressions in ADO
'     or VBA behavior during remote troubleshooting.
'
Private Function GetExcelBuildNumber() As String
    On Error Resume Next
    GetExcelBuildNumber = Application.Build
End Function

''
' VERSION AUDIT: Captures the granular build and version numbers of Excel.
' Used to identify specific Office updates or service packs that may
' impact VBA or ADO stability.
'
Private Function GetExcelFullVersion() As String
    On Error Resume Next
    GetExcelFullVersion = Application.version & " (Build " & Application.Build & ")"
End Function

''
' TRUST AUDIT: Verifies if the file is in an Excel 'Trusted Location'.
' Important for ADO because untrusted files may have restricted access
' to external database drivers (ACE/Jet).
'
Private Function IsWorkbookInTrustedLocation(wb As Workbook) As Boolean
    On Error Resume Next
    
    Dim trust As Object
    Dim loc As Object
    Dim p As String
    
    p = LCase$(wb.path)
    
    Set trust = Application.FileDialog(msoFileDialogOpen).Application _
                    .CommandBars("File").Controls("Options") ' dummy to force load
    
    ' Loop trusted locations
    For Each loc In Application.TrustCenter.TrustedLocations
        If Len(loc.path) > 0 Then
            If InStr(LCase$(p), LCase$(loc.path)) = 1 Then
                IsWorkbookInTrustedLocation = True
                Exit Function
            End If
        End If
    Next loc
    
    IsWorkbookInTrustedLocation = False
End Function

''
' FUNCTION: GetMacroSecurityMode
' PURPOSE:
'   Identifies the active Excel Automation Security level.
'
' TECHNICAL WORKFLOW:
'   1. SYSTEM QUERY: Polls Application.AutomationSecurity.
'   2. MAPPING: Converts the MsoAutomationSecurity enum into descriptive strings:
'      - 1 (Low), 2 (Medium), 3 (High), 4 (Very High).
'
' USAGE:
'   - Injected into the 'Environment Fingerprint' during a diagnostic log.
'   - Helps distinguish between a code bug and an environment-level macro block.
'
Private Function GetMacroSecurityMode() As String
    On Error Resume Next
    
    Dim sec As Long
    sec = Application.AutomationSecurity
    
    Select Case sec
        Case 1: GetMacroSecurityMode = "Low"
        Case 2: GetMacroSecurityMode = "Medium"
        Case 3: GetMacroSecurityMode = "High"
        Case 4: GetMacroSecurityMode = "Very High"
        Case Else: GetMacroSecurityMode = "Unknown"
    End Select
End Function

