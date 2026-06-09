Attribute VB_Name = "modUtilityADODBConnectionPool"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modUtilityADODBConnectionPool
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Utility / ADO SQL / Connection Pooling
'
' ROLE:
'   High-performance ADO connection-pooling subsystem for Excel and Access
'   data sources. Provides provider auto-detection, freshness validation,
'   retry-based resilience, and safe-mode fallback to ensure deterministic
'   SQL execution under Windows.
'
' RESPONSIBILITIES:
'   - Connection pooling:
'       o Maintain late-bound ADO connections keyed by file path
'       o Enforce freshness via timestamp-based TTL checks
'       o Provide safe-mode, non-pooled connections for recovery paths
'   - Provider negotiation:
'       o Auto-detect newest available OLEDB provider (ACE -> Jet fallback)
'   - Resilience and diagnostics:
'       o Retry-based open logic for transient locks and latency
'       o Defensive close logic to prevent lingering file handles
'       o Emit diagnostic telemetry for provider failures and stale handles
'   - Integration:
'       o SQL Engine (iterative SQL, enumeration, placeholder SQL, batch execution)
'       o Ribbon controls (pool reset, dev-mode toggles)
'       o Settings and Diagnostics worksheets
'
' ARCHITECTURAL NOTES:
'   - Late-bound ADO for cross-version compatibility.
'   - Windows-only subsystem; macOS SQL features degrade gracefully.
'   - Designed to mitigate Office/ACE instability and long-running handle issues.
'   - Works in concert with modUtilityADODBDiagnosticLogger and SQL engine modules.
'
' VERSION NOTES:
'   - v10.x: "Connection Pool" group added to address March 2025 Office update
'            causing ADO connections to slow down significantly.
'
' USAGE:
'   - Invoked by SQL execution pipeline via getConnection/returnConnection.
'   - Supports iterative SQL, recursive SQL, enumeration, and placeholder expansion.
'   - Ribbon "Reset Pool" button triggers full cleanup and provider re-evaluation.
'
' RELATED WIKI PAGES:
'   - SQL Engine & Connection Pooling
'   - Diagnostics & Environment Documentation
'   - SQL Worksheet Architecture
' =============================================================================


Option Explicit

#If Win32 Or Win64 Then
    Private Const MAX_CONN_OPEN_RETRIES As Long = 3
    Private Const RETRY_DELAY_MS As Long = 100
#End If

' Late-bound Scripting.Dictionary
'   key   = fileName (String)
'   value = Array(conn As ADODB.Connection, openedAt As Date)
Private ConnectionPool As Object

Private CachedProvider As String

' --- Public API ---

''
' INITIALIZER: Bootstraps the Scripting.Dictionary used for pooling.
' 1. Logic is restricted to Windows environments (Win32/Win64).
' 2. Ensures the pool exists before any SQL execution is attempted.
' Called by: InitializeRibbon and SQL execution entry points.
'
Public Sub InitializeConnectionPool()
#If Win32 Or Win64 Then
    If ConnectionPool Is Nothing Then
        Set ConnectionPool = CreateObject("Scripting.Dictionary")
    End If
#End If
End Sub

''
' FUNCTION: getConnection
' PURPOSE:
'   Retrieves an active ADO connection for the specified Excel or Access file.
'   Implements a high-performance "Retrieve-or-Create" pattern with pooling.
'
' TECHNICAL WORKFLOW:
'   1. VALIDATION: Verifies file existence and initializes the pool if null.
'   2. POOL CHECK: Looks for a cached handle. Discards if 'stale' (maxConnectionMinutes)
'      or if the connection state is closed/broken.
'   3. SCHEMA NEGOTIATION: Dynamically builds 'Extended Properties' based on
'      extension (.xlsx, .xlsm, .xlsb, .xls, .accdb).
'   4. DATA INTEGRITY: Forces IMEX=1 and ImportMixedTypes=Text to prevent
'      Excel driver "type-guessing" errors in mixed columns.
'   5. RESILIENCE: Implements a retry-loop (MAX_CONN_OPEN_RETRIES) with delays
'      to handle transient file-locks or network latency.
'   6. SAFE-MODE: Automatically falls back to a non-pooled "Safe Connection"
'      if the primary driver negotiation fails.
'
' PARAMETERS:
'   - fileName [String]: Path to the target database.
'   - maxConnectionMinutes [Long]: Freshness threshold for the pool.
' RETURNS:
'   - Late-bound ADODB.Connection (Object).
'
Public Function getConnection(ByVal fileName As String, ByVal maxConnectionMinutes As Long) As Object
#If Win32 Or Win64 Then
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Dim fso As Object
    Dim fileExtension As String
    Dim provider As String
    Dim properties As String
    Dim attempts As Long
    
    ' ------------------------------------------------------------
    ' Preconditions
    ' ------------------------------------------------------------
    If Len(fileName) = 0 Then
        err.Raise vbObjectError + 1001, , "Filename cannot be empty."
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(fileName) Then
        Dim message As String
        message = GetMessage("msgboxSqlFileNotFound")
        message = replace(message, "{filePath}", fileName)
        err.Raise vbObjectError + 1002, , message
    End If
    
    If ConnectionPool Is Nothing Then
        Set ConnectionPool = CreateObject("Scripting.Dictionary")
    End If
    
    ' ------------------------------------------------------------
    ' Determine provider and file type
    ' ------------------------------------------------------------
    fileExtension = LCase$(Mid$(fileName, InStrRev(fileName, ".") + 1))
    
    provider = DetectBestExcelProvider()
    If Len(provider) = 0 Then
        err.Raise vbObjectError + 9001, , "No suitable OLEDB provider found on this system."
    End If
    
    ' =============================================================================
    ' Connection String Properties for Microsoft ACE OLEDB Excel Driver
    ' Used for ADO queries against .xlsx files (Excel 2007 and later)
    ' =============================================================================
    
    ' Full Extended Properties string - assign this to your ADO Connection
    ' Example usage:
    '   conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFilePath & _
    '             ";Extended Properties=""" & properties & """;"
    '
    ' Breakdown of each setting (in the order they typically appear):
    '
    ' "Excel 12.0 Xml"
    '   - Specifies the Excel file format this connection targets
    '   - "Excel 12.0"  = Excel 2007->2019 / Microsoft 365 (.xlsx, .xlsm, etc.)
    '   - "Xml"         = indicates the modern Office Open XML format (not legacy .xls binary)
    '   - Use "Excel 8.0" instead if connecting to old .xls files (with Jet 4.0 provider)
    '
    ' "HDR=YES"
    '   - Header Row = YES -> Treats the **first row** of the used range as column headers
    '   - Field names in SQL become the values in row 1 (e.g. [Year], [Shell], etc.)
    '   - Alternatives:
    '       HDR=NO   -> No headers; columns become F1, F2, F3-> (useful for raw data)
    '       HDR=YES;IMEX=1 is the most common combo for structured tables
    '
    ' "IMEX=1"
    '   - IMport EXport mode = 1 (read-only import mode)
    '   - Critical setting! Tells the driver to **prefer text** over numeric/date when scanning
    '   - Without IMEX=1 (or IMEX=0), the driver usually guesses types aggressively
    '   - IMEX=2 is export mode (rarely used with Excel reading)
    '   - Most important when columns contain mixed numbers + text (like your 'future' issue)
    '
    ' "ImportMixedTypes=Text"
    '   - Forces the driver to treat **mixed-type columns as Text** (most reliable for mixed data)
    '   - Works in combination with IMEX=1
    '   - Without this, even with IMEX=1, the driver may still guess numeric if the first 8 rows are all numbers
    '   - Alternative (older style): IMEX=1 + Registry change (TypeGuessRows=0) -> but this is cleaner
    '
    ' =============================================================================
    ' Recommended full connection string patterns (pick one):
    '
    ' Modern .xlsx (most common today):
    '   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Path\To\File.xlsx;Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1;ImportMixedTypes=Text"""
    '
    ' Legacy .xls (Excel 97-2003):
    '   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Path\To\File.xls;Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"""
    '
    ' =============================================================================
    ' Quick troubleshooting checklist when data types are wrong:
    '   1. Verify IMEX=1 is present (missing = most common mistake)
    '   2. Make sure quotes are doubled correctly around the whole Extended Properties value
    '   3. Add a dummy text value in the first 8 data rows of mixed columns (temporary workaround)
    '   4. Format entire suspect columns as Text in Excel before saving
    '   5. Try HDR=NO to see raw data (helps diagnose)
    ' =============================================================================
    
    Select Case LCase$(fileExtension)
        Case "xlsx"
            ' This is the standard & recommended for .xlsx (XML-based, macros disabled)
            properties = "Excel 12.0 Xml;HDR=YES;IMEX=1;ImportMixedTypes=Text"
    
        Case "xlsb"
            ' .xlsb is the binary format -> use "Excel 12.0" (no "Xml")
            properties = "Excel 12.0;HDR=YES;IMEX=1;ImportMixedTypes=Text"
    
        Case "xlsm"
            ' .xlsm requires "Excel 12.0 Macro" to properly recognize macro-enabled workbooks
            '  (even when just reading data; prevents "external table not in expected format" errors in some cases)
            properties = "Excel 12.0 Macro;HDR=YES;IMEX=1;ImportMixedTypes=Text"
    
        Case "xls"
            ' Legacy binary .xls files use "Excel 8.0"
            ' (Note: IMEX=1 and ImportMixedTypes=Text are **not** supported/ignored by the ACE driver for .xls,
            ' but including them does no harm -> they're safely passed through)
            properties = "Excel 8.0;HDR=YES"
    
         Case "accdb"
            ' Access uses NO Extended Properties
            properties = vbNullString

         Case "mdb"
            ' Access 2000 uses NO Extended Properties
            properties = vbNullString

       Case Else
            err.Raise vbObjectError + 1003, , replace(GetMessage("msgboxSqlFileTypeNotSupported"), "{fileExtension}", fileExtension)
    End Select
    
    ' ------------------------------------------------------------
    ' Try to reuse pooled connection
    ' ------------------------------------------------------------
    If ConnectionPool.Exists(fileName) Then
        Dim pooledConn As Object
        Dim openedAt As Date
        
        GetPooledConnection fileName, pooledConn, openedAt
        
        ' Reject stale connections
        If Not IsConnectionFresh(openedAt, maxConnectionMinutes) Then
            LogDiagnostic "Connection stale -> discarding: " & fileName
            SafeCloseConnection pooledConn
            ConnectionPool.Remove fileName
        ElseIf IsConnectionValid(pooledConn) Then
            Set getConnection = pooledConn
            Exit Function
        Else
            LogDiagnostic "Connection invalid -> discarding: " & fileName
            SafeCloseConnection pooledConn
            ConnectionPool.Remove fileName
        End If
    End If
    
    ' ------------------------------------------------------------
    ' Create new connection
    ' ------------------------------------------------------------
    Set conn = CreateObject("ADODB.Connection")

    If fileExtension = "accdb" Or fileExtension = "mdb" Then
        ' -----------------------------
        ' ACCESS CONNECTION
        ' -----------------------------
        conn.provider = provider
        conn.ConnectionString = _
            "Provider=" & provider & ";" & _
            "Data Source=" & fileName & ";" & _
            "Persist Security Info=False;"
        conn.CursorLocation = CursorLocationEnum.adUseClient

    Else
        ' -----------------------------
        ' EXCEL CONNECTION
        ' -----------------------------
        conn.provider = provider
        conn.properties("Extended Properties").value = properties
        conn.CursorLocation = CursorLocationEnum.adUseClient
    End If

    ' ------------------------------------------------------------
    ' Normal open with retry
    ' ------------------------------------------------------------
    For attempts = 1 To MAX_CONN_OPEN_RETRIES
        On Error Resume Next
        
        If LCase$(fileExtension) = "accdb" Or LCase$(fileExtension) = "mdb" Then
            conn.Open
        Else
            conn.Open fileName
        End If
        
        If err.number = 0 Then Exit For
        
        LogDiagnostic err.Description & " - getConnection.Open failed for file: " & fileName, _
                      errorNumber:=err.number, attempt:=attempts

        err.Clear
        SleepMilliseconds RETRY_DELAY_MS
    Next attempts
    
    ' ------------------------------------------------------------
    ' Safe-mode fallback
    ' ------------------------------------------------------------
    If conn Is Nothing Or Not IsConnectionValid(conn) Then
        LogDiagnostic "Normal connection failed. Entering SAFE MODE for: " & fileName
        
        SafeCloseConnection conn
        Set conn = CreateSafeModeConnection(fileName, provider, properties)
        
        ' Do NOT pool safe-mode connections
        Set getConnection = conn
        Exit Function
    End If
    
    ' ------------------------------------------------------------
    ' Add to pool
    ' ------------------------------------------------------------
    ' Some organizations see ACE connections silently die after ~5 minutes.
    ' Normal, valid connection -> add to pool with "connection freshness" timestamp
    ConnectionPool.Add fileName, Array(conn, Now)
    Set getConnection = conn
    Exit Function

ErrorHandler:
    SafeCloseConnection conn
    err.Raise err.number, , "getConnection(): " & err.Description
#Else
    err.Raise vbObjectError + 1003, , "ADO is not supported on macOS."
#End If
End Function

''
' PROCEDURE: CleanupConnectionPool
' PURPOSE:
'   Systematically closes and destroys all active database handles.
'
' TECHNICAL WORKFLOW:
'   1. ITERATION: Loops through the Dictionary keys to access every pooled
'      handle.
'   2. SAFE-CLOSE: Calls SafeCloseConnection on each object to ensure
'      the database is released even if the connection is in an unstable state.
'   3. MEMORY RELEASE: Clears the Dictionary and sets the global
'      ConnectionPool reference to Nothing to reclaim system resources.
'
' USAGE:
'   - Called during Ribbon reset, Workbook_BeforeClose, or after fatal errors
'     to prevent 'File In Use' locks for the user.
'
Public Sub CleanupConnectionPool()
#If Win32 Or Win64 Then
    On Error Resume Next
    Dim key As Variant
    Dim conn As Object
    Dim entry As Variant
    
    If Not ConnectionPool Is Nothing Then
        For Each key In ConnectionPool.keys
            entry = ConnectionPool.item(key)
            Set conn = entry(0)
            SafeCloseConnection conn
        Next key
        ConnectionPool.RemoveAll
        Set ConnectionPool = Nothing
    End If
#End If
End Sub

''
' FUNCTION: GetConnectionCount
' PURPOSE:
'   Returns the current number of active database handles in the pool.
'
' TECHNICAL WORKFLOW:
'   1. STATE CHECK: Verifies if the ConnectionPool Dictionary has been initialized.
'   2. TELEMETRY: Retrieves the 'Count' property from the internal Dictionary.
'
' USAGE:
'   - Used by the Diagnostics worksheet to report system health.
'   - Used by the SQL Ribbon tab to indicate if any connections are currently 'hot'.
'
Public Function GetConnectionCount() As Long
#If Win32 Or Win64 Then
    If ConnectionPool Is Nothing Then
        GetConnectionCount = 0
    Else
        GetConnectionCount = ConnectionPool.count
    End If
#End If
End Function

' ==========================================================================
' SECTION: INTERNAL HEALTH CHECKS & DRIVER DETECTION
' ==========================================================================

''
' CONNECTION VALIDATOR: Performs a "heartbeat" check on a pooled handle.
' 1. State Verification: Ensures the connection is currently 'adStateOpen'.
' 2. Trivial Execution: Executes a dummy query (SELECT 1) to verify the
'    underlying database file hasn't been moved or locked externally.
' @returns Boolean: True if the connection is alive and responsive.
'
Private Function IsConnectionValid(ByVal conn As Object) As Boolean
#If Win32 Or Win64 Then
    On Error GoTo ErrorHandler
    If Not conn Is Nothing Then
        If conn.State = ObjectStateEnum.adStateOpen Then
            conn.Execute "SELECT 1", , ExecuteOptionEnum.adExecuteNoRecords
            IsConnectionValid = True
            Exit Function
        End If
    End If
ErrorHandler:
    IsConnectionValid = False
#Else
    IsConnectionValid = False
#End If
End Function

''
' PROCEDURE: SafeCloseConnection
' PURPOSE:
'   Defensively closes and destroys an ADO connection object.
'
' TECHNICAL WORKFLOW:
'   1. NULL CHECK: Verifies the object exists before attempting operations.
'   2. BITWISE STATE CHECK: Uses a bitwise AND comparison against 'adStateOpen'
'      to verify the connection is truly active before calling .Close.
'   3. GRACEFUL FAILURE: Employs 'On Error Resume Next' to prevent the
'      application from crashing if the database file is already locked or inaccessible.
'   4. MEMORY RECLAMATION: Explicitly sets the object to 'Nothing' to ensure
'      the COM handle is released by the Windows OS.
'
Private Sub SafeCloseConnection(ByRef cn As Object)
#If Win32 Or Win64 Then
    On Error Resume Next
    If Not cn Is Nothing Then
        If (cn.State And ObjectStateEnum.adStateOpen) = ObjectStateEnum.adStateOpen Then
            cn.Close
        End If
        Set cn = Nothing
    End If
    err.Clear
#End If
End Sub

''
' PROCEDURE: GetPooledConnection
' PURPOSE:
'   Extracts a cached connection and its metadata from the internal registry.
'
' TECHNICAL WORKFLOW:
'   1. DICTIONARY LOOKUP: Retrieves the Variant array associated with the
'      specified fileName key.
'   2. OBJECT RECOVERY: Unpacks the 'ADODB.Connection' object from the first
'      array element (index 0).
'   3. METADATA RECOVERY: Unpacks the 'openedAt' timestamp from the second
'      array element (index 1) for age verification.
'
' USAGE:
'   - Internal utility for the 'getConnection' workflow to facilitate
'     freshness checks before reusing a handle.
'
Private Sub GetPooledConnection(ByVal fileName As String, _
                                ByRef conn As Object, _
                                ByRef openedAt As Date)
    Dim entry As Variant
    entry = ConnectionPool.item(fileName)
    Set conn = entry(0)
    openedAt = entry(1)
End Sub

''
' FRESHNESS CHECK: Calculates the age of a pooled handle in minutes.
' Uses the '1440' constant (minutes in a day) to compare the 'openedAt'
' timestamp against the user-defined 'maxConnectionMinutes' threshold.
'
Private Function IsConnectionFresh(ByVal openedAt As Date, maxConnectionMinutes As Long) As Boolean
    IsConnectionFresh = ((Now - openedAt) * 1440) < maxConnectionMinutes
End Function

''
' SAFE-MODE GENERATOR: Creates an un-pooled, high-timeout connection.
' 1. Logic: Used as a fallback when standard pooling fails.
' 2. Configuration: Forces CommandTimeout and ConnectionTimeout to 0 (infinite)
'    to ensure critical data retrieval is prioritized over speed.
'
Private Function CreateSafeModeConnection(ByVal fileName As String, _
                                          ByVal provider As String, _
                                          ByVal properties As String) As Object
#If Win32 Or Win64 Then
    On Error GoTo SafeModeError
    
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    With conn
        .provider = provider
        .properties("Extended Properties").value = properties
        .CursorLocation = CursorLocationEnum.adUseClient
        .CommandTimeout = 0
        .ConnectionTimeout = 0
        .Open fileName
    End With
    
    Set CreateSafeModeConnection = conn
    Exit Function

SafeModeError:
    err.Raise err.number, , "Safe-mode connection failed: " & err.Description
#End If
End Function

''
' DRIVER AUTO-DETECTION: Identifies the newest OLEDB provider on the system.
' 1. Version Hunt: Prioritizes Microsoft ACE (16.0 -> 15.0 -> 12.0) for modern
'    Excel formats (.xlsx, .xlsm, .xlsb).
' 2. Legacy Fallback: Drops down to Jet 4.0 if ACE is missing.
' 3. Performance: Caches the result in 'CachedProvider' to prevent
'    redundant registry probes during the session.
'
Public Function DetectBestExcelProvider() As String
    ' Return cached result if already detected
    If Len(CachedProvider) > 0 Then
        DetectBestExcelProvider = CachedProvider
        Exit Function
    End If
    
    Dim providers As Variant
    Dim p As Variant
    
    providers = Array( _
        "Microsoft.ACE.OLEDB.16.0", _
        "Microsoft.ACE.OLEDB.15.0", _
        "Microsoft.ACE.OLEDB.12.0" _
    )
    
    For Each p In providers
        If ProviderExists(CStr(p)) Then
            CachedProvider = CStr(p)
            DetectBestExcelProvider = CachedProvider
            Exit Function
        End If
    Next p
    
    If ProviderExists("Microsoft.Jet.OLEDB.4.0") Then
        CachedProvider = "Microsoft.Jet.OLEDB.4.0"
        DetectBestExcelProvider = CachedProvider
        Exit Function
    End If
    
    CachedProvider = vbNullString
    DetectBestExcelProvider = vbNullString
End Function

''
' PROBE UTILITY: Tests for the physical presence of a database driver.
' Attempts to assign the provider to a late-bound connection object and
' catches the resulting error if the driver is not installed.
'
Private Function ProviderExists(ByVal providerName As String) As Boolean
#If Win32 Or Win64 Then
    On Error Resume Next
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.provider = providerName
    ProviderExists = (err.number = 0)
    err.Clear
    Set conn = Nothing
#Else
    ProviderExists = False
#End If
End Function



