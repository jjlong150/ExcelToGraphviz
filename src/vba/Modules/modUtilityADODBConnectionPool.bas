Attribute VB_Name = "modUtilityADODBConnectionPool"
' Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved
'
' Connection pool with:
' - Provider auto-detection
' - Timestamped freshness
' - Safe-mode fallback
' - Hardened open/close
' - Diagnostic logging hooks
'
'@Folder("Utility.Excel")

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

' ===========================
' Public API
' ===========================

Public Sub InitializeConnectionPool()
#If Win32 Or Win64 Then
    If ConnectionPool Is Nothing Then
        Set ConnectionPool = CreateObject("Scripting.Dictionary")
    End If
#End If
End Sub

' Returns a pooled or fresh connection to the given Excel file.
' Uses:
' - Provider auto-detection
' - Connection freshness
' - Safe-mode fallback
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
        Err.Raise vbObjectError + 1001, , "Filename cannot be empty."
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(fileName) Then
        Dim message As String
        message = GetMessage("msgboxSqlFileNotFound")
        message = replace(message, "{filePath}", fileName)
        Err.Raise vbObjectError + 1002, , message
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
        Err.Raise vbObjectError + 9001, , "No suitable OLEDB provider found on this system."
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
    '   - "Excel 12.0"  = Excel 2007?2019 / Microsoft 365 (.xlsx, .xlsm, etc.)
    '   - "Xml"         = indicates the modern Office Open XML format (not legacy .xls binary)
    '   - Use "Excel 8.0" instead if connecting to old .xls files (with Jet 4.0 provider)
    '
    ' "HDR=YES"
    '   - Header Row = YES ? Treats the **first row** of the used range as column headers
    '   - Field names in SQL become the values in row 1 (e.g. [Year], [Shell], etc.)
    '   - Alternatives:
    '       HDR=NO   ? No headers; columns become F1, F2, F3? (useful for raw data)
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
    '   - Alternative (older style): IMEX=1 + Registry change (TypeGuessRows=0) ? but this is cleaner
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
            ' .xlsb is the binary format ? use "Excel 12.0" (no "Xml")
            properties = "Excel 12.0;HDR=YES;IMEX=1;ImportMixedTypes=Text"
    
        Case "xlsm"
            ' .xlsm requires "Excel 12.0 Macro" to properly recognize macro-enabled workbooks
            '  (even when just reading data; prevents "external table not in expected format" errors in some cases)
            properties = "Excel 12.0 Macro;HDR=YES;IMEX=1;ImportMixedTypes=Text"
    
        Case "xls"
            ' Legacy binary .xls files use "Excel 8.0"
            ' (Note: IMEX=1 and ImportMixedTypes=Text are **not** supported/ignored by the ACE driver for .xls,
            ' but including them does no harm ? they're safely passed through)
            properties = "Excel 8.0;HDR=YES"
    
         Case "accdb"
            ' Access uses NO Extended Properties
            properties = vbNullString

         Case "mdb"
            ' Access 2000 uses NO Extended Properties
            properties = vbNullString

       Case Else
            Err.Raise vbObjectError + 1003, , replace(GetMessage("msgboxSqlFileTypeNotSupported"), "{fileExtension}", fileExtension)
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
        
        If Err.number = 0 Then Exit For
        
        LogDiagnostic Err.Description & " - getConnection.Open failed for file: " & fileName, _
                      errorNumber:=Err.number, attempt:=attempts

        Err.Clear
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
    Err.Raise Err.number, , "getConnection(): " & Err.Description
#Else
    Err.Raise vbObjectError + 1003, , "ADO is not supported on macOS."
#End If
End Function

' Cleans up all pooled connections.
Public Sub CleanupConnectionPool()
#If Win32 Or Win64 Then
    On Error Resume Next
    Dim key As Variant
    Dim conn As Object
    Dim entry As Variant
    
    If Not ConnectionPool Is Nothing Then
        For Each key In ConnectionPool.Keys
            entry = ConnectionPool.item(key)
            Set conn = entry(0)
            SafeCloseConnection conn
        Next key
        ConnectionPool.RemoveAll
        Set ConnectionPool = Nothing
    End If
#End If
End Sub

' Returns the number of pooled connections.
Public Function GetConnectionCount() As Long
#If Win32 Or Win64 Then
    If ConnectionPool Is Nothing Then
        GetConnectionCount = 0
    Else
        GetConnectionCount = ConnectionPool.count
    End If
#End If
End Function

' ===========================
' Internal helpers
' ===========================

' Returns True if the connection is open and can execute a trivial query.
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

' Safely closes and releases a connection.
Private Sub SafeCloseConnection(ByRef cn As Object)
#If Win32 Or Win64 Then
    On Error Resume Next
    If Not cn Is Nothing Then
        If (cn.State And ObjectStateEnum.adStateOpen) = ObjectStateEnum.adStateOpen Then
            cn.Close
        End If
        Set cn = Nothing
    End If
    Err.Clear
#End If
End Sub

' Extracts connection and timestamp from the pool entry.
Private Sub GetPooledConnection(ByVal fileName As String, _
                                ByRef conn As Object, _
                                ByRef openedAt As Date)
    Dim entry As Variant
    entry = ConnectionPool.item(fileName)
    Set conn = entry(0)
    openedAt = entry(1)
End Sub

' Returns True if the connection age is less than MAX_CONN_AGE_MINUTES.
Private Function IsConnectionFresh(ByVal openedAt As Date, maxConnectionMinutes As Long) As Boolean
    IsConnectionFresh = ((Now - openedAt) * 1440) < maxConnectionMinutes
End Function

' Creates a fresh, non-pooled "safe-mode" connection.
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
    Err.Raise Err.number, , "Safe-mode connection failed: " & Err.Description
#End If
End Function

' ===========================
' Provider auto-detection
' ===========================
' Returns the best available provider string for Excel files.
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

' Checks if a provider is installed by attempting to assign it to a connection.
Private Function ProviderExists(ByVal providerName As String) As Boolean
#If Win32 Or Win64 Then
    On Error Resume Next
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.provider = providerName
    ProviderExists = (Err.number = 0)
    Err.Clear
    Set conn = Nothing
#Else
    ProviderExists = False
#End If
End Function



