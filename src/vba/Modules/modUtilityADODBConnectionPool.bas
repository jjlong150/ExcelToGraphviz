Attribute VB_Name = "modUtilityADODBConnectionPool"
' Copyright (c) 2015-2025 Jeffrey J. Long. All rights reserved

'@Folder("Utility.Excel")

Option Explicit

' Global dictionary to store connections
Private ConnectionPool As Object ' Late-bound Scripting.Dictionary

' Initialize the connection pool
Public Sub InitializeConnectionPool()
    #If Win32 Or Win64 Then
        Set ConnectionPool = CreateObject("Scripting.Dictionary")
    #Else
        ' Connection pooling is not initialized on macOS because ADO is not supported.
    #End If
End Sub

' Get a connection from the pool or create a new one
Public Function getConnection(ByVal fileName As String) As Object ' Late-bound ADODB.Connection
    #If Win32 Or Win64 Then
        On Error GoTo ErrorHandler
        
        Dim conn As Object ' Late-bound ADODB.Connection
        Dim StrCon As String
        Dim fso As Object ' Late-bound Scripting.FileSystemObject
        
        ' Create FileSystemObject
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        ' Validate filename
        If Len(fileName) = 0 Then
            Err.Raise vbObjectError + 1001, , "Filename cannot be empty."
        End If
        
        ' Check if file exists
        If Not fso.FileExists(fileName) Then
            Err.Raise vbObjectError + 1002, , "File not found: " & fileName
        End If
        
        ' Check if connection pool is initialized (lazy creation of the connection pool).
        If ConnectionPool Is Nothing Then
            Set ConnectionPool = CreateObject("Scripting.Dictionary")
        End If
        
        ' Check if connection is already in the pool
        If ConnectionPool.Exists(fileName) Then
            Set conn = ConnectionPool.item(fileName)
            
            ' Verify if the connection is still valid
            If IsConnectionValid(conn) Then
                Set getConnection = conn
                Exit Function
            Else
                ' Remove invalid connection from pool
                ConnectionPool.Remove fileName
                Set conn = Nothing
            End If
        End If
        
        ' Determine the connection string settings based upon the file extension
        ' of the file we will executed the query against
        Dim fileExtension As String
        fileExtension = Right$(fileName, Len(fileName) - InStrRev(fileName, "."))
        
        Dim provider As String
        Dim properties As String
        Select Case LCase$(fileExtension)
            Case "xlsx"
                provider = "Microsoft.ACE.OLEDB.12.0;"
                properties = "Excel 12.0 Xml;HDR=YES;"
              
            Case "xlsb"
                provider = "Microsoft.ACE.OLEDB.12.0;"
                properties = "Excel 12.0;HDR=YES;"
              
            Case "xlsm"
                provider = "Microsoft.ACE.OLEDB.12.0;"
                properties = "Excel 12.0 Macro;HDR=YES;"
              
            Case "xls"
                provider = "Microsoft.ACE.OLEDB.12.0;"
                properties = "Excel 8.0;HDR=YES;"
                
            Case Else
                Err.Raise vbObjectError + 1003, , replace(GetMessage("msgboxSqlFileTypeNotSupported"), "{fileExtension}", fileExtension)
        End Select

        ' Establish connection to the file containing the relational data using
        ' late binding as we do not know which version of Excel this spreadsheet
        ' will be running on
        Set conn = CreateObject("ADODB.Connection")
        
        ' Specify connection options
        With conn
            .provider = provider
            .properties("Extended Properties").value = properties
            .CursorLocation = CursorLocationEnum.adUseClient
            .Open fileName
        End With

        ' Add to connection pool
        ConnectionPool.Add fileName, conn
        
        Set getConnection = conn
        Exit Function

ErrorHandler:
        If Not conn Is Nothing Then
            If conn.State = ObjectStateEnum.adStateOpen Then
                conn.Close
            End If
            Set conn = Nothing
        End If
        Err.Raise Err.number, , "getConnection Error: " & Err.Description
    #Else
        Err.Raise vbObjectError + 1003, , "ADO is not supported on macOS."
    #End If
End Function

' Helper function to check if a connection is valid
Private Function IsConnectionValid(ByVal conn As Object) As Boolean
    #If Win32 Or Win64 Then
        On Error GoTo ErrorHandler
        If Not conn Is Nothing Then
            If conn.State = ObjectStateEnum.adStateOpen Then
                ' Attempt a lightweight operation to test connection
                conn.Execute "SELECT 1", , ExecuteOptionEnum.adExecuteNoRecords
                IsConnectionValid = True
                Exit Function
            End If
        End If
ErrorHandler:
        IsConnectionValid = False
    #Else
        ' Connection validation is not supported on macOS.
        IsConnectionValid = False
    #End If
End Function

' Clean up all connections
Public Sub CleanupConnectionPool()
    #If Win32 Or Win64 Then
        On Error Resume Next
        Dim key As Variant
        Dim conn As Object ' Late-bound ADODB.Connection
        
        If Not ConnectionPool Is Nothing Then
            For Each key In ConnectionPool.Keys
                Set conn = ConnectionPool.item(key)
                If Not conn Is Nothing Then
                    If conn.State = ObjectStateEnum.adStateOpen Then
                        conn.Close
                    End If
                    Set conn = Nothing
                End If
            Next key
            ConnectionPool.RemoveAll
            Set ConnectionPool = Nothing
        End If
    #Else
        ' Connection pool cleanup is not performed on macOS because ADO is not supported."
    #End If
End Sub

Public Function GetConnectionCount() As Long
    #If Win32 Or Win64 Then
        If ConnectionPool Is Nothing Then
            GetConnectionCount = 0
        Else
            GetConnectionCount = ConnectionPool.count
        End If
    #End If
End Function

' Example usage of getConnection
Private Sub TestConnectionPoolSQL()
    TestConnectionPool "SELECT * FROM [lists$]"
End Sub

Private Sub TestConnectionPool(ByVal sqlStatement As String)
    #If Win32 Or Win64 Then
        On Error GoTo ErrorHandler
        Dim conn As Object ' Late-bound ADODB.Connection
        Dim rst As Object  ' Late-bound ADODB.Recordset
        Dim fileName As String
        
        ' Example: Use the current workbook
        fileName = ThisWorkbook.FullName
        
        ' Get connection from pool, timing how long it takes to get the connection.
        Dim timex As Stopwatch
        Set timex = New Stopwatch
        
        timex.start
        Set conn = getConnection(fileName)
        timex.stop_it
        
        Debug.Print "Getting a connection to " & fileName & " took " & timex.Elapsed_sec & " seconds" & vbNewLine
        Debug.Print "SQL Statement: " & sqlStatement & vbNewLine
        
        ' Example query
        Set rst = CreateObject("ADODB.Recordset")
        rst.Open sqlStatement, conn, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly
        
        ' Check if recordset has data
        If Not rst.EOF And Not rst.BOF Then
            Dim field As Object ' Late-bound ADODB.Field
            ' Print column headers
            Debug.Print "=== Recordset Contents ==="
            For Each field In rst.fields
                Debug.Print field.name & vbTab;
            Next field
            Debug.Print ' New line after headers
            
            ' Iterate through rows
            Dim rowCount As Long
            'Dim colCount As Long
            Dim i As Long
            rowCount = 0
            rst.MoveFirst
            While Not rst.EOF
                rowCount = rowCount + 1
                ' Print each field in the current row
                For i = 0 To rst.fields.count - 1
                    Debug.Print rst.fields(i).value & vbTab;
                Next i
                Debug.Print ' New line after each row
                rst.MoveNext
            Wend
            Debug.Print "=== End of Recordset (" & rowCount & " rows) ==="
        Else
            Debug.Print "No records found in [Sheet1$]."
        End If
        Debug.Print vbNewLine
        
        ' Clean up recordset
        rst.Close
        Set rst = Nothing
        
        ' Note: Do NOT close the connection here; it remains in the pool
        Exit Sub

ErrorHandler:
        If Not rst Is Nothing Then
            If rst.State = ObjectStateEnum.adStateOpen Then
                rst.Close
            End If
            Set rst = Nothing
        End If
        MsgBox "Error: " & Err.Description
    #Else
        MsgBox "This operation is not supported on macOS because ADO is not available."
    #End If
End Sub


