Attribute VB_Name = "modUtilityADODBConstants"
' Copyright (c) 2015-2025 Jeffrey J. Long. All rights reserved

'@Folder("Utility.Excel")

Option Explicit

''' Enums  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'''  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' Note to future Jeff...
' So here is the problem. We want to use ADODB capabilities for SQL commands and UTF-8 streams.
' If we use early binding the ADODB constants are available so long as a reference to the Microsoft
' ActiveX Data Objects x.x Library is added. This spreadsheet is used around the world by
' people with different versions of Excel all the way back to Excel 2007. It is not possible to
' specify an early binding reference for all users as it changes across versions of Excel.
' Instructions could be provided to tell the user how to go into Visual Basic and add the
' reference, but that is not ideal. The intent of this spreadsheet is to make things easy,
' not turn people into programmers, so early binding is out of the picture. We can use late
' binding without much difficulty, however since the objects are created at runtime there
' is no way for the compiler to know the enumeration values the objects use, and compiler
' errors occur. The work around is to use the values the enumerations represent. I don't like
' magic numbers, so I've replicated the enumerations for each object. Approximately 8 values are used total,
' but if this code ever needs to be changed, the other enum values are here. This is kind of
' a brittle solution, but these constants have been around for about 20 years so the risk
' of them being redefined is probably small.

' https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/commandtypeenum?view=sql-server-2017
Public Enum CommandTypeEnum
    adCmdUnspecified = -1        ' Does not specify the command type argument.
    adCmdText = 1                ' Evaluates CommandText as a textual definition of a command or stored procedure call.
    adCmdTable = 2               ' Evaluates CommandText as a table name whose columns are all returned by an internally generated SQL query.
    adCmdStoredProc = 4          ' Evaluates CommandText as a stored procedure name.
    adCmdUnknown = 8             ' Default. Indicates that the type of command in the CommandText property is not known.
    adCmdFile = 256              ' Evaluates CommandText as the file name of a persistently stored Recordset. Used with Recordset.Open or Requery only.
    adCmdTableDirect = 512       ' Evaluates CommandText as a table name whose columns are all returned. Used with Recordset.Open or Requery only. To use the Seek method, the Recordset must be opened with adCmdTableDirect.
End Enum

' https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/connectmodeenum?view=sql-server-ver15
Public Enum ConnectModeEnum
    adModeRead = 1               ' Indicates read-only permissions.
    adModeReadWrite = 3          ' Indicates read/write permissions.
    adModeRecursive = 4194304    ' = 0x400000, Used in conjunction with the other *ShareDeny* values (adModeShareDenyNone, adModeShareDenyWrite, or adModeShareDenyRead) to propagate sharing restrictions to all sub-records of the current Record. It has no affect if the Record does not have any children. A run-time error is generated if it is used with adModeShareDenyNone only. However, it can be used with adModeShareDenyNone when combined with other values. For example, you can use "adModeRead Or adModeShareDenyNone Or adModeRecursive".
    adModeShareDenyNone = 16     ' Allows others to open a connection with any permissions. Neither read nor write access can be denied to others.
    adModeShareDenyRead = 4      ' Prevents others from opening a connection with read permissions.
    adModeShareDenyWrite = 8     ' Prevents others from opening a connection with write permissions.
    adModeShareExclusive = 12    ' Prevents others from opening a connection.
    adModeUnknown = 0            ' Default. Indicates that the permissions have not yet been set or cannot be determined.
    adModeWrite = 2              ' Indicates write-only permissions.
End Enum

' https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/cursorlocationenum
Public Enum CursorLocationEnum
    adUseClient = 3              ' Uses client-side cursors supplied by a local cursor library.
    adUseNone = 1                ' Does not use cursor services. (This constant is obsolete and appears solely for the sake of backward compatibility.)
    adUseServer = 2              ' Default. Uses data-provider or driver-supplied cursors.
End Enum

' https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/cursortypeenum?view=sql-server-2017
Public Enum CursorTypeEnum
    adOpenDynamic = 2            ' Uses a dynamic cursor. Additions, changes, and deletions by other users are visible, and all types of movement through the Recordset are allowed, except for bookmarks, if the provider doesn't support them.
    adOpenForwardOnly = 0        ' Default. Uses a forward-only cursor. Identical to a static cursor, except that you can only scroll forward through records. This improves performance when you need to make only one pass through a Recordset.
    adOpenKeyset = 1             ' Uses a keyset cursor. Like a dynamic cursor, except that you can't see records that other users add, although records that other users delete are inaccessible from your Recordset. Data changes by other users are still visible.
    adOpenStatic = 3             ' Uses a static cursor, which is a static copy of a set of records that you can use to find data or generate reports. Additions, changes, or deletions by other users are not visible.
    adOpenUnspecified = -1       ' Does not specify the type of cursor.
End Enum

' https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/locktypeenum?view=sql-server-2017
Public Enum LockTypeEnum
    adLockBatchOptimistic = 4    ' Indicates optimistic batch updates. Required for batch update mode.
    adLockOptimistic = 3         ' Indicates optimistic locking, record by record. The provider uses optimistic locking, locking records only when you call the Update method.
    adLockPessimistic = 2        ' Indicates pessimistic locking, record by record. The provider does what is necessary to ensure successful editing of the records, usually by locking records at the data source immediately after editing.
    adLockReadOnly = 1           ' Indicates read-only records. You cannot alter the data.
    adLockUnspecified = -1       ' Does not specify a type of lock. For clones, the clone is created with the same lock type as the original.
End Enum

' https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/objectstateenum?view=sql-server-ver15
Public Enum ObjectStateEnum
    adStateClosed = 0            ' Indicates that the object is closed.
    adStateOpen = 1              ' Indicates that the object is open.
    adStateConnecting = 2        ' Indicates that the object is connecting.
    adStateExecuting = 4         ' Indicates that the object is executing a command.
    adStateFetching = 8          ' Indicates that the rows of the object are being retrieved.
End Enum

' https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/saveoptionsenum?view=sql-server-ver15
Public Enum SaveOptionsEnum
    adSaveCreateNotExist = 1     ' Default. Creates a new file if the file specified by the FileName parameter does not already exist.
    adSaveCreateOverWrite = 2    ' Overwrites the file with the data from the currently open Stream object, if the file specified by the Filename parameter already exists. If the file specified by the Filename parameter does not exist, a new file is created.
End Enum

' https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/streamtypeenum?view=sql-server-ver15
Public Enum StreamTypeEnum
    adTypeBinary = 1             ' Indicates binary data.
    adTypeText = 2               ' Default. Indicates text data, which is in the character set specified by Charset.
End Enum

' https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/executeoptionenum
Public Enum ExecuteOptionEnum
    adExecuteNoRecords = 128      ' Does not return records (e.g., for action queries)
    adExecuteRecord = 512         ' Returns a single record as a Record object
    adAsyncExecute = 16           ' Executes the command asynchronously
    adAsyncFetch = 32             ' Fetches remaining rows asynchronously
    adAsyncFetchNonBlocking = 64  ' Fetches rows asynchronously without blocking
    adOptionUnspecified = -1      ' Unspecified option
End Enum

