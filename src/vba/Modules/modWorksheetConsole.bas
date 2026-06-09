Attribute VB_Name = "modWorksheetConsole"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modWorksheetConsole
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Relationship Visualizer / Sheets / Console
'
' ROLE:
'   Manage the on-sheet Console subsystem used for real-time diagnostics,
'   Graphviz CLI logging, verbose-mode control, and cross-platform export.
'   Provides a persistent audit trail of commands, stdout/stderr output,
'   and internal status messages.
'
' RESPONSIBILITIES:
'   - Console lifecycle:
'       o ClearConsoleWorksheet: purge prior command/output history
'       o SaveConsoleToFile / ConsoleWorksheetToFile: export logs to UTF-8
'       o CopyConsoleToClipboard (Windows): aggregate and copy log text
'
'   - Logging:
'       o DisplayTextOnConsoleWorksheet: log CLI command + parsed output
'       o LogToConsoleWorksheet: append raw diagnostic text
'       o Platform-aware delimiter handling (vbCr on macOS, vbLf on Windows)
'       o High-speed bulk writes using Application.Transpose
'
'   - Verbose-mode logic:
'       o RunGraphvizInVerboseMode: enable Graphviz "-v" only when
'         Console is visible and user settings permit verbose logging
'
' ARCHITECTURAL NOTES:
'   - Integrates with SETTINGS_ Named Range API for append/overwrite,
'     verbose mode, and log-to-console toggles.
'   - macOS uses AppleScriptTask for Save As dialogs and command logging
'     normalization; Windows uses native dialogs and clipboard APIs.
'   - ConsoleSheet is treated as a UI surface and a diagnostic buffer,
'     optimized for minimal overhead during graph generation.
'
' USAGE:
'   - Ideal for capturing Graphviz engine behavior, debugging DOT output,
'     and providing transparent audit trails during graph generation.
'
' RELATED WIKI PAGES:
'   - DOT Source Viewer & Console Architecture
'   - Logging & Diagnostics Pipeline
'   - Cross-Platform Execution Model
' =============================================================================


Option Explicit

' ==========================================================================
' PROCEDURE: ClearConsoleWorksheet
'
' PURPOSE:
'   Purges all previous Graphviz execution logs and command history from the
'   'Console' worksheet to provide a clean state for new diagnostic data.
'
' TECHNICAL WORKFLOW:
'   1. BOUNDARY DETECTION: Identifies the 'lastRow' of data using the
'      worksheet's 'UsedRange'.
'   2. BULK REMOVAL: Defines a dynamic range spanning columns A and B
'      (Command and Message columns) and executes '.ClearContents'.
'
' TECHNICAL NOTES:
'   - Layer: UI / Diagnostics (Console).
'   - DeepWiki Context: Implements the "Console Architecture" described
'     in the 'DOT Source Viewer & Console' page.
' ==========================================================================
Public Sub ClearConsoleWorksheet()
    
    ' Determine the range of the cells which need to be cleared
    Dim lastRow As Long
    With ConsoleSheet.UsedRange
        lastRow = .Cells.item(.Cells.count).row
    End With

    ' Remove any existing content
    Dim cellRange As String
    cellRange = "A1:B" & lastRow
    ConsoleSheet.Range(cellRange).ClearContents
    
End Sub

' ==========================================================================
' PROCEDURE: DisplayTextOnConsoleWorksheet
'
' PURPOSE:
'   Logs the external Graphviz CLI command and its resulting output (stdout/stderr)
'   to the 'Console' worksheet for real-time debugging and audit trails.
'
' TECHNICAL WORKFLOW:
'   1. PRE-FLIGHT CHECK: Evaluates 'SETTINGS_LOG_TO_CONSOLE' via the Settings
'      sheet; exits immediately if logging is disabled.
'   2. STATE MANAGEMENT: Conditionally invokes 'ClearConsoleWorksheet' unless
'      'SETTINGS_APPEND_CONSOLE' is enabled.
'   3. COMMAND LOGGING:
'      - Records the executed command string with a ">" prefix in Column A.
'      - MAC LOGIC: Injects a "dot " prefix to the command to simulate the
'        AppleScript-executed CLI call for consistent logging.
'   4. DATA PARSING: Splits the 'textBlob' based on platform delimiters
'      (vbCr for macOS, vbLf for Windows).
'   5. BULK TRANSFER: Uses 'Application.Transpose' to write the parsed log
'      lines into Column B as a single range operation for maximum performance.
'
' TECHNICAL NOTES:
'   - Platform: Cross-Platform (delimiters and command faking).
'   - Layer: UI / Diagnostics (Console).
'   - Contract: Adheres to the 'SETTINGS_' Named Range API for logging behavior.
' ==========================================================================
Public Sub DisplayTextOnConsoleWorksheet(ByVal dotCommand As String, ByVal textBlob As String)

    ' Exit if logging is disabled
    If Not GetCellBoolean(SettingsSheet.name, SETTINGS_LOG_TO_CONSOLE) Then
        Exit Sub
    End If

    ' Clear console if not in append mode
    If Not GetCellBoolean(SettingsSheet.name, SETTINGS_APPEND_CONSOLE) Then
        ClearConsoleWorksheet
    End If
        
    ' Initialize row counter to first unused row
    Dim row As Long
    With ConsoleSheet.UsedRange
        row = .Cells.item(.Cells.count).row
    End With
    
    ' Leave some white space between invocations
    If row = 1 Then
        row = row + 1
    Else
        row = row + 2
    End If
        
    If Trim$(dotCommand) <> vbNullString Then
        ' Log the command used to invoke Graphviz
        Dim commandExecuted As String
#If Mac Then
        ' dot command is actually specified in the ExcelToGraphviz.applescript file. Fake it for the console
        commandExecuted = "dot " & dotCommand
#Else
        commandExecuted = dotCommand
#End If
        ConsoleSheet.Cells.item(row, 1).value = ">"
        ConsoleSheet.Cells.item(row, 2).value = commandExecuted
        row = row + 2
    End If
    
    ' Split the text into an array of lines
    Dim parsedText As Variant
#If Mac Then
    parsedText = split(textBlob, vbCr) ' lines are delimited by Carriage Return
#Else
    parsedText = split(textBlob, vbLf) ' lines are delimited by Line Feed
#End If
    
    If UBound(parsedText) >= 0 Then
        ' Transfer the array of lines to the worksheet in one swift action
        Dim writeToRange As String
        writeToRange = "B" & row & ":B" & (row + (UBound(parsedText) - LBound(parsedText)))
        ConsoleSheet.Range(writeToRange).value = Application.Transpose(parsedText)
    End If
    
End Sub

' ==========================================================================
' PROCEDURE: LogToConsoleWorksheet
'
' PURPOSE:
'   A simplified logging utility that appends raw text messages to the
'   'Console' worksheet without requiring an associated CLI command string.
'
' TECHNICAL WORKFLOW:
'   1. DISPOSITION CHECK: Evaluates 'SETTINGS_APPEND_CONSOLE' to decide
'      whether to purge the sheet via 'ClearConsoleWorksheet' or append data.
'   2. ROW RESOLUTION: Identifies the next available row using 'UsedRange'.
'   3. DELIMITER HANDLING: Splits the 'textBlob' based on the host OS
'      standards (vbCr for macOS, vbLf for Windows).
'   4. BULK INJECTION: Writes the array to Column B using 'Application.Transpose'
'      for efficient range-based updates.
'
' TECHNICAL NOTES:
'   - Layer: UI / Diagnostics (Console).
'   - Usage: Ideal for internal VBA status messages or non-CLI error reporting.
' ==========================================================================
Public Sub LogToConsoleWorksheet(ByVal textBlob As String)

    ' Clear console if not in append mode
    If Not GetCellBoolean(SettingsSheet.name, SETTINGS_APPEND_CONSOLE) Then
        ClearConsoleWorksheet
    End If
        
    ' Initialize row counter to first unused row
    Dim row As Long
    With ConsoleSheet.UsedRange
        row = .Cells.item(.Cells.count).row + 1
    End With
    
    ' Leave some white space between invocations
    'If row = 1 Then
    '    row = row + 1
    'Else
    '    row = row + 2
    'End If
    
    ' Split the text into an array of lines
    Dim parsedText As Variant
#If Mac Then
    parsedText = split(textBlob, vbCr) ' lines are delimited by Carriage Return
#Else
    parsedText = split(textBlob, vbLf) ' lines are delimited by Line Feed
#End If
    
    If UBound(parsedText) >= 0 Then
        ' Transfer the array of lines to the worksheet in one swift action
        Dim writeToRange As String
        writeToRange = "B" & row & ":B" & (row + (UBound(parsedText) - LBound(parsedText)))
        ConsoleSheet.Range(writeToRange).value = Application.Transpose(parsedText)
    End If
    
End Sub

' ==========================================================================
' PROCEDURE: CopyConsoleToClipboard
'
' PURPOSE:
'   Aggregates all log entries from the 'Console' worksheet into a single
'   string and copies it to the system clipboard for external support or debugging.
'
' TECHNICAL WORKFLOW:
'   1. STRING AGGREGATION: Iterates through the 'ConsoleSheet' rows,
'      concatenating the contents of Column B (Message column) with
'      Line Feed (vbLf) separators.
'   2. CLIPBOARD TRANSFER (Windows): Invokes 'ClipBoard_SetData' to commit
'      the log data to the Windows clipboard buffer.
'   3. UX FEEDBACK: Updates the Excel StatusBar with a localized success
'      or failure message for 5 seconds via 'UpdateStatusBarForNSeconds'.
'
' TECHNICAL NOTES:
'   - Platform: Windows Only (#If Not Mac).
'   - Layer: UI / Logic.
' ==========================================================================
Public Sub CopyConsoleToClipboard()
#If Not Mac Then

    ' Pull all the rows into a single string
    Dim consoleMessage As String
    consoleMessage = vbNullString
    
    Dim lastRow As Long
    With ConsoleSheet.UsedRange
        lastRow = .Cells.item(.Cells.count).row
    End With

    Dim i As Long
    For i = 1 To lastRow
        consoleMessage = consoleMessage & ConsoleSheet.Cells.item(i, 2).value & vbLf
    Next i

    If ClipBoard_SetData(consoleMessage) Then
        UpdateStatusBarForNSeconds GetMessage("statusbarClipboardCopyConsoleSuccess"), 5
    Else
        UpdateStatusBarForNSeconds GetMessage("statusbarClipboardCopyConsoleFailed"), 5
    End If
    
#End If
End Sub

' ==========================================================================
' PROCEDURE: ConsoleWorksheetToFile
'
' PURPOSE:
'   Exports the entire 'Console' log to a physical file, ensuring consistent
'   UTF-8 encoding and cross-platform line-ending normalization.
'
' TECHNICAL WORKFLOW:
'   1. BOUNDARY DETECTION: Identifies the 'lastRow' of the log using
'      the worksheet's 'UsedRange'.
'   2. MAC EXECUTION (#If Mac):
'      - Concatenates Column B (Messages) into a single string using
'        Carriage Returns (vbCr).
'      - Persists to disk via the native 'WriteTextToFile' wrapper.
'   3. WINDOWS EXECUTION (#Else):
'      - ADODB.STREAM ENCODING: Uses ADODB.Stream to force UTF-8 encoding.
'      - BOM STRIPPING: Manually shifts the stream position to skip the
'        3-byte Byte Order Mark (BOM) for cleaner text file interoperability.
'      - PERSISTENCE: Saves via a binary stream to ensure the BOM-less
'        state is preserved on disk.
'   4. RESOURCE HYGIENE: Implements 'EndMacro' cleanup to close handles and
'      nullify late-bound ADO objects.
'
' TECHNICAL NOTES:
'   - Platform: Cross-Platform (ADODB/Windows vs. Native/Mac).
'   - Layer: UI / File System.
' ==========================================================================
Public Sub ConsoleWorksheetToFile(ByVal fileName As String)

    Dim rowNumber As Long
    Dim lastRow As Long
    With ConsoleSheet.UsedRange
        lastRow = .Cells.item(.Cells.count).row
    End With
    
#If Mac Then
    Dim consoleText As String
    consoleText = vbNullString
    
    For rowNumber = 1 To lastRow
        consoleText = consoleText & ConsoleSheet.Cells(rowNumber, 2).value & vbCr
    Next rowNumber
    
    WriteTextToFile consoleText, fileName
#Else
    
    ' Output file objects
    Dim utf8Stream As Object
    Set utf8Stream = CreateObject("ADODB.Stream")
    If utf8Stream Is Nothing Then GoTo EndMacro
    
    Dim binaryStream As Object
    Set binaryStream = CreateObject("ADODB.Stream")
    If binaryStream Is Nothing Then GoTo EndMacro
    
    ' Initialize the utf8Stream object
    utf8Stream.Type = StreamTypeEnum.adTypeText
    utf8Stream.Charset = UTF8_CHARSET
    utf8Stream.Open
    
    For rowNumber = 1 To lastRow
        utf8Stream.WriteText ConsoleSheet.Cells.item(rowNumber, 2).value & vbLf
    Next rowNumber
    
    ' Initialize the object which is used to remove the Byte Order Mark (BOM) from the UTF-8 stream
    binaryStream.Type = StreamTypeEnum.adTypeBinary
    binaryStream.mode = ConnectModeEnum.adModeReadWrite
    binaryStream.Open

    ' Position the start of the utf8 stream past the Byte Order Mark (BOM) (i.e. BOM = first 3 bytes)
    ' and copy the contents to the binary stream
    utf8Stream.position = 3
    utf8Stream.CopyTo binaryStream
    
    ' Write out UTF-8 data without the BOM
    binaryStream.SaveToFile fileName, SaveOptionsEnum.adSaveCreateOverWrite

EndMacro:
    ' Clean up our objects so we don't get a memory leak
    If Not (utf8Stream Is Nothing) Then
        If (utf8Stream.State And ObjectStateEnum.adStateOpen) = ObjectStateEnum.adStateOpen Then utf8Stream.Close
        Set utf8Stream = Nothing
    End If
    
    If Not (binaryStream Is Nothing) Then
        If (binaryStream.State And ObjectStateEnum.adStateOpen) = ObjectStateEnum.adStateOpen Then binaryStream.Close
        Set binaryStream = Nothing
    End If
#End If
End Sub

' ==========================================================================
' PROCEDURE: SaveConsoleToFile
'
' PURPOSE:
'   Triggers a user-interactive dialog to save the current 'Console' log
'   to a physical text file.
'
' TECHNICAL WORKFLOW:
'   1. DIALOG INVOCATION:
'      - MAC (#If Mac): Calls 'RunAppleScriptTask' with the "getSaveAsFileName"
'        command to bypass sandboxed file system restrictions.
'      - WINDOWS (#Else): Uses the native 'GetSaveAsFilename' method with
'        a ".txt" filter.
'   2. PERSISTENCE: If a path is selected, it invokes 'ConsoleWorksheetToFile'
'      to handle the UTF-8 encoding and BOM stripping logic.
'   3. UX FEEDBACK: Alerts the user with a localized "Saved to File" message
'      displaying the full target path.
'
' TECHNICAL NOTES:
'   - Platform: Cross-Platform (AppleScript vs. Windows Native Dialog).
'   - Layer: UI / File System.
' ==========================================================================
Public Sub SaveConsoleToFile()
    Dim fileName As String
    
#If Mac Then
    fileName = RunAppleScriptTask("getSaveAsFileName", ".txt")
#Else
    fileName = GetSaveAsFilename("Text Files (*.txt), *txt")
#End If

    If fileName <> vbNullString Then
        ConsoleWorksheetToFile (fileName)
        EmitMessage GetMessage("msgboxSavedToFile") & vbNewLine & fileName
    End If
End Sub

' ==========================================================================
' FUNCTION: RunGraphvizInVerboseMode
'
' PURPOSE:
'   Determines if the Graphviz engine should be executed with the verbose
'   flag (-v) based on a combination of UI state and user settings.
'
' TECHNICAL WORKFLOW:
'   1. UI STATE CHECK: Verifies if the 'ConsoleSheet' is currently visible
'      to the user.
'   2. PREFERENCE CHECK: Validates that both 'SETTINGS_GRAPHVIZ_VERBOSE'
'      and 'SETTINGS_LOG_TO_CONSOLE' are enabled in the project settings.
'   3. LOGICAL AND: Returns TRUE only if all three conditions are met,
'      ensuring verbose output is suppressed if there is no visible
'      destination or logging is disabled.
'
' TECHNICAL NOTES:
'   - Layer: Logic / Diagnostics.
'   - DeepWiki Context: Part of the "DOT Source Viewer & Console" logic
'     used to toggle detailed engine feedback.
' ==========================================================================
Public Function RunGraphvizInVerboseMode() As Boolean
    RunGraphvizInVerboseMode = ConsoleSheet.visible And GetSettingBoolean(SETTINGS_GRAPHVIZ_VERBOSE) And GetSettingBoolean(SETTINGS_LOG_TO_CONSOLE)
End Function

