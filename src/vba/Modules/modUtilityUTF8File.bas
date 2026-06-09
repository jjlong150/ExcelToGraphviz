Attribute VB_Name = "modUtilityUTF8File"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modUtilityUTF8File
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Utility / File I/O
'
' ROLE:
'   High-fidelity UTF-8 read/write subsystem for Graphviz compatibility.
'   Provides BOM-aware streaming, cross-platform fallbacks, and Unicode-safe
'   persistence for DOT source, logs, and diagnostic output.
'
' RESPONSIBILITIES:
'   - UTF-8 writing (BOM-free):
'       o WriteTextToUTF8FileFileWithoutBOM: generate UTF-8 files without the
'         EF BB BF signature required by Graphviz's strict parsers
'       o Implements dual-stream ADODB pipeline (Text -> Binary) to strip BOM
'
'   - UTF-8 writing (with BOM):
'       o WriteTextToUTF8FileFileWithBOM: produce standard UTF-8 files with
'         ADO-generated BOM for external editors and log viewers
'
'   - UTF-8 reading:
'       o ReadUTF8File: load UTF-8 text using ADODB.Stream on Windows or
'         line-input fallback on macOS
'
' ARCHITECTURAL NOTES:
'   - Windows:
'       o Uses late-bound ADODB.Stream for high-performance Unicode I/O
'       o Charset = "UTF-8" ensures correct encoding and BOM handling
'       o Binary stream copy removes BOM cleanly for Graphviz ingestion
'
'   - macOS:
'       o Sandbox-safe fallback using native Line Input loops
'       o Ensures core read capability even without ADODB support
'
'   - Data integrity:
'       o Preserves full Unicode range for multi-language labels
'       o Ensures deterministic output for DOT rendering pipelines
'
' USAGE:
'   - BOM-free writer is used for DOT files consumed by Graphviz engines
'   - BOM-included writer is used for logs and human-readable exports
'   - Reader is used for previewing DOT output and ingesting diagnostic logs
'
' RELATED WIKI PAGES:
'   - UTF-8 Encoding & BOM Rules
'   - Graphviz File Requirements
'   - Cross-Platform File I/O Architecture
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: WriteTextToUTF8FileFileWithoutBOM
' PURPOSE:
'   Generates a clean UTF-8 text file compatible with the 'dot' engine.
'
' TECHNICAL WORKFLOW:
'   1. STREAM INITIALIZATION: Creates two late-bound 'ADODB.Stream' objects
'      (one for Text/UTF-8 and one for Binary).
'   2. ENCODED WRITE: Writes the VBA string into the UTF-8 text stream,
'      which automatically prepends the 3-byte BOM.
'   3. BOM STRIPPING: Sets the stream position to '3', effectively jumping
'      past the BOM, and copies the remaining data into the Binary stream.
'   4. FILE PERSISTENCE: Saves the Binary stream to disk, resulting in a
'      standardized, BOM-free UTF-8 file.
'   5. RESOURCE HYGIENE: Explicitly closes and destroys both streams to
'      prevent memory leaks during batch rendering.
'
' PLATFORM NOTE:
'   - Windows: Uses ADODB logic.
'   - macOS: Currently stubbed, as file-writing on Mac is handled via
'     dedicated AppleScript or Shell calls.
' ==========================================================================
Public Sub WriteTextToUTF8FileFileWithoutBOM(ByVal textToWrite As String, ByVal fileNameToWriteTo As String)
#If Mac Then
    EmitMessage "Sub 'WriteTextToUTF8FileFileWithoutBOM' is not supported on MacOS"
#Else
    ' Output file objects
    Dim utf8Stream As Object
    Set utf8Stream = CreateObject("ADODB.Stream")
    
    Dim binaryStream As Object
    Set binaryStream = CreateObject("ADODB.Stream")
    
    ' Initialize the utf8Stream object
    utf8Stream.Type = StreamTypeEnum.adTypeText
    utf8Stream.Charset = UTF8_CHARSET
    utf8Stream.Open
    
    ' Write the text to the stream
    utf8Stream.WriteText textToWrite
    
    ' Initialize the object which is used to remove the Byte Order Mark (BOM) from the UTF-8 stream
    binaryStream.Type = StreamTypeEnum.adTypeBinary
    binaryStream.mode = ConnectModeEnum.adModeReadWrite
    binaryStream.Open

    ' Position the start of the utf8 stream past the Byte Order Mark (BOM) (i.e. BOM = first 3 bytes)
    ' and copy the contents to the binary stream
    utf8Stream.position = 3
    utf8Stream.CopyTo binaryStream
    
    ' Write out UTF-8 data without the BOM
    binaryStream.SaveToFile fileNameToWriteTo, SaveOptionsEnum.adSaveCreateOverWrite

    ' Clean up our objects so we don't get a memory leak
    utf8Stream.Close
    Set utf8Stream = Nothing

    binaryStream.Close
    Set binaryStream = Nothing
#End If
End Sub

' ==========================================================================
' PROCEDURE: WriteTextToUTF8FileFileWithBOM
' PURPOSE:
'   Generates a standard UTF-8 text file including the leading BOM.
'
' TECHNICAL WORKFLOW:
'   1. STREAM INITIALIZATION: Instantiates a late-bound 'ADODB.Stream'
'      object for text processing.
'   2. ENCODING CONFIGURATION: Sets the character set to 'UTF-8', which
'      causes ADO to automatically include the 3-byte signature (EF BB BF).
'   3. DATA PERSISTENCE: Writes the VBA string to the stream and saves it
'      directly to the specified file path.
'   4. RESOURCE RECLAMATION: Closes the stream handle to prevent file
'      locking or memory leaks.
'
' USAGE:
'   - Used for exporting logs or DOT source files intended for manual
'     review in external Windows text editors.
' ==========================================================================
Public Sub WriteTextToUTF8FileFileWithBOM(ByVal textToWrite As String, ByVal fileNameToWriteTo As String)
#If Mac Then
    EmitMessage "Sub 'WriteTextToUTF8FileFileWithBOM' is not supported on MacOS"
#Else
    ' Output file objects
    Dim utf8Stream As Object
    Set utf8Stream = CreateObject("ADODB.Stream")
    
    ' Initialize the utf8Stream object
    utf8Stream.Type = StreamTypeEnum.adTypeText
    utf8Stream.Charset = UTF8_CHARSET
    utf8Stream.Open
    
    ' Write the text to the stream
    utf8Stream.WriteText textToWrite
    
    ' Write out UTF-8 data without the BOM
    utf8Stream.SaveToFile fileNameToWriteTo, SaveOptionsEnum.adSaveCreateOverWrite

    ' Clean up our objects so we don't get a memory leak
    utf8Stream.Close
    Set utf8Stream = Nothing
#End If
End Sub

' ==========================================================================
' FUNCTION: ReadUTF8File
' PURPOSE:
'   Reads the full content of a UTF-8 text file into a VBA string variable.
'
' TECHNICAL WORKFLOW:
'   1. WINDOWS EXECUTION (#If Not Mac):
'      - Uses an 'ADODB.Stream' with the 'UTF-8' charset.
'      - Handles the 'LoadFromFile' method, which automatically recognizes
'        and processes Byte Order Marks (BOM).
'      - High-performance: Reads the entire file content into memory at once.
'   2. MACOS EXECUTION (#If Mac):
'      - Employs the native 'Line Input' loop as a fallback for the
'        sandboxed Mac environment.
'      - Iteratively reconstructs the string by reading the file line-by-line.
'   3. RESOURCE HYGIENE: Ensures file handles and stream objects are closed
'      immediately after the read operation to prevent file-locking.
'
' USAGE:
'   - Used to ingest generated DOT source for previewing or to read
'     diagnostic log files into the 'Console' worksheet.
' ==========================================================================
Public Function ReadUTF8File(ByVal fileName As String) As String
#If Mac Then
    ReadUTF8File = ""
    
    Dim fileNum As Integer
    Dim dataLine As String

    fileNum = FreeFile()

    Open fileName For Input As #fileNum

    While Not EOF(fileNum)
        Line Input #fileNum, dataLine ' read in data 1 line at a time
        ReadUTF8File = ReadUTF8File & dataLine
    Wend
    
    Close #fileNum
#Else
    ' Read the file into a stream object
    Dim utf8Stream As Object
    Set utf8Stream = CreateObject("ADODB.Stream")
    
    utf8Stream.Charset = UTF8_CHARSET
    utf8Stream.Open
    utf8Stream.LoadFromFile fileName
    
    ' Pass back the file contents
    ReadUTF8File = utf8Stream.ReadText
          
    ' Clean up our objects so we don't get a memory leak
    utf8Stream.Close
    Set utf8Stream = Nothing
#End If
End Function

