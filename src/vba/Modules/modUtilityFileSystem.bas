Attribute VB_Name = "modUtilityFileSystem"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modUtilityFileSystem
' COPYRIGHT: Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Utility / File System
'
' ROLE:
'   Cross-platform file-system abstraction layer providing directory and file
'   existence checks, safe deletion and creation routines, and text-file
'   read/write helpers compatible with both Windows and macOS sandbox rules.
'
' RESPONSIBILITIES:
'   - Existence checks:
'       • DirectoryExists: macOS via AppleScriptTask; Windows via FileSystemObject
'       • FileExists: macOS via AppleScriptTask; Windows via FileSystemObject
'   - File and directory operations:
'       • DeleteFile: wrapper over Kill with defensive error handling
'       • CreateDirectory: wrapper over MkDir with suppressed errors
'   - Text I/O:
'       • WriteTextToFile: overwrite-mode text output using FreeFile
'       • ReadFileToString: macOS line-by-line Input; Windows via OpenTextFile
'   - Enumerations:
'       • IOMode and FileFormat enums mirror VBScript/OpenTextFile semantics
'
' ARCHITECTURAL NOTES:
'   - macOS uses AppleScriptTask to bypass sandbox restrictions for file
'     existence checks.
'   - Windows uses late-bound Scripting.FileSystemObject for compatibility
'     across Office versions.
'   - All routines are defensive and avoid raising errors to callers.
'   - Consumed by SQL engine, SVG export, logging, and Graphviz file output.
'
' USAGE:
'   - Ideal for cross-platform file checks, safe deletion, and lightweight
'     text-file read/write operations.
'
' RELATED WIKI PAGES:
'   - File System Abstraction (Windows/macOS)
'   - Text Output & Graphviz File Generation
'   - Sandbox-Aware macOS Operations
' =============================================================================

Option Explicit

' https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/opentextfile-method
Public Enum IOMode
    ForReading = 1      ' Opens a file for reading only
    ForWriting = 2      ' Opens a file for writing. If the file already exists, the contents are overwritten.
    ForAppending = 8    ' Opens a file and starts writing at the end (appends). Contents are not overwritten.
End Enum

Public Enum FileFormat
    TristateUseDefault = -2 ' Use default system setting
    TristateTrue = -1       ' Opens the file as Unicode
    TristateFalse = 0       ' Opens the file as ASCII
End Enum

Public Function DirectoryExists(ByVal dirPath As String) As Boolean

#If Mac Then
    DirectoryExists = False
   
    Dim applescriptResult As String
    applescriptResult = RunAppleScriptTask("doesFolderExist", dirPath)

    If applescriptResult = "true" Then
        DirectoryExists = True
    End If
    
#Else
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    DirectoryExists = False
    
    If Len(dirPath) > 0 Then
        If fso.FolderExists(dirPath) = True Then
            DirectoryExists = True
        End If
    End If

    Set fso = Nothing
#End If
End Function

Public Function FileExists(ByVal filePath As String) As Boolean
    FileExists = False

#If Mac Then
    '  Use Apple Script to get around sandbox restrictions
    Dim applescriptResult As String
    applescriptResult = RunAppleScriptTask("doesFileExist", filePath)

    If applescriptResult = "true" Then
        FileExists = True
    Else
        FileExists = False
    End If
#Else
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Len(filePath) > 0 Then
        If fso.FileExists(filePath) = True Then
            FileExists = True
        End If
    End If

    Set fso = Nothing
#End If
End Function

Public Sub DeleteFile(ByVal fileToDelete As String)
    On Error Resume Next
    Kill fileToDelete
    On Error GoTo 0
End Sub

Public Sub CreateDirectory(ByVal directoryName As String)
    On Error Resume Next
    MkDir directoryName
    On Error GoTo 0
End Sub

Public Sub WriteTextToFile(ByVal textToWrite As String, ByVal fileNameToWriteTo As String)
    
    ' Output file handle
    Dim fileNum As Long
    
    On Error GoTo EndMacro:
    fileNum = FreeFile

    ' Open file for output. Any existing file by the same name will be overwritten
    Open fileNameToWriteTo For Output Access Write As #fileNum

    ' Write the Graphviz commands to a file
    Print #fileNum, textToWrite

EndMacro:
    On Error GoTo 0
    Close #fileNum

End Sub

Public Function ReadFileToString(ByVal fileName As String) As String
    
#If Mac Then
    Dim fileNum As Integer
    Dim dataLine As String

    fileNum = FreeFile()

    Open fileName For Input As #fileNum

    While Not EOF(fileNum)
        Line Input #fileNum, dataLine ' read in data 1 line at a time
        ReadFileToString = ReadFileToString & dataLine & vbNewLine
    Wend
    
    Close #fileNum
#Else
    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")

    Dim textFile As Object
    
    Set textFile = fileSystem.OpenTextFile(fileName, IOMode:=IOMode.ForReading, format:=FileFormat.TristateFalse)
    
    ReadFileToString = textFile.ReadAll

    textFile.Close
    Set textFile = Nothing
    Set fileSystem = Nothing
#End If
    
End Function
