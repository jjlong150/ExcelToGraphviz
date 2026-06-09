Attribute VB_Name = "modUtilityExcelCell"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modUtilityExcel
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Utility / Excel Interop
'
' ROLE:
'   Thin abstraction layer over worksheet cell access. Provides consistent,
'   centralized helpers for reading and writing typed values, toggling
'   settings, clearing ranges, and loading external file contents into cells.
'
' RESPONSIBILITIES:
'   - Typed cell accessors:
'       o GetCellLong, GetCellString, GetCellBoolean, GetCellUCase
'       o GetCell(row, col) with trimming and normalization
'   - Cell mutation helpers:
'       o SetCell, SetCellString, ClearCell, ClearNamedCellContents
'       o ToggleCell and Toggle for boolean-driven value switching
'   - File ingestion:
'       o ReadFileIntoCell: binary-safe file read into a worksheet cell
'   - Directory selection:
'       o SelectDirectoryToCell: integrates ChooseDirectory with worksheet storage
'
' ARCHITECTURAL NOTES:
'   - Uses ActiveWorkbook.Sheets.[_Default] for late-bound sheet resolution.
'   - Ensures consistent trimming, case normalization, and boolean coercion.
'   - File ingestion uses FreeFile + Binary mode for full-file reads.
'   - Consumed by Settings, SQL, SVG, Source, Styles, and Diagnostics workflows.
'
' USAGE:
'   - Provides a stable, centralized API for all worksheet cell interactions.
'   - Used throughout the project to avoid duplicated Range/Cells logic.
'
' RELATED WIKI PAGES:
'   - Worksheet Access Patterns
'   - Settings Sheet Architecture
'   - File Ingestion & Binary Read Guidelines
' =============================================================================

Option Explicit

' -----------------------------------------------------------------------------
' WORKSHEET RESOLUTION CACHE
'
' Every cell accessor below used to re-walk the full COM chain
'   ActiveWorkbook.Sheets.[_Default](worksheetName)
' on every single call. During graph generation the parsing pipeline reads tens
' of thousands of cells (multiple passes over the data sheet, each row touching
' ~11 columns), and AutoDraw re-runs the whole pipeline on every keystroke in
' Automatic mode, so that collection lookup was being paid repeatedly for the
' same sheet.
'
' ResolveSheet caches the last resolved worksheet and reuses it while the request
' is for the same sheet name in the same active workbook. The cache is validated
' on every call by an object-identity check that the cached parent is still
' ActiveWorkbook plus an exact name match. If either check fails the sheet is
' re-resolved from scratch, so the returned object is always identical to the
' original direct lookup -- the optimisation only removes repeated collection
' traversal and never changes which cell is read or written.
' -----------------------------------------------------------------------------
Private mCachedWorkbook As Object
Private mCachedSheetName As String
Private mCachedSheet As Object

Private Function ResolveSheet(ByVal worksheetName As String) As Object
    If Not mCachedSheet Is Nothing Then
        If mCachedSheetName = worksheetName Then
            If mCachedWorkbook Is ActiveWorkbook Then
                Set ResolveSheet = mCachedSheet
                Exit Function
            End If
        End If
    End If

    Set mCachedWorkbook = ActiveWorkbook
    Set mCachedSheet = ActiveWorkbook.Sheets.[_Default](worksheetName)
    mCachedSheetName = worksheetName
    Set ResolveSheet = mCachedSheet
End Function

Public Function GetCellLong(ByVal worksheetName As String, ByVal cellName As String) As Long
    GetCellLong = CLng(ResolveSheet(worksheetName).Range(cellName).value)
End Function

Public Function GetCellString(ByVal worksheetName As String, ByVal cellName As String) As String
    GetCellString = ResolveSheet(worksheetName).Range(cellName).value
End Function

Public Function GetCell(ByVal worksheetName As String, ByVal row As Long, ByVal col As Long) As String
    GetCell = Trim$(ResolveSheet(worksheetName).Cells(row, col).value)
End Function

Public Sub SetCell(ByVal worksheetName As String, ByVal row As Long, ByVal col As Long, ByVal cellValue As Variant)
    ResolveSheet(worksheetName).Cells(row, col).value = cellValue
End Sub

Public Sub ClearCell(ByVal worksheetName As String, ByVal row As Long, ByVal col As Long)
    ResolveSheet(worksheetName).Cells(row, col).ClearContents
End Sub

Public Function GetCellUCase(ByVal worksheetName As String, ByVal row As Long, ByVal col As Long) As String
    GetCellUCase = UCase$(Trim$(ResolveSheet(worksheetName).Cells(row, col).value))
End Function

Public Sub SetCellString(ByVal worksheetName As String, ByVal cellName As String, ByVal cellValue As String)
    ResolveSheet(worksheetName).Range(cellName).value = cellValue
End Sub

Public Sub ClearNamedCellContents(ByVal worksheetName As String, ByVal cellName As String)
    ResolveSheet(worksheetName).Range(cellName).ClearContents
End Sub

Public Function GetCellBoolean(ByVal worksheetName As String, ByVal cellName As String) As Boolean
    
    GetCellBoolean = False
    
    Select Case UCase$(GetCellString(worksheetName, cellName))
        Case "ON", "YES", "TRUE", "AUTO", "SHOW", "INCLUDE", "DEFAULT"
            GetCellBoolean = True
        Case Else
            GetCellBoolean = False
    End Select
    
End Function

Public Sub SelectDirectoryToCell(ByVal worksheetName As String, ByVal cellName As String)
    SetCellString worksheetName, cellName, ChooseDirectory(GetCellString(worksheetName, cellName))
End Sub

Public Sub ReadFileIntoCell(ByVal worksheetName As String, ByVal cellName As String, ByVal fileName As String)

    ' Clear out any previous data in the cell
    ActiveSheet.Range(cellName).ClearContents

    ' Make sure the file exists before attempting to read it
    If FileExists(fileName) Then
        
        ' Obtain a file handle
        Dim fileHandle As Long
        fileHandle = FreeFile()
        
        ' Open the file as binary
        Open fileName For Binary Access Read As #fileHandle

        Dim stringToHoldFile As String
        
        ' Create a string with enough space to hold the file contents
        '@Ignore AssignmentNotUsed
        stringToHoldFile = Space(FileLen(fileName))
        
        ' Read the entire file into the string
        Get #fileHandle, , stringToHoldFile

        ' Close the file
        Close #fileHandle
 
        ' Add to cell
        SetCellString worksheetName, cellName, stringToHoldFile
    End If
    
End Sub

Public Sub ToggleCell(ByVal Worksheet As String, ByVal cellName As String, ByVal bool As Boolean, ByVal trueValue As String, ByVal falseValue As String)
    SetCellString Worksheet, cellName, Toggle(bool, trueValue, falseValue)
End Sub

Public Function Toggle(ByVal bool As Boolean, ByVal trueValue As String, ByVal falseValue As String) As String
    
    If bool Then
        Toggle = trueValue
    Else
        Toggle = falseValue
    End If

End Function


