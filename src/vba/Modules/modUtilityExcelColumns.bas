Attribute VB_Name = "modUtilityExcelColumns"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modUtilityExcelColumns
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Utility / Excel Interop
'
' ROLE:
'   Column-oriented worksheet utilities for determining last-used columns,
'   converting column numbers to A1-style letters, and showing/hiding columns
'   with automatic autofit behavior.
'
' RESPONSIBILITIES:
'   - Column discovery:
'       o GetLastColumn: determine the rightmost non-empty column in a row
'   - Column visibility:
'       o ShowColumn: toggle visibility of a specific column and apply AutoFit
'   - Column name conversion:
'       o ConvertColumnNumberToLetters: convert numeric column index -> A, B, ..., AA, AB
'
' ARCHITECTURAL NOTES:
'   - Uses ActiveWorkbook.Worksheets.[_Default] for late-bound sheet resolution.
'   - Column-letter conversion supports 1-702 (A-ZZ) using a compact algorithm.
'   - AutoFit is applied before visibility toggling to ensure consistent layout.
'   - Consumed by data-sheet workflows, diagnostics, and UI-driven column toggles.
'
' USAGE:
'   - Ideal for worksheet automation, dynamic UI toggles, and data-driven
'     column visibility logic.
'
' RELATED WIKI PAGES:
'   - Worksheet Access Patterns
'   - Column Visibility & Layout Management
' =============================================================================

Option Explicit

Public Function GetLastColumn(ByVal worksheetName As String, ByVal row As Long) As Long

    ' Determine which columns have data
    With ActiveWorkbook.worksheets.[_Default](worksheetName)
        GetLastColumn = .Cells(row, .columns.count).End(xlToLeft).Column
    End With

End Function

Public Sub ShowColumn(ByVal worksheetName As String, ByVal ColumnNumber As Long, ByVal show As Boolean)
    Dim alphabeticColumnName As String
    alphabeticColumnName = ConvertColumnNumberToLetters(ColumnNumber)
    ActiveWorkbook.worksheets.[_Default](worksheetName).columns(alphabeticColumnName & ":" & alphabeticColumnName).AutoFit
    ActiveWorkbook.worksheets.[_Default](worksheetName).Range(alphabeticColumnName & ":" & alphabeticColumnName).EntireColumn.Hidden = Not show
End Sub

Public Function ConvertColumnNumberToLetters(ByVal ColumnNumber As Long) As String
    Dim alpha As Long
    Dim remainder As Long
    alpha = Int(ColumnNumber / 27)
    remainder = ColumnNumber - (alpha * 26)
    If alpha > 0 Then
        ConvertColumnNumberToLetters = Chr$(alpha + 64)
    End If
    If remainder > 0 Then
        ConvertColumnNumberToLetters = ConvertColumnNumberToLetters & Chr$(remainder + 64)
    End If
End Function


