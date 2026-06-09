Attribute VB_Name = "modUtilityExcelSheets"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modUtilityExcelSheets
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Utility / Excel Interop
'
' ROLE:
'   Lightweight worksheet-existence checker used throughout the project to
'   defensively validate sheet references before performing read/write or
'   structural operations.
'
' RESPONSIBILITIES:
'   - WorksheetExists:
'       o Test whether a worksheet with a given name exists in the active
'         workbook using late-bound resolution
'       o Return Boolean without raising errors
'
' ARCHITECTURAL NOTES:
'   - Uses ActiveWorkbook.Sheets.[_Default] for name-based lookup.
'   - Error-suppressed resolution ensures safe use in initialization,
'     validation, and conditional-creation workflows.
'   - Consumed by Settings, Data, SQL, and utility modules that must avoid
'     invalid sheet references.
'
' USAGE:
'   - Ideal for guard clauses before creating, deleting, or modifying sheets.
'
' RELATED WIKI PAGES:
'   - Worksheet Access Patterns
'   - Defensive Workbook Operations
' =============================================================================

Option Explicit

Public Function WorksheetExists(ByVal worksheetName As String) As Boolean
    Dim sheetTest As Worksheet
    On Error Resume Next
    Set sheetTest = ActiveWorkbook.Sheets.[_Default](worksheetName)
    On Error GoTo 0
    WorksheetExists = Not sheetTest Is Nothing
End Function

