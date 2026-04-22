Attribute VB_Name = "modWorksheetData"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modWorksheetData
' COPYRIGHT: Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Relationship Visualizer / Sheets / Data
'
' ROLE:
'   Manage lifecycle operations for Data-model worksheets, including safe
'   clearing of record rows, schema-aware range targeting, and multi-sheet
'   compatibility. Ensures all operations honor the Named Range API contract
'   and the dataWorksheet UDT.
'
' RESPONSIBILITIES:
'   - Data clearing:
'       • ClearDataWorksheet: purge record rows while preserving header rows,
'         formatting, and structural integrity
'
'   - Schema adaptation:
'       • Resolve firstRow and headingRow via GetSettingsForDataWorksheet
'       • Use GetLastColumn to dynamically determine horizontal boundaries
'         regardless of column reordering
'
'   - Multi-target support:
'       • Operates on any worksheet conforming to the Data-model schema
'         (Data, SQL, Imports, etc.)
'
' ARCHITECTURAL NOTES:
'   - Uses dataWorksheet UDT to remain independent of physical cell locations.
'   - Ensures header rows are never cleared, even when UsedRange reports
'     minimal content.
'   - Dynamic range construction ensures compatibility with user-driven
'     column movement and schema evolution.
'   - Consumed by Import, SQL Refresh, and initialization workflows.
'
' USAGE:
'   - Ideal for resetting Data-model sheets before imports, SQL refreshes,
'     or batch graph-generation cycles.
'
' RELATED WIKI PAGES:
'   - Data Worksheet Contract (Named Range API)
'   - Schema-Aware Clearing Logic
'   - Data Lifecycle & Refresh Pipeline
' =============================================================================

Option Explicit

' ==========================================================================
' PROCEDURE: ClearDataWorksheet
'
' PURPOSE:
'   Purges all record data from a specified Data worksheet while preserving
'   the integrity of the header row and worksheet structure.
'
' TECHNICAL WORKFLOW:
'   1. SCHEMA DISCOVERY: Invokes 'GetSettingsForDataWorksheet' to retrieve
'      the 'dataWorksheet' UDT, identifying the 'firstRow' and 'headingRow'
'      per the Named Range API "Contract."
'   2. BOUNDARY CALCULATION:
'      - Vertical: Identifies the 'lastRow' using 'UsedRange'. Includes
'        safety logic to ensure the header row is never targeted.
'      - Horizontal: Resolves the 'lastColumn' via 'GetLastColumn' to
'        capture all dynamic attribute fields.
'   3. BULK PURGE: Constructs a dynamic 'cellRange' and executes
'      '.ClearContents' for high-performance data removal.
'
' USAGE:
'   - Essential for resetting the workspace before a new "Import" or "SQL
'     Refresh" operation.
'   - Supports multi-sheet targets by passing the 'worksheetName' parameter.
' ==========================================================================
Public Sub ClearDataWorksheet(ByVal worksheetName As String)
    Dim lastColumn As Long
    Dim cellRange As String
    Dim lastRow As Long
    Dim dataLayout As dataWorksheet
    
    ' Get the layout of the 'data' worksheet
    dataLayout = GetSettingsForDataWorksheet(worksheetName)

    ' Determine the range of the cells which need to be cleared
    With ActiveWorkbook.worksheets.[_Default](worksheetName).UsedRange
        lastRow = .Cells(.Cells.count).row
    End With
    
    ' If the worksheet is already empty we do not want to wipe out the heading row
    If lastRow < dataLayout.firstRow Then
        lastRow = dataLayout.firstRow
    End If
    
    ' Determine the columns to clear
    lastColumn = GetLastColumn(worksheetName, dataLayout.headingRow)

    ' Remove any existing content
    cellRange = "A" & dataLayout.firstRow & ":" & ConvertColumnNumberToLetters(lastColumn) & lastRow
    ActiveWorkbook.worksheets.[_Default](worksheetName).Range(cellRange).ClearContents
End Sub

