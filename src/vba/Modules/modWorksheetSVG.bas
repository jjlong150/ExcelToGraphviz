Attribute VB_Name = "modWorksheetSVG"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modWorksheetSVG
' COPYRIGHT: Copyright (c) 2015-2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Relationship Visualizer / Sheets / SVG
'
' ROLE:
'   The SVG post-processing and XML-transformation engine. Applies user-defined
'   find/replace rules, injects CSS/JavaScript, and produces enhanced SVG output
'   after Graphviz rendering.
'
' RESPONSIBILITIES:
'   - Stream-based XML editing:
'       o Load full SVG into memory, apply ordered rule substitutions,
'         and write the transformed output to disk.
'
'   - Worksheet-driven rule execution:
'       o Iterate SVG worksheet rows, respecting comment flags (#)
'       o Apply case-insensitive replacements for robust XML matching
'
'   - Interactive editing:
'       o Launch CellValueEditForm for multi-line CSS/JS editing
'       o Support large replacement strings beyond Excel's row display limits
'
'   - UI responsiveness:
'       o Use DoEvents during long replacement loops to keep Excel responsive
'
' ARCHITECTURAL NOTES:
'   - Designed for post-Graphviz enhancement (animations, tooltips, interactivity)
'   - Uses a simple, deterministic rule engine for predictable transformations
'   - Integrates with the SVG worksheet schema via enumerated row/column indices
'
' VERSION NOTES:
'   - v6.0.00 (May 14, 2023):
'       o Introduced SVG post-processing subsystem and new SVG worksheet
'       o Added JavaScript-based node/edge highlighting
'       o Added on/off toggle for post-processing
'
'   - v8.0.0 (Aug 27, 2025):
'       o Added pop-up editor for large replacement strings
'       o Improved JavaScript animation logic and added macOS-style variant
'       o Added Copy to Clipboard, Graph to File, and All Views to File buttons
'
'   - v9.0.0 (Dec 22, 2025):
'       o Updated SVG animation logic to support rounded-corner edge rendering
'
' USAGE:
'   - Automatically invoked during Publish/Preview when SVG post-processing is enabled
'   - Used to inject animations, tooltips, CSS, JS, and structural SVG modifications
'
' RELATED WIKI PAGES:
'   - SVG Post-Processing
'   - JavaScript Animation Injection
'   - Find/Replace Rule Design
' =============================================================================

Option Explicit

' ==========================================================================
' ENUM: svgLayoutRow / svgLayoutColumn
' PURPOSE:
'   Defines the physical schema of the SVG transformation worksheet.
'
' TECHNICAL ROLE:
'   Acts as the "Data Contract" for the post-processor. By using enums instead
'   of hard-coded numbers, the engine remains resilient to future column
'   reordering in the 'SVG' worksheet.
' ==========================================================================
Public Enum svgLayoutRow
    headingRow = 1
    firstDataRow = 2
End Enum

Public Enum svgLayoutColumn
    flagColumn = 1
    findColumn = 2
    replaceColumn = 3
End Enum

' ==========================================================================
' PROCEDURE: FindAndReplaceSVG
' @Service: Performs surgical XML injection on generated SVG files.
'
' TECHNICAL WORKFLOW:
'   1. INGESTION: Loads the raw Graphviz SVG output into a memory string.
'   2. RULE ITERATION: Scans the 'SVG' worksheet for user-defined patterns.
'   3. LOGICAL FILTERING: Skips rows marked with '#' to allow for non-
'      destructive rule testing.
'   4. XML TRANSFORMATION: Executes a global 'replace' for each rule. This
'      is where CSS animations or tooltips are injected into the XML.
'   5. ASYNCHRONOUS UI: Employs 'DoEvents' to prevent Excel from freezing
'      during high-volume replacements.
'   6. EMISSION: Saves the transformed XML string to the final output file.
' ==========================================================================
Public Sub FindAndReplaceSVG(ByVal svgFileIn As String, ByVal svgFileOut As String)
    Dim svgText As String
    svgText = ReadFileToString(svgFileIn)
    
    ' Determine the last row with data
    Dim lastRow As Long
    With SvgSheet.UsedRange
        lastRow = .Cells.item(.Cells.count).row
    End With
    
    ' Loop through the data rows of SVG find/replace statements
    Dim row As Long
    For row = svgLayoutRow.firstDataRow To lastRow
        If SvgSheet.Cells.item(row, svgLayoutColumn.flagColumn).value <> FLAG_COMMENT Then
            svgText = replace(svgText, _
                SvgSheet.Cells.item(row, svgLayoutColumn.findColumn).value, _
                SvgSheet.Cells.item(row, svgLayoutColumn.replaceColumn).value, _
                1, -1, vbTextCompare)
        End If
        DoEvents
    Next row
    
    ' Write the modified string to a file
    WriteTextToFile svgText, svgFileOut
End Sub

' ==========================================================================
' PROCEDURE: ShowSVGEditForm
' @Interface: Launches the large-format code editor.
'
' TECHNICAL ROLE:
'   Bridge to 'CellValueEditForm'. Required because SVG replacements often
'   involve multi-line CSS or JavaScript which is difficult to manage
'   directly within an Excel cell. Used by the "Edit" button that appears
'   on the "Replace" cell.
' ==========================================================================
Public Sub ShowSVGEditForm()
    CellValueEditForm.show
End Sub



