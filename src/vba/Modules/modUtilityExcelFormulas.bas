Attribute VB_Name = "modUtilityExcelFormulas"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modUtilityExcelFormulas
' COPYRIGHT: Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Relationship Visualizer / Data Sheet Helpers
'
' ROLE:
'   Generate Graphviz-ready subgraph fragments from worksheet ranges. Supports
'   rank-constrained subgraphs, unconstrained subgraphs, and HTML-table labels
'   for node styling and layout control.
'
' RESPONSIBILITIES:
'   - Rank-constrained subgraphs:
'       • RangeToSubgraphWithRank: build { rank="…" ; node1 ; node2 ; … }
'       • Validate rank keywords (min, max, same, sink, source)
'       • Convenience wrappers: SameRank, MaxRank, MinRank, SinkRank, SourceRank
'   - Unconstrained subgraphs:
'       • RangeToSubgraph: build { node1 ; node2 ; … }
'       • Wrapper: subgraph()
'   - HTML-table labels:
'       • RangeToHtmlTable: convert a cell block into <table><tr><td>…</td></tr></table>
'       • TableLabel: wrap HTML table in <…> for Graphviz HTML-label syntax
'
' ARCHITECTURAL NOTES:
'   - All node identifiers are quoted via AddQuotes / AddQuotesConditionally.
'   - Output strings are fully Graphviz-ready and require no post-processing.
'   - HTML-table generation preserves row/column order exactly as in Excel.
'   - Consumed by SQL-driven pipelines, Data sheet helpers, and style workflows.
'
' USAGE:
'   - Ideal for Data-sheet-driven layout control, rank grouping, and HTML labels.
'
' RELATED WIKI PAGES:
'   - Subgraph & Rank Construction
'   - HTML Labels in Graphviz
'   - Data Sheet -> Graphviz Pipeline
' =============================================================================

Option Explicit

' @method RangeToSubgraphWithRank
' @param {Range} itemIds A set of cells which should all have the same rank
' @param {String} rankType How to rank the nodes. Valid values: min | max | same | sink | source
' @return {String} Subgraph for the nodes in the cell range

Public Function RangeToSubgraphWithRank(ByVal itemIds As Range, ByVal rankType As String) As String

    ' Ensure valid rankType
    Dim rankTypeOut As String
    '@Ignore AssignmentNotUsed
    rankTypeOut = "same"
    
    Select Case UCase$(rankType)
        Case "MAX":     rankTypeOut = "max"
        Case "MIN":     rankTypeOut = "min"
        Case "SAME":    rankTypeOut = "same"
        Case "SINK":    rankTypeOut = "sink"
        Case "SOURCE":  rankTypeOut = "source"
    End Select

    Dim peers As String
    peers = vbNullString
    
    ' Iterate the range of cells
    Dim item As Range
    For Each item In itemIds.Cells
        peers = peers & "; " & AddQuotesConditionally(Trim$(item.value))
    Next item
    
    ' Build the rank statement
    RangeToSubgraphWithRank = "{rank=" & AddQuotes(rankTypeOut) & peers & ";}"
    
End Function

'Convenience wrappers
Public Function SameRank(ByVal itemIds As Range) As String
    SameRank = RangeToSubgraphWithRank(itemIds, "same")
End Function

Public Function MaxRank(ByVal itemIds As Range) As String
    MaxRank = RangeToSubgraphWithRank(itemIds, "max")
End Function

Public Function MinRank(ByVal itemIds As Range) As String
    MinRank = RangeToSubgraphWithRank(itemIds, "min")
End Function

Public Function SinkRank(ByVal itemIds As Range) As String
    SinkRank = RangeToSubgraphWithRank(itemIds, "sink")
End Function

Public Function SourceRank(ByVal itemIds As Range) As String
    SourceRank = RangeToSubgraphWithRank(itemIds, "source")
End Function

' @method RangeToSubgraph
' @param {Range} itemIds A set of cells which should be in the subgraph
' @return {String} Subgraph for the nodes in the cell range

Public Function RangeToSubgraph(ByVal itemIds As Range) As String

    Dim peers As String
    peers = vbNullString
    
    ' Iterate the range of cells
    Dim item As Range
    For Each item In itemIds.Cells
        peers = peers & "; " & AddQuotes(Trim$(item.value))
    Next item
    
    ' Build the rank statement
    RangeToSubgraph = "{ " & peers & ";}"
    
End Function

'Convenience wrappers
Public Function subgraph(ByVal itemIds As Range) As String
    subgraph = RangeToSubgraph(itemIds)
End Function


' @method RangeToHtmlTable
' @param {Range} rng A set of cells which should be in the table
' @return {String} HTML-like string for the nodes in the cell range

Public Function RangeToHtmlTable(ByVal tableCells As Range) As String

    Dim htmlLabel As String
    htmlLabel = "<table>" & vbNewLine
    
    Dim rowIndex As Long
    Dim columnIndex As Long
  
    For rowIndex = 1 To tableCells.rows.count
        htmlLabel = htmlLabel & "<tr>"
        For columnIndex = 1 To tableCells.columns.count
            htmlLabel = htmlLabel & "<td>" & tableCells.Cells.item(rowIndex, columnIndex).value & "</td>"
        Next columnIndex
        htmlLabel = htmlLabel & "</tr>" & vbNewLine
    Next rowIndex

    htmlLabel = htmlLabel & "</table>"
    RangeToHtmlTable = htmlLabel
    
End Function

'Convenience wrappers
Public Function TableLabel(ByVal tableCells As Range) As String
    TableLabel = LESS_THAN & RangeToHtmlTable(tableCells) & GREATER_THAN
End Function




