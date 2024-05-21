Attribute VB_Name = "modWorksheetSVG"
' Copyright (c) 2015-2023 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.SVG")

Option Explicit

Public Enum svgLayoutRow
    headingRow = 1
    firstDataRow = 2
End Enum

Public Enum svgLayoutColumn
    flagColumn = 1
    findColumn = 2
    replaceColumn = 3
End Enum

Public Sub FindAndReplaceSVG(ByVal svgFileIn As String, ByVal svgFileOut As String)
    Dim svgText As String
    svgText = ReadFileToString(svgFileIn)
    
    ' Determine the last row with data
    Dim lastRow As Long
    With SvgSheet.UsedRange
        lastRow = .Cells.Item(.Cells.Count).row
    End With
    
    ' Loop through the data rows of SVG find/replace statements
    Dim row As Long
    For row = svgLayoutRow.firstDataRow To lastRow
        If SvgSheet.Cells.Item(row, svgLayoutColumn.flagColumn).Value <> FLAG_COMMENT Then
            svgText = replace(svgText, _
                SvgSheet.Cells.Item(row, svgLayoutColumn.findColumn).Value, _
                SvgSheet.Cells.Item(row, svgLayoutColumn.replaceColumn).Value, _
                1, -1, vbTextCompare)
        End If
        DoEvents
    Next row
    
    ' Write the modified string to a file
    WriteTextToFile svgText, svgFileOut
End Sub


