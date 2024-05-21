Attribute VB_Name = "modWorksheetData"
' Copyright (c) 2015-2022 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.Data")

Option Explicit

Public Sub ClearDataWorksheet(ByVal worksheetName As String)
    Dim lastColumn As Long
    Dim cellRange As String
    Dim lastRow As Long
    Dim dataLayout As dataWorksheet
    
    ' Get the layout of the 'data' worksheet
    dataLayout = GetSettingsForDataWorksheet(worksheetName)

    ' Determine the range of the cells which need to be cleared
    With ActiveWorkbook.Worksheets.[_Default](worksheetName).UsedRange
        lastRow = .Cells(.Cells.Count).row
    End With
    
    ' If the worksheet is already empty we do not want to wipe out the heading row
    If lastRow < dataLayout.firstRow Then
        lastRow = dataLayout.firstRow
    End If
    
    ' Determine the columns to clear
    lastColumn = GetLastColumn(worksheetName, dataLayout.headingRow)

    ' Remove any existing content
    cellRange = "A" & dataLayout.firstRow & ":" & ConvertColumnNumberToLetters(lastColumn) & lastRow
    ActiveWorkbook.Worksheets.[_Default](worksheetName).Range(cellRange).ClearContents
End Sub

