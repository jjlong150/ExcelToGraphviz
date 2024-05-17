Attribute VB_Name = "modUtilityExcelColumns"
' Copyright (c) 2015-2022 Jeffrey J. Long. All rights reserved

'@Folder("Utility.Excel")

Option Explicit

Public Function GetLastColumn(ByVal worksheetName As String, ByVal row As Long) As Long

    ' Determine which columns have data
    With ActiveWorkbook.Worksheets.[_Default](worksheetName)
        GetLastColumn = .Cells(row, .columns.Count).End(xlToLeft).column
    End With

End Function

Public Sub ShowColumn(ByVal worksheetName As String, ByVal columnNumber As Long, ByVal show As Boolean)
    Dim alphabeticColumnName As String
    alphabeticColumnName = ConvertColumnNumberToLetters(columnNumber)
    ActiveWorkbook.Worksheets.[_Default](worksheetName).columns(alphabeticColumnName & ":" & alphabeticColumnName).AutoFit
    ActiveWorkbook.Worksheets.[_Default](worksheetName).Range(alphabeticColumnName & ":" & alphabeticColumnName).EntireColumn.Hidden = Not show
End Sub

Public Function ConvertColumnNumberToLetters(ByVal columnNumber As Long) As String
    Dim alpha As Long
    Dim remainder As Long
    alpha = Int(columnNumber / 27)
    remainder = columnNumber - (alpha * 26)
    If alpha > 0 Then
        ConvertColumnNumberToLetters = Chr$(alpha + 64)
    End If
    If remainder > 0 Then
        ConvertColumnNumberToLetters = ConvertColumnNumberToLetters & Chr$(remainder + 64)
    End If
End Function


