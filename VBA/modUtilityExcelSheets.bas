Attribute VB_Name = "modUtilityExcelSheets"
' Copyright (c) 2015-2022 Jeffrey J. Long. All rights reserved

'@Folder("Utility.Excel")

Option Explicit

Public Function WorksheetExists(ByVal worksheetName As String) As Boolean
    Dim sheetTest As Worksheet
    On Error Resume Next
    Set sheetTest = ActiveWorkbook.Sheets.[_Default](worksheetName)
    On Error GoTo 0
    WorksheetExists = Not sheetTest Is Nothing
End Function

