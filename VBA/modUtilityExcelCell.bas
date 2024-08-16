Attribute VB_Name = "modUtilityExcelCell"
' Copyright (c) 2015-2022 Jeffrey J. Long. All rights reserved

'@Folder("Utility.Excel")
'@IgnoreModule ProcedureNotUsed

Option Explicit

Public Function GetCellLong(ByVal worksheetName As String, ByVal cellName As String) As Long
    GetCellLong = CLng(ActiveWorkbook.Worksheets.[_Default](worksheetName).Range(cellName).Value)
End Function

Public Function GetCellString(ByVal worksheetName As String, ByVal cellName As String) As String
    GetCellString = ActiveWorkbook.Worksheets.[_Default](worksheetName).Range(cellName).Value
End Function

Public Function GetCell(ByVal worksheetName As String, ByVal row As Long, ByVal col As Long) As String
    GetCell = Trim$(ActiveWorkbook.Sheets.[_Default](worksheetName).Cells(row, col).Value)
End Function

Public Sub SetCell(ByVal worksheetName As String, ByVal row As Long, ByVal col As Long, ByVal cellValue As Variant)
    ActiveWorkbook.Worksheets.[_Default](worksheetName).Cells(row, col).Value = cellValue
End Sub

Public Sub ClearCell(ByVal worksheetName As String, ByVal row As Long, ByVal col As Long)
    ActiveWorkbook.Worksheets.[_Default](worksheetName).Cells(row, col).ClearContents
End Sub

Public Function GetCellUCase(ByVal worksheetName As String, ByVal row As Long, ByVal col As Long) As String
    GetCellUCase = UCase$(Trim$(ActiveWorkbook.Sheets.[_Default](worksheetName).Cells(row, col).Value))
End Function

Public Sub SetCellString(ByVal worksheetName As String, ByVal cellName As String, ByVal cellValue As String)
    ActiveWorkbook.Worksheets.[_Default](worksheetName).Range(cellName).Value = cellValue
End Sub

Public Sub ClearNamedCellContents(ByVal worksheetName As String, ByVal cellName As String)
    ActiveWorkbook.Worksheets.[_Default](worksheetName).Range(cellName).ClearContents
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

Public Sub ReadFileIntoCell(ByVal worksheetName As String, ByVal cellName As String, ByVal filename As String)

    ' Clear out any previous data in the cell
    ActiveSheet.Range(cellName).ClearContents

    ' Make sure the file exists before attempting to read it
    If FileExists(filename) Then
        
        ' Obtain a file handle
        Dim fileHandle As Long
        fileHandle = FreeFile()
        
        ' Open the file as binary
        Open filename For Binary Access Read As #fileHandle

        Dim stringToHoldFile As String
        
        ' Create a string with enough space to hold the file contents
        '@Ignore AssignmentNotUsed
        stringToHoldFile = Space(FileLen(filename))
        
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

