Attribute VB_Name = "modLocaleAudit"
'@IgnoreModule ProcedureNotUsed, ModuleWithoutFolder
' Copyright (c) 2015-2023 Jeffrey J. Long. All rights reserved

Option Explicit

Public Sub CompareTranslationKeys()

    Dim masterSheet As String
    masterSheet = "locale_pl-PL"

    Dim sheetToTest As String
    sheetToTest = "locale_it-IT"
    
    Dim lastRow As Long
    With LocaleEnUsSheet.UsedRange
        lastRow = .Cells.Item(.Cells.Count).row
    End With

    Dim row As Long
    For row = 2 To lastRow
        If GetCell(masterSheet, row, 1) <> GetCell(sheetToTest, row, 1) Then
            Debug.Print "Worksheets are out of sync at row " & row
            Debug.Print GetCell(masterSheet, row, 1) & " != " & GetCell(sheetToTest, row, 1)
           Exit Sub
        End If
    Next
    Debug.Print "keys are in sync"
End Sub

