Attribute VB_Name = "modLocaleAudit"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Locale")
'@IgnoreModule ProcedureNotUsed, ModuleWithoutFolder

Option Explicit

Public Sub SyncTranslations()
    SyncLocaleToMaster "locale_en-US", "locale_de-DE"
    SyncLocaleToMaster "locale_en-US", "locale_en-GB"
    SyncLocaleToMaster "locale_en-US", "locale_fr-FR"
    SyncLocaleToMaster "locale_en-US", "locale_it-IT"
    SyncLocaleToMaster "locale_en-US", "locale_pl-PL"
End Sub

Public Sub AuditTranslations()
    CompareTranslationKeys "locale_en-US", "locale_de-DE"
    CompareTranslationKeys "locale_en-US", "locale_en-GB"
    CompareTranslationKeys "locale_en-US", "locale_fr-FR"
    CompareTranslationKeys "locale_en-US", "locale_it-IT"
    CompareTranslationKeys "locale_en-US", "locale_pl-PL"
End Sub

Public Sub CompareTranslationKeys(ByVal masterSheet As String, ByVal sheetToTest As String)
    Dim lastRow As Long
    With LocaleEnUsSheet.UsedRange
        lastRow = .Cells.Item(.Cells.count).row
    End With

    Dim row As Long
    For row = 2 To lastRow
        If GetCell(masterSheet, row, 1) <> GetCell(sheetToTest, row, 1) Then
            Debug.Print sheetToTest & " is out of sync at row " & row
            Debug.Print GetCell(masterSheet, row, 1) & " != " & GetCell(sheetToTest, row, 1)
           Exit Sub
        End If
    Next
    Debug.Print sheetToTest & " keys are in sync with " & masterSheet
End Sub

Public Sub SyncLocaleToMaster(ByVal masterSheet As String, ByVal sheetToSync As String)

    ' Determine the lasr row of the master locale
    Dim lastRow As Long
    With LocaleEnUsSheet.UsedRange
        lastRow = .Cells.Item(.Cells.count).row
    End With

    ' Loop through all the rows in the master, saving the key and row
    ' number in a dictionary
    Dim masterDictionary As Dictionary
    Set masterDictionary = New Dictionary
    
    Dim row As Long
    For row = 2 To lastRow
        masterDictionary.Add GetCell(masterSheet, row, 1), row
    Next
    
    ' Find last row with data in the worksheet to be synced with the master
    With ActiveWorkbook.worksheets.[_Default](sheetToSync).UsedRange
        lastRow = .Cells(.Cells.count).row
    End With
    
    ' Loop through the sheet to sync with master, and fetch the row number
    ' in the master sheet that corresponds to the key. This information
    ' lets us sort the modified sheet so that the rows are in the same
    ' order as the master worksheet.
    Dim key As String
    Dim value As Long
    
    For row = 2 To lastRow
        key = GetCell(sheetToSync, row, 1)
        value = masterDictionary.Item(key)
        SetCell sheetToSync, row, 6, value
    Next
    
    ' Loop through the sheet to sync again, removing the keys from the master
    ' dictionary. Any keys left in the dictionary are missing from the sheet
    ' to sync.
    For row = 2 To lastRow
        key = GetCell(sheetToSync, row, 1)
        If masterDictionary.Exists(key) Then
            masterDictionary.Remove key
        End If
    Next
    
    ' Add column headings to the columns we will write to
    SetCell sheetToSync, 1, 6, "Sort Order"
    SetCell sheetToSync, 1, 7, "Messages"
    
    ' Iterate through the remaining keys in the master dictionary. Fetch the default
    ' text from the master, and write it to the corresponding column in the
    ' sheet to sync. Include the row number, and a value which can be filtered on
    ' which identifies the row as needing to be translated.
    Dim dictKey As Variant
    For Each dictKey In masterDictionary.Keys()
        SetCell sheetToSync, row, 1, dictKey    ' Control ID
        SetCell sheetToSync, row, 2, GetCell(masterSheet, masterDictionary.Item(dictKey), 2) ' Compact Control Labels
        SetCell sheetToSync, row, 3, GetCell(masterSheet, masterDictionary.Item(dictKey), 3) ' Verbose Control Labels
        SetCell sheetToSync, row, 4, GetCell(masterSheet, masterDictionary.Item(dictKey), 4) ' Screentip
        SetCell sheetToSync, row, 5, GetCell(masterSheet, masterDictionary.Item(dictKey), 5) ' Supertip
        SetCell sheetToSync, row, 6, masterDictionary.Item(dictKey)  ' Master row number
        SetCell sheetToSync, row, 7, "Requires translation"     ' Flag for translation
        row = row + 1
    Next
    
    Debug.Print sheetToSync & " has been synced to " & masterSheet

End Sub

