Attribute VB_Name = "modLocaleAudit"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modLocaleAudit
' COPYRIGHT: Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Locale / Maintenance Utilities
'
' ROLE:
'   Translation integrity and synchronization engine. Ensures all non-English
'   locale worksheets remain structurally aligned with the master locale
'   (en-US), preserving key order, detecting drift, and inserting missing
'   translation rows.
'
' RESPONSIBILITIES:
'   - Key auditing:
'       • Compare locale worksheets against the master key list
'       • Detect out-of-sync rows and report discrepancies
'   - Synchronization:
'       • Rebuild locale sheets to match master ordering
'       • Insert missing keys with default English text
'       • Flag untranslated rows for downstream filtering
'   - Structural preservation:
'       • Maintain row-to-row alignment across all locale sheets
'       • Ensure control IDs, compact labels, verbose labels, screentips,
'         and supertips remain in consistent column positions
'
' ARCHITECTURAL NOTES:
'   - Uses Scripting.Dictionary for fast key-to-row mapping.
'   - Operates directly on locale worksheets; no dependency on the runtime
'     localization cache.
'   - Designed for translator workflows and release-engineering validation.
'   - Ensures that the i18n engine (modLocalize) always receives clean,
'     predictable locale sheets.
'
' USAGE:
'   - Run in VBA editor before shipping a new release to ensure all locales
'     are aligned.
'   - Used by translators to identify missing or outdated strings.
'   - Supports automated QA workflows for localization completeness.
'
' RELATED WIKI PAGES:
'   - Locale Worksheet Specification
'   - Translation Workflow & Release Checklist
'   - Localization Integrity Tools
' =============================================================================

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
        lastRow = .Cells.item(.Cells.count).row
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
        lastRow = .Cells.item(.Cells.count).row
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
        value = masterDictionary.item(key)
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
        SetCell sheetToSync, row, 2, GetCell(masterSheet, masterDictionary.item(dictKey), 2) ' Compact Control Labels
        SetCell sheetToSync, row, 3, GetCell(masterSheet, masterDictionary.item(dictKey), 3) ' Verbose Control Labels
        SetCell sheetToSync, row, 4, GetCell(masterSheet, masterDictionary.item(dictKey), 4) ' Screentip
        SetCell sheetToSync, row, 5, GetCell(masterSheet, masterDictionary.item(dictKey), 5) ' Supertip
        SetCell sheetToSync, row, 6, masterDictionary.item(dictKey)  ' Master row number
        SetCell sheetToSync, row, 7, "Requires translation"     ' Flag for translation
        row = row + 1
    Next
    
    Debug.Print sheetToSync & " has been synced to " & masterSheet

End Sub

