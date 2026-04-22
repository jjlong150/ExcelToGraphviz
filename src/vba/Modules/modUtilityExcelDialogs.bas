Attribute VB_Name = "modUtilityExcelDialogs"
' =============================================================================
' PROJECT:   Excel to Graphviz
' MODULE:    modUtilityExcelDialogs
' COPYRIGHT: Copyright (c) 2015–2026 Jeffrey J. Long. All rights reserved.
' LAYER:     Utility / Excel Interop
'
' ROLE:
'   Cross-platform directory-selection and Save-As filename utilities.
'   Provides a unified abstraction over macOS AppleScriptTask folder pickers
'   and Windows FileDialog APIs, ensuring consistent behavior across all
'   workflows that require directory or file-output selection.
'
' RESPONSIBILITIES:
'   - Directory selection:
'       • ChooseDirectory: macOS folder picker via AppleScriptTask
'         or Windows FileDialog(msoFileDialogFolderPicker)
'       • Normalize initial directory and handle user cancellation
'   - Save-As filename selection:
'       • GetSaveAsFilename: wrapper over Application.GetSaveAsFilename
'         with project-specific defaults and trimming
'
' ARCHITECTURAL NOTES:
'   - Fully cross-platform: AppleScriptTask on macOS, FileDialog on Windows.
'   - Defensive handling of missing folder-picker support (older Office builds).
'   - Integrates with Settings, SQL, SVG, Source, and export workflows.
'   - Emits localized messages via GetMessage/GetLabel when folder pickers
'     are unavailable.
'
' USAGE:
'   - Used by Settings sheet, SQL engine, SVG export, and any workflow
'     requiring user-selected directories or output filenames.
'
' RELATED WIKI PAGES:
'   - Directory Selection (Windows/macOS)
'   - File Output & Save-As Conventions
'   - Cross-Platform UI Interop
' =============================================================================

Option Explicit

Public Function ChooseDirectory(ByVal startDir As String) As String
    ChooseDirectory = startDir
#If Mac Then
    Dim dirName As String
    dirName = RunAppleScriptTask("chooseAFolder", startDir)
    If dirName = vbNullString Then
        ' User clicked on CANCEL
    Else
        ChooseDirectory = dirName
    End If
#Else
    Dim fileDialogHandle As FileDialog
    Set fileDialogHandle = Application.FileDialog(msoFileDialogFolderPicker)
    
    If fileDialogHandle Is Nothing Then
        EmitMessage GetMessage("msgboxNoFolderPicker"), GetLabel("msgboxNoFolderPicker")
    Else
        If Trim$(startDir) = vbNullString Then
            fileDialogHandle.InitialFileName = ActiveWorkbook.path & "\"
        Else
            fileDialogHandle.InitialFileName = Trim$(startDir) & "\"
        End If
    
        '  Get the number of the button chosen
        Dim selected As Long
        selected = fileDialogHandle.show
        '@Ignore EmptyIfBlock
        If selected <> -1 Then
            ' User clicked on CANCEL)
        Else
            ' Set path of directory chosen
            ChooseDirectory = fileDialogHandle.SelectedItems.item(1)
        End If
    End If
    Set fileDialogHandle = Nothing
#End If
End Function

Public Function GetSaveAsFilename(ByRef fileFilter As String) As String
    Dim saveAsFilename As Variant

    saveAsFilename = Application.GetSaveAsFilename(fileFilter:=fileFilter, _
        title:="Save As", _
        InitialFileName:=Application.ActiveWorkbook.path)
    
    If saveAsFilename <> False Then
        GetSaveAsFilename = Trim$(saveAsFilename)
    End If
End Function


