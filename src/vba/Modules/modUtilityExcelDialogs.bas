Attribute VB_Name = "modUtilityExcelDialogs"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Utility.Excel")

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
        MsgBox GetMessage("msgboxNoFolderPicker"), vbOKOnly, GetLabel("msgboxNoFolderPicker")
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


