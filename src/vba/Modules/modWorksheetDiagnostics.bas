Attribute VB_Name = "modWorksheetDiagnostics"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.Diagnostics")
'@IgnoreModule ProcedureNotUsed

Option Explicit

Public Sub ReportDiagnostics()
    ' Show the hourglass cursor
    Application.Cursor = xlWait
    DoEvents

    ' Turn off screen updating and events
    OptimizeCode_Begin
    
    ' Current Workbook File Name
    DiagnosticsSheet.Range(DIAGNOSTICS_WORKBOOK_NAME).value = ThisWorkbook.name
    
    ' Operating System
    DiagnosticsSheet.Range(DIAGNOSTICS_APPLICATION_OPERATING_SYSTEM).value = Application.OperatingSystem
    
    ' Excel version and build number
    DiagnosticsSheet.Range(DIAGNOSTICS_APPLICATION_VERSION).value = Application.version & Application.Build
    
     ' Graphviz version number
    DiagnosticsSheet.Range(DIAGNOSTICS_GRAPHVIZ_VERSION).value = GetGraphvizVersion
   
    ' User name as seen by Excel Application
    DiagnosticsSheet.Range(DIAGNOSTICS_APPLICATION_USER_NAME).value = Application.username
    
    ' User name as returned by OS
    DiagnosticsSheet.Range(DIAGNOSTICS_USERNAME).value = GetUsername()

    ' Temp file directory
    DiagnosticsSheet.Range(DIAGNOSTICS_TEMP_DIRECTORY).value = GetTempDirectory()
        If DirectoryExists(GetTempDirectory()) Then
        DiagnosticsSheet.Range(DIAGNOSTICS_TEMP_DIRECTORY_EXISTS).value = 1
    Else
        DiagnosticsSheet.Range(DIAGNOSTICS_TEMP_DIRECTORY_EXISTS).value = 0
    End If

    ' Style Designer Image Cache Directory of font preview images
    DiagnosticsSheet.Range(DIAGNOSTICS_FONT_IMAGE_DIR).value = GetFontImageDir()
    If DirectoryExists(GetFontImageDir()) Then
        DiagnosticsSheet.Range(DIAGNOSTICS_FONT_IMAGE_DIR_EXISTS).value = 1
    Else
        DiagnosticsSheet.Range(DIAGNOSTICS_FONT_IMAGE_DIR_EXISTS).value = 0
    End If
    
    ' Style Designer Image Cache Directory of color scheme preview images
    DiagnosticsSheet.Range(DIAGNOSTICS_COLOR_IMAGE_DIR).value = GetColorImageDir()
    If DirectoryExists(GetColorImageDir()) Then
        DiagnosticsSheet.Range(DIAGNOSTICS_COLOR_IMAGE_DIR_EXISTS).value = 1
    Else
        DiagnosticsSheet.Range(DIAGNOSTICS_COLOR_IMAGE_DIR_EXISTS).value = 0
    End If
    
    ' Name of the environment variable which can be defined to point to a folder of images
    DiagnosticsSheet.Range(DIAGNOSTICS_IMAGE_ENV_VARIABLE_NAME).value = "ExcelToGraphvizImages"
    
    ' The folder of images pointed to by the environment variable
    DiagnosticsSheet.Range(DIAGNOSTICS_EXCELTOGRAPHVIZ_IMAGE_DIRECTORY).value = GetExcelToGraphvizImageDirectory()
    If DirectoryExists(GetExcelToGraphvizImageDirectory()) Then
        DiagnosticsSheet.Range(DIAGNOSTICS_EXCELTOGRAPHVIZ_IMAGE_DIRECTORY_EXISTS).value = 1
    Else
        DiagnosticsSheet.Range(DIAGNOSTICS_EXCELTOGRAPHVIZ_IMAGE_DIRECTORY_EXISTS).value = 0
    End If
    
    ' The directory paths to be searched for images when creating a graph
    DiagnosticsSheet.Range(DIAGNOSTICS_IMAGE_PATH).value = GetImagePath()
    If DirectoryExists(GetImagePath()) Then
        DiagnosticsSheet.Range(DIAGNOSTICS_IMAGE_PATH_EXISTS).value = 1
    Else
        DiagnosticsSheet.Range(DIAGNOSTICS_IMAGE_PATH_EXISTS).value = 0
    End If
              
#If Mac Then
    ' Security sandbox where applescript files must reside to be executed by AppleScriptTask command
    DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FOLDER).value = "/Users/" & GetUsername & "/Library/Application Scripts/com.microsoft.Excel"
    If DirectoryExists(DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FOLDER).value) Then
        DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FOLDER_EXISTS).value = 1
    Else
        DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FOLDER_EXISTS).value = 0
    End If

    ' Name of file containing the AppleScriptTask commands needed by the Excel version of Excel to Graphviz
    DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FILE).value = "ExcelToGraphviz.applescript"
    
    ' Was the file of AppleScriptTask commands found in the sandbox directory?
    Dim applescriptfile As String
    applescriptfile = DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FOLDER).value & "/" & DiagnosticsSheet.Range("Diagnostics.AppleScriptFile").value
    If FileExists(applescriptfile) Then
        DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FILE_EXISTS).value = 1
        ' Version of the AppleScriptTask commands
        DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_VERSION).value = RunAppleScriptTask("getVersion", vbNullString)
    Else
        DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FILE_EXISTS).value = 0
        DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_VERSION).value = vbNullString
    End If

#Else
    DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FOLDER).value = vbNullString
    DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FOLDER_EXISTS).value = 0
    DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FILE).value = vbNullString
    DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_FILE_EXISTS).value = 0
    DiagnosticsSheet.Range(DIAGNOSTICS_APPLE_SCRIPT_VERSION).value = vbNullString
#End If
    
    ' Turn on screen updating and events
    OptimizeCode_End
    
    ' Reset the cursor back to the default
    Application.Cursor = xlDefault
End Sub

Public Sub ClearDiagnostics()
    DiagnosticsSheet.Range("D4:D15").ClearContents
    DiagnosticsSheet.Range("D19:D21").ClearContents
End Sub

Private Sub DeleteFolderContents(ByVal folder As String)
#If Mac Then
    On Error Resume Next
    Kill folder & "/*"
    On Error GoTo 0
#Else
    Dim fileSystemObject As Object
    Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
    If Not fileSystemObject Is Nothing Then
        fileSystemObject.DeleteFile folder & "\*.*", True
        Set fileSystemObject = Nothing
    End If
#End If

End Sub

Public Sub ClearFontImageFolder()
    Dim folder As String
    folder = GetFontImageDir()
    DeleteFolderContents folder
End Sub

Public Sub ClearColorsImageFolder()
    Dim folder As String
    folder = GetColorImageDir()
    DeleteFolderContents folder
End Sub

Public Function GetGraphvizVersion() As String
#If Mac Then
    GetGraphvizVersion = RunAppleScriptTask("runDot", "-V")
#Else
    Dim stdOut As String
    Dim stdErr As String
    ExecuteAndCapture "dot -V", stdOut, stdErr

    GetGraphvizVersion = replace(stdErr, vbNewLine, vbNullString)
#End If
End Function

Public Sub TestGetGraphvizVersion()
    Debug.Print "|" & GetGraphvizVersion() & "|"
End Sub

