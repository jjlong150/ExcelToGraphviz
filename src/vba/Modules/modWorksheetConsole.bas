Attribute VB_Name = "modWorksheetConsole"
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Relationship Visualizer.Sheets.Console")
'@IgnoreModule UseMeaningfulName

Option Explicit

Public Sub ClearConsoleWorksheet()
    
    ' Determine the range of the cells which need to be cleared
    Dim lastRow As Long
    With ConsoleSheet.UsedRange
        lastRow = .Cells.item(.Cells.count).row
    End With

    ' Remove any existing content
    Dim cellRange As String
    cellRange = "A1:B" & lastRow
    ConsoleSheet.Range(cellRange).ClearContents
    
End Sub

Public Sub DisplayTextOnConsoleWorksheet(ByVal dotCommand As String, ByVal textBlob As String)

    ' Exit if logging is disabled
    If Not GetCellBoolean(SettingsSheet.name, SETTINGS_LOG_TO_CONSOLE) Then
        Exit Sub
    End If

    ' Clear console if not in append mode
    If Not GetCellBoolean(SettingsSheet.name, SETTINGS_APPEND_CONSOLE) Then
        ClearConsoleWorksheet
    End If
        
    ' Initialize row counter to first unused row
    Dim row As Long
    With ConsoleSheet.UsedRange
        row = .Cells.item(.Cells.count).row
    End With
    
    ' Leave some white space between invocations
    If row = 1 Then
        row = row + 1
    Else
        row = row + 2
    End If
        
    ' Log the command used to invoke Graphviz
    Dim commandExecuted As String
#If Mac Then
    ' dot command is actually specified in the ExcelToGraphviz.applescript file. Fake it for the console
    commandExecuted = "dot " & dotCommand
#Else
    commandExecuted = dotCommand
#End If
    ConsoleSheet.Cells.item(row, 1).value = ">"
    ConsoleSheet.Cells.item(row, 2).value = commandExecuted
    row = row + 2

    ' Split the text into an array of lines
    Dim parsedText As Variant
#If Mac Then
    parsedText = split(textBlob, vbCr) ' lines are delimited by Carriage Return
#Else
    parsedText = split(textBlob, vbLf) ' lines are delimited by Line Feed
#End If
    
    If UBound(parsedText) >= 0 Then
        ' Transfer the array of lines to the worksheet in one swift action
        Dim writeToRange As String
        writeToRange = "B" & row & ":B" & (row + (UBound(parsedText) - LBound(parsedText)))
        ConsoleSheet.Range(writeToRange).value = Application.Transpose(parsedText)
    End If
    
End Sub

Public Sub CopyConsoleToClipboard()
#If Not Mac Then

    ' Pull all the rows into a single string
    Dim consoleMessage As String
    consoleMessage = vbNullString
    
    Dim lastRow As Long
    With ConsoleSheet.UsedRange
        lastRow = .Cells.item(.Cells.count).row
    End With

    Dim i As Long
    For i = 1 To lastRow
        consoleMessage = consoleMessage & ConsoleSheet.Cells.item(i, 2).value & vbLf
    Next i

    If ClipBoard_SetData(consoleMessage) Then
        UpdateStatusBarForNSeconds GetMessage("statusbarClipboardCopyConsoleSuccess"), 5
    Else
        UpdateStatusBarForNSeconds GetMessage("statusbarClipboardCopyConsoleFailed"), 5
    End If
    
#End If
End Sub

Public Sub ConsoleWorksheetToFile(ByVal fileName As String)

    Dim rowNumber As Long
    Dim lastRow As Long
    With ConsoleSheet.UsedRange
        lastRow = .Cells.item(.Cells.count).row
    End With
    
#If Mac Then
    Dim consoleText As String
    consoleText = vbNullString
    
    For rowNumber = 1 To lastRow
        consoleText = consoleText & ConsoleSheet.Cells(rowNumber, 2).value & vbCr
    Next rowNumber
    
    WriteTextToFile consoleText, fileName
#Else
    
    ' Output file objects
    Dim utf8Stream As Object
    Set utf8Stream = CreateObject("ADODB.Stream")
    If utf8Stream Is Nothing Then GoTo EndMacro
    
    Dim binaryStream As Object
    Set binaryStream = CreateObject("ADODB.Stream")
    If binaryStream Is Nothing Then GoTo EndMacro
    
    ' Initialize the utf8Stream object
    utf8Stream.Type = StreamTypeEnum.adTypeText
    utf8Stream.Charset = UTF8_CHARSET
    utf8Stream.Open
    
    For rowNumber = 1 To lastRow
        utf8Stream.WriteText ConsoleSheet.Cells.item(rowNumber, 2).value & vbLf
    Next rowNumber
    
    ' Initialize the object which is used to remove the Byte Order Mark (BOM) from the UTF-8 stream
    binaryStream.Type = StreamTypeEnum.adTypeBinary
    binaryStream.mode = ConnectModeEnum.adModeReadWrite
    binaryStream.Open

    ' Position the start of the utf8 stream past the Byte Order Mark (BOM) (i.e. BOM = first 3 bytes)
    ' and copy the contents to the binary stream
    utf8Stream.position = 3
    utf8Stream.CopyTo binaryStream
    
    ' Write out UTF-8 data without the BOM
    binaryStream.SaveToFile fileName, SaveOptionsEnum.adSaveCreateOverWrite

EndMacro:
    ' Clean up our objects so we don't get a memory leak
    If Not (utf8Stream Is Nothing) Then
        If (utf8Stream.State And ObjectStateEnum.adStateOpen) = ObjectStateEnum.adStateOpen Then utf8Stream.Close
        Set utf8Stream = Nothing
    End If
    
    If Not (binaryStream Is Nothing) Then
        If (binaryStream.State And ObjectStateEnum.adStateOpen) = ObjectStateEnum.adStateOpen Then binaryStream.Close
        Set binaryStream = Nothing
    End If
#End If
End Sub

'@Ignore ParameterNotUsed
Public Sub SaveConsoleToFile()
    Dim fileName As String
    
#If Mac Then
    fileName = RunAppleScriptTask("getSaveAsFileName", ".txt")
#Else
    fileName = GetSaveAsFilename("Text Files (*.txt), *txt")
#End If

    If fileName <> vbNullString Then
        ConsoleWorksheetToFile (fileName)
        MsgBox GetMessage("msgboxSavedToFile") & vbNewLine & fileName, vbOKOnly, GetMessage(MSGBOX_PRODUCT_TITLE)
    End If
End Sub

Public Function RunGraphvizInVerboseMode() As Boolean
    RunGraphvizInVerboseMode = ConsoleSheet.visible And GetSettingBoolean(SETTINGS_GRAPHVIZ_VERBOSE) And GetSettingBoolean(SETTINGS_LOG_TO_CONSOLE)
End Function

