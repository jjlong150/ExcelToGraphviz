Attribute VB_Name = "modUtilityUTF8File"
' Copyright (c) 2015-2022 Jeffrey J. Long. All rights reserved

'@Folder("Utility.Excel")
'@IgnoreModule ProcedureNotUsed

Option Explicit

Public Sub WriteTextToUTF8FileFileWithoutBOM(ByVal textToWrite As String, ByVal fileNameToWriteTo As String)
#If Mac Then
    MsgBox "Sub 'WriteTextToUTF8FileFileWithoutBOM' is not supported on MacOS"
#Else
    ' Output file objects
    Dim utf8Stream As Object
    Set utf8Stream = CreateObject("ADODB.Stream")
    
    Dim binaryStream As Object
    Set binaryStream = CreateObject("ADODB.Stream")
    
    ' Initialize the utf8Stream object
    utf8Stream.Type = StreamTypeEnum.adTypeText
    utf8Stream.Charset = UTF8_CHARSET
    utf8Stream.Open
    
    ' Write the text to the stream
    utf8Stream.WriteText textToWrite
    
    ' Initialize the object which is used to remove the Byte Order Mark (BOM) from the UTF-8 stream
    binaryStream.Type = StreamTypeEnum.adTypeBinary
    binaryStream.mode = ConnectModeEnum.adModeReadWrite
    binaryStream.Open

    ' Position the start of the utf8 stream past the Byte Order Mark (BOM) (i.e. BOM = first 3 bytes)
    ' and copy the contents to the binary stream
    utf8Stream.position = 3
    utf8Stream.CopyTo binaryStream
    
    ' Write out UTF-8 data without the BOM
    binaryStream.SaveToFile fileNameToWriteTo, SaveOptionsEnum.adSaveCreateOverWrite

    ' Clean up our objects so we don't get a memory leak
    utf8Stream.Close
    Set utf8Stream = Nothing

    binaryStream.Close
    Set binaryStream = Nothing
#End If
End Sub

Public Sub WriteTextToUTF8FileFileWithBOM(ByVal textToWrite As String, ByVal fileNameToWriteTo As String)
#If Mac Then
    MsgBox "Sub 'WriteTextToUTF8FileFileWithBOM' is not supported on MacOS"
#Else
    ' Output file objects
    Dim utf8Stream As Object
    Set utf8Stream = CreateObject("ADODB.Stream")
    
    ' Initialize the utf8Stream object
    utf8Stream.Type = StreamTypeEnum.adTypeText
    utf8Stream.Charset = UTF8_CHARSET
    utf8Stream.Open
    
    ' Write the text to the stream
    utf8Stream.WriteText textToWrite
    
    ' Write out UTF-8 data without the BOM
    utf8Stream.SaveToFile fileNameToWriteTo, SaveOptionsEnum.adSaveCreateOverWrite

    ' Clean up our objects so we don't get a memory leak
    utf8Stream.Close
    Set utf8Stream = Nothing
#End If
End Sub

Public Function ReadUTF8File(ByVal filename As String) As String
#If Mac Then
    ReadUTF8File = ""
    
    Dim fileNum As Integer
    Dim dataLine As String

    fileNum = FreeFile()

    Open filename For Input As #fileNum

    While Not EOF(fileNum)
        Line Input #fileNum, dataLine ' read in data 1 line at a time
        ReadUTF8File = ReadUTF8File & dataLine
    Wend
    
    Close #fileNum
#Else
    ' Read the file into a stream object
    Dim utf8Stream As Object
    Set utf8Stream = CreateObject("ADODB.Stream")
    
    utf8Stream.Charset = UTF8_CHARSET
    utf8Stream.Open
    utf8Stream.LoadFromFile filename
    
    ' Pass back the file contents
    ReadUTF8File = utf8Stream.ReadText
          
    ' Clean up our objects so we don't get a memory leak
    utf8Stream.Close
    Set utf8Stream = Nothing
#End If
End Function

