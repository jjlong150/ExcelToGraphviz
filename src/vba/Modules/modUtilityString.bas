Attribute VB_Name = "modUtilityString"
'@IgnoreModule UseMeaningfulName
' Copyright (c) 2015-2024 Jeffrey J. Long. All rights reserved

'@Folder("Utility.String")

Option Explicit

Public Function EndsWith(ByVal sourceString As String, ByVal endingString As String) As Boolean
    Dim endingLen As Long
    endingLen = Len(endingString)
    EndsWith = (Right$(Trim$(UCase$(sourceString)), endingLen) = UCase$(endingString))
End Function

Public Function StartsWith(ByVal sourceString As String, ByVal startingString As String) As Boolean
    Dim startLen As Long
    startLen = Len(startingString)
    StartsWith = (Left$(Trim$(UCase$(sourceString)), startLen) = UCase$(startingString))
End Function

Public Function AddQuotes(ByVal Text As String) As String
    AddQuotes = Join(Array(Chr$(34), Text, Chr$(34)), vbNullString)
End Function

Public Function AddQuotesConditionally(ByVal Text As String) As String
    If Text Like "*[!A-Za-z0-9]*" Then
        AddQuotesConditionally = Join(Array(Chr$(34), Text, Chr$(34)), vbNullString)
    Else
        AddQuotesConditionally = Text
    End If
End Function

Public Function GetStringBetweenDelimiters(ByVal inString As String, ByVal leftDelimiter As String, ByVal rightDelimiter As String) As String

    GetStringBetweenDelimiters = inString
    
    Dim outputString As String
    outputString = Trim$(inString)
    
    If Len(outputString) >= Len(leftDelimiter) + Len(rightDelimiter) Then
        If StartsWith(outputString, leftDelimiter) And EndsWith(outputString, rightDelimiter) Then
            outputString = Left$(outputString, Len(outputString) - Len(rightDelimiter))
            outputString = Right$(outputString, Len(outputString) - Len(leftDelimiter))
            GetStringBetweenDelimiters = outputString
        End If
    End If
    
End Function

Public Function GetStringTokenAtPosition(ByVal inputString As String, ByVal tokenSeparator As String, ByVal tokenPosition As Long) As String

    Dim token As String
    
    If InStr(inputString, tokenSeparator) Then
        Dim tokenArray() As String
        tokenArray = split(inputString, tokenSeparator)
        If tokenPosition - 1 <= UBound(tokenArray) Then
            token = tokenArray(tokenPosition - 1)
        End If
    End If
    
    GetStringTokenAtPosition = token
End Function

Public Sub AddNameValue(ByRef styleAttributes As String, ByVal attrName As String, ByVal attrValue As String)
    ' If a value is present, write it out as a name/value pair
    If Trim$(attrName) <> vbNullString Then
        styleAttributes = Join(Array(styleAttributes, " ", Trim$(attrName), "=", attrValue), vbNullString)
    End If
End Sub

' @method WrapText
' @param {Range} items A set of cells which should be concatenated and wrapped as a single text string
' @param {Long} wrapLength The desired maximum number of characters in a line segment. Line segments may exceed this size under certain circumstances.
' @param {String} lineEnding The character(s) to append to each line segment when wrapping the text.
'@Ignore ProcedureNotUsed
Public Function WrapText(ByVal itemIds As Range, Optional ByVal wrapLength As Long = 1, Optional ByVal lineEnding As String = "\n") As String
    Dim Text As String
    Text = vbNullString
    
    ' Concatenate the range of cell values into one long string
    Dim item As Range
    For Each item In itemIds.Cells
        Text = Text & " " & Trim$(item.value)
    Next item
    
    ' Split the one long string using the length and lineEnding specified by the caller
    WrapText = SplitText(Text, wrapLength, lineEnding)
End Function

Public Function SplitMultilineText(ByVal fullText As String, ByVal wrapLength As Long, Optional ByVal lineEnding As String = vbNewLine) As String
    Dim Text As String
    Text = Trim$(fullText)
    
    ' Convert all instances of multiple " " spaces to a single space
    Do While InStr(Text, "  ")
        Text = replace(Text, "  ", " ")
    Loop
    
    ' Convert all the symbolic line breaks to a line feed character
    Text = replace(Text, "\n", Chr$(10))
    Text = replace(Text, "\r", Chr$(10))
    Text = replace(Text, "\l", Chr$(10))
    Text = replace(Text, "<br/>", Chr$(10))
    
    ' Break the scrubbed string into an array of string segments
    Dim stringSplitArray() As String
    stringSplitArray = split(Text, Chr$(10))
    
    ' Process each segment individually, splitting long strings
    Dim i As Long
    For i = LBound(stringSplitArray, 1) To UBound(stringSplitArray, 1)
        stringSplitArray(i) = SplitText(stringSplitArray(i), wrapLength, lineEnding)
    Next i
    
    ' Reassemble the string from the split string segments
    For i = LBound(stringSplitArray, 1) To UBound(stringSplitArray, 1) - 1
        SplitMultilineText = SplitMultilineText & stringSplitArray(i) & lineEnding
    Next i
    
    ' Don't add the line ending to the last segment
    SplitMultilineText = SplitMultilineText & stringSplitArray(UBound(stringSplitArray, 1))
    
End Function

Public Function SplitText(ByVal fullText As String, ByVal wrapLength As Long, Optional ByVal lineEnding As String = vbNewLine) As String
    Dim Text As String
    Dim remainder As String
    
    Text = Trim$(fullText)
    
    ' Convert all instances of multiple " " spaces to a single space
    Do While InStr(Text, "  ")
        Text = replace(Text, "  ", " ")
    Loop
    
    ' Split the text into segments, and insert the lineEnding value between segments
    SplitText = GetTextSegment(Text, wrapLength, remainder)
    Do While Len(remainder) > 0
        Text = remainder
        SplitText = SplitText & lineEnding & GetTextSegment(Text, wrapLength, remainder)
    Loop
End Function

Private Function GetTextSegment(ByVal Text As String, ByVal segmentLength As Long, ByRef remainder As String) As String
    Dim wrapLength As Long
    Dim i As Long
    Dim positionOfSpace As Long
    
    ' Scrub input parameter
    If segmentLength < 1 Then
        wrapLength = 1
    Else
        wrapLength = segmentLength
    End If
    
    ' Make sure values are returned
    GetTextSegment = Text
    remainder = vbNullString

    ' Is the text shorter than the segment length?
    If Len(Text) <= wrapLength Then
        Exit Function
    End If
    
    ' Did text segment end on a full word?
    If Mid$(Text, wrapLength + 1, 1) = " " Then  ' Yes - full word
        GetTextSegment = Left$(Text, wrapLength)
        remainder = Trim$(Mid$(Text, wrapLength + 1))
        Exit Function
    End If
    
    ' Text segment did not end on a full word, see if space " " exists in left segnemt
    For i = wrapLength To 1 Step -1
        If Mid$(Text, i, 1) = " " Then
            positionOfSpace = i
            Exit For
        End If
    Next i
        
    ' If positionOfSpace > 0 a space " " was found within the desired segment length.
    ' Return the string up to where the space " " was located.
    If positionOfSpace > 0 Then
        GetTextSegment = Trim$(Left$(Text, positionOfSpace - 1))
        remainder = Trim$(Mid$(Text, positionOfSpace))
        Exit Function
    End If
        
    ' If we have gotten this far, then there is not a space " " within the desired
    ' segment length. GetTextSegment must exceed the desired segment length.
    ' Look for the first space " " to the right of the desired segment length.
    For i = 1 To Len(Text)
        If Mid$(Text, i, 1) = " " Then
            positionOfSpace = i
            Exit For
        End If
    Next i
                
    ' If positionOfSpace = 0 then there are not any space " " characters
    ' in the text. Return all the text as the segment.
    If positionOfSpace = 0 Then ' No " " found in text
        GetTextSegment = Text
        remainder = vbNullString
        Exit Function
    End If
    
    ' If we got this far, a space " " was found to the right
    ' of a desired segment length
    If positionOfSpace < Len(Text) Then
        GetTextSegment = Mid$(Text, 1, positionOfSpace - 1)
        remainder = Trim$(Mid$(Text, positionOfSpace))
    End If
End Function

Public Function ScrubText(ByVal rawData As String) As String
    If rawData = Chr$(34) & Chr$(34) Then
        ScrubText = vbNullString                 ' "" to blank a label
    Else
        ScrubText = replace(rawData, Chr$(10), NEWLINE) ' Chr(10) 0x0a LF  Line Feed
        ScrubText = replace(ScrubText, "\" & Chr$(34), Chr$(34)) ' In case they already escaped the double quote
        ScrubText = replace(ScrubText, Chr$(34), "\" & Chr$(34)) ' Chr(34)      " Double quotes (or speech marks)
    End If
End Function


