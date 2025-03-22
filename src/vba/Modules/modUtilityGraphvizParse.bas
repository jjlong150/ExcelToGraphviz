Attribute VB_Name = "modUtilityGraphvizParse"
'@IgnoreModule UseMeaningfulName
'@Folder("Utility.Excel")
Option Explicit

Public Function ParseAttributeString(ByVal attributes As String) As Dictionary
    Dim pipedAttributes As String
    pipedAttributes = AddPipeDelimitersToAttributeString(attributes)
    Set ParseAttributeString = ParsePipedAttributeString(pipedAttributes)
End Function

Private Function AddPipeDelimitersToAttributeString(ByVal attributes As String) As String
    Dim inValue As Boolean
    Dim equalsFound As Boolean
    Dim oneChar As String
    Dim pipedAttributes As String
    
    inValue = False
    '@Ignore AssignmentNotUsed
    equalsFound = False
    
    pipedAttributes = vbNullString
    
    Dim i As Long
    ' Examine the attribute string once character at a time
    For i = 1 To Len(attributes)
        ' Grab a character
        oneChar = Mid$(attributes, i, 1)
        
        If oneChar = "=" Then
            ' We are transitioning from the key to the value
            pipedAttributes = pipedAttributes & oneChar
            equalsFound = True
        ElseIf oneChar = """" Then
            ' We found a quote. It can either be either the start or
            ' the end of the value string.
            
            If equalsFound Then
                ' We are are past an equals character,
                If inValue Then
                    ' if inValue is true this is the second quote
                    ' since the equals character, so append a pipe
                    ' character in place of the quote.
                    inValue = False
                    equalsFound = False
                    pipedAttributes = pipedAttributes & "|"
                Else
                    ' This is the first quote found to the right of
                    ' an equals character, meaning this is the start
                    ' of a value string.
                    inValue = True
                End If
            End If
        ElseIf oneChar = ";" Or oneChar = "," Or oneChar = " " Then
            ' An optional attribute terminator was encountered
            If equalsFound Then
                If inValue Then
                    ' allow commas in the value, ignore it
                    inValue = True
                    pipedAttributes = pipedAttributes & oneChar
                Else
                    ' honor the terminator string, append a pipe
                    inValue = False
                    equalsFound = False
                    pipedAttributes = pipedAttributes & "|"
                End If
            End If
        Else
            ' Ordinary, boring character. Concatenate it to the piped
            ' attribute string.
            pipedAttributes = pipedAttributes & oneChar
        End If
    Next i
    
    AddPipeDelimitersToAttributeString = pipedAttributes
End Function

Private Function ParsePipedAttributeString(ByVal pipedAttributes As String) As Dictionary

    Dim i As Long

    Dim pairs() As String
    Dim keyValue() As String
    
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary
   
    pairs = split(pipedAttributes, "|") ' create an array of key/value pairs
    For i = LBound(pairs) To UBound(pairs)
        ' Safety check to ensure the array element contains an equal string
        If InStr(1, pairs(i), "=") Then
            ' split the key/value pair into individual elemnts
            keyValue = split(pairs(i), "=")
            
            ' Ensure that an attribute only is specified once. Retain only
            ' the last value if the attribute is a duplicate.
            If dictionaryObj.Exists(Trim$(keyValue(0))) Then
                dictionaryObj.Remove (Trim$(keyValue(0)))
            End If
            ' Add the pair into the dictionary
            dictionaryObj.Add Trim$(keyValue(0)), Trim$(keyValue(1))
        End If
    Next

    Set ParsePipedAttributeString = dictionaryObj
End Function
