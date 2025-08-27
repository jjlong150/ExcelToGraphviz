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

' Function to parse a Graphviz arrowhead string into individual arrowheads (max 3)
' Input: strArrowheads - concatenated string like "normalonormalobox"
' Output: Array of strings containing up to 3 arrowhead names, or empty array if parsing fails
Public Function ParseGraphvizArrowheads(strArrowheads As String) As String()
    Dim validArrowheads As Variant
    Dim result() As String
    Dim currentString As String
    Dim i As Integer
    Dim maxArrowheads As Integer
    Dim parseSuccessful As Boolean
    
    ' Define base shapes and modifiers
    Dim baseShapes As Variant
    Dim modifiers As Variant
    baseShapes = Array("box", "crow", "diamond", "dot", "inv", "none", "normal", "tee", "vee", _
                       "odot", "invdot", "invodot", "ediamond", "open", "halfopen", "empty", "invempty")
    modifiers = Array("", "o", "l", "r") ' Empty string for base shapes without modifiers
    
    ' Generate valid arrowhead combinations
    Dim arrowheadList As Collection
    Set arrowheadList = New Collection
    Dim base As Variant, modifier As Variant
    For Each base In baseShapes
        For Each modifier In modifiers
            arrowheadList.Add modifier & base
        Next modifier
    Next base
    
    ' Convert collection to array for faster access
    ReDim validArrowheads(0 To arrowheadList.count - 1)
    For i = 1 To arrowheadList.count
        validArrowheads(i - 1) = arrowheadList(i)
    Next i
    
    ' Set maximum number of arrowheads to parse
    maxArrowheads = 3
    ReDim result(0 To maxArrowheads - 1)
    
    ' Initialize variables
    currentString = LCase(Trim(strArrowheads)) ' Convert to lowercase
    i = 0
    parseSuccessful = False
    
    ' Try parsing with recursive helper function
    Dim tempResult() As String
    ReDim tempResult(0 To maxArrowheads - 1)
    parseSuccessful = ParseArrowheadsRecursive(currentString, validArrowheads, tempResult, 0, maxArrowheads)
    
    If parseSuccessful Then
        ' Copy results
        For i = 0 To maxArrowheads - 1
            If tempResult(i) = "" Then Exit For
            result(i) = tempResult(i)
        Next i
        If i > 0 Then
            ReDim Preserve result(0 To i - 1)
        Else
            ReDim result(0 To 0)
            result(0) = ""
        End If
    Else
        ReDim result(0 To 0)
        result(0) = ""
    End If
    
    ParseGraphvizArrowheads = result
End Function

' Recursive helper function to parse arrowheads, handling ambiguities
Private Function ParseArrowheadsRecursive(ByVal currentString As String, _
                                         validArrowheads As Variant, _
                                         result() As String, _
                                         ByVal index As Integer, _
                                         ByVal maxArrowheads As Integer) As Boolean
    If currentString = "" And index > 0 Then
        ParseArrowheadsRecursive = True
        Exit Function
    End If
    
    If index >= maxArrowheads Then
        ParseArrowheadsRecursive = False
        Exit Function
    End If
    
    Dim arrowhead As Variant
    Dim found As Boolean
    found = False
    
    ' Try each valid arrowhead at the current position
    For Each arrowhead In validArrowheads
        If Len(arrowhead) > 0 And Left(currentString, Len(arrowhead)) = arrowhead Then
            result(index) = arrowhead
            Dim remainingString As String
            remainingString = Mid(currentString, Len(arrowhead) + 1)
            
            ' Recurse on the remaining string
            If ParseArrowheadsRecursive(remainingString, validArrowheads, result, index + 1, maxArrowheads) Then
                ParseArrowheadsRecursive = True
                Exit Function
            End If
            result(index) = "" ' Reset on backtrack
        End If
    Next arrowhead
    
    ParseArrowheadsRecursive = False
End Function

' Test subroutine to verify parsing
Public Sub TestParseGraphvizArrowheads()
    Dim testString As String
    Dim result() As String
    Dim i As Integer
    
    ' Test case
    testString = "normalonormalobox"
    result = ParseGraphvizArrowheads(testString)
    
    ' Output results to Immediate Window
    Debug.Print "Input: " & testString
    If result(0) = "" Then
        Debug.Print "No valid arrowheads found"
    Else
        For i = 0 To UBound(result)
            Debug.Print "Arrowhead " & i + 1 & ": " & result(i)
        Next i
    End If
End Sub

Public Function ParseGraphvizPackmode(packmode As String) As Object
    Dim result As Dictionary
    Dim regex As Object
    Dim matches As Object
    Dim validModes As Variant
    Dim i As Integer
    Dim mode As String
    Dim Flags As String
    Dim suffix As String
    
    ' Initialize result dictionary
    Set result = New Dictionary
    result.Add "Mode", ""
    result.Add "Flags", ""
    result.Add "Suffix", ""
    result.Add "IsValid", False
    
    ' Define valid modes
    validModes = Array("node", "cluster", "graph", "array")
    
    ' Convert input to lowercase for consistency
    packmode = LCase(Trim(packmode))
    
    ' Check if packmode is a simple mode
    For i = 0 To UBound(validModes)
        If packmode = validModes(i) Then
            result("Mode") = packmode
            result("IsValid") = True
            Set ParseGraphvizPackmode = result
            Exit Function
        End If
    Next i
    
    ' Check for array mode with optional flags and suffix
    ' Pattern: array(_[flags])?([0-9]+)?
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Pattern = "^array(_[a-z]*)?([0-9]+)?$"
        .IgnoreCase = True
    End With
    
    If regex.Test(packmode) Then
        Set matches = regex.Execute(packmode)
        If matches.count > 0 Then
            result("Mode") = "array"
            If matches(0).SubMatches.count >= 1 And Not IsEmpty(matches(0).SubMatches(0)) Then
                Flags = Mid(matches(0).SubMatches(0), 2) ' Remove leading underscore
                ' Validate flags (only 'u', 'c', 't', 'b', 'l', 'r' allowed)
                Dim validFlags As String
                validFlags = "uctblr"
                For i = 1 To Len(Flags)
                    If InStr(validFlags, Mid(Flags, i, 1)) = 0 Then
                        Debug.Print "Invalid flag found: " & Mid(Flags, i, 1)
                        Set ParseGraphvizPackmode = result
                        Exit Function
                    End If
                Next i
                result("Flags") = Flags
            End If
            If matches(0).SubMatches.count >= 2 And Not IsEmpty(matches(0).SubMatches(1)) Then
                suffix = matches(0).SubMatches(1)
                result("Suffix") = suffix
            End If
            result("IsValid") = True
            'Debug.Print "Parsed as array mode: Mode=" & result("Mode") & ", Flags=" & result("Flags") & ", Suffix=" & result("Suffix")
        End If
    Else
        Debug.Print "Invalid packmode string: " & packmode
    End If
    
    Set ParseGraphvizPackmode = result
End Function

' Test subroutine to verify parsing
Public Sub TestParseGraphvizPackmode()
    Dim testStrings As Variant
    Dim result As Object
    Dim i As Integer
    Dim testString As String
    
    ' Test cases
    testStrings = Array("node", "array_c4", "array_u", "array_ctblr8", "invalid", "array_x")
    
    ' Test each case
    For i = 0 To UBound(testStrings)
        testString = testStrings(i)
        Set result = ParseGraphvizPackmode(testString)
        Debug.Print "Input: " & testString
        Debug.Print "Mode: " & result("Mode")
        Debug.Print "Flags: " & result("Flags")
        Debug.Print "Suffix: " & result("Suffix")
        Debug.Print "IsValid: " & result("IsValid")
        Debug.Print "------------------------"
    Next i
End Sub
