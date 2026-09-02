Attribute VB_Name = "modUtilityDictionary"
' ==========================================================================
' FUNCTION: MergeDictionaries
'
' PURPOSE:
'   Combines two normalized attribute Dictionaries into a single unified set
'   of Graphviz attributes. Implements override precedence rules used across
'   all style pipelines (node, edge, and graph), ensuring that template-level
'   attributes supersede row-level attributes when both define the same key.
'
' TECHNICAL WORKFLOW:
'   1. BASE DICTIONARY CLONE:
'        - Creates a new Dictionary initialized with all key/value pairs from
'          the base attribute set.
'        - Preserves normalized keys and pre-sanitized values exactly as
'          provided by upstream parsing.
'
'   2. OVERRIDE APPLICATION:
'        - Iterates through all keys in the override Dictionary.
'        - For each key:
'            - Replaces the existing value in the merged Dictionary, or
'            - Adds the key/value pair if it does not already exist.
'        - Mirrors Graphviz's "last attribute wins" semantics.
'
'   3. PIPELINE INTEGRATION:
'        - Returns the merged Dictionary for use by attribute handlers such as
'          'HandleNodeAttribute', 'HandleEdgeAttribute', and
'          'HandleGraphAttribute'.
'        - Ensures consistent override behavior across all style layers.
'
' TECHNICAL NOTES:
'   - Assumes both input Dictionaries have already passed through
'     'NormalizeKeys' to enforce case-insensitive key handling.
'   - Does not perform any quoting or HTML-label sanitation; those steps are
'     handled later by 'FormatLabel' and 'RebuildStyleAttributeString'.
'   - DeepWiki Context: Implements the override-resolution rules described in
'     the "Style Layers" and "Attribute Normalization" documentation.
' ==========================================================================
Public Function MergeDictionaries( _
        ByVal d1 As Dictionary, _
        ByVal d2 As Dictionary) As Dictionary

    Dim result As New Dictionary
    Dim key As Variant

    If Not d1 Is Nothing Then
        For Each key In d1.keys
            If Not result.Exists(key) Then
                result.Add key, d1(key)
            Else
                result(key) = d1(key)
            End If
        Next key
    End If

    If Not d2 Is Nothing Then
        For Each key In d2.keys
            If Not result.Exists(key) Then
                result.Add key, d2(key)
            Else
                result(key) = d2(key)
            End If
        Next key
    End If

    Set MergeDictionaries = result
End Function

Public Function CloneDictionary(ByVal source As Dictionary) As Dictionary
    If source Is Nothing Then
        Set CloneDictionary = New Dictionary
    Else
        ' Reuse MergeDictionaries (second dict wins, first is base)
        Set CloneDictionary = MergeDictionaries(Nothing, source)
    End If
End Function

Public Function DictionaryValuesToCollection(ByVal dict As Dictionary) As Collection
    Dim c As New Collection
    Dim k As Variant

    For Each k In dict.keys
        c.Add dict(k)
    Next k

    Set DictionaryValuesToCollection = c
End Function

Public Sub DebugDictionary(dict As Dictionary, ByVal name As String)
    Dim key As Variant

    Debug.Print "---- " & name & " ----"
    Debug.Print "Count: "; dict.Count
    For Each key In dict.keys
        Debug.Print "Key: [" & key & "]  Value: [" & dict(key) & "]"
    Next key
    Debug.Print "----------------------"
    Debug.Print vbCrLf
End Sub





