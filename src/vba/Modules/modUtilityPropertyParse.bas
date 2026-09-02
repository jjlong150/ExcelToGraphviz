Attribute VB_Name = "modUtilityPropertyParse"
' =============================================================================
' MODULE:    modUtilityPropertyParse
' LAYER:     Utility / Generic Property Parser
'
' ROLE:
'   Standalone, dependency-free parser for generic "key=value" property
'   strings (e.g. weight=200 domestic=true costperunit=25.45). Unlike
'   modUtilityGraphvizParse's ParseAttributeString - which is intentionally
'   HTML/DOT-aware and returns everything as text so it can be written back
'   out as valid Graphviz syntax - this module infers and preserves real
'   VBA data types (Boolean, Long, Double, String - including validated
'   ISO-8601 date/datetime text kept as String) so downstream JSON
'   emission can produce weight:200 instead of weight:"200".
'
'   This module has no dependency on Graphviz concepts and can be dropped
'   into any project as-is.
'
' RESPONSIBILITIES:
'   - ParsePropertyString: tokenize a property string into a Dictionary of
'     key -> typed Variant value.
'   - SerializePropertyString: the reverse direction - convert a Dictionary
'     of typed key/value pairs back into a property string. Round-trips
'     with ParsePropertyString (see that function's header for the one
'     unavoidable exception: explicit empty-string values).
'
' TYPE INFERENCE RULES (in order):
'   1. Quoted values ("...") are ALWAYS treated as strings, even if they
'      look like a number, boolean, or date. Quoting is an explicit
'      "keep this as text" override.
'   2. Unquoted "true" / "false" (any case) -> Boolean.
'   3. Unquoted ISO-8601 dates, "yyyy-mm-dd" or "yyyy-mm-dd hh:mm[:ss]"
'      (T or space separator) -> validated and returned as-is, as a
'      String (not a VBA Date). Two deliberate choices here:
'        - Only this unambiguous format is recognized, since anything
'          else (e.g. mm/dd/yyyy vs dd/mm/yyyy) is locale-dependent and
'          would silently produce wrong dates on machines with different
'          regional settings.
'        - The value is kept as a String rather than converted to a VBA
'          Date, because JSON serializers (e.g. VBA-JSON's ConvertToJson)
'          typically round Date values through a UTC conversion before
'          formatting them, which can silently shift a date-only value
'          to the previous calendar day in timezones ahead of UTC. JSON
'          has no native date type regardless, so every consumer expects
'          an ISO-8601 string - this just guarantees the text you typed
'          is the text that comes out.
'   4. Unquoted numeric text (integer, decimal, or scientific notation,
'      optionally signed) -> Long (if it fits in 32-bit signed range and
'      has no decimal point/exponent) or Double otherwise. Parsing uses
'      Val(), which is locale-independent and always expects "." as the
'      decimal separator - matching JSON's number format regardless of
'      the workbook's regional settings.
'   5. Anything else -> String, as-is.
'
' EMPTY VALUES:
'   A property with no value (e.g. "note=" or a bare "note") is OMITTED
'   from the returned dictionary entirely, rather than added as "" or
'   Null. Callers should treat a missing key as "not specified".
'
' DELIMITERS:
'   Pairs may be separated by space, comma, or semicolon (mixed freely),
'   matching the convention already used in the Graphviz attribute column.
'   Values may be double-quoted to protect embedded delimiters, e.g.:
'     label="Detroit, MI" weight=200 domestic=true opened=2019-03-14
'
'   A literal quote character inside a quoted value is written as two
'   consecutive quotes (CSV-style escaping), e.g.:
'     quip="She said ""hi"" back"
'   parses to the String  She said "hi" back . SerializePropertyString
'   produces this same doubled-quote form automatically when needed.
'
' VERSION NOTES:
'   Introduced as a general-purpose companion to modUtilityGraphvizParse,
'   for use wherever a properties column is parsed into typed JSON.
' =============================================================================

Option Explicit

' Shared detection patterns - used both when INFERRING a type on parse and
' when deciding whether a String value needs quoting on serialize, so the
' two directions of the round trip can never drift out of sync.
Private Const NUMBER_PATTERN As String = "^[+-]?(\d+(\.\d+)?|\.\d+)([eE][+-]?\d+)?$"
Private Const DATE_PATTERN As String = "^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?$"

' ==========================================================================
' FUNCTION: ParsePropertyString
'
' PURPOSE:
'   Converts a generic "key=value" property string into a Dictionary of
'   key/typed-value pairs, preserving Boolean, numeric, and Date types
'   instead of collapsing everything to String.
'
' PARAMETERS:
'   properties - the raw property string, e.g. weight=200 domestic=true
'
' RETURNS:
'   Dictionary(key As String, value As Variant). Empty-valued properties
'   are omitted. On a fully empty/whitespace input, an empty Dictionary
'   is returned (never Nothing).
' ==========================================================================
Public Function ParsePropertyString(ByVal properties As String) As Dictionary
    Dim pipedProperties As String
    pipedProperties = AddPipeDelimitersToPropertyString(properties)
    Set ParsePropertyString = ParsePipedPropertyString(pipedProperties)
End Function

' --------------------------------------------------------------------------
' Tokenizer: walks the raw string once, inserting "|" delimiters between
' distinct key=value pairs while respecting quoted values (so a delimiter
' character inside quotes doesn't split the pair). Deliberately simpler
' than the Graphviz tokenizer - no HTML-label awareness is needed here.
' --------------------------------------------------------------------------
Private Function AddPipeDelimitersToPropertyString(ByVal properties As String) As String
    Dim inQuotes As Boolean
    Dim equalsFound As Boolean
    Dim oneChar As String
    Dim nextChar As String
    Dim pipedProperties As String
    Dim i As Long
    Dim propLen As Long

    inQuotes = False
    equalsFound = False
    pipedProperties = vbNullString
    propLen = Len(properties)
    i = 1

    Do While i <= propLen
        oneChar = Mid$(properties, i, 1)

        If oneChar = "=" And Not inQuotes Then
            pipedProperties = pipedProperties & oneChar
            equalsFound = True
            i = i + 1

        ElseIf oneChar = """" Then
            If inQuotes Then
                ' A doubled quote ("") inside a quoted value is an escaped
                ' literal quote character (CSV-style), not the closing
                ' quote - keep both characters; ParsePipedPropertyString
                ' collapses the pair back to one quote once the wrapping
                ' quotes are stripped.
                nextChar = Mid$(properties, i + 1, 1)
                If nextChar = """" Then
                    pipedProperties = pipedProperties & """"""
                    i = i + 2
                Else
                    ' Real closing quote.
                    pipedProperties = pipedProperties & oneChar & "|"
                    inQuotes = False
                    equalsFound = False
                    i = i + 1
                End If
            ElseIf equalsFound Then
                ' Opening quote.
                pipedProperties = pipedProperties & oneChar
                inQuotes = True
                i = i + 1
            Else
                ' Quote outside a value position - preserve literally
                ' (shouldn't normally occur in well-formed input).
                pipedProperties = pipedProperties & oneChar
                i = i + 1
            End If

        ElseIf (oneChar = ";" Or oneChar = "," Or oneChar = " ") And Not inQuotes Then
            If equalsFound Then
                ' End of an unquoted value (quoted values are handled above
                ' and already emit their own "|" on the closing quote).
                equalsFound = False
                pipedProperties = pipedProperties & "|"
            Else
                pipedProperties = pipedProperties & oneChar
            End If
            i = i + 1

        Else
            pipedProperties = pipedProperties & oneChar
            i = i + 1
        End If
    Loop

    ' Trailing unquoted value with no terminating delimiter.
    If equalsFound And Not inQuotes Then
        pipedProperties = pipedProperties & "|"
    End If

    AddPipeDelimitersToPropertyString = pipedProperties
End Function

' --------------------------------------------------------------------------
' Splits the piped string into key/value pairs, strips surrounding quotes
' (remembering that they were present), infers a type for unquoted values,
' and builds the result Dictionary. Empty values are omitted.
' --------------------------------------------------------------------------
Private Function ParsePipedPropertyString(ByVal pipedProperties As String) As Dictionary
    Dim dictionaryObj As Dictionary
    Set dictionaryObj = New Dictionary

    Dim pairs() As String
    Dim keyValue() As String
    Dim i As Long

    pairs = split(pipedProperties, "|")

    For i = LBound(pairs) To UBound(pairs)
        Dim pair As String
        pair = Trim$(pairs(i))

        If InStr(1, pair, "=") > 0 Then
            keyValue = split(pair, "=", 2)

            Dim key As String
            Dim rawValue As String
            key = Trim$(keyValue(0))
            rawValue = Trim$(keyValue(1))

            If key <> vbNullString Then
                Dim wasQuoted As Boolean
                wasQuoted = False

                If Len(rawValue) >= 2 Then
                    If Left$(rawValue, 1) = """" And Right$(rawValue, 1) = """" Then
                        wasQuoted = True
                        rawValue = Mid$(rawValue, 2, Len(rawValue) - 2)
                        rawValue = replace(rawValue, """""", """")
                    End If
                End If

                ' Empty value (quoted "" or bare "key=" or bare "key") -> omit entirely.
                If rawValue <> vbNullString Then
                    Dim typedValue As Variant
                    typedValue = InferPropertyValue(rawValue, wasQuoted)

                    If dictionaryObj.Exists(key) Then
                        dictionaryObj.Remove key
                    End If
                    dictionaryObj.Add key, typedValue
                End If
            End If
        End If
    Next i

    Set ParsePipedPropertyString = dictionaryObj
End Function

' --------------------------------------------------------------------------
' Infers the VBA type for a single unquoted value. Quoted values always
' return as String unchanged. See module header for the full rule order.
' --------------------------------------------------------------------------
Private Function InferPropertyValue(ByVal rawValue As String, ByVal wasQuoted As Boolean) As Variant
    If wasQuoted Then
        InferPropertyValue = rawValue
        Exit Function
    End If

    Dim trimmed As String
    trimmed = Trim$(rawValue)

    ' --- Boolean ---
    If LCase$(trimmed) = "true" Then
        InferPropertyValue = True
        Exit Function
    ElseIf LCase$(trimmed) = "false" Then
        InferPropertyValue = False
        Exit Function
    End If

    ' --- ISO-8601 Date/DateTime ---
    ' NOTE: This deliberately returns the validated ISO-8601 text as a
    ' String, NOT a VBA Date. If a real Date/Time value were returned here,
    ' VBA-JSON's ConvertToJson runs Date values through ConvertToUtc() -
    ' using the system's local timezone offset - before formatting them as
    ' "yyyy-mm-ddTHH:mm:ss.000Z". For a date-only value like "2019-03-14"
    ' (parsed as local midnight), that UTC conversion silently rolls the
    ' calendar date back a day in any timezone ahead of UTC. Keeping the
    ' validated original text as a String sidesteps that entirely - JSON
    ' has no native date type anyway, so every consumer expects an
    ' ISO-8601 string, not a distinguishable "date" JSON value.
    Dim dateRegex As Object
    Set dateRegex = CreateObject("VBScript.RegExp")
    dateRegex.pattern = DATE_PATTERN
    If dateRegex.Test(trimmed) Then
        Dim m As Object
        Set m = dateRegex.Execute(trimmed)(0)
        Dim y As Integer, mo As Integer, d As Integer
        Dim hh As Integer, mm As Integer, ss As Integer

        y = CInt(m.SubMatches(0))
        mo = CInt(m.SubMatches(1))
        d = CInt(m.SubMatches(2))

        ' Guard against impossible calendar values (e.g. month 13, Feb 30)
        ' that the regex shape alone wouldn't catch; fall through to
        ' String/Number below if the date isn't real.
        If mo >= 1 And mo <= 12 And d >= 1 And d <= 31 Then
            On Error Resume Next
            Dim candidate As Date
            Err.Clear
            candidate = DateSerial(y, mo, d)
            If Err.number = 0 And Day(candidate) = d And Month(candidate) = mo Then
                On Error GoTo 0

                Dim timeIsValid As Boolean
                timeIsValid = True
                If Not IsEmpty(m.SubMatches(3)) Then
                    hh = CInt(m.SubMatches(3))
                    mm = CInt(m.SubMatches(4))
                    If Not IsEmpty(m.SubMatches(5)) Then
                        ss = CInt(m.SubMatches(5))
                    Else
                        ss = 0
                    End If
                    timeIsValid = (hh >= 0 And hh <= 23) And (mm >= 0 And mm <= 59) And (ss >= 0 And ss <= 59)
                End If

                If timeIsValid Then
                    InferPropertyValue = trimmed  ' validated ISO-8601 text, returned as String
                    Exit Function
                End If
            Else
                On Error GoTo 0
            End If
        End If
    End If

    ' --- Number (integer, decimal, or scientific notation; optionally signed) ---
    Dim numberRegex As Object
    Set numberRegex = CreateObject("VBScript.RegExp")
    numberRegex.pattern = NUMBER_PATTERN
    If numberRegex.Test(trimmed) Then
        ' Val() is locale-independent and always treats "." as the decimal
        ' separator, unlike CDbl/CLng which honor the workbook's regional
        ' settings - important since JSON numbers are always period-based.
        Dim numValue As Double
        numValue = val(trimmed)

        Dim isWholeNumberFormat As Boolean
        isWholeNumberFormat = (InStr(trimmed, ".") = 0) And _
                              (InStr(LCase$(trimmed), "e") = 0)

        If isWholeNumberFormat And numValue >= -2147483648# And numValue <= 2147483647# Then
            InferPropertyValue = CLng(numValue)
        Else
            InferPropertyValue = numValue
        End If
        Exit Function
    End If

    ' --- Fallback: String ---
    InferPropertyValue = trimmed
End Function

' ==========================================================================
' FUNCTION: SerializePropertyString
'
' PURPOSE:
'   The reverse of ParsePropertyString: converts a Dictionary of typed
'   key/value pairs back into a property string, e.g.:
'     weight=200 domestic=true costperunit=25.45 label="Detroit, MI"
'
' PARAMETERS:
'   props     - Dictionary(key As String, value As Variant), typically one
'               returned by ParsePropertyString (or hand-built/edited).
'   delimiter - separator written between pairs. Defaults to a single
'               space; pass "; " or "," if you prefer a different style.
'
' RETURNS:
'   The serialized property string. An empty Dictionary returns "".
'
' TYPE -> TEXT RULES:
'   - Boolean          -> unquoted true / false
'   - Long/Double/etc.  -> unquoted number, via a locale-independent
'                          formatter (see FormatNumberForSerialization)
'   - String            -> written bare UNLESS it would be misread as
'                          something else on re-parse (contains a
'                          delimiter/quote character, or its content looks
'                          like a boolean/number/ISO-date) - in which case
'                          it's wrapped in quotes, with any embedded quote
'                          character doubled ("") so it survives re-parsing.
'   - Date (defensive)  -> ParsePropertyString never produces a VBA Date
'                          (see module header), but if a caller stores one
'                          by hand it's formatted as ISO-8601 text using a
'                          fixed custom format string, which - unlike
'                          Format()'s locale-dependent named formats - is
'                          not affected by regional settings.
'   - Empty/Null/other  -> written as "" (an explicit empty string). Note
'                          that re-parsing "" hits the EMPTY VALUES rule
'                          and OMITS the key again, exactly mirroring
'                          ParsePropertyString's own behavior.
'
' ROUND-TRIP NOTE:
'   ParsePropertyString(SerializePropertyString(d)) reproduces d, with one
'   unavoidable exception: an explicit empty-string value is not
'   preserved, since empty values are omitted by design on parse (see
'   module header, EMPTY VALUES). Everything else - including embedded
'   quotes, commas, and values that merely look like booleans/numbers/
'   dates - round-trips exactly.
' ==========================================================================
Public Function SerializePropertyString(ByVal props As Dictionary, Optional ByVal delimiter As String = " ") As String
    Dim parts As String
    Dim key As Variant

    parts = vbNullString

    For Each key In props.keys
        If parts <> vbNullString Then
            parts = parts & delimiter
        End If
        parts = parts & CStr(key) & "=" & SerializePropertyValue(props(key))
    Next key

    SerializePropertyString = parts
End Function

' --------------------------------------------------------------------------
' Converts one typed Variant back to its property-string text form. See
' SerializePropertyString's header for the full rule set.
' --------------------------------------------------------------------------
Private Function SerializePropertyValue(ByVal value As Variant) As String
    Select Case VarType(value)
        Case vbBoolean
            SerializePropertyValue = IIf(value, "true", "false")

        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            SerializePropertyValue = FormatNumberForSerialization(value)

        Case vbDate
            ' Defensive only - see module header; ParsePropertyString
            ' itself never returns a VBA Date.
            If value = Int(value) Then
                SerializePropertyValue = Format$(value, "yyyy-mm-dd")
            Else
                SerializePropertyValue = Format$(value, "yyyy-mm-dd hh:mm:ss")
            End If

        Case vbString
            Dim textValue As String
            textValue = CStr(value)
            If PropertyValueNeedsQuoting(textValue) Then
                SerializePropertyValue = """" & replace(textValue, """", """""") & """"
            Else
                SerializePropertyValue = textValue
            End If

        Case Else
            ' vbEmpty, vbNull, or anything unexpected.
            SerializePropertyValue = """"""
    End Select
End Function

' --------------------------------------------------------------------------
' Locale-independent number-to-text conversion. Str$() - like Val() on the
' parsing side - always uses "." as the decimal separator regardless of
' regional settings, unlike CStr (which would emit "25,45" on a
' comma-decimal locale and silently break both JSON output and re-parsing).
' Str$() prefixes non-negative numbers with a leading space; Trim$ removes it.
' --------------------------------------------------------------------------
Private Function FormatNumberForSerialization(ByVal value As Variant) As String
    FormatNumberForSerialization = Trim$(Str$(value))
End Function

' --------------------------------------------------------------------------
' Decides whether a String value must be quoted on serialization to
' survive re-parsing intact: either because it contains a character the
' tokenizer treats as a delimiter/quote, or because its content alone
' would be reinterpreted as a Boolean/Number/Date instead of a String.
' Reuses the exact same patterns InferPropertyValue tests against, so the
' two directions of the round trip can't drift apart.
' --------------------------------------------------------------------------
Private Function PropertyValueNeedsQuoting(ByVal value As String) As Boolean
    If value = vbNullString Then
        PropertyValueNeedsQuoting = True
        Exit Function
    End If

    If InStr(value, " ") > 0 Or InStr(value, ",") > 0 Or _
       InStr(value, ";") > 0 Or InStr(value, """") > 0 Then
        PropertyValueNeedsQuoting = True
        Exit Function
    End If

    If LCase$(value) = "true" Or LCase$(value) = "false" Then
        PropertyValueNeedsQuoting = True
        Exit Function
    End If

    Dim numberRegex As Object
    Set numberRegex = CreateObject("VBScript.RegExp")
    numberRegex.pattern = NUMBER_PATTERN
    If numberRegex.Test(value) Then
        PropertyValueNeedsQuoting = True
        Exit Function
    End If

    Dim dateRegex As Object
    Set dateRegex = CreateObject("VBScript.RegExp")
    dateRegex.pattern = DATE_PATTERN
    If dateRegex.Test(value) Then
        PropertyValueNeedsQuoting = True
        Exit Function
    End If

    PropertyValueNeedsQuoting = False
End Function

' ==========================================================================
' Test subroutine - run with F5 / Immediate Window to verify behavior.
' ==========================================================================
Public Sub TestParsePropertyString()
    Dim testString As String
    testString = "weight=200 domestic=true costperunit=25.45 label=""Detroit, MI"" " & _
                 "opened=2019-03-14 lastupdated=""2019-03-14"" ratio=-0.5 code=1e3 note=""""" & _
                 " quip=""She said """"hi"""" back"""

    Dim result As Dictionary
    Set result = ParsePropertyString(testString)

    Debug.Print "--- Parsed ---"
    Dim key As Variant
    For Each key In result.keys
        Debug.Print key & " = " & CStr(result(key)) & "  [" & TypeName(result(key)) & "]"
    Next key

    Debug.Print
    Debug.Print "--- Round trip ---"
    Dim serialized As String
    serialized = SerializePropertyString(result)
    Debug.Print "Serialized: " & serialized

    Dim reparsed As Dictionary
    Set reparsed = ParsePropertyString(serialized)

    Debug.Print
    For Each key In reparsed.keys
        Debug.Print key & " = " & CStr(reparsed(key)) & "  [" & TypeName(reparsed(key)) & "]"
    Next key
End Sub


