Attribute VB_Name = "Wv2Json"
''''''''''''''''''''''''''''''''''
' --- Wv2Json.bas  Stage D-2 (shared decoding of JSON strings) ---
'
'   Additions in D-2 (issue 6, option b):
'     *Added JsonUnescapeAt / JsonUnescape / JsonPickStr*
'       JsonPickStr, which D-1 placed in Wv2Pane as Private, was moved into this module
'       and made Public. In D-2 two places, Wv2Pane (unpacking EvalSync results) and
'       Wv2Element (extracting property values), need the same decoding, so it is
'       gathered here as a stateless pure transformation (design principle 70).
'
'     JsonUnescapeAt(json, startPos, outVal) As Long
'       Scans the *contents* of a JSON string literal from startPos, resolves the escapes
'       and puts the result in outVal. The return value is the position of the closing quote (0 if not found).
'       startPos must point at the character *after* the opening quote.
'
'     JsonUnescape(raw) As String
'       Takes one whole JSON value and, if it is a string literal, strips the quotes and
'       resolves the escapes. If it is not a string (number / true / null / object, etc.)
'       *it passes it through unchanged*. Shaped so the EvalSync return value can be passed straight in.
'
'     JsonPickStr(json, key, outVal) As Boolean
'       Extracts the value of "key":"value" with the escapes resolved.
'
'   *When to use JsonGetStr vs. these*
'     JsonGetStr is a light implementation that *does not resolve escapes*, meant for pulling out
'     values that "cannot contain escapes", such as the cmd / index sent by our own JS
'     (design principle 64). Strings that come from pages (innerText / innerHTML / attribute values)
'     always contain escapes, so use JsonPickStr / JsonUnescape for those. JsonGetStr has
'     existing callers, so it has been *left completely unchanged*.
'
'   *Why decoding is mandatory (spec fact 30)*
'     In result JSON, < arrives escaped as <. When D-2 takes innerHTML,
'     every tag arrives in this form, so decoding \uXXXX is not insurance but a necessity.
'
'   *Surrogate pairs*
'     A pair such as U+1F600 (an emoji) comes back correctly if you call ChrW$ twice and concatenate
'     (a VBA String is a sequence of UTF-16 code units, so adding the two halves forms the pair).
'     *Be careful when simulating in Python* Python str concatenation does not recombine them, so unless you
'     hold code units as a sequence and decode all at once at the end, you get false NGs.
'
''''''''''''''''''''''''''''''''''
' --- Wv2Json.bas  Stage 9.13 (cleanup: shared JSON helpers) ---
'
'   *The minimal JSON helpers used for VBA<->JS communication*
'
'   In stages 9.11c-9.12d, Wv2TabBar and Wv2NavBar each ended up holding the same functions
'   as Private (JsonEscape / GetJsonStr / GetJsonNum / BoolToJson). They are extracted into a
'   shared module before a third UI appears (an agreed cleanup).
'
'   *Why not put them in Wv2Thunks.bas*
'     Wv2Thunks.bas is a low-level-layer module of "machine-code thunks, hand-built vtables and SAFEARRAY
'     memory primitives" and is already over 2600 lines. JSON string handling is clearly a
'     separate concern, so it gets its own module instead of mixing responsibilities.
'
'   *Design principle 64* We only talk to the JS in our own HTML, so we do not build a general JSON parser.
'     The sender is limited to JS we wrote ourselves, so the input is fully trusted and there is no need to
'     handle the general form with nesting, arrays and escapes. A minimal function that extracts the needed keys is enough.
'     (VBA has no standard JSON parser, and external COM is not usable under this project's policy)
'
'   *Naming rule* These are Public functions of a standard module, so to avoid name
'     collisions with other modules we unify on the Json prefix:
'       JsonEscape  ... safely embed a string in a JSON string literal
'       JsonGetStr  ... extract the value of "key":"value"
'       JsonGetNum  ... extract the 123 of "key":123
'       JsonBool    ... turn a Boolean into a true/false literal
'
'   Correspondence with the old names (renamed in 9.13):
'       JsonEscape   <- JsonEscape  (unchanged)
'       JsonGetStr   <- GetJsonStr
'       JsonGetNum   <- GetJsonNum
'       JsonBool     <- BoolToJson
''''''''''''''''''''''''''''''''''

Option Explicit


' ============================================================
' JsonEscape ? safely embed a string in a JSON string literal
'
'   Escapes " and \ and control characters (0x00-0x1F).
'   The minimum needed to embed tab titles and URLs in JSON.
'
'   *Fixed in 9.13: high-range characters now pass through (spec fact 16)*
'     VBA's AscW returns characters at 0x8000 and above as a signed Integer.
'     Example: the kanji U+9032 ( = 36914) -> AscW returns -28622.
'     The old implementation wrote the control-character test as Case Is < 32, so high-range characters such as Japanese
'     fell into that branch as negative values and were escaped in \uXXXX form.
'     *The result was valid JSON and came back as the correct character on the JS side, so there was no real harm*, but
'     it was an unintended path, made the log hard to read and needlessly increased the bytes sent.
'     -> At the time of consolidation we changed it to ch >= 0 And ch < 32, and high-range characters pass through.
'       Behavior (how it looks on the JS side) is unchanged.
' ============================================================
Public Function JsonEscape(ByVal s As String) As String
    Dim out As String
    Dim i As Long
    Dim ch As Long
    Dim c As String

    out = ""
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        ch = AscW(c)
        Select Case ch
            Case 34            ' " (double quote)
                out = out & "\"""
            Case 92            ' \ (backslash)
                out = out & "\\"
            Case 8             ' backspace
                out = out & "\b"
            Case 9             ' tab
                out = out & "\t"
            Case 10            ' LF
                out = out & "\n"
            Case 12            ' FF
                out = out & "\f"
            Case 13            ' CR
                out = out & "\r"
            Case Else
                ' *9.13 fix* Turn only control characters into \uXXXX.
                '   The ch >= 0 condition is needed (because AscW returns high-range characters as negative values; spec fact 16).
                '   Without it, Japanese and the like get mangled into \uXXXX.
                If ch >= 0 And ch < 32 Then
                    out = out & "\u" & Right$("0000" & Hex$(ch), 4)
                Else
                    out = out & c
                End If
        End Select
    Next i

    JsonEscape = out
End Function


' ============================================================
' JsonGetStr ? extract the value of "key":"value"
'
'   Example: JsonGetStr("{""cmd"":""activate"",""index"":3}", "cmd") -> "activate"
'   Returns an empty string if not found.
'
'   *Escapes are not resolved* The sender is our own JS, and cmd and url
'     practically never contain characters that need escaping. Extend this if exact decoding becomes necessary (design principle 64).
' ============================================================
Public Function JsonGetStr(ByVal json As String, ByVal key As String) As String
    Dim pat As String
    pat = """" & key & """:"""
    ' the contents of pat are  "key":"  = 7 characters + key length

    Dim p As Long
    p = InStr(json, pat)
    If p = 0 Then
        JsonGetStr = ""
        Exit Function
    End If

    Dim st As Long
    st = p + Len(pat)                ' start of the value

    Dim en As Long
    en = InStr(st, json, """")       ' closing quote
    If en = 0 Then
        JsonGetStr = ""
        Exit Function
    End If

    JsonGetStr = Mid$(json, st, en - st)
End Function


' ============================================================
' JsonGetNum ? extract the 123 of "key":123
'
'   Example: JsonGetNum("{""cmd"":""close"",""index"":2}", "index") -> 2
'   Returns 0 if not found.
'
'   *Negative numbers and decimals are not handled* Currently it is only used for index (a 1-based positive integer).
'     Extend it when needed (design principle 64).
' ============================================================
Public Function JsonGetNum(ByVal json As String, ByVal key As String) As Long
    Dim pat As String
    pat = """" & key & """:"
    ' the contents of pat are  "key":  = 4 characters + key length

    Dim p As Long
    p = InStr(json, pat)
    If p = 0 Then
        JsonGetNum = 0
        Exit Function
    End If

    Dim st As Long
    st = p + Len(pat)

    ' skip whitespace (JSON.stringify does not emit it, but just in case)
    Do While st <= Len(json)
        If Mid$(json, st, 1) <> " " Then Exit Do
        st = st + 1
    Loop

    ' read while digits continue
    Dim buf As String
    Dim c As String
    Do While st <= Len(json)
        c = Mid$(json, st, 1)
        If c < "0" Or c > "9" Then Exit Do
        buf = buf & c
        st = st + 1
    Loop

    If Len(buf) = 0 Then
        JsonGetNum = 0
    ElseIf Len(buf) > 10 Then
        JsonGetNum = 2147483647   ' English fork: clamp (VBA does not short-circuit, so test length first)
    ElseIf CDbl(buf) > 2147483647# Then
        JsonGetNum = 2147483647   ' English fork: clamp (CLng raised overflow error 6 on page-supplied numbers)
    Else
        JsonGetNum = CLng(buf)
    End If
End Function


' ============================================================
' JsonBool ? turn a Boolean into the JSON true/false literal
'
'   *Why CStr(True) is not used* CStr(True) returns "True" (capitalized) and
'     may also depend on the locale. JSON requires lowercase true/false, so
'     we convert explicitly.
' ============================================================
Public Function JsonBool(ByVal b As Boolean) As String
    If b Then
        JsonBool = "true"
    Else
        JsonBool = "false"
    End If
End Function


' ============================================================
' JsonUnescapeAt (D-2) - read the contents of a JSON string literal while decoding
'
'   Arguments:
'     json     : the target JSON string (whole)
'     startPos : scan start position. Must point at *the character after the opening quote* (1-based)
'     outVal   : the extracted value (escapes resolved)
'   Return value:
'     Position of the closing quote. 0 if not closed (outVal is then an empty string)
'
'   Supported escapes: quote / backslash / slash / b / f / n / r / t / uXXXX
'   (each is one character after the backslash; only uXXXX is followed by 4 hex digits)
'   For any other \x, x is taken as is (invalid as JSON, but we do not fail).
' ============================================================
Public Function JsonUnescapeAt(ByVal json As String, _
                               ByVal startPos As Long, _
                               ByRef outVal As String) As Long
    Dim i As Long
    Dim buf As String
    Dim c As String
    Dim d As String

    outVal = ""
    JsonUnescapeAt = 0

    If startPos < 1 Then Exit Function

    i = startPos
    Do While i <= Len(json)
        c = Mid$(json, i, 1)

        If c = "\" Then
            d = Mid$(json, i + 1, 1)
            Select Case d
                Case """"
                    buf = buf & """"
                Case "\"
                    buf = buf & "\"
                Case "/"
                    buf = buf & "/"
                Case "b"
                    buf = buf & Chr$(8)
                Case "f"
                    buf = buf & Chr$(12)
                Case "n"
                    buf = buf & vbLf
                Case "r"
                    buf = buf & vbCr
                Case "t"
                    buf = buf & vbTab
                Case "u"
                    ' *Values beyond 4 digits also come back correctly* CLng("&HD83D") becomes -10179 as a 16-bit
                    '   signed value, but ChrW$ interprets a negative value as 65536 + n,
                    '   so you get the character 55357 (= D83D).
                    '   Concatenating the two surrogate halves forms a valid pair.
                    buf = buf & ChrW$(CLng("&H" & Mid$(json, i + 2, 4)))
                    i = i + 4
                Case Else
                    buf = buf & d
            End Select
            i = i + 2

        ElseIf c = """" Then
            outVal = buf
            JsonUnescapeAt = i
            Exit Function

        Else
            buf = buf & c
            i = i + 1
        End If
    Loop
End Function


' ============================================================
' JsonUnescape (D-2) - turn one JSON value back into a plain string
'
'   If it is a string literal (starts with a quote), strip the quotes and resolve the escapes.
'   Anything else (number / true / false / null / undefined / object / array) is
'   *returned as is*. The EvalSync return value can be passed straight in.
'
'   Broken input with no closing quote is returned as the original text (not swallowed,
'   so the caller can notice the anomaly).
' ============================================================
Public Function JsonUnescape(ByVal raw As String) As String
    Dim v As String
    Dim en As Long

    If Len(raw) = 0 Then
        JsonUnescape = ""
        Exit Function
    End If

    If Left$(raw, 1) <> """" Then
        JsonUnescape = raw
        Exit Function
    End If

    en = JsonUnescapeAt(raw, 2, v)
    If en > 0 Then
        JsonUnescape = v
    Else
        JsonUnescape = raw
    End If
End Function


' ============================================================
' JsonPickStr (D-2; moved here from Wv2Pane, where D-1 placed it)
'   Extracts the value of "key":"value" with the escapes resolved.
'
'   Arguments:
'     json   : the target JSON string
'     key    : the key to extract (its value must be a string)
'     outVal : the extracted value (escapes resolved)
'   Return value:
'     True  = the key was found and reading reached the closing quote
'     False = no such key, or not closed
' ============================================================
Public Function JsonPickStr(ByVal json As String, _
                            ByVal key As String, _
                            ByRef outVal As String) As Boolean
    Dim pat As String
    Dim p As Long
    Dim en As Long

    outVal = ""
    JsonPickStr = False

    pat = """" & key & """:"""
    p = InStr(1, json, pat)
    If p = 0 Then Exit Function

    en = JsonUnescapeAt(json, p + Len(pat), outVal)
    JsonPickStr = (en > 0)
    If Not JsonPickStr Then outVal = ""
End Function

