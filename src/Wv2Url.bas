Attribute VB_Name = "Wv2Url"
''''''''''''''''''''''''''''''''''
' --- Wv2Url.bas  Stage 9.16 (cleanup: extracting the URL conversion utilities) ---
'
'   A standard module split out of Wv2Browser containing the pure data conversions of URL
'   handling, where the output is determined by the input alone. The counterpart of Wv2Json.bas
'   (JSON conversion): a "group of stateless conversion utilities". Created in the cleanup stage (same rhythm as the JSON extraction in 9.13).
'
'   **What it contains (mechanism = pure functions only)**
'     - IsAlphaOnly(s)          : whether a string is letters only (used to detect scheme names)
'     - UrlEncodeComponent(s)   : percent-encoding equivalent to encodeURIComponent
'     - Utf8Bytes(s)            : String(UTF-16LE) -> UTF-8 byte array (ADODB.Stream)
'
'   **What it does not contain (policy = state-dependent parts stay in Wv2Browser)**
'     NormalizeUrl / BuildSearchUrl depend on instance state, the search engine setting
'     (m_searchTemplate), i.e. they are "browser behavior", so they stay in Wv2Browser and
'     call the pure functions of this Wv2Url. Separation of policy and mechanism.
'
'   **How ADODB.Stream is used** Utf8Bytes uses ADODB.Stream, but this is plain data conversion
'     unrelated to WebView2 control (thunks, vtables, memory primitives) and is
'     already used for writing HTML. It does not violate the core idea of the project, "drive
'     WebView2 without external COM" (follows the decision made in 9.14).
''''''''''''''''''''''''''''''''''

Option Explicit


' ============================================================
' IsAlphaOnly (moved from Wv2Browser in 9.16, logic unchanged)
'
'   Returns whether a string consists only of letters (A-Z / a-z).
'   NormalizeUrl uses it to detect scheme names such as "about:"
'   (anything containing digits is False, so "example.com:8080" is not mistaken for a scheme).
' ============================================================
Public Function IsAlphaOnly(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    Dim i As Long
    Dim c As String
    For i = 1 To Len(s)
        c = LCase$(Mid$(s, i, 1))
        If c < "a" Or c > "z" Then Exit Function
    Next i
    IsAlphaOnly = True
End Function


' ============================================================
' UrlEncodeComponent (moved from Wv2Browser in 9.16, logic unchanged)
'
'   Percent-encodes a string per RFC3986
'   (equivalent to JS encodeURIComponent. Only unreserved characters pass through; the rest become %XX).
'
'   **Unreserved characters** A-Z a-z 0-9 and "-" "_" "." "~" are kept as they are.
'     Everything else is converted to a UTF-8 byte sequence and each byte becomes %XX.
'     (A space becomes %20. Like encodeURIComponent, it does not become "+")
'
'   **UTF-8 conversion is done by Utf8Bytes (via ADODB.Stream)** so that high characters such as
'     Japanese expand correctly into multiple bytes. Processing one character at a time with AscW
'     would break surrogate pairs (emoji etc.), so the whole string is converted to UTF-8 at once
'     and then escaped byte by byte.
' ============================================================
Public Function UrlEncodeComponent(ByVal s As String) As String
    If Len(s) = 0 Then Exit Function

    Dim bytes() As Byte
    bytes = Utf8Bytes(s)

    ' If Utf8Bytes is empty (normally cannot happen since Len(s)=0 is rejected above), return empty
    On Error Resume Next
    Dim lb As Long, ub As Long
    lb = LBound(bytes)
    ub = UBound(bytes)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Dim sb As String
    Dim i As Long
    Dim b As Integer
    Dim c As String
    For i = lb To ub
        b = bytes(i)            ' 0..255
        c = Chr$(b)             ' for the ASCII range check (only meaningful when b<128)
        If (b >= 65 And b <= 90) _
        Or (b >= 97 And b <= 122) _
        Or (b >= 48 And b <= 57) _
        Or b = 45 Or b = 95 Or b = 46 Or b = 126 Then
            ' unreserved: A-Z a-z 0-9 - _ . ~
            sb = sb & c
        Else
            sb = sb & "%" & Right$("0" & Hex$(b), 2)
        End If
    Next i

    UrlEncodeComponent = sb
End Function


' ============================================================
' Utf8Bytes (moved from Wv2Browser in 9.16, logic unchanged)
'
'   Converts a VBA String (UTF-16LE) to a UTF-8 byte array.
'   Writes to ADODB.Stream with Charset="utf-8" and reads back with Type=adTypeBinary.
'
'   **No leading BOM is added** A text write with Charset="utf-8" adds no BOM
'     (the bytes read back with adTypeBinary do not contain EF BB BF).
'     As insurance against future changes, a guard removes the first 3 bytes if they are a BOM.
' ============================================================
Public Function Utf8Bytes(ByVal s As String) As Byte()
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                 ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.WriteText s
    st.Position = 0
    st.Type = 1                 ' adTypeBinary
    ' Switching from text to binary resets Position to the start.
    Dim raw() As Byte
    raw = st.Read                ' get all bytes
    st.Close
    Set st = Nothing

    ' BOM insurance (normally none is added, but if present, drop EF BB BF)
    On Error Resume Next
    Dim lb As Long, ub As Long
    lb = LBound(raw)
    ub = UBound(raw)
    If Err.Number <> 0 Then
    ' When there are no bytes, e.g. empty string input. Return an empty array.
        On Error GoTo 0
        Utf8Bytes = raw
        Exit Function
    End If
    On Error GoTo 0

    If (ub - lb) >= 2 Then
        If raw(lb) = &HEF And raw(lb + 1) = &HBB And raw(lb + 2) = &HBF Then
            Dim out() As Byte
            ReDim out(0 To (ub - lb) - 3)
            Dim j As Long
            For j = lb + 3 To ub
                out(j - (lb + 3)) = raw(j)
            Next j
            Utf8Bytes = out
            Exit Function
        End If
    End If

    Utf8Bytes = raw
End Function

