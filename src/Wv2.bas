Attribute VB_Name = "Wv2"
''''''''''''''''''''''''''''''''''
' --- Wv2.bas  v1.0 (English fork) - the simple entry point ---
'
'   One small facade over UserForm1 / Wv2Browser / Wv2Pane, so that a macro
'   needs no knowledge of the start-up order:
'
'       Dim pg As Wv2Pane
'       Set pg = Wv2.OpenBrowser("https://example.com")
'       If pg Is Nothing Then Exit Sub            ' start-up or load failed
'       Debug.Print pg.WaitFor("h1", 10).InnerText
'       Wv2.CloseBrowser
'
'   Members:
'     OpenBrowser(url, withBars, timeoutSec)  show the form, start WebView2, load url
'     NewTab(url, timeoutSec)                 open another tab and wait for it to load
'     Navigate(url, timeoutSec, pane)         load url in a tab and wait for it to load
'     WaitLoaded(pane, timeoutSec)            wait for the current navigation to finish
'     Page                                    the active tab (Wv2Pane) or Nothing
'     ActiveBrowser                           the running Wv2Browser or Nothing
'     IsBrowserOpen                           True while the browser is running
'     Eval(js, timeoutSec, pane)              run a JS expression, result as JSON text
'     EvalText(js, timeoutSec, pane)          same, with a JS string result decoded
'     CloseBrowser                            close everything
'     Demo                                    30-second smoke test
'
'   Rules that still apply (see docs/USAGE.md):
'     - 64-bit Excel only.
'     - Never break into the debugger / single-step while a page is loading or
'       while one of these functions is waiting (Excel crashes).
'     - Do not call these from inside a WebView2 event handler; they return
'       False / Nothing / "" immediately there instead of deadlocking.
''''''''''''''''''''''''''''''''''
Option Explicit

Private Const DEFAULT_TIMEOUT As Single = 30


' ============================================================
' OpenBrowser
'   Shows UserForm1 (modeless), starts WebView2 and loads url.
'   withBars = True  : tab bar + address bar + content (a small browser)
'   withBars = False : content only (for automation, or your own HTML GUI)
'   Returns the tab's Wv2Pane once the page has loaded, or Nothing when
'   start-up failed or the page did not load within timeoutSec (the browser
'   then stays open; Wv2.Page still returns the tab).
'   If the browser is already running, url is opened in a new tab.
' ============================================================
Public Function OpenBrowser(Optional ByVal url As String = "about:blank", _
                            Optional ByVal withBars As Boolean = True, _
                            Optional ByVal timeoutSec As Single = DEFAULT_TIMEOUT) As Wv2Pane
    Dim pane As Wv2Pane

    If IsBrowserOpen() Then
        Set OpenBrowser = NewTab(url, timeoutSec)
        Exit Function
    End If

    ' Show first, then log: the other order can crash Excel when an IME is
    ' used in the Immediate window afterwards (upstream spec fact 67).
    UserForm1.Show vbModeless
    If LenB(Wv2Log.LogPath) = 0 Then Wv2Log.LogStart

    If Not UserForm1.StartBrowser(FullUrl(url), withBars) Then
        Wv2Log.LogE "Wv2.OpenBrowser: start-up failed - see the lines above"
        Unload UserForm1
        Exit Function
    End If

    Set pane = UserForm1.GetActivePane
    If WaitLoaded(pane, timeoutSec) Then Set OpenBrowser = pane
End Function


' ============================================================
' CloseBrowser - closes all tabs and the form. Safe to call when not open.
' ============================================================
Public Sub CloseBrowser()
    If FormLoaded() Then Unload UserForm1
End Sub


' ============================================================
' IsBrowserOpen / ActiveBrowser / Page
'   None of these loads the form when it is not open.
' ============================================================
Public Function IsBrowserOpen() As Boolean
    If Not FormLoaded() Then Exit Function
    IsBrowserOpen = Not (UserForm1.CurrentBrowser Is Nothing)
End Function

Public Function ActiveBrowser() As Wv2Browser
    If FormLoaded() Then Set ActiveBrowser = UserForm1.CurrentBrowser
End Function

Public Function Page() As Wv2Pane
    Dim b As Wv2Browser
    Set b = ActiveBrowser()
    If Not b Is Nothing Then Set Page = b.ActivePane
End Function


' ============================================================
' NewTab - opens url in a new tab (which becomes the active tab) and waits
'   for it to load. Returns Nothing if the browser is not open, the tab could
'   not be created, or the page did not load in time.
' ============================================================
Public Function NewTab(ByVal url As String, _
                       Optional ByVal timeoutSec As Single = DEFAULT_TIMEOUT) As Wv2Pane
    Dim b As Wv2Browser
    Dim pane As Wv2Pane

    Set b = ActiveBrowser()
    If b Is Nothing Then
        Wv2Log.LogW "Wv2.NewTab: the browser is not open (call Wv2.OpenBrowser first)"
        Exit Function
    End If

    Set pane = b.AddTabWithUrl(FullUrl(url))
    If pane Is Nothing Then Exit Function
    If WaitLoaded(pane, timeoutSec) Then Set NewTab = pane
End Function


' ============================================================
' Navigate - loads url in pane (default: the active tab) and waits.
'   True = the page loaded. False = not open / bad URL / timeout / load error.
' ============================================================
Public Function Navigate(ByVal url As String, _
                         Optional ByVal timeoutSec As Single = DEFAULT_TIMEOUT, _
                         Optional ByVal pane As Wv2Pane = Nothing) As Boolean
    Dim hr As Long

    If pane Is Nothing Then Set pane = Page()
    If pane Is Nothing Then
        Wv2Log.LogW "Wv2.Navigate: no tab (call Wv2.OpenBrowser first)"
        Exit Function
    End If

    hr = pane.View_Navigate(FullUrl(url))
    If hr <> 0 Then
        Wv2Log.LogW "Wv2.Navigate: Navigate failed hr=&H" & Hex$(hr) & " (" & url & ")"
        Exit Function
    End If
    Navigate = WaitLoaded(pane, timeoutSec)
End Function


' ============================================================
' WaitLoaded - waits until the navigation in progress has completed.
'   True only if it completed AND succeeded (HTTP error pages still count as
'   success; DNS failure, refused connection, cancelled load do not).
'   "Loaded" is the browser's NavigationCompleted: a single-page app may still
'   be drawing. Follow with pane.WaitFor(selector) or pane.WaitSettled.
' ============================================================
Public Function WaitLoaded(Optional ByVal pane As Wv2Pane = Nothing, _
                           Optional ByVal timeoutSec As Single = DEFAULT_TIMEOUT) As Boolean
    Dim t0 As Single

    If pane Is Nothing Then Set pane = Page()
    If pane Is Nothing Then Exit Function
    If timeoutSec <= 0 Then timeoutSec = DEFAULT_TIMEOUT

    ' Inside a WebView2 callback no further event can arrive: waiting would
    ' only burn the whole timeout (same rule as EvalSync, spec fact 14).
    If pane.InCallbackDepth > 0 Then
        Wv2Log.LogW "Wv2.WaitLoaded: called from inside an event handler - not waiting"
        Exit Function
    End If

    t0 = Timer
    Do While Not pane.IsPageLoaded
        If Wv2Log.SecondsSince(t0) >= timeoutSec Then
            Wv2Log.LogW "Wv2.WaitLoaded: timeout after " & timeoutSec & " s (" & pane.View_GetSource() & ")"
            Exit Function
        End If
        DoEvents
        If Not FormLoaded() Then Exit Function      ' the user closed the form while we waited
    Loop

    WaitLoaded = pane.LastNavigationOk
    If Not WaitLoaded Then
        Wv2Log.LogW "Wv2.WaitLoaded: the page failed to load (" & pane.View_GetSource() & ")"
    End If
End Function


' ============================================================
' Eval / EvalText - run a JavaScript *expression* in a tab and return its value.
'   Eval     : the value as JSON text   ("abc" comes back with its quotes, 42 as 42)
'   EvalText : a string result decoded  (abc); other types unchanged
'   On failure both return "" - check Wv2.Page.LastEvalOk / LastEvalError.
'   Several statements: wrap them, e.g. "(function(){ ...; return x; })()".
' ============================================================
Public Function Eval(ByVal jsExpr As String, _
                     Optional ByVal timeoutSec As Single = 5, _
                     Optional ByVal pane As Wv2Pane = Nothing) As String
    If pane Is Nothing Then Set pane = Page()
    If pane Is Nothing Then Exit Function
    Eval = pane.EvalSync(jsExpr, timeoutSec)
End Function

Public Function EvalText(ByVal jsExpr As String, _
                         Optional ByVal timeoutSec As Single = 5, _
                         Optional ByVal pane As Wv2Pane = Nothing) As String
    EvalText = Wv2Json.JsonUnescape(Eval(jsExpr, timeoutSec, pane))
End Function


' ============================================================
' Demo - smoke test. Run it from the Immediate window:  Wv2.Demo
'   Opens example.com, reads the heading through the DOM API and through JS,
'   and prints PASS / FAIL. Leaves the browser open.
' ============================================================
Public Sub Demo()
    Dim pg As Wv2Pane
    Dim el As Wv2Element
    Dim heading As String
    Dim docTitle As String

    Set pg = OpenBrowser("https://example.com")
    If pg Is Nothing Then
        Debug.Print "Wv2.Demo: FAIL - the browser did not start or the page did not load."
        Debug.Print "  Log file: " & Wv2Log.LogPath
        Exit Sub
    End If

    Set el = pg.WaitFor("h1", 10)
    If Not el Is Nothing Then heading = el.InnerText
    docTitle = EvalText("document.title", 5, pg)

    Debug.Print "Wv2.Demo: h1 = [" & heading & "], title = [" & docTitle & "]"
    If Len(heading) > 0 And Len(docTitle) > 0 Then
        Debug.Print "Wv2.Demo: PASS"
    Else
        Debug.Print "Wv2.Demo: FAIL - see " & Wv2Log.LogPath
    End If
End Sub


' ------------------------------------------------------------
' Private helpers
' ------------------------------------------------------------

' True when UserForm1 is loaded. Asking UserForm1 directly would load it.
Private Function FormLoaded() As Boolean
    Dim f As Object
    For Each f In VBA.UserForms
        If TypeOf f Is UserForm1 Then
            FormLoaded = True
            Exit Function
        End If
    Next f
End Function

' WebView2's Navigate needs a full URL: "example.com" alone fails with
' E_INVALIDARG. Adds https:// when there is no scheme.
Private Function FullUrl(ByVal url As String) As String
    Dim u As String
    u = Trim$(url)
    If LenB(u) = 0 Then
        u = "about:blank"
    ElseIf Mid$(u, 2, 2) = ":" & Chr$(92) Then
        u = "file:///" & Replace(u, Chr$(92), "/")       ' a local path (drive letter + backslashes)
    ElseIf InStr(1, u, "://") = 0 And LCase$(Left$(u, 6)) <> "about:" And LCase$(Left$(u, 5)) <> "data:" Then
        u = "https://" & u
    End If
    FullUrl = u
End Function
