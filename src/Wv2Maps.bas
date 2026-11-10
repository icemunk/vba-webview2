Attribute VB_Name = "Wv2Maps"
''''''''''''''''''''''''''''''''''
' --- Wv2Maps.bas  stage K-4d (stamp the time at each step of MapsOpen) ---
'
'   *Not one line of logic was changed. Only log lines were added.*
'
'   Why it is needed: MapsOpen measurably takes 17 seconds, yet when the Wv2Tests side reproduces its contents
'   step by step, it *finishes in 3.9 seconds*. We could only reproduce 1/4.
'   *We do not know where the 13 seconds are spent*, so we measure the real thing directly.
'
'   This is the common entrance of 2 phenomena:
'     Phenomenon 1: MapsOpen is leisurely waiting (17 seconds)
'     Phenomenon 2: after MapsOpen, closing the form is slow (5 seconds)
'   *Phenomenon 2 never appeared in the Test_K4_Step reproduction*, so
'   the difference from the real thing is highly likely to lead directly to the cause.
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Maps.bas  stage D-7d (practical use: address -> coordinates / progress and abort) ---
'
'   *The first example of folding the parts built on the D axis into a form usable in business as is.*
'   D-4d demonstrated that Google Maps address search works on the real site,
'   so that procedure was made into a single function.
'
'   # What D-7 added (progress and abort)
'     MapsCancel    ... request an abort / read the current request (Property Get / Let)
'     MapsCanceled  ... whether the previous call ended by abort (read-only)
'     MapsCountRows ... count the rows to process first (an entry that gives the denominator)
'
'     *Pressing Esc during a batch run stops it*
'     Progress is shown in Application.StatusBar.
'
'   # *How to restore the status bar (measured in D-7d)*
'     *Application.StatusBar = False does not restore it.* The string "FALSE" stays
'     displayed (FALSE also appears on screen). Measured:
'
'       False        -> [FALSE] type=String    <- stringified as a value
'       Empty        -> [False] type=Boolean   <- *this is the right answer*
'       empty string "" -> [False] type=Boolean
'       vbNullString -> [False] type=Boolean
'       CVar(False)  -> [FALSE] type=String
'
'     The commonly said "putting in False returns control to Excel" is *the opposite in this environment*.
'     Test_D7_StatusBar can reproduce the same measurement at any time.
'
'   # *How to catch Esc (remade twice: D-7 -> D-7b -> D-7c)*
'
'     *The premise was wrong both times on the real machine. The history is kept.*
'
'     D-7  : Looked only with GetAsyncKeyState -> *it never fired*.
'            The VBA runtime's Esc arrives regardless of focus, so before we could read it
'            execution had already stopped at "Code execution has been interrupted".
'     D-7b : Set EnableCancelKey = xlErrorHandler and received Esc as error 18
'            -> *it did not reach the handler*. Error 18 mostly occurs
'            *inside WebView2 COM callbacks (View_On*)*, where the caller is not
'            VBA (via a thunk), so it does not trace back to On Error GoTo.
'            The Wv2Pane guard re-raised it, and it became an unhandled dialog at the COM boundary.
'     D-7c : *EnableCancelKey = xlDisabled* so that error 18 itself is not raised.
'            Execution does not stop, so the next time MapsPump runs
'            *GetAsyncKeyState picks it up*. The path becomes a single one.
'
'     *Why we want to avoid a break*: because it is exactly the crash window of spec fact 20.
'     A break while WebView2 events are running can become an AV.
'     It is a matter of safety, not only of ease of use.
'
'     *What is lost*: during a batch run, Excel's Esc / Ctrl+Break do not work.
'     However every wait has a timeout (EvalSync / WaitFor / WaitSettled), and
'     MapsPump always runs every 0.2-0.3 seconds, so it does not freeze forever.
'
'     *If it is not restored, Esc stops working in the whole of Excel*, so it is restored at every exit.
'     For that, MapsOpen / MapsGeocode are made in the form "wrapper + Core"
'     (Core has many exits, so the wrapper takes the cleanup in one place).
'
'   # Usage
'     Dim p As Wv2Pane, lat As Double, lng As Double, nm As String
'     Set p = Wv2Maps.MapsOpen(UserForm1.CurrentBrowser)
'     If Wv2Maps.MapsGeocode(p, "1-9-1 Marunouchi, Chiyoda-ku, Tokyo", lat, lng, nm) Then
'         Debug.Print lat & "," & lng & " " & nm
'     Else
'         Debug.Print "Failed: " & Wv2Maps.MapsLastError
'     End If
'
'     *The same Pane can be reused*, so addresses can be processed one after another (batch).
'     This is faster than reopening a tab.
'
'   # Why Maps-specific knowledge is isolated here
'     Both the selectors and the URL shape *change without notice at Google's convenience* (design principle 75).
'     If mixed into the product core (Wv2Pane / Wv2Element), it becomes unclear where to fix
'     when it breaks. **Things that change are confined to one place.**
'
'   # *What is used to judge completion*
'     `ArmUrlSignal` is not used. From the 2nd item on we are already on /place/, so
'     it hits the moment we arm (a trap documented in D-4c). Instead
'     *remember the URL before the search and wait until it becomes a different URL containing /place/ and !3d*.
'     "Done when the URL changes" is too early (at first it only changes to /search/,
'     and the coordinates are still those of the previous place). See the comment of MapsWaitPlace for details.
'     Getting the URL is via COM (View_GetSource), so it is light without EvalSync.
'     After the URL changes, WaitSettled is applied to wait for the DOM to settle.
'
'   # Where the coordinates are taken from
'     The !3d<latitude>!4d<longitude> in the URL is the first candidate (*the coordinates of the place itself*).
'     If absent, fall back to /@<latitude>,<longitude>,<zoom> (the center of the map).
'     *Use Val*: CDbl is locale-dependent and breaks in environments where the decimal point is ",".
'     Val always reads "." as the decimal point and stops at a non-digit character.
'
'   # Return value convention (following design principle 93)
'     True                       ... arrived at /place/ = settled to 1 result
'     False + MapsLastError      ... reason:
'       no-pane / no-searchbox / no-search-trigger ... the structure changed (Test_D4_Dom needed)
'       timeout-url              ... the URL did not change within the time
'       not-found                ... not found (*the map does not move even after searching*).
'                                  *No coordinates are returned in this case* since they are leftovers from the previous search
'       ambiguous                ... *multiple candidates* (arrived at /search/ but the map moved).
'                                  *With pickFirstIfAmbiguous:=True, it enters the first of the candidates
'                                  and settles on it (MapsLastPicked = True)*
'                                  In this case outLat / outLng hold the center of the map
'       no-coords                ... could not read the coordinates from the URL
'       canceled                 ... *aborted by Esc or MapsCancel* (D-7)
''''''''''''''''''''''''''''''''''
Option Explicit


' --- Reason for the previous failure ---
Private m_lastError As String

' --- D-6b: whether the previous result was *picked from the candidate list* ---
'   *Do not line up "settled" and "picked from candidates" with the same face* (the same line of thought as design principle 111).
'   If things of different precision get mixed in as the same ok, they can no longer be told apart afterwards.
Private m_lastPicked As Boolean

' --- D-7: abort (Esc) and progress ---
'   *Choose a means that does not depend on focus* Application.EnableCancelKey (Esc /
'   Ctrl+Break) does not arrive unless Excel has focus. During a batch the focus is on the
'   WebView2 side, so we use GetAsyncKeyState, which reads the key state directly.
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" ( _
    ByVal vKey As Long) As Integer

Private Const MAPS_VK_ESCAPE As Long = &H1B&

' Whether an abort was requested. *It can also be raised from outside* (Wv2Maps.MapsCancel = True)
Private m_cancel As Boolean

' Whether the previous call ended by abort
Private m_canceled As Boolean

' Whether a batch (MapsGeocodeSheet) is running.
'   *Discarding the abort state at the entrance is done only by "the outermost"* A single MapsGeocode also
'   discards it at its entrance, so without this the request vanishes after each row and the batch does not stop.
Private m_inBatch As Boolean

' D-7d: the number of times MapsCheckCancel was called. *Count what can be counted* (design principle 112)
'   When Esc cannot be picked up, whether the cause is "it is not being called" or "it is called but cannot see"
'   cannot be separated by guessing. Measure how many times we can poll in one batch.
Private m_checkCount As Long

' D-7b: Application.EnableCancelKey before arming. *Always restore it*
'   xlDisabled is 0 so it cannot be told apart from "not saved",
'   but when it is 0 we return it to xlInterrupt (the default) - restoring the fortress is safer.
Private m_savedCancelKey As Long

' Candidates for links in the candidate list (D-6b). *Pick by href first* because class changes easily.
Private Const MAPS_CAND_SELECTORS As String = _
    "a[href*='/maps/place/']" & vbLf & "div[role='feed'] a[href*='/place/']" & vbLf & "a.hfpxzc"

' Candidates for the search box. *id may change to an auto-generated one (spec fact 59)*, so name / structure first.
'   *The delimiter is vbLf* Chr$() cannot be used in Const (it is not a constant expression).
'   A newline does not appear in a CSS selector, so there is no collision.
Private Const MAPS_BOX_SELECTORS As String = _
    "input[name='q']" & vbLf & "#searchboxinput" & vbLf & "form input[type='text']"

' Candidates for running the search. aria-label changes with the UI language, so English is lined up too.
Private Const MAPS_GO_SELECTORS As String = _
    ("button[aria-label='" & ChrW(&H691C&) & ChrW(&H7D22&) & "']") & vbLf & "#searchbox-searchbutton" & vbLf & "button[aria-label='Search']"


' ============================================================
' MapsOpen - Open a Google Maps tab and wait until it is operable
'
'   Return value: a usable Wv2Pane. Nothing on failure (the reason is in MapsLastError)
'
'   *Right after opening the tab EvalSync fails with no-view*, so
'   first wait until JS passes (lesson of D-4d).
' ============================================================
Public Function MapsOpen(ByVal b As Wv2Browser, _
                         Optional ByVal timeoutSec As Single = 30) As Wv2Pane
    Dim p As Wv2Pane
    Dim solo As Boolean
    Dim errNo As Long
    Dim errDesc As String
    Dim errSrc As String

    ' --- D-7b: arm only when called standalone (the batch side has already armed) ---
    solo = Not m_inBatch
    If solo Then
        MapsResetCancel
        MapsArmCancelKey
        On Error GoTo Cleanup
    End If

    Set p = MapsOpenCore(b, timeoutSec)

Cleanup:
    errNo = Err.Number
    errDesc = Err.Description
    errSrc = Err.source
    On Error GoTo 0

    If errNo = 18 Then
        ' D-7c: xlDisabled was set so this normally does not come. Left for the window before/after arming.
        Set p = Nothing
        m_cancel = True
        m_canceled = True
        m_lastError = "canceled"
        Wv2Log.LogW "Wv2Maps.MapsOpen: *aborted by Esc*"
        errNo = 0
    End If

    If solo Then MapsDisarmCancelKey
    Set MapsOpen = p

    ' *Other errors are not swallowed* Only cleanup is done and it is raised as is
    If errNo <> 0 Then Err.Raise errNo, errSrc, errDesc
End Function


' ============================================================
' MapsOpenCore (D-7b, Private) - the body of MapsOpen
'   *There are many exits, so the cleanup of EnableCancelKey is left to the wrapper*
' ============================================================
Private Function MapsOpenCore(ByVal b As Wv2Browser, _
                              ByVal timeoutSec As Single) As Wv2Pane
    Dim p As Wv2Pane
    Dim el As Wv2Element
    Dim t0 As Single
    Dim tStep As Single      ' K-4d: for measuring each step
    Dim evalTries As Long    ' K-4d: how many times EvalSync was fired

    m_lastError = ""

    If b Is Nothing Then
        m_lastError = "no-browser"
        Exit Function
    End If

    tStep = Timer
    Set p = b.AddTabWithUrl("https://www.google.com/maps")
    Wv2Log.LogD "[MapsOpen] AddTabWithUrl ... " & _
                Format$(MapsElapsed(tStep), "0.00") & " s"   ' K-4d
    If p Is Nothing Then
        m_lastError = "no-pane"
        Exit Function
    End If

    ' --- Wait until JS passes ---
    t0 = Timer
    tStep = Timer
    Do
        evalTries = evalTries + 1   ' K-4d
        p.EvalSync "1", 3
        If p.LastEvalOk Then Exit Do
        If MapsCheckCancel() Then           ' D-7
            m_lastError = "canceled"
            m_canceled = True
            Exit Function
        End If
        If MapsElapsed(t0) >= timeoutSec Then
            m_lastError = "pane-not-ready"
            Exit Function
        End If
        MapsPump 0.3
    Loop
    Wv2Log.LogD "[MapsOpen] until JS passes ... " & _
                Format$(MapsElapsed(tStep), "0.00") & " s (EvalSync " & _
                evalTries & " times)"   ' K-4d

    ' --- Wait until the search box appears ---
    tStep = Timer
    Set el = MapsFindFirst(p, MAPS_BOX_SELECTORS, timeoutSec)
    Wv2Log.LogD "[MapsOpen] waiting for search box ... " & _
                Format$(MapsElapsed(tStep), "0.00") & " s " & _
                IIf(el Is Nothing, "(miss)", "(hit)") & _
                " registry " & p.ElementCount & " items"   ' K-4d
    If el Is Nothing Then
        m_lastError = "no-searchbox"
        Exit Function
    End If

    ' --- *Wait until the map has a position* ---
    '   The URL right after opening is https://www.google.com/maps and has no coordinates.
    '   MapsGeocode judges whether it was found by "whether the map moved after searching", so
    '   *without a baseline coordinate to compare against it cannot judge*. A /@ is attached in 1-2 seconds, so we wait.
    '   It is not fatal if it is not attached (in that case it is conservatively treated as not-found).
    t0 = Timer
    tStep = Timer
    Do
        If InStr(1, p.View_GetSource(), "/@") > 0 Then Exit Do
        If MapsCheckCancel() Then Exit Do   ' D-7 (not fatal, so just break out)
        If MapsElapsed(t0) >= 8 Then
            Wv2Log.LogW "Wv2Maps.MapsOpen: coordinates did not enter the URL (" & _
                        p.View_GetSource() & ")"
            Exit Do
        End If
        MapsPump 0.3
    Loop

    ' --- *Apply Maps-specific knowledge here* ---
    '   The search keeps streaming results with /search?tbm=map left open (spec fact 58).
    '   If it is not excluded from the quiescence check, we would wait the whole StaleInflightMs (default 10 seconds)
    '   for every single item. Excluding it by name settles in 1 second.
    '   *This is the gain of "confining things that change to one place"*
    Wv2Log.LogD "[MapsOpen] waiting for /@ ... " & _
                Format$(MapsElapsed(tStep), "0.00") & " s"   ' K-4d

    p.AddIgnoreNetwork "/search?tbm=map"

    Wv2Log.LogI "Wv2Maps.MapsOpen: ready (" & p.View_GetSource() & ")"
    Set MapsOpenCore = p
End Function


' ============================================================
' MapsGeocode - Search an address and get the coordinates
'
'   Arguments:
'     p           : the Pane returned by MapsOpen (can be reused)
'     addressText : the address to search
'     outLat / outLng : latitude / longitude
'     outName     : the normalized name returned by Maps (e.g. the postcode and the full Japanese address)
'     timeoutSec  : the upper limit for waiting for the URL to change
'
'   Return value: True = settled to 1 result / False = the reason is in MapsLastError
' ============================================================
Public Function MapsGeocode(ByVal p As Wv2Pane, _
                            ByVal addressText As String, _
                            ByRef outLat As Double, _
                            ByRef outLng As Double, _
                            ByRef outName As String, _
                            Optional ByVal timeoutSec As Single = 20, _
                            Optional ByVal pickFirstIfAmbiguous As Boolean = False) As Boolean
    Dim r As Boolean
    Dim solo As Boolean
    Dim errNo As Long
    Dim errDesc As String
    Dim errSrc As String

    ' --- D-7b: arm only when called standalone (the batch side has already armed) ---
    '   *Discarding the abort state at the entrance* is also only for the standalone case (design principle 106).
    '   If discarded inside a batch, the request vanishes after each row and it can no longer stop.
    solo = Not m_inBatch
    If solo Then
        MapsResetCancel
        MapsArmCancelKey
        On Error GoTo Cleanup
    End If

    r = MapsGeocodeCore(p, addressText, outLat, outLng, outName, _
                        timeoutSec, pickFirstIfAmbiguous)

Cleanup:
    errNo = Err.Number
    errDesc = Err.Description
    errSrc = Err.source
    On Error GoTo 0

    If errNo = 18 Then
        ' D-7c: xlDisabled was set so this normally does not come. Left for the window before/after arming.
        '   Values that cannot be judged are not written down (design principle 111).
        r = False
        outLat = 0
        outLng = 0
        outName = ""
        m_cancel = True
        m_canceled = True
        m_lastError = "canceled"
        Wv2Log.LogW "Wv2Maps.MapsGeocode: *aborted by Esc* (" & addressText & ")"
        errNo = 0
    End If

    If solo Then MapsDisarmCancelKey
    MapsGeocode = r

    ' *Other errors are not swallowed* Only cleanup is done and it is raised as is
    If errNo <> 0 Then Err.Raise errNo, errSrc, errDesc
End Function


' ============================================================
' MapsGeocodeCore (D-7b, Private) - the body of MapsGeocode
'   *There are many exits, so the cleanup of EnableCancelKey is left to the wrapper*
' ============================================================
Private Function MapsGeocodeCore(ByVal p As Wv2Pane, _
                                 ByVal addressText As String, _
                                 ByRef outLat As Double, _
                                 ByRef outLng As Double, _
                                 ByRef outName As String, _
                                 ByVal timeoutSec As Single, _
                                 ByVal pickFirstIfAmbiguous As Boolean) As Boolean
    Dim el As Wv2Element
    Dim prevUrl As String
    Dim newUrl As String
    Dim cands As Collection

    m_lastError = ""
    m_lastPicked = False
    outLat = 0
    outLng = 0
    outName = ""

    If p Is Nothing Then
        m_lastError = "no-pane"
        Exit Function
    End If

    prevUrl = p.View_GetSource()

    ' --- Write the address (D-3. *It goes through the native setter so it reaches the SPA*) ---
    Set el = MapsFindFirst(p, MAPS_BOX_SELECTORS, 10)
    If MapsCheckCancel() Then GoTo Canceled     ' D-7
    If el Is Nothing Then
        m_lastError = "no-searchbox"
        Exit Function
    End If

    el.value = addressText
    If Not el.LastOk Then
        m_lastError = "write-failed: " & el.LastError
        Exit Function
    End If

    ' --- Run the search ---
    Set el = MapsFindFirst(p, MAPS_GO_SELECTORS, 3)
    If el Is Nothing Then
        ' If there is no button, synthesize Enter
        p.EvalSync "(function(){var e=document.querySelector('input[name=""q""]');" & _
                   "if(!e){return 0;}e.focus();" & _
                   "e.dispatchEvent(new KeyboardEvent('keydown'," & _
                   "{key:'Enter',code:'Enter',keyCode:13,which:13,bubbles:true}));" & _
                   "return 1;})()"
        If Not p.LastEvalOk Then
            m_lastError = "no-search-trigger"
            Exit Function
        End If
    Else
        If Not el.Click() Then
            m_lastError = "click-failed: " & el.LastError
            Exit Function
        End If
    End If

    ' --- *Wait until the place is settled* (see the header for why arm is not used) ---
    If Not MapsWaitPlace(p, prevUrl, timeoutSec, newUrl) Then
        ' --- It did not settle ---
        '   *Directly count whether the candidate list is shown* (QuerySelectorAll of D-6)
        '     0 items      ... nothing was found. No coordinates are returned
        '     1 or more    ... multiple candidates. The center coordinates are meaningful
        '   *Previously we looked at "whether the map moved after searching", but when the same search term is
        '     searched consecutively the map does not move, and it became not-found even though candidates were shown*
        '     (found in the D-6b real-machine test). Stop indirect guessing and count directly.

        ' *Check abort before "not found"* (D-7)
        If MapsCheckCancel() Then GoTo Canceled

        If InStr(1, newUrl, "/@") = 0 Then
            m_lastError = "timeout-url"
            Exit Function
        End If

        Set cands = MapsFindCandidates(p)
        Wv2Log.LogI "Wv2Maps.MapsGeocode: candidate links " & cands.Count & " items"

        If cands.Count = 0 Then
            m_lastError = "not-found"
            Wv2Log.LogW "Wv2Maps.MapsGeocode: not found (" & addressText & _
                        ") no candidate list is shown either, so no coordinates are returned"
            Exit Function
        End If

        If pickFirstIfAmbiguous Then
            If MapsPickFirst(p, cands, timeoutSec, newUrl) Then
                m_lastPicked = True
                m_lastError = ""
                Wv2Log.LogI "Wv2Maps.MapsGeocode: *took the first candidate* (" & _
                            addressText & ")"
                GoTo Resolved
            End If
            Wv2Log.LogW "Wv2Maps.MapsGeocode: could not pick a candidate (" & addressText & ")"
            If MapsCheckCancel() Then GoTo Canceled     ' D-7
        End If

        MapsExtractLatLng newUrl, outLat, outLng
        outName = MapsTitleOf(p)
        m_lastError = "ambiguous"
        Wv2Log.LogW "Wv2Maps.MapsGeocode: multiple candidates (" & addressText & _
                    ") returning the center coordinates"
        Exit Function
    End If

' *Settled* From here on, the same processing is done for "it was settled from the start" and
' "it entered the first candidate".
Resolved:

    ' --- Wait for the DOM to settle (D-4. Lingering requests are cut off by default) ---
    p.WaitSettled 10
    Wv2Log.LogD "Wv2Maps.MapsGeocode: " & p.LastSettleInfo

    outName = MapsTitleOf(p)

    If Not MapsExtractLatLng(newUrl, outLat, outLng) Then
        m_lastError = "no-coords"
        Exit Function
    End If

    MapsGeocodeCore = True
    Exit Function

' --- D-7: *aborted* ---
'   *Values that cannot be judged are not written down* (design principle 111). Both coordinates and name are returned empty.
Canceled:
    outLat = 0
    outLng = 0
    outName = ""
    m_lastError = "canceled"
    m_canceled = True
    Wv2Log.LogW "Wv2Maps.MapsGeocodeCore: *aborted* (" & addressText & ")"
End Function


' ============================================================
' MapsTitleOf (Private) - Drop ' - Google Maps' from the title
'   *The normalized address is obtained* It can also be used for tidying notation variations.
' ============================================================
Private Function MapsTitleOf(ByVal p As Wv2Pane) As String
    Dim t As String

    t = p.View_GetDocumentTitle()
    If InStr(1, t, " - ") > 0 Then
        t = Left$(t, InStrRev(t, " - ") - 1)
    End If

    MapsTitleOf = t
End Function


' ============================================================
' MapsLastPicked (D-6b) - whether the previous result was picked from the candidate list
'
'   *If True, it is not "that address itself" but "the first of the candidates"*
'   Its precision differs from a settled one (True + MapsLastPicked = False), so when recording it
'   make it distinguishable. MapsGeocodeSheet writes ok(candidate 1) in the status column.
' ============================================================
Public Property Get MapsLastPicked() As Boolean
    MapsLastPicked = m_lastPicked
End Property


' ============================================================
' MapsFindCandidates (D-6b, Private) - grab the links of the candidate list
'
'   *An entry for directly counting "whether candidates are shown"*
'   0 items means "nothing was found", 1 or more means "multiple candidates".
'   The selectors are the candidate method (design principle 107). Those picked by href are placed first.
' ============================================================
Private Function MapsFindCandidates(ByVal p As Wv2Pane) As Collection
    Dim sels As Variant
    Dim i As Long
    Dim els As Collection

    Set MapsFindCandidates = New Collection
    sels = Split(MAPS_CAND_SELECTORS, vbLf)

    For i = LBound(sels) To UBound(sels)
        Set els = p.QuerySelectorAll(CStr(sels(i)), 20)
        If els.Count > 0 Then
            Wv2Log.LogD "Wv2Maps.MapsFindCandidates: " & els.Count & _
                        " items [" & sels(i) & "]"
            Set MapsFindCandidates = els
            Exit Function
        End If
    Next i
End Function


' ============================================================
' MapsPickFirst (D-6b, Private) - enter the first of the candidate list
'
'   *The QuerySelectorAll of D-6 pays off here* Grab the candidate links together,
'   click the first in document order, and wait until the place is settled.
'
'   Return value: True = settled (outUrl holds the URL after settling)
' ============================================================
Private Function MapsPickFirst(ByVal p As Wv2Pane, _
                               ByVal cands As Collection, _
                               ByVal timeoutSec As Single, _
                               ByRef outUrl As String) As Boolean
    Dim el As Wv2Element
    Dim prevUrl As String

    If cands Is Nothing Then Exit Function
    If cands.Count = 0 Then Exit Function

    prevUrl = p.View_GetSource()
    Set el = cands(1)   ' *The first in document order* (confirmed in D-6 that the order is preserved)

    If Not el.Click() Then
        Wv2Log.LogW "Wv2Maps.MapsPickFirst: cannot click err=" & el.LastError
        Exit Function
    End If

    MapsPickFirst = MapsWaitPlace(p, prevUrl, timeoutSec, outUrl)
End Function


' ============================================================
' MapsLastError - the reason for the previous failure
' ============================================================
Public Property Get MapsLastError() As String
    MapsLastError = m_lastError
End Property


' ============================================================
' MapsWaitPlace (Private) - wait *until the place is settled*
'
'   *"Done when the URL changes" is too early* (stepped on in the D-5 real-machine test)
'   When search is pressed, Maps first changes to /maps/search/<search term>/@<previous center>,
'   and then settles at /maps/place/<name>/@<new coordinates>/data=!3d<latitude>!4d<longitude>.
'   If grabbed at the first change, you get a one-item shift where *the name is new but the coordinates are still the previous ones*.
'   In fact when 3 searches were run in a row, the coordinates of the 2nd and 3rd were those of the one before.
'
'   Therefore what we wait for is *a URL containing both /place/ and !3d*.
'   !3d<latitude>!4d<longitude> is attached only after the place is settled.
'
'   Return value: True = settled (outUrl holds that URL)
'           False = timed out (outUrl holds the current URL. The caller judges whether there are multiple candidates)
'
'   *When the same address is searched twice in a row* the URL does not change so it times out,
'   but if it is already /place/ + !3d at that point, it is treated as settled (same as the previous result, so correct).
' ============================================================
Private Function MapsWaitPlace(ByVal p As Wv2Pane, _
                               ByVal prevUrl As String, _
                               ByVal timeoutSec As Single, _
                               ByRef outUrl As String) As Boolean
    Dim t0 As Single
    Dim u As String

    t0 = Timer
    Do
        u = p.View_GetSource()

        ' D-7: do not wait if aborted. Return False and let the caller judge.
        If MapsCheckCancel() Then
            outUrl = u
            Exit Function
        End If

        If u <> prevUrl And MapsIsPlaceUrl(u) Then
            outUrl = u
            MapsWaitPlace = True
            Exit Function
        End If

        If MapsElapsed(t0) >= timeoutSec Then
            outUrl = u
            ' If the same address was searched again we come here. If it is in the settled form, that is fine.
            MapsWaitPlace = MapsIsPlaceUrl(u)
            Exit Function
        End If

        MapsPump 0.2
    Loop
End Function

Private Function MapsIsPlaceUrl(ByVal u As String) As Boolean
    If InStr(1, u, "/place/") = 0 Then Exit Function
    If InStr(1, u, "!3d") = 0 Then Exit Function
    MapsIsPlaceUrl = True
End Function

' ============================================================
' MapsExtractLatLng (Private) - Pull the coordinates from the URL
'
'   First candidate: !3d<latitude>!4d<longitude>   *the coordinates of the place itself*
'   Second candidate: /@<latitude>,<longitude>      the center of the map
'
'   *Use Val* CDbl is locale-dependent. Val always reads "." as the decimal point and
'   stops at a non-digit character (such as ! or ,).
' ============================================================
Private Function MapsExtractLatLng(ByVal url As String, _
                                   ByRef outLat As Double, _
                                   ByRef outLng As Double) As Boolean
    Dim i As Long
    Dim j As Long

    i = InStr(1, url, "!3d")
    j = InStr(1, url, "!4d")
    If i > 0 And j > 0 Then
        outLat = Val(Mid$(url, i + 3))
        outLng = Val(Mid$(url, j + 3))
        If MapsLatLngOk(outLat, outLng) Then
            MapsExtractLatLng = True
            Exit Function
        End If
    End If

    i = InStr(1, url, "/@")
    If i > 0 Then
        outLat = Val(Mid$(url, i + 2))
        j = InStr(i, url, ",")
        If j > 0 Then outLng = Val(Mid$(url, j + 1))
        If MapsLatLngOk(outLat, outLng) Then
            MapsExtractLatLng = True
            Exit Function
        End If
    End If

    outLat = 0
    outLng = 0
End Function

Private Function MapsLatLngOk(ByVal la As Double, ByVal lo As Double) As Boolean
    If la = 0 And lo = 0 Then Exit Function
    If la < -90 Or la > 90 Then Exit Function
    If lo < -180 Or lo > 180 Then Exit Function
    MapsLatLngOk = True
End Function


' ============================================================
' MapsFindFirst (Private) - Try the candidate selectors in order
'   *Do not write real-site selectors by guessing (design principle 107)*
'   The one that hit is left in the log, so the next time it breaks we can tell what changed.
' ============================================================
Private Function MapsFindFirst(ByVal p As Wv2Pane, _
                               ByVal selectorList As String, _
                               ByVal timeoutSec As Single) As Wv2Element
    Dim cands As Variant
    Dim i As Long
    Dim el As Wv2Element

    cands = Split(selectorList, vbLf)

    For i = LBound(cands) To UBound(cands)
        If i = LBound(cands) Then
            Set el = p.WaitFor(CStr(cands(i)), timeoutSec)
        Else
            Set el = p.QuerySelector(CStr(cands(i)))
        End If

        If Not el Is Nothing Then
            Set MapsFindFirst = el
            Exit Function
        End If
    Next i

    Wv2Log.LogW "Wv2Maps.MapsFindFirst: none of the candidates hit [" & _
                Replace$(selectorList, vbLf, " / ") & "]"
End Function


' ============================================================
' MapsPump / MapsElapsed (Private)
'   *Timer returns to 0 at midnight (spec fact 53)*, so elapsed time is measured with correction.
' ============================================================
Private Sub MapsPump(ByVal waitSec As Single)
    Dim t0 As Single

    t0 = Timer
    Do
        DoEvents
        ' *D-7: abort detection is placed in this one spot* All waits pass through here,
        '   so it need not be added to each wait loop. The judgment after breaking out is up to the caller.
        If MapsCheckCancel() Then Exit Do
        If MapsElapsed(t0) >= waitSec Then Exit Do
    Loop
End Sub

Private Function MapsElapsed(ByVal sinceTimer As Single) As Single
    Dim d As Single

    d = Timer - sinceTimer
    If d < 0 Then d = d + 86400
    MapsElapsed = d
End Function

' ============================================================
' MapsCancel (D-7) - request an abort / read the current request
'
'   *It also stops with Esc, but an entry that can be raised from outside is left too* (when you want to stop it from the Immediate window,
'   or when a cancel button is added to the form in the future). It arrives because DoEvents is running.
'     Wv2Maps.MapsCancel = True
'
'   *Even if left raised, it does not kill the next run* It is discarded at the entrance of the next MapsOpen / MapsGeocode /
'   MapsGeocodeSheet (design principle 106: do not leave the previous state).
' ============================================================
Public Property Get MapsCancel() As Boolean
    MapsCancel = m_cancel
End Property

Public Property Let MapsCancel(ByVal newValue As Boolean)
    m_cancel = newValue
End Property


' ============================================================
' MapsCanceled (D-7) - whether the previous call ended by abort
'
'   *The meaning of the return value is not changed* MapsGeocodeSheet still returns
'   "the number of rows that became ok". Whether it was aborted is looked at here
'   (MapsLastError also becomes "canceled").
' ============================================================
Public Property Get MapsCanceled() As Boolean
    MapsCanceled = m_canceled
End Property


' ============================================================
' MapsCountRows (D-7) - count the rows to process first
'
'   *The denominator is needed* The body is "finish when an empty cell appears", so scan once with the same rule.
'   It only reads cells so it is fast. If 0 items, we can return without opening a tab.
'   Made Public for when the caller wants to show its own progress.
' ============================================================
Public Function MapsCountRows(ByVal targetSheet As Object, _
                              Optional ByVal firstRow As Long = 2, _
                              Optional ByVal addrCol As Long = 1) As Long
    Dim r As Long
    Dim n As Long

    If targetSheet Is Nothing Then Exit Function

    r = firstRow
    Do
        If Len(Trim$(CStr(targetSheet.Cells(r, addrCol).value))) = 0 Then Exit Do
        n = n + 1
        r = r + 1
    Loop

    MapsCountRows = n
End Function


' ============================================================
' MapsCheckCancel / MapsResetCancel (D-7, Private)
'
'   *Look at the OR of bit15 (currently pressed) and bit0 (pressed since last time)*
'   With a held-down check alone, it is missed during the few seconds' gap in which MapsPump does not run,
'   such as WaitSettled. With bit0, even a single press made during that gap is picked up.
'   GetAsyncKeyState returns both bits, so anything nonzero is regarded as an abort.
'
'   *Read and discard once at the start* Otherwise a previous press (an Esc pressed
'   before typing the verification, etc.) is carried over and it aborts the moment it starts.
'
'   *Side effect* Esc is also a key pressed in browser operation. Pressing it during a batch run stops it.
'   Since it is premised on running unattended, the real harm was judged small (D-7 point 3).
' ============================================================
Private Function MapsCheckCancel() As Boolean
    Dim k As Integer

    m_checkCount = m_checkCount + 1   ' D-7d: for measurement

    If m_cancel Then
        MapsCheckCancel = True
        Exit Function
    End If

    k = GetAsyncKeyState(MAPS_VK_ESCAPE)
    If k <> 0 Then
        m_cancel = True
        MapsCheckCancel = True
        Wv2Log.LogI "Wv2Maps: *abort requested by Esc*"
    End If
End Function

Private Sub MapsResetCancel()
    Dim k As Integer

    m_cancel = False
    m_canceled = False
    m_checkCount = 0
    k = GetAsyncKeyState(MAPS_VK_ESCAPE)   ' *Read and discard the previous press*
End Sub


' ============================================================
' MapsArmCancelKey / MapsDisarmCancelKey (D-7b, Private)
'
'   *Stop the VBA runtime's Esc only while inside Wv2Maps*
'   Without this, VBA grabs Esc first and breaks (confirmed on the D-7 real machine).
'   *A break is also the crash window of spec fact 20*, so the value of blocking it is double.
'
'   *Why xlDisabled rather than xlErrorHandler (D-7c)*
'   With xlErrorHandler Esc becomes error 18, but *it occurs mainly inside
'   WebView2 COM callbacks*, where the caller is not VBA, so it does not trace back
'   to our On Error GoTo (confirmed on the D-7b real machine).
'   With xlDisabled the error itself does not occur and execution does not stop, so
'   *GetAsyncKeyState can pick it up the next time MapsPump runs*.
'
'   *Always restore it* If forgotten, neither Esc nor Ctrl+Break works in the whole of Excel.
'   Arming nested would make the saved value xlDisabled,
'   but m_inBatch guarantees that "only the outermost arms", so it does not happen.
' ============================================================
Private Sub MapsArmCancelKey()
    m_savedCancelKey = Application.EnableCancelKey
    Application.EnableCancelKey = xlDisabled
End Sub

Private Sub MapsDisarmCancelKey()
    ' Even if the saved value is xlDisabled (0), return to xlInterrupt - restoring the fortress is safer
    If m_savedCancelKey = 0 Then m_savedCancelKey = xlInterrupt
    Application.EnableCancelKey = m_savedCancelKey
    m_savedCancelKey = 0
End Sub


' ============================================================
' MapsCheckCount (D-7d) - how many times abort detection could be called in the previous batch
'
'   *If this is only a few dozen against 15 seconds per item, most of the waiting is in Wv2Pane's
'   loops (EvalSync / WaitFor / WaitSettled), and there is no detection point there.*
' ============================================================
Public Property Get MapsCheckCount() As Long
    MapsCheckCount = m_checkCount
End Property


' ============================================================
' MapsRawEscState (D-7d) - the raw value of GetAsyncKeyState
'
'   *An observation entry to keep the declaration in one place* If the verification side writes the same Declare,
'   there would be two declarations. The probe (Test_D7_Key) hits this.
'   Return value: 0 = not pressed / anything else is the raw bits
'           equivalent to &H8000 (negative value) = currently pressed / 1 = pressed since last time
' ============================================================
Public Function MapsRawEscState() As Long
    MapsRawEscState = GetAsyncKeyState(MAPS_VK_ESCAPE)
End Function


' ============================================================
' MapsProgress (D-7, Private) - show progress in the status bar
'
'   *The sheet fills up during processing, but the denominator and remainder are not visible*, so
'   the count and the means to stop are shown in the status bar. When done it is always restored with False.
' ============================================================
Private Sub MapsProgress(ByVal doneRows As Long, ByVal totalRows As Long, _
                         ByVal okRows As Long, ByVal ngRows As Long)
    Application.StatusBar = "Address -> coordinates " & doneRows & "/" & totalRows & _
                           " (ok " & okRows & " / failed " & ngRows & ")" & _
                           "  -  press Esc to stop"
End Sub


' ============================================================
' MapsGeocodeSheet (D-5b) - Turn the address column of a sheet into coordinates in bulk
'
'   Arguments:
'     targetSheet : the target worksheet
'     firstRow    : starting row (default 2. Row 1 is assumed to be the header)
'     addrCol     : column number of the address (default 1 = column A)
'     outCol      : first column to write out to (default 2 = column B)
'     skipDone    : whether to skip rows whose status starts with ok (default True)
'     pickFirst   : whether to take the first one when there are multiple candidates (default False)
'
'   Layout written out (4 columns from outCol):
'     outCol   : latitude
'     outCol+1 : longitude
'     outCol+2 : normalized address (the name Maps returned)
'     outCol+3 : status (ok / *ok(candidate 1)* / ambiguous / not-found / timeout-url ...)
'                *ok(candidate 1) means "took the first of the candidate list"*, which has lower precision than settled
'
'   Return value: *the number of rows that became ok* (not the number of rows processed)
'
'   # *Handling when there are multiple candidates (ambiguous)*
'     **Write the center coordinates and note ambiguous in the status column.**
'     If silently skipped, the reason for "processed but empty" is not understood, and if silently treated as settled,
'     *wrong coordinates get mixed in*. The data is kept in a form a person can judge.
'
'   # *Even if stopped midway it can resume from where it left off*
'     With skipDone = True, rows whose status is ok are skipped. Even if stopped partway through 100 items,
'     calling again processes only the remainder. To redo, clear the status column.
'
'   # Progress and abort (D-7)
'     It writes row by row and runs DoEvents, so *the screen fills up during processing*.
'     A line is also left in the log for each. In addition *it is shown in the status bar in the form 12/100*.
'     *Press Esc when you want to stop* (no need to hold it down).
'     The abort takes effect *at a row boundary*, so no half-written row remains.
'     Whether it was aborted can be known with MapsCanceled (MapsLastError is also canceled).
'     *Nothing is written to the aborted row, so calling again with skipDone as is continues from there.*
'
'   Usage example:
'     Wv2Maps.MapsGeocodeSheet ActiveSheet
'     Wv2Maps.MapsGeocodeSheet Sheets("Addresses"), 3, 2, 5   ' From row 3, the addresses in column B, results to column E onward
' ============================================================
Public Function MapsGeocodeSheet(ByVal targetSheet As Object, _
                                 Optional ByVal firstRow As Long = 2, _
                                 Optional ByVal addrCol As Long = 1, _
                                 Optional ByVal outCol As Long = 2, _
                                 Optional ByVal skipDone As Boolean = True, _
                                 Optional ByVal pickFirst As Boolean = False) As Long
    Dim p As Wv2Pane
    Dim r As Long
    Dim addr As String
    Dim lat As Double
    Dim lng As Double
    Dim nm As String
    Dim ok As Boolean
    Dim doneCount As Long
    Dim total As Long
    Dim totalRows As Long
    Dim errNo As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim tStart As Single

    m_lastError = ""
    tStart = Timer

    If targetSheet Is Nothing Then
        m_lastError = "no-sheet"
        Exit Function
    End If

    ' --- D-7: *this is the outermost, so discard the abort state here* ---
    MapsResetCancel

    ' --- D-7: *count the denominator first* If 0 items, return without opening a tab ---
    totalRows = MapsCountRows(targetSheet, firstRow, addrCol)
    If totalRows = 0 Then
        m_lastError = "no-rows"
        Wv2Log.LogW "Wv2Maps.MapsGeocodeSheet: no rows to process"
        Exit Function
    End If
    Wv2Log.LogI "Wv2Maps.MapsGeocodeSheet: *all " & totalRows & " items* (press Esc to stop)"

    ' --- D-7b: *from here on everything goes through Cleanup* ---
    '   m_inBatch is raised before MapsOpen (so that MapsOpen does not arm).
    m_inBatch = True
    MapsArmCancelKey
    On Error GoTo Cleanup

    Set p = MapsOpen(UserForm1.CurrentBrowser)
    If p Is Nothing Then
        Wv2Log.LogE "Wv2Maps.MapsGeocodeSheet: cannot open Maps (" & m_lastError & ")"
        GoTo Cleanup
    End If

    r = firstRow
    Do
        ' --- D-7: *make the abort take effect at a row boundary* Do not leave a half-written row ---
        If MapsCheckCancel() Then
            m_canceled = True
            Wv2Log.LogW "Wv2Maps.MapsGeocodeSheet: *aborted* (just before row " & _
                        r & ")"
            Exit Do
        End If

        addr = Trim$(CStr(targetSheet.Cells(r, addrCol).value))
        If Len(addr) = 0 Then Exit Do

        total = total + 1
        MapsProgress total, totalRows, doneCount, (total - 1) - doneCount

        If skipDone And Left$(CStr(targetSheet.Cells(r, outCol + 3).value), 2) = "ok" Then
            Wv2Log.LogI "  [" & r & "] already done, skipping: " & addr
            doneCount = doneCount + 1
        Else
            ok = MapsGeocode(p, addr, lat, lng, nm, 20, pickFirst)

            ' --- D-7: *write nothing to an aborted row* (design principle 111) ---
            '   If canceled were written to the status column, afterwards it would read
            '   as if "that was the result". It was not processed, so leave it empty.
            If m_canceled Then
                total = total - 1
                Wv2Log.LogW "Wv2Maps.MapsGeocodeSheet: *aborted* (while processing row " & _
                            r & ")"
                Exit Do
            End If

            If ok Then
                targetSheet.Cells(r, outCol).value = lat
                targetSheet.Cells(r, outCol + 1).value = lng
                targetSheet.Cells(r, outCol + 2).value = nm
                ' *Write "settled" and "picked from candidates" differently* (do not line up things of different precision
                '   with the same face). The check on resume is done with Left$(..., 2) = "ok".
                If MapsLastPicked Then
                    targetSheet.Cells(r, outCol + 3).value = "ok(candidate 1)"
                Else
                    targetSheet.Cells(r, outCol + 3).value = "ok"
                End If
                doneCount = doneCount + 1
            Else
                ' *Even with multiple candidates the coordinates (center of the map) are left written*
                If lat <> 0 Or lng <> 0 Then
                    targetSheet.Cells(r, outCol).value = lat
                    targetSheet.Cells(r, outCol + 1).value = lng
                End If
                If Len(nm) > 0 Then targetSheet.Cells(r, outCol + 2).value = nm
                targetSheet.Cells(r, outCol + 3).value = MapsLastError
            End If

            Wv2Log.LogI "  [" & r & "] " & addr & " -> " & _
                        IIf(ok, lat & "," & lng, "*" & MapsLastError & "*")
        End If

        DoEvents
        r = r + 1
    Loop

' --- D-7b: *this is the only exit* Both normal termination and errors pass through here ---
Cleanup:
    errNo = Err.Number
    errDesc = Err.Description
    errSrc = Err.source
    On Error GoTo 0

    If errNo = 18 Then
        ' D-7c: xlDisabled was set so this normally does not come. Left for the window before/after arming.
        m_cancel = True
        m_canceled = True
        Wv2Log.LogW "Wv2Maps.MapsGeocodeSheet: *aborted by Esc* (while processing row " & _
                    r & ")"
        errNo = 0
    End If

    ' --- *Always restore* Leaving the status bar or EnableCancelKey would pollute the whole of Excel ---
    MapsDisarmCancelKey
    m_inBatch = False

    ' *D-7d: Empty, not False* With False, "FALSE" stays displayed
    Application.StatusBar = Empty
    If m_canceled Then m_lastError = "canceled"

    Wv2Log.LogD "Wv2Maps.MapsGeocodeSheet: after StatusBar restore = [" & _
                CStr(Application.StatusBar) & "] type=" & _
                TypeName(Application.StatusBar)

    ' *D-7d: how many times abort detection could be called* Material for isolating why it cannot be picked up (design principle 112)
    Wv2Log.LogI "Wv2Maps.MapsGeocodeSheet: *abort detection " & m_checkCount & " times / " & _
                Format$(MapsElapsed(tStart), "0.0") & " s*"

    Wv2Log.LogI "Wv2Maps.MapsGeocodeSheet: " & total & " rows, " & _
                doneCount & " ok" & IIf(m_canceled, " *aborted*", "")
    MapsGeocodeSheet = doneCount

    ' *Other errors are not swallowed* Only cleanup is done and it is raised as is
    If errNo <> 0 Then Err.Raise errNo, errSrc, errDesc
End Function

