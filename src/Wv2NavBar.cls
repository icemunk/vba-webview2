VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Wv2NavBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''
' --- Wv2NavBar.cls  stage K-4a (move the release-path logging to Wv2Log) ---
'
'   *Purpose is measurement* Releasing on form close is slow (a pause of about 0.5 s per item)
'   was the report, but this path was still Debug.Print and left nothing in a log,
'   so it was impossible to tell where the time was going.
'
'   Wv2Log has millisecond timestamps, so *just swapping it in lets us read the durations*.
'   Furthermore, measuring twice with Wv2Log.LogEcho = False / True separates
'   *"COM release is slow" from "output to the Immediate window is slow"*.
'
'   *We do not fix anything at this stage. We only measure.* (D-7 guessed wrong 3 times in a row)
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2NavBar.cls  stage 9.25b (removal of the verification probe / zero arguments confirmed permanently) ---
'
'   The 9.25 verification settled on "pattern A = safe", so this round removed the temporary code.
'
'   *Verification result (9.25 real machine)* The probe's log came out in this order:
'       (1) just before the get
'       (2) just after the get            <- *there is no HostProbe execution log in between*
'       *HostProbe was executed* (1st overall)
'       (3) f() call succeeded
'       *HostProbe was executed* (2nd overall)
'       (5) direct nb.HostProbe() call succeeded
'     - No execution log between (1)-(2) = the VBA side was not called at the time of the property get.
'     - (3) appeared = the value obtained is a function object, and it can be called later.
'       In other words the sync proxy defers the "method or property" decision until the actual ().
'     - (5) appeared = the form nb.HostProbe() that is actually used passes normally.
'     - Exactly 2 overall = zero extra executions.
'     With the real buttons too, HostBack / HostForward / HostReload each logged just one line,
'     there was no fall to the postMessage fallback, and <- went back exactly one page.
'
'   -> Recorded as **spec fact 22** (details in the dev notes). It is now confirmed that the dummy argument
'     reserved of 9.24 was unnecessary insurance, so zero arguments is confirmed permanently.
'     Note: however, the order itself of "start from the safe side and trim later" was correct. A double execution
'       would show up in the confusing form "one click goes back 2 pages", and since the hybrid's
'       insurance backfires, it is the kind of defect that is hard to notice afterwards.
'
'   *Removed (temporary code put in at 9.25)*
'     - the two Public Function HostProbe() / HostProbeMark(phase)
'     - probeParamless() and setTimeout(probeParamless, 500) inside BuildNavBarHtml
'     With this NavBar settles into exactly the 9.24 structure, with the 3 Host methods taking zero arguments.
'     Calls from JS are nb.HostBack() / nb.HostForward() / nb.HostReload() /
'     nb.HostNavigate(url).
'
'   *Not touched* the bodies of the 4 Host methods / the delegation structure of OnPaneWebMessage /
'     the three tiers of navHost() / attach / ready / Escape / editing / PushNavSyncToJs.
'     Design principles stay at 71. Spec facts go from 21 to 22.
'
' --- Wv2NavBar.cls  stage 9.25 (verification: whether zero-argument Host methods are viable + deleting reserved) ---
'
'   Scope of 9.25 (a small verification stage that clears the open item of 9.24. Plan C adopted):
'     *Question* When a zero-argument Public Function is touched from the hostObjects sync proxy,
'       does it get "executed at the time of the property get"? In 9.24 we feared this and kept the dummy argument
'       reserved, but the phenomenon did not appear only because we erred on the safe side, and whether it was needed
'       was never verified ("did not appear" is not proof of "unnecessary").
'
'     *Plan C = add a probe + delete reserved at the same time*
'       - Plan A (just remove reserved and try with the real buttons) is the smallest change, but when a double execution
'         occurs we could only guess whether the cause is "executed on get".
'       - Plan B (add only the probe and leave reserved) makes causality fully clear but takes 2 stages.
'       - -> With Plan C one launch gives both "the mechanism's facts (probe)" and "the real behavior (buttons)"
'         at the same time. Even if it misses, the real harm is only "<- goes back 2 pages once", and
'         one observation is enough, so both answers come out cheaply.
'
'     *Change 1: delete reserved*
'       Removed Optional ByVal reserved As Long from HostBack / HostForward / HostReload,
'       returning them to plain zero-argument functions. On the JS side nb.HostBack(0) -> nb.HostBack().
'       *This is the most breakable change this time.* If "executed on get" happens,
'         (1) one execution on get -> (2) the return value is Boolean so () gives a TypeError ->
'         (3) catch and postMessage fallback -> one more execution
'       and <- once goes back 2 pages. In the log HostBack appears as 2 lines (with the back of
'       OnPaneWebMessage between them) so it is obvious at a glance.
'
'     *Change 2: add a verification-only probe (temporary code removed in the next stage)*
'       - Public Function HostProbeMark(ByVal phase As Long) As Boolean
'           Has an argument so it is safe. Just stamps the phase into the log.
'       - Public Function HostProbe() As Boolean
'           *Zero arguments*. Just emits a log when called, with zero side effects. It involves no real
'           browser operation, so a double execution does no harm.
'       The JS side sets up setTimeout(probeParamless, 500) right after sending ready,
'       and at every startup runs the following once:
'           Mark(1) -> var f = nb.HostProbe (get only) -> Mark(2)
'           -> f() success Mark(3) / failure Mark(4)
'           -> direct nb.HostProbe() call success Mark(5) / failure Mark(6)
'           (if the get itself throws, Mark(9))
'       *How to read* If the execution log of HostProbe appears *between* Mark(1) and Mark(2),
'         then "it is executed on get" = reserved was needed. If it appears only after Mark(2)
'         it is safe = reserved is unnecessary. If Mark(5) appears it is direct proof that the
'         form nb.HostProbe() actually used passes.
'       *Reason for the 500ms delay* So as not to disturb the ready handshake and the round trip of the initial navsync,
'         and to keep a little away from the WebView2 initialization burst. It also makes the log blocks easier to read.
'
'     *Why the trigger is a setTimeout inside JS*
'       Wv2Pane.View_ExecuteScript (9.8c) already exists, but the NavBar instance is held by
'       UserForm1 as a Private field, so hitting it from outside would require adding a public entry
'       to UserForm1.frm. We do not want to grow production files for verification-only temporary code,
'       so we took the way of letting it run by itself inside JS.
'       tarboh only needs to launch normally and read the log.
'
'   *Not touched* the bodies of the 4 Host methods / the delegation structure of OnPaneWebMessage /
'     the three tiers of navHost() / ready / Escape / editing / PushNavSyncToJs.
'     Design principles stay at 71. Spec facts are set up after seeing the verification result.
'
' --- Wv2NavBar.cls  stage 9.24 (hostObjects conversion of NavBar + consolidating handling into Host) ---
'
'   Scope of 9.24 (roll out to NavBar, in the same form, the pattern established for TabBar in 9.23):
'     *Aim* In 9.21-9.23 TabBar's activate/close/reorder settled into the form
'       "JS is a hybrid (hostObjects first, postMessage fallback),
'       VBA consolidates the handling in one Host method, and the Select Case is a thin adapter".
'       NavBar's back/forward/reload/navigate are also aligned to exactly
'       the same pattern. Since the pattern is already established in TabBar, we do not take a stage like 9.21/9.22
'       of "first only hostObjects conversion", but assemble it all at once with Plan C-1 from the start (Plan 2).
'
'     *Entry points added (4 Host methods)*
'       - Public Function HostBack(Optional reserved) As Boolean    -> GoBackActive
'       - Public Function HostForward(Optional reserved) As Boolean -> GoForwardActive
'       - Public Function HostReload(Optional reserved) As Boolean  -> ReloadActive
'       - Public Function HostNavigate(url) As Boolean              -> NavigateActive
'       All are fire-and-forget on the Browser side (no completion wait = Plan 4-A), so even a direct
'       sync call from JS does not fall into the spec fact 14 deadlock.
'
'     **Reason for the dummy argument reserved (point 4b)**
'       All 3 of TabBar's had arguments, but back/forward/reload naturally
'       take zero arguments. The hostObjects sync proxy cannot distinguish from the JS side
'       "method or property", so it may be implemented so that at the moment of *getting* ho.NavBar.HostBack
'       IDispatch::Invoke(DISPATCH_PROPERTYGET) is sent.
'       VBA tolerantly accepts this for a Public Function without arguments and executes
'       it, so in the worst case (1) one execution on get -> (2) the return value is Boolean so ()
'       gives a TypeError -> (3) catch and postMessage fallback -> one more execution,
'       a double execution where "pressing <- once goes back twice" could happen. Since it is a form where the hybrid's
'       insurance backfires, the fallback cannot prevent it.
'       -> Give it one argument to eliminate the ambiguity in principle (JS calls nb.HostBack(0)).
'       From the VBA-side Case it is Optional so HostBack can be called plainly. If the real machine shows
'       "zero arguments is safe too", reserved may be dropped in the next stage.
'
'     *Where the empty-URL guard goes (point 3)* The If Len(Trim$(url)) = 0 check
'       that was inside the Case was moved into HostNavigate. Now whether it arrives from the hostObjects path
'       or the postMessage path, it passes through the same guard (same DRY thinking as 9.23).
'       The JS-side if(v && v.length) is left as insurance.
'
'     *The JS side is made hybrid with the navHost() helper* TabBar wrote out 3 places
'       plainly, but NavBar has 4 places and the form is uniform, so they are bundled into a helper.
'         navHost(msgObj, callHost)
'           (1) if ho.NavBar exists, call callHost(ho.NavBar) as a direct sync call -> [host]
'           (2) on exception, catch and postCmd(msgObj)      -> [host-fail->postmsg]
'           (3) if it could not be obtained at all, postCmd(msgObj)     -> [postmsg]
'       The JSON sent by the fallback is exactly the same as 9.12d, so the VBA-side
'       Select Case (thin adapter) can take it as is.
'       *Access to ho.NavBar is with dot notation only* (the form with a track record in TabBar.
'       With a computed property notation like ho.NavBar[name] the proxy's behavior cannot be predicted.)
'
'     *attach* In Init, before Navigate, put
'       m_barPane.View_AddHostObjectToScript("NavBar", Me) (order A of spec facts 18/19.
'       The same position as 5b of TabBar). Failure is not fatal, so hr is logged
'       and Init continues (JS just falls to postMessage).
'
'   *Not touched* the ready case (it is a notification so postMessage is right, design principle 61) /
'     Case "" / Case Else / PushNavSyncToJs / SendNavSync / the JS on the navsync receiving side /
'     the behavior of the editing flag (editing) and lastUrl / discarding the edit on Escape / the Browser's
'     5-event subscription / the 3-tier layout. Design principles stay at 71 (it is a practice of Plan 4-A and DRY,
'     not a new principle).
'
'   *Candidate for a new fact to confirm on the real machine (spec fact 21)*
'     Passing a "string argument" from JS to VBA via hostObjects.sync is a first this time
'     (TabBar had only Long). If URLs / search terms containing Japanese or symbols arrive as BSTR
'     without loss, record it as spec fact 21. The postMessage path went through JSON
'     escaping (a party to spec fact 16), so the prior expectation is that hostObjects should
'     if anything be safer.
'
' --- Wv2NavBar.cls  stage 9.13 ---
'
'   Changes in 9.13 (tidying: sharing the JSON helpers):
'     Removed JsonEscape / GetJsonStr / BoolToJson that were held as Private, and
'     switched to using the Public functions of the newly created standard module Wv2Json.bas.
'     The same functions were duplicated in Wv2TabBar too; they are consolidated in one place.
'       GetJsonStr -> Wv2Json.JsonGetStr
'       BoolToJson -> Wv2Json.JsonBool
'       JsonEscape -> Wv2Json.JsonEscape (same name)
'     *The \u path of JsonEscape was also fixed.* High characters (Japanese etc.) no longer
'     get garbled into \uXXXX (spec fact 16. Behavior is unchanged but the log is easier to read).
'
' --- Wv2NavBar.cls  stage 9.12d ---
'
'   Scope of 9.12d (JS->VBA: enabling input and clicks). With this 9.12 is complete:
'     Connect the nav bar's buttons and URL input to VBA navigation operations.
'       - <- click    -> {"cmd":"back"}                 -> Browser.GoBackActive
'       - -> click    -> {"cmd":"forward"}              -> Browser.GoForwardActive
'       - [Reload] click    -> {"cmd":"reload"}               -> Browser.ReloadActive
'       - Enter in the URL field -> {"cmd":"navigate","url":"..."} -> Browser.NavigateActive
'
'   *Re-entrancy safety = design principle 63* All navigation operations are ones that "contain no completion wait"
'     (Navigate / GoBack / GoForward / Reload are all fire-and-forget), so
'     it is safe to call them synchronously from inside a WebMessage event handler. The delayed execution of newtab in 9.11d
'     (waiting for Controller creation -> spec fact 14 deadlock) is not needed.
'
'   *Scheme completion is done on the Browser side (NavigateActive) (Plan 3-A)*
'     JS just sends the input string as is. The normalization "example.com" -> "https://example.com" is
'     the responsibility of Wv2Browser.NormalizeUrl (implemented in 9.12a).
'
'   *Discard the edit on Escape (Plan A: restore immediately from lastUrl on the JS side)*
'     At first Escape only set editing = false and blurred, but after that there was nothing to put back
'     the value of the URL field, so the typed string stayed (found on the real machine).
'     u.value gets rewritten only when the next navsync arrives, and with Escape nothing happens
'     on the VBA side (= no sync is sent).
'     -> The JS side keeps lastUrl (the "URL that should be shown now" conveyed by the latest navsync) and
'        restores it immediately on Escape. No VBA round trip is needed and the feel is good. lastUrl is
'        *always updated even while editing* (so that even if a navigation happens in the background during editing,
'        Escape can return to "the real current URL").
'
'   *Set editing = false on Enter*
'     When navigation starts, TabUrlChanged -> navsync replaces the URL field with the normalized URL
'     (typing "example.com" turns into "https://example.com/").
'     This is the same behavior as Chrome/Edge. If editing is not dropped, the URL field stays frozen
'     with the typed string.
'
' --- Wv2NavBar.cls  stage 9.12c ---
'
'   Scope of 9.12c (VBA->JS: display of the URL and button states. Input and clicks are 9.12d):
'     Receiving Browser's events, send the nav bar JS the active tab's "URL /
'     can go back / can go forward" and have it displayed.
'
'   *Full-sync method (design principle 60)* Same as 9.11c. Sent in full every time the state changes:
'       {"type":"navsync","url":"...","canBack":true,"canForward":false}
'     The JS side just overwrites the URL field and the buttons' disabled with the received values.
'
'   *ready handshake (design principle 61)* Same as 9.11c. JS postMessages
'     {"cmd":"ready"} when loading completes -> only after VBA receives it does it return the initial sync.
'     A PostWebMessageAsString sent before the JS message listener is registered gets dropped
'     (spec fact 15).
'
'   **index filter (a point specific to 9.12c)**
'     TabUrlChanged conveys by index "which tab's" URL changed. The nav bar
'     shows only the active tab's state, so events with index <> ActiveIndex
'     *must be ignored*. If not ignored, the moment a background tab finishes loading
'     the URL field turns into the background tab's URL (in fact in the 9.12b log, a TabUrlChanged of index=2
'     was sent while index=3 was active).
'
'   *On ActiveChanged, ask Browser for the current state*
'     When tabs are switched, the new active tab's URL does not come through
'     TabUrlChanged (because nothing happened in that tab). So on receiving ActiveChanged,
'     re-read the current state with Browser.GetActiveUrl / GetActiveCanGoBack / GetActiveCanGoForward
'     and send it. These read entries were prepared in 9.12a.
'
'   *Bug found on the real machine and its fix, No. 3 (9.12c: the URL being typed disappears)*
'     When a link was clicked while typing in the URL field and a new tab opened, the
'     string being typed disappeared. The cause was the u.disabled sync introduced in "No. 2":
'       assigning disabled makes the input blur
'         -> the blur handler sets editing = false
'         -> at the next sync u.value is overwritten and the input vanishes
'     *The 0-tab handling (No. 2) broke the protection while editing (No. 1)*
'     -> The countermeasure is two-tiered:
'        (1) While editing, do not touch the URL field at all (neither value nor disabled)
'        (2) Assign disabled only when the value changes (prevents useless blur)
'        Buttons do not steal focus, so they may always be synced.
'     Lesson: attaching/detaching disabled throws focus away. For an element that holds focus-dependent
'           state (such as an editing flag), do not assign the attribute unconditionally.
'
'   *Bug found on the real machine and its fix, No. 2 (9.12c: the URL field freezes at 0 tabs)*
'     When the last tab was closed, the nav bar's URL field kept the closed tab's URL.
'     Cause: Wv2Browser.CloseTab does not call ActivateTab when "there is no tab to activate", so
'     *ActiveChanged is not fired* (the 4-branch design of 9.11a: 0 tabs is
'     allowed but ActiveChanged(0) is not fired). NavBar only subscribed to
'     ActiveChanged, so it was not told that the tab count became 0.
'     -> Fix (Plan A): NavBar also subscribes to TabClosed and does PushNavSyncToJs.
'        PushNavSyncToJs already handles 0 tabs correctly (empty URL, all buttons disabled).
'        *The Browser's event contract is not changed.* It follows the loose-coupling principle that
'        "the display side subscribes to the events needed for display". In cases where both TabClosed and ActiveChanged
'        are fired it fires twice, but with the full-sync method it is harmless (design principle 60).
'     Along with this hasTab was added to navsync, and the reload button is also disabled at 0 tabs.
'
'   *Bug found on the real machine and its fix, No. 1 (9.12c)*
'     At first protection while editing the URL field was judged with document.activeElement !== u, but
'     in WebView2's standalone document #url stayed as activeElement and never came back, giving the symptom
'     *it worked only for the single initial sync and the URL field was never updated again*
'     (the buttons' enabled state followed correctly, so applyNavSync itself ran every time).
'     -> Changed to a method that holds "whether the user has actually started editing" in an explicit flag (editing).
'        Merely focusing does not make it editing; key input (the input event) makes it editing,
'        and blur / Enter / Escape release it.
'     Lesson: do not use focus state as a proxy for "whether editing" (design principle 65).
'
'   *About the duplicated JSON helpers*
'     JsonEscape / GetJsonStr also exist identically in Wv2TabBar (they are Private so
'     cannot be shared). The policy is to do the extraction to a standard module together in a future tidying stage,
'     and for now the duplication in 2 places is tolerated (agreed).
'
' --- Wv2NavBar.cls  stage 9.12b ---
'
'   *WebView2-based navigation bar (address bar + back/forward/reload)*
'
'   Follows as is the pattern of "using WebView2 as a UI component" (Wv2TabBar) established in 9.11.
'   Another thin, long WebView2 is placed in the 2nd row at the top of the screen,
'   and the nav bar is drawn with HTML/CSS/JS.
'
'   Layout (UserForm_Resize splits it into 3):
'     +--------------------------+
'     | [Bing][MSN x] [+]        | <- Wv2TabBar   (height TABBAR_HEIGHT_PX)
'     +--------------------------+
'     | [<][>][R] [https://...  ]| <- Wv2NavBar   (height NAVBAR_HEIGHT_PX) *this class
'     +--------------------------+
'     |         Content          | <- the tab group of Wv2Browser
'     +--------------------------+
'
'   *Ownership structure (same as Wv2TabBar, Plan b)*
'     Wv2NavBar holds one Wv2Pane of its own (m_barPane). This m_barPane is not put into
'     Wv2Browser.m_tabs (it is not placed under tab management). Therefore AddTab is not
'     used and Wv2Pane is created directly.
'
'   *Loose coupling (uses the entries prepared in 9.12a)*
'     Wv2Browser does not know the existence of Wv2NavBar. NavBar subscribes to Browser's events
'     with WithEvents and only calls Browser's public API (NavigateActive / GoBackActive etc.).
'     On the Browser side the following were already prepared in 9.12a:
'       - Public Event TabUrlChanged(index, url, canGoBack, canGoForward)
'       - Public Event ActiveChanged(index)  (switches the URL display when tabs are switched)
'       - NavigateActive / GoBackActive / GoForwardActive / ReloadActive / StopActive
'       - GetActiveUrl / GetActiveCanGoBack / GetActiveCanGoForward
'
'   *Environment sharing*
'     m_barPane shares the same Environment as Wv2Browser. Init receives
'     browser.Environment and passes it to m_barPane.Init.
'
'   Scope of 9.12b (skeleton only):
'     Build : Init (pane creation -> get View -> virtual host mapping -> show static HTML)
'             / Resize / Shutdown / IsReady / empty implementations of the WithEvents handlers
'     Not built: display of the URL and button states (9.12c) / enabling input and clicks (9.12d)
'     -> In 9.12b a "look-only nav bar" is shown in the 2nd row. The buttons cannot be pressed and
'       the URL field is not updated. The purpose is to confirm that the 3-way split layout coexists correctly.
'
'   *HTML delivery*
'     Like Wv2TabBar, uses SetVirtualHostNameToFolderMapping.
'     *Use a folder separate from TabBar.* The host name is appassets.navbar
'     (TabBar is appassets.tabbar). The delivery folder is also separate (specified by the caller).
'
'   *Spec fact 13* Specify the Navigate target down to the file name
'     (https://appassets.navbar/navbar.html).
'   *Design principle 56* If Ctrl_GetCoreWebView2 is omitted, Navigate gives E_POINTER.
''''''''''''''''''''''''''''''''''

Option Explicit

' ============================================================
' Fields
' ============================================================
Private WithEvents m_browser As Wv2Browser   ' Subscribes to Browser's events
Attribute m_browser.VB_VarHelpID = -1
Private m_barPane    As Wv2Pane              ' Single pane for drawing the nav bar
Private m_hwndHost   As LongPtr              ' hwnd of the host (parent frame)
Private m_folderPath As String               ' Delivery folder of the nav bar HTML
Private m_ready      As Boolean              ' Init-complete flag

' ============================================================
' Constants
' ============================================================
Private Const NAVBAR_HOST As String = "appassets.navbar"
Private Const NAVBAR_FILE As String = "navbar.html"

' Timeout for waiting for Controller Ready (seconds)
Private Const READY_TIMEOUT As Single = 10#


Private Sub Class_Initialize()
    Set m_browser = Nothing
    Set m_barPane = Nothing
    m_hwndHost = 0
    m_folderPath = ""
    m_ready = False
End Sub


' ============================================================
' Init ? Starts the nav bar
'
'   Arguments:
'     browser    : Wv2Browser (must be Init-ed. Its Environment is borrowed)
'     hwnd       : hwnd of the parent frame (placed on the same host as Browser)
'     folderPath : folder to write navbar.html to
'                  *Use a folder separate from TabBar*
'
'   Procedure (same as Wv2TabBar.Init):
'     1. Write out the nav bar HTML
'     2. Create the pane (AddTab is not used = not placed under tab management)
'     3. Wait for Controller Ready
'     4. Get View (design principle 56)
'     5. Virtual host mapping (required before Navigate)
'     6. Navigate (spec fact 13: specify down to the file name)
'
'   *Set m_browser = browser starts the WithEvents subscription*
'     From then on, Browser's TabUrlChanged / ActiveChanged reach this class.
'     In 9.12b the handlers are Debug.Print only (the content is 9.12c).
' ============================================================
Public Function Init( _
    ByVal browser As Wv2Browser, _
    ByVal hwnd As LongPtr, _
    ByVal folderPath As String) As Boolean

    If browser Is Nothing Then
        Debug.Print "Wv2NavBar.Init: browser is Nothing"
        Exit Function
    End If

    Dim env As Wv2Environment
    Set env = browser.Environment
    If env Is Nothing Then
        Debug.Print "Wv2NavBar.Init: browser.Environment is Nothing (Browser.Init not finished?)"
        Exit Function
    End If

    Set m_browser = browser        ' *The WithEvents subscription starts with this
    m_hwndHost = hwnd
    m_folderPath = folderPath

    ' --- 1. Write out the nav bar HTML ---
    If Not WriteNavBarHtml() Then
        Debug.Print "Wv2NavBar.Init: HTML write failed"
        Exit Function
    End If

    ' --- 2. Create the pane for the nav bar (AddTab is not used) ---
    Set m_barPane = New Wv2Pane
    m_barPane.Init env, m_hwndHost

    ' --- 3. Wait for Controller Ready ---
    If Not WaitPaneReady(m_barPane) Then
        If m_barPane.IsFailed Then
            Debug.Print "Wv2NavBar.Init: Controller creation failed (lastError=&H" & _
                        Hex(m_barPane.LastError) & ")"
        Else
            Debug.Print "Wv2NavBar.Init: Controller creation timed out"
        End If
        Exit Function
    End If
    Debug.Print "Wv2NavBar.Init: Controller creation OK (pCtrl=" & m_barPane.CtrlPtr & ")"

    ' --- 4. Get View (if omitted, Navigate gives E_POINTER, design principle 56) ---
    Dim hr As Long
    hr = m_barPane.Ctrl_GetCoreWebView2()
    If hr <> 0 Or m_barPane.ViewPtr = 0 Then
        Debug.Print "Wv2NavBar.Init: View acquisition failed hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2NavBar.Init: View acquisition OK (pView=" & m_barPane.ViewPtr & ")"

    ' --- 5. Virtual host mapping (required before Navigate) ---
    hr = m_barPane.View3_SetVirtualHostNameToFolderMapping(NAVBAR_HOST, m_folderPath, 2)  ' 2 = DENY_CORS (English fork; was 1 = ALLOW, which let any site read the mapped folder)
    If hr <> 0 Then
        Debug.Print "Wv2NavBar.Init: mapping failed hr=&H" & Hex(hr)
        Exit Function
    End If

    ' --- 5b. Expose hostObjects (9.24, required before Navigate: order A of spec facts 18/19) ---
    '   Same position and same form as 5b (9.21) of TabBar. What is attached is NavBar itself (Me).
    '   From JS it is called as window.chrome.webview.hostObjects.sync.NavBar.HostBack() /
    '   HostForward() / HostReload() / HostNavigate(url) (made zero-argument in 9.25).
    '   Failure is not fatal (the JS side falls back to postMessage), so
    '   here we only log hr and Init continues.
    Dim hrHost As Long
    hrHost = m_barPane.View_AddHostObjectToScript("NavBar", Me)
    If hrHost <> 0 Then
        Debug.Print "Wv2NavBar.Init: AddHostObjectToScript(NavBar) failed hr=&H" & Hex(hrHost) & " (JS falls back to postMessage)"
    Else
        Debug.Print "Wv2NavBar.Init: NavBar bridge attach OK (back/forward/reload/navigate prefer hostObjects)"
    End If

    ' --- 6. Navigate (spec fact 13: specify down to the file name) ---
    Dim url As String
    url = "https://" & NAVBAR_HOST & "/" & NAVBAR_FILE
    hr = m_barPane.View_Navigate(url)
    If hr <> 0 Then
        Debug.Print "Wv2NavBar.Init: Navigate failed hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2NavBar.Init: Navigate OK (" & url & ")"

    ' --- 7. Register WebMessage sink + enable reception (9.12c, Plan P) ---
    '   Make it possible to receive postMessage from the nav bar JS (ready / future clicks and URL input)
    '   in OnPaneWebMessage.
    '   *The initial sync is "not fired here"* A PostMessage before the JS message
    '   listener is registered gets dropped (spec fact 15). We do PushNavSyncToJs inside OnPaneWebMessage
    '   after JS sends {"cmd":"ready"} (design principle 61).
    m_barPane.SetWebMsgSink Me
    hr = m_barPane.View_AddWebMessageReceived()
    If hr <> 0 Then
        Debug.Print "Wv2NavBar.Init: View_AddWebMessageReceived failed hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2NavBar.Init: WebMessage reception enabled (waiting for ready)"

    m_ready = True
    Init = True
    Debug.Print "Wv2NavBar.Init: done (nav bar started, initial sync after ready is received)"
End Function


' ============================================================
' Resize ? Applies the rectangle of the nav bar band
'   Takes as is the rectangle of the "2nd-row band" that UserForm_Resize computed.
'   *NavBar does not compute where it is placed.* The source of truth for the layout is
'   the upper layer (UserForm). Here it just follows the rectangle it is told (the thinking of design principle 59).
' ============================================================
Public Sub Resize(ByRef barRect As RECT)
    If m_barPane Is Nothing Then Exit Sub
    Dim hr As Long
    hr = m_barPane.Ctrl_PutBounds(barRect)
    If hr <> 0 Then
        Debug.Print "Wv2NavBar.Resize: Ctrl_PutBounds failed hr=&H" & Hex(hr)
    End If
    m_barPane.Ctrl_PutIsVisible True   ' The nav bar is always visible
End Sub


' ============================================================
' Shutdown ? Release explicitly (design principle 58)
'   Cuts the WithEvents reference and the own pane. Called from UserForm_Terminate.
' ============================================================
Public Sub Shutdown()
    Wv2Log.LogD "Wv2NavBar.Shutdown start"
    If Not m_barPane Is Nothing Then
        m_barPane.ClearWebMsgSink   ' The sink is registered in 9.12d, so build in the habit of cutting it first
        m_barPane.Ctrl_Close
        Set m_barPane = Nothing
    End If
    Set m_browser = Nothing   ' Cancel the WithEvents subscription
    m_ready = False
    Wv2Log.LogD "Wv2NavBar.Shutdown done"
End Sub


Public Property Get IsReady() As Boolean
    IsReady = m_ready
End Property


' ============================================================
' *Wv2Browser event handlers (9.12c: reflect to the nav bar via full sync)*
'
'   TabUrlChanged : the active tab's URL / whether back and forward are possible changed.
'                   *The index filter is mandatory* Events of background tabs are ignored
'                   (if not ignored the URL field turns into the background tab's URL).
'                   It is fired from both NavigationCompleted and HistoryChanged, so
'                   multiple arrivals for one navigation are normal (harmless thanks to full sync).
'
'   ActiveChanged : the tab was switched. *The new active tab's URL does not
'                   come through TabUrlChanged* (because nothing happened
'                   in that tab). So ask Browser for the current state and send it.
' ============================================================
Private Sub m_browser_TabUrlChanged(ByVal index As Long, ByVal url As String, _
                                    ByVal canGoBack As Boolean, ByVal canGoForward As Boolean)
    ' --- *index filter* ignore events of background tabs ---
    If m_browser Is Nothing Then Exit Sub
    If index <> m_browser.ActiveIndex Then
        Debug.Print "Wv2NavBar[EVT] TabUrlChanged(index=" & index & ") -> ignored (background tab) " & _
                    "(ActiveIndex=" & m_browser.ActiveIndex & ")"
        Exit Sub
    End If

    Debug.Print "Wv2NavBar[EVT] TabUrlChanged(index=" & index & ", url=" & url & _
                ", back=" & canGoBack & ", fwd=" & canGoForward & ") -> PushNavSyncToJs"

    ' Use the value the event carried as is (do not query Browser again)
    SendNavSync url, canGoBack, canGoForward
End Sub

Private Sub m_browser_ActiveChanged(ByVal index As Long)
    Debug.Print "Wv2NavBar[EVT] ActiveChanged(index=" & index & ") -> PushNavSyncToJs"
    PushNavSyncToJs
End Sub

' --- TabClosed (added in 9.12c, Plan A) ---
'   *Why it is needed* CloseTab does not call ActivateTab when "there is no tab to activate" (= when the last
'   tab is closed), so ActiveChanged is not fired. If only ActiveChanged
'   is subscribed, it is not told that the tab count became 0 and the URL field freezes (found on the real machine).
'   If TabClosed is also subscribed and PushNavSyncToJs is done, at 0 tabs the URL becomes empty and all buttons are disabled
'   (PushNavSyncToJs already handles 0 tabs).
'   When tabs remain it fires twice together with ActiveChanged, but this is harmless with the full-sync method.
Private Sub m_browser_TabClosed(ByVal index As Long)
    Debug.Print "Wv2NavBar[EVT] TabClosed(index=" & index & ") -> PushNavSyncToJs"
    PushNavSyncToJs
End Sub


' ============================================================
' OnPaneWebMessage (9.12c, Plan P: receiving end of the m_barPane WebMessage sink)
'
'   The nav bar JS's postMessage arrives here via Wv2Pane.
'   In 9.12c only "ready" is interpreted. The command interpretation for the click types (back/forward/reload/navigate)
'   is added in 9.12d.
' ============================================================
Public Sub OnPaneWebMessage(ByVal pane As Wv2Pane, ByVal source As String, ByVal msg As String)
    Debug.Print "Wv2NavBar.OnPaneWebMessage: msg=" & msg

    If m_browser Is Nothing Then
        Debug.Print "  -> ignored because m_browser is Nothing"
        Exit Sub
    End If

    Dim cmd As String
    cmd = JsonGetStr(msg, "cmd")

    Select Case cmd

        Case "ready"
            Debug.Print "  -> ready received. Sending the initial navsync."
            PushNavSyncToJs

        Case "back"
            ' --- 9.12d/9.24: <- click (postMessage fallback path) ---
            '   *9.24: processing consolidated into HostBack (Plan C-1)* Only when JS in the hybrid
            '   could not grab hostObjects does it fall here. The GoBackActive call and the
            '   failure log are already consolidated inside HostBack, so here it is fire-and-forget
            '   delegation only (the return value is not taken). That once navigation starts, HistoryChanged /
            '   NavigationCompleted -> TabUrlChanged -> PushNavSyncToJs makes the nav bar
            '   follow automatically (full sync) is as before.
            Debug.Print "  -> back [postmsg fallback -> HostBack]"
            HostBack

        Case "forward"
            ' --- 9.12d/9.24: -> click (postMessage fallback path) ---
            Debug.Print "  -> forward [postmsg fallback -> HostForward]"
            HostForward

        Case "reload"
            ' --- 9.12d/9.24: [Reload] click (postMessage fallback path) ---
            Debug.Print "  -> reload [postmsg fallback -> HostReload]"
            HostReload

        Case "navigate"
            ' --- 9.12d/9.24: Enter in the URL field (postMessage fallback path) ---
            '   *9.24: processing consolidated into HostNavigate (Plan C-1)* The empty-URL guard /
            '   the NavigateActive call / the failure log have all been moved into HostNavigate
            '   (point 3). Here we only restore url from JSON and delegate fire-and-forget.
            '   The JSON parse (JsonGetStr) is needed to restore the arguments on the
            '   fallback path, so it stays on the Case side (on the hostObjects path JS passes the string
            '   directly so no parsing is needed).
            '   *Scheme completion and routing to the search engine are the responsibility of the Browser side (NormalizeUrl)*
            '   Neither this nor the Host method touches the input string; they pass it through as is (Plan 3-A).
            Dim url As String
            url = JsonGetStr(msg, "url")
            Debug.Print "  -> navigate(url=" & url & ") [postmsg fallback -> HostNavigate]"
            HostNavigate url

        Case ""
            Debug.Print "  -> could not extract cmd (msg=" & msg & ")"

        Case Else
            Debug.Print "  -> unknown command: " & cmd

    End Select
End Sub


' ============================================================
' HostBack (9.24, the back entry of the hostObjects path)
'
'   Called directly from JS as window.chrome.webview.hostObjects.sync.NavBar.HostBack()
'   (a synchronous RPC that does not go through postMessage). The body is the same as the back
'   case of OnPaneWebMessage: it just calls m_browser.GoBackActive synchronously.
'
'   *Safety of the synchronous call (Plan 4-A)* GoBackActive only fires View_GoBack and
'     contains no completion wait, so even a JS->VBA sync direct call does not fall into the spec fact 14
'     deadlock. When navigation starts, HistoryChanged /
'     NavigationCompleted -> TabUrlChanged -> PushNavSyncToJs makes the nav bar follow
'     through full sync.
'
'   *Zero arguments confirmed (9.25 -> 9.25b)* In 9.24 we had given it
'     Optional ByVal reserved As Long out of the concern that "a zero-argument Public Function
'     would be mistaken by the hostObjects sync proxy for a property get, causing a double execution:
'     one execution at the time of the get, and another by the fallback".
'     As a result of measuring with a dedicated probe in 9.25, it was confirmed that the sync proxy
'     does not call VBA at the time of the get and defers the decision until the actual ()
'     (spec fact 22). So reserved was confirmed unnecessary and returned to plain zero arguments. JS calls it as nb.HostBack().
' ============================================================
Public Function HostBack() As Boolean
    Debug.Print "Wv2NavBar.HostBack: (via host)"
    If m_browser Is Nothing Then
        Debug.Print "  -> ignored because m_browser is Nothing"
        HostBack = False
        Exit Function
    End If
    Dim ok As Boolean
    ok = m_browser.GoBackActive()
    If Not ok Then
        Debug.Print "    [warning] GoBackActive failed / cannot go back"
    End If
    HostBack = ok
End Function


' ============================================================
' HostForward (9.24, the forward entry of the hostObjects path)
'
'   Called directly from JS as hostObjects.sync.NavBar.HostForward(). The body is
'   m_browser.GoForwardActive, same as the forward case. It only fires View_GoForward
'   with no completion wait (Plan 4-A). In 9.25 reserved was deleted and it returned to zero arguments.
' ============================================================
Public Function HostForward() As Boolean
    Debug.Print "Wv2NavBar.HostForward: (via host)"
    If m_browser Is Nothing Then
        Debug.Print "  -> ignored because m_browser is Nothing"
        HostForward = False
        Exit Function
    End If
    Dim ok As Boolean
    ok = m_browser.GoForwardActive()
    If Not ok Then
        Debug.Print "    [warning] GoForwardActive failed / cannot go forward"
    End If
    HostForward = ok
End Function


' ============================================================
' HostReload (9.24, the reload entry of the hostObjects path)
'
'   Called directly from JS as hostObjects.sync.NavBar.HostReload(). The body is
'   m_browser.ReloadActive, same as the reload case. It only fires View_Reload
'   with no completion wait (Plan 4-A). In 9.25 reserved was deleted and it returned to zero arguments.
' ============================================================
Public Function HostReload() As Boolean
    Debug.Print "Wv2NavBar.HostReload: (via host)"
    If m_browser Is Nothing Then
        Debug.Print "  -> ignored because m_browser is Nothing"
        HostReload = False
        Exit Function
    End If
    Dim ok As Boolean
    ok = m_browser.ReloadActive()
    If Not ok Then
        Debug.Print "    [warning] ReloadActive failed"
    End If
    HostReload = ok
End Function


' ============================================================
' HostNavigate (9.24, the navigate entry of the hostObjects path)
'
'   Called directly from JS as hostObjects.sync.NavBar.HostNavigate(url). The body is
'   m_browser.NavigateActive, same as the navigate case.
'
'   ** The highlight of this verification = the string argument ** TabBar's 3 methods accepted only
'     Long, so this is the first time a String is passed from JS->VBA via hostObjects.sync.
'     A JS string should arrive as VARIANT(BSTR) and passes through no JSON escaping
'     like the postMessage path. If Japanese search terms and URLs with symbols arrive without
'     loss, record it as spec fact 21. -> For judging, the received url is
'     Debug.Print-ed as is (garbling can be seen in the Immediate window).
'
'   *The empty-URL guard is consolidated here (point 3)* Whether it comes from the hostObjects path or the postMessage path,
'     it passes through the same guard.
'
'   *Safety of the synchronous call (Plan 4-A)* NavigateActive does NormalizeUrl (pure conversion) and
'     View_Navigate (fire-and-forget) only, with no completion wait.
' ============================================================
Public Function HostNavigate(ByVal url As String) As Boolean
    Debug.Print "Wv2NavBar.HostNavigate: url=" & url & " (via host)"
    If m_browser Is Nothing Then
        Debug.Print "  -> ignored because m_browser is Nothing"
        HostNavigate = False
        Exit Function
    End If
    If Len(Trim$(url)) = 0 Then
        Debug.Print "    [warning] URL is empty (ignored)"
        HostNavigate = False
        Exit Function
    End If
    Dim ok As Boolean
    ok = m_browser.NavigateActive(url)
    If Not ok Then
        Debug.Print "    [warning] NavigateActive failed / ignored (no active tab?)"
    End If
    HostNavigate = ok
End Function


' ============================================================
' PushNavSyncToJs (9.12c, ask Browser for the current state and send it to JS)
'
'   Gets the active tab's URL / whether it can go back / whether it can go forward from Browser's read entries
'   (prepared in 9.12a) and sends them. At 0 tabs: empty URL, both buttons disabled.
'
'   Callers:
'     - when ready is received (initial sync)
'     - ActiveChanged (tab switch. The new tab's state is not carried by the event, so it is queried)
' ============================================================
Private Sub PushNavSyncToJs()
    If m_browser Is Nothing Then Exit Sub

    Dim url As String
    Dim canBack As Boolean
    Dim canFwd As Boolean

    If m_browser.TabCount = 0 Or m_browser.ActiveIndex < 1 Then
        ' 0 tabs (no active tab): empty URL, both buttons disabled
        url = ""
        canBack = False
        canFwd = False
    Else
        url = m_browser.GetActiveUrl()
        canBack = m_browser.GetActiveCanGoBack()
        canFwd = m_browser.GetActiveCanGoForward()
    End If

    SendNavSync url, canBack, canFwd
End Sub


' ============================================================
' SendNavSync (9.12c, build the navsync JSON and send it to JS)
'
'   JSON sent:
'     {"type":"navsync","url":"...","canBack":true,"canForward":false}
'
'   The common sending part called from both PushNavSyncToJs (the version that queries Browser) and the TabUrlChanged handler
'   (the version that uses the values the event carried as is).
' ============================================================
Private Sub SendNavSync(ByVal url As String, ByVal canBack As Boolean, ByVal canFwd As Boolean)
    If m_barPane Is Nothing Then Exit Sub
    If Not m_ready Then Exit Sub

    ' --- hasTab: whether there is an active tab (9.12c) ---
    '   At 0 tabs the reload button is also disabled. Back/forward are automatically disabled because canBack/canForward
    '   are False to begin with, but reload is independent of them,
    '   so it must be conveyed explicitly.
    Dim hasTab As Boolean
    If Not m_browser Is Nothing Then
        hasTab = (m_browser.TabCount > 0 And m_browser.ActiveIndex >= 1)
    End If

    ' about:blank is the state of "nothing loaded yet", so it is not shown in the URL field
    ' (it comes briefly right after Controller creation. Spec fact 6)
    Dim shown As String
    If url = "about:blank" Then
        shown = ""
    Else
        shown = url
    End If

    Dim json As String
    json = "{""type"":""navsync"",""url"":""" & JsonEscape(shown) & """" & _
           ",""canBack"":" & JsonBool(canBack) & _
           ",""canForward"":" & JsonBool(canFwd) & _
           ",""hasTab"":" & JsonBool(hasTab) & "}"

    Dim hr As Long
    hr = m_barPane.View_PostWebMessageAsString(json)
    If hr <> 0 Then
        Debug.Print "Wv2NavBar.SendNavSync: PostWebMessageAsString failed hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2NavBar.SendNavSync: send OK"
        Debug.Print "  json=" & json
    End If
End Sub


' ============================================================
' WaitPaneReady ? Wait for Controller creation to complete
'   *This is inside Init (= outside an event handler), so a synchronous wait is fine.*
'   Doing the same thing inside an event handler deadlocks (spec fact 14 / design principle 63).
' ============================================================
Private Function WaitPaneReady(ByVal pane As Wv2Pane) As Boolean
    Dim t As Single
    t = Timer
    Do While (Not pane.IsReady) And (Not pane.IsFailed) And (Wv2Log.SecondsSince(t) < READY_TIMEOUT)
        DoEvents
    Loop
    WaitPaneReady = pane.IsReady
End Function


' ============================================================
' WriteNavBarHtml ? Writes navbar.html in UTF-8 (no BOM)
'   In preparation for Japanese and future dynamic insertion, the same ADODB.Stream method as Wv2TabBar.
' ============================================================
Private Function WriteNavBarHtml() As Boolean
    On Error GoTo eh

    ' Create the folder if it does not exist
    If Dir(m_folderPath, vbDirectory) = "" Then
        MkDir m_folderPath
        Debug.Print "Wv2NavBar.WriteNavBarHtml: creating folder " & m_folderPath
    End If

    Dim fullPath As String
    fullPath = m_folderPath
    If Right$(fullPath, 1) <> "\" Then fullPath = fullPath & "\"
    fullPath = fullPath & NAVBAR_FILE

    Dim html As String
    html = BuildNavBarHtml()

    ' Write UTF-8 (no BOM) with ADODB.Stream
    Dim bin As Object
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 2                ' adTypeText
    bin.Charset = "UTF-8"
    bin.Open
    bin.WriteText html
    ' Write as text, then re-read as binary and discard the first 3 bytes (BOM)
    bin.Position = 0
    bin.Type = 1                ' adTypeBinary
    bin.Position = 3            ' Skip EF BB BF
    Dim bytes() As Byte
    bytes = bin.Read
    bin.Close

    Dim outStream As Object
    Set outStream = CreateObject("ADODB.Stream")
    outStream.Type = 1          ' adTypeBinary
    outStream.Open
    outStream.Write bytes
    outStream.SaveToFile fullPath, 2   ' 2 = adSaveCreateOverWrite
    outStream.Close

    Debug.Print "Wv2NavBar.WriteNavBarHtml: write OK " & fullPath & " (" & Len(html) & " chars, UTF-8)"
    WriteNavBarHtml = True
    Exit Function
eh:
    Debug.Print "Wv2NavBar.WriteNavBarHtml: error " & Err.Number & " " & Err.Description
    WriteNavBarHtml = False
End Function


' ============================================================
' BuildNavBarHtml ? Builds the nav bar HTML (9.12b is static)
'
'   At 9.12b it is "look only". The buttons cannot be pressed and the URL field is not updated.
'     - Back / forward / reload buttons (back and forward are shown disabled)
'     - URL input field (placeholder only)
'   In 9.12c it was made to receive navsync and reflect the URL and button enabled states.
'   Sending of clicks/input is 9.12d (for now still a commented-out vessel).
'
'   *Protection while editing the URL field* If a sync runs in the background while the user is typing a URL,
'     the string being typed disappears. When document.activeElement is the URL field,
'     value is not rewritten.
'
'   *Button symbols* To avoid environment dependence, written with HTML numeric character references
'     (&#8592; = left arrow / &#8594; = right arrow / &#10227; = circular arrow).
' ============================================================
Private Function BuildNavBarHtml() As String
    Dim s As String
    s = "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    s = s & "<style>" & vbLf
    s = s & "  html,body{margin:0;padding:0;height:100%;box-sizing:border-box;" & _
            "font-family:'Segoe UI','Meiryo',sans-serif;}" & vbLf
    s = s & "  *{box-sizing:border-box;}" & vbLf
    s = s & "  body{display:flex;align-items:center;gap:4px;padding:0 6px;" & _
            "overflow:hidden;background:#fafafa;border-bottom:1px solid #ccc;}" & vbLf
    s = s & "  button{width:30px;height:26px;border:1px solid transparent;" & _
            "border-radius:4px;background:transparent;cursor:pointer;" & _
            "font-size:15px;color:#333;line-height:1;padding:0;flex:none;}" & vbLf
    s = s & "  button:hover:not(:disabled){background:#e6e6e6;}" & vbLf
    s = s & "  button:disabled{color:#bbb;cursor:default;}" & vbLf
    s = s & "  #url{flex:1;height:26px;padding:0 10px;border:1px solid #ccc;" & _
            "border-radius:13px;background:#fff;font-size:13px;color:#222;" & _
            "outline:none;min-width:0;}" & vbLf
    s = s & "  #url:focus{border-color:#0078d4;box-shadow:0 0 0 1px #0078d4;}" & vbLf
    s = s & "  #url:disabled{background:#f0f0f0;color:#999;}" & vbLf
    s = s & "</style></head><body>" & vbLf
    s = s & "  <button id=""back""    title=""Back""       disabled>&#8592;</button>" & vbLf
    s = s & "  <button id=""forward"" title=""Forward""       disabled>&#8594;</button>" & vbLf
    s = s & "  <button id=""reload""  title=""Reload"" disabled>&#10227;</button>" & vbLf
    s = s & "  <input  id=""url"" type=""text"" placeholder=""Enter URL"" spellcheck=""false"" disabled>" & vbLf
    s = s & "<script>" & vbLf
    s = s & "  var W = window.chrome && window.chrome.webview;" & vbLf
    s = s & "  function postCmd(obj){ if(W){ W.postMessage(JSON.stringify(obj)); } }" & vbLf
    ' --- 9.24: hostObjects first + postMessage fallback (same form as TabBar 9.21) ---
    '   HO() returns hostObjects.sync if it can be obtained, otherwise null (same form as the settings tab's
    '   null guard). navHost(msgObj, callHost) is the three-tier structure common to the 4 commands.
    '   The JSON sent by the fallback is exactly the same as 9.12d (the thin VBA-side Case
    '   takes it as is).
    s = s & "  function HO(){ return (W && W.hostObjects && W.hostObjects.sync) ? W.hostObjects.sync : null; }" & vbLf
    s = s & "  function navHost(msgObj, callHost){" & vbLf
    s = s & "    var ho = HO();" & vbLf
    s = s & "    if(ho && ho.NavBar){" & vbLf
    s = s & "      try{ callHost(ho.NavBar); console.log('[host] '+msgObj.cmd); return; }" & vbLf
    s = s & "      catch(e){ console.log('[host-fail->postmsg] '+e); postCmd(msgObj); return; }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    console.log('[postmsg] '+msgObj.cmd); postCmd(msgObj);" & vbLf
    s = s & "  }" & vbLf
    ' *Declare the editing flag before all handlers (do not rely on hoisting)*
    s = s & "  var editing = false;" & vbLf
    ' *lastUrl (9.12d)* The "URL that should be shown now" conveyed by the latest navsync.
    '   When Escape discards the edit, restore immediately from here without asking VBA.
    '   It is updated at every navsync so it always holds the correct value (a cache for display).
    s = s & "  var lastUrl = '';" & vbLf
    ' --- The 3 kinds of buttons (enabled in 9.12d -> hostObjects first in 9.24) ---
    s = s & "  document.getElementById('back').onclick = function(){ navHost({cmd:'back'}, function(nb){ nb.HostBack(); }); };" & vbLf
    s = s & "  document.getElementById('forward').onclick = function(){ navHost({cmd:'forward'}, function(nb){ nb.HostForward(); }); };" & vbLf
    s = s & "  document.getElementById('reload').onclick = function(){ navHost({cmd:'reload'}, function(nb){ nb.HostReload(); }); };" & vbLf
    ' --- URL field: Enter navigates, Escape discards the edit (9.12d) ---
    '   *Set editing = false on Enter* When navigation starts, navsync replaces the URL field with the normalized
    '   URL. If editing is not dropped it stays frozen with the typed string.
    '   Scheme completion is done by the VBA side (NormalizeUrl), so here the value is sent as is.
    s = s & "  document.getElementById('url').onkeydown = function(ev){" & vbLf
    s = s & "    if(ev.key === 'Enter'){" & vbLf
    s = s & "      var v = this.value;" & vbLf
    s = s & "      editing = false;" & vbLf
    s = s & "      if(v && v.length){ navHost({cmd:'navigate', url:v}, function(nb){ nb.HostNavigate(v); }); }" & vbLf
    s = s & "    } else if(ev.key === 'Escape'){" & vbLf
    s = s & "      editing = false;" & vbLf
    s = s & "      this.value = lastUrl;   /* edit discarded -> restore immediately to the URL that should be shown now */" & vbLf
    s = s & "      this.blur();" & vbLf
    s = s & "    }" & vbLf
    s = s & "  };" & vbLf
    ' --- 9.12c: receive navsync from VBA and update the URL field and button enabled states ---
    ' *Editing flag (9.12c fix)*
    '   At first this was judged with document.activeElement !== u, but in WebView2's
    '   standalone document #url sometimes stays as activeElement and does not come back, so
    '   the URL field was never updated again (found on the real machine).
    '   Holding "whether the user has actually started editing" in an explicit flag is more reliable.
    '     - Merely focusing does not make it editing (so it does not freeze just from a click)
    '     - When key input (the input event) arrives, make it editing
    '     - Release editing on blur / Enter / Escape
    s = s & "  function applyNavSync(st){" & vbLf
    s = s & "    var u = document.getElementById('url');" & vbLf
    ' *lastUrl is always updated even while editing* So that even if a navigation happens in the background during editing,
    ' pressing Escape can return to "the real current URL".
    s = s & "    lastUrl = st.url || '';" & vbLf
    s = s & "    var b = document.getElementById('back');" & vbLf
    s = s & "    var f = document.getElementById('forward');" & vbLf
    s = s & "    var r = document.getElementById('reload');" & vbLf
    ' --- Buttons may always be synced (they do not steal focus) ---
    s = s & "    b.disabled = !st.canBack;" & vbLf
    s = s & "    f.disabled = !st.canForward;" & vbLf
    s = s & "    r.disabled = !st.hasTab;" & vbLf
    ' --- *If the URL field is being edited, do not touch it at all (neither value nor disabled)* ---
    '   ** Bug found on the real machine ** Assigning disabled makes the input blur, and
    '   the blur handler sets editing = false, so at the next sync the string being typed
    '   disappears (occurred the moment a new tab opened by a link click).
    '   The countermeasure is two-tiered:
    '     (1) While editing, do not touch the URL field (neither value nor disabled)
    '     (2) Assign disabled only when the value changes (prevents useless blur)
    s = s & "    if(!editing){" & vbLf
    s = s & "      var want = !st.hasTab;" & vbLf
    s = s & "      if(u.disabled !== want){ u.disabled = want; }" & vbLf
    s = s & "      u.value = st.url || '';" & vbLf
    s = s & "    }" & vbLf
    s = s & "  }" & vbLf
    s = s & "  (function(){" & vbLf
    s = s & "    var u = document.getElementById('url');" & vbLf
    s = s & "    u.addEventListener('input', function(){ editing = true; });" & vbLf
    s = s & "    u.addEventListener('blur',  function(){ editing = false; });" & vbLf
    s = s & "  })();" & vbLf
    s = s & "  if(W){" & vbLf
    s = s & "    W.addEventListener('message', function(ev){" & vbLf
    s = s & "      var data = ev.data;" & vbLf
    s = s & "      if(typeof data === 'string'){ try{ data = JSON.parse(data); }catch(e){ return; } }" & vbLf
    s = s & "      if(data && data.type === 'navsync'){ applyNavSync(data); }" & vbLf
    s = s & "    });" & vbLf
    s = s & "  }" & vbLf
    ' --- ready handshake: send ready after the listener is registered (design principle 61) ---
    s = s & "  postCmd({cmd:'ready'});" & vbLf
    s = s & "</script>" & vbLf
    s = s & "</body></html>"
    BuildNavBarHtml = s
End Function

