VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   8850.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12375
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'UserForm UserForm1.frm
'
' English fork v1.1 (StartBrowser + TearDown)
'   - Added StartBrowser(startUrl, withBars): one tab, optional tab/nav bars. It is
'     what Wv2.OpenBrowser calls; macros should go through the Wv2 module.
'   - The release sequence moved from UserForm_Terminate into TearDown, and every
'     failed start now calls it, so a failed start can simply be retried.
'
' Stage 9.18 (settings tab + search engine UI via AddHostObjectToScript)
'   -- As a continuation of aligning the entries of the main body in 9.17 (UseSearchEngine / CurrentBrowser),
'      a ? settings button is attached to the tab bar, and clicking it opens the search engine selection screen in a dedicated tab.
'      For the JS->VBA link we use AddHostObjectToScript rather than postMessage commands,
'      letting the JS of the settings screen call VBA (Wv2SettingsBridge) directly.
'
'   - The changes on the UserForm1 side are only 2 points:
'       (1) The search engine preset constants (SE_*) and name resolution are moved to Wv2SettingsBridge
'           (DRY). It shares the same resolution path as the settings screen.
'       (2) UseSearchEngine is rewritten as create a temporary bridge -> delegate to SetEngine.
'           The behavior from the Immediate window (UserForm1.UseSearchEngine "bing") is the same as 9.17.
'
'   - Creation of the settings tab, bridge attach and the settings screen HTML are the responsibility of Wv2Browser
'       (OpenSettingsTab / BuildSettingsHtml). TabBar only conveys the ? click to Browser with the settings
'       command (loose coupling kept). CurrentBrowser is as in 9.17.
'
' Stage 9.17 (search engine switching entry + exposing the running Browser reference)
'   -- Collects the open item that was put on hold in 9.16 as "build only the main-body Property (SearchEngine), and split
'      the UI and real-machine check into another stage". No UI parts (dropdown etc.) are placed yet;
'      we only grow "a foothold to touch the running Browser from outside" and "a minimal entry to switch by preset name"
'      (Plan A-1 + B-1). GUI integration (selection UI in NavBar) is the next stage.
'
'   - Added Public Property Get CurrentBrowser() As Wv2Browser (Plan A-1).
'       m_browser is Private WithEvents and cannot be touched from outside, so exactly one entry that returns
'       a plain reference is exposed. WithEvents is a declaration constraint on this side, and
'       handing out the reference through a Get is not a problem. From now on it is the foothold for "grabbing the running Browser from outside" in any extension.
'       While not started (m_browser Is Nothing) it returns Nothing.
'
'   - Added the search engine preset constants (SE_GOOGLE / SE_BING / SE_DUCKDUCKGO /
'       SE_YAHOO) at the top of the module. All are "query prefix type templates"
'       (just concatenate the encoded search term at the end, the same form as Plan 1-A of 9.16).
'
'   - Added Public Sub UseSearchEngine(ByVal engineName As String) (Plan B-1).
'       Resolves "google"/"bing"/"duckduckgo"/"yahoo" (case and leading/trailing whitespace ignored) to
'       the preset's template and does Let to CurrentBrowser.SearchEngine.
'       An unknown name falls to the default (Google) with a warning by Debug.Print (the empty guard is
'       guaranteed by the Browser-side Property Let). If Browser is not started, it does nothing and gives a notice.
'       *It is a logic entry, not a UI.* On the real machine, after starting StartWebView2_Full, just type
'       UserForm1.UseSearchEngine "bing" in the Immediate window to switch.
'
'   - The point where real-machine checks were stuck on "there is no entry for the reference" (the proviso in the 9.16 Help)
'       is resolved with CurrentBrowser. The 9.17 verification is gathered in Wv2Tests.Test_9_17_*.
'
' Stage 9.13 (tidying: sharing the JSON helpers + removal of verification code)
'   - Removed all of the throwaway verification Subs accumulated in 9.11a?9.12d (35 of them).
'     Now that the tab bar and nav bar provide all functions through the GUI, the role of the verification Subs
'     is over. They can be regenerated if needed, and the record remains in the dev notes.
'   - The old startup Subs were also removed, narrowing the practical startup entries to 2:
'       StartWebView2_Full ... 3 tabs + tab bar + nav bar (the normal startup entry)
'       StartWebView2_Spa  ... SPA template (the one decided in 9.10e to "leave just one")
'     Removed: StartWebView2 / StartWebView2_WithTabBar / StartWebView2_WithTabBar_3
'   - Line count: 1456 -> 700 lines
'
'   Note: on the Wv2TabBar / Wv2NavBar side the duplicated JSON helpers (JsonEscape / GetJsonStr /
'     GetJsonNum / BoolToJson) were shared in Wv2Json.bas.
'     The UserForm itself does not touch JSON, so it is unaffected.
'
' Stage 9.12d (JS->VBA: enabling input and clicks) ? 9.12 complete
'   - Code changes are for verification only (the main-body logic is on the Wv2NavBar side).
'   - Added Test_9_12d_Help / Test_9_12d_Inject*. Startup reuses StartWebView2_Full.
'
' Stage 9.12c (VBA->JS: display of the URL and button states)
'   - Code changes are for verification only (the main-body logic is on the Wv2NavBar side).
'   - Added Test_9_12c_Help. Startup reuses StartWebView2_Full.
'
' Stage 9.12b (skeleton of Wv2NavBar + 3-way layout split)
'   - Added the m_navBar field. Treated the same as Wv2TabBar (creation, Resize, Shutdown).
'   - UserForm_Resize becomes a 3-way split:
'       1st row = tab bar   (height TABBAR_HEIGHT_PX)
'       2nd row = nav bar   (height NAVBAR_HEIGHT_PX) *added in 9.12b
'       3rd row = content (all the rest)
'     *The split calculation is the responsibility of here (the UserForm)* Browser, TabBar and NavBar
'     do not know each other. The UserForm goes around handing out rectangles (loose coupling).
'   - Added StartWebView2_Full (3 tabs + tab bar + nav bar).
'   - In 9.12b the nav bar is static display only (buttons cannot be pressed, and the URL field is not updated).
'
' Stage 9.12a (foundation for the address bar/nav buttons ? no UI, Immediate window verification)
'   - The UserForm also subscribes to Wv2Browser's new event TabUrlChanged and logs it.
'   - Verify the active-tab delegation methods (NavigateActive / GoBackActive / GoForwardActive /
'     ReloadActive / StopActive / GetActiveUrl etc.) from the Immediate window.
'   - In 9.12b Wv2NavBar (the WebView2-based nav bar) is mounted. This is its foundation.
'
' Stage 9.11d (verification of enabling JS->VBA clicks)
'   - Code changes are for verification only (the main-body logic is on the Wv2TabBar side).
'   - Test_9_11d_Help: shows the procedure for manual click verification.
'   - Test_9_11d_Inject*: pours command strings directly into OnPaneWebMessage without going through JS.
'     Only the parser (GetJsonStr/GetJsonNum) and command handling can be separated
'     and verified (can be hit from the Immediate window without click operations).
'   - To reuse the tab bar, StartWebView2_WithTabBar_3 is also used in 9.11d.
'
' Stage 9.11c (verification of VBA->JS tab list drawing)
'   - Code changes are for verification only (the main-body logic is on the Wv2TabBar / Wv2Pane side).
'   - Added StartWebView2_WithTabBar_3: 3 tabs of bing/example/msn + tab bar.
'     We see the tab bar JS draw the 3 tabs and each tab's title reflected.
'   - Added Test_9_11c_Help / Test_9_11c_AddTab / Test_9_11c_CloseActive.
'     Confirm that after AddTab/CloseTab the tab bar drawing follows via full sync.
'
' Stage 9.11b (integration of the WebView2-based tab bar Wv2TabBar)
'   - New field: m_tabBar As Wv2TabBar (tab bar UI. If not used it is simply not
'     created = loose coupling).
'   - UserForm_Resize supports a 2-way split: if m_tabBar exists, a tab bar band is secured at the top
'     (height TABBAR_HEIGHT_PX) and the content area is shifted below it.
'     The rectangle of the tab bar band goes to m_tabBar.Resize, the content area to m_browser.ResizeAll.
'     If there is no m_tabBar, the whole area goes to content as before (*Browser does not
'     know the tab bar*; the split calculation is the UserForm's responsibility).
'   - Added m_tabBar.Shutdown to UserForm_Terminate (before m_browser.Shutdown).
'   - For verification: added StartWebView2_WithTabBar / Test_9_11b_*.
'
' Stage 9.11a (receiving events for the loosely coupled tab bar)
'   - Made m_browser WithEvents and added handlers that receive the 4 events of Wv2Browser (TabAdded /
'     TabClosed / ActiveChanged / TabTitleChanged).
'   - For verification: added the Test_9_11a_* group (Setup / CloseActiveMiddle / CloseActiveLast /
'     CloseBeforeActive / CloseAll / Help) and the internal helper DumpState.
'
' Stage 9.10e (NewWindowRequested to new tab)
'   - Added m_browser.Shutdown to UserForm_Terminate (cut the circular reference Browser?Pane
'     and then release. Plan Y. The implementation cost on the UserForm side is just this 1 line, design principle 58)
'   - 9.10 tidying: removed the verification Sub group (SpaSmoke / SpaMapping / SpaFull etc.).
'     SPA startup is consolidated in StartWebView2_Spa (template), and the m_env/m_pane fields were also removed
'
' Stage 9.9c (batch handling of the supplementary group)
'   - New Public Sub: OnExecuteScriptResult
'     The receiving end of result notifications called directly from Wv2Pane.OnExecuteScriptCompleted.
'     For now the first implementation is log output only (it can be extended to reflect in controls on the Form
'     when SPA integration happens in the future).
'   - Added a Nothing-check warning to the GetActivePane property
'     If m_pane is Nothing, a warning is given by Debug.Print and then Nothing is returned
'     (prevents getting lost in the "Object variable not set" error during Immediate
'      verification, Plan G1-a).
'
' Stage 9.9b (SPA support)
'   - New Public Sub: StartWebView2_SpaTest
'     A creation chain that displays the SPA test page (BuildSpaTestHtml).
'     Registers 7 kinds of handlers:
'       NavigationStarting / NavigationCompleted / DocumentTitleChanged /
'       WebMessageReceived / NewWindowRequested /
'       HistoryChanged (9.9b base) / DOMContentLoaded (9.9b derived _2)
'   - StartWebView2 / StartWebView2_LocalHtml have no main-body change
'   - Real-machine verification is called from Test_SpaTest_Real in the standard module
'
' Stage 9.9a (foundation for extracting ICoreWebView2_2)
'   - The UserForm side is untouched. Verification was done from the Immediate window with
'     ?UserForm1.GetActivePane.View2_Ensure
'
' Stage 9.8b (exhaustive View base IF + Settings group + Ctrl_Close)
'   - New Public property: GetActivePane (for Immediate window verification, exposes m_pane)
'   - The bodies of StartWebView2 / StartWebView2_LocalHtml are unchanged
'   - For how to hit the View_/Settings_/Ctrl_ methods added on the Wv2Pane side, see
'     the comment of the GetActivePane property (in this file)
'
' Stage 9.8a (class names formally fixed, rename only, no logic change)
'   - Module variable m_ctrl -> renamed to m_pane (consistent with the type change to Wv2Pane)
'   - Type changed: As Class8 -> As Wv2Environment / As Class9 -> As Wv2Pane
'   - Changed New Class8 / New Class9 -> New Wv2Environment / New Wv2Pane
'   - References to Class7 / Class8 / Class9 in the header and body were replaced with
'     ComCallbackHandler / Wv2Environment / Wv2Pane
'   - The old header comments (the rename list from 9.4c -> 9.6a, the explanation of the role
'     change at the TabWebView2 migration, etc.) were deleted from this header; for details see docs/dev-notes-*.md
'
' Mechanisms already added in stage 9.7b (continued):
'   - Module variables: hWnd_Form, m_gap{Left,Right,Top,Bottom}, m_resizeReady
'   - UserForm_Initialize: grant WS_THICKFRAME + save the gaps + raise the flag
'   - UserForm_Resize: resize Frame1 in a way that reproduces the gaps, then
'     reissue Ctrl_PutBounds so the WebView2 rendering area follows
'   - The assignments to Me.Width / Me.Height have been abolished (design principle 36)
'
' Design principles established in 9.7b (restated, see docs/dev-notes-*.md):
'   - Design principle 34: UserForm resize following uses the "save the gaps and reproduce them" method
'   - Design principle 35: beware the fake events of UserForm_Resize during Initialize
'                  (return early with the m_resizeReady flag)
'   - Design principle 36: decide the UserForm size in the designer, and do not assign to
'                  Me.Width / Me.Height in Initialize
'   - Design principle 37: WebView2 resize following works smoothly just by issuing
'                  Ctrl_PutBounds from UserForm_Resize (left to DWM)
'
' Manual setup in the VBE:
'   1. Add one UserForm to the project (name it UserForm1)
'   2. Drop one Frame control from the toolbox (name it Frame1)
'      (any size is fine; the value decided in the designer is the basis of the gap calculation)
'   3. Paste the following code into the UserForm's code module
'
' Role of the UserForm:
'   - Holds the window handle of Frame1 in hWnd_Frame
'   - Holds the UserForm's own window handle in hWnd_Form (for resizing)
'   - Holds the gaps on all four sides between Frame1 and the UserForm in m_gap{L,R,T,B}
'   - The creation chain for Bing display in StartWebView2 (4 kinds of handlers:
'     NavigationStarting / NavigationCompleted / DocumentTitleChanged /
'     NewWindowRequested)
'   - The creation chain for test HTML display in StartWebView2_LocalHtml (5 kinds:
'     the above + WebMessageReceived)
'   - The creation chain for SPA testing in StartWebView2_SpaTest (7 kinds:
'     the above 5 kinds + HistoryChanged + DOMContentLoaded, 9.9b)
'   - SendMessageToJS lets you fire VBA -> JS sending from outside
'   - The IsPageLoaded property observes NavigationCompleted
'   - UserForm_Terminate sets m_pane -> m_env to Nothing and
'     starts the ComRelease chain (the remove_ of all persistent handlers also runs
'     inside Wv2Pane.Class_Terminate)
'
' Future role (after the TabWebView2 class is introduced, 9.10 onward):
'   - m_env / m_pane are consolidated into a single m_tab, reducing the UserForm's fields
'   - IsPageLoaded moves to m_tab.IsPageLoaded
'   - StartWebView2 / StartWebView2_LocalHtml / StartWebView2_SpaTest are replaced by
'     m_tab.Init / m_tab.LoadHtml etc.


Option Explicit

' The target window handle on which WebView2 is placed
Private hWnd_Frame As LongPtr

' --- Added in stage 9.7b (for mouse resizing) ---
'   The UserForm's own HWND (obtained by tracing up from Frame1.HWND with GA_ROOT)
'   - The target to which WS_THICKFRAME is granted to make mouse resizing possible
'   - There is no need to reissue SetWindowLongPtrW inside the normal UserForm_Resize, so
'     it is obtained once at Initialize time and just held
Private hWnd_Form    As LongPtr

' --- Added in stage 9.7b (gaps between the UserForm and Frame1, design principle 34) ---
'   The "gaps on all four sides between Frame1 and the UserForm" in the design state at Initialize time are saved
'   in point units. By resizing Frame1 in a way that reproduces them in UserForm_Resize,
'   the positions of controls other than Frame1 (planned to be added in the future)
'   are not destroyed.
'
'   The point units are to match the units of Frame1.Left/Top/Width/Height and Me.InsideWidth/Height
'   (= the MSForms standard coordinate system). WebView2 ICoreWebView2Controller's
'   put_Bounds is in pixels, but that is re-obtained with GetClientRect (pixels),
'   so no conversion formula appears.
Private m_gapLeft    As Single
Private m_gapRight   As Single
Private m_gapTop     As Single
Private m_gapBottom  As Single

' --- Added in stage 9.7b (countermeasure for fake Resize events during Initialize, design principle 35) ---
'   The VBA UserForm fires a Resize event when Me.Width/Me.Height are set,
'   but at that point m_gap* is not yet determined. Return early with a flag.
Private m_resizeReady As Boolean
Private Const TABBAR_HEIGHT_PX As Long = 36   ' 9.11b: height of the tab bar band (px)
Private Const NAVBAR_HEIGHT_PX As Long = 34   ' 9.12b: height of the nav bar band (px)

' The Wv2Browser held by the UserForm (the comprehensive class that bundles all tabs, formally connected in 9.10c).
'   The main field that manages 1 Environment + multiple tabs (Wv2Pane).
'   (9.10 tidying: the old m_env / m_pane fields were removed and unified into m_browser)
Private WithEvents m_browser As Wv2Browser   ' 9.11a: made WithEvents to receive events
Attribute m_browser.VB_VarHelpID = -1
Private m_tabBar As Wv2TabBar                ' 9.11b: WebView2-based tab bar (optional)
Private m_navBar As Wv2NavBar   ' 9.12b: WebView2-based navigation bar

' --- 9.18: the search engine preset constants were moved to Wv2SettingsBridge ---
'   SE_GOOGLE / SE_BING / SE_DUCKDUCKGO / SE_YAHOO and the name resolution logic that
'   UserForm1 held until 9.17 were consolidated in Wv2SettingsBridge to share them with the settings screen
'   (DRY). UseSearchEngine creates a temporary bridge and delegates.


' ============================================================
' UserForm_Initialize
'   Sets the UserForm size and obtains and holds Frame1's HWND.
'   Premise: Frame1's size is matched at design time to fill the UserForm's client area
'   (a guideline of about Width=792, Height=572).
'   The creation chain is started by calling StartWebView2 separately.
'
'   Added in stage 9.7b:
'     - Obtain the UserForm's own HWND with GetAncestor(hWnd_Frame, GA_ROOT)
'     - Grant WS_THICKFRAME to make mouse resizing possible
'     - Save the gaps on all four sides between Frame1 and the UserForm
'     - Set the m_resizeReady flag to True and enable Resize events from here on
'   Note: assigning Me.Width / Me.Height fires a Resize event, but
'         m_resizeReady = False so UserForm_Resize returns early.
' ============================================================
Private Sub UserForm_Initialize()
    ' --- 9.7b fix: abolished the assignments to Me.Width / Me.Height ---
    '   The UserForm size at startup respects the value set at design time.
    '   If the size decided at design time is overwritten in Initialize,
    '   - code overwrites the designer = design-time consistency collapses
    '   - Frame1 stays as at design time so the gap calculation breaks
    '   - it contradicts design principle 34, "uphold the design-time layout"
    '   If you want to display at a different size at startup, adjust the UserForm size
    '   (and Frame1's size and position) together in the designer.

    hWnd_Frame = Frame1.[_GethWnd]
    Debug.Print "UserForm1.Initialize: Frame1.[_GethWnd] = " & hWnd_Frame

    ' --- 9.7b: obtain the UserForm's own HWND ---
    hWnd_Form = GetAncestor(hWnd_Frame, GA_ROOT)
    Debug.Print "UserForm1.Initialize: UserForm hWnd = " & hWnd_Form

    ' --- 9.7b: grant WS_THICKFRAME to make mouse resizing possible ---
    Dim style As LongPtr
    style = GetWindowLongPtrW(hWnd_Form, GWL_STYLE)
    SetWindowLongPtrW hWnd_Form, GWL_STYLE, style Or WS_THICKFRAME
    SetWindowPos hWnd_Form, 0, 0, 0, 0, 0, _
                 SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED
    Debug.Print "UserForm1.Initialize: WS_THICKFRAME granted (style was &H" & _
                Hex(style) & ")"

    ' --- 9.7b: save the gaps on all four sides between Frame1 and the UserForm (design principle 34) ---
    m_gapLeft = Frame1.Left
    m_gapTop = Frame1.Top
    m_gapRight = Me.InsideWidth - (Frame1.Left + Frame1.width)
    m_gapBottom = Me.InsideHeight - (Frame1.Top + Frame1.Height)
    Debug.Print "UserForm1.Initialize: gaps L=" & m_gapLeft & _
                " T=" & m_gapTop & _
                " R=" & m_gapRight & _
                " B=" & m_gapBottom

    ' --- 9.7b: enable Resize events from here on (design principle 35) ---
    m_resizeReady = True
End Sub


' ============================================================
' UserForm_Resize (new in stage 9.7b)
'
'   Called when the form is resized with the mouse.
'   Following design principle 34, Frame1 is resized in a way that reproduces the saved gaps,
'   and then the pixel-unit rc obtained with GetClientRect is passed to
'   Ctrl_PutBounds so the WebView2 rendering area follows.
'
'   Following design principle 35, fake Resize events during Initialize
'   return early with the m_resizeReady flag.
'
'   If the WebView2 controller (m_pane) is still being prepared (right after Initialize
'   before StartWebView2 is called, or when a resize event comes in the middle of StartWebView2),
'   only Frame1 follows and Ctrl_PutBounds is
'   skipped.
' ============================================================
Private Sub UserForm_Resize()
    If Not m_resizeReady Then Exit Sub

    ' --- Resize Frame1 "while keeping the gaps" ---
    Frame1.width = Me.InsideWidth - m_gapLeft - m_gapRight
    Frame1.Height = Me.InsideHeight - m_gapTop - m_gapBottom

    ' --- If there is no Wv2Browser yet, finish with only the Frame1 resize (9.10c-2) ---
    If m_browser Is Nothing Then Exit Sub

    ' --- Get Frame1's client area in pixels and reissue to all tabs ---
    '   ResizeAll reissues Ctrl_PutBounds to all tabs (visible or invisible)
    '   (point 8 a). By updating invisible tabs too, the size does not
    '   drift when they are activated later.
    Dim rc As RECT
    GetClientRect hWnd_Frame, rc

    ' --- 9.11b: if there is a tab bar, secure the top band and shift the content area down ---
    '   *Browser does not know the tab bar* The split is calculated here (UserForm), and
    '   the rectangle of the tab bar band goes to m_tabBar.Resize and the rectangle of the content area to
    '   m_browser.ResizeAll. If there is no tab bar, the whole area goes to content as before.
    ' --- 9.12b: stack bands from the top and make the rest content (3-way split) ---
    '   *The split calculation is the responsibility of here (the UserForm)* Browser / TabBar / NavBar
    '   do not know each other. The UserForm goes around handing out rectangles (loose coupling).
    '   If there is no bar, that row is skipped, so it works correctly in any configuration:
    '   tab bar only / nav bar only / neither.
    Dim y As Long
    y = rc.Top          ' The y coordinate at which to place the next band (stack from the top)

    ' --- 1st row: tab bar ---
    If Not m_tabBar Is Nothing Then
        Dim barRect As RECT
        barRect.Left = rc.Left
        barRect.Top = y
        barRect.Right = rc.Right
        barRect.Bottom = y + TABBAR_HEIGHT_PX
        If barRect.Bottom > rc.Bottom Then barRect.Bottom = rc.Bottom
        m_tabBar.Resize barRect
        y = barRect.Bottom
    End If

    ' --- 2nd row: nav bar (9.12b) ---
    If Not m_navBar Is Nothing Then
        Dim navRect As RECT
        navRect.Left = rc.Left
        navRect.Top = y
        navRect.Right = rc.Right
        navRect.Bottom = y + NAVBAR_HEIGHT_PX
        If navRect.Bottom > rc.Bottom Then navRect.Bottom = rc.Bottom
        m_navBar.Resize navRect
        y = navRect.Bottom
    End If

    ' --- 3rd row: content area (all the rest) ---
    Dim contentRect As RECT
    contentRect.Left = rc.Left
    contentRect.Top = y
    contentRect.Right = rc.Right
    contentRect.Bottom = rc.Bottom
    If contentRect.Top > contentRect.Bottom Then contentRect.Top = contentRect.Bottom
    m_browser.ResizeAll contentRect
End Sub
' ============================================================
' SendMessageToJS (new in stage 9.4c)
'
'   A public method for firing VBA -> JS string sending from outside (Test_MessageEcho_Real etc.).
'   A thin wrapper of Wv2Pane.View_PostWebMessageAsString.
'
'   It is desirable to confirm that loading is complete with IsPageLoaded before calling
'   (firing when the listener is not yet registered makes the message vanish).
' ============================================================
Public Sub SendMessageToJS(ByVal text As String)
    ' 9.10 tidying: from direct m_pane reference to via m_browser.ActivePane (GetActivePane).
    Dim pane As Wv2Pane
    Set pane = GetActivePane
    If pane Is Nothing Then
        Debug.Print "UserForm1.SendMessageToJS: no active Pane"
        Exit Sub
    End If
    If Not pane.IsReady Then
        Debug.Print "UserForm1.SendMessageToJS: Pane is not Ready"
        Exit Sub
    End If

    Dim hr As Long
    hr = pane.View_PostWebMessageAsString(text)
    If hr <> 0 Then
        Debug.Print "UserForm1.SendMessageToJS: failed hr=&H" & Hex(hr)
    Else
        Debug.Print "UserForm1.SendMessageToJS OK: """ & text & """"
    End If
End Sub


' ============================================================
' IsPageLoaded (new in stage 9.4c)
'
'   Returns whether the currently displayed page has progressed to NavigationCompleted.
'   The substance just delegates m_pane.IsPageLoaded (a thin wrapper).
'
'   Used by Test_MessageEcho_Real for the synchronization of
'   "wait for page loading to complete, then PostMessage".
'
'   Note on the provisional design:
'     It has meaning only for the first navigation completion, once. SPA support (tracking screen updates
'     by pushState etc.) will be redesigned when the comprehensive class is introduced in the future.
' ============================================================
Public Property Get IsPageLoaded() As Boolean
    ' 9.10 tidying: from direct m_pane reference to via GetActivePane.
    Dim pane As Wv2Pane
    Set pane = GetActivePane
    If pane Is Nothing Then
        IsPageLoaded = False
    Else
        IsPageLoaded = pane.IsPageLoaded
    End If
End Property


' ============================================================
' GetActivePane property (added in stage 9.8b, for Immediate verification, improved in 9.9c)
'
'   Returns a reference to the Wv2Pane instance the UserForm currently holds.
'   From the Immediate window you can hit Wv2Pane's methods directly as follows:
'
'     ?UserForm1.GetActivePane.View_GetSource()
'     ?UserForm1.GetActivePane.View_GetCanGoBack()
'     ?UserForm1.GetActivePane.View_GetBrowserProcessId()
'     UserForm1.GetActivePane.View_Reload
'     UserForm1.GetActivePane.View_Stop
'     UserForm1.GetActivePane.Settings_PutIsScriptEnabled False
'     UserForm1.GetActivePane.Settings_PutAreDefaultContextMenusEnabled False
'     UserForm1.GetActivePane.Settings_PutAreDevToolsEnabled False
'     UserForm1.GetActivePane.Ctrl_Close
'     UserForm1.GetActivePane.View_ExecuteScript "1+2"
'     UserForm1.GetActivePane.View_ExecuteScriptSafe "throw new Error('x')"  (9.9c)
'
'   9.9c improvement (Plan G1-a):
'     When m_pane is Nothing, a warning is now given by Debug.Print.
'     With this, during Immediate verification when you hit
'       ?UserForm1.GetActivePane.View_ExecuteScript "x"
'     instead of getting lost in the "Object variable not set" error,
'     a warning "m_pane is Nothing" now appears.
'     Nothing itself is still returned unchanged, so the path where the user does an Is Nothing check
'     is also maintained.
'
'   When TabWebView2 comes in, GetActivePane is planned to become a thin wrapper such as m_tab.ActivePane
'   (= meaning "get the active tab's Pane").
' ============================================================
Public Property Get GetActivePane() As Wv2Pane
    ' 9.10c-2: re-pointed from direct m_pane reference to via m_browser.ActivePane.
    '   It now has its original meaning of "get the active tab's Pane".
    If m_browser Is Nothing Then
        Debug.Print "[UserForm1.GetActivePane] m_browser is Nothing (the UserForm is not started, or " & _
                    "the StartWebView2 family has not been called)"
        Set GetActivePane = Nothing
        Exit Property
    End If
    Set GetActivePane = m_browser.ActivePane
End Property


' ============================================================
' OnExecuteScriptResult (new in stage 9.9c, receiving end of the second implementation of Plan alpha)
'
'   The notification entry for ExecuteScript results, called directly from Wv2Pane.OnExecuteScriptCompleted.
'   Wv2Pane does not hold a reference to UserForm1; it is called through the global
'   UserForm reference (UserForm1.OnExecuteScriptResult ...)
'   (no circular reference, Plan E1-a-i).
'
'   Arguments:
'     callbackId : the sequential ID returned by View_ExecuteScript
'     errorCode  : HRESULT (0 = S_OK)
'     resultJson : the JS return value as a JSON-encoded string
'                  Via View_ExecuteScriptSafe (9.9c F), structured JSON of the form {"ok":true,...} or
'                  {"ok":false,...} comes
'
'   Current implementation:
'     Log output only (coexists with the log inside Wv2Pane, Plan E2-X).
'     When SPA integration is taken up in earnest in the future, this can be extended to display the result
'     in controls on the Form (e.g. a TextBox or ListBox).
'
'   * Behavior on error *
'     Wv2Pane.OnExecuteScriptCompleted calls this Sub without error protection, so
'     an exception must not be thrown inside this Sub (= it would destroy the COM callback path).
'     Keep to logging only and avoid complex processing.
' ============================================================
Public Sub OnExecuteScriptResult( _
    ByVal callbackId As Long, _
    ByVal errorCode As Long, _
    ByVal resultJson As String)

    Debug.Print "UserForm1.OnExecuteScriptResult: callbackId=" & callbackId & _
                ", errorCode=" & errorCode & _
                ", resultJson=" & resultJson
End Sub


' ============================================================
' FrameHwnd (stage 9.10b verification-only accessor)
'   Exposes Frame1's window handle (hWnd_Frame) to the outside.
'   Used as the parent-frame HWND passed to Wv2Browser.Init.
'   Note that it stays 0 until after UserForm_Initialize has run (= after Show).
' ============================================================
Public Property Get FrameHwnd() As LongPtr
    FrameHwnd = hWnd_Frame
End Property


' ============================================================
' BuildSpaAppHtml (9.10d-2a, returns the minimal SPA HTML)
'   ASCII only (no Japanese included) so that
'   even writing out with VBA's Print # (system locale) does not garble. A real page on which
'   DOMContentLoaded reliably fires. pushState etc. are controlled by firing ExecuteScript as in d-1,
'   so no buttons are placed on the HTML side (point C: minimal configuration).
' ============================================================
Public Function BuildSpaAppHtml() As String
    Dim s As String
    s = "<!DOCTYPE html>" & vbCrLf
    s = s & "<html><head><meta charset=""utf-8""><title>WV2 SPA d-2</title></head>" & vbCrLf
    s = s & "<body style=""font-family:sans-serif"">" & vbCrLf
    s = s & "<h1>WV2 SPA (appassets.example)</h1>" & vbCrLf
    s = s & "<p id=""msg"">initial</p>" & vbCrLf
    s = s & "<script>" & vbCrLf
    s = s & "document.addEventListener('DOMContentLoaded',function(){" & vbCrLf
    s = s & "  document.getElementById('msg').textContent='DOMContentLoaded fired';" & vbCrLf
    s = s & "});" & vbCrLf
    s = s & "</script>" & vbCrLf
    s = s & "</body></html>" & vbCrLf
    BuildSpaAppHtml = s
End Function

' ============================================================
' WriteSpaAppFolder (9.10d-2a, folder creation + text file writing)
'   Create folderPath if it does not exist (the parent C:\Temp is assumed to exist).
'   Write content to folderPath\fileName (ASCII assumed, Print # is enough).
'   Return value: True on success.
' ============================================================
Public Function WriteSpaAppFolder( _
    ByVal folderPath As String, _
    ByVal fileName As String, _
    ByVal content As String) As Boolean

    On Error GoTo eh

    ' Create the folder if it does not exist
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
        Debug.Print "WriteSpaAppFolder: creating folder " & folderPath
    End If

    Dim fullPath As String
    fullPath = folderPath
    If Right$(fullPath, 1) <> "\" Then fullPath = fullPath & "\"
    fullPath = fullPath & fileName

    Dim fnum As Integer
    fnum = FreeFile
    Open fullPath For Output As #fnum
    Print #fnum, content
    Close #fnum

    Debug.Print "WriteSpaAppFolder: write OK " & fullPath & " (" & Len(content) & " chars)"
    WriteSpaAppFolder = True
    Exit Function
eh:
    Debug.Print "WriteSpaAppFolder: error " & Err.Number & " " & Err.Description
    WriteSpaAppFolder = False
End Function

' ============================================================
' StartWebView2_Spa (template for SPA startup)
'
'   Maps a local folder to https://appassets.example/ and opens the
'   home-made SPA (index.html) there. With the single AddTabWithUrlForSpa
'   SetMapping + Navigate + EnableSpaHandlers (HistoryChanged/DOMContentLoaded)
'   everything is done in one go (spec fact 13: specify the Navigate target down to the file name).
'
'   The startup entry for running an app that uses an SPA (history.pushState etc.) in WebView2.
'   The HTML to deliver is prepared with BuildSpaAppHtml/WriteSpaAppFolder
'   (in real operation, replace with any HTML in the folder).
'
'   Usage:
'     UserForm1.Show vbModeless
'     UserForm1.StartWebView2_Spa
'   The active tab's Pane can be obtained with ?UserForm1.GetActivePane.
' ============================================================
Public Sub StartWebView2_Spa()
    Debug.Print "UserForm1.StartWebView2_Spa start (SPA startup)"
    If AlreadyStarted("StartWebView2_Spa") Then Exit Sub

    ' --- 1. Prepare the SPA folder (write out the HTML) ---
    Dim folderPath As String
    folderPath = Wv2Log.AppFolder("SPA")
    If Not WriteSpaAppFolder(folderPath, "index.html", BuildSpaAppHtml()) Then
        Debug.Print "StartWebView2_Spa: failed to prepare the SPA folder"
        Exit Sub
    End If

    ' --- 2. Create Wv2Browser and initialize the Environment ---
    Set m_browser = New Wv2Browser
    If Not m_browser.Init(hWnd_Frame, Wv2Log.AppFolder("Profile")) Then
        Debug.Print "StartWebView2_Spa: Wv2Browser.Init failed (state=" & _
                    m_browser.state & ")"
        TearDown True
        Exit Sub
    End If
    Debug.Print "StartWebView2_Spa: Wv2Browser.Init OK"

    ' --- 3. Create the SPA tab in one go (AddTab->SetMapping->Navigate->EnableSpaHandlers) ---
    Dim pane As Wv2Pane
    Set pane = m_browser.AddTabWithUrlForSpa("appassets.example", folderPath, "index.html")
    If pane Is Nothing Then
        Debug.Print "StartWebView2_Spa: failed in AddTabWithUrlForSpa"
        Exit Sub
    End If
    Debug.Print "StartWebView2_Spa: SPA tab creation OK (ViewPtr=" & pane.ViewPtr & ")"
    Debug.Print "  Success if https://appassets.example/index.html is displayed."
End Sub

' ============================================================
' UserForm_Terminate
'   When the form is closed, release in the order Wv2Pane -> Wv2Environment.
'   The ComRelease chain runs in each class's Class_Terminate,
'   and the references of ICoreWebView2Controller / ICoreWebView2Environment are released.
' ============================================================
' ============================================================
' *9.11a event receiving handlers (the receivers of the loosely coupled tab bar's event entries)
'   They just receive Browser's RaiseEvent here and Debug.Print.
'   This will be replaced by the actual tab bar implementation (WebView2-based / MsForms-based).
'   *How to read the verification* The following logs appear in the Immediate window (Ctrl+G):
'     [EVT] TabAdded / [EVT] ActiveChanged / [EVT] TabClosed / [EVT] TabTitleChanged
' ============================================================
Private Sub m_browser_TabAdded(ByVal index As Long)
    Debug.Print "[EVT] TabAdded(index=" & index & ")"
End Sub

Private Sub m_browser_ActiveChanged(ByVal index As Long)
    Debug.Print "[EVT] ActiveChanged(index=" & index & ")"
End Sub

Private Sub m_browser_TabClosed(ByVal index As Long)
    Debug.Print "[EVT] TabClosed(index=" & index & ")"
End Sub

Private Sub m_browser_TabTitleChanged(ByVal index As Long, ByVal title As String)
    Debug.Print "[EVT] TabTitleChanged(index=" & index & ", title=" & title & ")"
End Sub

' --- 9.12a: changes in URL state (event for updating the address bar) ---
'   It is fired from both NavigationCompleted and HistoryChanged, so multiple arrivals
'   for one navigation are normal (harmless thanks to the full-sync method, design principle 60).
Private Sub m_browser_TabUrlChanged(ByVal index As Long, ByVal url As String, _
                                    ByVal canGoBack As Boolean, ByVal canGoForward As Boolean)
    Debug.Print "[EVT] TabUrlChanged(index=" & index & ", url=" & url & _
                ", back=" & canGoBack & ", fwd=" & canGoForward & ")"
End Sub
' --- Internal helper that dumps the current state ---
Private Sub DumpState(ByVal tag As String)
    If m_browser Is Nothing Then
        Debug.Print "  [state:" & tag & "] m_browser is Nothing"
        Exit Sub
    End If
    Debug.Print "  [state:" & tag & "] TabCount=" & m_browser.TabCount & _
                ", ActiveIndex=" & m_browser.ActiveIndex
End Sub


Private Sub UserForm_Terminate()
    Debug.Print "UserForm1.Terminate start"
    TearDown False      ' the frame window dies with the form and takes the WebViews with it
    Debug.Print "UserForm1.Terminate done"
End Sub

' ============================================================
' TearDown (English fork)
'   The release sequence that used to live in UserForm_Terminate, shared with
'   the start-up failure paths: a start that failed half-way used to leave
'   m_browser set, so every later start was refused by AlreadyStarted until
'   the form was closed, and a half-built tab bar stayed on screen.
'
'   closeViews = True also closes every WebView2 controller. Needed when the
'   form stays open (a failed start): otherwise the tabs created so far would
'   stay visible in the frame with nothing managing them. Unloading the form
'   passes False - destroying the frame window closes them, and Close costs
'   about 1.4 s per tab (K-4).
' ============================================================
Private Sub TearDown(ByVal closeViews As Boolean)
    ' --- 9.12b: drop the nav bar first (reverse order of creation) ---
    If Not m_navBar Is Nothing Then
        m_navBar.Shutdown
        Set m_navBar = Nothing
    End If

    ' 9.11b: release the tab bar first (cancel the WithEvents subscription).
    If Not m_tabBar Is Nothing Then
        m_tabBar.Shutdown
        Set m_tabBar = Nothing
    End If

    ' 9.10e: cut the circular reference (Browser ? Pane) and then release.
    '   Shutdown cuts each Pane's m_browser back reference. If Set m_browser = Nothing
    '   is done without calling it, the circular reference prevents release and leaks (design principle 58).
    If Not m_browser Is Nothing Then
        If closeViews Then m_browser.CloseAll
        m_browser.Shutdown
    End If

    ' Release Wv2Browser (internally all Panes -> Environment in that order, design principle 43)
    Set m_browser = Nothing
End Sub

' ============================================================
' StartBrowser (English fork) - the start-up entry used by Wv2.OpenBrowser
'   One tab showing startUrl. withBars adds the tab bar and the navigation
'   bar; without them the page fills the whole frame.
'   Returns False (and releases everything) when any step fails. Does not
'   wait for the page to load - Wv2.WaitLoaded does that.
' ============================================================
Public Function StartBrowser(ByVal startUrl As String, ByVal withBars As Boolean) As Boolean
    If AlreadyStarted("StartBrowser") Then Exit Function
    If hWnd_Frame = 0 Then
        Debug.Print "StartBrowser: the form has no window yet (Show it first)"
        Exit Function
    End If

    Set m_browser = New Wv2Browser
    If Not m_browser.Init(hWnd_Frame, Wv2Log.AppFolder("Profile")) Then
        Debug.Print "StartBrowser: Browser.Init failed (state=" & m_browser.state & ")"
        TearDown True
        Exit Function
    End If

    If m_browser.AddTabWithUrl(startUrl) Is Nothing Then
        Debug.Print "StartBrowser: the first tab could not be created"
        TearDown True
        Exit Function
    End If

    If withBars Then
        Set m_tabBar = New Wv2TabBar
        If Not m_tabBar.Init(m_browser, hWnd_Frame, Wv2Log.AppFolder("TabBar")) Then
            Debug.Print "StartBrowser: TabBar.Init failed"
            TearDown True
            Exit Function
        End If
        Set m_navBar = New Wv2NavBar
        If Not m_navBar.Init(m_browser, hWnd_Frame, Wv2Log.AppFolder("NavBar")) Then
            Debug.Print "StartBrowser: NavBar.Init failed"
            TearDown True
            Exit Function
        End If
    End If

    UserForm_Resize
    StartBrowser = True
End Function
' ============================================================
' AlreadyStarted (English fork)
'   Running a StartWebView2_* sub a second time used to replace m_browser /
'   m_tabBar / m_navBar without shutting the old ones down. Their reference
'   cycles (Browser <-> Pane, bar pane <-> sink) kept the old WebViews alive,
'   visible in the frame and still receiving events. Refuse instead: close
'   the form and Show it again to restart.
' ============================================================
Private Function AlreadyStarted(ByVal caller As String) As Boolean
    If Not m_browser Is Nothing Then
        Debug.Print caller & ": already started on this form - close the form and Show it again to restart"
        AlreadyStarted = True
    End If
End Function

' --- Start in the full configuration (3 tabs + tab bar + nav bar) ---
'   *The HTML delivery folders are separate for TabBar and NavBar*
'   The virtual host names are also separate (appassets.tabbar / appassets.navbar).
Public Sub StartWebView2_Full()
    Debug.Print "==== StartWebView2_Full start (3 tabs + tab bar + nav bar) ===="
    If AlreadyStarted("StartWebView2_Full") Then Exit Sub

    ' --- 1. Start Browser and add 3 tabs ---
    Set m_browser = New Wv2Browser
    If Not m_browser.Init(hWnd_Frame, Wv2Log.AppFolder("Profile")) Then
        Debug.Print "StartWebView2_Full: Browser.Init failed (state=" & m_browser.state & ")"
        TearDown True
        Exit Sub
    End If
    Dim p As Wv2Pane
    Set p = m_browser.AddTabWithUrl("https://www.bing.com")
    If p Is Nothing Then Debug.Print "  tab add failed (bing)": Exit Sub
    Set p = m_browser.AddTabWithUrl("https://example.com")
    If p Is Nothing Then Debug.Print "  tab add failed (example)": Exit Sub
    Set p = m_browser.AddTabWithUrl("https://www.msn.com")
    If p Is Nothing Then Debug.Print "  tab add failed (msn)": Exit Sub

    ' --- 2. Start the tab bar ---
    Set m_tabBar = New Wv2TabBar
    If Not m_tabBar.Init(m_browser, hWnd_Frame, Wv2Log.AppFolder("TabBar")) Then
        Debug.Print "StartWebView2_Full: TabBar.Init failed"
        TearDown True
        Exit Sub
    End If

    ' --- 3. Start the nav bar (9.12b, the delivery folder is separate from TabBar) ---
    Set m_navBar = New Wv2NavBar
    If Not m_navBar.Init(m_browser, hWnd_Frame, Wv2Log.AppFolder("NavBar")) Then
        Debug.Print "StartWebView2_Full: NavBar.Init failed"
        TearDown True
        Exit Sub
    End If

    ' --- 4. Recompute the layout (3-way split) ---
    UserForm_Resize

    Debug.Print "==== StartWebView2_Full done (TabCount=" & m_browser.TabCount & _
                ", ActiveIndex=" & m_browser.ActiveIndex & _
                ", TabBar.IsReady=" & m_tabBar.IsReady & _
                ", NavBar.IsReady=" & m_navBar.IsReady & ") ===="
    Debug.Print "  Success if the 3 rows (tab bar / nav bar / content) line up without overlapping."
End Sub

' ============================================================
' CurrentBrowser (9.17, Plan A-1)
'
'   Returns a read-only reference to the running Wv2Browser. m_browser is
'   Private WithEvents and cannot be touched from outside (Immediate window / search engine switching / future UI),
'   so exactly one entry that hands out a plain reference is exposed. WithEvents is a constraint on the
'   declaration side, and returning a reference with a Get is not a problem.
'
'   Before startup (StartWebView2_Full not run) m_browser is Nothing, so
'   it returns Nothing. The caller must check Is Nothing.
'
'   *Uses*
'     - In the Immediate window, grab it with  Set b = UserForm1.CurrentBrowser  and
'       switch the search engine on the real machine with b.SearchEngine = ... (the open item of 9.16).
'     - A common foothold for "touching the running Browser from outside" in all future extensions.
' ============================================================
Public Property Get CurrentBrowser() As Wv2Browser
    Set CurrentBrowser = m_browser
End Property


' ============================================================
' UseSearchEngine (9.17, Plan B-1)
'
'   A minimal entry for switching the search engine by preset name. Not a GUI part but
'   a logic entry (on the real machine, hit it from the Immediate window).
'
'   Argument engineName: "google" / "bing" / "duckduckgo" / "yahoo"
'     Case and leading/trailing whitespace are ignored. An unknown name falls to the default (Google)
'     with a warning. The final guard against empty string/whitespace only is guaranteed by the Browser-side Property Let
'     (falls back to Google if empty), so here we concentrate on name resolution.
'
'   If Browser is not started (CurrentBrowser Is Nothing), it does nothing and gives a notice.
'
'   Usage example (real machine, after StartWebView2_Full has been started):
'     UserForm1.UseSearchEngine "bing"
'     -> After that, typing a search phrase (e.g.  excel vba ) in the address bar shows Bing search results.
' ============================================================
Public Sub UseSearchEngine(ByVal engineName As String)
    If m_browser Is Nothing Then
        Debug.Print "UseSearchEngine: Browser is not started (run StartWebView2_Full first)."
        Exit Sub
    End If

    ' 9.18: the resolution of preset names is consolidated in Wv2SettingsBridge. Here a temporary
    '   bridge is created, Browser is bound, and it delegates to SetEngine. It passes through exactly the same
    '   name resolution path the settings screen uses, so the behavior matches between the Immediate window and the GUI.
    Dim bridge As Wv2SettingsBridge
    Set bridge = New Wv2SettingsBridge
    bridge.BindBrowser m_browser

    Dim applied As String
    applied = bridge.SetEngine(engineName)

    Debug.Print "UseSearchEngine: set the search engine to '" & applied & "' (" & _
                m_browser.Debug_SearchEngine() & ")"
End Sub


