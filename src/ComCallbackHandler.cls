VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComCallbackHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''
' --- ComCallbackHandler.cls Stage N-4 (response body) ---
'
'   Changes in N-4:
'     - Added Case HK_ResponseGetContent to Handler_Invoke.
'       **One-shot type**, same shape as HK_ExecuteScriptCompleted:
'         Invoke(HRESULT errorCode, IStream* result)
'       -> m_owner.OnResponseContentCompleted m_callbackId, errorCode, arg2
'
'   **The IStream is not read here**
'     Settings such as the maximum byte count are held by Wv2Pane,
'     so the pointer is passed on as is and Wv2Pane reads it.
'     **The IStream is valid only during Invoke**, so the Wv2Pane side must
'     read it all synchronously before returning.
'
'   **The request's serial number (seq) goes into CallbackId**
'     Invoke carries no identifier, so which response a body belongs to is told apart by CallbackId
'     (the mechanism of stage 9.8c is reused as is).
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- ComCallbackHandler.cls Stage N-3 (response status and headers) ---
'
'   Changes in N-3:
'     - Added Case HK_WebResourceResponseReceived to Handler_Invoke.
'       It is a 2-argument event (sender, args), so the same shape as the existing ones.
'       -> m_owner.View2_OnWebResourceResponseReceived arg1, arg2
'
'   **Naming uses the View2_ prefix** (design principle 31 / stage 9.9b convention)
'     This event belongs to ICoreWebView2_2 (a derived IF). add / remove are also called
'     on m_pView2. Do not mix it with the base HK_WebResourceRequested (View_).
'
'   **This branch also does not fire Debug.Print**
'     **This event has no filter**, so images and CSS all arrive.
'     It fires more often than N-1's WebResourceRequested. Just record and return immediately.
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- ComCallbackHandler.cls Stage N-1 (wiring of WebResourceRequested) ---
'
'   Changes in N-1:
'     - Added Case HK_WebResourceRequested to Handler_Invoke.
'       It is a 2-argument event (sender, args), so the same shape as the DocumentTitleChanged branch.
'       -> m_owner.View_OnWebResourceRequested arg1, arg2
'
'   **Only this branch does not fire Debug.Print (N-1 issue 5)**
'     WebResourceRequested fires tens to hundreds of times per page, and
'     **if the handler is slow the traffic itself jams** (it is an event for stopping / rewriting
'     requests, so the runtime advances the traffic only after Invoke returns).
'     The 4-line Debug.Print present in the other Cases is deliberately left out for this kind only.
'     Recording is done by the ring buffer on the Wv2Pane side, and retrieval is done later in a batch
'     with NetLogDrain (issue 6).
'
'   **Do not call EvalSync inside the handler** (the territory of spec facts 14 / 20).
'     Wv2Pane's in-callback guard works, so calling it would not hang,
'     but the design is such that it is not called in the first place.
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- ComCallbackHandler.cls Stage 9.9b (formerly Class7) ---
'
'   Changes in 9.9b (SPA support):
'     - Added 2 kinds of branches to Handler_Invoke:
'         Case HK_HistoryChanged   -> m_owner.View_OnHistoryChanged arg1, arg2
'         Case HK_DOMContentLoaded -> m_owner.View2_OnDOMContentLoaded arg1, arg2
'       Both are 2-argument events (sender, args). Same shape as the DocumentTitleChanged branch.
'     - Core logic (Handler_QueryInterface / AddRef / Release) unchanged
'       (kind is accessed dynamically, so the new kind=9/10 also pass automatically)
'
'   Changes in 9.9a: none (ComCallbackHandler was not touched)
'
'   Changes in 9.8c (continued):
'     - Added new module variable m_callbackId As Long
'       (for HK_ExecuteScriptCompleted, identifies concurrent ExecuteScript calls)
'     - Added new Friend Property: CallbackId (Get/Let; Wv2Pane.View_ExecuteScript sets it with Let
'       right after Acquire)
'     - Added Case HK_ExecuteScriptCompleted to Handler_Invoke (one-shot type,
'       converts the LPCWSTR result to a String with PtrToString and then notifies Wv2Pane)
'     - Notifies Wv2Pane.OnExecuteScriptCompleted(callbackId, errorCode, resultStr) via
'       dynamic binding (plan alpha, first implementation only logs)
'
'   Carried over from 9.8a/9.8b:
'     - Class name Class7 -> ComCallbackHandler (9.8a)
'     - No change to ComCallbackHandler in 9.8b
'
'   Role: the VBA-side wrapper for COM callbacks (via thunks).
'     An instance is created in a "1 event handler = 1 instance" relationship, and
'     the AddRef/Release/Invoke of the vtable that WebView2 holds are implemented in VBA with
'     m_refCount + m_owner dynamic-binding calls.
'
'   Implemented kinds:
'     HK_EnvironmentCompleted    (one-shot, plan D1-alpha)
'     HK_ControllerCompleted     (one-shot, plan D1-alpha)
'     HK_NavigationStarting      (persistent, plan P1)
'     HK_NavigationCompleted     (persistent, plan P1)
'     HK_WebMessageReceived      (persistent, plan P1)
'     HK_DocumentTitleChanged    (persistent, plan P1)
'     HK_NewWindowRequested      (persistent, plan P1)
'     HK_ExecuteScriptCompleted  (one-shot, plan a = concurrency OK, added in 9.8c)
'     HK_HistoryChanged          (persistent, added in 9.9b, base ICoreWebView2)
'     HK_DOMContentLoaded        (persistent, added in 9.9b, ICoreWebView2_2)
'     HK_WebResourceRequested    (persistent, added in N-1, base ICoreWebView2)
'     HK_WebResourceResponseReceived (persistent, added in N-3, ICoreWebView2_2)
'     HK_ResponseGetContent      (one-shot, added in N-4, concurrency OK)
'
'   Rule for interpreting arguments (design principle 11):
'     The (this, arg1, arg2) that ComCallbackHandler's Handler_Invoke receives always
'     matches WebView2's view of (C++ this, 1st argument, 2nd argument).
'     For the Invoke of a 2-argument event (sender, args):
'       this = handler obj (unused)
'       arg1 = sender (should be the same as m_pView)
'       arg2 = args (the args interface specific to each event)
'     For 2-argument one-shot types (Environment/Controller/ExecuteScript Completed):
'       this = handler obj (unused)
'       arg1 = errorCode (HRESULT, low 32 bits extracted with LongPtrLowDword)
'       arg2 = result pointer (pEnvironment/pController/LPCWSTR result)
'
'   Naming convention (design principle 31):
'     Event receivers of ICoreWebView2 take the form View_On<EventName>
'     (base IF) / View2_On<EventName> (derived IF, first introduced in 9.9b).
'     All are Public Subs in Wv2Pane, called from ComCallbackHandler.Handler_Invoke
'     via dynamic binding.
'
'   Important: m_owner.View_OnXxx cannot be called via dynamic binding unless it is a
'     **Public Sub**. All View_OnXxx / View2_OnXxx / OnXxxCompleted
'     methods of Wv2Pane must be Public.
''''''''''''''''''''''''''''''''''
Option Explicit

#If Win64 Then
    ' OK
#Else
    #Error "This class requires 64-bit VBA (x64)"
#End If

Private m_kind     As HandlerKind
Private m_owner    As Object
Private m_pSlot    As LongPtr
Private m_refCount As Long

' Added in 9.8c:
'   HK_ExecuteScriptCompleted allows concurrency (plan a). To identify each call,
'   the ComCallbackHandler instance holds a callback ID.
'   For kinds other than HK_ExecuteScriptCompleted it stays 0 and is not used.
Private m_callbackId As Long


' ============================================================
' Initialization (assumed to be called from the foundation's AcquireHandlerFor / Thunks_AcquireSlot)
'
'   refcount starts from 0 (plan X-1 adopted):
'     - COM semantics: once the handler is passed to WebView2, WebView2 itself
'       calls AddRef and takes responsibility. The VBA side holds no reference.
'     - It is confirmed on a real machine that WebView2 calls AddRef inside
'       CreateCoreWebView2EnvironmentWithOptions (refcount: 0 -> 1).
'     - After Invoke, WebView2 calls Release and refcount = 0, and
'       HandlerReleaseInternal's automatic release path runs.
' ============================================================
Friend Sub Init( _
    ByVal kind As HandlerKind, _
    ByVal owner As Object, _
    ByVal pSlot As LongPtr)

    m_kind = kind
    Set m_owner = owner
    m_pSlot = pSlot
    m_refCount = 0
End Sub


' ============================================================
' Property group (for operations from the foundation side)
' ============================================================
Friend Property Get RefCount() As Long
    RefCount = m_refCount
End Property

Friend Property Let RefCount(ByVal v As Long)
    m_refCount = v
End Property

Friend Property Get Slot() As LongPtr
    Slot = m_pSlot
End Property

Friend Property Get kind() As HandlerKind
    kind = m_kind
End Property


' Added in 9.8c: CallbackId property (used by HK_ExecuteScriptCompleted)
'   Wv2Pane.View_ExecuteScript assigns a serial ID for each call and
'   sets it with this Let right after AcquireHandlerFor.
'   It is then referenced as m_callbackId inside Handler_Invoke and
'   passed to Wv2Pane.OnExecuteScriptCompleted.
Friend Property Get callbackId() As Long
    callbackId = m_callbackId
End Property

Friend Property Let callbackId(ByVal v As Long)
    m_callbackId = v
End Property


' ============================================================
' ClearOwner
'   Called from the foundation side (HandlerReleaseInternal) when refcount = 0.
'   To cut the circular reference, sets m_owner to Nothing.
' ============================================================
Friend Sub ClearOwner()
    Set m_owner = Nothing
End Sub


' ============================================================
' Handler_Invoke (fixed at vTable slot 7)
'
'   The only entry that receives callbacks from WebView2.
'   The thunk injects ObjPtr(this) into Me and shifts the arguments by one,
'   so they reach this method.
'
'   Argument spec:
'     this : rdx after the thunk (= the obj WebView2 put in rcx)
'            not used for HK_EnvironmentCompleted (identification is done by m_kind)
'     arg1 : r8 after the thunk   (= the value WebView2 put in rdx)
'     arg2 : r9 after the thunk   (= the value WebView2 put in r8)
'
'   Interpretation per m_kind:
'     HK_EnvironmentCompleted: WebView2's view Invoke(this, errorCode, pEnv)
'                              -> this = handler obj, arg1 = errorCode (low 32 bits),
'                                 arg2 = pEnvironment
'     HK_ControllerCompleted:  WebView2's view Invoke(this, errorCode, pCtrl)
'                              -> this = handler obj, arg1 = errorCode (low 32 bits),
'                                 arg2 = pController
'     HK_NavigationStarting:   WebView2's view Invoke(this, sender, args)
'                              -> this = handler obj, arg1 = sender (=m_pView),
'                                 arg2 = args (= ICoreWebView2NavigationStartingEventArgs)
'     HK_NavigationCompleted:  WebView2's view Invoke(this, sender, args)
'                              -> same as above, args type is NavigationCompletedEventArgs
'     HK_WebMessageReceived:   same as above, args type is WebMessageReceivedEventArgs
'     HK_DocumentTitleChanged: same as above, args type is IUnknown* (no information)
'     HK_NewWindowRequested:   same as above, args type is
'                              ICoreWebView2NewWindowRequestedEventArgs
'                              (added in 9.6a)
'
'   Important (established in stage 9.4a):
'     The this argument of ComCallbackHandler receives the "handler obj address" for all kinds.
'     WebView2's "first argument" (the one placed after C++ this) is arg1.
'     Writing "this = obj" in stage 9.3 was correct, but with HK_EnvironmentCompleted etc.
'     "arg1 = errorCode" happened to be 0 and went unnoticed; the same rule holds
'     for both 3-argument handlers and 2-argument handlers.
' ============================================================
Public Function Handler_Invoke( _
    ByVal this As LongPtr, _
    ByVal arg1 As LongPtr, _
    ByVal arg2 As LongPtr) As Long

    ' English fork: never let an error escape to WebView2. The thunk returns the
    ' retval slot, which VBA does not write when this method raises, so WebView2
    ' used to receive stack garbage as the HRESULT and the error vanished.
    ' (Wv2Pane's GuardExit blocks deliberately re-raise; they end up here.)
    On Error GoTo InvokeFailed

    Select Case m_kind
        Case HK_EnvironmentCompleted
            Dim errorCode As Long
            errorCode = LongPtrLowDword(arg1)
            Debug.Print "ComCallbackHandler.Invoke[HK_EnvironmentCompleted]"
            Debug.Print "  this        :", this
            Debug.Print "  errorCode   :", errorCode
            Debug.Print "  pEnvironment:", arg2
            Debug.Print "  m_refCount  :", m_refCount

            ' Notify the upper class (only when m_owner exists)
            ' Dynamic binding: m_owner is of type Object, but it can be called as long as it has
            ' a method named OnEnvironmentCompleted.
            ' As of stage 9.3a, Wv2Environment (formerly Class8) implements this method.
            '
            ' Note (9.6a): one-shot callbacks keep the naming without an origin prefix.
            ' There is no function that directly receives a derived IF (ICoreWebView2Environment2 etc.),
            ' so there is no need to identify the version.
            If Not (m_owner Is Nothing) Then
                m_owner.OnEnvironmentCompleted errorCode, arg2
            End If

        Case HK_ControllerCompleted
            Dim errorCode2 As Long
            errorCode2 = LongPtrLowDword(arg1)
            Debug.Print "ComCallbackHandler.Invoke[HK_ControllerCompleted]"
            Debug.Print "  this        :", this
            Debug.Print "  arg1 (raw)  :", arg1
            Debug.Print "  errorCode   :", errorCode2
            Debug.Print "  pController :", arg2
            Debug.Print "  m_refCount  :", m_refCount

            ' Notify the upper class (only when m_owner exists)
            ' Assumes m_owner points to Wv2Pane (formerly Class9, the combined Controller+View class)
            ' One-shot type, so no origin prefix (per the 9.6a convention)
            If Not (m_owner Is Nothing) Then
                m_owner.OnControllerCompleted errorCode2, arg2
            End If

        Case HK_NavigationStarting
            Debug.Print "ComCallbackHandler.Invoke[HK_NavigationStarting]"
            Debug.Print "  this        :", this
            Debug.Print "  arg1 (raw)  :", arg1
            Debug.Print "  arg2 (raw)  :", arg2
            Debug.Print "  m_refCount  :", m_refCount

            ' Notify the upper class (only when m_owner exists)
            '
            ' Thunk behavior (re-check):
            '   The thunk is a mechanism that "shifts the arguments right by one and injects ObjPtr
            '   into rcx". From the VBA view:
            '     Me      = the ObjPtr injected by the thunk (consumed implicitly by VBA)
            '     this    = rdx (the C++ this WebView2 originally put in rcx) = the handler obj
            '     arg1    = r8  (the 1st argument WebView2 originally put in rdx)
            '     arg2    = r9  (the 2nd argument WebView2 originally put in r8)
            '
            ' C++ signature (confirmed on Microsoft Learn):
            '   HRESULT Invoke(ICoreWebView2* sender, ICoreWebView2NavigationStartingEventArgs* args)
            ' This is a C++ member function = 3 arguments including the implicit this:
            '   rcx = this (handler), rdx = sender, r8 = args
            ' VBA view after the thunk:
            '   this (ours) = handler obj (just received the C++ this, normally unused)
            '   arg1        = sender = ICoreWebView2 pointer (should be the same as m_pView)
            '   arg2        = args   = ICoreWebView2NavigationStartingEventArgs pointer
            '
            ' Caution (lesson from stage 9.4a):
            '   The "this = obj" established in 9.3 is the same for HK_EnvironmentCompleted/Controller.
            '   At that stage errorCode = 0 happened to succeed every time, so arg1 = errorCode
            '   was merely interpreted correctly by luck. The (this, arg1, arg2) that ComCallbackHandler receives is always
            '   the post-thunk (rdx, r8, r9), which in WebView2's view are the (1st, 2nd, 3rd) arguments.
            If Not (m_owner Is Nothing) Then
                m_owner.View_OnNavigationStarting arg1, arg2
            End If

        Case HK_NavigationCompleted
            Debug.Print "ComCallbackHandler.Invoke[HK_NavigationCompleted]"
            Debug.Print "  this        :", this
            Debug.Print "  arg1 (raw)  :", arg1
            Debug.Print "  arg2 (raw)  :", arg2
            Debug.Print "  m_refCount  :", m_refCount

            ' 2-argument event: Invoke(sender, args)
            '   arg1 = sender (= m_pView)
            '   arg2 = args   (= ICoreWebView2NavigationCompletedEventArgs)
            If Not (m_owner Is Nothing) Then
                m_owner.View_OnNavigationCompleted arg1, arg2
            End If

        Case HK_WebMessageReceived
            Debug.Print "ComCallbackHandler.Invoke[HK_WebMessageReceived]"
            Debug.Print "  this        :", this
            Debug.Print "  arg1 (raw)  :", arg1
            Debug.Print "  arg2 (raw)  :", arg2
            Debug.Print "  m_refCount  :", m_refCount

            ' 2-argument event: Invoke(sender, args)
            '   arg1 = sender (= m_pView)
            '   arg2 = args   (= ICoreWebView2WebMessageReceivedEventArgs)
            If Not (m_owner Is Nothing) Then
                m_owner.View_OnWebMessageReceived arg1, arg2
            End If

        Case HK_DocumentTitleChanged
            Debug.Print "ComCallbackHandler.Invoke[HK_DocumentTitleChanged]"
            Debug.Print "  this        :", this
            Debug.Print "  arg1 (raw)  :", arg1
            Debug.Print "  arg2 (raw)  :", arg2
            Debug.Print "  m_refCount  :", m_refCount

            ' 2-argument event: Invoke(sender, args)
            '   arg1 = sender (= m_pView)
            '   arg2 = args   (= IUnknown* with no information; you must call
            '                  ICoreWebView2::get_DocumentTitle via sender)
            If Not (m_owner Is Nothing) Then
                m_owner.View_OnDocumentTitleChanged arg1, arg2
            End If

        Case HK_NewWindowRequested
            Debug.Print "ComCallbackHandler.Invoke[HK_NewWindowRequested]"
            Debug.Print "  this        :", this
            Debug.Print "  arg1 (raw)  :", arg1
            Debug.Print "  arg2 (raw)  :", arg2
            Debug.Print "  m_refCount  :", m_refCount

            ' 2-argument event: Invoke(sender, args)
            '   arg1 = sender (= m_pView)
            '   arg2 = args   (= ICoreWebView2NewWindowRequestedEventArgs)
            '
            ' vtable layout of args (from the IDL comments of the old WebView2.cls):
            '    3 : get_Uri(LPWSTR* uri)
            '    4 : put_NewWindow(ICoreWebView2*)
            '    5 : get_NewWindow(ICoreWebView2**)
            '    6 : put_Handled(BOOL)
            '    7 : get_Handled(BOOL*)
            '    8 : get_IsUserInitiated(BOOL*)
            '    9 : GetDeferral(ICoreWebView2Deferral**)
            '   10 : get_WindowFeatures(ICoreWebView2WindowFeatures**)
            If Not (m_owner Is Nothing) Then
                m_owner.View_OnNewWindowRequested arg1, arg2
            End If

        Case HK_HistoryChanged
            ' Added in 9.9b. Persistent, base ICoreWebView2 (vtable 13/14).
            Debug.Print "ComCallbackHandler.Invoke[HK_HistoryChanged]"
            Debug.Print "  this        :", this
            Debug.Print "  arg1 (raw)  :", arg1
            Debug.Print "  arg2 (raw)  :", arg2
            Debug.Print "  m_refCount  :", m_refCount

            ' 2-argument event: Invoke(sender, args)
            '   arg1 = sender (= m_pView)
            '   arg2 = args   (= IUnknown* with no information. The args of HistoryChanged carry
            '                  no information, so call get_Source /
            '                  get_CanGoBack / get_CanGoForward via sender)
            If Not (m_owner Is Nothing) Then
                m_owner.View_OnHistoryChanged arg1, arg2
            End If

        Case HK_DOMContentLoaded
            ' Added in 9.9b. Persistent, ICoreWebView2_2 (vtable 64/65).
            Debug.Print "ComCallbackHandler.Invoke[HK_DOMContentLoaded]"
            Debug.Print "  this        :", this
            Debug.Print "  arg1 (raw)  :", arg1
            Debug.Print "  arg2 (raw)  :", arg2
            Debug.Print "  m_refCount  :", m_refCount

            ' 2-argument event: Invoke(sender, args)
            '   arg1 = sender (= m_pView, but the Invoke of ICoreWebView2_2 passes
            '          sender as ICoreWebView2*. The same value as m_pView)
            '   arg2 = args   (= ICoreWebView2DOMContentLoadedEventArgs)
            '          vtable 3 = get_NavigationId(UINT64*)
            '
            ' The receiver is named View2_OnDOMContentLoaded (View2_ because it comes from a derived IF,
            ' 9.6a/9.9b naming convention).
            If Not (m_owner Is Nothing) Then
                m_owner.View2_OnDOMContentLoaded arg1, arg2
            End If

        Case HK_WebResourceRequested
            ' Added in N-1. Persistent, base ICoreWebView2 (vtable 55/56).
            '
            ' **Do not fire Debug.Print here (N-1 issue 5, see the header)**
            '   Tens to hundreds arrive in succession, and if the handler is slow
            '   the traffic itself jams. Just record and return immediately.
            '
            ' 2-argument event: Invoke(sender, args)
            '   arg1 = sender (= m_pView)
            '   arg2 = args   (= ICoreWebView2WebResourceRequestedEventArgs)
            '
            ' vtable layout of args (**confirmed from WebView2.h with tools/vtable.py**):
            '    3 : get_Request(ICoreWebView2WebResourceRequest**)
            '    4 : get_Response      5 : put_Response      6 : GetDeferral
            '    7 : get_ResourceContext(COREWEBVIEW2_WEB_RESOURCE_CONTEXT*)
            If Not (m_owner Is Nothing) Then
                m_owner.View_OnWebResourceRequested arg1, arg2
            End If

        Case HK_WebResourceResponseReceived
            ' Added in N-3. Persistent, ICoreWebView2_2 (vtable 61/62).
            '
            ' **Do not fire Debug.Print here either** (same reason as N-1. Moreover
            '   this event has **no filter**, so it fires even more often)
            '
            ' 2-argument event: Invoke(sender, args)
            '   arg1 = sender (= m_pView. Even with the _2 IF the sender is ICoreWebView2*)
            '   arg2 = args   (= ICoreWebView2WebResourceResponseReceivedEventArgs)
            '
            ' vtable of args (**confirmed with tools/vtable.py**):
            '    3 : get_Request(ICoreWebView2WebResourceRequest**)
            '    4 : get_Response(ICoreWebView2WebResourceResponseView**)
            If Not (m_owner Is Nothing) Then
                m_owner.View2_OnWebResourceResponseReceived arg1, arg2
            End If

        Case HK_ResponseGetContent
            ' Added in N-4. One-shot type, concurrency OK.
            '   ICoreWebView2WebResourceResponseViewGetContentCompletedHandler
            '     Invoke(HRESULT errorCode, IStream* result)
            '
            ' Interpretation of arguments (same shape as ExecuteScriptCompleted):
            '   arg1 = errorCode (low 32 bits of the LongPtr)
            '   arg2 = IStream*  (**0 for a response with no body**)
            '
            ' **Not read here** Settings such as the limit are held by Wv2Pane, so
            '   the pointer is passed on as is. **The IStream is valid only during Invoke**, so
            '   the Wv2Pane side must read it all synchronously before returning.
            '
            ' **Do not fire Debug.Print** (same as N-1 onward. Responses are numerous)
            Dim errorCode4 As Long
            errorCode4 = LongPtrLowDword(arg1)

            If Not (m_owner Is Nothing) Then
                m_owner.OnResponseContentCompleted m_callbackId, errorCode4, arg2
            End If

        Case HK_ExecuteScriptCompleted
            ' Added in 9.8c. One-shot type, concurrency OK (plan a adopted).
            '
            ' SDK signature:
            '   ICoreWebView2ExecuteScriptCompletedHandler::Invoke(
            '       HRESULT errorCode,
            '       LPCWSTR result)
            '
            ' Interpretation of arguments:
            '   arg1 = errorCode  (low 32 bits of the LongPtr, extracted with LongPtrLowDword)
            '   arg2 = result     (LPCWSTR, the JSON-encoded result string)
            '
            ' Responsibility for freeing result: by convention the LPCWSTR is not freed by the caller (= VBA)
            '   (= it is passed as a one-shot Invoke argument and becomes invalid after Invoke returns).
            '   If it is copied into a String here with PtrToString, later references are safe.
            Dim errorCode3 As Long
            errorCode3 = LongPtrLowDword(arg1)
            Dim resultStr As String
            If arg2 <> 0 Then
                resultStr = PtrToString(arg2)
            End If

            Debug.Print "ComCallbackHandler.Invoke[HK_ExecuteScriptCompleted]"
            Debug.Print "  this        :", this
            Debug.Print "  callbackId  :", m_callbackId
            Debug.Print "  errorCode   :", errorCode3
            Debug.Print "  resultLen   :", Len(resultStr)
            Debug.Print "  m_refCount  :", m_refCount

            ' Notify the upper class (Wv2Pane).
            ' Plan alpha first implementation: Wv2Pane.OnExecuteScriptCompleted logs + discards.
            ' If we later want to hand it over via a UserForm, extend the Wv2Pane side.
            ' Dynamic binding: assumes m_owner points to Wv2Pane
            ' One-shot type, so no origin prefix (per the 9.6a convention)
            If Not (m_owner Is Nothing) Then
                m_owner.OnExecuteScriptCompleted m_callbackId, errorCode3, resultStr
            End If

            ' Because of concurrency, Wv2Pane holds no reference to ComCallbackHandler
            ' (= it does Set h = Nothing and discards it right after AcquireHandlerFor).
            ' The completion notification has been received, so cut our own m_owner reference here.
            ' When the WebView2 runtime side's Release runs, refCount=0, and
            ' ClearOwner is also called from Handler_Release's automatic release path, so the
            ' circular reference is cleanly resolved.
            ' In practice, even without cutting m_owner = Wv2Pane immediately, ClearOwner fires
            ' inside Handler_Release, so no explicit Set is done here.
            ' (= same flow as the existing HK_EnvironmentCompleted / HK_ControllerCompleted)

        Case Else
            Debug.Print "ComCallbackHandler.Invoke[unknown kind=" & m_kind & "]"
    End Select

    Handler_Invoke = 0   ' S_OK
Exit Function

InvokeFailed:
    Wv2Log.LogE "ComCallbackHandler.Invoke: unhandled error " & Err.Number & " " & Err.Description & _
                " (kind=" & m_kind & ") - logged and swallowed, S_OK returned"
    Handler_Invoke = 0   ' S_OK
End Function


' ============================================================
' Extract the low 32 bits of a LongPtr as a signed Long
'
'   Under the x64 calling convention, when a 32-bit argument (HRESULT etc.) is passed,
'   the upper 32 bits may be left undefined in rdx etc.
'   The thunk shifts by 64 bits like "mov rdx, rcx",
'   so if received as a LongPtr as is, garbage in the upper 32 bits remains.
'
'   Extract the low 32 bits and, if it is 0x80000000 or above, wrap it to a negative Long
'   (CLng overflows outside the Long range, so we handle it ourselves).
' ============================================================
Private Function LongPtrLowDword(ByVal v As LongPtr) As Long
    Dim u As LongLong
    u = v And &HFFFFFFFF^
    If u >= &H80000000^ Then
        LongPtrLowDword = CLng(u - &H100000000^)
    Else
        LongPtrLowDword = CLng(u)
    End If
End Function

