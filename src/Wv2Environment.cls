VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Wv2Environment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''
' --- Wv2Environment.cls Stage 9.8a (formerly Class8) ---
'
'   Changes in 9.8a (rename only, no logic change):
'     - Class name Class8 -> Wv2Environment, now fixed officially
'     - References to Class7 changed to ComCallbackHandler
'     - The class's own debug log changed from "Class8.xxx" to "Wv2Environment.xxx"
'     - Responsibilities, lifetime management and state transitions established in 9.3a are unchanged
'
'   Role: a VBA-side wrapper for ICoreWebView2Environment.
'     - Creates and holds one HK_EnvironmentCompleted handler (ComCallbackHandler)
'     - Calls CreateCoreWebView2EnvironmentWithOptions to request creation of the WebView2 environment
'     - Completion is notified as follows: ComCallbackHandler.Handler_Invoke is hit via the thunk,
'       and from its m_kind branch our OnEnvironmentCompleted is called
'     - Fires AddRef on the received pEnvironment per the COM convention, and is
'       responsible for Releasing it in Class_Terminate
'
'   Lifetime management (plan D1-alpha: one-shot handler pattern):
'     - HK_EnvironmentCompleted is a handler that is called once and is done.
'     - At the end of OnEnvironmentCompleted, Set m_handler = Nothing to drop the
'       external reference to ComCallbackHandler.
'     - When WebView2 Releases ComCallbackHandler the refcount reaches 0 and the
'       foundation's automatic release path runs (ComCallbackHandler.ClearOwner cuts
'       m_owner=Me, resolving the circular reference between Wv2Environment and
'       ComCallbackHandler).
'     - After that Wv2Environment (Me) is released, e.g. when the test procedure's local variable
'       goes out of scope, and m_pEnvironment->Release runs in Class_Terminate.
'
'   State transitions:
'     Es_Created       : right after Class_Initialize
'     Es_Waiting       : after Init calls the API, waiting for the callback
'     Es_Ready         : OnEnvironmentCompleted received errorCode=0
'     Es_Failed        : OnEnvironmentCompleted received errorCode<>0
''''''''''''''''''''''''''''''''''
Option Explicit

#If Win64 Then
    ' OK
#Else
    #Error "This class requires 64-bit VBA (x64)"
#End If


Public Enum EnvironmentState
    Es_Created = 0
    Es_Waiting = 1
    Es_Ready = 2
    Es_Failed = 3
End Enum


' --- Internal state ---
' m_handler:      strong reference to the HK_EnvironmentCompleted handler
'                 (set to Nothing after OnEnvironmentCompleted)
' m_pEnvironment: the ICoreWebView2Environment received from WebView2
'                 (responsible for Releasing it in Class_Terminate)
Private m_handler      As ComCallbackHandler
Private m_pEnvironment As LongPtr
Private m_state        As EnvironmentState
Private m_lastError    As Long


' ============================================================
' Class_Initialize / Class_Terminate
' ============================================================
Private Sub Class_Initialize()
    m_state = Es_Created
    m_pEnvironment = 0
    m_lastError = 0
End Sub

Private Sub Class_Terminate()
    ' COM ownership convention: OnEnvironmentCompleted fired AddRef on the pEnvironment received in Init
    ' to secure the VBA side's share of ownership,
    ' so fire Release here to release it.
    If m_pEnvironment <> 0 Then
        Debug.Print "Wv2Environment.Terminate: Releasing m_pEnvironment = " & m_pEnvironment
        ComRelease m_pEnvironment
        m_pEnvironment = 0
    End If

    ' m_handler should normally already be Nothing after OnEnvironmentCompleted, but
    ' if it remains, e.g. because an exception occurred in Init, cut it here.
    ' (The ComCallbackHandler side is also referenced from the foundation's m_handlers(idx), so
    '  cutting it here makes the slot's automatic release path run.)
    If Not (m_handler Is Nothing) Then
        Debug.Print "Wv2Environment.Terminate: m_handler remains, cutting it (abnormal termination path)"
        Set m_handler = Nothing
    End If
End Sub


' ============================================================
' Init
'
'   Starts creation of ICoreWebView2Environment.
'   This method returns immediately without waiting for the callback (asynchronous).
'   Completion is notified through OnEnvironmentCompleted.
'
'   The caller waits with DoEvents until env.IsReady becomes True.
' ============================================================
Public Sub Init(ByVal userDataPath As String)

    If m_state <> Es_Created Then
        MsgBox "Wv2Environment.Init cannot be called twice", vbCritical
        Exit Sub
    End If

    ' English fork: load WebView2Loader.dll by full path first (see Wv2Thunks.EnsureWebView2Loader).
    '   Without it, a missing DLL raised run-time error 53 from the Declare. Checked before a
    '   handler slot is acquired so that this failure path leaks nothing.
    If LenB(Wv2Thunks.EnsureWebView2Loader()) = 0 Then
        MsgBox "WebView2Loader.dll (x64) was not found." & vbCrLf & _
               "Put it next to the workbook or set WEBVIEW2_LOADER_PATH to its full path." & vbCrLf & _
               "(It is in the Microsoft.Web.WebView2 NuGet package: runtimes\win-x64\native.)", vbCritical
        m_state = Es_Failed
        m_lastError = &H8007007E   ' HRESULT_FROM_WIN32(ERROR_MOD_NOT_FOUND)
        Exit Sub
    End If


    ' --- 1. Acquire one handler (ComCallbackHandler) ---
    '   Inside AcquireHandlerFor:
    '     - New ComCallbackHandler
    '     - Thunks_AcquireSlot secures a slot + writes the thunk
    '     - ComCallbackHandler.Init(HK_EnvironmentCompleted, owner=Me, pSlot)
    '     - registers in m_handlers(idx)
    '   are all done together, and an initialized ComCallbackHandler is returned.
    Set m_handler = AcquireHandlerFor(HK_EnvironmentCompleted, Me)
    If m_handler Is Nothing Then
        MsgBox "Wv2Environment.Init: failed to Acquire the handler", vbCritical
        m_state = Es_Failed
        Exit Sub
    End If

    ' --- 2. Compute the obj address that WebView2 holds ---
    '   pSlot + VTABLE_OBJ_OFFSET is the start of vtableObj.
    '   This address is passed to WebView2 as obj and is used as the
    '   this of later AddRef/Release/Invoke.
    Dim pVTableObj As LongPtr
    pVTableObj = m_handler.Slot + VTABLE_OBJ_OFFSET

    ' --- 3. Prepare the user data folder ---
    Dim path As String
    path = userDataPath
    If LenB(path) = 0 Then path = Wv2Log.AppFolder("Profile")
    EnsureFolder path

    ' --- 4. Call CreateCoreWebView2EnvironmentWithOptions ---
    '
    '   Important: set m_state = Es_Waiting BEFORE the API call.
    '   The reason is that WebView2Loader may fire the callback synchronously inside the API
    '   (confirmed in real-machine logs):
    '
    '     CreateCoreWebView2EnvironmentWithOptions call starts
    '       +- internally WebView2 runs AddRef -> Invoke -> Release synchronously
    '          +- Wv2Environment.OnEnvironmentCompleted is called and m_state = Es_Ready
    '       <- the API returns S_OK
    '
    '   If m_state = Es_Waiting were set AFTER the API call, the Es_Ready set by the callback
    '   would be overwritten with Es_Waiting, and
    '   the caller's IsReady loop would never become True.
    '
    '   If Es_Waiting is set first, then for both synchronous and asynchronous callbacks the
    '   final state is the value set by OnEnvironmentCompleted (Es_Ready / Es_Failed),
    '   which stays consistent. Only when the API fails is it overwritten with Es_Failed afterwards.
    m_state = Es_Waiting

    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(path), 0, pVTableObj)

    If hr <> 0 Then
        MsgBox "CreateCoreWebView2EnvironmentWithOptions failed. HRESULT: &H" & Hex(hr), _
               vbCritical
        Wv2Thunks.Thunks_DiscardHandler m_handler   ' English fork: return the slot
        Set m_handler = Nothing
        m_state = Es_Failed
        m_lastError = hr
        ' Leak handling on failure is not implemented in 9.3a.
        ' If WebView2 returns failure without firing AddRef, refcount stays 0 and
        ' m_handlers(idx) remains, leaking the slot.
        ' Recovery relies on calling Thunks_Shutdown to force-clean all slots.
        ' Plan to design a unified helper in 9.3b or later together with the Controller creation failure path.
        Exit Sub
    End If
End Sub


' ============================================================
' OnEnvironmentCompleted
'
'   From the HK_EnvironmentCompleted branch of ComCallbackHandler.Handler_Invoke,
'   m_owner.OnEnvironmentCompleted is called via dynamic binding.
'
'   This method must be a Public Sub:
'     ComCallbackHandler's m_owner is of type Object, so the call is late-bound
'     (via IDispatch::Invoke). Friend does not go onto IDispatch, so the
'     runtime error "Object doesn't support this property or method." occurs.
'     With Public it goes onto IDispatch and can be called.
'
'   Arguments:
'     errorCode    : HRESULT (S_OK = 0 means success)
'     pEnvironment : ICoreWebView2Environment pointer
'                    (passed with a temporary AddRef by WebView2.
'                     WebView2 Releases it after the callback returns,
'                     so to keep holding it on the VBA side we must AddRef it ourselves.)
' ============================================================
Public Sub OnEnvironmentCompleted( _
    ByVal errorCode As Long, _
    ByVal pEnvironment As LongPtr)

    Debug.Print "Wv2Environment.OnEnvironmentCompleted: errorCode=" & errorCode & _
                ", pEnvironment=" & pEnvironment

    m_lastError = errorCode

    If errorCode = 0 And pEnvironment <> 0 Then
        ' Secure the reference that the VBA side keeps
        m_pEnvironment = pEnvironment
        ComAddRef m_pEnvironment
        m_state = Es_Ready
    Else
        m_pEnvironment = 0
        m_state = Es_Failed
    End If

    ' One-shot handler pattern (plan D1-alpha):
    '   The handler has served its purpose, so drop the reference.
    '   At this point, inside ComCallbackHandler.Handler_Invoke, WebView2
    '   still holds refcount=1 (WebView2 fires Release right after).
    '   Cutting m_handler here removes the Wv2Environment -> ComCallbackHandler reference, but since
    '   it is still referenced from the foundation's m_handlers(idx), ComCallbackHandler
    '   stays alive. The moment WebView2 fires Release and refcount reaches 0,
    '   the automatic release path of the foundation's HandlerReleaseInternal runs, and
    '   ComCallbackHandler.ClearOwner also cuts the reference to m_owner (= Me, Wv2Environment).
    Set m_handler = Nothing
End Sub


' ============================================================
' Property group
' ============================================================
Public Property Get state() As EnvironmentState
    state = m_state
End Property

Public Property Get IsReady() As Boolean
    IsReady = (m_state = Es_Ready)
End Property

Public Property Get IsFailed() As Boolean
    IsFailed = (m_state = Es_Failed)
End Property

Public Property Get LastError() As Long
    LastError = m_lastError
End Property

Public Property Get EnvironmentPtr() As LongPtr
    EnvironmentPtr = m_pEnvironment
End Property




