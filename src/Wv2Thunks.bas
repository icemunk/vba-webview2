Attribute VB_Name = "Wv2Thunks"
''''''''''''''''''''''''''''''''''
' --- Wv2Thunks.bas English fork v1.1 ---
'
'   EnsureWebView2Loader no longer probes candidates with Dir$ (it raises on a
'   malformed or unreachable path); LoadLibraryW is the existence test.
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Thunks.bas Stage N-4 (response body) ---
'
'   Changes in N-4:
'     - Added HK_ResponseGetContent = 13 to HandlerKind
'       **One-shot type, concurrency OK** = treated the same as HK_ExecuteScriptCompleted
'       (not put into the persistent array. Calls are identified by CallbackId)
'     - Extended the upper bound of the m_iidTable array to HK_ResponseGetContent
'     - IID {875738e1-9fa2-40e3-8b74-2e8972dd6fe7}
'
'   **The shape of Invoke is the same as ExecuteScriptCompleted**
'     Invoke(HRESULT errorCode, IStream* result)
'     -> arg1 = errorCode (low 32 bits), arg2 = IStream* (may be 0)
'     The CallbackId mechanism built in stage 9.8c can be used as is.
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Thunks.bas Stage N-3 (response status and headers) ---
'
'   Changes in N-3:
'     - Added HK_WebResourceResponseReceived = 12 to HandlerKind
'       **Persistent, ICoreWebView2_2 (vtable 61/62)** = an event of a derived IF
'     - Extended the upper bound of the m_iidTable array from HK_WebResourceRequested to HK_WebResourceResponseReceived
'     - IID {7de9898a-24f5-40c3-a2de-d4f458e69828}
'       **The source is the MIDL_INTERFACE line of the SDK header (design principle 115)**
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Thunks.bas Stage N-2 (request headers and body) ---
'
'   Changes in N-2:
'     - Added the `MultiByteToWideChar` declaration (kernel32)
'     - Added `Utf8BytesToString`. **Turns a UTF-8 byte sequence into a VBA String**
'
'   **Why ADODB.Stream is not used**
'     Product code is bound by the rule of using no external COM libraries. ADODB is used only in
'     verification code (WriteUtf8NoBom in Wv2Tests). `MultiByteToWideChar` is a
'     Win32 API, so it falls in the same category as the already-used `lstrlenW` / `RtlMoveMemory`.
'
'   **StrConv(b, vbUnicode) will not do**
'     It interprets using the project's code page (CP932 in a Japanese environment).
'     Passing a UTF-8 byte sequence garbles it. An API that can specify the code page explicitly is needed.
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Thunks.bas Stage N-1 (wiring of WebResourceRequested) ---
'
'   Changes in N-1:
'     - Added HK_WebResourceRequested = 11 to HandlerKind (persistent, base ICoreWebView2)
'     - Extended the upper bound of the m_iidTable array from HK_DOMContentLoaded to HK_WebResourceRequested
'     - Added the IID of ICoreWebView2WebResourceRequestedEventHandler to InitIIDTable
'       {ab00b74c-15f1-4646-80e8-e76341d25d71}
'       **The source is the MIDL_INTERFACE line of the SDK header WebView2.h (design principle 115)**
'     - Core logic (Handler_QueryInterface / AddRef / Release / dcf) unchanged
'       (kind is accessed dynamically, so the new kind=11 also passes automatically)
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Thunks.bas Stage 9.16 (formerly Module9_9_9.bas) ---
'
'   Changes in 9.16 (cleanup: moved the feature verification Subs to Wv2Tests.bas):
'     **Purified this module into the thunk foundation (plan 1-A)**
'       Wv2Thunks is by nature the heart module that carries "machine-code thunks, vtables and SAFEARRAY memory primitives".
'       The feature-level pure-logic verifications that had piled up there
'       (Test_9_14_SearchFallback / Test_9_15_Reorder / Test_9_16_SearchEngine and
'        the helpers Check / CheckOrder) were moved to the newly created Wv2Tests.bas.
'       They only call Wv2Browser's Debug_* and never touch the thunk layer,
'       so moving them out breaks nothing (the same "separate by concern" principle as when Wv2Json was split out).
'
'     **Verification of the thunk foundation itself stays**
'       Test_TwoHandlers_Mock / Test_Environment_Real / Test_ResetSurvival_Manual /
'       Test_VirtualAllocFree_Roundtrip / Test_SentinelStatus are tightly coupled to the
'       Private implementation inside this module (Thunks_* / Sentinel_* / VirtualAlloc etc.), so
'       they stay here. Moving them out would require making Private members Public for verification,
'       which would break encapsulation instead (criterion: "verification tightly coupled to production code stays in that module").
'
'     - Core logic (the Public/Private API of the thunk foundation) is completely unchanged.
'
' --- Wv2Thunks.bas Stage 9.15 (formerly Module9_9_9.bas) ---
'
'   Changes in 9.15 (added a logic verification Sub for tab reordering):
'     - Added Test_9_15_Reorder. It calls Wv2Browser's verification ports (Debug_SeedDummyTabs /
'       MoveTab / Debug_TabOrderTitles / Debug_ActiveIndex) and verifies, in one batch without starting WebView2,
'       the order after tab reordering and the active-tab tracking (a pure logic test).
'       It goes through the same MoveTab that the production OnPaneWebMessage(reorder) calls, so
'       there is no double maintenance (same position as Debug_NormalizeUrl in Test_9_14).
'     - Added the CheckOrder helper (compares an order string and the expected activeIndex).
'     - Core logic and existing tests unchanged. This time only a verification Sub was added.
'
' --- Wv2Thunks.bas Stage 9.14 (formerly Module9_9_9.bas) ---
'
'   Changes in 9.14 (added a logic verification Sub for the search-term fallback):
'     - Added Test_9_14_SearchFallback. It calls Wv2Browser.Debug_NormalizeUrl and verifies, in one batch
'       without starting WebView2, that address bar input is correctly sorted into
'       "URL" and "Google search" (a pure logic test).
'       The verification goes through Wv2Browser's Debug_NormalizeUrl, so it runs the same logic as the NormalizeUrl
'       used by the production NavigateActive (no double maintenance).
'     - Found in real-machine testing: "localhost:8080" has only letters before the colon (IsAlphaOnly), so
'       it is treated as a scheme and passed through, and Edge/WebView2 supplies http:// and connects
'       (= correct behavior for a browser). Corrected the expected values of the verification Sub to match actual behavior.
'     - Core logic and existing tests unchanged. This time only a verification Sub was added.
'
''''''''''''''''''''''''''''''''''
' --- Wv2Thunks.bas Stage 9.13 (formerly Module9_9_9.bas) ---
'
'   Changes in 9.13 (cleanup stage, sweeping up leftovers, this session):
'     - Removed 4 UserForm-dependent integration test Subs. All of them called StartWebView2 (and IsPageLoaded /
'       GetActivePane / OnExecuteScriptResult), which was already removed from
'       UserForm1 in 9.13, and leaving them would cause
'       compile errors:
'         * Test_Controller_Real   (9.3b)  <- depends on StartWebView2
'         * Test_Navigate_Real     (9.3c)  <- depends on StartWebView2
'         * Test_AllEvents_Real    (9.4b)  <- depends on StartWebView2
'         * Test_ExecuteScript_Parallel (9.9c) <- depends on StartWebView2 and others
'       Parallel ExecuteScript verification already passed in 9.9c. When SPA integration is
'       taken up in earnest in the future, rewriting on top of StartWebView2_Spa makes more sense,
'       so they are removed for now (agreed).
'     - The 5 low-level unit tests that remain (thunks, vtables, memory primitives.
'       Living assets that a GUI cannot replace):
'         * Test_TwoHandlers_Mock           (mock verification of vtable/thunk)
'         * Test_Environment_Real           (Environment creation only)
'         * Test_VirtualAllocFree_Roundtrip (memory primitives)
'         * Test_SentinelStatus             (sentinel)
'         * Test_ResetSurvival_Manual       (VBA Reset survival guide)
'     - The procedure comment of Test_ResetSurvival_Manual referred to the removed Test_Navigate_Real, so
'       it was rewritten to a procedure using UserForm1.StartWebView2_Full
'       (made self-contained as a procedure guide).
'     - Core logic (Handler_xxx / dcf / Thunks_Init / memory primitives etc.) is
'       completely unchanged. This time only "removal of throwaway test Subs".
'
''''''''''''''''''''''''''''''''''
' --- Wv2Thunks.bas Stage 9.10a (formerly Module9_9_9.bas) ---
'
'   Changes in 9.10a (this session, handling the Sentinel MEM_FREE bug = spec fact 10):
'     - Added constants: MEM_FREE (&H10000&) / PAGE_READWRITE (&H4&)
'     - Rewrote Sentinel_RecoverIfNeeded as a "State guard" approach:
'         Only when the result of VirtualQuery (State=MEM_COMMIT and AllocationBase=prevBase and
'         Protect is read/write) confirms that "the region is alive" does it run
'         ReadLongPtr / MemLongPtr / VirtualFree.
'         If it is not alive, it fires only Sentinel_ClearPrevRegion and safely does Exit Sub.
'       -> The old implementation unconditionally fired ReadLongPtr / MemLongPtr at the freed address,
'         causing an AV (= SEH exception) in native array access through the SAFEARRAY.
'         This AV could not be caught with On Error Resume Next and Excel
'         crashed immediately (spec fact 10). Spelled out as design principle 55.
'     - New Private Function: IsProtectReadWritable(prot) As Boolean
'         (PAGE_GUARD bit detection + PAGE_READWRITE / PAGE_EXECUTE_READWRITE check)
'     - Core logic (Handler_xxx / dcf / Thunks_Init etc.) otherwise unchanged
'
'   Changes in 9.9c (continued, complementary group):
'     - Newly added Test_ExecuteScript_Parallel (item D)
'         Issues 5 ExecuteScript calls in succession and verifies on a real machine that
'         concurrency (plan a, 9.8c) works correctly, by matching callback IDs to result values.
'         Throws sequential calculations ((1+1), (2+1), ... ) and looks at the match between callback ID and result.
'     - Core logic (Handler_xxx / dcf / Thunks_Init etc.) unchanged
'         (9.9c only adds one test function to Wv2Thunks)
'
'   Changes in 9.9b (continued, SPA support):
'     - Added 2 kinds to HandlerKind (HK_HistoryChanged=9 / HK_DOMContentLoaded=10)
'     - Extended the upper bound of the m_iidTable array to HK_DOMContentLoaded
'     - Added 2 IIDs to InitIIDTable
'     - Newly created GetLongLongProperty (wrapped in #If Win64 = 64-bit only)
'     - Newly created BuildSpaTestHtml / Test_SpaTest_Real (for SPA testing)
'     - Plan R': added try-catch to doPush/doReplace of BuildSpaTestHtml
'              (show the SecurityError of the about:blank origin on screen)
'
'   ** Consideration for future 32-bit extension of 9.9 (a + b + c) (policy X) **
'     All stages were implemented 64-bit only (#Else #Error kept), but
'     "32-bit conversion point" comments are left at LongLong-dependent places in new code.
'     The large task of converting the thunk machine code to x86 is carved out as a separate stage (future).
'     EventRegistrationToken is 64-bit (LongLong) even on x86, so supporting both is no problem.
'
'   Changes in 9.9a (continued):
'     - Promoted FillGUID from Private to Public
'         (needed so Wv2Pane.EnsureView2 can initialize a local GUID structure)
'     - Implemented the ICoreWebView2_2 retrieval foundation on the Wv2Pane side (m_pView2 + EnsureView2)
'
'   Changes in 9.8c (continued):
'     - Added HK_ExecuteScriptCompleted = 8 to HandlerKind
'         (one-shot type, concurrency OK, plan a adopted)
'     - Extended the upper bound of the m_iidTable array from HK_NewWindowRequested to HK_ExecuteScriptCompleted
'     - Added IID_ICoreWebView2ExecuteScriptCompletedHandler to InitIIDTable
'         (49511172-CC67-4BCA-9923-137112F4C4CC)
'     - Handler_QueryInterface / Handler_AddRef / Handler_Release / Handler_Invoke
'       The core logic is unchanged (kind is accessed dynamically, so the new kind=8 also passes automatically)
'
'   Changes in 9.8a/9.8b (continued):
'     - Module name Module9_9_9 -> Wv2Thunks, now fixed officially (9.8a)
'     - Array type m_handlers() As ComCallbackHandler (9.8a)
'     - No change to Wv2Thunks in 9.8b (the Settings family is self-contained within Wv2Pane)
'
'   Role of the module:
'     The foundation layer that supports the actual WebView2 handlers. It handles thunk region allocation (VirtualAlloc),
'     slot management (m_handlers / m_freeSlots), reference counting (HandlerAddRefInternal /
'     HandlerReleaseInternal), Handler_QueryInterface /
'     Handler_AddRef / Handler_Release / Handler_Invoke arriving from the thunks, the IID check
'     (m_iidTable + IsEqualGUIDInPlace, 9.7a), region reclamation on VBA reset
'     via the sentinel mechanism (9.5), and consolidation of Win32 API declarations (the
'     mouse resize APIs were also added in 9.7b).
'
'   Main things established in past stages (see dev-notes-0-index.md for details):
'     9.2: established the thunk method (DispCallFunc -> turning back to VBA methods)
'     9.3: introduced the upper classes (Wv2Environment / Wv2Pane), plan D1-alpha
'     9.4: persistent handler mechanism (plan P1), HandlerKind-linked arrays (plan r)
'     9.5: complete elimination of region leaks via the sentinel mechanism, discovery of the &H literal trap
'            (MEM_RELEASE = &H8000& fix)
'     9.6: rollout of NewWindowRequested, naming convention by origin prefix (View_/Ctrl_)
'     9.7a: made Handler_QueryInterface genuine (introduced the IID check)
'     9.7b: APIs for UserForm mouse resizing (GetAncestor / SetWindowLongPtrW /
'             SetWindowPos) and related constants (GA_ROOT / WS_THICKFRAME etc.)
'
'   ** Important facts established in 9.7a **
'     Observing the logs of Test_AllEvents_Real revealed the following:
'     1. For a registered event handler, the WebView2 runtime does not request
'        anything other than IID_IUnknown and the proper IID via QueryInterface
'     2. WebView2 adopts the pattern "QI at registration and cache the vtable; at Invoke time
'        invoke directly without QI"
'     3. The true reason the "pass every IID through" implementation up to 9.6b caused no actual harm is that
'        WebView2 issues only a minimum of QIs, so the dangerous situation of
'        "returning a false IMarshal pointer" never arose
'
'   ** Troubleshooting guideline (from 9.7a onward) **
'     If puzzling behavior appears (= handler not registered, Invoke not called,
'     AddRef/Release mismatch, hang at Cleanup, etc.):
'       (1) Temporarily add at the top of Handler_QueryInterface
'           "ppvObject = this : HandlerAddRefInternal this :
'            Handler_QueryInterface = S_OK : Exit Function"
'           to return to the 9.6b-equivalent behavior
'       (2) If the function is fixed in this state, the IID check is confirmed as the cause
'       (3) If not fixed, the IID check is unrelated; look for another cause
'
'   Future tasks (what remains at 9.9c completion = all of 9.9 complete):
'     - TabWebView2 comprehensive class (planned to start in 9.10, the largest)
'     - A real SPA test environment (full verification of pushState/replaceState on an http(s) origin, 9.10 onward)
'     - Version-extension IFs (ICoreWebView2_3 and later, to be added when needed)
'     - 32-bit (x86) support: x86 conversion of the thunk machine code + type abstraction (a separate future stage)
'     Note: SPA support (HistoryChanged / DOMContentLoaded / postMessage convention sample) and
'        the ExecuteScript complement group (D + E + F + G) were completed up to this session (9.9c).
''''''''''''''''''''''''''''''''''

''
' PointerAccessor / SafeArray logic
' Copyright (c) 2025 Cristian Buse
' Licensed under the MIT License
' https://github.com/WNKLER/refTypes/discussions/3
''

Option Explicit

#If Win64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize = 8
#Else
    #Error "This module requires 64-bit VBA (x64)"
#End If

' --- Structures for SafeArray / PointerAccessor ---
Private Enum SAFEARRAY_FEATURES
    FADF_AUTO = &H1
    FADF_FIXEDSIZE = &H10
End Enum

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type

Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type

' --- WebView2Loader.dll ---
'   Changed to Public Declare in stage 9.3a (to be called from Wv2Environment.Init)
Public Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder As LongPtr, _
    ByVal environmentOptions As LongPtr, _
    ByVal environmentCreatedHandler As LongPtr) As Long

' --- Locating WebView2Loader.dll (English fork) ---
'   The Declare above has no path, so Windows only finds the DLL if it is
'   already loaded in Excel, next to EXCEL.EXE, in System32, in the current
'   directory or on PATH. The copy Office ships (Power Query's bin folder) is
'   on none of those, so it used to work only when Power Query had loaded it.
'   EnsureWebView2Loader loads it by full path first; once a module named
'   WebView2Loader.dll is loaded, the Declare resolves to it.
Private Declare PtrSafe Function GetModuleHandleW Lib "kernel32" ( _
    ByVal lpModuleName As LongPtr) As LongPtr
Private Declare PtrSafe Function LoadLibraryW Lib "kernel32" ( _
    ByVal lpLibFileName As LongPtr) As LongPtr

' --- Win32 API ---
Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, _
    ByVal flAllocationType As Long, _
    ByVal flProtect As Long) As LongPtr

Private Declare PtrSafe Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByVal dwSize As LongPtr, _
    ByVal dwFreeType As Long) As Long

' --- DispCallFunc (oleaut32) ---
'   Newly added in stage 9.3a.
'   Used by the minimal helpers (ComRelease / ComAddRef)
'   for calling things like ICoreWebView2Environment::Release from the VBA side.
'
'   For a ByRef ... As Any parameter, when calling a method with 0 arguments
'   the caller writes "ByVal 0&" to pass NULL.
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" ( _
    ByVal pvInstance As LongPtr, _
    ByVal oVft As LongPtr, _
    ByVal cc As Long, _
    ByVal vtReturn As Integer, _
    ByVal cActuals As Long, _
    ByRef prgvt As Any, _
    ByRef prgpvarg As Any, _
    ByRef pvargResult As Any) As Long

' --- user32: GetClientRect (added in stage 9.3c) ---
'   Gets the client area of an HWND. The position is always (0,0,width,height).
'   Can be used directly as the argument of put_Bounds.
'   Public Declare on the assumption that it is called from Wv2Pane.
Public Declare PtrSafe Function GetClientRect Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByRef lpRect As RECT) As Long

' --- API group for UserForm mouse resizing (added in stage 9.7b) ---
'   Used on the UserForm1.frm side to obtain "Frame1's HWND -> the UserForm's own HWND" and to give
'   the UserForm's own HWND WS_THICKFRAME (= thick frame =
'   mouse-resizable).
'
'   Declarations are consolidated in Module9 (by design principle Win32 APIs are managed together in Module9).
'   Constants are also placed in Module9 as Public Const (referenceable from UserForm1).
'
'   GetAncestor:
'     Gets the ancestor HWND of the specified HWND.
'     Passing gaFlags = GA_ROOT (=2) yields the "root window" that contains that HWND
'     (= the top-level/popup window directly visible from the desktop).
'     Used to get Frame1 -> the UserForm's own HWND.
'
'   GetWindowLongPtrW / SetWindowLongPtrW:
'     Gets/sets extended window information (Window Style etc.).
'     Passing GWL_STYLE (=-16) as the second argument nIndex lets you get/set the window style
'     as a LongPtr (LONG_PTR width on x64, LONG width on x86,
'     abstracted by PtrSafe + LongPtr).
'     Setting WS_THICKFRAME (=&H40000) with OR makes it resizable with the mouse.
'
'   SetWindowPos:
'     Updates the window's position/size/Z order/appearance.
'     After giving WS_THICKFRAME, calling with SWP_FRAMECHANGED triggers
'     redrawing of the non-client area (= reflecting the thick frame display).
'     We do not want to change position/size/Z order, so SWP_NOMOVE | SWP_NOSIZE |
'     SWP_NOZORDER are also set at the same time.
Public Declare PtrSafe Function GetAncestor Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal gaFlags As Long) As LongPtr

Public Declare PtrSafe Function GetWindowLongPtrW Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIndex As Long) As LongPtr

Public Declare PtrSafe Function SetWindowLongPtrW Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As LongPtr) As LongPtr

Public Declare PtrSafe Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal hWndInsertAfter As LongPtr, _
    ByVal X As Long, _
    ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal uFlags As Long) As Long

' --- Constants for UserForm mouse resizing (added in stage 9.7b) ---
'   Standard values used by GetAncestor / GetWindowLongPtrW / SetWindowLongPtrW / SetWindowPos.
'   Matched to the standard Win32 symbol names.
'
'   None exceeds the Integer range, but following the trap of design principle 29,
'   the & suffix is attached to &H... just in case (this rules out the possibility of
'   being interpreted as Integer, and treats them as definitely Long).
Public Const GA_ROOT          As Long = 2
Public Const GWL_STYLE        As Long = -16
Public Const WS_THICKFRAME    As Long = &H40000
Public Const SWP_NOSIZE       As Long = &H1&
Public Const SWP_NOMOVE       As Long = &H2&
Public Const SWP_NOZORDER     As Long = &H4&
Public Const SWP_FRAMECHANGED As Long = &H20&

' --- APIs for string helpers (added in stage 9.4a, derived from the old mWebView2Helper) ---
'   Used to convert an LPWSTR (UTF-16 null-terminated string) to a VBA String.
'   Use of each method:
'     lstrlenW       : gets the character count of an LPWSTR (excluding the NULL terminator)
'     CoTaskMemFree  : frees a COM out-parameter LPWSTR* (a string allocated with
'                      CoTaskMemAlloc). By COM convention, the caller's responsibility.
'     RtlMoveMemory  : memory-to-memory copy (used to follow the old version)
Public Declare PtrSafe Function lstrlenW Lib "kernel32" ( _
    ByVal lpString As LongPtr) As Long

Public Declare PtrSafe Sub CoTaskMemFree Lib "ole32" ( _
    ByVal pv As LongPtr)

Public Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" ( _
    ByVal destination As LongPtr, _
    ByVal source As LongPtr, _
    ByVal length As LongPtr)

' --- Character code conversion (added in N-2) ---
'   Turns a UTF-8 byte sequence into UTF-16 (a VBA String).
'   **The key is being able to specify the code page explicitly** StrConv(b, vbUnicode)
'   interprets using the project's code page (CP932), so UTF-8 gets garbled.
Public Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal codePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long) As Long

Public Const CP_UTF8 As Long = 65001

' --- Environment variable APIs (added in stage 9.5, for the sentinel mechanism) ---
'   Saves "the base address of the region last allocated in Thunks_Init" in a process environment variable
'   as a decimal string. The same school as MouseScroll.bas (Cristian Buse style).
'
'   The W version (UTF-16) is adopted. StrPtr() can pass the VBA String's internal UTF-16 buffer
'   as is, so no A/W conversion is involved.
'
'   SetEnvironmentVariableW(lpName, NULL) means deleting the environment variable.
'   To pass a NULL pointer from VBA, just call with 0 for the second argument
'   (it is a LongPtr ByVal, so 0 is treated as NULL).
Private Declare PtrSafe Function GetEnvironmentVariableW Lib "kernel32" ( _
    ByVal lpName As LongPtr, _
    ByVal lpBuffer As LongPtr, _
    ByVal nSize As Long) As Long

Private Declare PtrSafe Function SetEnvironmentVariableW Lib "kernel32" ( _
    ByVal lpName As LongPtr, _
    ByVal lpValue As LongPtr) As Long

' --- Getting Win32 errors (added in stage 9.5, for diagnostics) ---
'   Gets the error code when the previous Win32 API call failed.
'   Used to isolate the cause of VirtualFree returning 0 (ERROR_BUSY=170 /
'   ERROR_INVALID_ADDRESS=487 / ERROR_INVALID_PARAMETER=87 etc.).
'
'   Note: GetLastError has per-thread state. It is updated every time a Win32 API is
'         called from VBA, so get it right after the API you want to observe.
Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

' --- Structure for VirtualQuery (added in stage 9.5) ---
'   Layout of MEMORY_BASIC_INFORMATION on x64 (48 bytes in total):
'     +0  BaseAddress       LongPtr (8)
'     +8  AllocationBase    LongPtr (8)
'     +16 AllocationProtect Long    (4)
'     +20 (pad1)            Long    (4) <- padding for alignment
'     +24 RegionSize        LongPtr (8)
'     +32 State             Long    (4)
'     +36 Protect           Long    (4)
'     +40 Type              Long    (4)  (named Type_ because it would clash with the VBA reserved word)
'     +44 (pad2)            Long    (4) <- trailing alignment
'   A VBA Type is laid out in declaration order + automatic alignment, so inserting
'   pad1/pad2 explicitly also secures accuracy as documentation.
Private Type MEMORY_BASIC_INFORMATION
    BaseAddress       As LongPtr
    AllocationBase    As LongPtr
    AllocationProtect As Long
    pad1              As Long
    RegionSize        As LongPtr
    state             As Long
    Protect           As Long
    Type_             As Long
    pad2              As Long
End Type

' --- VirtualQuery (added in stage 9.5, for diagnostics) ---
'   Gets the state of the memory page containing the specified address.
'   The return value is the number of bytes of the copied MEMORY_BASIC_INFORMATION (0 on failure).
'   The Type definition must come first, so the Declare is written right after this.
Private Declare PtrSafe Function VirtualQuery Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, _
    ByRef lpBuffer As MEMORY_BASIC_INFORMATION, _
    ByVal dwLength As LongPtr) As LongPtr

' Values of MEMORY_BASIC_INFORMATION.State:
'   MEM_COMMIT  = &H1000  (4096)   <- same value as the existing MEM_COMMIT
'   MEM_RESERVE = &H2000  (8192)   <- same value as the existing MEM_RESERVE
'   MEM_FREE    = &H10000 (65536)
' The values of MEMORY_BASIC_INFORMATION.Protect are judged individually inside FormatMemProtect, so
' individual constant definitions are omitted.

' --- RECT structure (added in stage 9.3c) ---
'   A 16-byte structure passed to ICoreWebView2Controller::put_Bounds.
'   By the x64 ABI, "structures of 16 bytes or less are passed by value", but in real-machine testing
'   the old version worked by passing a pointer with VarPtr(rect), so this implementation
'   follows the same method (= passes the pointer as the LongPtr argument of dcf).
'   Referenced from Wv2Pane, so it is a Public Type.
Public Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Const MEM_COMMIT             As Long = &H1000&
Private Const MEM_RESERVE            As Long = &H2000&
' ^ In stage 9.5g changed "&H8000" to "&H8000&". Without the suffix it
'   was evaluated as the Integer &H8000 (= -32768) and then widened to Long, so
'   -32768 was passed to VirtualFree and it failed with ERROR_INVALID_PARAMETER.
'   This trap is a classic bug that has been stepped on for years in VBA 7 + x64 environments.
Private Const MEM_RELEASE            As Long = &H8000&
' MEM_FREE (added in 9.10a): for the State check of VirtualQuery. Represents a freed region.
'   &H10000 = 65536. Always attach the & suffix (the same kind of prevention as the MEM_RELEASE trap).
Private Const MEM_FREE               As Long = &H10000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40&
' PAGE_READWRITE (added in 9.10a): used by the Sentinel guard to judge "whether it is readable/writable".
'   The thunk region is allocated with PAGE_EXECUTE_READWRITE (&H40), but the aliveness check
'   also accepts PAGE_READWRITE (&H4) as readable/writable, just in case.
Private Const PAGE_READWRITE         As Long = &H4&
Private Const S_OK                   As Long = 0
Private Const E_NOINTERFACE          As Long = &H80004002
Private Const CC_STDCALL             As Long = 4

' --- Constants for the sentinel mechanism (added in stage 9.5) ---
'
'   SENTINEL_ENV_NAME:
'     The name of the process environment variable. Using the same name across
'     multiple Excel workbooks on one machine causes a collision, but in that case it is only that
'     "one side's region cannot be reclaimed", not a crash (= it just falls back to tolerating the leak).
'
'   SENTINEL_BUFFER_SIZE:
'     The number of characters (in WCHAR units) of the buffer received by GetEnvironmentVariableW.
'     An x64 address turned into a decimal string with CStr is at most 19 digits
'     (2^63 - 1 = 9223372036854775807), and 20 characters including the null terminator
'     would suffice, but 32 is taken as a safety margin.
Private Const SENTINEL_ENV_BASE    As String = "WV2_VBA_LastRegion"
Private Const SENTINEL_BUFFER_SIZE As Long = 32

' --- Slot / region layout constants ---

' Position within the EntryPoint stub where the "address to call next" is embedded as an immediate
Private Const LATE_BIND_OFFSET As Long = 55

' Actual length of the stub (confirmed in stage 8.5, 91 bytes including up to C2 00 00 at +88)
Private Const STUB_LEN As Long = 91

' Length of the thunk (flag check 18 + existing 56 = 74 bytes)
Private Const THUNK_LEN As Long = 74

' Offset of the thunk region (stub 91 + padding 5 = 96)
Private Const THUNK_OFFSET   As Long = 96

' Offset of the vtableObj region (= the same value as the old SLOT_SIZE, easy to remember)
'   From this position the WebView2VTable structure (40 bytes) is placed:
'     +0..+7   : pVTable (= points to its own +8)
'     +8..+15  : Functions(0) = Handler_QueryInterface
'     +16..+23 : Functions(1) = Handler_AddRef
'     +24..+31 : Functions(2) = Handler_Release
'     +32..+39 : Functions(3) = pSlot (start of the stub clone)
'
'   Changed to Public Const in stage 9.3a (referenced from Wv2Environment.Init)
Public Const VTABLE_OBJ_OFFSET As Long = 176

' Allocation size per slot (aligned to a multiple of 8, 224 bytes)
'   +0..+90      stub clone (91 bytes)
'   +91..+95     padding
'   +96..+169    thunk (74 bytes)
'   +170..+175   padding (8-byte alignment)
'   +176..+183   vtableObj.pVTable
'   +184..+215   vtableObj.Functions(0..3)
'   +216..+223   spare padding (8-byte alignment)
Private Const SLOT_SIZE      As Long = 224

' Number of slots (fixed)
Private Const SLOT_COUNT     As Long = 512

' Header size at the start of the region (alive flag + future reserve)
'   +0          : alive flag (1 byte, but handled in 8-byte units)
'   +1..+63     : reserved (for future sentinel flag-clearing thunks etc.)
Private Const HEADER_SIZE    As Long = 64

' Size of the whole region (= 64 + 224x512 = 114,752 bytes, about 112 KB)
Private Const REGION_SIZE    As Long = HEADER_SIZE + SLOT_SIZE * SLOT_COUNT


' ============================================================
' GUID structure (Win32 standard, newly added in 9.7a)
'
'   A 16-byte structure passed as a COM REFIID.
'   The riid argument of QueryInterface is a pointer to this structure.
'
'   x64 layout in memory (16 bytes, Win32 standard):
'     +0..+3   : Data1     (Long, little-endian)
'     +4..+5   : Data2     (Integer, little-endian)
'     +6..+7   : Data3     (Integer, little-endian)
'     +8..+15  : Data4(0..7) (each Byte as is, big-endian style)
'
'   Example: the IID {4e8a3389-c9d8-4bd2-b6b5-124fee6cc14d} is
'     Data1 = &H4E8A3389
'     Data2 = &HC9D8 (it is an Integer, so when written in signed form it stays &HC9D8)
'     Data3 = &H4BD2
'     Data4 = (&HB6, &HB5, &H12, &H4F, &HEE, &H6C, &HC1, &H4D)
'
'   GUID comparison is done memcmp-style (binary match of all 16 bytes). In this module a
'   fast version that reads two LongLongs and compares them is implemented in IsEqualGUIDInPlace.
'
'   It is Public because it may be used in the future (version-extension IFs from 9.8 onward etc.)
'   when upper classes issue QueryInterface for a derived IF.
' ============================================================
Public Type GUID
    data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type


' ============================================================
' HandlerKind enumeration (referenceable from upper modules)
'   Identifies the kind of WebView2 callback. Stored in m_kind of ComCallbackHandler.cls
'   and used for branching inside Handler_Invoke.
' ============================================================
Public Enum HandlerKind
    HK_None = 0
    HK_EnvironmentCompleted = 1
    HK_ControllerCompleted = 2
    HK_NavigationStarting = 3
    HK_NavigationCompleted = 4
    HK_WebMessageReceived = 5
    HK_DocumentTitleChanged = 6
    HK_NewWindowRequested = 7
    HK_ExecuteScriptCompleted = 8   ' added in 9.8c, one-shot type, concurrency OK
    HK_HistoryChanged = 9           ' added in 9.9b, persistent, base ICoreWebView2 (vtable 13/14)
    HK_DOMContentLoaded = 10        ' added in 9.9b, persistent, ICoreWebView2_2 (vtable 64/65)
    HK_WebResourceRequested = 11    ' added in N-1, persistent, base ICoreWebView2 (vtable 55/56)
    HK_WebResourceResponseReceived = 12  ' added in N-3, persistent, ICoreWebView2_2 (vtable 61/62)
    HK_ResponseGetContent = 13      ' added in N-4, **one-shot type, concurrency OK**
End Enum


' --- Function addresses obtained once at startup and kept ---
'   Values written to Functions(0..2) of each slot's vtableObj.
'   The AddressOf operator can be written only at an actual-argument position, so it is obtained via GetAddr.
'   Obtained only once in Thunks_Init and reused at Acquire time.
Private m_pHandler_QI       As LongPtr
Private m_pHandler_AddRef   As LongPtr
Private m_pHandler_Release  As LongPtr

' --- Slot pool state ---
Private m_pRegionBase As LongPtr        ' start of the region allocated with VirtualAlloc
Private m_freeHead    As Long           ' head index of the free list (-1 means empty)
Private m_freeNext()  As Long           ' freeNext(i) = the next free slot after slot i
Private m_inUse       As Long           ' number of slots currently in use (for debugging)

' --- Correspondence table: slot index -> Handler object ---
'   From Handler_AddRef / Release / QueryInterface, the path
'   "this (= pSlot + VTABLE_OBJ_OFFSET) -> idx (SlotIndexFromVTableObjAddr)
'    -> m_handlers(idx)" reaches the corresponding ComCallbackHandler instance.
'   This array is also cleared in the automatic release path at refcount = 0.
Private m_handlers(0 To SLOT_COUNT - 1) As ComCallbackHandler

' --- IID table (newly added in stage 9.7a) ---
'   Holds "its own proper IID" for each HandlerKind.
'   Handler_QueryInterface compares riid with it, and returns
'   S_OK + itself if they match, or ppvObject = 0 + E_NOINTERFACE if not.
'
'   Initialized by InitIIDTable at the end of Thunks_Init.
'   The declared range is the full range of HandlerKind (HK_None = 0 to HK_ResponseGetContent = 13).
'   The HK_None entry is unused (not used as a handler, stays at the initial value 0).
Private m_iidTable(HK_None To HK_ResponseGetContent) As GUID

' --- IID_IUnknown (Win32 standard, initialized in Thunks_Init) ---
'   {00000000-0000-0000-C000-000000000046}. Used for the comparison in
'   Handler_QueryInterface. The value is the hard-coded COM standard value.
Private m_iidIUnknown As GUID

Private Const THUNK_BUF_SIZE As Long = 80     ' 74 rounded up to a multiple of 8


' ============================================================
' Source of the EntryPoint stub
'   Copies this stub's byte sequence into the VirtualAlloc region.
'   Must be called once at startup for Force-Compile.
' ============================================================
Private Sub EntryPoint(): End Sub
' ============================================================
' AcquireHandlerFor (newly added in stage 9.3a)
'
'   The integrated entry point that an upper class (Wv2Environment etc.) calls when it wants
'   "one handler of its own". It performs the following in one go:
'     1. Call Thunks_Init() (if not yet done)
'     2. New ComCallbackHandler
'     3. Get the address of ComCallbackHandler.Handler_Invoke from the vTable
'     4. Call Thunks_AcquireSlot to secure a slot + write the thunk + register in m_handlers
'     5. Call ComCallbackHandler.Init(kind, owner, pSlot)
'     6. Return the initialized ComCallbackHandler
'
'   The caller receives the return value as a strong reference and stores it in its own field.
'   After that, the upper class registers the callback by passing ComCallbackHandler.Slot + VTABLE_OBJ_OFFSET to
'   WebView2.
'
'   On failure (pool exhausted etc.) it returns Nothing. The caller must check for Nothing.
' ============================================================
Public Function AcquireHandlerFor( _
    ByVal kind As HandlerKind, _
    ByVal owner As Object) As ComCallbackHandler

    ' --- 1. Pool initialization ---
    If m_pRegionBase = 0 Then
        If Not Thunks_Init() Then Exit Function
    End If

    ' --- 2. New ComCallbackHandler ---
    Dim h As ComCallbackHandler
    Set h = New ComCallbackHandler

    ' --- 3. Get the address of ComCallbackHandler.Handler_Invoke ---
    Dim pHandlerInvoke As LongPtr
    pHandlerInvoke = GetClassMethodAddrAtFixedSlot(h, 7)
    If pHandlerInvoke = 0 Then Exit Function

    ' --- 4. Secure a slot + write the thunk + register in m_handlers ---
    Dim pSlot As LongPtr
    pSlot = Thunks_AcquireSlot(h, ObjPtr(h), pHandlerInvoke)
    If pSlot = 0 Then Exit Function

    ' --- 5. Initialize the ComCallbackHandler itself ---
    h.Init kind, owner, pSlot

    ' --- 6. Return the initialized ComCallbackHandler ---
    Set AcquireHandlerFor = h
End Function


' ============================================================
' EnsureFolder (newly added in stage 9.3a)
'
'   Checks for the existence of a folder such as the user data folder + creates it automatically.
'   Called from Wv2Environment.Init.
'
'   Even for multi-level paths such as "C:\Temp\VBA_WebView2", it creates only one level of parent.
'   If deeper levels are needed, the caller must handle it (currently one level is assumed to suffice).
' ============================================================
Public Sub EnsureFolder(ByVal path As String)
    If LenB(path) = 0 Then Exit Sub

    ' Parent directory
    Dim parentDir As String
    Dim slashPos As Long
    slashPos = InStrRev(path, "\")
    If slashPos > 0 Then
        parentDir = Left$(path, slashPos - 1)
        If LenB(parentDir) > 0 Then
            If Dir(parentDir, vbDirectory) = "" Then MkDir parentDir
        End If
    End If

    ' Itself
    If Dir(path, vbDirectory) = "" Then MkDir path
End Sub


' ============================================================
' FinalizeWebView2Environment (simplified in stage 9.3a)
'
'   The old version included processing to set gHandler to Nothing,
'   but gHandler was abolished, so it now only calls Thunks_Shutdown.
'
'   It is to be called only as the "final cleanup after all handlers and upper classes
'   have been cleaned up". In normal test flows the policy is to call Thunks_Shutdown
'   individually inside each Test (as Test_TwoHandlers_Mock does).
' ============================================================
Public Sub FinalizeWebView2Environment()
    Thunks_Shutdown
End Sub


' ============================================================
' dcf (newly added in stage 9.3b)
'
'   A general-purpose DispCallFunc wrapper. With variable-length arguments (ParamArray) it can
'   call COM methods with 0 to N arguments.
'
'   Points tidied up from the old version of dcf:
'     - Removed the "double acquisition of pVTable / pMethod via CopyMemory".
'       It was dead code whose acquired values were never used (DispCallFunc
'       does the same thing internally).
'     - Removed the "CLngPtr(CLng(...)) inside Case Else". It was unreachable because the
'       preceding Case vbLong: branch had already handled it.
'
'   Arguments:
'     pInterface : COM interface pointer (this)
'     vtblIndex  : vtable slot number (for IUnknown 0=QI, 1=AddRef, 2=Release,
'                  subsequent methods are numbered from 3 in the declaration order of the derived interface)
'     funcName   : for debugging (nothing is printed if empty)
'     args       : variable-length arguments passed to the method (LongPtr/Long/Double/other)
'
'   Argument type inference rules (based on VarType):
'     vbLong       -> passed as a 4-byte integer
'     vbLongLong   -> passed as an 8-byte pointer/integer (Win64 LongPtr compatible)
'     vbDouble     -> passed as an 8-byte double-precision float
'     everything else -> treated as vbLongLong (LongPtr fallback)
'
'   Return value: the return value of the called COM method (cast to Long)
'           - If DispCallFunc itself fails, its hr is returned and Debug.Print is done
'           - On success the COM method's return value is returned as is
'             (both for an HRESULT and for a ULONG return value like AddRef/Release,
'              the caller receives it as is)
'           - No automatic error check assuming "return value = HRESULT" is done:
'             AddRef/Release return the new reference count, so
'             treating value != 0 uniformly as an error causes false detections (confirmed on a real machine).
'             The failure check is done on the caller side when it
'             knows "this is a method that returns an HRESULT".
'
'   Note: when calling a method that has an out argument (e.g. a pp* double pointer), the premise is that
'         the caller passes the VarPtr of a LongPtr variable stored in a Variant.
'         The value stored in the Variant is returned modified, so after the call
'         take it out with CLngPtr etc. Used by Controller::get_CoreWebView2 etc. from 9.3c onward.
' ============================================================
Public Function dcf( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    ByVal funcName As String, _
    ParamArray args() As Variant) As Long

    If pInterface = 0 Then
        Debug.Print "dcf: null interface - " & funcName
        dcf = &H80004003   ' E_POINTER
        Exit Function
    End If

    Dim argc As Long
    argc = UBound(args) - LBound(args) + 1
    If argc < 0 Then argc = 0

    Dim res As Variant
    Dim hr As Long

    If argc = 0 Then
        ' No arguments: pass NULL to vt/vp
        hr = DispCallFunc(pInterface, vtblIndex * PtrSize, _
                          CC_STDCALL, vbLong, _
                          0, ByVal 0&, ByVal 0&, res)
    Else
        ' With arguments: build the vt / vp arrays
        Dim vt() As Integer
        Dim vp() As LongPtr
        Dim vals() As Variant
        ReDim vt(0 To argc - 1)
        ReDim vp(0 To argc - 1)
        ReDim vals(0 To argc - 1)

        Dim i As Long
        For i = 0 To argc - 1
            vals(i) = args(LBound(args) + i)
            Select Case VarType(vals(i))
                Case vbLong:     vt(i) = vbLong
                Case vbLongLong: vt(i) = vbLongLong   ' = vbLongPtr on Win64
                Case vbDouble:   vt(i) = vbDouble
                Case Else:       vt(i) = vbLongLong   ' fallback (treated as LongPtr)
            End Select
            vp(i) = VarPtr(vals(i))
        Next i

        hr = DispCallFunc(pInterface, vtblIndex * PtrSize, _
                          CC_STDCALL, vbLong, _
                          argc, vt(0), vp(0), res)
    End If

    If hr <> 0 Then
        If LenB(funcName) > 0 Then _
            Debug.Print "dcf CALL failed: " & funcName & " hr=&H" & Hex(hr)
        dcf = hr
    Else
        ' On success the COM method's return value is returned as is.
        ' For a method that returns an HRESULT, the caller checks <> 0 and handles the error.
        ' For a ULONG return value like AddRef/Release, the new reference count is returned
        ' (it is not uniformly judged here as "error if <> 0").
        dcf = CLng(res)
    End If
End Function


' ============================================================
' ComRelease / ComAddRef
'
'   A thin wrapper that calls IUnknown::Release / AddRef.
'   It was unified into dcf in 9.3b, but the Public Function is kept for readability
'   (so that calls from Wv2Environment.Class_Terminate etc. can make clear the intent
'   that "Release is being fired").
' ============================================================
Public Function ComRelease(ByVal pInterface As LongPtr) As Long
    If pInterface <> 0 Then ComRelease = dcf(pInterface, 2, "Release")
End Function

Public Function ComAddRef(ByVal pInterface As LongPtr) As Long
    If pInterface <> 0 Then ComAddRef = dcf(pInterface, 1, "AddRef")
End Function


' ============================================================
' PtrToString (added in stage 9.4a, derived from the old mWebView2Helper)
'
'   Converts an LPWSTR (pointer to a UTF-16 null-terminated string) to a VBA String.
'   Used to handle an LPWSTR received as a COM out argument.
'
'   Implementation notes:
'     The old version's implementation using RtlMoveMemory is reused as is (plan beta).
'     String conversion alone departs from our "no CopyMemory" policy, but
'     it is faster than building one character at a time with ChrW via PointerAccessor, and
'     fast enough even for URL-sized strings. A compromise point in stage 9.4a.
'
'   Note:
'     The premise is that the input pointer p was allocated with CoTaskMemAlloc.
'     After obtaining the String with PtrToString, the caller is always
'     responsible for freeing it with CoTaskMemFree(p).
'     GetStringProperty does get -> convert -> free all at once.
' ============================================================
Public Function PtrToString(ByVal p As LongPtr) As String
    If p = 0 Then Exit Function
    Dim cch As Long
    cch = lstrlenW(p)
    If cch = 0 Then Exit Function
    PtrToString = String$(cch, vbNullChar)
    RtlMoveMemory StrPtr(PtrToString), p, CLngPtr(cch * 2)
End Function


' ============================================================
' Utf8BytesToString (added in stage N-2)
'
'   Converts a UTF-8 byte sequence into a VBA String.
'   Used to make the request body (raw bytes read from an IStream) readable.
'
'   Arguments:
'     b  : byte array (0-based)
'     cb : number of bytes to convert (may be shorter than the array's real length = pass only the part cut at the limit)
'
'   Return value: the converted string. Empty string if cb <= 0.
'
'   **Does not fall over on invalid byte sequences**
'     MB_ERR_INVALID_CHARS is not set (dwFlags = 0). Because of cutting at the limit,
'     **being cut in the middle of a multibyte character** happens routinely, and
'     failing there would make "the whole truncated body unreadable".
'     An invalid sequence is replaced with U+FFFD and the rest remains readable.
'
'   **32-bit conversion point** None in particular (LongPtr only).
' ============================================================
Public Function Utf8BytesToString(ByRef b() As Byte, ByVal cb As Long) As String
    If cb <= 0 Then Exit Function

    Dim cch As Long
    cch = MultiByteToWideChar(CP_UTF8, 0, VarPtr(b(LBound(b))), cb, 0, 0)
    If cch <= 0 Then Exit Function

    Dim outStr As String
    outStr = String$(cch, vbNullChar)
    cch = MultiByteToWideChar(CP_UTF8, 0, VarPtr(b(LBound(b))), cb, StrPtr(outStr), cch)
    If cch <= 0 Then Exit Function

    Utf8BytesToString = Left$(outStr, cch)
End Function


' ============================================================
' GetStringProperty (added in stage 9.4a, derived from the old mWebView2Helper)
'
'   Calls a get_*String-type property of a COM interface and returns a String.
'   General-purpose helper: the vtable layout is HRESULT get_Xxx([out, retval] LPWSTR *value)
'   Usable for every method that has it.
'
'   Flow:
'     1. Get the out argument LPWSTR* with dcf (the value is written to pStr)
'     2. Convert to a VBA String with PtrToString
'     3. Free the LPWSTR with CoTaskMemFree (by COM convention, the caller's responsibility)
'
'   Arguments:
'     pInterface : COM pointer (e.g. ICoreWebView2NavigationStartingEventArgs*)
'     vtblIndex  : vtable index of the target method (e.g. get_Uri is 3)
'     funcName   : for debugging (nothing is printed if empty; meaningful only on HRESULT failure)
'
'   Return value: the obtained string (empty string means acquisition failed or an empty string)
'
'   Usage example in 9.4a:
'     Dim url As String
'     url = GetStringProperty(pArgs, 3, "NavStartingArgs.get_Uri")
' ============================================================
Public Function GetStringProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    Optional ByVal funcName As String = "") As String

    If pInterface = 0 Then Exit Function

    Dim pStr As LongPtr
    Dim hr As Long
    hr = dcf(pInterface, vtblIndex, funcName, VarPtr(pStr))
    If hr = 0 And pStr <> 0 Then
        GetStringProperty = PtrToString(pStr)
        CoTaskMemFree pStr
    End If
End Function


' ============================================================
' GetBoolProperty (added in stage 9.4b, derived from the old mWebView2Helper)
'
'   Calls a get_*Bool-type property of a COM interface and returns a Boolean.
'   vtable layout: HRESULT get_Xxx([out, retval] BOOL *value)
'
'   BOOL is a Win32 type, a 4-byte integer (0 = FALSE, anything else = TRUE).
'   To convert to a VBA Boolean a comparison with 0 is needed.
'
'   Usage example in 9.4b:
'     Dim isOk As Boolean
'     isOk = GetBoolProperty(pArgs, 3, "NavCompletedArgs.get_IsSuccess")
' ============================================================
Public Function GetBoolProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    Optional ByVal funcName As String = "") As Boolean

    If pInterface = 0 Then Exit Function

    Dim value As Long
    Dim hr As Long
    hr = dcf(pInterface, vtblIndex, funcName, VarPtr(value))
    If hr = 0 Then GetBoolProperty = (value <> 0)
End Function


' ============================================================
' GetLongProperty (added in stage 9.4b, derived from the old mWebView2Helper)
'
'   Calls a get_*Long-type property of a COM interface and returns a Long.
'   vtable layout: HRESULT get_Xxx([out, retval] long *value)
'
'   Can also be used to obtain an enumeration (e.g. COREWEBVIEW2_WEB_ERROR_STATUS).
' ============================================================
Public Function GetLongProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    Optional ByVal funcName As String = "") As Long

    If pInterface = 0 Then Exit Function

    Dim value As Long
    Dim hr As Long
    hr = dcf(pInterface, vtblIndex, funcName, VarPtr(value))
    If hr = 0 Then GetLongProperty = value
End Function


' ============================================================
' GetLongLongProperty (added in stage 9.9b)
'
'   Calls a get_*UINT64 / get_*INT64-type property of a COM interface and
'   returns a LongLong (signed 64-bit).
'   vtable layout: HRESULT get_Xxx([out, retval] UINT64 *value)
'
'   Purpose:
'     Obtaining ICoreWebView2DOMContentLoadedEventArgs.get_NavigationId (UINT64).
'     NavigationId is a monotonically increasing identifier and in practice fits in the positive range of LongLong.
'
'   ** 32-bit conversion point (policy X, plan a) **
'     32-bit VBA has no LongLong type, so this function is wrapped in #If Win64 and
'     is 64-bit only. In the 32-bit support stage, a separate implementation that receives
'     the UINT64 will be needed. The standard approaches are one of the following:
'       (a) receive it as Currency (8-byte fixed point) and correct by a factor of 10000
'       (b) receive it as 2 Longs (lo/hi) and compose manually
'       (c) NavigationId is for identification only, so compromise by receiving just the low 32 bits as a Long
'     Which to adopt will be designed together in the 32-bit support stage (no getting ahead now).
' ============================================================
#If Win64 Then
Public Function GetLongLongProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    Optional ByVal funcName As String = "") As LongLong

    If pInterface = 0 Then Exit Function

    Dim value As LongLong
    Dim hr As Long
    hr = dcf(pInterface, vtblIndex, funcName, VarPtr(value))
    If hr = 0 Then GetLongLongProperty = value
End Function
#End If


' ============================================================
' Test_TwoHandlers_Mock (the one from stage 9.2, kept as is)
'
'   Core function test of stage 9.2: Acquire 2 handlers,
'   fire AddRef/Release separately from each vtableObj address, and confirm that
'   the reverse lookup of idx works correctly. WebView2 is not called
'   (a pure unit test of the reverse lookup logic).
'
'   Points to confirm in stage 9.3a:
'     - It still passes as before after gHandler was abolished (regression)
'     - It works without error even when ComCallbackHandler is Acquired with owner = Nothing
'     - The removal of the "gHandler Is h" branch in the automatic release path of
'       HandlerReleaseInternal has no effect
'
'   This test does not use Wv2Environment at all. It also does not use AcquireHandlerFor, and
'   keeps the method of calling Thunks_AcquireSlot directly as before
'   (to preserve its nature as a unit test of the foundation).
' ============================================================
Public Sub Test_TwoHandlers_Mock()

    Debug.Print String(60, "=")
    Debug.Print "Test_TwoHandlers_Mock start"
    Debug.Print String(60, "=")

    ' --- 1. Pool initialization ---
    If Not Thunks_Init() Then
        Debug.Print "Thunks_Init failed"
        Exit Sub
    End If

    Debug.Print "Thunks_Init OK, m_inUse = " & m_inUse

    ' --- 2. Create 2 handlers + Acquire ---
    Dim h1 As ComCallbackHandler, h2 As ComCallbackHandler
    Set h1 = New ComCallbackHandler
    Set h2 = New ComCallbackHandler

    Dim pInvoke1 As LongPtr, pInvoke2 As LongPtr
    pInvoke1 = GetClassMethodAddrAtFixedSlot(h1, 7)
    pInvoke2 = GetClassMethodAddrAtFixedSlot(h2, 7)

    Dim pSlot1 As LongPtr, pSlot2 As LongPtr
    pSlot1 = Thunks_AcquireSlot(h1, ObjPtr(h1), pInvoke1)
    pSlot2 = Thunks_AcquireSlot(h2, ObjPtr(h2), pInvoke2)

    If pSlot1 = 0 Or pSlot2 = 0 Then
        Debug.Print "Acquire failed"
        Exit Sub
    End If

    h1.Init HK_EnvironmentCompleted, Nothing, pSlot1
    h2.Init HK_ControllerCompleted, Nothing, pSlot2

    Dim idx1 As Long, idx2 As Long
    idx1 = SlotIndexFromAddr(pSlot1)
    idx2 = SlotIndexFromAddr(pSlot2)
    Debug.Print "h1 -> idx " & idx1 & ", pSlot " & pSlot1
    Debug.Print "h2 -> idx " & idx2 & ", pSlot " & pSlot2
    Debug.Print "m_inUse = " & m_inUse & " (expected 2)"

    If idx1 = idx2 Then
        Debug.Print "[NG] different handlers obtained the same idx"
        Exit Sub
    End If

    ' --- 3. Directly compute the vtableObj address of each handler ---
    Dim pObj1 As LongPtr, pObj2 As LongPtr
    pObj1 = pSlot1 + VTABLE_OBJ_OFFSET
    pObj2 = pSlot2 + VTABLE_OBJ_OFFSET

    ' --- 4. Fire AddRef on each and confirm the counts rise independently ---
    Debug.Print "--- AddRef sequence ---"
    Dim r As Long
    r = Handler_AddRef(pObj1)
    Debug.Print "  Handler_AddRef(pObj1) -> " & r
    r = Handler_AddRef(pObj1)
    Debug.Print "  Handler_AddRef(pObj1) -> " & r
    r = Handler_AddRef(pObj2)
    Debug.Print "  Handler_AddRef(pObj2) -> " & r
    Debug.Print "h1.RefCount = " & h1.RefCount & " (expected 2)"
    Debug.Print "h2.RefCount = " & h2.RefCount & " (expected 1)"

    If h1.RefCount <> 2 Or h2.RefCount <> 1 Then
        Debug.Print "[NG] AddRef was dispatched differently from expected"
    Else
        Debug.Print "[OK] AddRef was dispatched as expected"
    End If

    ' --- 5. Fire Release and confirm each independently reaches refcount=0 ---
    Debug.Print "--- Release sequence ---"
    r = Handler_Release(pObj1)
    Debug.Print "  Handler_Release(pObj1) -> " & r
    Debug.Print "h1.RefCount = " & h1.RefCount & " (expected 1)"
    Debug.Print "m_inUse = " & m_inUse & " (expected 2, both still in use)"

    r = Handler_Release(pObj1)
    Debug.Print "  Handler_Release(pObj1) -> " & r
    Debug.Print "(after h1 release) m_inUse = " & m_inUse & " (expected 1)"

    Debug.Print "h1.RefCount (after release) = " & h1.RefCount & " (expected 0)"
    Debug.Print "h2.RefCount (independent of h1 release) = " & h2.RefCount & " (expected 1)"

    r = Handler_Release(pObj2)
    Debug.Print "  Handler_Release(pObj2) -> " & r
    Debug.Print "(after h2 release) m_inUse = " & m_inUse & " (expected 0)"
    Debug.Print "h2.RefCount (after release) = " & h2.RefCount & " (expected 0)"

    ' --- 6. Confirm that operations on released slots pass safely as no-ops ---
    Debug.Print "--- Post-release safety check ---"
    r = HandlerAddRefInternal(pObj1)
    Debug.Print "Return value of AddRef on released pObj1 = " & r & " (expected 0)"
    r = HandlerReleaseInternal(pObj2)
    Debug.Print "Return value of Release on released pObj2 = " & r & " (expected 0)"

    ' --- 7. Cleanup ---
    Set h1 = Nothing
    Set h2 = Nothing

    Debug.Print "--- Final state ---"
    Debug.Print "m_inUse = " & m_inUse & " (expected 0)"
    Debug.Print "m_freeHead = " & m_freeHead

    Thunks_Shutdown

    Debug.Print "Test_TwoHandlers_Mock done"
    Debug.Print String(60, "=")
End Sub


' ============================================================
' Test_Environment_Real (newly added in stage 9.3a)
'
'   Integration test using the Wv2Environment class.
'   Verifies the whole flow through the real WebView2Loader.dll: Environment creation -> receiving the completion notification ->
'   AddRef -> Release in Class_Terminate.
'
'   Verification items:
'     1. ComCallbackHandler is Acquired once in Wv2Environment.Init and the API is called
'     2. Invoke arrives from WebView2 and Wv2Environment.OnEnvironmentCompleted is called
'        via ComCallbackHandler.Handler_Invoke (dynamic binding)
'     3. errorCode = 0 and pEnvironment != 0 are received and it becomes IsReady
'     4. m_handler is cut to Nothing (the foundation's automatic release path runs)
'     5. AddRef on m_pEnvironment succeeds (= validity such as returning 1)
'     6. Class_Terminate runs at End Sub and the reference is released by ComRelease
'
'   The timeout is 5 seconds. If WebView2 startup takes time (first launch etc.)
'   it may need to be extended.
' ============================================================
Public Sub Test_Environment_Real()

    Debug.Print String(60, "=")
    Debug.Print "Test_Environment_Real start"
    Debug.Print String(60, "=")

    Dim env As Wv2Environment
    Set env = New Wv2Environment

    Debug.Print "Calling env.Init ..."
    env.Init Wv2Log.AppFolder("Profile")

    If env.IsFailed Then
        Debug.Print "[NG] env.Init failed immediately (LastError = &H" & Hex(env.LastError) & ")"
        GoTo Cleanup
    End If

    Debug.Print "State right after Init = " & env.state & _
                " (1=Es_Waiting / 2=Es_Ready synchronous callback / 3=Es_Failed)"
    Debug.Print "Waiting for the completion notification (up to 5 s, exits immediately if the synchronous callback has already happened) ..."

    ' The wait loop judges only by IsReady / IsFailed.
    ' (Referencing Wv2Environment's Public Enum from a standard module is
    '  implementation-dependent, so going through Boolean properties is safer)
    Dim t As Single: t = Timer
    Do While (Not env.IsReady) And (Not env.IsFailed) And (Wv2Log.SecondsSince(t) < 5#)
        DoEvents
    Loop

    If env.IsReady Then
        Debug.Print "[OK] env.IsReady = True"
        Debug.Print "  m_pEnvironment   = " & env.EnvironmentPtr
        Debug.Print "  elapsed time      = " & Format$(Timer - t, "0.000") & " s"
    ElseIf env.IsFailed Then
        Debug.Print "[NG] env.IsFailed = True (LastError = &H" & Hex(env.LastError) & ")"
    Else
        Debug.Print "[NG] timeout (state = " & env.state & ")"
    End If

Cleanup:
    Debug.Print "--- Releasing env (ComRelease is expected to run in Class_Terminate) ---"
    Set env = Nothing
    Debug.Print "Set env = Nothing done"

    ' Finally release the pool region too with Thunks_Shutdown
    Debug.Print "Running Thunks_Shutdown"
    Thunks_Shutdown

    Debug.Print "Test_Environment_Real done"
    Debug.Print String(60, "=")
End Sub


' ============================================================
' Test_ResetSurvival_Manual (newly added in stage 9.4a)
'
'   A real-machine verification guide for reset button resistance (plan R4).
'   This procedure itself only logs the steps; the actual verification is
'   done by hand by the user in the VBE.
'
'   Verification procedure (done by the user on a real machine):
'     1. Run this procedure Test_ResetSurvival_Manual (shows the guidance)
'     2. Run UserForm1.Show vbModeless -> UserForm1.StartWebView2_Full to start
'        showing Bing + registering the standard handlers (including NavigationStarting)
'     3. Once Bing is shown, click a link in the page once or twice and confirm with Debug.Print
'        that the NavigationStarting event fires correctly
'     4. **In that state press the VBE stop button (reset)**
'        -> module variables such as m_pRegionBase are reset to the initial value 0
'        -> however the VirtualAlloc region is not reclaimed by the OS, so it remains
'        -> the alive flag also remains (Thunks_Shutdown was not called)
'     5. Confirm that Excel does not go down
'     6. Wait a while (or try clicking another link) and
'        confirm Excel does not crash even when WebView2 hits the ComCallbackHandler thunk
'        (Note: with the alive flag mechanism the region remains even after reset, so
'         the thunk may try the real processing instead of returning immediately. Whether this
'         leads to a crash is the purpose of this verification)
'     7. Confirm that Test_Navigate_Real can be re-run in the same session
'        (Thunks_Init does "allocate new if m_pRegionBase = 0", so
'         a new region is allocated separately from the past leaked region)
'
'   Expected results:
'     - Excel does not crash after reset
'     - Re-running within the same session is possible
'     - From stage 9.5 onward, the sentinel mechanism automatically
'       VirtualFrees the old region on re-run (leak completely eliminated)
'
'   Use Test_SentinelStatus together to observe the sentinel mechanism's behavior.
' ============================================================
Public Sub Test_ResetSurvival_Manual()
    Debug.Print String(60, "=")
    Debug.Print "Test_ResetSurvival_Manual (reset test guide)"
    Debug.Print String(60, "=")
    Debug.Print "Verification procedure for reset button resistance:"
    Debug.Print ""
    Debug.Print "1. Separately run Test_Navigate_Real to show Bing"
    Debug.Print "2. Click a Bing link once or twice and confirm that the NavigationStarting event is"
    Debug.Print "   Debug.Printed"
    Debug.Print "3. In that state press the VBE stop button (reset)"
    Debug.Print "4. Confirm that Excel does not go down"
    Debug.Print "5. Wait a while and confirm that even trying to click a link"
    Debug.Print "   does not crash Excel"
    Debug.Print "6. Confirm that Test_Navigate_Real can be re-run in the same session"
    Debug.Print ""
    Debug.Print "Expected results:"
    Debug.Print "  - Excel does not crash after reset (effect of the alive flag)"
    Debug.Print "  - Re-running within the same session is possible"
    Debug.Print "  - From stage 9.5 onward, the sentinel mechanism"
    Debug.Print "    automatically VirtualFrees the old region on re-run (leak completely eliminated)"
    Debug.Print ""
    Debug.Print "To check the sentinel mechanism's behavior, also use Test_SentinelStatus:"
    Debug.Print "  - Called right after a reset, the old region address should remain in the environment variable"
    Debug.Print "  - Called after re-Init, it should have been replaced by the new region address"
    Debug.Print String(60, "=")
End Sub


' ============================================================
' Sentinel_LoadPrevRegion (added in stage 9.5)
'
'   Reads "the base address of the region allocated by the previous Thunks_Init" from
'   the environment variable SENTINEL_ENV_NAME. If there is no trace (= it ended normally, or
'   this is the first launch) it returns 0.
'
'   Note: this environment variable normally does not exist right after process startup.
'      It exists only when "Shutdown did not run because of a reset just before".
'      (Environment variables are per process, so a value from another Excel process
'       cannot get mixed in)
'
'   Return value: base address of the previous region (0 if none)
' ============================================================
' ============================================================
' SentinelEnvName (English fork - fixes a cross-workbook crash)
'   The sentinel variable used to be the fixed name WV2_VBA_LastRegion. The
'   environment block is shared by every VBA project in the Excel process, so
'   a second workbook containing this module read the FIRST workbook's live
'   region address, found it alive, zeroed its alive flag and VirtualFree'd
'   it -> the next WebView2 callback into the first workbook crashed Excel.
'   The name is now suffixed with a hash of this workbook's full path, so
'   each project only ever recovers its own region (a reset keeps the path,
'   so recovery after a reset still works).
' ============================================================
Private Function SentinelEnvName() As String
    Dim s As String, h As Double, i As Long
    On Error Resume Next
    s = LCase$(ThisWorkbook.FullName)
    On Error GoTo 0
    For i = 1 To Len(s)
        h = h * 31# + (AscW(Mid$(s, i, 1)) And &HFFFF&)
        h = h - Int(h / 2147483629#) * 2147483629#
    Next i
    SentinelEnvName = SENTINEL_ENV_BASE & "_" & Hex$(CLng(h))
End Function


Private Function Sentinel_LoadPrevRegion() As LongPtr
    Sentinel_LoadPrevRegion = 0^

    Dim buff As String
    buff = String$(SENTINEL_BUFFER_SIZE, vbNullChar)

    Dim n As Long
    n = GetEnvironmentVariableW(StrPtr(SentinelEnvName()), StrPtr(buff), SENTINEL_BUFFER_SIZE)

    ' n = 0 : the variable does not exist or acquisition failed
    If n = 0 Then Exit Function

    ' n >= SENTINEL_BUFFER_SIZE : buffer too small (= abnormal value, ignored)
    '   On success GetEnvironmentVariableW returns a value less than nSize; when the buffer is insufficient it returns the required
    '   buffer size. An address string never exceeds 19 digits, so
    '   if we come to this path the environment variable is judged to be corrupted.
    If n >= SENTINEL_BUFFER_SIZE Then Exit Function

    Dim s As String
    s = Left$(buff, n)

    ' If numeric conversion fails (empty string, non-numeric contamination, etc.) return 0
    On Error Resume Next
    Sentinel_LoadPrevRegion = CLngLng(s)
    On Error GoTo 0
End Function


' ============================================================
' Sentinel_StorePrevRegion (added in stage 9.5)
'
'   Writes the current region base address into the environment variable.
'   The format is a decimal string (follows MouseScroll.bas).
'
'   Call timing:
'     Right after VirtualAlloc succeeds in Thunks_Init. This way, even if a reset
'     occurs midway through Init, the region is reliably reclaimed at the next Init.
' ============================================================
Private Sub Sentinel_StorePrevRegion(ByVal addr As LongPtr)
    Dim s As String
    s = CStr(addr)
    SetEnvironmentVariableW StrPtr(SentinelEnvName()), StrPtr(s)
End Sub


' ============================================================
' Sentinel_ClearPrevRegion (added in stage 9.5)
'
'   Clears the environment variable.
'   Win32 spec: SetEnvironmentVariableW(lpName, NULL) deletes the variable.
'   From VBA, call it with ByVal 0 for the second argument (passing 0 with a LongPtr ByVal
'   is treated as a NULL pointer as is).
'
'   Call timing:
'     1. At the end of Thunks_Shutdown (normal termination path)
'     2. At the end of Sentinel_RecoverIfNeeded (after recovering from a reset, before the new Init begins)
' ============================================================
Private Sub Sentinel_ClearPrevRegion()
    SetEnvironmentVariableW StrPtr(SentinelEnvName()), 0
End Sub


' ============================================================
' Sentinel_RecoverIfNeeded (added in stage 9.5)
'
'   Called at the start of Thunks_Init. If there is a trace of the previous run, the old region is reclaimed.
'
'   Reclamation procedure (plan beta, double defense of alive flag + VirtualFree):
'     1. Set the old region's alive flag to 0
'        - Even if WebView2 hits the stub meanwhile, it returns S_OK at the start of the thunk
'        - A momentary safety valve until VirtualFree completes
'        - If the region has already been freed/remapped, an access violation is possible,
'          so defend with On Error
'     2. Free the old region with VirtualFree
'        - On failure, isolate the cause with GetLastError
'     3. Clear the environment variable
'        - Thunks_Init overwrites it with the new region address right after, but clear it first so
'          that no old value remains if VirtualAlloc should fail
'
'   Guarantee that the reclamation target is "only regions abandoned within the same process":
'     - Environment variables are per process (they are copied parent -> child, but are never
'       read from other processes)
'     - Excel is not restarted by a VBA reset (the same process continues)
'     - Therefore a trace present = abandonment within the same process is certain
'
'   Diagnostic log (kept because it helps isolate future incidents):
'     - "VirtualQuery result ...": the region state as seen from the OS
'         State = MEM_COMMIT (4096)  -> region alive (normal)
'         State = MEM_FREE   (65536) -> already freed
'         State = MEM_RESERVE(8192)  -> intermediate state
'     - "header LongPtr before clear: X" : the alive flag at the start of the old region
'         X = 1   -> region alive, flag still set (normal)
'         X = 0   -> someone already cleared the flag
'         access violation -> the region has already been freed/remapped
'     - "VirtualFree ... LastError: Y" : the result of VirtualFree
'         Y = 0   -> normally we come here (succeeded)
'         Y = 87  -> ERROR_INVALID_PARAMETER (root cause resolved in stage 9.5,
'                  if it recurs suspect a forgotten & suffix on a constant)
'         Y = 170 -> ERROR_BUSY (another thread is running code in the region)
'         Y = 487 -> ERROR_INVALID_ADDRESS (invalid address, already freed, etc.)
' ============================================================
Private Sub Sentinel_RecoverIfNeeded()
    Dim prevBase As LongPtr
    prevBase = Sentinel_LoadPrevRegion()
    If prevBase = 0 Then Exit Sub

    Debug.Print "Sentinel: detected previous region = " & prevBase & _
                " (decimal), recovering ..."

    ' --- Diagnostic 0 + aliveness check (added in stage 9.5 / promoted to an aliveness check in 9.10a) ---
    '   Observe the region state as seen from the OS with VirtualQuery, and use the result to judge "whether the region is alive".
    '   This isolates the following:
    '     State = MEM_COMMIT (4096)  -> region alive (suspect VirtualFree's arguments)
    '     State = MEM_FREE   (65536) -> already freed (someone Freed it first)
    '     State = MEM_RESERVE(8192)  -> only Reserve remains (intermediate state)
    '   Also check whether AllocationBase matches prevBase (= the true start of the region).
    '
    '   ** The most important fix of 9.10a (handling spec fact 10) **
    '     The old implementation "observed State with VirtualQuery" yet did not use the result
    '     for branching and unconditionally ran ReadLongPtr / MemLongPtr / VirtualFree.
    '     Firing ReadLongPtr / MemLongPtr at a freed address with State=MEM_FREE causes an
    '     AV (access violation) in native array access through the SAFEARRAY.
    '     This AV is a CPU SEH exception and cannot be caught by VBA's On Error Resume Next
    '     (= the Excel process dies immediately, spec fact 10).
    '     Therefore only when VirtualQuery confirms "alive" do we run
    '     ReadLongPtr / MemLongPtr / VirtualFree (design principle 55).
    Dim mbi As MEMORY_BASIC_INFORMATION
    Dim qSize As LongPtr
    Dim regionAlive As Boolean
    regionAlive = False

    qSize = VirtualQuery(prevBase, mbi, LenB(mbi))
    If qSize = 0 Then
        Debug.Print "Sentinel: VirtualQuery FAILED, LastError = " & GetLastError()
    Else
        Debug.Print "Sentinel: VirtualQuery result (" & qSize & " bytes copied):"
        Debug.Print "  BaseAddress       = " & mbi.BaseAddress
        Debug.Print "  AllocationBase    = " & mbi.AllocationBase & _
                    IIf(mbi.AllocationBase = prevBase, " (= prevBase, OK)", " (<> prevBase, mismatch)")
        Debug.Print "  AllocationProtect = &H" & Hex$(mbi.AllocationProtect) & _
                    " (" & FormatMemProtect(mbi.AllocationProtect) & ")"
        Debug.Print "  RegionSize        = " & mbi.RegionSize
        Debug.Print "  State             = " & mbi.state & " (" & FormatMemState(mbi.state) & ")"
        Debug.Print "  Protect           = &H" & Hex$(mbi.Protect) & _
                    " (" & FormatMemProtect(mbi.Protect) & ")"
        Debug.Print "  Type              = &H" & Hex$(mbi.Type_)

        ' Aliveness check (conservative): regarded as "alive" only when all 3 conditions are met.
        '   (1) State = MEM_COMMIT       : the page is committed (not merely freed/reserved)
        '   (2) AllocationBase = prevBase : the true start of the region and not remapped
        '                                   to another region
        '   (3) Protect is readable/writable : not PAGE_GUARD / PAGE_NOACCESS.
        '       The thunk region is allocated with PAGE_EXECUTE_READWRITE, but just in case
        '       PAGE_READWRITE is also allowed. If the PAGE_GUARD bit is set, the first
        '       access gives STATUS_GUARD_PAGE_VIOLATION, so it is excluded.
        regionAlive = (mbi.state = MEM_COMMIT) _
                      And (mbi.AllocationBase = prevBase) _
                      And IsProtectReadWritable(mbi.Protect)
    End If

    If Not regionAlive Then
        ' The region is not alive (= already freed / remapped / reserved only / not accessible).
        ' Never fire ReadLongPtr / MemLongPtr / VirtualFree; only clear the environment variable
        ' and exit safely. Why VirtualFree is not fired either: MEM_RELEASE on a freed region
        ' only gives ERROR_INVALID_PARAMETER and does no real harm, but
        ' simplifying the path by sticking to "do not touch it if it is not alive".
        Debug.Print "Sentinel: old region judged not alive -> skipped without touching (handling spec fact 10)"
        Sentinel_ClearPrevRegion
        Exit Sub
    End If

    ' --- From here on, VirtualQuery has confirmed that the region is alive ---

    ' Diagnostic 1: read the alive flag at the start of the region (alive is confirmed, so no AV occurs)
    '   Note: MemLongPtr is Property Let only (write-only). For reading,
    '       use the ReadLongPtr function.
    Dim headerByte As LongPtr
    headerByte = ReadLongPtr(prevBase)
    Debug.Print "Sentinel: header LongPtr before clear = " & headerByte & _
                " (alive flag = " & (headerByte And &HFF&) & ")"

    ' 1. Set the alive flag to 0 (double defense)
    MemLongPtr(prevBase) = 0^

    ' 2. Free the old region with VirtualFree (also get the error code with GetLastError)
    '    Win32 spec: VirtualFree returns 0 on failure, and the cause is obtained with GetLastError.
    Dim freeResult As Long
    Dim lastErr As Long
    freeResult = VirtualFree(prevBase, 0, MEM_RELEASE)
    lastErr = GetLastError()
    If freeResult <> 0 Then
        Debug.Print "Sentinel: VirtualFree succeeded for " & prevBase
    Else
        Debug.Print "Sentinel: VirtualFree returned 0, LastError = " & lastErr & _
                    " (" & VirtualFreeErrorName(lastErr) & ")"
    End If

    ' 3. Clear the environment variable first (before writing the new region address)
    Sentinel_ClearPrevRegion
End Sub

' ============================================================
' IsProtectReadWritable (added in 9.10a)
' ============================================================
'   Judges whether MEMORY_BASIC_INFORMATION.Protect is "a protection attribute safe to read/write from VBA".
'   Part of the aliveness check in Sentinel_RecoverIfNeeded (design principle 55).
'
'   Regarded as readable/writable: PAGE_READWRITE (&H4) / PAGE_EXECUTE_READWRITE (&H40).
'   However, if the PAGE_GUARD (&H100) bit is set, a guard page exception occurs,
'   so in that case it is treated as not readable/writable.
Private Function IsProtectReadWritable(ByVal prot As Long) As Boolean
    ' Not allowed if the PAGE_GUARD bit is set
    If (prot And &H100&) <> 0 Then
        IsProtectReadWritable = False
        Exit Function
    End If
    ' Judge by the basic protection attribute with the guard bit removed
    Dim base As Long
    base = prot And Not &H100& And Not &H200& And Not &H400&
    Select Case base
        Case PAGE_READWRITE, PAGE_EXECUTE_READWRITE
            IsProtectReadWritable = True
        Case Else
            IsProtectReadWritable = False
    End Select
End Function


' ============================================================
' VirtualFreeErrorName (added in stage 9.5, for diagnostics)
'
'   Converts a GetLastError value into a Win32 error name. Only the representative ones
'   are judged individually; unknown values return "UNKNOWN".
' ============================================================
Private Function VirtualFreeErrorName(ByVal errCode As Long) As String
    Select Case errCode
        Case 0: VirtualFreeErrorName = "ERROR_SUCCESS (= actually succeeded?)"
        Case 5: VirtualFreeErrorName = "ERROR_ACCESS_DENIED"
        Case 6: VirtualFreeErrorName = "ERROR_INVALID_HANDLE"
        Case 8: VirtualFreeErrorName = "ERROR_NOT_ENOUGH_MEMORY"
        Case 87: VirtualFreeErrorName = "ERROR_INVALID_PARAMETER"
        Case 170: VirtualFreeErrorName = "ERROR_BUSY (in use by another thread)"
        Case 487: VirtualFreeErrorName = "ERROR_INVALID_ADDRESS"
        Case Else: VirtualFreeErrorName = "UNKNOWN"
    End Select
End Function


' ============================================================
' FormatMemState (added in stage 9.5, for diagnostics)
'
'   Turns the MEMORY_BASIC_INFORMATION.State value into a string.
' ============================================================
Private Function FormatMemState(ByVal state As Long) As String
    Select Case state
        Case &H1000: FormatMemState = "MEM_COMMIT (region alive)"
        Case &H2000: FormatMemState = "MEM_RESERVE (Reserve only, not Committed)"
        Case &H10000: FormatMemState = "MEM_FREE (already freed)"
        Case Else: FormatMemState = "UNKNOWN(&H" & Hex$(state) & ")"
    End Select
End Function


' ============================================================
' FormatMemProtect (added in stage 9.5, for diagnostics)
'
'   Turns the MEMORY_BASIC_INFORMATION.Protect / AllocationProtect value into a string.
'   Flags such as PAGE_GUARD may be bit-ORed, so they are detected individually.
' ============================================================
Private Function FormatMemProtect(ByVal prot As Long) As String
    If prot = 0 Then
        FormatMemProtect = "NONE (= Protect=0 for a freed region)"
        Exit Function
    End If

    Dim s As String

    Select Case prot And &HFF&
        Case &H1: s = "PAGE_NOACCESS"
        Case &H2: s = "PAGE_READONLY"
        Case &H4: s = "PAGE_READWRITE"
        Case &H8: s = "PAGE_WRITECOPY"
        Case &H10: s = "PAGE_EXECUTE"
        Case &H20: s = "PAGE_EXECUTE_READ"
        Case &H40: s = "PAGE_EXECUTE_READWRITE"
        Case &H80: s = "PAGE_EXECUTE_WRITECOPY"
        Case Else: s = "UNKNOWN(&H" & Hex$(prot And &HFF&) & ")"
    End Select

    If (prot And &H100&) <> 0 Then s = s & "|PAGE_GUARD"
    If (prot And &H200&) <> 0 Then s = s & "|PAGE_NOCACHE"
    If (prot And &H400&) <> 0 Then s = s & "|PAGE_WRITECOMBINE"

    FormatMemProtect = s
End Function


' ============================================================
' Test_VirtualAllocFree_Roundtrip (added in stage 9.5)
'
'   Observes, in 5 kinds of scenarios, whether VirtualAlloc -> VirtualFree succeeds
'   in the normal flow that does not go through a VBA reset.
'   The purpose is to isolate the root cause of the phenomenon found in 9.5c, that "only a region that went
'   through a reset fails with ERROR_INVALID_PARAMETER".
'
'   How to use:
'     1. Restart Excel + VBE to get a clean state (optional)
'     2. Run Test_VirtualAllocFree_Roundtrip in the Immediate window
'     3. Observe the results of alpha to delta
'     4. (Optional) After that, run Thunks_Init with Test_AllEvents_Real etc. and
'        run Test_VirtualAllocFree_Eps with m_pRegionBase allocated,
'        which lets you observe scenario epsilon as well
'
'   If every scenario succeeds -> only the region that went through a reset fails specifically
'                       (the theory that the VBA runtime puts some mark on it is the leading one)
'   If any fails   -> the element included in that scenario is the root cause
' ============================================================
Public Sub Test_VirtualAllocFree_Roundtrip()
    Debug.Print String(60, "=")
    Debug.Print "Test_VirtualAllocFree_Roundtrip (check without going through a reset)"
    Debug.Print String(60, "=")

    Dim p As LongPtr, p2 As LongPtr
    Dim r As Long
    Dim errCode As Long

    ' --- Scenario alpha: alloc -> free immediately ---
    Debug.Print "--- Scenario alpha: alloc -> free immediately ---"
    p = VirtualAlloc(0, REGION_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    Debug.Print "  VirtualAlloc -> " & p
    If p = 0 Then
        Debug.Print "  [SKIP] alloc failed, skipping the following scenarios too"
        Exit Sub
    End If
    r = VirtualFree(p, 0, MEM_RELEASE)
    errCode = GetLastError()
    Debug.Print "  VirtualFree returned " & r & ", LastError = " & errCode & _
                IIf(r <> 0, " [OK]", " [NG] (" & VirtualFreeErrorName(errCode) & ")")

    ' --- Scenario beta: alloc -> write contents -> free ---
    Debug.Print "--- Scenario beta: alloc -> write contents -> free ---"
    p = VirtualAlloc(0, REGION_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    Debug.Print "  VirtualAlloc -> " & p
    If p <> 0 Then
        MemLongPtr(p) = 1^                       ' alive-flag-like write
        MemLongPtr(p + 8) = &H42^                ' an arbitrary value (safe within the Long range)
        Debug.Print "  alive flag written (header LongPtr = " & ReadLongPtr(p) & ")"
        r = VirtualFree(p, 0, MEM_RELEASE)
        errCode = GetLastError()
        Debug.Print "  VirtualFree returned " & r & ", LastError = " & errCode & _
                    IIf(r <> 0, " [OK]", " [NG] (" & VirtualFreeErrorName(errCode) & ")")
    End If

    ' --- Scenario gamma: alloc -> round-trip conversion via decimal string -> free ---
    Debug.Print "--- Scenario gamma: alloc -> address decimal round-trip conversion -> free ---"
    p = VirtualAlloc(0, REGION_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    Debug.Print "  VirtualAlloc -> " & p
    If p <> 0 Then
        Dim s As String
        s = CStr(p)
        Dim pRestored As LongPtr
        pRestored = CLngLng(s)
        Debug.Print "  CStr/CLngLng round-trip: " & p & " -> '" & s & "' -> " & pRestored & _
                    IIf(p = pRestored, " [match]", " [MISMATCH!]")
        r = VirtualFree(pRestored, 0, MEM_RELEASE)
        errCode = GetLastError()
        Debug.Print "  VirtualFree(pRestored) returned " & r & ", LastError = " & errCode & _
                    IIf(r <> 0, " [OK]", " [NG] (" & VirtualFreeErrorName(errCode) & ")")
    End If

    ' --- Scenario delta: alloc(A) -> alloc(B) -> free(A) -> free(B) ---
    '   The situation closest to the current Sentinel_RecoverIfNeeded. However, no reset
    '   is involved.
    Debug.Print "--- Scenario delta: allocate 2 regions -> free the older first -> free the newer ---"
    p = VirtualAlloc(0, REGION_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    Debug.Print "  region A alloc -> " & p
    p2 = VirtualAlloc(0, REGION_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    Debug.Print "  region B alloc -> " & p2
    If p <> 0 And p2 <> 0 Then
        MemLongPtr(p) = 1^                     ' alive-flag-like write to region A
        r = VirtualFree(p, 0, MEM_RELEASE)
        errCode = GetLastError()
        Debug.Print "  VirtualFree(A) returned " & r & ", LastError = " & errCode & _
                    IIf(r <> 0, " [OK]", " [NG] (" & VirtualFreeErrorName(errCode) & ")")
        r = VirtualFree(p2, 0, MEM_RELEASE)
        errCode = GetLastError()
        Debug.Print "  VirtualFree(B) returned " & r & ", LastError = " & errCode & _
                    IIf(r <> 0, " [OK]", " [NG] (" & VirtualFreeErrorName(errCode) & ")")
    End If

    ' --- Scenario epsilon: VirtualFree on the real Thunks_Init region ---
    '   If m_pRegionBase is already allocated (= Thunks_Init done),
    '   try VirtualQuery -> VirtualFree on that address.
    '   Note: running this does not go through Thunks_Shutdown, so
    '         m_pRegionBase does not become 0 (= risk of double free).
    '         This scenario is meaningful only in the state after Thunks_Shutdown,
    '         so run it only when m_pRegionBase = 0.
    Debug.Print "--- Scenario epsilon: a region that went through the real Thunks_Init ---"
    If m_pRegionBase <> 0 Then
        Debug.Print "  [SKIP] m_pRegionBase = " & m_pRegionBase & " (= already Init-ed)"
        Debug.Print "         Please re-run this scenario after Shutdown (= m_pRegionBase = 0)"
    Else
        Debug.Print "  Running Thunks_Init ..."
        If Not Thunks_Init() Then
            Debug.Print "  [SKIP] Thunks_Init failed"
        Else
            Dim baseBak As LongPtr
            baseBak = m_pRegionBase
            Debug.Print "  m_pRegionBase = " & baseBak
            ' Try VirtualFree directly without calling Thunks_Shutdown.
            ' To keep the state consistent afterwards, clear module variables such as m_pRegionBase
            ' manually and also call Sentinel_ClearPrevRegion.
            r = VirtualFree(baseBak, 0, MEM_RELEASE)
            errCode = GetLastError()
            Debug.Print "  VirtualFree(real) returned " & r & ", LastError = " & errCode & _
                        IIf(r <> 0, " [OK]", " [NG] (" & VirtualFreeErrorName(errCode) & ")")
            ' Cleanup
            m_pRegionBase = 0
            m_freeHead = -1
            m_inUse = 0
            Erase m_freeNext
            Sentinel_ClearPrevRegion
            Debug.Print "  cleanup done (m_pRegionBase = 0, environment variable cleared)"
        End If
    End If

    Debug.Print String(60, "=")
    Debug.Print "Test_VirtualAllocFree_Roundtrip done"
    Debug.Print String(60, "=")
End Sub


' ============================================================
' Test_SentinelStatus (added in stage 9.5, for debugging)
'
'   A test function that merely Debug.Prints the current environment variable contents and m_pRegionBase.
'   Can be run as "Test_SentinelStatus" from the Immediate window.
'
'   Points to observe:
'     - Right after first launch: no environment variable, m_pRegionBase = 0
'     - After StartWebView2: environment variable = m_pRegionBase = new region address
'     - After normal termination: no environment variable, m_pRegionBase = 0
'     - Right after a reset: environment variable = old region address, m_pRegionBase = 0
'         (VBA variables vanish on reset but environment variables survive within the process)
'     - After re-Init following a reset: environment variable = m_pRegionBase = new region address
'         (the old region has already been VirtualFreed)
' ============================================================
Public Sub Test_SentinelStatus()
    Dim prev As LongPtr
    prev = Sentinel_LoadPrevRegion()

    Debug.Print String(60, "=")
    Debug.Print "Test_SentinelStatus"
    Debug.Print "  Env var name : " & SentinelEnvName()
    If prev = 0 Then
        Debug.Print "  Env var value: (empty or unset) -> no leftover region"
    Else
        Debug.Print "  Env var value: " & prev & " (decimal pointer)"
        If prev = m_pRegionBase Then
            Debug.Print "  -> matches the region currently in use (normal state)"
        ElseIf m_pRegionBase = 0 Then
            Debug.Print "  -> scheduled to be reclaimed at the next Thunks_Init (after a reset)"
        Else
            Debug.Print "  -> the environment variable and m_pRegionBase disagree (abnormal)"
        End If
    End If
    Debug.Print "  m_pRegionBase: " & m_pRegionBase
    Debug.Print String(60, "=")
End Sub


' ============================================================
' Thunks_Init
'   Allocates the region with VirtualAlloc, sets the header's alive flag,
'   copies the stub into all slots, initializes vtableObj,
'   and initializes the free list.
'
'   Processing in stage 9.2 (unchanged in 9.3a):
'     - Get the addresses of Handler_QI / AddRef / Release only once
'     - Initialize the vtableObj of all 512 slots in one go
'         pVTable      = pFunctions
'         Functions(0) = m_pHandler_QI
'         Functions(1) = m_pHandler_AddRef
'         Functions(2) = m_pHandler_Release
'         Functions(3) = pSlot (= start of the stub clone)
'
'   Processing added in stage 9.5:
'     - Call Sentinel_RecoverIfNeeded at the start
'         (if a region was abandoned at the previous reset, VirtualFree it and reclaim)
'     - Call Sentinel_StorePrevRegion right after VirtualAlloc succeeds
'         (records the new region address in the environment variable, used at the next reset recovery)
'
'   Return value: success = True, failure = False
' ============================================================
Public Function Thunks_Init() As Boolean

    If m_pRegionBase <> 0 Then
        Thunks_Init = True
        Exit Function
    End If

    ' --- Sentinel mechanism (added in stage 9.5) ---
    '   If reset was pressed in the previous run within the process and Shutdown did not run,
    '   the old region address remains in the environment variable. Reclaim it
    '   before proceeding to allocate the new region.
    Sentinel_RecoverIfNeeded

    ' Force-Compile EntryPoint (materialize the stub source)
    EntryPoint

    ' Source address of the EntryPoint stub
    Dim pStubSrc As LongPtr
    pStubSrc = VBA.Int(AddressOf EntryPoint)
    If pStubSrc = 0 Then Exit Function

    ' Get the addresses of Handler_QI / AddRef / Release only once at startup and save them
    m_pHandler_QI = GetAddr(AddressOf Handler_QueryInterface)
    m_pHandler_AddRef = GetAddr(AddressOf Handler_AddRef)
    m_pHandler_Release = GetAddr(AddressOf Handler_Release)
    If m_pHandler_QI = 0 Or m_pHandler_AddRef = 0 Or m_pHandler_Release = 0 Then
        Exit Function
    End If

    ' Allocate the region in one go
    m_pRegionBase = VirtualAlloc(0, REGION_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If m_pRegionBase = 0 Then Exit Function

    ' --- Sentinel mechanism (added in stage 9.5) ---
    '   Record the address just allocated in the environment variable.
    '   This way, even if reset is pressed before Init completes after this,
    '   it can be reliably reclaimed at the next Init.
    Sentinel_StorePrevRegion m_pRegionBase

    ' Zero-initialize the header (VirtualAlloc is already zero-cleared by nature, but just in case)
    Dim k As Long
    For k = 0 To HEADER_SIZE - 1 Step 8
        MemLongPtr(m_pRegionBase + k) = 0^
    Next k

    ' Set the alive flag (1 byte at the start of the region = 1)
    MemLongPtr(m_pRegionBase) = 1^

    ' Initialize all slots (stub copy + vtableObj initialization)
    Dim i As Long, pSlot As LongPtr, pVTableObj As LongPtr, pFunctions As LongPtr
    For i = 0 To SLOT_COUNT - 1
        pSlot = SlotAddrAt(i)

        ' --- Stub copy (96 bytes each; the padding at +91..+95 is harmless) ---
        For k = 0 To 95 Step 8
            MemLongPtr(pSlot + k) = ReadLongPtr(pStubSrc + k)
        Next k

        ' --- vtableObj initialization (40 bytes, written in 8-byte units) ---
        pVTableObj = pSlot + VTABLE_OBJ_OFFSET
        pFunctions = pVTableObj + PtrSize        ' +8, right after pVTable

        MemLongPtr(pVTableObj) = pFunctions
        MemLongPtr(pFunctions + 0 * PtrSize) = m_pHandler_QI
        MemLongPtr(pFunctions + 1 * PtrSize) = m_pHandler_AddRef
        MemLongPtr(pFunctions + 2 * PtrSize) = m_pHandler_Release
        MemLongPtr(pFunctions + 3 * PtrSize) = pSlot
    Next i

    ' Initialize the free list
    ReDim m_freeNext(0 To SLOT_COUNT - 1)
    For i = 0 To SLOT_COUNT - 2
        m_freeNext(i) = i + 1
    Next i
    m_freeNext(SLOT_COUNT - 1) = -1
    m_freeHead = 0
    m_inUse = 0

    ' Clear the whole Handler correspondence table
    For i = 0 To SLOT_COUNT - 1
        Set m_handlers(i) = Nothing
    Next i

    ' --- IID table initialization (added in stage 9.7a) ---
    '   Embeds IID_IUnknown and the IID of each HandlerKind into the table.
    InitIIDTable

    Thunks_Init = True
End Function


' ============================================================
' Thunks_AcquireSlot
'   Takes one free slot from the free list and writes into that slot's
'   thunk region the machine code that "injects pSelfObj into Me and calls pTargetFunc",
'   and writes the thunk's address at +55.
'
'   In stage 9.3a it is basically called from AcquireHandlerFor, but
'   the case of calling it directly like Test_TwoHandlers_Mock remains valid.
'
'   Return value: slot start address, 0 on failure
' ============================================================
Public Function Thunks_AcquireSlot( _
    ByVal handler As ComCallbackHandler, _
    ByVal pSelfObj As LongPtr, _
    ByVal pTargetFunc As LongPtr) As LongPtr

    If m_pRegionBase = 0 Then Exit Function
    If m_freeHead < 0 Then Exit Function     ' pool exhausted
    If handler Is Nothing Then Exit Function

    ' Take one from the free list
    Dim idx As Long
    idx = m_freeHead
    m_freeHead = m_freeNext(idx)
    m_freeNext(idx) = -1                      ' -1 = in use

    Dim pSlot As LongPtr
    pSlot = SlotAddrAt(idx)

    ' Write the thunk (74 bytes from offset +96)
    WriteThunkMachineCode pSlot + THUNK_OFFSET, pSelfObj, pTargetFunc, m_pRegionBase

    ' Write the thunk's address at +55 of the stub clone
    MemLongPtr(pSlot + LATE_BIND_OFFSET) = pSlot + THUNK_OFFSET

    ' Associate the Handler object by idx
    Set m_handlers(idx) = handler

    m_inUse = m_inUse + 1
    Thunks_AcquireSlot = pSlot
End Function


' ============================================================
' Thunks_DiscardHandler (English fork)
'   Gives back a handler whose COM registration failed synchronously
'   (add_*, ExecuteScript, CreateController, GetContent ... returned an error).
'   Before this, callers only did Set h = Nothing, but m_handlers(idx) still
'   held the handler and its owner (Pane/Environment), refCount stayed 0 so
'   Release never came: the slot was lost from the pool for good and the
'   owner could never terminate. Only acts when WebView2 holds no reference
'   (RefCount = 0); otherwise the normal Release path frees it later.
' ============================================================
Public Sub Thunks_DiscardHandler(ByVal h As ComCallbackHandler)
    If h Is Nothing Then Exit Sub
    If m_pRegionBase = 0 Then Exit Sub
    If h.RefCount <> 0 Then Exit Sub
    Dim idx As Long
    idx = SlotIndexFromAddr(h.Slot)
    If idx < 0 Then Exit Sub
    If Not (m_handlers(idx) Is h) Then Exit Sub
    h.ClearOwner
    Thunks_ReleaseSlot h.Slot
End Sub

' ============================================================
' Thunks_ReleaseSlot
'   Returns the specified slot to the free list.
'   Also clears m_handlers(idx).
' ============================================================
Public Sub Thunks_ReleaseSlot(ByVal pSlot As LongPtr)
    If m_pRegionBase = 0 Then Exit Sub
    If pSlot = 0 Then Exit Sub

    Dim idx As Long
    idx = SlotIndexFromAddr(pSlot)
    If idx < 0 Then Exit Sub
    If m_freeNext(idx) <> -1 Then Exit Sub           ' already free = double release

    ' Also clear the Handler correspondence table
    Set m_handlers(idx) = Nothing

    ' Return to the head of the free list
    m_freeNext(idx) = m_freeHead
    m_freeHead = idx

    m_inUse = m_inUse - 1
End Sub


' ============================================================
' SlotIndexFromAddr (internal helper)
'   Works backward from the thunk's address (= slot start address) to
'   the slot index.
' ============================================================
Private Function SlotIndexFromAddr(ByVal pSlot As LongPtr) As Long
    SlotIndexFromAddr = -1
    If m_pRegionBase = 0 Then Exit Function

    Dim offset As LongPtr
    offset = pSlot - (m_pRegionBase + HEADER_SIZE)
    If offset < 0 Then Exit Function
    If (offset Mod SLOT_SIZE) <> 0 Then Exit Function

    Dim idx As Long
    idx = CLng(offset \ SLOT_SIZE)
    If idx < 0 Or idx >= SLOT_COUNT Then Exit Function

    SlotIndexFromAddr = idx
End Function


' ============================================================
' SlotIndexFromVTableObjAddr (internal helper)
'   The address of vtableObj (= the value WebView2 passes as this in AddRef/Release,
'                     = pSlot + VTABLE_OBJ_OFFSET) is worked backward to the slot index.
' ============================================================
Private Function SlotIndexFromVTableObjAddr(ByVal pVTableObj As LongPtr) As Long
    SlotIndexFromVTableObjAddr = -1
    If m_pRegionBase = 0 Then Exit Function

    Dim offset As LongPtr
    offset = pVTableObj - (m_pRegionBase + HEADER_SIZE + VTABLE_OBJ_OFFSET)
    If offset < 0 Then Exit Function
    If (offset Mod SLOT_SIZE) <> 0 Then Exit Function

    Dim idx As Long
    idx = CLng(offset \ SLOT_SIZE)
    If idx < 0 Or idx >= SLOT_COUNT Then Exit Function

    SlotIndexFromVTableObjAddr = idx
End Function


' ============================================================
' Thunks_Shutdown
'   Sets the alive flag to 0 and then frees the region with VirtualFree.
'
'   Processing added in stage 9.5:
'     - At the end call Sentinel_ClearPrevRegion to erase the trace in the environment variable.
'       If forgotten, the next Init would try to VirtualFree the
'       "already freed address" again, causing a double free. VirtualFree just returns
'       0 harmlessly, but leaving no trace is the right thing.
' ============================================================
Public Sub Thunks_Shutdown()
    If m_pRegionBase = 0 Then Exit Sub

    ' First clear the alive flag
    MemLongPtr(m_pRegionBase) = 0^

    ' Cut the owner of every Handler and release the references (prevents circular references)
    Dim i As Long
    For i = 0 To SLOT_COUNT - 1
        If Not (m_handlers(i) Is Nothing) Then
            m_handlers(i).ClearOwner
            Set m_handlers(i) = Nothing
        End If
    Next i

    ' Free the region (diagnostic log added in stage 9.5: observe the return value and LastError)
    Dim shutFreeResult As Long
    Dim shutFreeErr As Long
    shutFreeResult = VirtualFree(m_pRegionBase, 0, MEM_RELEASE)
    shutFreeErr = GetLastError()
    Debug.Print "Thunks_Shutdown: VirtualFree(" & m_pRegionBase & ") returned " & _
                shutFreeResult & ", LastError = " & shutFreeErr & _
                IIf(shutFreeResult <> 0, " [OK]", " [NG] (" & VirtualFreeErrorName(shutFreeErr) & ")")

    m_pRegionBase = 0
    m_freeHead = -1
    m_inUse = 0
    Erase m_freeNext

    ' --- Sentinel mechanism (added in stage 9.5) ---
'   Leaves no trace on the normal termination path. If this is forgotten, the next Init
'   tries to VirtualFree an already freed address, causing useless processing
'   (harmless, but the "VirtualFree returned 0" on Debug.Print
'    becomes confusing).
    Sentinel_ClearPrevRegion
End Sub


' ============================================================
' Slot index -> slot start address
' ============================================================
Private Function SlotAddrAt(ByVal idx As Long) As LongPtr
    SlotAddrAt = m_pRegionBase + HEADER_SIZE + CLngLng(idx) * SLOT_SIZE
End Function


' ============================================================
' Writes the thunk machine code to the specified address (effective length 74 bytes)
'
'   Layout:
'     +0..+17  : alive flag check (18 bytes)
'     +18..+73 : existing thunk body (56 bytes, same as stages 6-8)
'
'   Implementation notes:
'     The byte array b is declared as 80 bytes (= a multiple of 8), and the last 6 bytes are
'     treated as zero padding. The destination slot's +96..+175 (= 80 bytes' worth)
'     is originally a padding/unused area, so writing 80 bytes is harmless.
' ============================================================
Private Sub WriteThunkMachineCode( _
    ByVal addr As LongPtr, _
    ByVal pSelfObj As LongPtr, _
    ByVal pTargetFunc As LongPtr, _
    ByVal pRegionBase As LongPtr)

    Dim b(0 To THUNK_BUF_SIZE - 1) As Byte, i As Long
    i = 0

    ' --- Alive flag check (18 bytes) ---
    ' mov rax, imm64   (= pAliveFlag = pRegionBase + 0)
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pRegionBase: i = i + 8
    ' cmp byte ptr [rax], 1
    b(i) = &H80: b(i + 1) = &H38: b(i + 2) = &H1: i = i + 3
    ' je +3 (= alive label, skips the 3 bytes of xor + ret)
    b(i) = &H74: b(i + 1) = &H3: i = i + 2
    ' xor eax, eax
    b(i) = &H33: b(i + 1) = &HC0: i = i + 2
    ' ret
    b(i) = &HC3: i = i + 1

    ' --- From here on, the same 56-byte thunk body as stages 6-8 ---

    ' sub rsp, 0x38
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HEC: b(i + 3) = &H38: i = i + 4
    ' mov r9, r8
    b(i) = &H4D: b(i + 1) = &H89: b(i + 2) = &HC1: i = i + 3
    ' mov r8, rdx
    b(i) = &H49: b(i + 1) = &H89: b(i + 2) = &HD0: i = i + 3
    ' mov rdx, rcx
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &HCA: i = i + 3
    ' mov rcx, imm64   (pSelfObj)
    b(i) = &H48: b(i + 1) = &HB9: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pSelfObj: i = i + 8
    ' lea rax, [rsp+0x28]
    b(i) = &H48: b(i + 1) = &H8D: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H28: i = i + 5
    ' mov [rsp+0x20], rax
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &H44: b(i + 3) = &H24: b(i + 4) = &H20: i = i + 5
    ' mov rax, imm64   (pTargetFunc)
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    MemLongPtr(VarPtr(b(i))) = pTargetFunc: i = i + 8
    ' call rax
    b(i) = &HFF: b(i + 1) = &HD0: i = i + 2
    ' mov eax, [rsp+0x28]
    b(i) = &H8B: b(i + 1) = &H44: b(i + 2) = &H24: b(i + 3) = &H28: i = i + 4
    ' add rsp, 0x38
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HC4: b(i + 3) = &H38: i = i + 4
    ' ret
    b(i) = &HC3: i = i + 1
    ' int3 / int3 (padding)
    b(i) = &HCC: i = i + 1
    b(i) = &HCC: i = i + 1

    ' --- Write 80 bytes in 8-byte units (the last 6 bytes are 0 padding, harmless) ---
    Dim k As Long
    For k = 0 To THUNK_BUF_SIZE - 1 Step 8
        MemLongPtr(addr + k) = ReadLongPtrFromBytes(b, k)
    Next k
End Sub


' ============================================================
' Extracts 8 bytes (LongPtr) from the specified offset of a byte array
' ============================================================
Private Function ReadLongPtrFromBytes(ByRef b() As Byte, ByVal offset As Long) As LongPtr
    ReadLongPtrFromBytes = ReadLongPtr(VarPtr(b(offset)))
End Function


' ============================================================
' Gets the function address of a fixed slot from the class's vTable
' ============================================================
Private Function GetClassMethodAddrAtFixedSlot( _
    ByVal cls As Object, ByVal slotIndex As Long) As LongPtr

    If cls Is Nothing Then Exit Function
    Dim pObj As LongPtr: pObj = ObjPtr(cls)
    If pObj = 0 Then Exit Function

    Dim pVTable As LongPtr
    pVTable = ReadLongPtr(pObj)
    If pVTable = 0 Then Exit Function

    GetClassMethodAddrAtFixedSlot = ReadLongPtr(pVTable + slotIndex * PtrSize)
End Function


' ============================================================
' IUnknown stub group (standard module)
'
'   Specification in stage 9.2:
'     - Work backward from this to idx with SlotIndexFromVTableObjAddr
'     - Operate m_handlers(idx) directly
'     - This allows multiple handlers to be managed in parallel
'
'   IID check introduced in stage 9.7a:
'     - Handler_QueryInterface compares riid with the genuine IID
'     - Only IID_IUnknown or "the proper IID corresponding to its own HandlerKind" gives S_OK
'     - Anything else gives ppvObject = 0 + E_NOINTERFACE
'     - On failure riid.Data1 is logged in Hex (no log on success, to reduce noise)
'
'   Points for obeying the COM convention:
'     - On success (when S_OK is returned) do AddRef
'     - On failure (when E_NOINTERFACE is returned) do not AddRef, and write NULL to ppvObject
'
'   For the rollback procedure in case of trouble, see the comment at the top of the module.
' ============================================================
Private Function GetAddr(ByVal addr As LongPtr) As LongPtr
    GetAddr = addr
End Function

Private Function Handler_QueryInterface( _
    ByVal this As LongPtr, _
    ByVal riid As LongPtr, _
    ByRef ppvObject As LongPtr) As Long

    ' --- Defense: if riid is NULL, E_POINTER ---
    If riid = 0 Then
        ppvObject = 0
        Handler_QueryInterface = &H80004003   ' E_POINTER
        Exit Function
    End If

    ' --- Defense: if this is invalid, E_NOINTERFACE ---
    Dim idx As Long
    idx = SlotIndexFromVTableObjAddr(this)
    If idx < 0 Then
        ppvObject = 0
        Handler_QueryInterface = E_NOINTERFACE
        Exit Function
    End If
    If m_handlers(idx) Is Nothing Then
        ppvObject = 0
        Handler_QueryInterface = E_NOINTERFACE
        Exit Function
    End If

    ' --- Compare riid with IID_IUnknown ---
    If IsEqualGUIDInPlace(riid, m_iidIUnknown) Then
        ppvObject = this
        HandlerAddRefInternal this
        Handler_QueryInterface = S_OK
        Exit Function
    End If

    ' --- Compare riid with "its own proper IID" ---
    Dim kind As HandlerKind
    kind = m_handlers(idx).kind
    If IsEqualGUIDInPlace(riid, m_iidTable(kind)) Then
        ppvObject = this
        HandlerAddRefInternal this
        Handler_QueryInterface = S_OK
        Exit Function
    End If

    ' --- Matches no IID: E_NOINTERFACE ---
    '   Log riid.Data1 (the first 4 bytes, a little-endian Long).
    '   You can observe which IID WebView2 tried at the moment QueryInterface was called
    '   (also useful for detecting well-behaved IIDs = IMarshal / IAgileObject etc.).
    Dim data1 As Long
    data1 = LongPtrLowDword(ReadLongPtr(riid))
    Debug.Print "  QI rejected [idx=" & idx & " kind=" & kind & _
                "] riid.Data1=&H" & Hex(data1) & " -> E_NOINTERFACE"
    ppvObject = 0
    Handler_QueryInterface = E_NOINTERFACE
End Function

Private Function Handler_AddRef(ByVal this As LongPtr) As Long
    Handler_AddRef = HandlerAddRefInternal(this)
End Function

Private Function Handler_Release(ByVal this As LongPtr) As Long
    Handler_Release = HandlerReleaseInternal(this)
End Function


' ============================================================
' IsEqualGUIDInPlace (newly added in stage 9.7a)
'
'   Compares two GUIDs (16 bytes).
'   One is riid (a pointer, passed from WebView2),
'   the other is m_iidTable / m_iidIUnknown (a GUID-type variable on the VBA side).
'
'   Method: a fast version that reads two LongPtrs (= 8 bytes on x64) and compares them.
'     - First 8 bytes: Data1 (4) + Data2 (2) + Data3 (2)
'     - Last 8 bytes: Data4(0..7)
'
'   The address of a GUID-type variable on the VBA side is obtained with VarPtr.
'   On x64 the GUID type is laid out contiguously in 16 bytes (no padding, and the data order
'   matches the Win32 standard little-endian). This is because the Win32 GUID structure and
'   VBA's Type declaration coincide by chance (or rather, naturally).
'
'   Verification points:
'     - Whether a VBA Type really is laid out contiguously in 16 bytes with no padding
'     - Whether Long / Integer are arranged little-endian in x64 memory
'   Both are standard behavior of x64 VBA, but if they do not match the IID check itself
'   fails (riid and m_iidTable do not match), so it can be noticed in verification.
' ============================================================
Private Function IsEqualGUIDInPlace(ByVal pRiid As LongPtr, ByRef refGuid As GUID) As Boolean
    Dim pRef As LongPtr
    pRef = VarPtr(refGuid)

    If ReadLongPtr(pRiid) <> ReadLongPtr(pRef) Then Exit Function
    If ReadLongPtr(pRiid + 8) <> ReadLongPtr(pRef + 8) Then Exit Function

    IsEqualGUIDInPlace = True
End Function


' ============================================================
' LongPtrLowDword (auxiliary helper, newly added in stage 9.7a)
'
'   Extracts the low 32 bits of a LongPtr as a Long. riid.Data1 (the first
'   4 bytes) is used in the log output of Handler_QueryInterface.
'   ComCallbackHandler.cls also has a Private function of the same name, but this one is for Module9 only.
'   The implementation is the same.
' ============================================================
Private Function LongPtrLowDword(ByVal v As LongPtr) As Long
    Dim u As LongLong
    u = CLngLng(v) And &HFFFFFFFF^
    If u > &H7FFFFFFF^ Then
        LongPtrLowDword = CLng(u - &H100000000^)
    Else
        LongPtrLowDword = CLng(u)
    End If
End Function


' ============================================================
' InitIIDTable (newly added in stage 9.7a)
'
'   Initializes m_iidIUnknown and m_iidTable(HK_xxx).
'   Called only once from the end of Thunks_Init.
'
'   Source of the IID values: WebView2.h (Microsoft Edge WebView2 SDK).
'   Each IID is already listed in WebView2_IID_vtable_reference.md.
'
'   How to write to each field of the GUID:
'     Data1 (Long, 4-byte little-endian): integer literal
'     Data2 (Integer, 2-byte little-endian): integer literal (because of the signed range,
'       values of &H8000 and above take the trouble of becoming Long with the & suffix and then
'       returning to Integer with CInt. This implementation adopts the String parsing method in FillGUID,
'       and by not writing integer values directly in code avoids the &H literal trap (design principle 29))
' ============================================================
Private Sub InitIIDTable()
    ' --- IID_IUnknown (Win32 standard, {00000000-0000-0000-C000-000000000046}) ---
    FillGUID m_iidIUnknown, "00000000-0000-0000-C000-000000000046"

    ' --- The proper IID for each HandlerKind ---

    ' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
    FillGUID m_iidTable(HK_EnvironmentCompleted), _
             "4e8a3389-c9d8-4bd2-b6b5-124fee6cc14d"

    ' ICoreWebView2CreateCoreWebView2ControllerCompletedHandler
    FillGUID m_iidTable(HK_ControllerCompleted), _
             "6c4819f3-c9b7-4260-8127-c9f5bde7f68c"

    ' ICoreWebView2NavigationStartingEventHandler
    FillGUID m_iidTable(HK_NavigationStarting), _
             "9adbe429-f36d-432b-9ddc-f8881fbd76e3"

    ' ICoreWebView2NavigationCompletedEventHandler
    FillGUID m_iidTable(HK_NavigationCompleted), _
             "d33a35bf-1c49-4f98-93ab-006e0533fe1c"

    ' ICoreWebView2WebMessageReceivedEventHandler
    FillGUID m_iidTable(HK_WebMessageReceived), _
             "57213f19-00e6-49fa-8e07-898ea01ecbd2"

    ' ICoreWebView2DocumentTitleChangedEventHandler
    FillGUID m_iidTable(HK_DocumentTitleChanged), _
             "f5f2b923-953e-4042-9f95-f3a118e1afd4"

    ' ICoreWebView2NewWindowRequestedEventHandler
    FillGUID m_iidTable(HK_NewWindowRequested), _
             "d4c185fe-c81c-4989-97af-2d3fa7ab5651"

    ' ICoreWebView2ExecuteScriptCompletedHandler (added in 9.8c)
    FillGUID m_iidTable(HK_ExecuteScriptCompleted), _
             "49511172-cc67-4bca-9923-137112f4c4cc"

    ' ICoreWebView2HistoryChangedEventHandler (added in 9.9b)
    '   Source: WebView2.h L3790
    FillGUID m_iidTable(HK_HistoryChanged), _
             "c79a420c-efd9-4058-9295-3e8b4bcab645"

    ' ICoreWebView2DOMContentLoadedEventHandler (added in 9.9b)
    '   Source: WebView2.h L5718
    FillGUID m_iidTable(HK_DOMContentLoaded), _
             "4bac7e9c-199e-49ed-87ed-249303acf019"

    ' ICoreWebView2WebResourceRequestedEventHandler (added in N-1)
    '   Source: the MIDL_INTERFACE line of WebView2.h (**taken from the header, not guessed**)
    FillGUID m_iidTable(HK_WebResourceRequested), _
             "ab00b74c-15f1-4646-80e8-e76341d25d71"

    ' ICoreWebView2WebResourceResponseReceivedEventHandler (added in N-3)
    '   Source: the MIDL_INTERFACE line of WebView2.h
    FillGUID m_iidTable(HK_WebResourceResponseReceived), _
             "7de9898a-24f5-40c3-a2de-d4f458e69828"

    ' ICoreWebView2WebResourceResponseViewGetContentCompletedHandler (added in N-4)
    '   Source: the MIDL_INTERFACE line of WebView2.h
    FillGUID m_iidTable(HK_ResponseGetContent), _
             "875738e1-9fa2-40e3-8b74-2e8972dd6fe7"
End Sub


' ============================================================
' FillGUID (newly added in stage 9.7a, promoted to Public in 9.9a)
'
'   Fills a GUID structure from a string of the form "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx".
'   Writing &H literals directly in code can fall into the trap (design principle 29), so
'   the method of parsing via a string is adopted.
'
'   Input example: "4e8a3389-c9d8-4bd2-b6b5-124fee6cc14d"
'   Split:
'     g(0..7)   = "4e8a3389"  -> Data1 (Long)
'     g(9..12)  = "c9d8"      -> Data2 (Integer)
'     g(14..17) = "4bd2"      -> Data3 (Integer)
'     g(19..22) = "b6b5"      -> Data4(0..1)
'     g(24..35) = "124fee6cc14d" -> Data4(2..7)
'
'   Each hexadecimal string is converted to a number by HexStrToLong / HexStrToInt.
'   Data2 / Data3 are Integer (signed 16-bit). A value like &HC9D8 overflows if
'   simply passed to CInt(), so it must first be read as a Long and then
'   have sign inversion processing applied. This implementation provides a dedicated helper
'   HexStrToInt for that.
'
'   Promoted to Public in 9.9a:
'     Wv2Pane.EnsureView2 calls this function to initialize a local GUID structure
'     for QueryInterface. What had been a utility dedicated to Wv2Thunks' InitIIDTable
'     was upgraded to a common utility. It is a basic facility used as a pair with the GUID type (Public Type),
'     so making it public is natural (= newly established as design principle 50,
'     to be finalized after 9.9a verification).
' ============================================================
Public Sub FillGUID(ByRef g As GUID, ByVal s As String)
    g.data1 = HexStrToLong(Mid$(s, 1, 8))
    g.Data2 = HexStrToInt(Mid$(s, 10, 4))
    g.Data3 = HexStrToInt(Mid$(s, 15, 4))

    ' Data4(0..7): the remaining 4 bytes (Mid 20..23) + the latter 12 digits (Mid 25..36)
    g.Data4(0) = CByte("&H" & Mid$(s, 20, 2))
    g.Data4(1) = CByte("&H" & Mid$(s, 22, 2))
    g.Data4(2) = CByte("&H" & Mid$(s, 25, 2))
    g.Data4(3) = CByte("&H" & Mid$(s, 27, 2))
    g.Data4(4) = CByte("&H" & Mid$(s, 29, 2))
    g.Data4(5) = CByte("&H" & Mid$(s, 31, 2))
    g.Data4(6) = CByte("&H" & Mid$(s, 33, 2))
    g.Data4(7) = CByte("&H" & Mid$(s, 35, 2))
End Sub


' ============================================================
' HexStrToLong (newly added in stage 9.7a)
'
'   Converts an 8-digit hexadecimal string to a Long. If the value is &H80000000 or above it becomes
'   negative (Long is signed 32-bit, two's complement).
'
'   VBA's CLng("&H...") overflows when the value exceeds the Long range.
'   So it is read via LongLong, and when the top bit is set it is
'   converted to two's complement form and then dropped to Long with CLng.
'
'   Input example: "4e8a3389" -> 1317797257
'           "d33a35bf" -> -752318529 (signed Long)
' ============================================================
Private Function HexStrToLong(ByVal s As String) As Long
    Dim v As LongLong
    v = CLngLng("&H" & s)                     ' receive 0..&HFFFFFFFF as a LongLong (signed 64-bit)
    If v >= &H80000000^ Then
        HexStrToLong = CLng(v - &H100000000^) ' 32-bit sign inversion (two's complement)
    Else
        HexStrToLong = CLng(v)
    End If
End Function


' ============================================================
' HexStrToInt (newly added in stage 9.7a)
'
'   Converts a 4-digit hexadecimal string to an Integer (signed 16-bit).
'   If the value is &H8000 or above it becomes negative.
'
'   VBA's CInt("&H...") can overflow at &H8000 and above
'   (= "an exception from trying to fit 32768 into the Integer range", a problem with the same root as the
'   trap of design principle 29).
'
'   As a safety measure, take it once as a Long with CLng and, if &H8000 or above, convert it manually to
'   two's complement form and then CInt.
' ============================================================
Private Function HexStrToInt(ByVal s As String) As Integer
    Dim v As Long
    v = CLng("&H" & s)                        ' obtain as a Long of 0..&HFFFF
    If v >= &H8000& Then
        HexStrToInt = CInt(v - &H10000)       ' convert to a negative Integer
    Else
        HexStrToInt = CInt(v)
    End If
End Function


' ============================================================
' HandlerAddRefInternal / HandlerReleaseInternal
'
'   Implementation in stage 9.2:
'     - Find idx from this (= pSlot + VTABLE_OBJ_OFFSET)
'     - Operate on the object in m_handlers(idx)
'     - An invalid this safely returns 0
'
'   Changes in stage 9.3a:
'     - Removed the trailing "gHandler Is h" branch from the automatic release path of
'       HandlerReleaseInternal (because gHandler itself was abolished)
'
'   Return value: refcount after the operation
' ============================================================
Private Function HandlerAddRefInternal(ByVal this As LongPtr) As Long
    Dim idx As Long
    idx = SlotIndexFromVTableObjAddr(this)
    If idx < 0 Then Exit Function
    If m_handlers(idx) Is Nothing Then Exit Function

    Dim n As Long
    n = m_handlers(idx).RefCount + 1
    m_handlers(idx).RefCount = n
    Debug.Print "  AddRef [idx=" & idx & "] -> refCount = " & n
    HandlerAddRefInternal = n
End Function

Private Function HandlerReleaseInternal(ByVal this As LongPtr) As Long
    Dim idx As Long
    idx = SlotIndexFromVTableObjAddr(this)
    If idx < 0 Then Exit Function
    If m_handlers(idx) Is Nothing Then Exit Function

    Dim n As Long
    n = m_handlers(idx).RefCount - 1
    If n < 0 Then n = 0                ' defense (cannot happen by design)
    m_handlers(idx).RefCount = n
    Debug.Print "  Release [idx=" & idx & "] -> refCount = " & n

    If n = 0 Then
        ' refcount reached 0, so cut the circular reference and release the slot.
        ' The order is important (the same since stage 9.1):
        '   1. Cut the circular reference with ClearOwner (cut ComCallbackHandler -> the upper class reference)
        '   2. Hold a local reference, then set m_handlers(idx) to Nothing
        '   3. Return the slot to the free list with Thunks_ReleaseSlot
        '   4. Set the local reference to Nothing too (ComCallbackHandler is completely released here)
        '
        ' Change in 9.3a: the old 4., the branch "if gHandler Is h, also cut gHandler",
        ' became unnecessary (because gHandler was abolished).

        Dim h As ComCallbackHandler
        Set h = m_handlers(idx)

        Dim pSlot As LongPtr
        pSlot = h.Slot

        ' 1. Cut the circular reference (ComCallbackHandler.m_owner = Nothing)
        '    This cuts the reference to the upper class (Wv2Environment etc.), and
        '    if the upper class side has also cut its reference to ComCallbackHandler,
        '    both get onto the path of being released by GC.
        h.ClearOwner

        ' 2. Set m_handlers(idx) to Nothing
        Set m_handlers(idx) = Nothing

        ' 3. Return the slot to the free list
        Thunks_ReleaseSlot pSlot

        ' 4. h loses its reference and is released when this procedure exits
        Set h = Nothing

        Debug.Print "  Release [idx=" & idx & "] -> refCount = 0 -> Slot " & idx & " auto-released"
    End If

    HandlerReleaseInternal = n
End Function


' ============================================================
' Memory operation primitives via PointerAccessor
' ============================================================
Private Property Let MemLongPtr(ByVal addr As LongPtr, ByVal newValue As LongPtr)
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = PtrSize
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        .arr(0) = newValue
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Property

Private Function ReadLongPtr(ByVal addr As LongPtr) As LongPtr
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.cbElements = PtrSize
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        ReadLongPtr = .arr(0)
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Function

Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal Ptr As LongPtr)
    ptrs(0) = Ptr
End Sub


' ============================================================
' EnsureWebView2Loader (English fork)
'   Makes sure WebView2Loader.dll (x64) is loaded before the first call to
'   CreateCoreWebView2EnvironmentWithOptions. Search order:
'     1. already loaded in this process
'     2. %WEBVIEW2_LOADER_PATH% (full path to the DLL, for your own copy)
'     3. the workbook's folder
'     4. Excel's folder (Application.Path)
'     5. Office's Power Query copy:
'        <Application.Path>\ADDINS\Microsoft Power Query for Excel Integrated\bin
'   Returns the path that was loaded ("(already loaded)"), or "" if none.
' ============================================================
Public Function EnsureWebView2Loader() As String
    Const DLL_NAME As String = "WebView2Loader.dll"
    If GetModuleHandleW(StrPtr(DLL_NAME)) <> 0 Then
        EnsureWebView2Loader = "(already loaded)"
        Exit Function
    End If

    Dim cands(1 To 4) As String, i As Long
    cands(1) = Environ$("WEBVIEW2_LOADER_PATH")
    On Error Resume Next
    cands(2) = ThisWorkbook.Path & "\" & DLL_NAME
    cands(3) = Application.Path & "\" & DLL_NAME
    cands(4) = Application.Path & "\ADDINS\Microsoft Power Query for Excel Integrated\bin\" & DLL_NAME
    On Error GoTo 0

    ' LoadLibraryW itself is the existence test: Dir$ raises on a malformed or
    ' unreachable path, and a 32-bit DLL simply fails to load and is skipped.
    For i = 1 To 4
        If Len(cands(i)) > Len(DLL_NAME) Then
            If LoadLibraryW(StrPtr(cands(i))) <> 0 Then
                EnsureWebView2Loader = cands(i)
                Exit Function
            End If
        End If
    Next i
End Function
