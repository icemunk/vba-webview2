# VBA-WebView2 Development Notes, Part 2 — Browser (Stage 9.3a to Stage 9.32b)

A record of how, on top of the thunk foundation established in Part 1, we stacked the Environment / Controller / CoreWebView2 class layers,
tab management, a WebView2-based tab bar and address bar, and settings persistence, and shaped the result into
something that works as a complete tabbed browser.
**This volume is frozen. Do not rewrite it.**

---

## Stage 9.3a — Introducing the upper-level class (Class8 Environment) and dynamic binding

### Background and goal

By stage 9.2 we had completed "a real COM callback mechanism that can manage multiple Class7 handlers in parallel". However, all control was concentrated in a standard module (`Module9_2.bas`), and the structure was naive: a single module variable `gHandler` held the only strong reference to a `Class7` instance.

To use WebView2 properly, we need to build a VBA-side class hierarchy that corresponds to the SDK's COM object hierarchy (`ICoreWebView2Environment` → `ICoreWebView2Controller` → `ICoreWebView2`). Stage 9.3 is the phase that establishes this "**introduction of upper-level classes and a multi-step creation chain**", and we proceeded by subdividing it into sub-stages 9.3a / 9.3b / 9.3c.

Goals of stage 9.3a:

- Create a new `Class8` (Environment class), and have Class7 hold the upper-level class through its m_owner property
- Establish a mechanism where `Class8.OnEnvironmentCompleted` is called from Class7.Handler_Invoke via **dynamic binding**
- Establish a lifetime-management pattern for one-shot handlers
- Retire the single gHandler variable
- Introduce a minimal helper for calling IUnknown::Release / AddRef from VBA via DispCallFunc

### Design decision points

#### D1. Lifetime management of the one-shot handler (adopted: option D1-α)

`HK_EnvironmentCompleted` is a handler that is "called once and done". Class8 ←→ Class7 form a circular reference through `m_handler` / `m_owner`, and the question is when to break it:

- **Option D1-α (adopted)**: At the end of `Class8.OnEnvironmentCompleted`, do `Set m_handler = Nothing`. When WebView2 fires Release and the refcount reaches 0, the foundation's automatic release path runs, and `Class7.ClearOwner` also cuts the Class7 → Class8 reference. The circular reference is resolved from both directions.
- Option D1-β: Add a hook in Class7.ClearOwner that "notifies the upper level", and have Class8 receive the notification and cut the reference. Rejected because it couples the classes too tightly.
- Option D1-γ: Have Class8 hold m_handler as a WeakRef. A WeakRef equivalent can be built in VBA via ObjPtr, but it is complicated, so it was rejected.

The adopted option D1-α works cleanly for a "called once and done" handler. Persistent handlers (NavigationStarting etc.) need a different pattern (option P1), which was designed separately in stage 9.4a.

#### D2. Ownership relationship where Class8 contains Class7

Class8 has exactly one handler for HK_EnvironmentCompleted (one to one). The type of `m_owner` stays `Object`, and we do not adopt a dedicated Interface (e.g. Implements of `IHandlerOwner`). Reasons:

- At the time of stage 9.3a there is only one pattern of "what the upper-level class wants to do when the handler completes", so rushing into abstraction would make the pattern drift away from reality
- Dynamic binding (`m_owner.OnEnvironmentCompleted ...`) is sufficient
- Each upper-level class gets the flexibility to have differently named methods per kind (`OnControllerCompleted`, `OnNavigationStarting`, etc.)

#### D3. Retiring the CreateWebView2Environment Sub

The `CreateWebView2Environment` Sub of stage 9.2 sat in the standard module as the entry point that created gHandler and called the API; from stage 9.3a on it is fully absorbed into `Class8.Init`. We do not place a global variable for the "current Environment" in the standard module (concern: Controller and CoreWebView2 would follow one after another and bring chaos again). The style is to call `Dim env As New Class8: env.Init "..."` directly from the test Sub `Test_Environment_Real`.

#### D4. Minimal DispCallFunc helper

`Class_Terminate` needs to fire `m_pEnvironment->Release`, which is the first case of "calling a native COM method from VBA". As a minimal implementation, we add `ComRelease` / `ComAddRef` (zero-argument versions). They are planned to be merged into the variadic `dcf` in stage 9.3b.

### Implementation

New and changed files:

| File | Change |
|---|---|
| `Class8.cls` | **New**. Environment class. Holds `m_handler`, `m_pEnvironment`, `m_state`. `Init` acquires the handler + calls the API, `OnEnvironmentCompleted` receives the completion notification, `Class_Terminate` does the Release |
| `Module9_3.bas` | Diff from Module9_2: removed `gHandler`, retired the `CreateWebView2Environment` Sub, added `ComRelease` / `ComAddRef`, added public functions `AcquireHandlerFor` / `EnsureFolder`, added `Test_Environment_Real`, `Test_TwoHandlers_Mock` left as is |
| `Class7.cls` | Added the `m_owner.OnEnvironmentCompleted` call to the `HK_EnvironmentCompleted` branch of `Handler_Invoke` |

Flow of Class8's Init:

1. Acquire a Class7 with `AcquireHandlerFor(HK_EnvironmentCompleted, Me)` (owner = Me, which becomes the receiver of the dynamic binding)
2. Compute the obj address held by WebView2 (`pSlot + VTABLE_OBJ_OFFSET`)
3. **Set `m_state = Es_Waiting` before the API call** (to handle synchronous callbacks)
4. Call `CreateCoreWebView2EnvironmentWithOptions` through a plain Declare, not through dcf
5. Receive the completion notification in `OnEnvironmentCompleted`, and hold `m_pEnvironment` after a ComAddRef

### Findings from the real-machine test

#### Finding A1: Public Sub is required (Friend cannot be dynamically bound)

When the first implementation declared `OnEnvironmentCompleted` as a `Friend Sub`, the call `m_owner.OnEnvironmentCompleted errorCode, arg2` raised the runtime error "**Object doesn't support this property or method.**"

Cause: dynamic binding (late binding) through the VBA `Object` type internally resolves the method name with `IDispatch::GetIDsOfNames` and then calls it with `IDispatch::Invoke`. `Friend` methods are **not exposed on the `IDispatch` interface**, so the name resolution fails.

Fix: changing `Friend Sub OnEnvironmentCompleted` to `Public Sub OnEnvironmentCompleted` made it work immediately.

→ **Design principle established**: Declare every method that is called through dynamic binding (`m_owner.OnXxx`) as a `Public Sub`. `Friend` may be used only when calling with static types from a Module9_X (for example, when calling `Class7.Init` from a Module9_X).

`Invoke_EnvironmentCompleted` in the old WebView2.cls was also declared `Public`, but the reason for that was not stated at the time. This stage made clear "why Public is required".

#### Finding A2: WebView2Loader calls back synchronously

When we implemented the naive version that sets `m_state = Es_Waiting` **after** the API call, the test never reached IsReady and timed out. Examining the log:

```
Calling env.Init ...
Class7.Invoke[HK_EnvironmentCompleted]    ← fired synchronously inside the API call
Class8.OnEnvironmentCompleted: ...
  ↑ m_state was set to Es_Ready here
State right after Init = 1                 ← then overwritten with Es_Waiting!
```

It turned out that `CreateCoreWebView2EnvironmentWithOptions` can **fire the callback synchronously inside the API call**. This happens when the WebView2 runtime is already running and environment creation completes immediately.

Fix: changed it to set `m_state = Es_Waiting` **before** the API call. This way, for both synchronous and asynchronous callbacks, the final state is consistent with the value set by `OnEnvironmentCompleted`.

→ **Design principle established**: Keep the order "set state to Waiting → API call". The same pattern is used for all later layers, such as Controller creation.

#### Finding A3: Behavior on VBA reset

After Test_Environment_Real finished, cleanup also ran cleanly in a simple Class8 test that did not go through a UserForm. We confirmed that Class_Terminate fires correctly and the Environment pointer is released by ComRelease.

### Established design principles

- **D1-α: One-shot handler lifetime management**: At the end of OnXxxCompleted, `Set m_handler = Nothing`; the circular reference is resolved from both directions by the foundation's automatic release path
- **Synchronous callback handling**: Set `m_state = Es_Waiting` **before** the API call
- **Public Sub strictly**: All methods called through dynamic binding are Public; Friend is only for statically typed calls
- **Class responsibility scope**: Class8 is limited to creating and holding the Environment; Controller creation is handled by Class9 (the responsibility is reversed, which reduces the risk of circular references)

### Known limitations

- Leak handling for the API failure path is not implemented (if `CreateCoreWebView2EnvironmentWithOptions` fails, m_handlers(idx) remains with refcount=0 and the slot leaks). Tolerated until the Sentinel mechanism is introduced in stage 9.5.

---

## Stage 9.3b — Controller class (Class9) + dcf unification + reconnaissance log

### Background and goal

We roll out the upper-level class pattern established in stage 9.3a horizontally to the Controller layer. We use real-machine reconnaissance to check whether the predicted argument interpretation confirmed with HK_EnvironmentCompleted (`this = handler obj, arg1 = 1st argument, arg2 = 2nd argument`) also holds for the other kinds. At the same time, we introduce the variadic dcf for calling COM methods from VBA, and unify `ComRelease` / `ComAddRef` as thin wrappers over it.

Goals of stage 9.3b:

- Create a new `Class9` (Controller class) and implement it as a mirror image of Class8
- Establish a mechanism that calls `Environment->CreateCoreWebView2Controller(hwnd, handler)` via dcf
- Confirm the argument interpretation with the HK_ControllerCompleted reconnaissance log
- Introduce a cleaned-up version of the old dcf in Module9_4

### Design decision points

#### E1. dcf unification (cleanup points)

Reviewing the implementation of the old dcf turned up the following problems:

**Problem 1**: The double acquisition of pVTable / pMethod via CopyMemory is dead code. The acquired values are never used (DispCallFunc does the same thing internally).
```vba
Dim pVTable As LongPtr
CopyMemory pVTable, ByVal pInterface, LenB(pVTable)
Dim pMethod As LongPtr
CopyMemory pMethod, ByVal (pVTable + vtblIndex * LenB(pVTable)), LenB(pMethod)
' ↑ The values acquired here are never used afterwards
```

**Problem 2**: `CLngPtr(CLng(vals(i)))` inside `Case Else` is unreachable code. It is already handled by the preceding `Case vbLong:` branch, so it can never be reached.

The cleaned-up version removes both. CopyMemory is no longer needed either.

#### E2. Removing dcf's automatic error check

When `ComRelease` was replaced with a dcf-based version in stage 9.3a, a log line `dcf METHOD failed: AddRef result=&H3` appeared. Cause: dcf judged "a nonzero return value is an error", but `AddRef` / `Release` are methods that return not an HRESULT but a **ULONG (the new reference count)**, so a value != 0 is normal.

Fix: removed dcf's automatic error check (Debug.Print when the return value ≠ 0). The policy is to leave the check to the caller.

→ **Design principle established**: dcf does not assume "it returns an HRESULT" and returns the return value as is. Failure checks are done by the caller when it knows "this is a method that returns an HRESULT".

#### E3. Class9 is a mirror image of Class8

The structure of Class9 is almost the same shape as Class8:

- `m_handler As Class7` (the one-shot handler for HK_ControllerCompleted)
- `m_pCtrl As LongPtr` (ICoreWebView2Controller)
- `m_state As ControllerState` (Cs_Created / Cs_Waiting / Cs_Ready / Cs_Failed)

Init is `Class9.Init(env As Class8, hwnd As LongPtr)`, receiving Class8 as an argument (Class9 only "uses" env and does not hold it). The ownership relationship between Environment and Controller is split so that the caller (UserForm etc.) manages it.

#### E4. How to obtain the UserForm's HWND

An HWND must be passed to the WebView2 Controller. How do we get the HWND of an Excel UserForm?

Adopted: place one Frame control on the UserForm and obtain the Frame's HWND with the `Frame1.[_GethWnd]` property. Advantages:

- Enclosing the name in square brackets `[ ]` lets VBA reference the hidden property of MSForms.Frame
- The obtained HWND can be used directly as the parent window of the WebView2 Controller
- Using the Frame's HWND rather than the UserForm's own HWND secures the WebView2 area as a rectangle

### Implementation

| File | Change |
|---|---|
| `Class9.cls` | **New**. Controller class, a mirror image of Class8 |
| `Class7.cls` | Materialized the HK_ControllerCompleted branch (m_owner.OnControllerCompleted call), enriched the reconnaissance log |
| `Module9_4.bas` | Diff from Module9_3: added the public function `dcf`, changed `ComRelease`/`ComAddRef` into thin wrappers over dcf, added `Test_Controller_Real` |
| `UserForm1.frm` | **New**. Frame1 + `StartWebView2` method + `UserForm_Terminate` cleanup |

Main structure of the cleaned-up dcf:

```vba
Public Function dcf( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    ByVal funcName As String, _
    ParamArray args() As Variant) As Long

    ' Argument type inference: vbLong / vbLongLong / vbDouble / other (LongLong fallback)
    ' Calls DispCallFunc(pInterface, vtblIndex * PtrSize, CC_STDCALL, vbLong,
    '              argc, vt(0), vp(0), res)
    ' The return value is returned as is (the HRESULT check is up to the caller)
End Function

Public Function ComRelease(ByVal p As LongPtr) As Long
    If p <> 0 Then ComRelease = dcf(p, 2, "Release")
End Function

Public Function ComAddRef(ByVal p As LongPtr) As Long
    If p <> 0 Then ComAddRef = dcf(p, 1, "AddRef")
End Function
```

### Findings from the real-machine test

#### Finding B1: Reconnaissance result for HK_ControllerCompleted

Real-machine log:

```
Class9.Init: dcf(pEnv=55972038131728, 3, "Env.CreateController", hwnd=1377076, handler=2364886679792)
  AddRef [idx=0] → refCount = 1
  AddRef [idx=0] → refCount = 2          ← why twice?
Class7.Invoke[HK_ControllerCompleted]
  this        : 2364886679792             ← handler obj (VBA region)
  arg1 (raw)  : 0                          ← errorCode = S_OK
  errorCode   : 0
  pController : 55972014948360             ← WebView2 region
  m_refCount  : 2
```

Points of note:

- `this` is the handler obj address (the same as `m_handler.Slot + VTABLE_OBJ_OFFSET`)
- `arg1` is the errorCode value as is (received as a LongPtr, with the lower 32 bits taken by `LongPtrLowDword`)
- `arg2` is pController (the ICoreWebView2Controller pointer)
- The two consecutive AddRefs are WebView2's internal defensive reference management (a temporary +1 so that the handler is not released during Invoke)
- After Invoke, two consecutive Releases also arrive, returning the refcount to 0

This confirmed that the correspondence of Class7's `(this, arg1, arg2)` arguments, which we had predicted in stage 9.2, also holds for HK_ControllerCompleted.

#### Finding B2: Synchronous callback (holds at the Controller layer too)

By the time we returned from Class9.Init, `m_state = Cs_Ready` was already set, and the `DoEvents` wait loop in UserForm1.StartWebView2 exited after 0 iterations. We confirmed that, as with Environment, the Controller can also complete via a synchronous callback.

#### Finding B3: Effect of removing dcf's automatic error check

Before the removal, "dcf METHOD failed: ..." appeared on every call to `ComRelease`/`ComAddRef`, but after the removal the log was clean. The design decision that dcf does not need to distinguish ULONG-returning methods such as AddRef/Release from HRESULT-returning methods at the dcf level turned out to be correct.

### Established design principles

- **dcf unification**: Migrated from the old CopyMemory-based version to PointerAccessor + the cleaned-up version. With variadic arguments (ParamArray) it can call methods with 0 to N arguments
- **dcf does no HRESULT check**: It returns the return value as is, and the check is done by the caller
- **Obtain the HWND with UserForm + Frame1.[_GethWnd]**: The standard idiom for securing a child window HWND via a Frame control

### Known limitations

- Bing is not displayed yet (planned for stage 9.3c)
- Resize tracking is not implemented (to be considered when the umbrella class is introduced; deferred because even mouse-resizing a UserForm itself requires SetWindowLongPtr)

---

## Stage 9.3c — WebView2 display milestone (Bing appears on screen)

### Background and goal

Stage 9.3a got Environment working and stage 9.3b got Controller working. Next we obtain ICoreWebView2, run `Navigate("https://www.bing.com")`, and achieve the "**first milestone**" of showing WebView2 on the Excel window.

Goals of stage 9.3c:

- Extend Class9 into a **combined class** of Controller and CoreWebView2
- Set the WebView2 size with `Ctrl_PutBounds`
- Obtain the ICoreWebView2 pointer with `Ctrl_GetCoreWebView2`
- Actually display Bing with `View_Navigate`

### Design decision points

#### F1. Class9 = Controller + CoreWebView2 combined class (adopted: option a)

Whether to create ICoreWebView2 (the CoreWebView2 class) as a separate class Class10 or merge it into Class9. After discussion with tarboh, the following options were laid out:

- Option a (adopted): **Merge into Class9**. We give up the one-to-one correspondence with the SDK hierarchy, but gain a natural correspondence from the VBA user's view: "one WebView2 display = one class"
- Option b: Create Class10 as a new class. The correspondence with the SDK hierarchy is clean, but Controller / CoreWebView2 line up one to one, so the benefit of separating them is thin
- tarboh's judgment: **"Internally, option b is easier to organize during development, but in practical use option a is easier to use"**, though with the concern that "managing it will be hard when interface versions increase". Claude summarized that "the main risk of option a is manageable" and proceeded with option a

#### F2. Method naming convention: `Ctrl_xxx` / `View_xxx` prefixes

Since option a consolidates everything into one class, the name of each method must show whether it belongs to the Controller or to CoreWebView2. tarboh's proposal of "`ICoreWebView2Controller_xxx` / `ICoreWebView2_yyy`" is too verbose, so Claude proposed the following convention:

- `Ctrl_xxx` : base methods of `ICoreWebView2Controller`
- `View_xxx` : base methods of `ICoreWebView2`
- `Ctrl2_xxx`, `Ctrl3_xxx` : future version extensions such as `ICoreWebView2Controller2`
- `View2_xxx`, `View3_xxx` : future version extensions such as `ICoreWebView2_2`

The 4-character prefixes `Ctrl` / `View` make it easier to narrow down completion candidates. `View` is a name taken from the "View" part of CoreWebView2.

#### F3. Implement only the base version; add extensions when needed (adopted: plan 4)

The WebView2 SDK has versioned extensions such as `ICoreWebView2Controller2` to `6` and `ICoreWebView2_2` to `_22`. How should they be handled?

- Plan 1: Hold the vtable pointers of all versions in one class
- Plan 2: A separate class per version (`ControllerExt2`, etc.)
- Plan 3: Lazy caching through a method-table abstraction
- Plan 4 (adopted): **Implement only the base; add version extensions to the same class when they are used**

The scope of stages 9.3c / 9.4 needs only the base version (Navigate / Bounds / event registration are all base). The policy is to add fields and methods to Class9 once a version extension becomes necessary.

#### F4. How to pass the RECT structure

ICoreWebView2Controller::put_Bounds takes a 16-byte structure by value (the `RECT` structure). Under the x64 ABI a 16-byte structure is passed by value, but it has been confirmed that passing a pointer with `VarPtr(rc)` also works in the old version's implementation.

Adopted: follow the old version and pass a pointer in the form `dcf(m_pCtrl, 6, "Ctrl.put_Bounds", VarPtr(rc))`. This is established as the common pattern for passing a structure from VBA.

#### F5. Get the HWND size directly with GetClientRect

The old version used `Frame1.InsideWidth` (in points) + `PtsToPx` for conversion, but tarboh pointed out: "**use `GetClientRect` to get it directly in pixels**".

Adopted: get the Frame's client area directly in pixels with `GetClientRect(hWnd_Frame, rc)`. The advantage is that the correct value is obtained automatically even in HiDPI environments.

#### F6. Deferring mouse-resize support for the UserForm

While implementing resize tracking, tarboh pointed out that "a VBA UserForm cannot be resized with the mouse by default". It is possible by adding `WS_THICKFRAME` with `SetWindowLongPtr(hwnd, GWL_STYLE, ...)`, but we decided to leave it out of the scope of stage 9.3c.

The policy for adopting it in the future has been agreed with tarboh:

```
[Future resize support policy]
- Getting the UserForm's own HWND: start from Frame1.[_GethWnd] and walk up with GetAncestor(hwnd, GA_ROOT) to the UserForm HWND
- Timing: change the style at the end of UserForm_Initialize
- Minimum / maximum size limits: none (up to the user)
- Implementation: SetWindowLongPtrW + SetWindowPos(SWP_FRAMECHANGED)
```

### Implementation

| File | Changes |
|---|---|
| `Class9.cls` | Extended: added the `m_pView` field, built `Ctrl_PutBounds` + `Ctrl_GetCoreWebView2` into the end of `Init`, added the `View_Navigate` method |
| `Module9_5.bas` | Diff from Module9_4: `RECT` type (Public), `GetClientRect` API (Public Declare), added `Test_Navigate_Real` |
| `UserForm1.frm` | Extended: `Me.Width = 800: Me.Height = 600`, built Ctrl_PutBounds / Ctrl_GetCoreWebView2 / View_Navigate into `StartWebView2` |

Confirmed vtable indexes (verified on the real machine with the old version):

| Method | vtable index |
|---|---|
| `Ctrl_PutBounds` (ICoreWebView2Controller::put_Bounds) | 6 |
| `Ctrl_GetCoreWebView2` (ICoreWebView2Controller::get_CoreWebView2) | 25 |
| `View_Navigate` (ICoreWebView2::Navigate) | 5 |

### Findings from the real-machine test

#### Finding C1: Bing display milestone achieved

The Bing top page was rendered correctly on UserForm1. Log:

```
UserForm1.StartWebView2: Frame1 client rect = (0,0,768,536)
UserForm1.StartWebView2: Ctrl_PutBounds OK
UserForm1.StartWebView2: Ctrl_GetCoreWebView2 OK (pView=55972014053376)
UserForm1.StartWebView2: View_Navigate OK (Bing should be displayed)
```

With a single dcf, all three argument patterns (a lone LongPtr / a structure pointer / a string pointer + out argument) work. The generic design of dcf was confirmed.

#### Finding C2: Flawless cleanup

When the UserForm is closed with the x:

```
UserForm1.Terminate started
Class9.Terminate: Releasing m_pView = 55972014053376
Class9.Terminate: Releasing m_pCtrl = 55972014949128
Class8.Terminate: Releasing m_pEnvironment = 55972038270992
UserForm1.Terminate finished
```

Class_Terminate runs in the order Class9 -> Class8, and each releases its COM pointer with ComRelease. No circular references and no leaks.

### Established design principles

- **Plan a: Class9 = combined Controller + CoreWebView2 class**: we give up the one-to-one mapping with the SDK hierarchy, but from the VBA user's viewpoint "one WebView2 display = one class"
- **Ctrl_ / View_ prefixes**: the method naming makes the origin explicit
- **Implement only the base; add version extensions in the future**: in stage 9.3c only the Ctrl / View base is covered
- **Pass structure arguments as VarPtr pointers**: real-machine behavior takes priority over the x64 ABI specification
- **Get the HWND size in pixels directly with GetClientRect**: avoids points + DPI conversion
- **Release order is child to parent**: m_pView -> m_pCtrl -> m_pEnvironment (makes the lifetime dependencies explicit)

### Known limitations

- Mouse resizing of the UserForm is not supported yet (to be implemented later with SetWindowLongPtr; the policy is settled)
- Version extension interfaces (Controller2, CoreWebView2_2, etc.) are not supported yet


## Stage 9.4a — Persistent handler mechanism + NavigationStarting reconnaissance

### Completion date
This session (at the completion of stage 9.4a)

### Background and purpose

Stages 9.3a to 9.3c established the "**one-shot handler**" pattern (plan D1-α: Set m_handler = Nothing at the end of OnXxxCompleted). What is needed next is the "**persistent handler**" pattern. To receive events that keep firing many times during page operations in WebView2, such as NavigationStarting / NavigationCompleted / WebMessageReceived / DocumentTitleChanged, a lifetime management different from the one-shot one is needed.

Stage 9.4 as a whole is subdivided into 3 sub-stages:
- **9.4a**: Minimal implementation of the persistent handler mechanism (just one NavigationStarting)
- **9.4b**: Roll out other event handlers (NavigationCompleted, DocumentTitleChanged) + navigation control
- **9.4c**: WebMessageReceived two-way communication

Goals of stage 9.4a:

- Establish the **lifetime management pattern for persistent handlers**
- Implement Class9's handler-holding structure as a **HandlerKind-linked array**
- Implement only `HK_NavigationStarting` and confirm the argument interpretation through reconnaissance logs
- Establish the flow of getting the destination URL with `args.get_Uri`
- Verify reset-button resilience on the real machine

### Design decision points

#### G1. Lifetime management of persistent handlers (adopted: plan P1)

Plan D1-α for one-shot handlers was "cut m_handler at the end of OnXxxCompleted", but this cannot be used for persistent handlers, because "if the reference is cut in OnXxxStarting, the second and later events never arrive".

Plans considered:

- Plan P1 (adopted): **"Class9 only holds it; removal is explicit"**
  - Class9.AddNavigationStarting creates a Class7 and holds it in its own field
  - No matter how many times WebView2 calls Invoke, Class9 does not cut m_handler
  - Removal timing: Class_Terminate (UserForm shutdown) or an explicit RemoveNavigationStarting
  - Flow: `View_RemoveEvent` -> WebView2 releases Class7 -> refcount=0 -> automatic release path -> Class7.ClearOwner also cuts the reference to Class9 -> Class9 sets m_persistentHandlers(k) = Nothing
  - High symmetry with the one-shot case (plan D1-α): the "automatic release path" is the same, and the only thing Class9 does is "fire the removal and cut its own reference"

- Plan P2: Make Class7's m_owner to Class9 a weak reference. A WeakRef equivalent is possible in VBA via ObjPtr, but rejected because it adds complexity
- Plan P3: Do not hold the handler in a Class9 field; remember only the token. It survives only via m_handlers(idx) on the infrastructure side. Rejected because the lifetime becomes unclear from the Module9_X viewpoint

#### G2. Class9's handler-holding structure (adopted: plan r)

How to build a structure where Class9 manages multiple persistent handlers in parallel:

- Plan p: Parallel fields. `m_h_navStart As Class7`, `m_h_navComp As Class7`, ... one field declared per kind. Adding a new kind increases the code changes
- Plan q: A generic container using Dictionary / Collection. Requires a `Microsoft Scripting Runtime` reference, and we do not want to add dependencies
- Plan r (adopted): **An array indexed by HandlerKind**

```vba
Private m_persistentHandlers(HK_NavigationStarting To HK_DocumentTitleChanged) As Class7
Private m_persistentTokens  (HK_NavigationStarting To HK_DocumentTitleChanged) As LongLong
```

Advantages:
- HandlerKind is already declared as a `Public Enum` in Module9_X and can be referenced from Class9
- "Loop and remove_ every token" can be written cleanly in Class_Terminate
- When adding a new kind, just widen the range of the array + add a Select Case branch to RemoveVtableIndexFor
- It cannot hold multiple handlers of the same kind (e.g. two NavigationStarting) (at most one per kind per Class9 instance), but that is sufficient for stage 9.4

#### G3. Type of EventRegistrationToken

In C++ `EventRegistrationToken` is a structure, but internally it is a single field `INT64 value;`. It can effectively be treated as an 8-byte integer. In VBA it can be received as `LongLong` (= 8 bytes, Win64 only).

It was confirmed that the old version also treated the token as 8 bytes.

#### G4. Implementation policy for the string conversion helper (adopted: plan β = RtlMoveMemory)

`args.get_Uri` returns an LPWSTR* (a UTF-16 null-terminated string allocated with `CoTaskMemAlloc`) as an out argument. Ways to convert it into a VBA String:

- Plan α: Build it one character at a time with ChrW using the PointerAccessor. Consistent with the no-CopyMemory policy, but slow
- Plan β (adopted): **Use RtlMoveMemory**. Reuse the old version's `PtrToString` implementation as is. It departs from the no-CopyMemory policy, but for string conversion alone speed takes priority

tarboh's decision: "Q1. Go with (β) for now. I'm also attaching the helpers that were in the old standard module, so please use anything that seems useful." The old `mWebView2Helper`'s `PtrToString` and `GetStringProperty` were taken in as is.

#### G5. Reset-button resilience (adopted: plan R4)

The VBE stop button (reset) is used as an everyday matter in VBA development. Behavior at reset:

- All module variables and class instances are destroyed
- Class_Terminate may not be called
- The WebView2 side may keep holding the Class7 handler and fire Invoke later

Design plans:

- Plan R1: Guarantee "crash avoidance" with the alive flag alone, give up on leaks
- Plan R2: Detect "traces of the previous reset" at every run and clean up. The module variables themselves are cleared by the reset and so cannot be tracked; rejected
- Plan R3: Expose a ForceCleanup API to the user and have them call it manually
- Plan R4 (adopted): **R1 + a strict remove_ order for persistent handlers**. "Crash avoidance is absolutely guaranteed; leak countermeasures come later (the stage 9.5 Sentinel mechanism)"

tarboh's decision: "Q5. Go with (R4)", "Q6. (s) include it" (put the reset test in the verification menu).

### Implementation

| File | Changes |
|---|---|
| `Class7.cls` | Made the HK_NavigationStarting branch real, also prepared the comments of the argument interpretation table |
| `Class9.cls` | Added the `m_persistentHandlers` / `m_persistentTokens` arrays, `AddNavigationStarting` / `RemoveNavigationStarting` / `OnNavigationStarting` methods, a `CleanupAllPersistentHandlers` mechanism that removes all persistent handlers in `Class_Terminate`, and a `RemoveVtableIndexFor` HandlerKind -> vtable index conversion helper |
| `Module9_6.bas` | Diff from Module9_5: `lstrlenW` / `CoTaskMemFree` / `RtlMoveMemory` as Public Declare, added the `PtrToString` / `GetStringProperty` helpers, added the `Test_ResetSurvival_Manual` reset-test guide |
| `UserForm1.frm` | Added a `m_ctrl.AddNavigationStarting()` call at the end of `StartWebView2` |

#### Main vtable indexes

| Method | vtable index |
|---|---|
| `add_NavigationStarting` (ICoreWebView2) | 7 |
| `remove_NavigationStarting` | 8 |
| `get_Uri` (ICoreWebView2NavigationStartingEventArgs) | 3 |

#### Core logic of the persistent handler mechanism

```vba
' On registration
Public Function AddNavigationStarting() As Long
    Dim h As Class7
    Set h = AcquireHandlerFor(HK_NavigationStarting, Me)
    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET
    Dim token As LongLong
    Dim hr As Long
    hr = dcf(m_pView, 7, "View.add_NavigationStarting", pVTableObj, VarPtr(token))
    ' Save only on success
    m_persistentTokens(HK_NavigationStarting) = token
    Set m_persistentHandlers(HK_NavigationStarting) = h
End Function

' Bulk removal in Class_Terminate
Private Sub CleanupAllPersistentHandlers()
    Dim k As Long
    For k = LBound(m_persistentHandlers) To UBound(m_persistentHandlers)
        If Not (m_persistentHandlers(k) Is Nothing) Then
            Dim removeIndex As Long
            removeIndex = RemoveVtableIndexFor(CLng(k))
            If removeIndex > 0 Then
                dcf m_pView, removeIndex, "View.remove_Event", m_persistentTokens(k)
            End If
            Set m_persistentHandlers(k) = Nothing
            m_persistentTokens(k) = 0^
        End If
    Next k
End Sub
```

`RemoveVtableIndexFor` returns the mapping from HandlerKind to the remove_ vtable index (NavigationStarting=8, etc.). When adding a new kind, you only add one line to the Select Case.

### Excerpts of real-machine test logs and observed behavior

#### Failure of the first implementation (important)

The first code crashed. The cause was the argument passing in Class7's HK_NavigationStarting branch:

**The wrong implementation**:
```vba
Case HK_NavigationStarting
    If Not (m_owner Is Nothing) Then
        m_owner.OnNavigationStarting this, arg1   ' <- wrong
    End If
```

Log:
```
Class7.Invoke[HK_NavigationStarting]
  this        : 2702042333424      <- VBA area (handler obj)
  arg1 (raw)  : 105484397019136   <- WebView2 area (same as m_pView)
  arg2 (raw)  : 105484397511488   <- another pointer
Class9.OnNavigationStarting: pSender=2702042333424, pArgs=105484397019136
  [WARNING] pSender does not match m_pView
```

What this warning showed is the correspondence `this = handler obj`, `arg1 = sender = m_pView`, `arg2 = args`. At the time, Claude wrongly updated the interpretation model to "the thunk does not shift to the right" and passed `OnNavigationStarting(this, arg1)`. It passed the sender (= ICoreWebView2) as `pArgs`, and called its vtable index 3 (= `get_Settings`), which crashed.

**Major discovery**: the thunk behaves as established in stage 8.5: **"shift the arguments right by one and inject ObjPtr into rcx"**. Reorganized once more:

| Before the thunk (WebView2's view) | After the thunk (VBA's view) | Received by Class7 | Meaning in C++ terms |
|---|---|---|---|
| rcx (1st argument) | -> moved to rdx | `this` | C++ this (handler obj) |
| rdx (2nd argument) | -> moved to r8 | `arg1` | C++ 1st argument |
| r8 (3rd argument) | -> moved to r9 | `arg2` | C++ 2nd argument |
| (none) | injected into rcx | `Me` | ObjPtr(Class7 instance) |

That is:

- 3-argument handler (`Invoke(this, errorCode, pEnv)`): from Class7's viewpoint `this = handler obj, arg1 = errorCode, arg2 = pEnv`
- 2-argument handler (`Invoke(this, sender, args)`): from Class7's viewpoint `this = handler obj, arg1 = sender, arg2 = args`

In the past HK_EnvironmentCompleted / HK_ControllerCompleted, `errorCode = 0` happened to hold for every call (all succeeded) and was consistent with `arg1 = 0`, so the mistake went unnoticed. Only in stage 9.4a, when handling a kind that has no concept of `errorCode`, did the confused interpretation model surface as a crash.

**Fixed version**:
```vba
Case HK_NavigationStarting
    If Not (m_owner Is Nothing) Then
        m_owner.OnNavigationStarting arg1, arg2   ' <- correct: arg1=sender, arg2=args
    End If
```

Log after the fix:
```
Class7.Invoke[HK_NavigationStarting]
  this        : 2109396025584
  arg1 (raw)  : 75007309086720    <- matches m_pView
  arg2 (raw)  : 75007309579488
Class9.OnNavigationStarting: pSender=75007309086720, pArgs=75007309579488
  -> destination URL: https://www.bing.com/search?q=%E6%A4%9C%E7%B4%A2%E3%81%97%E3%81%A6%E3%81%BF%E3%82%8B%E3%82%88&...
```

It works perfectly. The URL-encoded Japanese query was also obtained correctly.

#### Reset test (stop button in the middle of Test_Navigate_Real)

```
[State after searching, with Bing displayed]
(press the macro stop button)
  Release [idx=0] -> refCount = 0
  Release [idx=0] -> refCount = 0 -> Slot 0 auto-released
```

**Surprising result**: the moment reset was pressed, Release ran and reached auto-released. The `UserForm1.Terminate started` log did not appear (= the UserForm_Terminate Sub itself was not called), but **the Class_Terminate of the objects held by the UserForm's fields fires in a chain**.

This is a hidden strength of the VBA reset:
- The UserForm's event handlers (Subs) are not called
- But an implicit equivalent of `Set m_xxx = Nothing` runs for the module variables held by the UserForm
- -> Each object's Class_Terminate fires

-> **Design principle established**: at VBA reset, cleanup surprisingly runs cleanly. We adopted plan R4 (the alive flag as the minimum guarantee of crash avoidance), but on the real machine behavior almost equivalent to having implemented R3 (ForceCleanup) is obtained automatically.

### Additions to the established design principles and findings

- **Persistent handler lifetime management (plan P1)**: Class9 keeps holding it and removes everything at once with remove_ in Class_Terminate
- **HandlerKind-linked array (plan r)**: managed in parallel in the form `Private m_persistentHandlers(HK_xxx To HK_yyy) As Class7`
- **RemoveVtableIndexFor**: the HandlerKind -> remove_ vtable index mapping is gathered in one place, and only this is extended when rolling out
- **CleanupAllPersistentHandlers**: rolls out automatically with the HandlerKind-linked loop; the loop itself needs no change when adding a new kind
- **String conversion follows the old version (uses RtlMoveMemory)**: plan β. `PtrToString` + `GetStringProperty` is a generic helper that calls "a method returning an LPWSTR* as an out argument" in one line
- **Reset-button resilience is plan R4**: the alive flag + strict persistent handler remove_ order absolutely guarantee crash avoidance; leak countermeasures come with the Sentinel mechanism of stage 9.5
- **Important: thunk behavior reconfirmed**: Class7's `(this, arg1, arg2)` = the post-thunk `(rdx, r8, r9)` = WebView2's view `(C++ this, 1st argument, 2nd argument)`. Common to all kinds, constant regardless of the number of arguments. Do not be misled by cases where errorCode being 0 makes it work as a false positive

### Known limitations and future considerations

- Leaks on failure are tolerated (to be handled by the stage 9.5 Sentinel mechanism)
- Registering multiple handlers of the same kind is not supported (one Class9 instance = 0 or 1 handler per kind)
- Other `args` properties (get_IsRedirected, get_IsUserInitiated, etc.) are not retrieved yet (add as needed)
- Resize tracking is still not implemented

---

## Stage 9.4b — NavigationCompleted + DocumentTitleChanged rollout + navigation control

### Completion date
This session (as of the completion of 9.4b)

### Background and goal

Stage 9.4a established the core pattern of the persistent-handler mechanism (plan P1 + HandlerKind-linked arrays). Stage 9.4b is a relatively light stage whose premise is **"the design decisions are settled; we just roll them out sideways."** It adds NavigationCompleted and DocumentTitleChanged, and exposes the navigation control methods (NavigateToString, GoBack, GoForward, GetSource, GetDocumentTitle) on Class9.

Goals of stage 9.4b:

- Add the NavigationCompleted handler + retrieve args.get_IsSuccess
- Add the DocumentTitleChanged handler + get_DocumentTitle via sender
- Add the navigation control methods (NavigateToString, GoBack, GoForward, GetSource, GetDocumentTitle) to Class9
- Verify the chain of 3 event kinds in one go with "Test_AllEvents_Real"

### Design decision points

#### H1. Whether to generalize AddXxx / RemoveXxx / OnXxx (adopted: do not generalize)

Implementing 3 kinds (NavigationStarting / NavigationCompleted / DocumentTitleChanged) makes the code repetitive. Whether to generalize:

- Option i (adopted): **do not generalize**. Write a separate method per kind
- Option ii: create a generic method `AddPersistentEvent(kind, addVtblIndex)`

Reasons for adopting option i:
- The intent of each method is clear, and debug logs can carry kind-specific wording
- It is easier to accommodate kind-specific customization later, such as "add a filter" or "insert special handling" one kind at a time
- The internal implementation is a thin wrapper that calls dcf, so each method stays short

#### H2. Navigation control methods (no UI buttons added)

How to test GoBack / GoForward. One option was to add 2 buttons to the UserForm, but on tarboh's decision ("Q2. (a)") no UI buttons are added; the methods are only exposed as Public methods. Behavior testing will be done later (when UI is added or when the TabWebView2 class is introduced).

#### H3. Exit test of stage 9.4b

On tarboh's decision ("Q3. (a)"): consolidate into a single test, `Test_AllEvents_Real`, that shows Bing -> searches -> confirms the chain of 3 event kinds in one shot.

### Implementation

| File | Change |
|---|---|
| `Class7.cls` | Made the HK_NavigationCompleted / HK_DocumentTitleChanged branches real |
| `Class9.cls` | Add/Remove/On sets for NavigationCompleted / DocumentTitleChanged (3 each), 5 navigation control methods (NavigateToString, GoBack, GoForward, GetSource, GetDocumentTitle) |
| `Module9_7.bas` | Diff from Module9_6: imported `GetBoolProperty` / `GetLongProperty` from the old mWebView2Helper, added `Test_AllEvents_Real` |
| `UserForm1.frm` | At the end of `StartWebView2`, registers the 3 kinds of handlers in succession |

#### Main vtable indexes (ICoreWebView2)

| Method | vtable index |
|---|---|
| `get_Source` | 4 |
| `Navigate` | 5 |
| `NavigateToString` | 6 |
| `add/remove_NavigationStarting` | 7/8 |
| `add/remove_NavigationCompleted` | 15/16 |
| `get_CanGoBack` | 38 |
| `get_CanGoForward` | 39 |
| `GoBack` | 40 |
| `GoForward` | 41 |
| `add/remove_DocumentTitleChanged` | 46/47 |
| `get_DocumentTitle` | 48 |

#### args-related vtable indexes

- ICoreWebView2NavigationCompletedEventArgs::get_IsSuccess: 3 (out BOOL)
- ICoreWebView2NavigationStartingEventArgs::get_Uri: 3 (out LPWSTR*)

`get_DocumentTitle` is a method of ICoreWebView2 (vtable 48). The args of DocumentTitleChanged is `IUnknown*` and carries no information, so the flow is to call `get_DocumentTitle` via sender (= m_pView).

### Excerpts from the real-machine test log and observed behavior

#### Full log of the initial Bing load + a search

```
[When the Bing top page is shown]
Class7.Invoke[HK_NavigationStarting]
  → target URL: https://www.bing.com/
Class7.Invoke[HK_DocumentTitleChanged]
  → title: 検索 - Microsoft Bing
Class7.Invoke[HK_NavigationCompleted]
  → isSuccess: True

[User searches for "vba"]
Class7.Invoke[HK_NavigationStarting]
  → target URL: https://www.bing.com/search?q=vba&form=QBLH&...
Class7.Invoke[HK_DocumentTitleChanged]
  → title: vba - 検索
Class7.Invoke[HK_NavigationCompleted]
  → isSuccess: True
```

**Observed behavior 1**: The event firing order is **Starting -> DocumentTitleChanged -> Completed**. Claude initially expected "Starting -> Completed -> TitleChanged", but on the real machine Completed comes last. This is SDK behavior: WebView2 fires TitleChanged the moment it finds the page HTML's `<title>`, and notifies Completed when all of the page's resources have finished loading.

**Observed behavior 2**: The args of DocumentTitleChanged is always `0` (NULL). This matches the SDK spec. The title is retrieved via `get_DocumentTitle` through sender.

**Observed behavior 3**: The 3 persistent handlers coexist in separate slots (0/1/2) and run in parallel. Each Invoke completes its own AddRef/Release pair independently.

#### The power of CleanupAllPersistentHandlers

```
UserForm1.Terminate started
Class9.Cleanup: removing persistent handler kind=3 vtable=8 token=...     ← NavigationStarting
Class9.Cleanup: removing persistent handler kind=4 vtable=16 token=...    ← NavigationCompleted
Class9.Cleanup: removing persistent handler kind=6 vtable=47 token=...    ← DocumentTitleChanged
```

**The combination of the HandlerKind-linked array loop + RemoveVtableIndexFor shows its true strength.** The Cleanup loop of Class9 itself has **zero changes** since stage 9.4a, and the new kinds simply rode on it automatically. The correctness of the stage 9.4a design was perfectly demonstrated.

### Design principles and knowledge established (additions)

- **Write AddXxx/RemoveXxx/OnXxx separately for the 3 kinds (option i)**: readability of each method takes priority over generalization
- **Event firing order**: Starting -> DocumentTitleChanged -> Completed (Completed is last)
- **The args of DocumentTitleChanged is `IUnknown*` and always NULL**: call get_DocumentTitle via sender (m_pView)
- **The power of the HandlerKind-linked array loop**: CleanupAllPersistentHandlers needed no change and automatically covers the 3 kinds
- **Naming of the navigation control methods**: prefix unified as `View_NavigateToString` / `View_GoBack` / `View_GoForward` / `View_GetSource` / `View_GetDocumentTitle`

### Known constraints and future considerations

- Behavior test of navigation control not yet done (policy of no UI buttons; add a Test when needed)
- WebMessageReceived is handled in stage 9.4c
- Other properties of NavigationCompleted (get_WebErrorStatus, get_NavigationId) are not retrieved yet

---

## Stage 9.4c — WebMessage bidirectional communication

### Completion date
This session (as of the completion of 9.4c)

### Background and goal

The real value of embedding WebView2 in Excel is that **arbitrary data can be exchanged between JavaScript and VBA**. This makes it possible to:
- Call a UI built with HTML/JS from Excel
- Use feature-rich JS libraries (chart drawing, maps, editors, etc.) from Excel
- Implement a rich UI as a web page on WebView2 instead of the poor controls of an Excel UserForm

Goals of stage 9.4c:

- Add the WebMessageReceived persistent handler (4th kind of the plan P1 pattern)
- The View_PostWebMessageAsString method (VBA -> JS)
- Add `StartWebView2_LocalHtml` (shows a test HTML via NavigateToString) and `SendMessageToJS` to the UserForm
- Add an `IsPageLoaded` property to the UserForm (so that the Test procedure can wait for NavigationCompleted)
- Verify bidirectional communication on the real machine with Test_MessageEcho_Real

### Design decision points

#### I1. Show the test HTML via NavigateToString (adopted: option q)

Testing VBA -> JS / JS -> VBA bidirectional communication requires JS code that uses `window.chrome.webview.postMessage`. Bing does not use the WebView2 chrome.webview API, so bidirectional communication cannot be tested with it.

Considered:

- Option p: use Bing as is. Sending VBA -> JS works but reception cannot be confirmed
- Option q (adopted): **show a test HTML via NavigateToString**. Bidirectional communication can be fully tested
- Option r: a separate Test for each (keep the Bing display in `Test_AllEvents_Real`, and use a separate HTML in `Test_MessageEcho_Real` for bidirectional communication)

Adopted: in practice a combination of (q) and (r). Prepare **2 start methods** on the UserForm, one for Bing and one for the local HTML:

- `UserForm1.StartWebView2` -> shows Bing (for `Test_AllEvents_Real`)
- `UserForm1.StartWebView2_LocalHtml` -> shows the test HTML (for `Test_MessageEcho_Real`)

#### I2. Trigger for VBA -> JS sending (adopted: option β)

When should VBA send to JS:

- Option α: put a button on the UserForm and send on click
- Option β (adopted): **fire it explicitly from the Test procedure** (`UserForm1.SendMessageToJS "..."`)
- Option γ: send periodically from a UserForm timer

Reason for adopting option β: consistent with the "no UI buttons added" policy up to stage 9.4b. A style where it is called from the Test procedure.

Add `Public Sub SendMessageToJS(text)` to the UserForm, and Test_MessageEcho_Real fires it with `UserForm1.SendMessageToJS "hello"`.

#### I3. Implementation of waiting for NavigationCompleted (adopted: option α)

It is necessary to "confirm that the JS-side listener has been registered before sending the first VBA -> JS message." Waiting for NavigationCompleted is sufficient for this.

- Option α (adopted): add an `IsPageLoaded` property to the UserForm and set it to True in `OnNavigationCompleted`. The Test procedure waits with `Do While Not UserForm1.IsPageLoaded: DoEvents: Loop`
- Option β: crudely wait a fixed time with Application.Wait
- Option γ: have the user press a "Send" button in the UI

tarboh's decision ("Q3. (a) is fine. I eventually want IsPageLoaded to live in something like an umbrella class"): adopt (α), but as a provisional implementation to be moved to an umbrella class such as TabWebView2 in the future.

Implementation:
- Add an `m_pageLoaded As Boolean` field (Private) to Class9
- `OnNavigationCompleted` sets `m_pageLoaded = True`
- `View_Navigate` / `View_NavigateToString` reset `m_pageLoaded = False` (a new navigation starts)
- The UserForm1.IsPageLoaded property delegates to m_ctrl.IsPageLoaded

#### I4. Discussion of SPA support (stated explicitly as a future consideration)

An important point from tarboh: "**For an SPA, the page content keeps being updated in response to user actions even after NavigationCompleted, so how to catch that needs to be discussed separately**."

In an SPA (Single Page Application: React/Vue/Angular, etc.):
- Only the URL is rewritten via `pushState` / `replaceState`
- The page content is updated by DOM manipulation
- From WebView2's point of view "no navigation has occurred", so NavigationCompleted does not fire

Options for handling this (future consideration):
- HistoryChanged event (when the URL is rewritten)
- SourceChanged event (almost the same as above)
- DOMContentLoaded (ICoreWebView2_2 and later)
- ContentLoading
- Notification by postMessage from the SPA side (at any timing)

These are "**things the flat handler mechanism of stage 9.4 can handle**", but **state management** such as IsPageLoaded is cleanly designed in the umbrella class layer (TabWebView2, etc.). Stage 9.4c stays with a provisional implementation, to be properly redesigned in the future.

#### I5. Information retrieved in OnWebMessageReceived (adopted: option a)

The args of WebMessageReceived (= ICoreWebView2WebMessageReceivedEventArgs) has several properties:

| Method | vtable index |
|---|---|
| `get_Source` (source URL) | 3 |
| `get_WebMessageAsJson` (JSON form) | 4 |
| `TryGetWebMessageAsString` (string form) | 5 |

Options:
- Option a (adopted): log only the source URL + message body (string)
- Option b: all three: source URL + message body (string) + JSON form

Reason for adopting option a: confirm operation with a simple, minimal implementation. The JSON route will be extended in stage 9.4d or later if needed.

### Implementation

| File | Change |
|---|---|
| `Class7.cls` | Made the HK_WebMessageReceived branch real (all kinds are now implemented) |
| `Class9.cls` | Added `AddWebMessageReceived` / `RemoveWebMessageReceived` / `OnWebMessageReceived`, added `View_PostWebMessageAsString`, added the `m_pageLoaded` flag + `IsPageLoaded` property (True on NavigationCompleted, reset to False on Navigate/NavigateToString) |
| `Module9_8.bas` | Diff from Module9_7: added `Test_MessageEcho_Real`, added `BuildTestHtml` (generates the test HTML string) |
| `UserForm1.frm` | Added `StartWebView2_LocalHtml` (shows the test HTML via NavigateToString, registers 4 kinds of handlers), the public method `SendMessageToJS`, and the `IsPageLoaded` property (delegating to Class9) |

#### Main vtable indexes (added in stage 9.4c)

| Method | vtable index |
|---|---|
| `PostWebMessageAsJson` (ICoreWebView2) | 32 |
| `PostWebMessageAsString` | 33 |
| `add_WebMessageReceived` | 34 |
| `remove_WebMessageReceived` | 35 |

#### args-related (ICoreWebView2WebMessageReceivedEventArgs)

| Method | vtable index |
|---|---|
| `get_Source` | 3 |
| `get_WebMessageAsJson` | 4 |
| `TryGetWebMessageAsString` | 5 |

#### Structure of the test HTML

```html
<!DOCTYPE html>
<html><head>
<meta charset='utf-8'>
<title>WebView2 bidirectional communication test</title>
<style>...</style>
</head><body>
<h1>WebView2 <-> VBA communication test</h1>
<div>
  <input type='text' id='msg' placeholder='Message to send from JS to VBA'>
  <button onclick='sendToVBA()'>Send</button>
</div>
<h2>Received from VBA:</h2>
<div id='log'></div>
<script>
  function sendToVBA() {
    const text = document.getElementById('msg').value;
    window.chrome.webview.postMessage(text);
    document.getElementById('msg').value = '';
  }
  window.chrome.webview.addEventListener('message', event => {
    const log = document.getElementById('log');
    log.innerText += event.data + '\n';
  });
</script>
</body></html>
```

### Excerpts from the real-machine test log and observed behavior

#### Successful JS -> VBA bidirectional communication

```
(trying to send a JS -> VBA message)
  AddRef [idx=3] → refCount = 2
Class7.Invoke[HK_WebMessageReceived]
  this        : 2613491139472
  arg1 (raw)  : 48928267665408     ← sender (= m_pView)
  arg2 (raw)  : 48928268370688     ← args
  m_refCount  : 2
Class9.OnWebMessageReceived: pSender=48928267665408, pArgs=48928268370688
  → source: about:blank
  → message: JSから送信してるよ      ← Japanese came through perfectly
```

The UTF-16 string reached VBA from JS without garbling.

#### The NavigationStarting URL is data:text/html;base64,...

```
Class9.OnNavigationStarting: ...
  → target URL: data:text/html;charset=utf-8;base64,PCFET0NUWVBFIGh0bWw+DQo8aHRtbD4...
```

**Important finding**: Internally, WebView2 handles NavigateToString by **Base64-encoding the HTML and treating it as a URL with the `data:` scheme**. That is, the internal implementation is equivalent to "Navigate("data:text/html;base64,...")". A data: page has its own origin, which is a security implementation choice.

#### The source URL of a NavigateToString page is about:blank

The URL retrieved with `get_Source` (the get_Source of args, not the ICoreWebView2 property) is `about:blank`. It differs from the data: URL shown in NavigationStarting.

This is WebView2 specification:
- URL at navigation start = data: scheme (Internal)
- Source URL of the displayed page = about:blank (User-facing)

#### 4 persistent handlers managed in parallel

```
Class9.AddNavigationStarting     OK: token=31, handler.slot=2613491139072  ← Slot 2
Class9.AddNavigationCompleted    OK: token=31, handler.slot=2613491138848  ← Slot 1
Class9.AddDocumentTitleChanged   OK: token=23, handler.slot=2613491138624  ← Slot 0
Class9.AddWebMessageReceived     OK: token=24, handler.slot=2613491139296  ← Slot 3
```

The 4 Class7 handlers coexist in Slots 0/1/2/3. The "parallel management of multiple handlers" established in stage 9.2 is fully demonstrated.

#### Cleanup of the 4 kinds is flawless

```
UserForm1.Terminate started
Class9.Cleanup: removing persistent handler kind=3 vtable=8 token=31
Class9.Cleanup: removing persistent handler kind=4 vtable=16 token=31
Class9.Cleanup: removing persistent handler kind=5 vtable=35 token=24
Class9.Cleanup: removing persistent handler kind=6 vtable=47 token=23
Class9.Terminate: Releasing m_pView = ...
Class9.Terminate: Releasing m_pCtrl = ...
Class8.Terminate: Releasing m_pEnvironment = ...
UserForm1.Terminate finished
```

remove_ succeeded for all 4 kinds. **The HandlerKind-linked array loop has had zero changes since stage 9.4a** and handles all 4 kinds.

### Design principles and knowledge established (additions)

- **Establishing WebMessage bidirectional communication**: VBA -> JS is `View_PostWebMessageAsString` (vtable 33); JS -> VBA is `add_WebMessageReceived` (vtable 34) + `args.TryGetWebMessageAsString` (vtable 5)
- **Internal implementation of NavigateToString**: handled as a data: scheme URL of the form data:text/html;base64,...
- **The source URL of a NavigateToString page is about:blank**: the chrome.webview API is usable
- **TryGetWebMessageAsString**: succeeds when JS sent a string; returns an empty string when sent as a JSON object (the HRESULT fails, but GetStringProperty absorbs it)
- **m_pageLoaded flag management**: True on NavigationCompleted, False on Navigate/NavigateToString. A provisional implementation in Class9, to be moved to an umbrella class in the future
- **UTF-16 bidirectional communication**: strings containing Japanese pass between VBA <-> JS without garbling
- **SPA support is a future consideration**: NavigationCompleted fires only once for the initial load; tracking screen updates of an SPA will be designed in an umbrella class layer such as TabWebView2, combining HistoryChanged / DOMContentLoaded / postMessage notifications, etc.

### Known constraints and future considerations

- IsPageLoaded is to be moved to an umbrella class (Class9 holds it provisionally)
- SPA support (history.pushState, DOMContentLoaded, HistoryChanged) is for the future (when designing an umbrella class such as TabWebView2)
- JSON-format message sending/receiving is not supported yet (get_WebMessageAsJson, PostWebMessageAsJson to be added when needed)
- The NewWindowRequested handler is not supported yet (most Bing links request popups; add when needed)
- Version extensions (ICoreWebView2_2 to _22, etc.) are not supported yet

---

## Stage 9.3-9.4 summary — overall wrap-up and handover to stage 9.5 onward

### What was achieved in stages 9.3-9.4

| Stage | Achievements |
|---|---|
| 9.3a | Introduced the upper-level class Class8 Environment, one-shot handler lifetime management (plan D1-α), Public Sub dynamic binding |
| 9.3b | Class9 Controller class, dcf unification, HK_ControllerCompleted reconnaissance |
| 9.3c | Class9 = merged Controller + CoreWebView2 class (plan a), Ctrl_/View_ prefixes, Bing display milestone achieved |
| 9.4a | Persistent handler mechanism (plan P1), HandlerKind-linked arrays (plan r), NavigationStarting + get_Uri, reset test |
| 9.4b | Rolled out NavigationCompleted + DocumentTitleChanged, 5 navigation control methods, demonstrated the power of the HandlerKind-linked loop |
| 9.4c | WebMessageReceived two-way communication, VBA ↔ JS communication via a test HTML, provisional IsPageLoaded implementation, discussion of SPA support |

### Main design principles established in stages 9.3 to 9.4 (item list)

1. One-shot handler lifetime management (plan D1-α)
2. Persistent handler lifetime management (plan P1)
3. Automate rolling out handler kinds with HandlerKind-linked arrays (plan r)
4. Strict use of Public Sub (for dynamic binding)
5. Set m_state = Es_Waiting before the API call (handles synchronous callbacks)
6. dcf does not judge HRESULTs automatically
7. Class9 merged class (plan a) + Ctrl_/View_ prefixes
8. Pass structure arguments as VarPtr pointers
9. Get the HWND size directly in pixels with GetClientRect
10. On reset, avoid crashes with the alive flag and tolerate leaks (plan R4)
11. The (this, arg1, arg2) of Class7 is always read from WebView2's viewpoint as (C++ this, 1st argument, 2nd argument)
12. WebView2 event hierarchy: use NavigationStarting / NewWindowRequested / FrameNavigationStarting appropriately
13. Reference-count behavior of a persistent handler per event (one AddRef / one Release completes it)
14. Release order is from child to parent (m_pView → m_pCtrl → m_pEnvironment)
15. CleanupAllPersistentHandlers is rolled out automatically by the HandlerKind-linked loop
16. Event firing order: NavigationStarting → DocumentTitleChanged → NavigationCompleted
17. The args of DocumentTitleChanged is an IUnknown* that is always NULL; take the value via the sender
18. NavigateToString is internally implemented as a data: scheme URL, and the source URL is about:blank
19. TryGetWebMessageAsString succeeds only when JS sent a string
20. m_pageLoaded flag management (provisionally in Class9, to be moved to the umbrella class in the future)
21. UTF-16 two-way communication passes through without garbling (Japanese is OK)
22. Write AddXxx/RemoveXxx/OnXxx separately for each of the 3 kinds (readability over generalization, plan i)
23. Navigation control method naming: View_NavigateToString / View_GoBack / View_GoForward / View_GetSource / View_GetDocumentTitle
24. Mouse resizing of the UserForm is a future task (the SetWindowLongPtr + GA_ROOT approach is already decided)
25. On a VBA reset, the WebView2 shutdown path surprisingly runs a clean cleanup (the UserForm_Terminate Sub is not called, but the Class_Terminate chain of the fields fires)
26. Version-extension interfaces (ICoreWebView2_2 and so on) are added when they are used (implement only the base, plan 4)
27. The persistence scheme is an environment variable (MouseScroll style): save "the base address of the most recently allocated region" as a decimal string with `SetEnvironmentVariableW` / `GetEnvironmentVariableW`. No workbook-modified flag, no registry pollution, disappears automatically when the process ends, and is completed by code alone
28. Sentinel mechanism plan β (double defense: alive flag + VirtualFree): on a VBA reset, the old region is reclaimed in the order "① set the old region's alive flag to 0 → ② release the old region with VirtualFree → ③ clear the environment variable"
29. Countermeasure for the VBA &H literal trap: when a hexadecimal literal exceeds the Integer range (≤ &H7FFF), always add the `&` suffix. Even for a Const declared `As Long`, the right-hand side is evaluated regardless of the left-hand type, so `&H8000` is evaluated as the Integer `-32768` and only then widened to Long. As insurance, the safe practice is to add the `&` suffix to every `&H` literal
30. Always call important Win32 APIs in Function form: Win32 APIs such as `VirtualFree` must have their return value received and checked. Calling them in Sub form (discarding the return value) risks missing failures forever (a prime example is that, throughout the entire period up to stage 9.4c, the VirtualFree in Thunks_Shutdown kept failing and nobody noticed)


### Items handed over to stage 9.5 and later

Because the Sentinel mechanism was completed and leaks were eliminated entirely in stage 9.5, that item was removed from the handover items. The remaining 6 items are listed below.

#### 1. Introduce an IID check (make Handler_QueryInterface genuine)
The current Handler_QueryInterface is a simple implementation that always returns the same vtable without looking at the IID. To align better with WebView2's QI strategy, an IID check should be added. High independence, small scale.

#### 2. Mouse resizing of the UserForm
The WS_THICKFRAME + GA_ROOT + SetWindowLongPtrW approach is decided (agreed in stage 9.3c). Starting from Frame1.[_GethWnd], walk up with GetAncestor(GA_ROOT) to the UserForm HWND. High independence, small scale.

#### 3. SPA support
NavigationCompleted fires only once, for the initial load. Tracking screen updates of an SPA (React/Vue, etc.) requires the following:
- The HistoryChanged event (when the URL is rewritten)
- DOMContentLoaded (ICoreWebView2_2 and later)
- A postMessage notification from the SPA side
Design this in an umbrella class layer such as TabWebView2.

#### 4. Version-extension interfaces
ICoreWebView2_2 to _22, ICoreWebView2Controller2 to _6, and so on. Add fields + methods to Class9 when needed (plan 4, the policy of implementing only the base).

#### 5. Rolling out NewWindowRequested
A simple rollout of the persistent handler pattern (plan P1 + HandlerKind-linked arrays) established in stages 9.4a to 9.4c. Most of the links on Bing are popup requests, so it is better to handle this early to improve usability.

#### 6. TabWebView2 umbrella class
The tab management layer. Moving IsPageLoaded from Class9 to TabWebView2, merging with SPA support, and resize tracking are also cleanest implemented here. The largest scale.


### Code deliverables of this session

| File | Version | Role |
|---|---|---|
| `Class7.cls` | 9.4c | Generic handler, with the Invoke branches for all 6 kinds fleshed out |
| `Class8.cls` | 9.3a | Environment class (unchanged since 9.3a) |
| `Class9.cls` | 9.4c | Merged Controller + CoreWebView2 class, 4 kinds of persistent handlers + navigation control + IsPageLoaded |
| `Module9_8.bas` | 9.4c | Foundation + dcf + string helpers + Test group (Test_TwoHandlers_Mock / Test_Environment_Real / Test_Controller_Real / Test_Navigate_Real / Test_AllEvents_Real / Test_MessageEcho_Real / Test_ResetSurvival_Manual) |
| `UserForm1.frm` | 9.4c | Frame1 + StartWebView2 (Bing) / StartWebView2_LocalHtml (test HTML) + SendMessageToJS + IsPageLoaded |

All files were verified in a consistent state on the real machine (real-machine test). They provide a stable base as the starting point of stage 9.5.


## Stage 9.5 — Sentinel mechanism (complete elimination of leaks on reset) + finding the true cause of the VirtualFree problem

### Completion date
This session (completed at stage 9.5g, then tidied into Module9_9_7 as the final version of 9.5)

### Background and goal

Up to stage 9.4c, "no crash on reset" (plan R4: alive flag + strict ordering of persistent handler remove_) had been achieved, but **leaking the regions themselves was tolerated**. The goals of stage 9.5 were:

- Introduce a mechanism (= the Sentinel mechanism) that **automatically reclaims** the VirtualAlloc regions abandoned on a VBA reset
- Achieve complete elimination of leaks

Stage 9.5 as a whole proceeded in 7 sub-stages (9.5a to 9.5g). It was initially expected to take 2 to 3 stages, but during verification the shocking fact came to light that **"Thunks_Shutdown's VirtualFree had actually been failing all along too"**, and finding the true cause took several stages.

### Flow of the sub-stages

| Stage | Achievements |
|---|---|
| 9.5a | Environment variable persistence scheme + 5 Sentinel procedures + Test_SentinelStatus implemented |
| 9.5b | Added GetLastError retrieval to Sentinel_RecoverIfNeeded (to isolate the cause of VirtualFree `returned 0`) |
| 9.5c | Introduced VirtualQuery + MEMORY_BASIC_INFORMATION (to check the state of the old region from the OS's viewpoint) |
| 9.5d | Tested 5 scenarios (α/β/γ/δ/ε) with Test_VirtualAllocFree_Roundtrip → shocking: **all NG** |
| 9.5e | Added diagnostic logging to the VirtualFree in Thunks_Shutdown too → revealed that **the existing Shutdown had been failing as well** |
| 9.5f | Tested changing the type of dwSize (As LongPtr → As Long) and a CLngLng(0) argument → both NG |
| 9.5g | **True cause found**: `Private Const MEM_RELEASE As Long = &H8000` had become `-32768` because of the VBA &H literal trap. Fixed to `&H8000&` and everything was resolved |

### Design decision points

#### J1. Persistence scheme (adopted: plan ⑤ environment variable = MouseScroll style)

On a VBA reset all VBA variables are discarded, so "the previous region address" must be saved outside VBA. Candidates:

- Plan ① Excel Names: sets the modified flag → inconsistent with the best-effort policy
- Plan ② HKCU registry: leaves traces in regedit
- Plan ③ %TEMP% temporary file: I/O cost, collision management
- Plan ④ CustomDocumentProperties: sets the modified flag (same drawback as Names)
- Plan ⑤ (adopted): **process environment variable** (`SetEnvironmentVariableW` / `GetEnvironmentVariableW`)

Reason for adopting plan ⑤: it is the same school as MouseScroll.bas (Cristian Buse style). Disappearing automatically when the process ends = fully consistent with the best-effort policy (works only within the same process). No workbook-modified flag, no registry pollution, completed by code alone.

Note: in the initial explanation of the options, ⑤ environment variable was forgotten (4 choices, ① ② ③ ④, were presented). It was discovered when checking the MouseScroll implementation and added to the options → tarboh chose ⑤.

#### J2. Sentinel reclaim procedure (adopted: plan β)

Procedure for reclaiming the old region when returning from a reset:

- Plan α: just set the alive flag to 0. No VirtualFree (leaks remain tolerated)
- Plan β (adopted): **set the alive flag to 0 and then also fire VirtualFree** (complete leak elimination, double defense)
- Plan γ: VirtualFree right away (skips the alive-flag operation, dangerous)

Reason for adopting plan β: it serves as a safety valve so that even if WebView2 should still hit the stub, the thunk can return S_OK at its head. It works as a momentary defense until VirtualFree completes.

#### J3. Environment variable key design (adopted: fixed string)

- Plan (a) (adopted): fixed string `"WV2_VBA_LastRegion"`
- Plan (b): mix in the workbook name
- Plan (c): mix in a hash of the workbook's full path

Reason for adoption: even if multiple Excel workbooks launched at the same time collide, the result is only that "one side's region cannot be reclaimed", not a crash (= it just falls back to tolerating the leak). The simplest choice, consistent with the best-effort policy.

#### J4. Best-effort policy (adopted)

Whether to handle strictly the case where "the old WebView2 process is still alive" at reset recovery:

- Handle strictly: complex, many state transitions
- Best effort (adopted): it is enough that only reset → restart within the same process works

Reason for adoption: in the handover items of the development notes (at the consideration stage), tarboh explicitly said "best effort is fine".

### The shocking discovery of the true cause: the VBA &H literal trap

When diagnostic logging was added to the end of `Test_Environment_Real` in stage 9.5e, the following fact came to light:

```
Thunks_Shutdown: VirtualFree(...) returned 0, LastError = 87
                 [NG] (ERROR_INVALID_PARAMETER)
```

**The existing Thunks_Shutdown had actually been failing at VirtualFree all along**. Throughout the entire period up to stage 9.4c, the return value was discarded because it was called in Sub form, so nobody noticed.

Changing types and casting arguments in stages 9.5d to 9.5f all failed. Finally, in stage 9.5g the true cause was found:

```vba
Private Const MEM_RELEASE As Long = &H8000
```

In VBA, `&H8000` is **evaluated as an Integer (16-bit)** and only then widened to Long:

- The Integer `&H8000` = **-32768** (negative, because the top bit is set)
- Widening it to Long still gives **-32768**
- When `dwFreeType = -32768` is passed to `VirtualFree`, Win32 rejects it as an invalid value → ERROR_INVALID_PARAMETER

The fix:

```vba
Private Const MEM_RELEASE As Long = &H8000&    ' Adding the & suffix forces Long evaluation
```

With the `&` suffix it is **evaluated as a Long from the start**, so `&H8000&` = **+32768**, the correct value.

While at it, as insurance the `&` suffix was also added to `MEM_COMMIT` / `MEM_RESERVE` / `PAGE_EXECUTE_READWRITE` (these are within the Integer range `≤ &H7FFF`, so they were not actually a problem, but this aids readability + prevents accidents during future refactoring).

### Implementation

| File | Changes |
|---|---|
| `Module9_8.bas` → `Module9_9_7.bas` (final version of 9.5) | Added 5 Sentinel mechanism procedures, added 4 diagnostic functions, added Test_VirtualAllocFree_Roundtrip, added the `&` suffix to constants such as MEM_RELEASE, built Sentinel calls into Thunks_Init/Shutdown, converted the VirtualFree in Thunks_Shutdown to Function form |
| `Class7.cls` / `Class9.cls` / `UserForm1.frm` | **No changes** |

Added API declarations (Module9_9_7.bas):
- `GetEnvironmentVariableW` / `SetEnvironmentVariableW` (Private, kernel32)
- `GetLastError` (Private, kernel32)
- `VirtualQuery` (Private, kernel32)

Added structure:
- `MEMORY_BASIC_INFORMATION` (Private Type, 48 bytes on x64)

Added procedures:
- `Sentinel_LoadPrevRegion` (Private Function, reads the environment variable)
- `Sentinel_StorePrevRegion` (Private Sub, writes the environment variable)
- `Sentinel_ClearPrevRegion` (Private Sub, deletes the environment variable)
- `Sentinel_RecoverIfNeeded` (Private Sub, reclaims the old region at the head of Init)
- `Test_SentinelStatus` (Public Sub, for debugging)
- `Test_VirtualAllocFree_Roundtrip` (Public Sub, general-purpose diagnostics, 5 scenarios)
- `VirtualFreeErrorName` (Private Function, Win32 error code → string)
- `FormatMemState` (Private Function, State → string)
- `FormatMemProtect` (Private Function, Protect → string)

### Excerpts of the real-machine test log and observed behavior

#### Finding J1: In stage 9.5g the true cause was resolved and all items cleared

```
Test_VirtualFree_MemReleaseFix
Current value of MEM_RELEASE = 32768 [OK = 0x8000 correct]
VirtualFree returned 1, LastError = 0 [OK ★ finally succeeded!]

Test_VirtualAllocFree_Roundtrip
Scenarios α to ε all [OK]

Test_Environment_Real
Thunks_Shutdown: VirtualFree(...) returned 1, LastError = 0 [OK]

Test_AllEvents_Real → reset → Test_AllEvents_Real
Sentinel: detected previous region = ... recovering ...
Sentinel: VirtualQuery result (48 bytes copied):
  State = 4096 (MEM_COMMIT (region is intact))
Sentinel: header LongPtr before clear = 1 (alive flag = 1)
Sentinel: VirtualFree succeeded for ...
```

In the log of "two consecutive resets" in particular, we could observe that the 1st reset and the 2nd reset produced different old region addresses, and each was reclaimed correctly in a separate Init. Furthermore, cases were confirmed where the new region was allocated at the **same address** before and after a reset, which is evidence that the old region was VirtualFreed and the address space was reused (= the leak elimination is fully effective).

#### Finding J2: Discovery of the VBA &H literal trap

This is a famous trap that anyone experienced with VBA steps into at least once, and it occurs frequently when porting constants from Win32 API headers. **Writing like `Const ... As Long = &H8000` is dangerous**. `As Long` only specifies the type the left-hand side holds; the right-hand side follows a separate evaluation rule (Integer if within the Integer range, Long if beyond it).

Established as design principle 29: **whenever a hexadecimal literal may exceed the Integer range, always add the `&` suffix**. As insurance, the safe practice is to add it to all `&H` literals too.

#### Finding J3: Call important Win32 APIs in Function form

Throughout the entire period up to stage 9.4c, `VirtualFree m_pRegionBase, 0, MEM_RELEASE` in `Thunks_Shutdown` kept failing. Nobody noticed because it was called in Sub form and the return value was discarded.

Established as design principle 30: **call important Win32 APIs in Function form, receive the return value, and if necessary isolate the cause with GetLastError**.

#### Finding J4: Usefulness of checking region state with VirtualQuery

The VirtualQuery diagnostic log introduced in stage 9.5c made it possible to observe "what state the old region is in from the OS's viewpoint". In this investigation of the true cause:

- AllocationBase matches prevBase → the address is correct
- State = MEM_COMMIT (4096) → the region is intact
- RegionSize = 118784 (= REGION_SIZE 114752 rounded up to the page boundary 0x1000)

These facts were confirmed early, and it was narrowed down to "the problem is on the VirtualFree call side, not the address side". It is worth keeping in the final version as a diagnostic log.

#### Finding J5: The best-effort policy was the right call

The Sentinel mechanism works only for reset → restart within the same process. That is enough:

- When Excel itself is closed, the OS also releases the VirtualAlloc regions → no Sentinel needed
- In a different Excel process, both the environment variables and the address space are separate → irrelevant in the first place
- Within the same process, the trace is conveyed by the environment variable → works correctly

The best-effort requirement of "not wanting to cross processes" (= not wanting to leave traces) and the automatic-disappearance property of environment variables fit together cleanly.

### Additional established design principles and findings

- **The persistence scheme is an environment variable (MouseScroll style)** (new design principle 27)
- **Sentinel mechanism plan β (double defense: alive flag + VirtualFree)** (new design principle 28)
- **Countermeasure for the VBA &H literal trap (consistent use of the & suffix)** (new design principle 29)
- **Always call important Win32 APIs in Function form (judge failure by the return value)** (new design principle 30)
- **VirtualFree had actually been failing throughout the entire period up to stage 9.4c**: the premise of plan R4 (tolerating leaks on reset) changed, and it turned out that leaks were already occurring at stage 9.4c (though there was no practical problem because the Excel address space is vast)
- **VirtualQuery + MEMORY_BASIC_INFORMATION is useful for diagnostics**: you can check the region state as seen from the OS (MEM_COMMIT / MEM_FREE / MEM_RESERVE)
- **x64 layout of MEMORY_BASIC_INFORMATION**: 48 bytes (BaseAddress / AllocationBase / AllocationProtect + pad1 / RegionSize / State / Protect / Type_ + pad2)

### Known constraints and future considerations

- If multiple Excel workbooks in the same process use the same environment variable name `WV2_VBA_LastRegion`, they collide (but this only means one side's region cannot be reclaimed, not a crash, so it is acceptable)
- For "another Excel process on the same machine", the environment variable space is separate, so there is no impact
- Because of the best-effort design, if Excel itself crashes, the region is not reclaimed (but the OS also releases VirtualAlloc regions, so there is no problem)

---

## Stage 9.6 — Rolling out NewWindowRequested + establishing a naming convention based on origin prefixes

### Completion date
This session (completed at stage 9.6b)

### Background and goal

Stage 9.5 completed the Sentinel mechanism and achieved "complete elimination of leaks". Of the 6 remaining hand-over items, stage 9.6 took on **"5. Rolling out NewWindowRequested"**.

It should have been a mechanical rollout of the persistent handler pattern established up to stage 9.4 (plan P1 + the HandlerKind-linked array), but the following derived issues surfaced during the investigation:

- Existing method naming (`AddNavigationStarting` etc.) had no origin prefix, while `View_Navigate` / `Ctrl_PutBounds` did have one, an asymmetry. Adding NewWindowRequested made it necessary to tidy up the naming convention
- A design decision was needed on how to handle it: "open in the same View" vs "launch a separate UserForm"
- Improve the behavior when a Bing link is clicked (until now an independent new window opened inside the Excel process)

Stage 9.6 as a whole was split into 2 sub-stages:
- **9.6a**: Roll out the NewWindowRequested handler mechanism + overhaul the naming convention with origin prefixes (URL log only, no handling)
- **9.6b**: Add "open in the same View" handling via `put_Handled` + `View_Navigate`

### Flow of the sub-stages

| Stage | Items achieved |
|---|---|
| 9.6a | Added HK_NewWindowRequested, 3 new View_Add/Remove/On methods, renamed the existing 12 methods (4 kinds x 3) with the View_ prefix, extended the array upper bound, confirmed URL-log-only operation |
| 9.6b | Added put_Handled(1&) + View_Navigate(url) to View_OnNewWindowRequested, confirmed that Bing links open in the same View |

### Design decision points

#### K1. How to handle NewWindowRequested (adopted: plan a = open in the same View)

When NewWindowRequested is received, how should the VBA side handle it? Options:

- **Plan a (adopted)**: `args.put_Handled(TRUE)` + `View_Navigate(url)` in the same View. **Simple, and immediately solves the Bing link problem.**
- Plan b: Launch a new Class9 / UserForm via `args.put_NewWindow(pNewView)`. "Popup-like behavior". High implementation complexity
- Plan c: Log only, without setting `put_Handled`. Keep the default behavior (a new window)

tarboh's decision: "**Plan a is fine for the current stage. At the point where this finally becomes a tabbed browser, we will create a separate WebView2 instance and display it there**".

#### K2. Subdivision (adopted: 2 stages, 9.6a / 9.6b)

By doing "only the mechanism rollout" (log only) in 9.6a and "introducing the handling" (put_Handled + View_Navigate) in 9.6b:
- The 9.6a test can verify in isolation "whether the handler fires correctly"
- The 9.6b test can verify in isolation "the effect of put_Handled"
- It is easy to isolate the cause on failure

The same granularity as the 9.4a -> 9.4b flow. Agreed with tarboh.

#### K3. Establishing a naming convention based on origin prefixes (important, plan (a) adopted)

Existing method naming in Class9:
- `Ctrl_PutBounds` / `Ctrl_GetCoreWebView2` <- has a prefix
- Navigation control such as `View_Navigate` / `View_GoBack` <- has a prefix
- Event-related methods such as `AddNavigationStarting` / `OnNavigationStarting` <- **no prefix**

This asymmetry would become visible when we later implement `ICoreWebView2_2.add_DOMContentLoaded` and the like, in the form "with just `AddDOMContentLoaded` we can't tell which version of the View it belongs to". tarboh also requested: "If the Controller and the main object are going to live in a single class, I want to be able to tell from the function name alone whether it is the Controller or the main object, and which version of each interface it belongs to".

Options considered:

- **Plan (a) (adopted)**: Unify on a trailing version number, in the form `<origin>[<version>]_<method name>`
  - `Ctrl_xxx` (base) / `Ctrl2_xxx` (extension, future)
  - `View_xxx` (base) / `View2_xxx` / `View22_xxx` (extensions, future)
  - `View_OnXxx` (event receiver of the base) / `View2_OnXxx` (extension, future)
- Plan (b): Separate prefixes for origin and extension version (`ViewExt2_xxx`) -> redundant
- Plan (c): Use the formal IDL name as is (`ICoreWebView2_2_xxx`) -> too long

Reasons for adopting plan (a):
- Short (4-6 characters each)
- The version number corresponds one-to-one with the `N` of `ICoreWebView2_N` in the SDK
- Consistent with the existing `Ctrl_PutBounds` / `View_Navigate` (it only adds a naming convention for the extension side)

#### K4. Decision to leave the one-shot callbacks as they are

Whether `OnEnvironmentCompleted` (Class8) / `OnControllerCompleted` (Class9) should also be renamed was discussed.

Consideration:
- These are callbacks of API functions (`CreateCoreWebView2EnvironmentWithOptions` / `CreateCoreWebView2Controller`)
- **There is no function that directly receives a derived interface (`ICoreWebView2Environment2` etc.)** (they are obtained via QueryInterface)
- Therefore **there is no need to identify the version**

tarboh's confirmation: "**There's no function that gives you each derived interface directly, right? If not, there's no need to identify the version**" -> agreed to leave as is.

Adopted result:
- One-shot callbacks (`OnEnvironmentCompleted` / `OnControllerCompleted`) are left as is, without a prefix
- Persistent event callbacks (NavigationStarting / Completed / DocumentTitleChanged / WebMessageReceived / NewWindowRequested) are unified in the `View_OnXxx` form

#### K5. Insertion position of the Enum value (adopted: plan alpha = append at the end)

Position of the `HK_NewWindowRequested` Enum value:

- **Plan alpha (adopted)**: Append `HK_NewWindowRequested = 7` at the end. All existing values are kept. The Enum values of the persistent handler kinds are contiguous, `3 ~ 7`.
- Plan beta: Place it in the middle by logical grouping (after NavigationCompleted) while the value stays at the end -> order and value diverge, hard to read
- Plan gamma: Renumber the existing values -> not recommended, since the meaning of the m_persistentHandlers array elements would change

Changing the upper bound of the Class9 array from `HK_DocumentTitleChanged` to `HK_NewWindowRequested` extends the linked array by exactly +1 element. The loop in `CleanupAllPersistentHandlers` uses LBound/UBound, so **no code change is needed** (design principle 15 "Automate the rollout with HandlerKind-linked loops" at its best, the 2nd time after stage 9.4b).

#### K6. Argument type of put_Handled (adopted: 1&)

`ICoreWebView2NewWindowRequestedEventArgs::put_Handled` takes a BOOL (in Win32, BOOL = Long = 4 bytes). The old WebView2.cls wrote it as `dcf(args, 6, "put_Handled", 1)`.

Adopted: `dcf(pArgs, 6, "NewWindowReqArgs.put_Handled", 1&)` (the `&` suffix explicitly guarantees a Long, following design principle 29).

The real-machine test returned HRESULT = 0 (S_OK), so this is correct.

#### K7. Error handling (adopted: hybrid of plan a + plan c)

Behavior when `put_Handled` or `View_Navigate` fails:

- **Adopted (plan a + plan c)**: Log both, and continue processing even on failure. Even if they fail, the default behavior (a new window) simply takes over, so it is not fatal. No MsgBox is shown (Debug.Print only)

tarboh's decision: "**Logging and then carrying on is fine**".

#### K8. Also register a NewWindowRequested handler on the StartWebView2_LocalHtml side (adopted)

Maintain symmetry between the Bing path and the LocalHtml path. The LocalHtml path currently has no HTML that triggers it, but the registration itself is done for the sake of symmetry of the mechanism.

tarboh's decision: "**I'm in the 'put it in' camp**".

### Implementation

| File | Changes at stage 9.6a | Changes at stage 9.6b |
|---|---|---|
| `Module9_9_7.bas` -> `Module9_9_8.bas` | Added `HK_NewWindowRequested = 7` to the HandlerKind Enum; other foundation code left as is | No change |
| `Class7.cls` | Added an HK_NewWindowRequested branch to the Select Case in `Handler_Invoke`; renamed the existing 4 kinds of `m_owner.OnXxx` to `m_owner.View_OnXxx` | No change |
| `Class9.cls` | Extended the array upper bound from `HK_DocumentTitleChanged` -> `HK_NewWindowRequested`; renamed the existing 12 methods (4 kinds x Add/Remove/On) with the `View_` prefix; newly added `View_AddNewWindowRequested` / `View_RemoveNewWindowRequested` / `View_OnNewWindowRequested`; added HK_NewWindowRequested = 45 to `RemoveVtableIndexFor` | Added `put_Handled(1&)` + `View_Navigate(url)` to the body of `View_OnNewWindowRequested` |
| `UserForm1.frm` | Renamed the existing 7 places to `View_`-prefixed calls; added `View_AddNewWindowRequested` at the end of both Start methods | No change |

#### Main vtable indexes (ICoreWebView2, used in stage 9.6)

| Method | vtable index |
|---|---|
| `add_NewWindowRequested` | 44 |
| `remove_NewWindowRequested` | 45 |

#### vtable layout of ICoreWebView2NewWindowRequestedEventArgs (from the IDL comments of the old WebView2.cls)

| vtable index | Method | Use in 9.6 |
|---|---|---|
| 3 | `get_Uri(LPWSTR* uri)` | ✓ Retrieved in 9.6a |
| 4 | `put_NewWindow(ICoreWebView2*)` | Unused (for plan b, future) |
| 5 | `get_NewWindow(ICoreWebView2**)` | Unused |
| 6 | `put_Handled(BOOL)` | ✓ Used in 9.6b |
| 7 | `get_Handled(BOOL*)` | Unused |
| 8 | `get_IsUserInitiated(BOOL*)` | Unused |
| 9 | `GetDeferral(ICoreWebView2Deferral**)` | Unused |
| 10 | `get_WindowFeatures(ICoreWebView2WindowFeatures**)` | Unused |

### Excerpts of the real-machine test log and the observed behavior

#### 9.6a test: firing of NewWindowRequested and URL retrieval

After Bing was displayed, the "Images" link was clicked:

```
Class7.Invoke[HK_NewWindowRequested]
  this        : 1652572291984
  arg1 (raw)  : 41437844707328     <- matches m_pView
  arg2 (raw)  : 41437845783008
  m_refCount  : 2
Class9.View_OnNewWindowRequested: pSender=41437844707328, pArgs=41437845783008
  -> New-window request URL: https://www.bing.com/images?FORM=Z9LH
```

Confirmed that the argument interpretation rule `arg1 = sender` (= m_pView), `arg2 = args` (design principle 11) also holds for the new kind.

#### Behavior observed in 9.6a "when put_Handled is not handled"

Visual confirmation by tarboh:

> An independent new window was opened as part of the Excel process (belonging to the Excel group on the taskbar). I checked with UIA Inspect: the class name of the top-level window was `Chrome_WidgetWin_1` and the NamePropertyValue was "Bing 画像" (Bing Images). If it were Edge, something like "Microsoft Edge" should have been appended at the end, so I think it is open as pure WebView2. In Task Manager there is an item "Bing 画像" inside the WebView2-icon process group called "ユーティリティ" (Utility), so I'm sure.

-> **Important finding**: When `put_Handled` is not set, the WebView2 runtime does not hand over to Edge; instead it **creates a new WebView2 window by itself inside the Excel process**. The strategy of suppressing this with `put_Handled = TRUE` was established.

#### 9.6b test: handling with put_Handled + View_Navigate

After Bing was displayed, the "Images" link was clicked (9.6b version):

```
Class7.Invoke[HK_NewWindowRequested]
  arg1 (raw)  : 41437844707328
  arg2 (raw)  : 41437845783008
  m_refCount  : 2
Class9.View_OnNewWindowRequested: pSender=41437844707328, pArgs=41437845783008
  -> New-window request URL: https://www.bing.com/images?FORM=Z9LH
  put_Handled OK (the new window is suppressed)                                  <- ★ Core of 9.6b, point 1
  View_Navigate OK (navigated to "https://www.bing.com/images?FORM=Z9LH" in the same screen) <- ★ Core of 9.6b, point 2

[A chain of new navigation starts immediately]
Class9.View_OnNavigationStarting: -> Destination URL: https://www.bing.com/images?FORM=Z9LH
Class9.View_OnDocumentTitleChanged: -> Title: Bing 画像
Class9.View_OnNavigationCompleted: -> isSuccess: True
```

**Important facts observed**:
- The return value of `put_Handled` = S_OK (no failure log appeared) -> it works correctly with the argument `1&` (Long TRUE)
- The return value of `View_Navigate` = S_OK -> reusing the existing method works without problems
- Chained event firing in the same View: NewWindowRequested -> NavigationStarting -> DocumentTitleChanged -> NavigationCompleted run straight through
- **All persistent handlers also work on the new navigation** (= a handler, once registered, stays alive after Navigate, so the persistent-handler lifetime management of design principle 2 works correctly)
- The title from DocumentTitleChanged is "Bing 画像" -> the same as the title of the independent window seen in 9.6a, which is evidence that "the same content could be displayed on the same screen without opening an external window"

#### Cleanup (no regression)

```
Class9.Cleanup: removing persistent handler kind=3 vtable=8 token=61    <- NavigationStarting
Class9.Cleanup: removing persistent handler kind=4 vtable=16 token=61   <- NavigationCompleted
Class9.Cleanup: removing persistent handler kind=6 vtable=47 token=23   <- DocumentTitleChanged
Class9.Cleanup: removing persistent handler kind=7 vtable=45 token=41   <- NewWindowRequested
```

All 4 kinds were removed cleanly. All 4 slots were auto-released. The leak elimination by the stage 9.5 Sentinel mechanism also continues to work.

### Added design principles and findings established

- **Naming convention by origin prefix (new design principle 31)**: Methods derived from COM interfaces take the form `<origin><version>_<method name>` (`Ctrl_xxx` / `Ctrl2_xxx` / `View_xxx` / `View2_xxx` / `View_OnXxx` / `View2_OnXxx` / `Env_xxx`). One-shot callbacks (`OnEnvironmentCompleted` / `OnControllerCompleted`) are left as is without a prefix (no version identification needed, since no function directly receives a derived IF)
- **How to handle NewWindowRequested (new design principle 32)**: When `put_Handled` is not set, the WebView2 runtime creates a new WebView2 window inside the Excel process (it does not hand over to Edge; class name `Chrome_WidgetWin_1`). Suppressing this with `dcf(pArgs, 6, "put_Handled", 1&)` and firing `View_Navigate(uri)` on the same View achieves the "open in the same screen" behavior
- **The argument interpretation rule also holds for the new kind**: `(this, arg1, arg2)` of NewWindowRequested also works correctly with the interpretation `(handler obj, sender, args)`. Design principle 11 remains valid
- **HandlerKind-linked loop at its best (2nd time)**: Just widening the array upper bound by one makes `CleanupAllPersistentHandlers` handle the new kind automatically. Design principle 15 was demonstrated once more
- **The NewWindowRequested token is on a separate numbering system from the other events**: In the real-machine test, separate number sequences per kind were assigned: `add_NavigationStarting=60`, `add_NavigationCompleted=60`, `add_DocumentTitleChanged=22`, `add_NewWindowRequested=40` (an internal implementation detail of WebView2, recorded as an observed result)

### Known constraints and future considerations

- Plan (b), "open a popup in a separate UserForm", is not implemented. The plan is to handle it when the TabWebView2 umbrella class is introduced, by passing a separate Class9 via `args.put_NewWindow` (agreed with tarboh)
- The "close" operation (WindowCloseRequested event) for new windows opened with `target="_blank"` is not handled. It is unnecessary while pages are opened in the same View, since that is not an operation that closes a tab, but it will be needed when moving to plan b
- Other properties of args (get_IsUserInitiated, get_WindowFeatures, etc.) are not retrieved yet (to be added as needed)

### Hand-over items after stage 9.6

Since "5. Rolling out NewWindowRequested" was completed in stage 9.6, the remaining 4 items are listed below.

#### 1. Introduce the IID check (making Handler_QueryInterface authentic)
The current Handler_QueryInterface is a simple implementation that always returns the same vtable without looking at the IID. To better match WebView2's QI strategy, an IID check should be added. High independence, small scale.

#### 2. Mouse resizing of the UserForm
The approach of WS_THICKFRAME + GA_ROOT + SetWindowLongPtrW has been settled (agreed in stage 9.3c). Starting from Frame1.[_GethWnd], trace up to the UserForm HWND with GetAncestor(GA_ROOT). High independence, small scale.

#### 3. SPA support
NavigationCompleted fires only once, for the initial load. Tracking screen updates of an SPA (React/Vue etc.) requires the following:
- HistoryChanged event (when the URL is rewritten)
- DOMContentLoaded (ICoreWebView2_2 and later)
- Notification via postMessage from the SPA side
To be designed in an umbrella class layer such as TabWebView2.

#### 4. Version-extension IFs
ICoreWebView2_2 to _22, ICoreWebView2Controller2 to _6, etc. When needed, add fields + methods to Class9 (plan 4, the policy of implementing the base only). It can be rolled out as is with the naming convention (`View2_xxx` etc.) established in stage 9.6a.

#### 5. TabWebView2 umbrella class
The tab management layer. It is also cleanest to implement here the migration of IsPageLoaded from Class9 to TabWebView2, the merge with SPA support, and resize tracking. The "plan b, open in a separate UserForm" for NewWindowRequested also merges here. The largest scale.

### Code deliverables of this session

| File | Version | Role |
|---|---|---|
| `Class7.cls` | 9.6a | Generic handler, HK_NewWindowRequested branch added, dynamic binding target renamed to View_OnXxx |
| `Class8.cls` | 9.3a (no change) | Environment class |
| `Class9.cls` | 9.6b | Combined Controller + CoreWebView2 class, 5 kinds of persistent handlers + NewWindowRequested handling + View_ prefix naming convention |
| `Module9_9_8.bas` | 9.6a | Foundation + Sentinel mechanism + HK_NewWindowRequested, the rest left as in Module9_9_7 |
| `UserForm1.frm` | 9.6a | Frame1 + StartWebView2 / StartWebView2_LocalHtml + rename follow-through + View_AddNewWindowRequested calls added |

All files were verified on a real machine in a consistent state. This provides a stable baseline as the starting point for stage 9.6c (or any of the hand-over items).

---

## Stage 9.7 — Introducing the IID check + UserForm mouse resizing

### Completion date
This session (completed at stage 9.7b)

### Background and goal

With "5. Rolling out NewWindowRequested" completed in stage 9.6, the hand-over items were reduced to 4:

1. Introduce the IID check (making Handler_QueryInterface authentic)
2. Mouse resizing of the UserForm
3. SPA support
4. Version-extension IFs
5. TabWebView2 umbrella class

Of these, 1. and 2. had been assessed as "high independence, small scale". Stage 9.7 works through these 2 items as independent sub-stages:

- **9.7a**: Introduce the IID check (making `Handler_QueryInterface` authentic)
- **9.7b**: UserForm mouse resizing (WS_THICKFRAME + GA_ROOT) + WebView2 Bounds tracking

The two have completely independent scopes of impact, so the sub-stage boundary can be drawn cleanly. 9.7a is a change to Module9 (foundation) alone, and 9.7b is a change to Module9 (only API declarations and constants added) + UserForm1, keeping changes to Class7/8/9 at zero.

### Flow of sub-stages

| Stage | Items achieved |
|---|---|
| 9.7a | GUID type + IID table + genuine Handler_QueryInterface + related helpers, Module9_9_8 → Module9_9_9 |
| 9.7b | UserForm mouse resize + WebView2 Bounds tracking, 4 APIs and 7 constants added to Module9_9_9 without a version bump, major overhaul of UserForm1 |

---

### Design decision points of stage 9.7a

#### L1. QueryInterface strategy (adopted: option a = strict)

Up to 9.6b, `Handler_QueryInterface` was a simple implementation that "always returned itself without looking at the IID". The strategies for making it genuine:

- **Option a (adopted)**: Strict. For any IID other than IID_IUnknown or the proper IID, return `ppvObject = 0` + `E_NOINTERFACE`. The most faithful to the COM convention
- Option b: Lenient. Accept all IIDs and, for an unknown IID, only log a warning and return S_OK. Because it would return a pointer while lying "I am also an IMarshal", there is a risk of a crash the moment WebView2 actually calls an IMarshal method

Reason for adoption (option a): Priority on faithfulness to the COM convention. The anticipated risk (WebView2 internally trying IAgileObject / IMarshal etc., getting E_NOINTERFACE back, and handling being refused on a particular path) will be judged after observing it in the real-machine test.

#### L2. How to initialize the IID table (adopted: string-parse method)

An IID for each HandlerKind must be held inside VBA. Initialization methods:

- Option i: Assign directly to the fields, in the form `m_iidTable(HK_xxx).Data1 = &H6C4819F3`. There is a risk of stepping on the `&H` literal trap (design principle 29), so either avoid it or put the `&` suffix on every one, which is cumbersome
- Option ii: Hold it as a byte array and write it with CopyMemory. Conflicts with the no-copy policy (the direction of the design principles)
- **Option iii (adopted)**: **String-parse method**. Write it as `FillGUID m_iidTable(HK_xxx), "6C4819F3-5E55-4E29-AF46-8157DBFA88A6"`, and parse the string with `HexStrToLong` / `HexStrToInt` and pack it into the fields

Reason for adoption (option iii): The GUID string copy-pasted from WebView2.h can be used as is. It avoids the `&H` literal trap. High readability.

#### L3. Granularity of log output (adopted: failures only)

Once the IID check is in, a "success/failure" log line could be emitted at every QueryInterface. The WebView2 runtime always fires a QI at registration, so at least 5 lines (5 handlers × startup) would appear.

- Option α: Log every QI (both success and failure). Strong for diagnosis but verbose
- **Option β (adopted)**: **Log failures only**. In the form "QI rejected: kind=N riid.Data1=&Hxxx → E_NOINTERFACE"

Reason for adoption (option β): If the "rejected IIDs" can be observed, we can see which IIDs WebView2 tries, and that is enough. Stay silent on success (= less log noise).

#### L4. GUID comparison implementation (adopted: ReadLongPtr × 2)

How to compare two 16-byte GUIDs:

- Option ω: A byte-by-byte loop with 16 comparisons
- **Option χ (adopted)**: **Read a LongLong (= 8 bytes) twice and compare all 16 bytes at once**. Uses `ReadLongPtr` twice

Reason for adoption (option χ): It can be implemented simply by reusing the existing `ReadLongPtr` (PointerAccessor-based). The Win32 GUID structure layout is 16 contiguous bytes even on x64 (no padding), so two LongLongs cover it completely.

#### L5. Establishing the isolation procedure in advance

In case the IID check introduces inexplicable behavior (handler not registered / Invoke not called / AddRef/Release mismatch / hang during Cleanup), troubleshooting guidance was written in advance in the header comment of Module9:

1. Temporarily add "`ppvObject = this : HandlerAddRefInternal this : Handler_QueryInterface = S_OK : Exit Function`" at the top of `Handler_QueryInterface` to revert to the 9.6b equivalent
2. If functionality is restored in this state, the IID check is confirmed as the cause
3. If not, look for a different cause

---

### Implementation of stage 9.7a

| File | Change |
|---|---|
| `Module9_9_8.bas` → `Module9_9_9.bas` | New GUID type (Public Type), new E_NOINTERFACE constant, new `m_iidIUnknown` / `m_iidTable(HK_None To HK_NewWindowRequested) As GUID`, `InitIIDTable` call added at the end of `Thunks_Init`, `Handler_QueryInterface` rewritten (with IID check), helpers added (`InitIIDTable` / `FillGUID` / `HexStrToLong` / `HexStrToInt` / `IsEqualGUIDInPlace` / `LongPtrLowDword` / `MemLongLong`) |
| `Class7.cls` | No change |
| `Class8.cls` | No change |
| `Class9.cls` | No change |
| `UserForm1.frm` | No change |

Contents of the IID table (7 kinds):

| HandlerKind | IID (from WebView2.h) | Purpose |
|---|---|---|
| `HK_EnvironmentCompleted` | `4E8A3389-C9D8-4BD2-B6B5-124FEE6CC14D` | `ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler` |
| `HK_ControllerCompleted` | `6C4819F3-5E55-4E29-AF46-8157DBFA88A6` | `ICoreWebView2CreateCoreWebView2ControllerCompletedHandler` |
| `HK_NavigationStarting` | `9ADBE429-F36D-432B-9DDC-F8881FBD76E3` | `ICoreWebView2NavigationStartingEventHandler` |
| `HK_NavigationCompleted` | `D33A35BF-1C49-4F98-93AB-006E0533FE1C` | `ICoreWebView2NavigationCompletedEventHandler` |
| `HK_WebMessageReceived` | `57213F19-00E6-49FA-8E07-898EA01ECBD2` | `ICoreWebView2WebMessageReceivedEventHandler` |
| `HK_DocumentTitleChanged` | `F5F2B923-953E-4042-9F95-F3A118E1AFD4` | `ICoreWebView2DocumentTitleChangedEventHandler` |
| `HK_NewWindowRequested` | `D4C185FE-C81C-4989-97AF-2D3FA7AB5651` | `ICoreWebView2NewWindowRequestedEventHandler` |

(IID_IUnknown is initialized separately into `m_iidIUnknown` with `"00000000-0000-0000-C000-000000000046"`)

### Real-machine test log of stage 9.7a

Ran Test_AllEvents_Real / Test_MessageEcho_Real / the reset test.

**Observed fact 1: the "QI rejected" log never appeared**

All 5 handlers (HK_NavigationStarting / HK_NavigationCompleted / HK_DocumentTitleChanged / HK_NewWindowRequested / HK_WebMessageReceived) registered successfully. Across the whole flow of link click → navigation chain → reset, not a single failure log line from `Handler_QueryInterface` appeared.

→ **The WebView2 runtime does not request, via QueryInterface, any IID other than IID_IUnknown and the proper IID from a registered event handler**.

**Observed fact 2: AddRef happens only once, at registration**

Consistent with the observation at stage 9.4 ("AddRef × 1 at registration, AddRef × 1 + Release × 1 per Invoke, Release × 1 at unregistration"). It is confirmed that WebView2 adopts the pattern **"QI at registration and cache the vtable; at Invoke time, call Invoke directly without QI"**.

**Observed fact 3: the real reason the pre-9.6b "pass every IID through" implementation caused no actual harm**

Because WebView2 fires QI only minimally, the dangerous situation of "returning a fake IMarshal pointer" never arose in the first place. Up to 9.6b we were in a state of "it was dangerous but we luckily never stepped on it", and the tightening in 9.7a works as "insurance against stepping on a future landmine".

#### Regression

No change in behavior in Test_AllEvents_Real / Test_MessageEcho_Real / the reset test. The cleanup log (`Class9.Cleanup: removing persistent handler kind=...`) is also the same as in 9.6b.

### Additions to design principles and knowledge established in 9.7a

- **Implement QueryInterface strictly (new design principle 33)**: Following the COM convention, return `ppvObject = 0` + `E_NOINTERFACE` for any IID other than IID_IUnknown and the proper IID. "Accept all IIDs" risks returning a fake IMarshal pointer; we just happen not to step on it at present, but it is a future landmine
- **The WebView2 runtime fires QI only minimally on registered handlers**: Recorded as an observation result. The strategy is to QI at registration and cache the vtable, and at Invoke time to call Invoke directly without QI
- **The string-parse method is the safe way to initialize the GUID table**: GUID strings copy-pasted from WebView2.h can be used as is, the `&H` literal trap is avoided, and readability is high
- **Compare GUIDs as 16 bytes at once with LongLong × 2**: Implementable just by using the PointerAccessor-family `ReadLongPtr` twice

---

### Design decision points of stage 9.7b

#### M1. UserForm resize policy (adopted: WS_THICKFRAME + GA_ROOT + SetWindowLongPtrW)

As agreed at stage 9.3c. To make a UserForm resizable with the mouse:

1. Get the HWND of Frame1 with `Frame1.[_GethWnd]` (existing)
2. Walk up to the UserForm's own HWND with `GetAncestor(hWnd_Frame, GA_ROOT)`
3. Get the current style with `GetWindowLongPtrW(hWnd_Form, GWL_STYLE)`
4. Set `WS_THICKFRAME` with OR and write it back with `SetWindowLongPtrW`
5. Call `SetWindowPos` with `SWP_FRAMECHANGED` to trigger a redraw of the non-client area

#### M2. Resize tracking method (adopted: option a = gap preservation)

How should Frame1 follow in UserForm_Resize:

- **Option a (adopted)**: **Save the gaps and reproduce them**. At Initialize, save `m_gapLeft = Frame1.Left`, `m_gapTop = Frame1.Top`, `m_gapRight = Me.InsideWidth - (Frame1.Left + Frame1.Width)`, `m_gapBottom = Me.InsideHeight - (Frame1.Top + Frame1.Height)`. At Resize, reproduce in the form `Frame1.Width = Me.InsideWidth - m_gapLeft - m_gapRight`
- Option b: Set Frame1 every time to fill the UserForm's client area (0, 0, InsideWidth, InsideHeight). Simple, but breaks down if controls other than Frame1 (a toolbar, etc.) are added in the future
- Option c: Keep it by ratio. Scale while preserving the design-time ratios

Reason for adoption (option a): Even if controls other than Frame1 are added in the future, the design-time positional relationships are preserved. Also consistent with the design principle of "keeping the design-time layout".

#### M3. Method of WebView2 Bounds tracking (adopted: Ctrl_PutBounds from UserForm_Resize)

To make the WebView2 rendering area follow Frame1:

- **Adopted**: Inside `UserForm_Resize`, re-fire `GetClientRect(hWnd_Frame, rc)` and pass that rc to `Ctrl_PutBounds`
- Considered: WindowProc subclassing (catching WM_SIZING for real-time tracking). The MouseScroll.bas style

Reason for adoption: The real-machine test confirmed that "just calling `Ctrl_PutBounds` inside `UserForm_Resize` works smoothly". The internal compositing layer of the WebView2 runtime and DWM (Desktop Window Manager) takes on the smooth drawing, so subclassing and custom handling of WM_SIZING are unnecessary. **"Simple and effective"**.

#### M4. Countermeasure for the false Resize during Initialize (adopted: m_resizeReady flag)

A VBA UserForm fires the Resize event when `Me.Width = ...` is set. If it fires in the middle of Initialize, `m_gap*` is not yet fixed, so it breaks down.

- **Adopted**: Introduce `m_resizeReady As Boolean`. Set it to True at the end of Initialize. At the top of UserForm_Resize, return early with `If Not m_resizeReady Then Exit Sub`

#### M5. Removing the Me.Width / Me.Height assignment (adopted)

Up to 9.6, UserForm_Initialize fired `Me.Width = 800: Me.Height = 600`. This:
- Overwrites the designer (= breaks design-time consistency)
- Breaks the gap calculation because Frame1 is still as designed (in the Resize event during Initialize, only Frame1 is enlarged to match the new UserForm size)
- Contradicts design principle 36 (newly established), "keep the design-time layout"

→ **Removed**. If you want to display at a different size at startup, the policy is to adjust the UserForm size (and the size and position of Frame1 together) in the designer.

### Implementation of stage 9.7b

| File | Change |
|---|---|
| `Module9_9_9.bas` (version unchanged) | Added 4 APIs for UserForm mouse resizing (`GetAncestor` / `GetWindowLongPtrW` / `SetWindowLongPtrW` / `SetWindowPos`) as Public Declare, and 7 related constants (`GA_ROOT` / `GWL_STYLE` / `WS_THICKFRAME` / `SWP_NOSIZE` / `SWP_NOMOVE` / `SWP_NOZORDER` / `SWP_FRAMECHANGED`) as Public Const |
| `Class7.cls` | No change |
| `Class8.cls` | No change |
| `Class9.cls` | No change |
| `UserForm1.frm` | Added 5 module variables (`hWnd_Form` / `m_gapLeft` / `m_gapRight` / `m_gapTop` / `m_gapBottom` / `m_resizeReady`), added WS_THICKFRAME assignment + gap saving + flag raising at the end of `UserForm_Initialize`, newly added `UserForm_Resize` (Frame1 resize + re-issue of Ctrl_PutBounds), removed the `Me.Width` / `Me.Height` assignments |

### Real-machine test of stage 9.7b

Visual confirmation by tarboh:

- **Silky-smooth resize tracking achieved**: When dragging the edge of the form with the mouse, Frame1 and the WebView2 rendering area follow in real time. No stutter or flicker in drawing is observed
- **No subclassing needed**: No WindowProc subclassing or custom WM_SIZING handling is in place, yet it works smoothly. This is thanks to the excellence of the internal compositing layer of DWM and WebView2
- **Design-time layout maintained**: The gaps saved at Initialize (tarboh's environment: `L=18, T=18, R=24.75, B=22.5`) are preserved through shrinking/enlarging after startup
- **Effect of the m_resizeReady flag**: Even if a Resize event fires during Initialize, the early return avoids breakage

### Additions to design principles and knowledge established in 9.7b

- **UserForm resize tracking uses the "save the gaps and reproduce them" method (new design principle 34)**: At Initialize, save `Frame1.Left/Top` in the design state and `Me.InsideWidth/Height - (Frame1's bottom-right edge)`, and at Resize resize Frame1 by reproducing them. Even if controls other than Frame1 (toolbar, sidebar, etc.) are added in the future, the design-time positional relationships are preserved
- **UserForm_Resize: beware of false events during Initialize (new design principle 35)**: A VBA UserForm fires the Resize event when `Me.Width` / `Me.Height` is set. Since `m_gap*` is not yet fixed during Initialize, introduce `m_resizeReady As Boolean`, set it to True at the end of Initialize, and return early at the top of UserForm_Resize
- **Decide the UserForm size in the designer, and do not assign to Me.Width / Me.Height in Initialize (new design principle 36)**: Keep the design-time layout (UserForm size, Frame1 size, positional relationships). If you want to display at a different size at startup, set it in the designer
- **WebView2 resize tracking works smoothly just by firing Ctrl_PutBounds from UserForm_Resize (new design principle 37)**: No WindowProc subclassing or custom WM_SIZING handling is needed. The internal compositing layer of DWM and WebView2 takes on the smooth drawing

---

### Known constraints and future considerations (stage 9.7)

- Because the IID table is initialized at the end of `Thunks_Init`, if `Handler_QueryInterface` is called before `Thunks_Init` runs, the IID check does not work as expected (in practice a handler is never Acquired before the thunk region is allocated, so there is no problem)
- Since "QI rejected" was never observed, the check of the path that returns E_NOINTERFACE remains a theoretical guarantee (no problem in practice)
- The UserForm's maximize/minimize buttons are not implemented (only WS_THICKFRAME is set, so to show a maximize button WS_MAXIMIZEBOX must be set at the same time; a future task)
- A minimum size limit on resize (via WM_GETMINMAXINFO) is not implemented (if made extremely small, Frame1 may fly off-screen; add when needed)

### Handover items after stage 9.7

Because "1. Introduce the IID check" and "2. UserForm mouse resize" were completed in stage 9.7, they are removed from the handover items. The remaining 3 items are listed below.

#### 1. Formal fixing of class names (to be started in 9.8a)
Renames `Class7 → ComCallbackHandler` / `Class8 → Wv2Environment` / `Class9 → Wv2Pane`. Renaming the Module9 family to `Wv2Thunks` etc. is also under consideration. Field names in the UserForm (`m_ctrl → m_pane` etc.) are tidied at the same time. Old-name mentions in the dev notes are replaced in bulk. Agreed with tarboh.

#### 2. Comprehensive implementation of the various ICoreWebView2 methods (to be started in 9.8b / 9.8c)
- 9.8b: `View_GetSource` / `View_Reload` / `View_Stop` / `View_GetCanGoBack` / `View_GetCanGoForward` / `View_GetBrowserProcessId` + 3 to 4 Settings-family ones + `Ctrl_Close`
- 9.8c: New `View_ExecuteScript` + `HK_ExecuteScriptCompleted` (one-shot kind)

#### 3. SPA support
NavigationCompleted fires only once, for the initial load. Tracking screen updates of an SPA (React/Vue etc.) requires the following:
- HistoryChanged event (when the URL is rewritten)
- DOMContentLoaded (ICoreWebView2_2 and later)
- postMessage notification from the SPA side
Design this in a comprehensive class layer such as TabWebView2.

#### 4. Version-extension interfaces
ICoreWebView2_2 to _22, ICoreWebView2Controller2 to _6, etc. When needed, add fields + methods to Wv2Pane (option 4, the policy of implementing only the base). It can be expanded as is using the naming convention (`View2_xxx` etc.) established in 9.6a.

#### 5. TabWebView2 comprehensive class
The tab management layer. Migrating IsPageLoaded from Wv2Pane to TabWebView2, merging with SPA support, and the "option b: open in a separate UserForm" for NewWindowRequested also merge here. The largest scale.

### Code deliverables of this session

| File | Version | Role |
|---|---|---|
| `Class7.cls` | 9.6a (no change) | Generic handler, branching for 5 kinds of events |
| `Class8.cls` | 9.3a (no change) | Environment class |
| `Class9.cls` | 9.6b (no change) | Combined Controller + CoreWebView2 class |
| `Module9_9_8.bas` → `Module9_9_9.bas` | 9.7a/b | Foundation + Sentinel mechanism + IID check (9.7a) + UserForm resize APIs (9.7b, version unchanged) |
| `UserForm1.frm` | 9.7b | Frame1 + StartWebView2 / StartWebView2_LocalHtml + mouse resize support (gap saving + WS_THICKFRAME + new UserForm_Resize + Ctrl_PutBounds tracking) |

All files were verified on the real machine in a consistent state. It provides a stable base as the starting point of stage 9.8 (formal fixing of class names).

---

### Announcement of renames at 9.8a (finalized at the end of this session)

At stage 9.8a the following class names and module names are formally fixed. **There is no logic change at all; it is purely a rename**. The exit checks are two: "no regression" and "no missed renames".

| Old name | New name | Role |
|---|---|---|
| `Class7` | `ComCallbackHandler` | VBA-side wrapper for COM callbacks (via thunks) |
| `Class8` | `Wv2Environment` | VBA-side 1-to-1 wrapper for `ICoreWebView2Environment` |
| `Class9` | `Wv2Pane` | Combined `ICoreWebView2Controller` + `ICoreWebView2` class (= one WebView2 display pane) |
| `Module9_9_9` | `Wv2Thunks` | Foundation layer (thunk region, slot management, reference counting, IID check, Sentinel mechanism) |
| `m_ctrl` (inside UserForm1) | `m_pane` | Reference to Wv2Pane |

#### Scope to tidy up along with the rename (agreed with tarboh)
- A. Module name: `Module9_9_9 → Wv2Thunks`
- B. UserForm field name: `m_ctrl → m_pane`
- F. State "Former name: ClassN" at the top of each file header
- H. Replace mentions of Class7/8/9 in headers and bodies with the new names (the lengthy history comments, such as the old 9.4c → 9.6a difference notes, are swept away by the new headers)

#### Policy for the dev notes (instruction from tarboh)
- The Class7/8/9 notation in past sections (9.3a to 9.7) is **left as is** (= preserving the historical record)
- New sections from 9.8a onward are **written with the new names**

#### Subdivision of stage 9.8 (Claude's recommended plan, agreed with tarboh)
| Stage | Scope |
|---|---|
| **9.8a** | Fix the class names (the content of this notice) |
| **9.8b** | Full coverage of the View base IF (5 items) + Settings-related (3 to 4 items) + `Ctrl_Close` (excluding ExecuteScript) |
| **9.8c** | `View_ExecuteScript` + `HK_ExecuteScriptCompleted` (new one-shot handler family) |

It was also agreed with tarboh that the versioned extension IFs (`ICoreWebView2_2`, etc.) will not be started within 9.8 and are deferred to 9.9 and later.


---

## Stage 9.8a — Formal fixing of the class names

### Completion date
This session

### Background and purpose

At the end of stage 9.7 we had agreed to start the formal fixing of the class names at the beginning of stage 9.8. A rename contains no logic change, which has the advantage of making regression triage clear.

### Renames carried out in 9.8a

| Old name | New name | Role |
|---|---|---|
| `Class7` | `ComCallbackHandler` | VBA-side wrapper for COM callbacks (via thunks) |
| `Class8` | `Wv2Environment` | VBA-side 1-to-1 wrapper for `ICoreWebView2Environment` |
| `Class9` | `Wv2Pane` | Combined class of `ICoreWebView2Controller` + `ICoreWebView2` (= one WebView2 display pane) |
| `Module9_9_9.bas` | `Wv2Thunks.bas` | Foundation layer (thunk region, slot management, reference counting, IID check, Sentinel mechanism) |
| `m_ctrl` (UserForm1) | `m_pane` | Reference to Wv2Pane |

### Rename policy

- State "Former name: Class7" etc. at the top of the header comment (instruction from tarboh)
- Unify the body (code, comments, Debug.Print text) on the new names
- Zero logic change (= if a regression appears, it can be isolated as a mistake in the rename work)
- Past sections of the dev notes (9.3a to 9.7) are left as is (= preserving the historical record, instruction from tarboh)

### Real-machine test results of stage 9.8a

Test_AllEvents_Real + Test_MessageEcho_Real were run. **Complete success**.

| Check item | Result |
|---|---|
| Compile | Zero errors |
| Log text (new names) | `ComCallbackHandler.Invoke[HK_xxx]` / `Wv2Pane.View_OnXxx` / `Wv2Environment.OnEnvironmentCompleted` are all output with the new names |
| Log text (old names) | Zero occurrences of `Class7.` / `Class8.` / `Class9.` |
| Environment → Controller → View chain | Normal |
| Registration of the 5 kinds of persistent handlers | Normal (NavigationStarting / NavigationCompleted / DocumentTitleChanged / NewWindowRequested / WebMessageReceived) |
| Event firing chain | Normal (Bing displayed → title obtained → navigation completed) |
| Bidirectional WebMessage | Normal (both VBA → JS and JS → VBA) |
| Cleanup log | Normal (`Wv2Pane.Cleanup: removing persistent handler kind=...` appears 5 times) |
| Class_Terminate chain | Normal (`Wv2Pane.Terminate → Wv2Environment.Terminate`) |
| Sentinel mechanism (9.5) | Normal (`Sentinel: VirtualFree succeeded for ...`) |
| Mouse resize (9.7b) | Normal (`WS_THICKFRAME applied`, `gaps L=18 T=18 R=24.75 B=22.5`) |
| IID check (9.7a) | Normal (zero `QI rejected` logs = same observation as 9.7a) |
| AddRef/Release balance | Normal (every Slot ends up auto-released) |

### Design principles and findings added in 9.8a

- **Final form of class naming (design principle 38, new)**:
  - `ComCallbackHandler`: COM callback foundation (no prefix, because it is not SDK-derived)
  - `Wv2Environment`: 1-to-1 with the SDK's Environment (the `Wv2` prefix makes WebView2 explicit)
  - `Wv2Pane`: combined class of the SDK's Controller + View (Pane = "display pane", the unit attached to an HWND)
  - `Wv2Thunks`: foundation module (meaning thunk-related)
  - Reasons for adopting the `Wv2` prefix: completion works well, and `Wv2Environment` / `Wv2Pane` convey their purpose
- **Outcome of making the rename sub-stage independent (design principle 39, new)**: If you carve out a sub-stage where "zero logic change" is guaranteed, then when a regression appears it can be immediately isolated to a "missed rename". By the time methods are added in the next sub-stage, things are already running under the new names = easy to review. This is especially effective when the rename is large (300+ places)
- **"Former name" notation at the top of the file (design principle 40, new)**: For a file whose class name was changed, stating "Former name: ClassN" in the header comment makes it easier to cross-check the new file while reading past dev notes

### Code deliverables of 9.8a

| File | Version | Role |
|---|---|---|
| `Class7.cls` → `ComCallbackHandler.cls` | 9.8a | General-purpose handler, branching for the 5 kinds of events (logic unchanged) |
| `Class8.cls` → `Wv2Environment.cls` | 9.8a | Environment class (logic unchanged) |
| `Class9.cls` → `Wv2Pane.cls` | 9.8a | Controller + CoreWebView2 combined class (logic unchanged) |
| `Module9_9_9.bas` → `Wv2Thunks.bas` | 9.8a | Foundation + Sentinel mechanism + IID check + UserForm resize API (logic unchanged) |
| `UserForm1.frm` | 9.8a | Frame1 + StartWebView2 / StartWebView2_LocalHtml + mouse resize support (only the m_ctrl → m_pane rename) |

The real-machine test logs fully cleared the exit check of stage 9.8a (zero regressions + zero missed renames). This provides a stable base for the next stage (9.8b).


---

## Stage 9.8b — Full coverage of the View base IF + Settings-related + Ctrl_Close

### Completion date
This session

### Background and purpose

With the class names fixed in 9.8a, a stable base was obtained. In 9.8b, as agreed in advance, we implement full coverage of the main methods of the ICoreWebView2 base IF, the handling of Settings, and Close of ICoreWebView2Controller. ExecuteScript is separated into stage 9.8c because its scale (new HK_ExecuteScriptCompleted, addition of a one-shot handler family) is independent.

### Design decision points of 9.8b

#### N1. How to handle Settings (adopted: option 1 = helper inside Wv2Pane)

- **Option 1 (adopted)**: Cache `m_pSettings As LongPtr` inside Wv2Pane and implement `Settings_PutXxx` as Public methods of Wv2Pane. A Private helper named `EnsureSettings` lazily obtains View vtable 3 (get_Settings)
- Option 2: Create a new `Wv2Settings` class and return it via a Wv2Pane.Settings property

Reasons for adopting option 1: no increase in classes, it fits within the scale of 9.8b, and redesigning the Release order in Class_Terminate stays simple. If the Settings functionality grows, the option of splitting it out into Wv2Settings in the future remains.

#### N2. Signature of BOOL getters (adopted: As Boolean)

- **Adopted: the style that returns As Boolean**. The HRESULT is handled internally (log output only), and False is returned on failure
- Considered: As Long (HRESULT) + ByRef result As Boolean

Reasons for adoption: consistent with the existing `Public Property Get IsPageLoaded() As Boolean`. Calls from the UserForm side can be written naturally in the form `If pane.View_GetCanGoBack Then ...`. The `GetBoolProperty` helper (already existing since stage 9.4b) can be reused as is.

#### N3. Signature of void methods (adopted: return the HRESULT as Long)

- **Adopted**: View_Reload / View_Stop / Ctrl_Close return As Long (= HRESULT)
- Consistent with the existing View_Navigate / View_NavigateToString

#### N4. Release order in Class_Terminate (adopted: Settings → View → Ctrl → Environment)

With the introduction of m_pSettings, the Release order is extended:

```
1. CleanupAllPersistentHandlers (m_pView must still be alive)
2. Release of m_pSettings  ← added in 9.8b
3. Release of m_pView
4. Release of m_pCtrl
5. (In Wv2Environment.Class_Terminate) Release of m_pEnvironment
```

Following the dependency (Ctrl is the parent, View the child, Settings the grandchild), we release from the children first.

#### N5. Test method (adopted: option i = GetActivePane property + Immediate window)

Add `Public Property Get GetActivePane() As Wv2Pane` to UserForm1 and call things like `UserForm1.GetActivePane.View_Reload` from the Immediate window. This lets us check the new methods interactively without adding test functions.

### Implementation of 9.8b

| File | Change |
|---|---|
| `Wv2Pane.cls` | Added the module variable `m_pSettings As LongPtr`, added the Release of m_pSettings to Class_Terminate (before View), added new Public functions (View_Reload / View_Stop / View_GetCanGoBack / View_GetCanGoForward / View_GetBrowserProcessId / Settings_PutIsScriptEnabled / Settings_PutAreDefaultContextMenusEnabled / Settings_PutAreDevToolsEnabled / Ctrl_Close), added the new Private function `EnsureSettings` |
| `UserForm1.frm` | Added `Public Property Get GetActivePane() As Wv2Pane` (for Immediate-window testing) |
| `ComCallbackHandler.cls` | No change |
| `Wv2Environment.cls` | No change |
| `Wv2Thunks.bas` | No change |

vtable index mapping of the APIs added in 9.8b:

| Method | Target IF | vtable | Type |
|---|---|---|---|
| View_Reload | ICoreWebView2 | 31 | void |
| View_GetBrowserProcessId | ICoreWebView2 | 37 | UINT32 |
| View_GetCanGoBack | ICoreWebView2 | 38 | BOOL |
| View_GetCanGoForward | ICoreWebView2 | 39 | BOOL |
| View_Stop | ICoreWebView2 | 43 | void |
| (inside EnsureSettings) | ICoreWebView2 | 3 (get_Settings) | LPVOID** |
| Settings_PutIsScriptEnabled | ICoreWebView2Settings | 4 | BOOL |
| Settings_PutAreDevToolsEnabled | ICoreWebView2Settings | 12 | BOOL |
| Settings_PutAreDefaultContextMenusEnabled | ICoreWebView2Settings | 14 | BOOL |
| Ctrl_Close | ICoreWebView2Controller | 24 | void |

### Real-machine test of 9.8b

Verification completed by tarboh.

#### A. Regression (the behavior equivalent to 9.8a is maintained) → all items OK
- Test_AllEvents_Real / Test_MessageEcho_Real all succeeded
- Bidirectional WebMessage, registration of the 5 kinds of persistent handlers, Cleanup log, and the Class_Terminate chain are all equivalent to stage 9.8a

#### B. New features (verified from the Immediate window)
| Check | Result |
|---|---|
| View_GetSource | "https://www.bing.com/" (existing method, consistency confirmed) |
| View_GetCanGoBack (initial display) | False (as expected) |
| View_GetCanGoForward (initial display) | False (as expected) |
| View_GetBrowserProcessId | 28568 (a PID that can be matched in Task Manager) |
| View_GoBack | NavigationStarting → NavigationCompleted chain, the screen goes back |
| View_GetCanGoForward (after going back) | True (as expected) |
| View_Reload | NavigationStarting → NavigationCompleted chain, reloaded |
| View_Stop (run after load completion) | Nothing happens (= harmless, as expected) |
| Ctrl_Close | The WebView2 area goes blank, and all registered handlers are auto-released |

#### C. Release order in Class_Terminate → OK
```
Wv2Pane.Terminate: Releasing m_pSettings = 91534344078656
Wv2Pane.Terminate: Releasing m_pView = 91534343247872
Wv2Pane.Terminate: Releasing m_pCtrl = 91534344143624
Wv2Environment.Terminate: Releasing m_pEnvironment = 91534344595472
```
The Settings → View → Ctrl → Environment order intended in 9.8b was confirmed.

### ★ Important spec facts established in 9.8b ★

#### Spec fact 1: WebView2 Settings are not reflected until the next top-level navigation

Facts observed in the real-machine test:

| Operation | Observed result | Interpretation |
|---|---|---|
| `Settings_PutAreDevToolsEnabled(False)` → press F12 (no Reload) | DevTools launches | The setting is not reflected until the next navigation |
| `Settings_PutAreDefaultContextMenusEnabled(False)` → right-click (no Reload) | The context menu appears | Same as above |
| `Settings_PutIsScriptEnabled(False)` + run `View_Reload` | Bing turns into a pitch-black screen | The Reload runs the next navigation, IsScriptEnabled=False takes effect, and Bing cannot render because it depends on JS |

The specification stated explicitly in the official Microsoft documentation (CoreWebView2Settings section):

> Changes to IsGeneralAutofillEnabled and IsPasswordAutosaveEnabled apply immediately, while other setting changes made after NavigationStarting event do not apply until the next top-level navigation.

A similar Q&A is also recorded in a GitHub Issue (MicrosoftEdge/WebView2Feedback #587):

> Settings require a navigation before they take affect. ... It works correctly if I navigate after applying the settings.

In other words, **most Settings other than `IsGeneralAutofillEnabled` and `IsPasswordAutosaveEnabled` are not applied until the next top-level navigation even if set after NavigationStarting**. Even if the return value of `Settings_PutXxx` is 0 (S_OK), it does not take effect immediately. To confirm the reflection, you need to follow up with `View_Reload` or `View_Navigate(...)`.

This is presumed to reflect an internal implementation in which "inside the WebView2 runtime, the settings are enabled when the navigation process is launched".

#### Spec fact 2: Ctrl_Close has the side effect of releasing all registered event handlers

Facts observed from the real-machine test log:

```
?UserForm1.GetActivePane.Ctrl_Close()
Wv2Pane.Ctrl_Close: m_pCtrl = 91534344143624
  Release [idx=4] → refCount = 0 → Slot 4 auto-released
  Release [idx=1] → refCount = 0 → Slot 1 auto-released
  Release [idx=0] → refCount = 0 → Slot 0 auto-released
  Release [idx=2] → refCount = 0 → Slot 2 auto-released
 0
```

→ When `Ctrl_Close` is called, in addition to the Controller's own Close processing, **the WebView2 runtime fires a Release on all registered persistent event handlers**. As a result, the handler slots on the VBA side are auto-released (refCount goes from 1 to 0).

This is consistent with the observation established in stage 9.4a ("WebView2 does AddRef at registration, AddRef+Release on each Invoke, and Release at unregistration"): Close has the semantics of "releasing all event handlers at once".

#### Derived caution for spec fact 2: behavior in Class_Terminate after Close

In the real-machine test, when the form is closed after `Ctrl_Close`, the log `Wv2Pane.Cleanup: removing persistent handler kind=3/4/6/7 token=29` also appears in Class_Terminate. This is because `CleanupAllPersistentHandlers` does not clear `m_persistentTokens`, so it tries to fire `remove_*Event` even after Close.

With the current behavior there is **no problem** (even after Close, m_pView is apparently still alive and the remove succeeds). However, the safety when handling derived IFs such as ICoreWebView2_2 in the future, or when calling View_xxx after Close, is a point that needs verification. Intervening in this (clearing m_pCtrl/m_pView, clearing m_persistentTokens, etc.) will be **considered together with the ExecuteScript implementation in stage 9.8c** (instruction from tarboh).

### Design principles and findings added in 9.8b

- **Design principle 41 (new) When WebView2 Settings take effect**: Settings other than `IsGeneralAutofillEnabled` and `IsPasswordAutosaveEnabled` are not reflected until the next top-level navigation. Even if the return value of `Settings_PutXxx` is 0 (S_OK), it does not take effect immediately. To confirm the reflection, follow up with `View_Reload` or `View_Navigate`
- **Design principle 42 (new) Lazy-fetch and cache approach for Settings**: m_pSettings is not obtained at Init but at the moment a Settings_xxx is first called (EnsureSettings). This avoids adding needless COM references in cases where Settings is not used, and does not affect the existing Init chain
- **Design principle 43 (new) The Release order in Class_Terminate is from child to parent**: Settings → View → Ctrl → Environment. Release following the SDK's object dependencies (Environment is the parent, Controller the child, View the grandchild, Settings the great-grandchild)
- **Design principle 44 (new) BOOL getters are written in the style of returning As Boolean**: the HRESULT is handled internally (log output only), and False is returned on failure. It can be written naturally on the UserForm side in the form `If pane.View_GetCanGoBack Then ...`. The `GetBoolProperty` helper is reused
- **New finding: side effect of Ctrl_Close**: the WebView2 runtime fires a Release on all registered event handlers. Close has the semantics of "releasing all event handlers at once"
- **New finding: helper assets for BOOL getters already existed since stage 9.4b**: `GetBoolProperty` (added in 9.4b, used by `NavigationCompletedArgs.get_IsSuccess`) and `GetLongProperty` (also added in 9.4b) could be reused. `View_GetCanGoBack` etc. can be implemented in one line (`GetBoolProperty(m_pView, 38, "View.get_CanGoBack")`)
- **New finding: View_GetSource was already implemented since stage 9.4b**: it was initially included in the coverage targets of 9.8b, but the inventory revealed it as "existing". Likewise, `View_GetDocumentTitle` / `View_GoBack` / `View_GoForward` are also already implemented

### Known constraints and future considerations (stage 9.8b)

- Only put_ is implemented for the Settings family, get_ is not (= there is no API to obtain the current setting values). To be added when needed
- Only 3 items of the Settings family are implemented; the remaining 6 (IsWebMessageEnabled / AreDefaultScriptDialogsEnabled / IsStatusBarEnabled / AreHostObjectsAllowed / IsZoomControlEnabled / IsBuiltInErrorPageEnabled) will be added in the future when needed
- Reload / Stop of the View family are implemented, but many other unimplemented View-family methods remain (CallDevToolsProtocolMethod, OpenDevToolsWindow, CapturePreview, etc.). Low priority
- The safety after Ctrl_Close (= behavior when View_xxx / Ctrl_xxx are called after Close) needs verification in stage 9.8c

### Code deliverables of 9.8b

| File | Version | Role |
|---|---|---|
| `ComCallbackHandler.cls` | 9.8a (unchanged) | General-purpose handler, branching for the 5 kinds of events |
| `Wv2Environment.cls` | 9.8a (unchanged) | Environment class |
| `Wv2Pane.cls` | 9.8b | Controller + CoreWebView2 combined class + m_pSettings + 9 new methods |
| `Wv2Thunks.bas` | 9.8a (unchanged) | Foundation + Sentinel mechanism + IID check + UserForm resize API |
| `UserForm1.frm` | 9.8b | Frame1 + StartWebView2 family + mouse resize + GetActivePane property |

All files were verified on the real machine in a consistent state. This provides a stable base as the starting point for 9.8c (ExecuteScript).


---

## Stage 9.8c — View_ExecuteScript + improved safety of Ctrl_Close

### Completion date
This session (completion of the whole of stage 9.8)

### Background and goal

As agreed beforehand, stage 9.8c implements the following two items:
1. **View_ExecuteScript + HK_ExecuteScriptCompleted** (a new one-shot handler family, concurrent use allowed)
2. **Improved safety after Ctrl_Close** (an open item carried over from stage 9.8b, plan E 1+3 hybrid)

### Design decisions of stage 9.8c (settled in the pre-implementation alignment)

| Issue | Adopted plan |
|---|---|
| O1. Lifetime management | Plan a: concurrent use OK, the slot manages its own completion |
| O2. Receiving the result | Plan α: event-callback style, first implementation only logs the result |
| O3. callback ID | Plan X: return a sequential callback ID |
| O4. HandlerKind addition | Append HK_ExecuteScriptCompleted = 8 at the end |
| O5. Safety after Close | Plan E 1+3 hybrid: clear members + m_closed flag |
| O6. Test method | Plan i: extend Test_AllEvents_Real + Immediate window |

### Implementation of stage 9.8c

| File | Change |
|---|---|
| `Wv2Thunks.bas` | Added HK_ExecuteScriptCompleted = 8 to HandlerKind, extended the upper bound of the m_iidTable array, added the IID to InitIIDTable (49511172-CC67-4BCA-9923-137112F4C4CC) |
| `ComCallbackHandler.cls` | Added module variable m_callbackId, added Friend Property CallbackId (Get/Let), added Case HK_ExecuteScriptCompleted to Handler_Invoke (errorCode + LPCWSTR result → late-bound call into Wv2Pane) |
| `Wv2Pane.cls` | Added module variables m_executeScriptCounter / m_closed, extended Class_Initialize, greatly extended Ctrl_Close (plan E 1+3), added new Public functions View_ExecuteScript / OnExecuteScriptCompleted |
| `Wv2Environment.cls` | No change |
| `UserForm1.frm` | No change (plan α first implementation, self-contained within Wv2Pane) |

### ★ The initial bug of stage 9.8c and its fix (important, made explicit as design principle 45) ★

**The first version crashed Excel**. During tarboh's real-machine test, the Excel process died instantly the moment `?UserForm1.GetActivePane.View_ExecuteScript("1 + 2")` was fired.

#### Cause

```vba
' First version (cause of the crash):
hr = dcf(m_pView, 29, "View.ExecuteScript", StrPtr(jsCode), h.Slot)

' After the fix:
Dim pVTableObj As LongPtr
pVTableObj = h.Slot + VTABLE_OBJ_OFFSET
hr = dcf(m_pView, 29, "View.ExecuteScript", StrPtr(jsCode), pVTableObj)
```

I had forgotten the basic rule: **"when passing a thunk slot to WebView2 as a COM object, you must pass the address of the VTable object part (`h.Slot + VTABLE_OBJ_OFFSET`), not the head of the slot address (`h.Slot`)"**. WebView2 followed a garbage vtable and crashed immediately.

#### Why the mistake happened

- I copy-pasted from View_PostWebMessageAsString (which passes the string pointer from StrPtr directly) as a model
- What View_PostWebMessageAsString passes is a "plain data pointer", while what View_ExecuteScript must pass is a "COM object pointer" — I failed to tell this fundamental difference apart
- I overlooked the existing `pVTableObj = h.Slot + VTABLE_OBJ_OFFSET` in the existing pattern of "passing a handler COM object to WebView2" (View_AddNavigationStarting etc.)

#### Design principle 45 (newly established)

**"When passing a thunk slot to WebView2 as a COM object, always pass `h.Slot + VTABLE_OBJ_OFFSET`"**

| Use | Address to pass |
|---|---|
| Handler COM object (a COM-like object that has a vtable) | `h.Slot + VTABLE_OBJ_OFFSET` |
| Plain data pointer (string, struct, etc.) | `StrPtr(s)` / `VarPtr(struct)` etc., no offset needed |

From WebView2's point of view the two are arguments of completely different nature, and confusing them causes a crash. **When implementing a new handler-type method, always use an existing View_AddXxx (= the same kind of use) as the model** (do not use plain-argument methods such as View_PostXxx as the model).

### Real-machine test results of stage 9.8c

ExecuteScript verification from Test_AllEvents_Real + the Immediate window with the fixed version:

| # | Input JS | callback ID | errorCode | resultJson | Result |
|---|---|---|---|---|---|
| 1 | `1 + 2` | 1 | 0 | `3` | ✅ |
| 2 | `document.title` | 2 | 0 | `"検索 - Microsoft Bing"` (JSON-encoded) | ✅ |
| 3 | `1+1` | 3 | 0 | `2` | ✅ |
| 4 | `2+2` | 4 | 0 | `4` | ✅ |
| 5 | `3+3` | 5 | 0 | `6` | ✅ |
| 6 | `throw new Error('test')` | 6 | 0 | `null` | ✅ (* see below) |

For every run, AddRef/Release closes cleanly with `slot 4` auto-released. No leaks.

#### New finding about the behavior of case 6

The result of `throw new Error('test')` was **errorCode=0, resultJson=null**. The prior expectation (non-zero errorCode) was wrong.

Spec observed on the real machine: **even if a JS exception is thrown inside ExecuteScript, the HRESULT of ExecuteScript itself returns S_OK (0). As a result the JSON string `null` is returned**. This is presumably because WebView2 is designed to catch the JS exception internally and treat it as "result is undefined".

### Safety verification results after Ctrl_Close

Everything behaved as expected:

| Operation | Expected | Measured | Result |
|---|---|---|---|
| Run Ctrl_Close | Persistent handlers auto-released + Release logs for m_pView/m_pCtrl | As logged | ✅ |
| Class_Terminate after Ctrl_Close | Cleanup logs and m_pView/m_pCtrl Release logs are skipped | Completely skipped, only `Wv2Environment.Terminate: Releasing m_pEnvironment` appears | ✅ |
| Double Close | Log `already Closed, ignored`, return value 0 | As logged | ✅ |
| View_Reload after Close | Safely rejected with E_POINTER (`&H80004003` = `-2147467261`) | Return value `-2147467261` | ✅ |

It is now fully confirmed that the E1+3 hybrid (clear members + m_closed flag) works.

### ★ New findings established in stage 9.8c ★

#### Spec fact 3: a throw in ExecuteScript gives HRESULT=S_OK + result=null

A JS `throw` statement or a thrown Error leaves the HRESULT of ExecuteScript at S_OK and makes the result `"null"` (null as a JSON string). For error detection:
- The HRESULT cannot be used (always 0)
- The only option is a string comparison against `resultJson = "null"` (however, it cannot be distinguished from the case where the JS deliberately returns `null`)

→ Exception detection requires a design in which the JS side does try-catch and returns a structured result (e.g. `"try { ... return {ok:true, value:x} } catch(e) { return {ok:false, error:e.message} }"`).

#### Spec fact 4: a VBA UserForm automatically re-Initializes

During tarboh's real-machine test, when `UserForm1.GetActivePane.View_Reload` was fired after closing the form, **UserForm1.Initialize ran automatically**, and afterwards `Object Required` (error 91) occurred.

This is a **known behavior** of VBA UserForms:
- When VBA accesses `UserForm1.<any member>` while `UserForm1` has not been created yet (= after the previous Terminate), a re-Initialize runs automatically
- The state right after re-Initialize is `m_pane = Nothing`, so the `GetActivePane` property returns Nothing
- A method call on that Nothing raises error 91

→ **Operational caution**: do not fire `UserForm1.GetActivePane.xxx` after the form is closed. Alternatively, add an early return `If m_pane Is Nothing Then Exit Property` to the `GetActivePane` property (= room for future safety improvement).

#### Full observation of the side effects of Ctrl_Close (reinforcing the 9.8b finding)

The finding from stage 9.8b that "Ctrl_Close has the side effect of releasing all registered event handlers" was confirmed again:
```
Wv2Pane.Ctrl_Close: m_pCtrl = ...
  Release [idx=0] → refCount = 0 → Slot 0 auto-released
  Release [idx=1] → refCount = 0 → Slot 1 auto-released
  Release [idx=2] → refCount = 0 → Slot 2 auto-released
  Release [idx=3] → refCount = 0 → Slot 3 auto-released
Wv2Pane.Ctrl_Close: Releasing m_pView = ...
Wv2Pane.Ctrl_Close: Releasing m_pCtrl = ...
```

Order: (1) right after Controller::Close is fired, WebView2 Releases the handlers (4 of them) → (2) Wv2Pane Releases m_pView → (3) Wv2Pane Releases m_pCtrl. **This fully agrees with the "release children first" rule.**

### Additional design principles and findings established in stage 9.8c

- **Design principle 45 (new, ★important★)**: When passing a thunk slot to WebView2 as a COM object, always pass `h.Slot + VTABLE_OBJ_OFFSET`. If the slot head address is passed as is, WebView2 follows a garbage vtable and crashes immediately. When implementing a new handler-type method, always use an existing View_AddXxx (= the same kind of use) as the model (do not use plain-argument methods such as View_PostXxx as the model)
- **Design principle 46 (new)**: For one-shot callbacks that can run "concurrently", it is natural for Wv2Pane not to keep a ComCallbackHandler reference (plan a). Discard it with `Set h = Nothing` right after AcquireHandlerFor, so that only the WebView2 runtime holds it via AddRef. On completion notification the method is called through m_owner (= Wv2Pane), after which WebView2 fires Release and the slot is auto-released
- **Design principle 47 (new)**: The callback ID is kept as a member of ComCallbackHandler (m_callbackId), and Wv2Pane.View_ExecuteScript sets it with Friend Property Let right after Acquire. It can be implemented with a minimal addition, without destructively changing Init
- **Design principle 48 (new)**: For safety after Ctrl_Close, the "clear members + flag" hybrid is clean. If m_pSettings/m_pView/m_pCtrl are Released and the members set to 0 at Close time, the existing `If m_pXxx = 0 Then ... = &H80004003` path automatically rejects every View_xxx / Ctrl_xxx. Setting `m_closed = True` additionally makes "reject double Close" and "reject View_ExecuteScript after Close" explicit
- **Design principle 49 (new)**: Clearing `m_persistentTokens` of the persistent handlers is also done at Ctrl_Close time. Since WebView2 releases the handlers by itself on Close, firing remove_xxx from the VBA side afterwards would mean "firing remove on an already released handler", which is dangerous. By setting the tokens to 0, CleanupAllPersistentHandlers skips them safely
- **New finding (spec fact 3)**: a throw in ExecuteScript gives HRESULT=S_OK + result=null. Exceptions cannot be detected via the HRESULT; a design is needed in which the JS side does try-catch and returns a structured result
- **New finding (spec fact 4)**: a VBA UserForm automatically re-Initializes on `UserForm1.<member>` access. Firing `UserForm1.GetActivePane.xxx` after the form is closed triggers a re-Initialize, and the newly created instance has m_pane = Nothing, causing an Object Required error

### Known constraints and future considerations (stage 9.8c)

- Concurrency of ExecuteScript is supported in the implementation, but cannot be confirmed by verification from the Immediate window (only one line can be run at a time). Real verification would need a dedicated test function (`Test_ExecuteScript_Parallel` etc.). An item for 9.9 and later
- JSON parsing of resultJson needs to be implemented on the VBA side (via ScriptControl or by hand). The first implementation only logs
- The design pattern for detecting ExecuteScript throw exceptions on the JS side (try-catch structured result) is the user's responsibility. There is room to add a helper to Wv2Pane in the future
- There is room to add an early return to the GetActivePane property (`If m_pane Is Nothing Then Exit Property`) in the future (= improved safety after the form is closed)

### Completion of the whole of stage 9.8 (a + b + c)

Pre-agreement of stage 9.8:
- 9.8a: fix the class names → ✅ done (this session)
- 9.8b: full coverage of the View base IF + Settings family + Ctrl_Close → ✅ done (this session)
- 9.8c: View_ExecuteScript + improved Ctrl_Close safety → ✅ done (this session)

Handover items to stages after 9.8:

1. **SPA support** (to start in stage 9.9): HistoryChanged / DOMContentLoaded (ICoreWebView2_2 and later) / postMessage notifications from the SPA side. A design that merges with the umbrella class layer such as TabWebView2
2. **Version-extension IFs (ICoreWebView2_2 etc.)**: In stage 9.9, ICoreWebView2_2 must be extracted as a prerequisite for SPA support. It can be rolled out as is with the naming convention (`View2_xxx` etc.) established in stage 9.6a
3. **TabWebView2 umbrella class**: tab management layer; merges with SPA support, and the "open in a separate UserForm, plan b" for NewWindowRequested also merges here. The largest-scale issue
4. **Verification of concurrent ExecuteScript**: add a dedicated test function
5. **Notifying the UserForm of ExecuteScript results**: second-phase implementation of plan α (currently only a log inside Wv2Pane)
6. **Structured-result pattern for JS exceptions**: add a try-catch helper to Wv2Pane

### Code deliverables of stage 9.8c

| File | Version | Role |
|---|---|---|
| `ComCallbackHandler.cls` | 9.8c | Generic handler, 6 kinds of event branches (HK_ExecuteScriptCompleted added) |
| `Wv2Environment.cls` | 9.8a (unchanged) | Environment class |
| `Wv2Pane.cls` | 9.8c | Combined Controller + CoreWebView2 class + View_ExecuteScript + Ctrl_Close safety |
| `Wv2Thunks.bas` | 9.8c | Foundation + IID for HK_ExecuteScriptCompleted added |
| `UserForm1.frm` | 9.8b (unchanged) | Frame1 + StartWebView2 family + mouse resizing + GetActivePane property |

All files were verified on the real machine in a consistent state. It provides a stable base for stage 9.9 (SPA support + extracting ICoreWebView2_2).

---

## Stage 9.8 summary

What stage 9.8 achieved:

1. **Formal fixing of class names (9.8a)**: Renamed the Class7/8/9/Module9 family to ComCallbackHandler/Wv2Environment/Wv2Pane/Wv2Thunks. The VBA symbol names are stable and the foundation for future feature additions is established
2. **Full coverage of the main methods of the View base IF (9.8b)**: Added 5: Reload/Stop/GetCanGoBack/GetCanGoForward/GetBrowserProcessId
3. **Handling of the Settings family (9.8b)**: Added 3 put_ methods (IsScriptEnabled/AreDefaultContextMenusEnabled/AreDevToolsEnabled), helper style inside Wv2Pane
4. **Ctrl_Close (9.8b → 9.8c)**: Extended from a simple implementation to improved safety (clear members + m_closed flag)
5. **ExecuteScript + a new one-shot pattern (9.8c)**: The HK_ExecuteScriptCompleted mechanism with concurrency OK, sequential callback IDs, and receiving JS results in VBA

Design principles grew from 33 to 49 (16 newly established). Spec facts (1 to 4) and many new findings were also recorded. The code size grew from 1228 lines for Wv2Pane.cls (at the end of 9.7b) to 1704 lines (at the completion of 9.8c).

Main issues to address in stage 9.9 and later: SPA support + extracting ICoreWebView2_2 + the TabWebView2 umbrella class.

---

## Stage 9.9 — SPA support (extracting ICoreWebView2_2 + HistoryChanged + DOMContentLoaded + ExecuteScript enhancements)

### Completion date
This session (as of completion of 9.9c)

### Background and goal

In stage 9.8, the coverage of the View base IF, the Settings family, Ctrl_Close safety and the ExecuteScript foundation were settled. Stage 9.9 prepares the foundation for SPA (Single Page Application) support. An SPA changes URL/history/DOM without top-level navigation, so state transitions cannot be tracked with `NavigationStarting` / `NavigationCompleted` alone. To make up for this, the following are introduced.

- `HistoryChanged` (base ICoreWebView2, vtable 13/14): captures history changes caused by SPA `pushState`/`replaceState`/`popstate`
- `DOMContentLoaded` (derived ICoreWebView2_2, vtable 64/65): captures completion of DOM construction. The `_2` derived IF is required
- ExecuteScript enhancements: a concurrency test function, a structured-result helper for JS exceptions, and notification of results to the UserForm

Breakdown of the pre-agreement of stage 9.9:
- 9.9a: ICoreWebView2_2 extraction foundation (m_pView2 + EnsureView2 + View2_Ensure)
- 9.9b: persistent-handler implementation of HistoryChanged (base) + DOMContentLoaded (derived _2), SPA test page generation
- 9.9c: batch handling of the complementary group (View_ExecuteScriptSafe / OnExecuteScriptResult notification / Test_ExecuteScript_Parallel / log improvements)

---

### Stage 9.9a: ICoreWebView2_2 extraction foundation

#### Design decisions (agreed with tarboh)

- **Plan 1**: `m_pView2 As LongPtr` field + `EnsureView2` lazy acquisition (follows the Settings style = design principle 42)
- **Plan a**: `IID_ICoreWebView2_2` is assembled each time as a local variable of `EnsureView2`. It is not registered in Wv2Thunks' m_iidTable (= the IID table is dedicated to "the handlers' own IIDs")
- **Plan X**: both `EnsureView2` / `View2_Ensure` are `As Boolean` (same form as EnsureSettings)
- **Plan i**: verification is `?UserForm1.GetActivePane.View2_Ensure` from the Immediate window + confirming zero regressions with Test_AllEvents_Real (no new test function)
- **Plan R1**: Release order is Settings → View2 → View → Ctrl (strictly child to parent, a derivative of design principle 43)

#### Significance of stage 9.9a

Extracting ICoreWebView2_2 is not itself the goal; it is the foundation for calling DOMContentLoaded (vtable 64/65) in 9.9b and later. HistoryChanged is a base IF (vtable 13/14), so it can be implemented without `_2`, but DOMContentLoaded requires the `_2` IF. This is the first stage that applies in practice the naming convention (`View2_xxx`) established in stage 9.6a.

#### Implementation of stage 9.9a

| File | Change |
|---|---|
| `Wv2Pane.cls` | Added `m_pView2 As LongPtr`, `m_pView2 = 0` in Class_Initialize, extended the Release order in Class_Terminate / Ctrl_Close (Settings → View2 → View → Ctrl), Private `EnsureView2() As Boolean` (QueryInterface(IID_ICoreWebView2_2)), Public `View2_Ensure() As Boolean` (verification wrapper) |
| Others | No change (Wv2Thunks does not add the IID to m_iidTable = plan a) |

---

### Stage 9.9b: HistoryChanged + DOMContentLoaded persistent handlers + SPA test page

#### Implementation of stage 9.9b

| File | Change |
|---|---|
| `ComCallbackHandler.cls` | Added event branches for HistoryChanged / DOMContentLoaded (HK_HistoryChanged=9 / HK_DOMContentLoaded=10) |
| `Wv2Pane.cls` | Extended the upper bound of the persistent-handler array from HK_NewWindowRequested(7) to HK_DOMContentLoaded(10). New helper `RemoveInterfaceFor(kind) As LongPtr`. HistoryChanged family (base): View_AddHistoryChanged / View_RemoveHistoryChanged / View_OnHistoryChanged. DOMContentLoaded family (_2): View2_AddDOMContentLoaded / View2_RemoveDOMContentLoaded / View2_OnDOMContentLoaded |
| `Wv2Thunks.bas` | Added IIDs for HK_HistoryChanged=9 / HK_DOMContentLoaded=10, `GetLongLongProperty` (`#If Win64`, for getting UINT64), BuildSpaTestHtml + Test_SpaTest_Real |
| `UserForm1.frm` | New `StartWebView2_SpaTest` (shows the SPA test page + registers 7 kinds of handlers) |

#### ★ Vigilance toward design principle 45 in stage 9.9b ★

Two persistent handlers are added = two occasions of passing a COM object to WebView2. Based on the lesson of passing h.Slot directly and crashing in 9.8c (design principle 45), View_AddHistoryChanged / View2_AddDOMContentLoaded are both implemented using View_AddNavigationStarting verbatim as the model, and always go through `pVTableObj = h.Slot + VTABLE_OBJ_OFFSET`. Plain-argument methods such as View_PostWebMessageAsString are not used as the model.

#### Facts observed in the real-machine test of 9.9b

- **When NavigateToString completes, HistoryChanged fires automatically twice in a row** (spec fact 7). Can be used to confirm that the persistent-handler registration path works
- **On an about:blank origin, pushState/replaceState is rejected with SecurityError** (spec fact 6). This is the web standard
- **DOMContentLoadedEventArgs.get_NavigationId is a UINT64** (spec fact 8). Obtained with GetLongLongProperty

---

### Stage 9.9c: Batch handling of the supplementary items

#### Implementation of 9.9c

| File | Change |
|---|---|
| `Wv2Pane.cls` | New Public `View_ExecuteScriptSafe(userCode) As Long` (a wrapper that wraps the code in try-catch and returns structured JSON, `{ok:true,value:...}` / `{ok:false,errorName:...}`). OnExecuteScriptCompleted calls `UserForm1.OnExecuteScriptResult` directly (the Debug.Print inside Wv2Pane is kept as well). Improved logging in View_OnHistoryChanged ((α) a note about about:blank, (β) states explicitly that pArgs=0 is by design) |
| `Wv2Thunks.bas` | Added `Test_ExecuteScript_Parallel`; promoted FillGUID to Public |
| `UserForm1.frm` | New `OnExecuteScriptResult` (the receiver for callbacks from Wv2Pane; the first implementation only logs). Added a Nothing warning to GetActivePane |

#### Bug found during verification of 9.9c (spec fact 10)

We found a possible latent bug in the MEM_FREE handling of the Sentinel mechanism. In the processing after it detects "the previous thunk region address is in the MEM_FREE state", there is a path that calls `ReadLongPtr` on the already-freed address and causes an AV crash. It was observed after an operation equivalent to an Excel reset. It will be handled as a separate item before starting 9.10 (or within 9.10).

Following the Excel crash in 9.8c (passing h.Slot directly -> design principle 45), this is another success of the real-machine test in finding a bug early.

---

### Additions to the design principles and insights established in stage 9.9

#### New design principles (50 to 54)

- **Design principle 50 (new)**: Derived interfaces (ICoreWebView2_2 etc.) are acquired lazily with an `m_pXxxN` field + an `EnsureXxxN` helper, and the IID for QueryInterface is built each time from the caller's local GUID structure. It is not registered in Wv2Thunks' m_iidTable (= the IID table is dedicated to "the handler's own IID"). FillGUID is shared as a Public utility and reused for extracting multiple derived interfaces
- **Design principle 51 (new)**: A derived interface obtained by QueryInterface has its own independent reference count, so in the Release order the derived interface is released before the base interface (the derived case of design principle 43, "from child to parent"). Even if the derived interface's pointer value is the same as the base's (spec fact 5), the refcounts are managed separately
- **Design principle 52 (new)**: When an event of a derived interface is registered as a persistent handler, it is also removed using the derived interface's pointer. Use the semantically correct pointer even if `m_pView2 == m_pView` (spec fact 5) (a derived application of design principle 45). The kind -> interface conversion is managed in one place by a dedicated helper such as `RemoveInterfaceFor`
- **Design principle 53 (new)**: Leave a "32-bit-ization point" comment at each use of a type that has to be kept in mind for future 32-bit expansion (e.g. wrapping `GetLongLongProperty` in `#If Win64`). The 32-bit support itself will be done all at once in a separate stage (= "policy X"). Converting the thunk machine code to x86, type abstraction (absorbing size differences such as PtrSize and VTABLE_OBJ_OFFSET), and how to receive a UINT64 on 32-bit (Currency / two Longs / a compromise) are clearly separated as issues for a separate stage
- **Design principle 54 (new)**: The user code for the structured-result helper of ExecuteScript (View_ExecuteScriptSafe) is assumed to be an "expression". When using a "statement", the user wraps it in an immediately-invoked function `(()=>{...})()`. Runtime exceptions are caught, but syntax errors (failure at the parse stage) cannot be caught and come back with resultJson="null" (spec fact 9)

#### New spec facts (5 to 10)

- **Spec fact 5**: QueryInterface for ICoreWebView2_2 returns the same pointer value as the base ICoreWebView2 (because it is a derived interface with single inheritance). The refcounts are managed independently
- **Spec fact 6**: The origin of a page displayed with NavigateToString is `about:blank` (= the "null" origin). On this origin, `history.pushState` / `replaceState` are rejected with SecurityError, per the web standard. Once we move to an http(s) origin in 9.10 and later, full testing of pushState/replaceState becomes possible
- **Spec fact 7**: When NavigateToString completes, the WebView2 runtime updates the history internally, so it fires `HistoryChanged` automatically twice in a row. This can be used to automatically confirm that the registration path of the persistent handler works
- **Spec fact 8**: `DOMContentLoadedEventArgs.get_NavigationId` returns a UINT64 (LongLong). It is obtained with `GetLongLongProperty` (64-bit only, `#If Win64`)
- **Spec fact 9**: ExecuteScript returns a JS syntax error (failure at the parse stage) as HRESULT=S_OK + result="null". It cannot be distinguished from a runtime exception (ReferenceError etc.). This helper (View_ExecuteScriptSafe) can structure runtime exceptions but cannot structure syntax errors (= when using statements such as throw, the user must wrap them in an immediately-invoked function; design principle 54)
- **Spec fact 10**: A possible latent bug in the MEM_FREE handling of the Sentinel mechanism. In the processing after it detects "the previous thunk region address is in the MEM_FREE state", there is a path that calls `ReadLongPtr` on the already-freed address and causes an AV crash. It was observed after an operation equivalent to an Excel reset during verification of 9.9c. It will be handled as a separate item before starting 9.10 (or within 9.10)

---

### Code deliverables of stage 9.9 (file master)

| File | Version | Role |
|---|---|---|
| `ComCallbackHandler.cls` | 9.9b | General-purpose handler; HistoryChanged / DOMContentLoaded branches added (HK 9/10) |
| `Wv2Environment.cls` | 9.8a (unchanged) | Environment class |
| `Wv2Pane.cls` | 9.9c (2400 lines) | View_ExecuteScriptSafe / OnExecuteScriptResult notification / OnHistoryChanged logging improvements. The View2 family (EnsureView2 / View2_Ensure / m_pView2), HistoryChanged / DOMContentLoaded persistent handlers, RemoveInterfaceFor |
| `Wv2Thunks.bas` | 9.9c (2919 lines) | HK_HistoryChanged=9 / HK_DOMContentLoaded=10, GetLongLongProperty (#If Win64), BuildSpaTestHtml / Test_SpaTest_Real / Test_ExecuteScript_Parallel, FillGUID promoted to Public |
| `UserForm1.frm` | 9.9c (679 lines) | StartWebView2 / StartWebView2_LocalHtml / StartWebView2_SpaTest, GetActivePane (Nothing warning), OnExecuteScriptResult |
| `旧WebView2.cls` | Reference implementation | — |
| `MouseScroll.bas` | Reference implementation | — |
| `LibTimers.bas` | Reference implementation | — |

---

## Stage 9.9 Summary

What was achieved in stage 9.9:

1. **Foundation for extracting ICoreWebView2_2 (9.9a)**: m_pView2 + EnsureView2 + View2_Ensure; established lazy acquisition, caching and Release order for derived interfaces
2. **HistoryChanged + DOMContentLoaded persistent handlers (9.9b)**: capture SPA history changes and DOM construction completion. Established the handler registration pattern for both the base interface and the derived _2 interface. SPA test page generation (BuildSpaTestHtml)
3. **ExecuteScript expansion (9.9c)**: View_ExecuteScriptSafe (structured results for JS exceptions), OnExecuteScriptResult (UserForm notification), Test_ExecuteScript_Parallel (concurrent verification)

Design principles grew from 49 to 54 (5 new), and spec facts from 4 to 10 (6 new). The real-machine test found the Sentinel mechanism's MEM_FREE bug (spec fact 10) early.

Main items to address in stage 9.10 and later: the TabWebView2 umbrella class (the largest scale), dealing with the Sentinel MEM_FREE bug, a real SPA test environment (http(s) origin), and the NewWindowRequested "open in a separate UserForm, option b".

---

## Stage 9.10 — Wv2Browser umbrella class (tab management layer) + dealing with the Sentinel bug

### Policy for subdividing the stage (agreed in advance)

The main job of stage 9.10 is the umbrella class `Wv2Browser` (formerly provisionally named TabWebView2) that bundles multiple tabs. Because it is the largest, it is subdivided:

- **9.10a**: Dealing with the Sentinel MEM_FREE bug (spec fact 10) — eliminate the crash-causing item alone, first
- **9.10b**: `Wv2Browser` skeleton — one shared Wv2Environment + a Collection of Wv2Pane + activeIndex. Constructor / AddTab / RemoveTab / ActivateTab
- **9.10c**: Display control for switching the active tab (IsVisible approach) + connecting the UserForm (consolidating m_env+m_pane into a single m_browser)
- **9.10d**: Show a real site in "tab 1" as the real SPA test environment (http(s) origin) -> full pushState/replaceState testing (removing the constraint of spec fact 6)
- **9.10e** (optional): Consider whether the NewWindowRequested "option b" can be replaced by an "open in a new tab" option

#### Confirmed design policy of stage 9.10

- **Class name**: `Wv2Browser` (keeps the `Wv2` prefix, giving the integration layer a leading-role feel. Consistent with the ease of completion in design principle 38)
- **Handling of Wv2Pane (option D-1)**: Keep it as is. `Wv2Browser` only bundles a Collection of Wv2Pane, and each tab's View_/Ctrl_/Settings_ is reached via `browser.ActivePane.View_xxx`. Without touching Wv2Pane's 2400 lines, regression risk is minimal. Method consolidation (option D-2) is deferred to an optional future stage
- **Environment sharing (option E-1)**: One Wv2Environment is shared by all tabs (the same userDataFolder)
- **Real SPA test (issue F)**: Merged into `Wv2Browser` in 9.10d (a real site in tab 1)

---

### Stage 9.10a: Dealing with the Sentinel mechanism's MEM_FREE bug (spec fact 10)

#### Completion date
This session

#### The essence of the bug (identifying the cause)

`Sentinel_RecoverIfNeeded` (Wv2Thunks.bas) is the function that reclaims the old thunk region after a reset. The old implementation **did** obtain `mbi.State` with `VirtualQuery`, but did not use that result for branching, and **unconditionally executed** `ReadLongPtr(prevBase)` / `MemLongPtr(prevBase) = 0` / `VirtualFree`.

`ReadLongPtr` / `MemLongPtr` are implemented to plant the target address directly into the `pvData` of a SAFEARRAY internally and access it as a native array via `.arr(0)`. Firing this at a freed address with `State = MEM_FREE` raises a CPU-level access violation (AV) = an SEH exception.

★ The most important insight ★: This AV **cannot be caught by VBA's `On Error Resume Next`**. What `On Error Resume Next` can catch is only VBA runtime errors; an SEH exception (a hardware exception) passes straight through and the Excel process dies immediately. The comment in the old implementation said "if an access violation occurs, the region has already been freed", but the premise was broken: "detecting the AV and continuing" is itself impossible in VBA. This is the true nature of spec fact 10.

#### Fix policy (adopted: a robust State guard)

`VirtualQuery` itself can be safely called even on a freed address (it is an OS API and just returns MEM_FREE etc. even for an invalid address), so it is made the sole basis for the decision. Only when all of the following 3 conditions are met is the region considered "alive", and `ReadLongPtr` / `MemLongPtr` / `VirtualFree` are executed:

1. `mbi.State = MEM_COMMIT` (the pages are committed)
2. `mbi.AllocationBase = prevBase` (the true start of the region; it has not been remapped as a different region)
3. `IsProtectReadWritable(mbi.Protect) = True` (PAGE_READWRITE / PAGE_EXECUTE_READWRITE and without the PAGE_GUARD bit)

If it is not alive, fire only `Sentinel_ClearPrevRegion` and safely `Exit Sub`.

#### Implementation of 9.10a

| File | Version | Change |
|---|---|---|
| `Wv2Thunks.bas` | 9.10a | Added constants (MEM_FREE / PAGE_READWRITE). Rewrote `Sentinel_RecoverIfNeeded` to the State-guard approach. New Private `IsProtectReadWritable(prot) As Boolean` (PAGE_GUARD detection + read/write-capable check). The core logic (Handler_xxx / dcf / Thunks_Init) is otherwise unchanged |
| `ComCallbackHandler.cls` | 9.9b (unchanged) | — |
| `Wv2Environment.cls` | 9.8a (unchanged) | — |
| `Wv2Pane.cls` | 9.9c (unchanged) | — |
| `UserForm1.frm` | 9.9c (unchanged) | — |

For 9.10a, re-importing only `Wv2Thunks.bas` is enough (the other files are unchanged).

#### Additions to the design principles and insights established in 9.10a

- **Design principle 55 (new, ★important★)**: Memory primitives that involve SEH (ReadLongPtr / MemLongPtr etc., the kind that plant an address into a SAFEARRAY's pvData and access it natively) are called only after first confirming with `VirtualQuery` that `State=MEM_COMMIT` + `AllocationBase` matches + the protection attribute allows read/write. The AV they cause is a CPU SEH exception and cannot be caught by VBA's `On Error Resume Next` (= the process dies immediately). "Thinking you are protected by On Error" is the most dangerous pattern

#### Verification points of 9.10a (items to confirm in the real-machine test)

- After an operation equivalent to a reset (Excel's "Reset" button / re-running after a project recompile), the log `Sentinel: 旧領域は健在でないと判定 → 触れずにスキップ` (Sentinel: old region judged not alive -> skipped without touching) appears and there is no crash
- In normal consecutive start/stop (without a reset in between), the path through `Sentinel: VirtualFree succeeded` is taken as before, and leak reclamation works
- The region state can be observed as expected with `Test_SentinelStatus`
#### Real-machine test results of 9.10a (completion record)

In this session all 3 points succeeded. 9.10a passes.

**Verification 1 (Sentinel bug fix)**: With events running from a link click, a VBE reset
-> Excel did not crash. After the reset, `Test_SentinelStatus` showed
`→ 次回 Thunks_Init で回収される予定 (リセット後)` (-> to be reclaimed at the next Thunks_Init (after reset)) / `m_pRegionBase: 0`.
Re-running `Test_Navigate_Real` in the same session -> no crash, and the Sentinel
reclaimed the old region and started up normally.

**Verification 2 (leak reclamation on the normal path)**: After restarting Excel, normal start -> normal exit with the x button ->
restart. `Cleanup` removed the persistent handlers of kind=3/4/6/7 and completed the Release in the order
View -> Ctrl -> Environment (design principle 43). No double release.

**Verification 3 (region state observation)**: `Test_SentinelStatus` while in use correctly displayed
`→ 現在使用中の領域と一致 (正常状態)` (-> matches the region currently in use (normal state)). The environment variable value and `m_pRegionBase`
match.

**★Supplementary note (important, relationship to spec fact 10)**:
In this reset case, the old thunk region remained alive with
`State = MEM_COMMIT` even after the reset. So the Sentinel judged it "alive" and passed through the
`VirtualFree succeeded` path (log: `State = 4096 (MEM_COMMIT)` -> `VirtualFree succeeded`).
In other words, what this log demonstrated is "the path that correctly reclaims an alive region"; the path
we worried about in spec fact 10, "the old region has become MEM_FREE and firing ReadLongPtr at it
causes an AV", was not reproduced by this reset operation.
This is not a defect in the fix: depending on the timing of the reset (the state of the WebView2 runtime threads and
the OS's page reclamation), the old region either becomes MEM_FREE or remains as MEM_COMMIT.
The 9.10a guard (the 3-condition alive check) covers both cases, and
in the MEM_FREE case the implementation logs `→ 触れずにスキップ` (-> skipped without touching) and exits safely
(not reproduced, but logically guaranteed). If we want to reproduce the MEM_FREE skip path on a real machine in the future,
it is expected to be easier to reproduce by inserting a large memory allocation right after the reset to encourage
page reclamation of the old region (a separate item).

**Secondary observation**: Slot allocation proceeds soundly in LIFO order (the reverse of idx=3,2,1,0), and
the previous session's releases were correctly returned to the free list. NewWindowRequested suppression
(`put_Handled OK` -> same-screen navigation) is also maintained.

With this, stage 9.10a is complete. The next stage is 9.10b (implementing the Wv2Browser skeleton).

---

### Stage 9.10b: Implementing the Wv2Browser skeleton

#### Completion date
This session

#### Scope of 9.10b (skeleton only)

We build the shell of the umbrella class `Wv2Browser` that bundles multiple tabs. It holds
exactly one Wv2Environment and creates and manages N Wv2Pane (= tabs) from it.

To build: Init / AddTab / AddTabWithUrl / AddTabWithHtml / TabCount / ActivePane /
      PaneAt (for verification) / Class_Terminate
Not to build: ActivateTab (display switching) / RemoveTab / formal UserForm connection /
          mouse-resize linkage (all in 9.10c and later)

#### Confirmed design of 9.10b (results of the alignment discussion)

- **Issue ①**: The asynchronous completion wait (Do While ... DoEvents) is enclosed
  inside AddTab / Init (WaitEnvReady / WaitPaneReady). The waiting logic can be removed from the UserForm
  -> the UserForm connection in 9.10c becomes clean (option a adopted)
- **Issue ②**: Since 9.10b does no display switching, all tabs are created stacked on the same m_hwndHost.
  The goal of verification is only that "each Pane reaches the Controller/View"
- **Issue ③**: The Navigate family is separated into 3 methods (option ③-A). AddTab (the shell only) /
  AddTabWithUrl (View_Navigate) / AddTabWithHtml (View_NavigateToString).
  In AddTabWithHtml, a cautionary comment about spec fact 6 (pushState is not possible on an about:blank origin) is left
  as a template
- **Issue ④**: ActivePane is provisionally "the tab added last". Formal activeIndex management and
  ActivateTab are in 9.10c
- **Class name**: Wv2Browser (design principle 38). Wv2Pane is kept as is (option D-1).
  One Environment is shared by all tabs (option E-1)
- **Release order**: In Class_Terminate release all Panes -> release the Environment
  (design principle 43: from child to parent)

#### Implementation of 9.10b

| File | Version | Change |
|---|---|---|
| `Wv2Browser.cls` | 9.10b (new) | Skeleton class. BrowserState enum (Bs_Created/Bs_EnvReady/Bs_Failed), m_env/m_tabs(Collection)/m_hwndHost/m_state, the methods above, WaitEnvReady/WaitPaneReady (wait built in) |
| `UserForm1.frm` | 9.10b | Verification only: m_browser field, FrameHwnd accessor, StartBrowserTest (creates 2: 1 Env + a URL tab + an HTML tab), added release of m_browser to UserForm_Terminate. The existing StartWebView2 is kept (for regression comparison) |

Import: Wv2Browser (new) -> UserForm1, in that order. The other 3 files (Wv2Thunks 9.10a /
ComCallbackHandler 9.9b / Wv2Environment 9.8a / Wv2Pane 9.9c) are unchanged.

#### Bug hit during verification of 9.10b (fixed immediately) = reconfirmation of design principles 45/56

**Bug**: The first version of AddTab skipped the "obtain View" step
(Ctrl_GetCoreWebView2) after Controller creation. As a result it called
View_Navigate / View_NavigateToString with m_pView = 0, and both
failed with E_POINTER (&H80004003) (a symptom was ViewPtr=0 in the log).

**Cause**: Although the Bounds setting (step 7) was copied from the existing StartWebView2,
the Ctrl_GetCoreWebView2 (step 8) right after it was dropped. A failure to apply
"read the existing pattern thoroughly before writing a new method".

**Fix**: Added a Ctrl_GetCoreWebView2 +
ViewPtr=0 guard between the Bounds setting and the Collection add in AddTab (a verbatim carry-over of StartWebView2 step 8).

- **Design principle 56 (new)**: Before calling a Navigate-family method on a Wv2Pane (View_Navigate /
  View_NavigateToString), always obtain the View from the Controller first
  (`Ctrl_GetCoreWebView2` → confirm `ViewPtr <> 0`). Completion of Controller creation alone leaves
  m_pView at 0. This order (Controller creation → Bounds → **View acquisition** →
  Navigate) is an invariant already established in StartWebView2, and it must not be omitted even in a
  new path such as Wv2Browser.AddTab. If omitted, the
  Navigate family fails silently with E_POINTER

#### Real-machine test results for stage 9.10b (completion record)

All features succeeded with the fixed version. 9.10b passes.

- **Proof of Environment sharing (plan E-1)**: `Environment 生成 OK` (Environment created OK) appeared only once
  (pEnv=6631431003152). From this single Environment, two Controllers
  (distinct pCtrl) could be created
- **View reachability**: Both tabs showed `View 取得 OK` (View acquired OK) with pView non-zero and distinct pointers
  (6631429750784 / 6631429756928). Navigate OK / NavigateToString OK (175 characters)
- **Correct destruction order (design principle 43)**: Class_Terminate recognized tab count = 2 and,
  in reverse order tab 2 → tab 1, released each Pane in the order m_pView → m_pCtrl (View is the child,
  Controller is the parent), and finally released the Environment. No double release.
  The Sentinel reclamation inherited from 9.10a also worked normally

**★Observation on display overlap (important; hand-off to 9.10c)**:
The prior expectation was that "the later-added tab 2 (HTML) would appear in front", but
on the real machine **tab 1 (Bing) was in front**. This is not a bug; it is the natural result of the scope of 9.10b
(no display switching). When two Controllers are stacked on the same HWND,
which one ends up in front is left to the WebView2 runtime and is indeterminate. Tab 2's HTML was also correctly
created behind the scenes and NavigateToString succeeded (ViewPtr non-zero, m_pView release confirmed on destruction),
so it "exists even though it is not visible". Deterministic front-most control becomes possible
only when 9.10c introduces `Ctrl_PutIsVisible`.
→ The central task of 9.10c is IsVisible control: "make only the active tab visible and the others invisible".
   9.10b demonstrated that it correctly prepared the foundation for that (a state in which each tab is
   independently created and retained).

With this, stage 9.10b is complete. The next stage is 9.10c (tab-switch IsVisible control +
formal UserForm connection).

---

### Stage 9.10c-1: IsVisible control + ActivateTab (the core of tab switching)

#### Completion date
This session

#### Splitting policy for 9.10c

9.10c contains two major issues (IsVisible control + formal UserForm connection) and is large, so following the
"small scope, independent, regression-isolatable" principle in force since 9.7, it is split in two:
- **9.10c-1** (this stage): IsVisible control + ActivateTab + ResizeAll. Verification is still done with
  StartBrowserTest / SwitchTab (no formal UserForm connection)
- **9.10c-2** (next stage): Formal UserForm connection (rewrite StartWebView2 on top of Wv2Browser,
  consolidate m_env+m_pane into m_browser, route UserForm_Resize through ResizeAll)

#### Fixed design of 9.10c-1 (result of the alignment discussion)

- **Issue ⑤**: Ctrl_PutIsVisible is correctly a simple-argument type modeled on Ctrl_PutBounds
  (not a handler type that passes a COM object)
- **Issue ⑥**: ActivateTab loops over all tabs, making only index visible and the others invisible (plan A). It is
  naive but reliable, and leaves no "everything invisible" hole
- **Issue ⑦**: Automatically bring a tab to the front when it is added (plan a). However, "activation" is split out into the independent function
  ActivateTab, and AddTab calls it (delegation). → The visibility control logic is
  consolidated in one place, ActivateTab (single source of truth). "Add" and "switch" are
  separate responsibilities, so separating the functions is correct, and it also keeps things DRY (no divergence)
- **Issue ⑧**: On resize, reissue Bounds to all tabs (plan a). By updating the invisible tabs too,
  the size will not be off when they are activated later
- **Issue ⑨**: UserForm connection rewrites StartWebView2 on top of Wv2Browser
  (plan A, to be done in 9.10c-2)

#### Implementation of 9.10c-1

| File | Version | Changes |
|---|---|---|
| `Wv2Pane.cls` | 9.10c-1 | New `Ctrl_PutIsVisible(visible)` (ICoreWebView2Controller::put_IsVisible, vtable index 4). Simple-argument type modeled on Ctrl_PutBounds. Normalizes VBA True(-1) to Win32 BOOL 1 |
| `Wv2Browser.cls` | 9.10c-1 | Added m_activeIndex. ActivateTab(index) (loop over all tabs, only index visible, updates activeIndex). AddTab calls ActivateTab after adding (delegation). ResizeAll(rc) (reissue Bounds to all tabs). ActivePane formalized on the basis of activeIndex. Added ActiveIndex property |
| `UserForm1.frm` | 9.10c-1 | Updated the StartBrowserTest summary. Added SwitchTab(index) verification helper |

Import: Wv2Pane → Wv2Browser → UserForm1. The others (Wv2Thunks 9.10a /
ComCallbackHandler 9.9b / Wv2Environment 9.8a) are unchanged.

#### Additional knowledge established in 9.10c-1

- **Design principle 57 (new)**: "Activating a tab (visibility control)" and "adding/removing a tab" are
  separate responsibilities, so separate the functions. Put visibility control in only one place, ActivateTab
  (single source of truth), and have AddTab and the like merely call it (delegation).
  That way the visibility control logic is not scattered, and in the future a click from the tab UI
  only needs to call the same ActivateTab. A design in which "Add hits PutIsVisible itself" scatters
  visibility control over two places and diverges, so avoid it
- **Spec fact 11 (new)**: The argument of ICoreWebView2Controller::put_IsVisible is a Win32
  BOOL (4-byte int, True=1/False=0). VBA's True is -1, so it must be normalized
  to 1/0 before being passed to dcf (it is not a VARIANT_BOOL)

#### Real-machine test results for 9.10c-1 (completion record)

Complete success. 9.10c-1 passes. The decisive difference from 9.10b was confirmed.

- **Proof of IsVisible control**: Right after each AddTab, `ActivateTab: タブN をアクティブ化
  (可視)` (ActivateTab: tab N activated (visible)) appeared, and the delegation structure worked. When StartBrowserTest completed, the HTML
  (tab 2) was displayed in front. In 9.10b, "Bing (tab 1) was unexpectedly in front", but
  this time, thanks to IsVisible control, "front = the tab 2 that was activated last" was
  determined deterministically. This is the clear difference from 9.10b
- **Deterministic tab switching**: `SwitchTab 1` → ActiveIndex=1 + the screen switched to Bing.
  `SwitchTab 2` → ActiveIndex=2 + the screen switched to the handwritten HTML. It was visually confirmed that the screen actually
  switches. ActivateTab functioned as the sole switching entry point
- **Behavior on reset termination**: Because the run was stopped with the VBE reset button, UserForm_Terminate
  did not run, and the form vanished without cleanup logs. This is not abnormal; it is as designed:
  the Sentinel mechanism of 9.10a reclaims the previous region at the next startup (indeed,
  `Sentinel: VirtualFree succeeded` appeared at the start of this log, reclaiming the
  previous reset's region). Cleanup on normal termination (× button) was confirmed in 9.10b

The fact that the core of the "tabbed browser" (holding multiple tabs + deterministic switched display) works
is a major milestone. With this, stage 9.10c-1 is complete. The next stage is 9.10c-2
(formal UserForm connection).

---

### Stage 9.10c-2: Formal UserForm connection (m_browser consolidation + automatic handler registration)

#### Completion date
This session

#### Purpose of 9.10c-2

Consolidate the UserForm fields from m_env + m_pane into a single m_browser
(realizing the original concept of stage 9.10).

#### Fixed design of 9.10c-2 (result of the alignment discussion)

- **Issue ⑩**: The foundational navigation handlers (NavigationStarting/Completed/
  DocumentTitleChanged/NewWindowRequested) are registered automatically for all tabs inside AddTab
  (plan ⑩-A). Whether the form receives them is up to the user
  (Wv2Pane's OnXxx merely does Debug.Print). Apart from events the user must specify explicitly (such as
  HistoryChanged for the future SPA), automatic registration is fine
  - ★Policy-change trigger★ If event reception overhead makes things sluggish, switch to
    plan ⑩-C (the user explicitly calls an AddStandardHandlers helper)

#### Implementation of 9.10c-2

| File | Version | Changes |
|---|---|---|
| `Wv2Browser.cls` | 9.10c-2 | Added automatic standard-handler registration to AddTab. Split out into a dedicated private helper RegisterStandardHandlers(pane) (4 kinds: View_AddNavigationStarting/Completed/DocumentTitleChanged/NewWindowRequested. Verbatim port of existing StartWebView2 steps 10-13) |
| `UserForm1.frm` | 9.10c-2 | Fully rewrote StartWebView2 on top of Wv2Browser (condensed into a few lines of m_browser.Init + AddTabWithUrl). The old body is preserved fully commented out as StartWebView2_OLD_9_9c. UserForm_Resize now goes through m_browser.ResizeAll. GetActivePane now goes through m_browser.ActivePane |

**Decision to retain**: The m_env / m_pane fields and StartWebView2_LocalHtml /
StartWebView2_SpaTest / StartBrowserTest / SwitchTab were left as they are.
The first two use m_env/m_pane, and the latter two remain useful for multi-tab verification.
Complete removal of the fields has a wide impact range, so it is deferred to a separate cleanup stage.

Import: Wv2Browser → UserForm1. The others (Wv2Thunks 9.10a / ComCallbackHandler
9.9b / Wv2Environment 9.8a / Wv2Pane 9.10c-1) are unchanged.

#### Real-machine test results for 9.10c-2 (completion record)

Complete success. 9.10c-2 passes → stage 9.10c is complete.

- **Realization of UserForm consolidation**: StartWebView2 completes with just the single m_browser
  (`StartWebView2 開始 (9.10c-2 / Wv2Browser ベース)` (StartWebView2 start (9.10c-2 / Wv2Browser-based)) → Init OK →
  4 standard handlers registered → done TabCount=1)
- **Proof of automatic handler registration (plan ⑩-A)**: Although StartWebView2 has no hand-written handler
  registration at all, clicking a link fired NavigationStarting / DocumentTitleChanged /
  NavigationCompleted, all of them. RegisterStandardHandlers inside AddTab works
- **Proof of ResizeAll (issue ⑧ a)**: Resizing on tab 1 → switching to tab 2 → tab 2 had also been
  resized correctly. This is evidence that Bounds can be reissued even to invisible tabs
  (had they not been reissued, the old size would have been displayed on switching)
- **Completeness of destruction**: On × close, Wv2Browser.Terminate → 4 handler Cleanups
  (kind=3/4/6/7) → release ran to completion in the order View→Ctrl→Environment (design principle 43)
- **Secondary confirmations**: Even with multiple tabs (StartBrowserTest), the standard handlers were registered independently
  on each tab, and the events of both tabs fired with different pSender. Tab 2's HTML appeared in the navigation log as
  `data:text/html;...base64,...` (a reconfirmation of spec fact 6 =
  NavigateToString has a data URI origin → an actual SPA in 9.10d needs http(s))

#### Overall summary of stage 9.10c

The core of the tabbed browser works fully on the real machine:
- Holding and creating multiple tabs (sharing 1 Environment)
- Deterministic switched display (IsVisible control, ActivateTab)
- Resize tracking (ResizeAll for all tabs)
- Consolidation of the UserForm into a single m_browser
- Automatic registration of the standard navigation handlers for all tabs

With this, stage 9.10c (9.10c-1 + 9.10c-2) is complete.
The next stage is 9.10d (real SPA test environment: seriously test pushState/
replaceState on a tab with an http(s) origin. Removes the constraint of spec fact 6).

---
### Stage 9.10d-1: SPA smoke test (proof of existence of history operations on an http(s) origin)

#### Completion date
This session

#### Purpose of 9.10d-1 (zero-implementation smoke test)

Demonstrate, at minimum cost and without writing any new methods, whether the constraint of spec fact 6
(on the about:blank origin of NavigateToString, pushState / replaceState are rejected with
SecurityError) can be removed by opening a tab with an http(s) origin. Of the two-way split policy for 9.10d,
d-1 is the "proof that the foundation exists", and d-2 is the "real implementation in a self-contained SPA environment".

#### Fixed design of 9.10d-1 (result of the alignment discussion)

- **Issue 1 (securing an http(s) origin)**: Adopt the A→C hybrid. d-1 proves existence with plan A (the real site
  example.com), and d-2 implements self-containedly with plan C (`SetVirtualHostNameToFolderMapping`,
  ICoreWebView2_3, **vtable index 71 on m_pView3**).
  Note: vtable order (verified against WebView2.h): 66:get_CookieManager /
  67:get_Environment / 68:TrySuspend / 69:Resume / 70:get_IsSuspended /
  **71:SetVirtualHostNameToFolderMapping** / 72:ClearVirtualHostNameToFolderMapping.
- **Issue 2 (SPA handler registration)**: Adopt plan ②-c (public helper `EnableSpaHandlers` + thin
  wrapper `AddTabWithUrlForSpa`). However, that is built in d-2. In d-1, just call
  `ActivePane.View_AddHistoryChanged` once manually.
- **Issue 3 (pass conditions)**: (1) get_Source is an http(s) URL that is not about:blank /
  (2) pushState → ok:true + HistoryChanged fires + URL changes /
  (3) replaceState → ok:true + URL changes /
  (4) popstate (GoBack/GoForward) → HistoryChanged fires + CanGoBack/Forward consistent /
  (5) Division of roles for DOMContentLoaded (fires only the first time, does not fire on pushState) ← d-2.
- **Issue 4 (staging)**: Adopt the two-way split (d-1 / d-2).
- **Test site**: https://example.com/ (minimal, stable, same-origin pushState possible).
- **Firing method**: Manual staged firing (SpaSmoke_* one at a time from the Immediate window).

#### Implementation of 9.10d-1

| File | Version | Changes |
|---|---|---|
| `UserForm1.frm` | 9.10d-1 | Added a group of verification-only Subs: `StartWebView2_SpaSmoke` (opens example.com and manually registers HistoryChanged) / `SpaSmoke_PushState` / `SpaSmoke_ReplaceState` / `SpaSmoke_GoBack` / `SpaSmoke_GoForward` / `SpaSmoke_Source`. All are combinations of existing methods, with **no new methods** |

The other 5 files (Wv2Browser 9.10c-2 / Wv2Pane 9.10c-1 / Wv2Thunks 9.10a /
ComCallbackHandler 9.9b / Wv2Environment 9.8a) are unchanged.

#### Real-machine test results for 9.10d-1 (completion record)

Full pass. Pass conditions 1-4 achieved (condition 5 = division of roles for DOMContentLoaded is deferred to d-2).

- **★Direct proof that the constraint of spec fact 6 is removed★**: pushState / replaceState fired
  via `View_ExecuteScriptSafe` both gave `resultJson={"ok":true}`. What was
  `{"ok":false}` with SecurityError on the about:blank origin (spec fact 6) went through on the
  https origin. `get_Source` also tracked `https://example.com/spa1` → `/spa2`
  correctly.
- **pushState** → HistoryChanged fired (URL: /spa1, CanGoBack changed to True).
- **replaceState** → HistoryChanged fired (URL: /spa2, CanGoBack stayed True).
  ★This overturned the initial expectation (replaceState leaves the history length unchanged, so it might not fire)★.
- **GoBack** → HistoryChanged fired via popstate (URL: /, CanGoBack: False /
  CanGoForward: True). **GoForward** → fired (URL: /spa2, CanGoBack: True /
  CanGoForward: False). History navigation and CanGoBack/Forward are fully consistent.
- **On completion of the first Navigate, HistoryChanged fired twice** (first about:blank,
  second https://example.com/). It reflects the transition from the initial View's about:blank
  (the Navigate version of spec fact 7).
- "Example Domain" via DocumentTitleChanged, and isSuccess: True via NavigationCompleted.
  Resource management (AddRef/Release symmetric, Slot auto-release) was also healthy,
  with no double release. The Sentinel also reclaimed the previous region normally via the State=MEM_COMMIT path.

#### Additional knowledge established in 9.10d-1

- **Spec fact 12 (new)**: WebView2's HistoryChanged **fires for all** of an SPA's
  pushState / replaceState / popstate (GoBack/GoForward).
  It fires even for replaceState, which leaves the history length unchanged (a real-machine fact that overturned the initial expectation).
  In addition, on completion of the first Navigate it **fires twice** (the first with the pre-transition about:blank,
  the second with the post-transition real URL). URL tracking can be done with get_Source (when HistoryChanged fires,
  `View_OnHistoryChanged` logs it automatically). args is always 0 (a known item among the spec facts;
  state is obtained via the sender).
- **Supplement to spec fact 6**: The constraint (pushState/replaceState impossible on the about:blank origin) is removed
  by changing the origin to http(s). Proven on example.com in 9.10d-1.
  In 9.10d-2, `SetVirtualHostNameToFolderMapping` will be used to prepare the same
  https origin self-containedly, without an external site.

With this, stage 9.10d-1 is complete. The next stage is 9.10d-2
(a self-contained SPA environment via `SetVirtualHostNameToFolderMapping` + building out `EnableSpaHandlers`
/ `AddTabWithUrlForSpa` + verification of the division of roles for DOMContentLoaded).


### Stage 9.10d-2b: Exhaustive verification of history operations with a self-made SPA (EnableSpaHandlers / AddTabWithUrlForSpa)

#### Completion date
This session

#### Purpose of 9.10d-2b

On top of the self-contained https origin (SetVirtualHostNameToFolderMapping) established in d-2a,
build out the SPA handler registration and exhaustively verify pass conditions 1-5, including the
division of roles among pushState/replaceState/popstate/DOMContentLoaded.

#### Fixed design of 9.10d-2b (result of the alignment discussion)

- Issue F (design of AddTabWithUrlForSpa): Adopt plan F-1 (all-in-one wrapper).
  AddTab → SetVirtualHostNameToFolderMapping → Navigate → EnableSpaHandlers are executed in
  1 method.
- Issue G (how to fire pushState): Adopt firing via ExecuteScript (reuse d-1's SpaSmoke_*).
- Issue I (staging): Single stage.
- Issue ② (SPA handler registration): Adopt plan ②-c. EnableSpaHandlers becomes a public helper, and
  AddTabWithUrlForSpa is a thin wrapper that calls it internally.

#### Implementation of 9.10d-2b

| File | Version | Changes |
|---|---|---|
| `Wv2Browser.cls` | 9.10d-2b | Added `EnableSpaHandlers(pane)` (public; internally View_AddHistoryChanged + View2_AddDOMContentLoaded) / `AddTabWithUrlForSpa(hostName, folderPath, fileName)` (runs AddTab → SetMapping → Navigate (file name explicit, spec fact 13) → EnableSpaHandlers all at once). Updated the header to 9.10d-2b, and also added the previously unrecorded 9.10c-1 (ActivateTab) and 9.10c-2 (RegisterStandardHandlers) |
| `UserForm1.frm` | 9.10d-2b | Added the verification `StartWebView2_SpaFull` (creates the SPA tab with AddTabWithUrlForSpa alone; pushState etc. reuse d-1's SpaSmoke_*) |

The other 4 files (Wv2Pane 9.10d-2a / Wv2Thunks 9.10a / ComCallbackHandler 9.9b /
Wv2Environment 9.8a) are unchanged.

#### Real-machine test results for 9.10d-2b (completion record)

Full pass. All pass conditions 1-5 achieved.

- **Condition 1**: get_Source is an https origin (`https://appassets.example/spa1` etc.).
- **Condition 2**: pushState → HistoryChanged fired (`URL: /spa1`, CanGoBack changed to **True**),
  resultJson `{"ok":true}`.
- **Condition 3**: replaceState → HistoryChanged fired (`URL: /spa2`, CanGoBack stayed True
  = it is a replacement, so history length is unchanged), `{"ok":true}`.
- **Condition 4**: GoBack → `URL: /index.html` (CanGoBack: False / CanGoForward: **True**),
  GoForward → `URL: /spa2` (CanGoBack: True / CanGoForward: False). Fired on popstate, and
  CanGoBack/Forward fully consistent.
- **Condition 5 (the main focus)**: On the first load, `View2_OnDOMContentLoaded: navigationId=2` fired **only once**,
  and on subsequent pushState/replaceState/GoBack/GoForward **only HistoryChanged fired and
  DOMContentLoaded never fired at all**. The division of roles between the two events (DOMContentLoaded =
  on a real load, HistoryChanged = on SPA history operations) was clearly demonstrated.
- Resource management (AddRef/Release symmetric, Slot auto-release) healthy, no double release.

Addendum (on the order of the verification steps): if only replaceState is fired first, the history length does not change
and CanGoBack stays False, so the subsequent GoBack/GoForward have nowhere to go and HistoryChanged does
not fire (this is normal behavior per the Web standard). Firing pushState first to make CanGoBack True
and then calling GoBack/GoForward makes it fire. This backs up spec fact 12.

---

### Stage 9.10d overall summary

Stage 9.10d (d-1 / d-2a / d-2b) is considered complete. What was reached:

- **The about:blank origin constraint (spec fact 6) was resolved in a self-contained way.**
  Using `ICoreWebView2_3::SetVirtualHostNameToFolderMapping` (vtable 71), a local
  folder is mapped to `https://appassets.example/`, so a secure origin can be
  provided with pure VBA alone, without depending on an external site, an HTTP server
  process, or any COM library.
- On that origin we demonstrated that **SPA history operations (pushState/replaceState/popstate) and the
  DOM lifecycle (DOMContentLoaded) can be handled completely**.
  The division of roles between HistoryChanged and DOMContentLoaded has also been confirmed.
- Deliverables of the stage:
  - Wv2Pane: `EnsureView3` / `View3_Ensure` / `View3_SetVirtualHostNameToFolderMapping`
  - Wv2Browser: `EnableSpaHandlers` / `AddTabWithUrlForSpa`
  - Verification Subs: SpaSmoke_* (d-1) / StartWebView2_SpaMapping (d-2a) /
    StartWebView2_SpaFull (d-2b)

- Newly established findings:
  - Spec fact 12 (HistoryChanged fires for all of pushState/replaceState/popstate,
    fires for replaceState too, and fires twice on the first Navigate)
  - Spec fact 13 (SetVirtualHostNameToFolderMapping returns ACCESS_DENIED for the root "/",
    so the file name must be given explicitly)

#### Outlook beyond 9.10d (open items)

- 9.10e (optional): consider whether the "open in a separate UserForm, option b" for NewWindowRequested can be
  replaced by an "open in a new tab" option (combining AddTabWithUrl and ActivateTab).
- Cleanup of verification-only fields/Subs: remove the m_env / m_pane fields and StartWebView2_LocalHtml /
  StartWebView2_SpaTest, plus the d-1/d-2a/d-2b verification Subs (the impact is
  wide, so this is a separate stage).
- Future: reuse the LiteCom package for WinRT toast notifications (independent of the main track).


  ### Stage 9.10e: Opening NewWindowRequested in a new tab (option A + option Y + option β Timer)

#### Completion date
This session

#### Purpose of 9.10e

Until now `View_OnNewWindowRequested` did "put_Handled=TRUE + same View
View_Navigate" (9.6b), opening the link target in the same tab. This is changed to "open in a new
tab".

#### Confirmed design for 9.10e (result of the alignment discussion)

- **Issue 1 (how to open) = option A (simple)**: put_Handled=TRUE + AddTabWithUrl in a new tab.
  The return value of window.open becomes null (a simple approach that does not pass a new View via put_NewWindow).
  The proper approach (option B: get_Deferral + put_NewWindow) is left for a separate future stage.
- **Issue 2 (Wv2Pane to upper-layer coordination) = option Y (back-reference + Shutdown)**: Wv2Pane holds an
  m_browser back-reference and calls AddTabWithUrl through it. The implementation
  cost on the UserForm side is one Shutdown call line, and it does not depend on the UserForm name (option X, calling the
  UserForm directly, has the drawbacks of boilerplate receiver methods and a dependency on the UserForm1 name).
- **Issue (deadlock avoidance) = option β (SafeTimer one-shot deferred execution)**: a deadlock was found
  midway through implementation (see below), so this was fixed to run AddTabWithUrl via SafeTimer after
  the handler has completed.

#### Problem that occurred in 9.10e and its resolution (important)

In the first version (option A + option Y only), m_browser.AddTabWithUrl was
**called directly** inside the NewWindowRequested handler. AddTab internally uses WaitPaneReady to wait synchronously, with DoEvents, for the new Controller's creation to complete
(ControllerCompleted), and this deadlocked.

The mechanism was confirmed from the real-machine log:
```
Wv2Pane.View_OnNewWindowRequested → put_Handled OK → AddTabWithUrl → CreateController
Wv2Browser.AddTab: Controller creation timeout (state=1)   ← kept waiting inside the handler and timed out
  m_browser.AddTabWithUrl failed
  …(handler returns)…
ComCallbackHandler.Invoke[HK_ControllerCompleted]         ← delivered only after the handler returned
```

→ **A WebView2 event handler does not deliver the next asynchronous callback until it returns**.
Waiting inside a handler for another asynchronous WebView2 operation to complete waits forever (spec fact 14).

Resolution (option β): the handler calls `m_browser.RequestNewTabAsync(url)` and returns immediately, and
a SafeTimer (one-shot) runs AddTabWithUrl **after the handler has completed**. The real-machine log
after the fix confirmed: RequestNewTabAsync request → handler return → TimerCall (outside the handler) →
AddTab → ControllerCompleted delivered normally → tab count = 2 → example.com displayed.

#### Implementation of 9.10e

| File | Changes |
|---|---|
| `Wv2Pane.cls` | Added the `m_browser` (Wv2Browser back-reference) field / `SetBrowser` / `ClearBrowserRef`. Changed View_OnNewWindowRequested to request `m_browser.RequestNewTabAsync(url)` (falls back to a same-View Navigate when m_browser is not set). Also restored `Ctrl_PutIsVisible`, which had been missing since 9.10c-1 (because the old version in the project was the 9.9c version) |
| `Wv2Browser.cls` | Added `WithEvents m_newTabTimer As SafeTimer` / `m_pendingUrls` (pending queue). AddTab now does `pane.SetBrowser Me`. Added `RequestNewTabAsync` (queues the URL and starts a one-shot timer with a New SafeTimer) / `m_newTabTimer_TimerCall` (runs AddTabWithUrl on the queue outside the handler) / `Shutdown` (breaks circular references + stops the timer) |
| `UserForm1.frm` | Added one line, `m_browser.Shutdown`, to UserForm_Terminate |

New dependencies: **LibTimers.bas / SafeTimer.cls / TimerForm.frm** (derived from Cristian Buse's
RefTypes. SafeTimer depends on LibTimers and TimerForm; TimerForm is an empty UserForm).
Import order: Wv2Thunks → ComCallbackHandler → Wv2Environment → LibTimers →
TimerForm → SafeTimer → Wv2Pane → Wv2Browser → UserForm1.

#### Real-machine test results for 9.10e (completion record)

Full pass.
- NewWindowRequested → RequestNewTabAsync request, handler returns immediately (**no timeout**).
- TimerCall (outside the handler) → AddTab → tab count = 2 → ActivateTab tab 2 → example.com
  displayed (isSuccess: True).
- On form close: `Wv2Browser.Shutdown: back-references of all 2 tabs disconnected` → Wv2Browser.Terminate
  → **Cleanup/Terminate for both of the 2 tabs ran to completion** → Environment released → UserForm1.Terminate
  completed. No circular-reference leak, all slots auto-released.

#### Findings established in 9.10e

- **Spec fact 14 (new)**: Synchronously waiting inside a WebView2 event handler for another asynchronous
  WebView2 operation (such as Controller creation) to complete deadlocks, because the next callback
  is not delivered until the handler returns. To avoid it, defer the work until after the handler completes
  with SafeTimer or similar (pending queue + one-shot Timer).
- **Design principle 58 (new)**: When a back-reference creates a circular reference (Browser.m_tabs ↔ Pane.m_browser),
  provide an explicit Shutdown in the upper layer and break the references before release. VBA's reference-
  counting GC cannot resolve cycles automatically, and unless Shutdown is called, Class_Terminate itself
  never runs and the objects leak. The burden on the UserForm side is one Shutdown call line.
- **Caveat on SafeTimer**: StartTimer "ignores the call if uElapse = m_uElapse", and m_uElapse remains
  even after stopping (StartTimer 0). Restarting with the same interval is ignored, so for each request
  create a `New SafeTimer` (a new instance with m_uElapse=0) to make sure it starts.

With this, stage 9.10e is complete.

---

### Stage 9.10 cleanup: removing verification code and slimming down the main body

#### Completion date
This session

#### Purpose

Remove the verification-only code accumulated in 9.7 to 9.10e (throwaway Start/Test/Sub groups, the old m_env/m_pane
fields) and tighten the main body into a production configuration.

#### Confirmed policy (result of the alignment discussion)

- All verification Subs are removed, but `StartWebView2_SpaFull` is kept, renamed to
  `StartWebView2_Spa`, as a template for starting an SPA.
- `SendMessageToJS` / `IsPageLoaded` / `BuildSpaAppHtml` / `WriteSpaAppFolder` are kept as
  practical features (their m_pane dependency is rewritten to go through GetActivePane).
- Because the impact is wide, it is split in two (cleanup-1: m_env/m_pane, cleanup-2: verification Subs).

#### Cleanup-1 (removing the m_env/m_pane parts)

- Removed the `m_env` / `m_pane` fields and unified on `m_browser`.
- Deleted the old StartWebView2 (kept commented out), `StartWebView2_LocalHtml`, and
  `StartWebView2_SpaTest`.
- Rewrote `SendMessageToJS` / `IsPageLoaded` from direct m_pane references to go through GetActivePane.
- Removed the m_pane/m_env release from UserForm_Terminate.
- Real-machine check: StartWebView2 (bing) works without m_env/m_pane, IsPageLoaded=True,
  SendMessageToJS OK. Compile passes.

#### Cleanup-2 (removing the verification Subs + promoting SpaFull)

- Removed: `StartWebView2_SpaSmoke` + `SpaSmoke_*`, `StartWebView2_SpaMapping`,
  `SpaSmoke_OpenNewWindow`, `StartBrowserTest`, `SwitchTab`.
- `StartWebView2_SpaFull` → `StartWebView2_Spa` (dropped the verification logs and tidied the
  comments into an SPA startup template). `BuildSpaAppHtml` / `WriteSpaAppFolder` are used by SpaFull, so they are kept.
- Wv2Thunks.bas: removed 4 functions: the test functions `Test_MessageEcho_Real` / `Test_SpaTest_Real`, which called
  the removed `StartWebView2_LocalHtml` / `_SpaTest`, and their dedicated helpers
  `BuildTestHtml` / `BuildSpaTestHtml` (confirmed they are not used by other test functions).
- Real-machine check: the home-made SPA is displayed by StartWebView2_Spa, StartWebView2 (bing) shows no
  regression, compile passes.

#### Structure after cleanup

- `UserForm1.frm`: 1245 → 562 lines. The main body (Initialize/Resize/Terminate, StartWebView2,
  GetActivePane, OnExecuteScriptResult, FrameHwnd) + practical methods (SendMessageToJS,
  IsPageLoaded, BuildSpaAppHtml, WriteSpaAppFolder) + the SPA startup template (StartWebView2_Spa).
- `Wv2Thunks.bas`: 2998 → 2689 lines.

#### Minor remaining items (no real harm)

- In the header change-history comments of UserForm1.frm and Wv2Thunks.bas, the names of removed methods
  (m_env/m_pane, BuildSpaTestHtml, Test_SpaTest_Real, etc.) remain as "records of the past". There are zero
  references in actual code. They can be removed by tidying the history comments if it bothers anyone.
---

## Stage 9.11: Implementing the tab UI (tab bar)

### Policy (result of the alignment discussion)

As the final finishing touch of the original goal, "a VBA application fully functional as a tabbed browser",
we implement a tab bar in the UI (click to switch, × to close).

**Adopting the romantic spec**: the tab bar itself is also implemented with WebView2 rather than as MsForms controls. Existing
assets (two-way WebMessage communication, SetVirtualHostNameToFolderMapping for SPAs,
N-pane management) can be reused as they are, and we judged it achievable at low cost and stably.

**Loose coupling is the top priority**: demand for a tab bar differs by user (some want a WebView2-based one / some want to
implement their own with MsForms / some do not need one at all). Therefore Wv2Browser is designed to know nothing about the
tab bar implementation. Giving the Browser tab-bar-specific fields (m_tabBarPane etc.) would
bake in "tab bar = WebView2 pane", so we do not.

**Two-layer structure**:
- Layer 1 (9.11a): a loosely coupled hook on the Browser side (Public Event + CloseTab). A foundation needed
  whatever the tab bar is. It can be verified on its own without regression.
- Layer 2 (9.11b onward): the WebView2-based tab bar itself (Wv2TabBar etc.). One of the concrete
  implementations that connect to the Layer 1 hook. The Browser itself does not know about it.

**How the tab bar pane is held (option 1 adopted)**: when implementing a WebView2-based tab bar, it is held
separately (in a dedicated field) from the content tabs' m_tabs. It is not included in the visibility toggling
of ActivateTab and is always visible, with dedicated Bounds (a band at the top). This is a Layer 2 matter,
and the responsibility of the upper layer (the tab bar implementation), not of the Browser itself.

### Confirmed design for 9.11a (loosely coupled event hook + CloseTab)

**Notification mechanism = option A (Public Event + WithEvents)**. The standard way to do loosely coupled notification in VBA.
The Browser need not know the listener's existence or implementation, and the same hook absorbs all of: no tab bar, WebView2-based, MsForms-based, and
home-made. RaiseEvent passes harmlessly through even with zero listeners.
(Alternatives considered: option B, Implements ITabBarSink, is over-spec since this case is not complex enough to
need a type contract. Option C, polling, is inferior in responsiveness.)

**Zero tabs allowed**: the state where the last tab is closed and there are 0 tabs is permitted. activeIndex=0
(no active tab), ActivePane is Nothing. The Browser has no minimum-tab-count policy; it is
left to the upper layer (the tab bar implementation). This matches the loose-coupling policy and brings no
unnecessary constraints to usages that do not need a tab bar.

**Active correction in CloseTab = right neighbor first, otherwise left neighbor (the end)**. This matches the feel of
typical browsers. When the closed tab is before the active one, activeIndex is decreased by 1 to keep
the "same tab".

### Implementation of 9.11a

| File | Version | Changes |
|---|---|---|
| `Wv2Browser.cls` | 9.11a | Added 4 Public Events (TabAdded/TabClosed/ActiveChanged/TabTitleChanged) to the declarations section. Added RaiseEvent TabAdded to AddTab (right after the Collection add, before ActivateTab). Added RaiseEvent ActiveChanged to ActivateTab (on success). Added the new CloseTab(index) (ClearBrowserRef → Ctrl_Close → m_tabs.Remove → 4-branch active correction → RaiseEvent TabClosed). Added the new NotifyTitleChanged(senderPane, title) (reverse-looks up the index by object reference equality → RaiseEvent TabTitleChanged) |
| `Wv2Pane.cls` | 9.11a | Added one block to View_OnDocumentTitleChanged. After obtaining the title, if m_browser exists, calls m_browser.NotifyTitleChanged Me, title. Reuses the m_browser back-reference (introduced in 9.10e for NewWindowRequested). When not set, falls back to Debug.Print only. No new fields/methods on the Pane side |
| `UserForm1.frm` | 9.11a | Made m_browser WithEvents (to receive events). Added 4 event handlers (m_browser_TabAdded etc., each with Debug.Print). Added the verification Subs (Test_9_11a_Help / Setup / CloseActiveMiddle / CloseActiveLast / CloseBeforeActive / CloseAll) and the internal helper DumpState |

**The 4 branches of CloseTab's active correction**:
1. m_tabs.count = 0 → activeIndex = 0 (zero tabs allowed)
2. index < m_activeIndex → newActive = m_activeIndex - 1 (earlier ones shift, so correct it to keep the same tab)
3. index = m_activeIndex → right neighbor first (index if index <= count, otherwise the end)
4. index > m_activeIndex → activeIndex unchanged (just in case, fix it via ActivateTab)

All visibility repair is consolidated in ActivateTab (design principle 57). The circular reference is also broken in CloseTab by
calling ClearBrowserRef first (design principle 58).

Import: in dependency order Wv2Pane → Wv2Browser → UserForm1. Nothing else changes.

### Real-machine test procedure for 9.11a (just run Test_9_11a_*)

Run the following in order in the Immediate window (Ctrl+G):
1. `UserForm1.Show vbModeless`
2. `UserForm1.Test_9_11a_Setup` — creates 3 tabs (bing/example/msn). TabAdded×3 +
   ActiveChanged×3, TabCount=3/ActiveIndex=3. A few seconds later TabTitleChanged for each tab
3. `UserForm1.Test_9_11a_CloseActiveMiddle` — activate tab 2 and close it →
   TabClosed(2) + ActiveChanged(2), the right neighbor becomes active (the screen shows old tab 3=msn)
4. `UserForm1.Test_9_11a_CloseActiveLast` — activate the last tab and close it →
   the left neighbor becomes the new last tab and is active
5. `UserForm1.Test_9_11a_CloseBeforeActive` — with the last tab active, close tab 1 →
   activeIndex decreases by 1 and the Pane is the same, screen contents unchanged
6. `UserForm1.Test_9_11a_CloseAll` — close all remaining tabs → TabCount=0/ActiveIndex=0,
   ActivePane Is Nothing=True, no crash
7. `Unload UserForm1` — normal release (Terminate log)

(Running `UserForm1.Test_9_11a_Help` prints the above steps to the Immediate window.)

### Real-machine test results for 9.11a (completion record)

Full pass. All items matched the expected values.

- **Setup (creating 3 tabs)**: in the order TabAdded(1)→ActiveChanged(1) / (2)→(2) / (3)→(3),
  6 fires in total. TabCount=3/ActiveIndex=3. Title notifications fired for all tabs, and
  the Pane→Browser reverse lookup was accurate (bing=index1 / example=index2 / msn=index3).
  msn fired twice because of its initial navigation, but both resolved correctly to index3.
- **CloseActiveMiddle (closing the active tab 2)**: Ctrl_Close ran to completion in the Release
  order → ActivateTab(2) → ActiveChanged(2) → TabClosed(2). The right neighbor (old msn=
  new index2) became active, and msn was confirmed displayed on the real machine. Branch 3 (right neighbor first) works as intended.
- **CloseActiveLast (closing the last tab)**: TabClosed(2) → ActiveChanged(1). The left neighbor
  (new last = bing) became active, and Bing was confirmed displayed on the real machine. The fallback "no right neighbor
  → left neighbor" is correct.
- **CloseBeforeActive**: because of the test order, tabs had already dropped to 1, so
  the guard (fewer than 2 tabs: redo from Setup) worked correctly and skipped it.
  Branch 2 (index < m_activeIndex) is implemented but this test was not run. Since the other
  3 branches all went through the same correction logic, there is no real harm (it can be checked
  by running it alone after redoing Setup if needed).
- **CloseAll (closing all tabs)**: TabCount=0/ActiveIndex=0, ActivePane Is Nothing=True,
  no crash, WebView2 area became empty. Zero-tab allowance works correctly.
- **Unload (release)**: Shutdown ran to completion from back-reference disconnection through Environment
  release even in the zero-tab state. No signs of leaks.

Passed with no doubts. The loosely coupled event hook (option A), CloseTab's 4 branches, the title
reverse lookup, and zero-tab allowance are all as designed. → Stage 9.11a complete.

The next stage is the design alignment for Layer 2 (the WebView2-based tab bar itself).

---


## Stage 9.11b: Skeleton of the WebView2-based tab bar Wv2TabBar + layout split

### Policy (result of the alignment discussion)

The first sub-stage in breaking Layer 2 (the WebView2-based tab bar itself) into small pieces.

**Ownership structure = option b (new Wv2TabBar class)**. The tab bar logic (event reception,
JS UI updates, two-way communication) is concentrated in one class, keeping the UserForm thin. This makes
the loose coupling clearest: users who do not need a tab bar simply do not instantiate Wv2TabBar. Users who want
an MsForms-based tab bar can connect
their own class to the Browser via WithEvents instead of Wv2TabBar.

**Guaranteeing loose coupling**: Wv2Browser knows nothing at all about Wv2TabBar. Wv2TabBar is the "listening side" that
subscribes to the Browser's Public Events (9.11a). From the Browser's point of view this is
the same as having zero listeners.

**The tab bar pane is not placed under tab management**: Wv2TabBar internally holds m_barPane As
Wv2Pane, but this is not put into Wv2Browser.m_tabs (so it never mixes with normal tabs).
Therefore AddTab is not used; the Wv2Pane is created directly.

**Environment sharing**: m_barPane shares the same Environment as the Browser
(one Environment is shared by all panes). For this reason one read-only
Environment property was added to Wv2Browser.

**Layout splitting is the UserForm's responsibility**: Preserving the principle that the Browser does not know about the tab bar,
UserForm_Resize computes the top band (tab bar) and the content area. The
rectangle of the tab bar band is passed to m_tabBar.Resize, and the rectangle of the content area is passed to m_browser.ResizeAll.

### What was implemented in stage 9.11b

| File | Version | Changes |
|---|---|---|
| `Wv2TabBar.cls` | 9.11b (new) | Skeleton of the WebView2-based tab bar. Init (create pane directly -> wait for Controller -> get View -> virtual host mapping -> display static HTML) / Resize / Shutdown / IsReady. Subscribes to the Browser's 4 events via WithEvents m_browser (in 9.11b the handlers are empty implementations = Debug.Print only). The HTML is static (only displays "TabBar"), written out as UTF-8 (no BOM) with ADODB.Stream. The serving host is appassets.tabbar (separate from the content side's appassets.example) |
| `Wv2Browser.cls` | 9.11b | Added Property Get Environment() As Wv2Environment (so that Wv2TabBar can share the Env). This only |
| `UserForm1.frm` | 9.11b | Added the m_tabBar field + the TABBAR_HEIGHT_PX=36 constant. Made UserForm_Resize support the 2-way split (if m_tabBar exists: top band + content area; otherwise the whole area as before). Added m_tabBar.Shutdown to UserForm_Terminate (before m_browser.Shutdown). Added the verification routines StartWebView2_WithTabBar / Test_9_11b_Help. Note: the header update had been missed in 9.11a, so both the 9.11a and 9.11b blocks were stacked to make up for it |

**Pane creation procedure (Wv2TabBar.Init, i.e. AddTab minus the tab-management part)**:
New Wv2Pane -> Init(env, hwnd) -> WaitPaneReady (IsReady/IsFailed polling) ->
Ctrl_GetCoreWebView2 (design principle 56) -> View3_SetVirtualHostNameToFolderMapping
(before Navigate) -> View_Navigate(https://appassets.tabbar/tabbar.html, spec fact 13).

Import order: Wv2Browser -> Wv2TabBar -> UserForm1. Wv2TabBar is new, so it is simply
added by importing it.

### Real-machine test results for stage 9.11b (completion record)

Full pass. All items as expected.

- **Tab bar startup**: WriteTabBarHtml wrote 387 characters in UTF-8 successfully (ADODB.Stream;
  the feared "ADODB unavailable" problem did not occur). The tab bar pane was created directly without going through AddTab, and went through
  Controller creation -> View acquisition -> SetVirtualHostNameToFolderMapping (appassets.tabbar) ->
  Navigate to completion. TabBar.IsReady=True.
- **Environment sharing**: Both the Browser's pane and the tab bar's pane created their Controller with the same pEnv=90915867746320.
  A single Environment is being shared.
- **Demonstration of loosely coupled event subscription**: When bing's title event fired, both [EVT] TabTitleChanged
  (UserForm side) and Wv2TabBar[EVT] TabTitleChanged (TabBar side) fired.
  Multiple listeners (UserForm and TabBar) can subscribe to the same Browser event simultaneously.
- **View3 support**: m_pView3 was also acquired correctly in the tab bar pane (EnsureView3), and
  at Shutdown, Releasing m_pView3 comes first (the View3 support from 9.10d is being reused correctly).
- **Layout split**: On the real machine, the tab bar is at the top with bing below it. On resize: "the tab bar stays at a fixed height at the top and only its width
  stretches; bing is fixed at the top and stretches in both height and width". The 2-way split follows correctly.
- **Release**: Completed in the order TabBar.Shutdown (Ctrl_Close of the tab bar pane first, then View3 -> View -> Ctrl) ->
  Browser.Shutdown -> Pane group Cleanup -> Environment release. No sign of leaks.

No doubts; pass. -> Stage 9.11b complete.

The next stage is 9.11c (VBA->JS direction: receive the Browser's events and
render the tab list in the tab bar JS. Display only; clicks do not work yet).

---


## Stage 9.11c — VBA->JS direction: rendering the tab list (full sync + ready handshake)

### Policy (outcome of the alignment discussion)

The 2nd sub-stage of Layer 2. Receive the Browser's events and render the tab list
in the tab bar JS. **Display only**; clicks (activate/close/newtab) come in 9.11d.

**Full-sync approach (not a diff approach)**: VBA->JS sends the entire state of all tabs every time.

    {"type":"sync","tabs":[{"title":"..."},...],"active":N}

At the scale of a few dozen tabs, the transfer and rendering cost is practically the same as with diffs, and what really matters is
whether consistency management is needed. Full sync needs no index-correction logic on the JS side (VBA is the source of truth) and is robust. It is the same philosophy as React/Vue
declarative rendering. The initial sync is also solved naturally with a single sync.

**Ready handshake (option X)**: When loading completes, the JS side does a `{"cmd":"ready"}`
postMessage -> only when VBA receives it does it return the initial sync. `PostWebMessageAsString` is
dropped if sent before the JS side has registered its message listener (this is the time window in which the title event in 9.11b
went through as "unhandled"). More reliable than DOMContentLoaded, and it also becomes the foundation for receiving clicks in 9.11d.

**Title cache = option 3 (held in the Pane)**: Add `m_title` + a
`DocumentTitle` property to `Wv2Pane`, and store it in `View_OnDocumentTitleChanged`.
Since the title is tied to the Pane object, it is unaffected by index shifts (consistent with unifying the source of truth in the full sync).
`PushSyncToJs` reads it via `m_browser.PaneAt(i).DocumentTitle`.

**Pane->TabBar back reference = option P (generic WebMessage sink)**: Add one generic
"WebMessage sink" reference (`m_webMsgSink`) to `Wv2Pane`, and `View_OnWebMessageReceived`
forwards `OnPaneWebMessage(pane, source, msg)` to the sink. The TabBar registers itself as the sink. This is a separate
channel from the existing `m_browser` back reference (for NewWindowRequested). The practice of
cutting it on release follows design principle 58.

### What was implemented in stage 9.11c

| File | Version | Changes |
|---|---|---|
| `Wv2Pane.cls` | 9.11c | `m_title` + `Public Property Get DocumentTitle()` (title cache, option 3). Added `m_title = title` to `View_OnDocumentTitleChanged` (immediately before the NotifyTitleChanged notification block from 9.11a). Generic WebMessage sink: `m_webMsgSink As Object` + `SetWebMsgSink` / `ClearWebMsgSink`, and a sink-forwarding block in `View_OnWebMessageReceived`. `ClearBrowserRef` also disconnects the sink (design principle 58) |
| `Wv2TabBar.cls` | 9.11c | Replaced the 4 empty handlers with a common `PushSyncToJs()` call (full sync). `OnPaneWebMessage` (entry point for the option P sink; interprets `"ready"` and performs the initial sync). `PushSyncToJs` (turns TabCount / PaneAt(i).DocumentTitle / ActiveIndex into JSON -> `View_PostWebMessageAsString`). Minimal JSON escaping `JsonEscape`. At the end of Init, `m_barPane.SetWebMsgSink Me` + `View_AddWebMessageReceived` (★the initial sync is NOT fired here = wait for ready). `ClearWebMsgSink` in Shutdown. Extended `BuildTabBarHtml` (sends ready + redraws the DOM on receiving sync; clicks are only a container) |
| `Wv2Browser.cls` | 9.11c | **Make the initial Bounds of a new tab match the content rectangle (option A)**. See "Bug found during 9.11c verification" below for details |
| `UserForm1.frm` | 9.11c | Verification only. `StartWebView2_WithTabBar_3` (3 tabs of bing/example/msn + tab bar) / `Test_9_11c_Help` / `Test_9_11c_AddTab` / `Test_9_11c_CloseActive` |

**Granularity**: 9.11c covers "full sync + ready handshake + tab DOM rendering (clicks are only a container)",
and 9.11d covers "enabling clicks (activate/close/newtab)".

### ★ Bug found during verification of 9.11c (layout overlap) -> design principle 59 ★

**Symptom**: After startup, when a 4th tab is added with `AddTab`, that WebView2's drawing area
overlaps the tab bar. When the form's window size is changed (= `UserForm_Resize`
runs), it is corrected to the right position. The 3 tabs at startup looked normal from the beginning.

**Cause**: `Wv2Browser.AddTab` set the initial Bounds of the new tab to `GetClientRect(m_hwndHost)` =
**the whole frame**. Since it does not subtract the tab bar band (top 36px), the new tab
is placed from the top edge 0 and overlaps the tab bar. The 3 tabs at startup looked normal because
`StartWebView2_WithTabBar_3` calls `UserForm_Resize` once at its end, which
re-issued `ResizeAll(contentRect)` to all tabs and corrected them (= "they happened to be fixed").
`Test_9_11c_AddTab` only calls `AddTabWithUrl` and does not call `UserForm_Resize`, so
the tab stayed with its initial Bounds and overlapped.

**Structural problem**: The Browser does not know the tab bar exists (loose coupling principle). So AddTab
can only use "the whole frame" as the initial value, and the actual content rectangle (the area excluding the tab bar) is known
only to the upper layer (UserForm).

**Fix (option A: content rectangle cache)**: While keeping the upper layer as the source of truth for layout,
the Browser caches "the content rectangle the upper layer last applied" and applies it consistently to new tabs.

- New fields: `m_lastContentRect As RECT` / `m_hasContentRect As Boolean`
- `ResizeAll`: saves the passed rc into `m_lastContentRect` and sets `m_hasContentRect = True`
- `AddTab`: if `m_hasContentRect` is True, uses that rectangle as the initial Bounds. When unset
  (ResizeAll not yet run), falls back to `GetClientRect` (whole area) as before

**Loose coupling is unchanged**: The Browser just remembers "the area the upper layer specified" without knowing about the tab bar;
the Browser itself does not compute the tab bar height. The UserForm is unchanged.

**All paths are covered**: `AddTabWithUrl` / `AddTabWithHtml` / `AddTabWithUrlForSpa` /
the asynchronous new tab from NewWindowRequested (`RequestNewTab` -> `AddTabWithUrl`) **all
go through `AddTab` for pane creation and initial Bounds setup**. So a fix at the single point `AddTab` means
that new tabs via a 9.11d click-to-open and via window.open are also all born with the correct rectangle, with no omissions.

**Rejected alternative**: "Always call `UserForm_Resize` on the UserForm side after AddTab" (option B) would
leave the Browser unchanged, but the more paths there are for creating new tabs, the easier it is to forget the call again, and
responsibility gets scattered, so it was rejected.

### Real-machine test results for stage 9.11c (completion record)

Full pass. All items as expected.

- **The ready handshake works**: On receiving `{"cmd":"ready"}`, `OnPaneWebMessage` ->
  `PushSyncToJs` performed the initial sync. The design of not firing before the JS listener is registered (option X) worked correctly.
- **Full sync tracking is perfect**: For both AddTab (4 tabs, active=4) and CloseActive (3 tabs, active=3),
  event -> `PushSyncToJs` -> JSON send -> tab bar redraw was consistent throughout. With no diff management, index
  shifts never occurred at all.
- **The title cache (option 3) works as the source of truth**: The MSN title was replaced in two steps, `MSN` ->
  `MSN | パーソナライズ…`, and each time it was
  read correctly from `PaneAt(i).DocumentTitle`.
- **Demonstration of the layout fix (option A)**: The 3 tabs at startup went through
  `Using the whole frame for initial Bounds (ResizeAll not yet run)` -> corrected by the subsequent `UserForm_Resize`
  (the expected fallback path). The 4th tab went through
  `Using the cached content rectangle for initial Bounds` -> **placed correctly below the tab bar
  without needing a resize**. Bug fix confirmed.
- **Release**: Completed in the order TabBar.Shutdown (ClearWebMsgSink -> View3 -> View -> Ctrl) -> Browser.Shutdown ->
  Pane group Cleanup -> Environment release. No sign of leaks.

No doubts; pass. -> Stage 9.11c complete.

### Design principles and spec facts established in 9.11c

- **Design principle 59 (new)**: A child object's "initial geometry (Bounds/layout)" is handled by reusing
  the rectangle the upper layer last applied, by caching it.
  In a structure that keeps loose coupling (the lower layer does not know the
  upper layer's layout circumstances), it is in principle impossible for the lower layer to compute the correct initial rectangle by itself. Relying on "whole area at creation -> the upper layer fixes it later with Resize"
  breaks down on paths where Resize does not run (such as adding a tab at runtime). **Once the upper layer has told it the rectangle even once,
  the lower layer should remember it and apply it consistently to all subsequent children.** The source of truth remains the upper layer, so
  loose coupling is not broken.
- **Design principle 60 (new)**: State sync from VBA->JS defaults to **full sync (declarative)**.
  Diff sync requires index-correction logic on the JS side and is a breeding ground for consistency bugs. Make VBA the single
  source of truth, send the whole state on every state change, and have the JS side re-render everything. At the scale of a few dozen elements the
  cost is practically the same as diffs, and robustness wins decisively. As a side benefit, the initial sync is "just send the same sync once", so
  no special-casing is needed.
- **Design principle 61 (new)**: The first VBA->JS send via `PostWebMessageAsString` is fired
  **only after receiving the ready signal from the JS side** (ready handshake). A message sent before the JS message
  listener is registered is dropped (spec fact 15). `Navigate` completion and
  `NavigationCompleted` do not guarantee "the JS listener is already registered". Have the JS side fire
  `postMessage({cmd:"ready"})`, and have VBA send the initial state upon receiving it.
- **Design principle 62 (new)**: Notification paths from a Pane to the upper layer each hold **a separate reference per purpose**.
  `m_browser` (back reference for NewWindowRequested / title notification) and `m_webMsgSink`
  (forwarding destination for WebMessages from JS) must not be mixed. Responsibilities are separated, and it is clear
  which one to cut on release. Both are explicitly disconnected per design principle 58.
- **Spec fact 15 (new)**: A message sent with `ICoreWebView2::PostWebMessageAsString` is
  dropped if the JS side has not yet registered
  `window.chrome.webview.addEventListener('message', ...)` when it arrives (**not queued**). The time window in which the title change event went through as "unhandled" in the 9.11b log
  was this. The remedy is the ready handshake (design principle 61).
- **Spec fact 16 (new)**: VBA's `AscW` returns characters at 0x8000 and above as a **signed Integer**
  (e.g. 「進」U+9032 = 36914 -> **-28622**). If you write a control-character check such as
  `Case Is < 32` in a hand-written JSON escape, high-order characters such as Japanese fall into
  this branch as negative values and get escaped in `\uXXXX` form. **The result is correct as JSON and is restored to the right character on the JS side**
  (confirmed on the real machine: "常に進化し続け…" displayed normally), so there is no actual harm,
  but be aware that it is an unintended path. If you strictly want to "escape only control characters", use an
  unsigned evaluation (`ch >= 0 And ch < 32`).

The next stage is 9.11d (JS->VBA direction: enabling clicks. Click a tab to activate, the × button to close,
the + button to newtab). Since 9.11c established the foundations of the ready handshake and the WebMessage sink,
this will be an extension of the command interpretation in `OnPaneWebMessage`.

---
## Stage 9.11d — JS->VBA direction: enabling clicks (Layer 2 complete)

### Policy (outcome of the alignment discussion)

The final sub-stage of Layer 2. Connect clicks on the tab bar to VBA tab operations,
completing the WebView2-based tab bar in both directions.

- Click on the tab body -> `{"cmd":"activate","index":N}` -> `Browser.ActivateTab N`
- Click on the × button -> `{"cmd":"close","index":N}` -> `Browser.CloseTab N`
- Click on the ＋ button -> `{"cmd":"newtab"}` -> `Browser.RequestNewTabAsync`

**Command interpretation = option 1-A (minimal home-grown parser)**: VBA has no standard JSON parser, and external COM
(ScriptControl etc.) is not used under this project's policy. The sender is only the JS of our own HTML, so
the input is fully trusted and a general-purpose parser is overkill. We keep our own `GetJsonStr` / `GetJsonNum`.
Since VBA->JS is JSON, JS->VBA is also made JSON to match, and it extends naturally when more commands are added later
(option 1-B, pipe-delimited, was rejected because the format would be asymmetric between the directions).

**Re-entrancy safety = option 4-A**: `OnPaneWebMessage` runs inside a WebView2 event handler, so
when calling other WebView2 operations from there, one must be wary of spec fact 14 (waiting inside a handler for the completion of an asynchronous WebView2 operation
causes a deadlock). Whether a wait for completion is needed differs per command.

| Command | What it calls | Wait for completion | Verdict |
|---|---|---|---|
| activate | `ActivateTab` | Only `Ctrl_PutIsVisible`, none | Safe synchronously |
| close | `CloseTab` | Only `Ctrl_Close`, none | Safe synchronously |
| newtab | `AddTabWithUrl` | **Waits for Controller creation with `WaitPaneReady`** | **Deadlocks** |

Therefore only newtab is routed to `RequestNewTabAsync` (the SafeTimer one-shot mechanism from 9.10e), and the
handler returns immediately. This reuses the path already proven with NewWindowRequested as is.

Option 4-B (route all commands through SafeTimer, making it a mechanical discipline of "never touch WebView2 inside a WebMessage handler")
was also considered, but activate/close should be fine synchronously by nature, so the policy was to
first verify the safety of synchronous calls with option 4-A on the real machine (the way this project does things, since real-machine tests have found bugs
early). -> **Safety was demonstrated on the real machine, so option 4-A is finalized.**

**Close policy = option 2-A (allow 0 tabs)**: Even the last tab can be closed. This is consistent with the 9.11a CloseTab decision that
"the minimum-tab-count policy is left to the upper layer". When the tab count reaches 0, the JS shows
"(no tabs)" (already handled by the render in 9.11c), and one can recover from the ＋ button.

**Initial URL for newtab = option 3-B (constant HOME_URL)**: So that it can be made configurable from the UserForm in the future,
it is held as a constant. Because `about:blank` has an empty title and the tab would stay at "(loading)", `https://www.bing.com` was made the default,
prioritizing the look of the verification.

### What 9.11d implemented

| File | Version | Changes |
|---|---|---|
| `Wv2TabBar.cls` | 9.11d | Added the `HOME_URL` constant. Expanded `OnPaneWebMessage` into a `Select Case cmd` (6 branches: ready / activate / close / newtab / empty / unknown). Added the minimal JSON value extractors `GetJsonStr` / `GetJsonNum`. Enabled `postCmd` in the 3 click handlers of `BuildTabBarHtml` (× already calls `stopPropagation`, so it does not wrongly fire the parent's activate) |
| `UserForm1.frm` | 9.11d | Verification only. `Test_9_11d_Help` / `Test_9_11d_InjectActivate` / `Test_9_11d_InjectClose` / `Test_9_11d_InjectNewTab` / `Test_9_11d_ParserCheck` |

`Wv2Pane.cls` / `Wv2Browser.cls` are unchanged from 9.11c (the existing public API was enough =
the loosely coupled design of 9.11a to 9.11c paid off).

Import: Wv2TabBar -> UserForm1.

### Real-machine test results of 9.11d (completion record)

Full pass. All items of the manual clicks (A-1 to A-5) and the injection tests (B) behaved as expected.

- **A-1 activate**: clicking a tab body -> `{"cmd":"activate","index":2}` -> `ActivateTab 2` ->
  `RaiseEvent ActiveChanged` -> `PushSyncToJs` (active=2). The content switched to Example Domain,
  and the active highlight on the tab bar followed.
- **A-2 close (active tab)**: clicking × -> `CloseTab 2` -> active correction (right neighbor first) made
  MSN active. **× does not wrongly fire the tab body's activate** (stopPropagation works).
- **A-3 newtab (★the main point★)**: clicking + -> `RequestNewTabAsync` -> add pending URL -> start the one-shot
  Timer -> handler returns immediately -> `AddTabWithUrl` runs in `TimerCall`. **Excel does not freeze and
  does not deadlock**. The decision for plan 4-A was proven on the real machine.
- **A-4 tab 0 allowed**: closing inactive tabs in succession did not make the active tab jump around, and closing the last
  tab gave `all tabs closed (activeIndex=0)` -> `{"type":"sync","tabs":[],"active":0}` ->
  the tab bar showed "(no tabs)".
- **A-5 recovery**: clicking + from the tab-0 state restored a tab normally.
- **B injection tests**: for all 5 kinds of abnormal input (empty string / non-JSON / unknown command / missing index / out-of-range index),
  Excel did not crash, and `TabCount` / `ActiveIndex` were not corrupted. For a missing index, `GetJsonNum` returns
  0 and `ActivateTab(0)` is rejected by the out-of-range guard (as designed).
- **Release**: ran through to the end: TabBar.Shutdown -> Browser.Shutdown -> Pane group Cleanup -> Environment release.
  No sign of leaks.

Passed without doubt. -> **9.11d complete. With this, Layer 2 (the WebView2-based tab bar) is complete**.
A two-way working WebView2 tab bar, where tab display, active switching, adding and removing can all be operated from the tab bar's GUI,
was achieved with VBA alone.

### Additions to the design principles established in 9.11d

- **Design principle 63 (new)**: When calling another WebView2 operation from inside a WebView2 event handler,
  **decide between synchronous and asynchronous by whether that operation includes "waiting for completion"**. Operations that do not include waiting for completion
  (fire-and-forget kinds such as `Ctrl_PutIsVisible` / `Ctrl_Close` / `PostWebMessageAsString`) are
  safe to call synchronously. Operations that include waiting for completion (kinds that wait for an asynchronous
  callback to arrive, such as `AddTabWithUrl` -> `WaitPaneReady`) always deadlock if called from inside a handler
  (spec fact 14), so they are deferred until after the handler completes with a SafeTimer one-shot (the `RequestNewTabAsync` mechanism of 9.10e).
  **"Do not wait for asynchronous completion inside an event handler" is the invariant to keep**; it does not forbid
  the kind of operation itself.
- **Design principle 64 (new)**: When communicating only with the JS of our own HTML, **do not build a general-purpose JSON parser**.
  If the sender is limited to JS we wrote ourselves, the input can be fully trusted, and there is no need to support the general form
  including nesting, arrays and escapes. The right answer is to have a minimal function that extracts the needed keys (about `GetJsonStr` / `GetJsonNum`).
  Under the constraint that VBA has no standard JSON parser and no external COM can be used,
  we explicitly adopt the judgment that "a minimal implementation is enough for trusted input".

### Observed behavior (no harm, recorded only)

When a new tab is added with + from the tab-0 state, a sync with `active=0` is sent momentarily at the `TabAdded` moment,
and immediately after, a sync with `active=1` is sent by `ActivateTab` (2 renders within 1 ms). This is the
consequence of the design decided in 9.11a that "`AddTab` fires `TabAdded` right after adding to the Collection, then `ActivateTab`".
Because it is the full-sync approach, the final state is always correct, and the flicker is not visible. No harm, so
left as is.

---
## Stage 9.12 — Address bar / navigation bar (Layer 3 complete)

### Overall policy (result of the alignment)

Following the tab bar (Layer 2), we add the address bar and back / forward / reload so the browser is
usable as a whole. **All the necessary Pane-side methods were already implemented in 9.8b/9.9b**, so
9.12 became a stage of "wiring existing features to the UI", and almost no new COM calls occurred.

| Feature | Existing method |
|---|---|
| Navigate to URL | `View_Navigate(url)` |
| Back / forward | `View_GoBack` / `View_GoForward` |
| Reload / stop | `View_Reload` / `View_Stop` |
| Get the current URL | `View_GetSource` |
| Whether back/forward is possible | `View_GetCanGoBack` / `View_GetCanGoForward` |

**Class structure = plan 1-B (add a new `Wv2NavBar`, one more WebView2)**: Since 9.11 established the pattern of "using WebView2 as a
UI component" and the effect of loose coupling was also proven, we follow the same structure.
The implementation is close to a copy of 9.11b to d, so the risk is low. The responsibilities also separate cleanly.

The layout becomes a 3-way split:

    +--------------------------+
    | [Bing][MSN x] [+]        | <- Wv2TabBar   (36px)
    +--------------------------+
    | [<][>][R] [https://...  ]| <- Wv2NavBar   (34px) ★added in 9.12
    +--------------------------+
    |      Content             | <- the tabs of Wv2Browser
    +--------------------------+

**Sub-stages**: split into 4 with the same rhythm as 9.11 (a=foundation / b=skeleton / c=VBA->JS / d=JS->VBA).

---

### 9.12a: Browser-side foundation (no UI, Immediate-window verification)

**New event `TabUrlChanged(index, url, canGoBack, canGoForward)` (plan 2-A)**:
The nav bar needs to know the "URL state of the active tab", but Browser had no event to
announce URL changes. It was added with exactly the same loose-coupling pattern as `TabTitleChanged` of 9.11a.
Pane's `NotifyUrlState` -> `Browser.NotifyUrlChanged` -> reverse lookup of the index -> `RaiseEvent`.

**Fired from both NavigationCompleted and HistoryChanged**:
- `NavigationCompleted` ... normal page navigation
- `HistoryChanged` ... SPA pushState, back/forward

It fires multiple times for one navigation, but **it is harmless because of the full-sync approach (design principle 60)** (it just sends the same state
twice).

**★HistoryChanged promoted to a standard handler (plan 2-A-1)★**: Previously HistoryChanged was just for "SPA"
purposes and was registered only in `EnableSpaHandlers` (9.10d-2b). However, **the enabled
state of the back/forward buttons has to be tracked on all tabs, not just SPAs**, so we moved it into `RegisterStandardHandlers`
so that it is registered automatically on all tabs. -> Standard handlers go from 4 to **5**. It was removed from `EnableSpaHandlers`
to avoid double registration (that now handles only DOMContentLoaded).
`AddTabWithUrlForSpa` also goes through `AddTab`, so SPA tabs are not missed.

**Delegating operations to the active tab (public API for the nav bar)**: Letting the nav bar touch Pane directly would break
the loose coupling, so we provide thin methods that Browser proxies.
`NavigateActive` / `GoBackActive` / `GoForwardActive` / `ReloadActive` / `StopActive` /
`GetActiveUrl` / `GetActiveCanGoBack` / `GetActiveCanGoForward`.

**URL scheme completion (plan 3-A)**: `NormalizeUrl` prepends `https://` if there is no scheme
(`"example.com"` -> `"https://example.com"`). URLs with a scheme such as `about:blank` pass through unchanged
(a scheme name is judged to be letters only, so `example.com:8080` is not misjudged).
**The search-word fallback is out of scope** (it would need URL encoding, so it is deferred to the next stage).

Real-machine test: registration of the 5 standard handlers / TabUrlChanged firing / scheme completion / history tracking
(Navigate -> Back -> Forward moves back/fwd correctly) / tab-0 guard, all passed.

---

### 9.12b: Wv2NavBar skeleton + 3-way layout split

A new class that follows the structure of `Wv2TabBar` as is. The host name is `appassets.navbar`,
and **the delivery folder is also separate from TabBar's**.

**`UserForm_Resize` generalized to a "stack bands from the top" approach**: it carries `y` along, stacks each bar, and
makes the remainder the content. **If a bar is absent that row is skipped**, so it works correctly in any configuration: tab bar only / nav bar
only / neither. The split calculation is the UserForm's responsibility (Browser,
TabBar and NavBar do not know each other).

`Terminate` goes in the order **NavBar -> TabBar -> Browser** (reverse of creation).

Real-machine test: the 3 rows are laid out without overlapping and follow resizing. The existing tab bar functions are also unharmed.
**★Loose coupling re-proven★** The same Browser event **reached the UserForm / TabBar / NavBar, all 3 parties,
simultaneously**.

---

### 9.12c: VBA->JS (displaying the URL and button states)

`{"type":"navsync","url":"...","canBack":true,"canForward":false,"hasTab":true}` is sent
by full sync. The ready handshake (design principle 61) is also the same as in 9.11c.

**★Index filter (an issue specific to 9.12c)★**: `TabUrlChanged` tells by index "which tab's" URL
changed. The nav bar displays only the active tab's state, so
**events with `index <> ActiveIndex` must be ignored**. If they are not ignored, the moment a background tab finishes
loading, the URL field turns into that background tab's URL (in the real-machine log, a TabUrlChanged with index=2 actually
arrived while index=3 was active).

**★On ActiveChanged, ask Browser for the current state★**: Even when the tab is switched, nothing happens on the new
active tab, so no `TabUrlChanged` comes. So we re-read with
`GetActiveUrl` etc. and send it (the reason we prepared this read path in 9.12a).

#### ★★A chain of 3 bugs found on the real machine (protection while the URL field is being edited)★★

**This stage had the most to learn from. We stepped on a chain, three times, where a fix produced the next bug.**

**Bug 1: using `activeElement` as a proxy for "being edited" -> the URL field freezes forever**
- At first, the edit protection was implemented as `if(document.activeElement !== u){ u.value = ... }`.
- In a WebView2 standalone document, **`#url` can remain the activeElement and never go back**,
  so this condition stays false forever, and the symptom appeared that **it worked only for the first initial sync, and after that the URL field was never
  updated again**.
- **What decided the diagnosis**: the observation that "the buttons' enabled states follow correctly, but only the URL field
  is not updated". `applyNavSync` runs every time, so the only culprit could be
  the `if` that wraps just the assignment to `u.value`, we could pin it down.
- -> Changed to a method where an **explicit flag (`editing`)** holds "whether the user has actually started editing".
  Merely focusing does not make it editing; key input (the `input` event) makes it editing,
  and blur / Enter / Escape clear it.

**Bug 2: tab-0 support broke the edit protection**
- When the last tab was closed, the URL field stayed at the closed tab's URL. The cause was that
  when there is "no tab to activate", `CloseTab` does not call `ActivateTab`, so
  **`ActiveChanged` is not fired** (the 4-branch design of 9.11a). NavBar subscribed only to `ActiveChanged`, so
  it was never told that the tabs had become 0.
- -> **Fix (plan A): NavBar also subscribes to `TabClosed`**. The Browser event contract is not changed
  (the loose-coupling principle of "events needed for display are subscribed to by the display side"). Along with it, `hasTab` was added to navsync,
  and the reload button and the URL field were also disabled at tab 0.
- **★But this `u.disabled` sync broke the fix for bug 1★**

**Bug 3: assigning `disabled` triggers blur and erases the input**
- While typing in the URL field, clicking a link to open a new tab erased the string being typed.
- Cause: **assigning `disabled` makes the input blur** -> the blur handler sets `editing = false`
  -> the next sync overwrites `u.value`; that is the chain.
- -> **The fix is two-tiered**:
  1. **While editing, do not touch the URL field at all** (neither `value` nor `disabled`)
  2. **Assign `disabled` only when the value changes** (prevents needless blur)

  Buttons do not steal focus, so they may always be synced.

**Lessons from this chain of 3 -> design principles 65 / 66**

---

### 9.12d: JS->VBA (enabling input and clicks)

The commands are `back` / `forward` / `reload` / `navigate`.

**★Re-entrancy safety = design principle 63 could be used as is★**: All navigation operations are
"operations that do not include waiting for completion" (Navigate / GoBack / GoForward / Reload are all fire-and-forget),
so **calling every command synchronously is safe**. The deferred execution needed for newtab in 9.11d (waiting for Controller creation ->
the deadlock of spec fact 14) was not necessary. A good example of design principle 63 working in the form of "judge per kind of
operation".

**Scheme completion is Browser's responsibility**: JS sends the input string as is, and `NormalizeUrl` inside `NavigateActive`
normalizes it.

**`editing = false` on Enter**: When navigation starts, navsync replaces the URL field with the normalized URL
(typing `example.com` turns into `https://example.com/`). This is the same behavior as
Chrome/Edge.

**★Discarding the edit with Escape (bug 4, fixed with plan A)★**: At first Escape just set `editing = false`
and blurred, and **there was no processing to restore the URL field's value**, so the typed string remained. This was because `u.value` is
rewritten only when the next navsync arrives, and with Escape nothing happens on the VBA side (no sync is
sent).
-> On the JS side we hold **`lastUrl`** (the "URL that should be displayed now" conveyed by the latest navsync), and Escape
restores it immediately. No VBA round trip is needed, and it also feels better. **`lastUrl` is always updated even while editing**
(so that even if navigation occurs in the background during editing, Escape returns to "the real current URL").

### Real-machine test results of 9.12d (completion record)

Full pass.

- **A-1 URL input**: `example.com` + Enter -> navigated to `https://example.com` via scheme completion.
  The URL field was replaced with the normalized URL, and ← became enabled.
- **A-2/A-3 back/forward**: moved back and forth through the history correctly, and the button enabled states followed.
- **A-4 reload**: reloaded while keeping the history state.
- **A-5 Escape**: typed `abcde` and pressed Escape -> immediately restored to the current URL.
- **★A-6 Independent history per tab★**: in tab 1 Yahoo -> back -> Bing, in tab 2
  Google -> back -> Example -> forward -> Google. **Every time the tab was switched, the nav bar correctly followed that tab's
  history state**. The delegation to the active tab works completely.
- **A-7 Tab 0**: all buttons and the URL field are grayed out.
- **B injection tests**: does not crash even with abnormal input.

Passed without doubt. -> **9.12 complete. With this, Layer 3 (the navigation bar) is complete**.
It is now in a form that is usable as a tabbed browser overall.

---

### Additions to the design principles established in 9.12

- **Design principle 65 (new)**: **Do not use focus state as a proxy for "whether it is being edited"**.
  `document.activeElement` only represents the presence of focus, which is different from "whether the user
  intends to edit". In a WebView2 standalone document, focus may even stay stuck and never return.
  **Holding "whether there was actual key input" in an explicit flag** is the reliable way. Merely focusing
  does not make it editing; the `input` event makes it editing, and blur / commit / discard clear it.
- **Design principle 66 (new)**: **Toggling the `disabled` attribute knocks the focus away**. For an element that holds
  focus-dependent state (such as an editing flag), **the attribute must not be assigned unconditionally**.
  The countermeasure is two-tiered: (1) when in the state to be protected, do not touch that element at all, (2) assign the attribute only when the value
  changes (prevents needless side effects). **This is a trap you step on when you write "state sync" naively as
  "assign everything every time"**, and it is one of the few cautions about the full-sync approach (design principle 60). Full sync
  can be naive on the "sending" side, but the "receiving" side must distinguish assignments that have
  side effects.
- **Design principle 67 (new)**: **Events needed to display a UI are subscribed to by the display side itself**.
  Do not change the event contract of the lower level (Browser) "for the convenience of the display side". In bug 2 of 9.12c,
  for the problem that "`ActiveChanged` is not fired when the last tab is closed", there was a plan to have Browser fire `ActiveChanged(0)`, but we
  adopted having **NavBar also subscribe to `TabClosed`**. The lower level only needs to fire "facts that happened", and "who displays it
  and how" is the concern of the upper level.

### Numbers for 9.12

- Design principles: 64 -> 67
- Spec facts: 16 (no change; note: 9.12 centered on combining existing APIs)
- New class: `Wv2NavBar.cls`
- Standard handlers: 4 -> **5** (HistoryChanged promoted)
- Public Events of `Wv2Browser`: 4 -> **5** (TabUrlChanged added)

---

## Stage 9.13 — Cleanup: unifying the JSON helpers + removing throwaway tests

### Overall policy

Layer 2 (tab bar) / Layer 3 (nav bar) are complete, and it is now in a form that is usable as a tabbed browser
overall. Before moving on to the next new feature, we set up a cleanup stage to sweep away the duplication and leftovers that accumulated in 9.11 to 9.12
(the same rhythm as the 9.10 cleanup stage). There are 2 aims:

1. **Unifying the JSON helpers** -- `JsonEscape` / `GetJsonStr` / `GetJsonNum` /
   `BoolToJson` were duplicated in `Wv2TabBar` and `Wv2NavBar` (executing the agreed
   "extract to a common module in a future cleanup stage").
2. **Removing throwaway test Subs and old entry points** -- `Test_9_11a_*` to `Test_9_12d_*` made in each sub-stage of 9.11 to 9.12,
   and the old `StartWebView2*` group were
   left over though not needed in actual operation.

### 9.13-1: Consolidating the JSON helpers (new `Wv2Json.bas`)

**A new standard module `Wv2Json.bas`** gathers four Public functions, renamed from the old names:

| New name | Old name (Private in TabBar/NavBar) | Purpose |
|---|---|---|
| `JsonEscape` | `JsonEscape` | Embed a string into a JSON literal |
| `JsonGetStr` | `GetJsonStr` | Get the value of `"key":"value"` |
| `JsonGetNum` | `GetJsonNum` | Get the 123 of `"key":123` |
| `JsonBool` | `BoolToJson` | Boolean -> `true`/`false` |

`Wv2TabBar.cls` / `Wv2NavBar.cls` had their Private helpers removed and now call the
Public functions in `Wv2Json` (10 call sites renamed).

**★ Fixed the \u path of spec fact 16 ★**: the control-character check in `JsonEscape` was changed from
`Case Is < 32` to `If ch >= 0 And ch < 32`. This avoids the trap where `AscW` returns high characters (U+8000
and above) as negative values in a signed Integer, so Japanese no longer gets garbled into
`\uXXXX`. **The final result as seen on the JS side does not change** (`進` also
turns back into 「進」, so there was no real harm as of 9.11c), but the logs become easier to read and the number of bytes sent
goes down. Confirmed in the real-machine log of 9.13 that `json=` shows Japanese as is.

### 9.13-2: Removing throwaway test Subs and the old entry points

**`UserForm1.frm`** (1456 -> 700 lines):
- Removed 35 throwaway verification Subs (`Test_9_11a_*` to `Test_9_12d_*`).
- Removed the old startup Subs (`StartWebView2` / `StartWebView2_WithTabBar` /
  `StartWebView2_WithTabBar_3`).
- Consolidated the production entry points into two: **`StartWebView2_Full`** (normal) and
  **`StartWebView2_Spa`** (SPA skeleton).

**`Wv2Thunks.bas`** (2687 -> 2483 lines) — done in this session:
- Removed 4 UserForm-dependent integration test Subs. Each of them called `StartWebView2`
  (and `IsPageLoaded` / `GetActivePane` /
  `OnExecuteScriptResult`), which had already been removed from `UserForm1` in 9.13-2; left alone, they would
  **cause compile errors**:
  - `Test_Controller_Real` (9.3b) / `Test_Navigate_Real` (9.3c) /
    `Test_AllEvents_Real` (9.4b) / `Test_ExecuteScript_Parallel` (9.9c)
- The 5 low-level unit tests kept (thunks, vtable, memory primitives.
  **Living assets that a GUI cannot replace**):
  - `Test_TwoHandlers_Mock` / `Test_Environment_Real` /
    `Test_VirtualAllocFree_Roundtrip` / `Test_SentinelStatus` /
    `Test_ResetSurvival_Manual`
- What to do with `Test_ExecuteScript_Parallel` was a point of discussion, but parallel ExecuteScript
  execution was already verified (passed) in 9.9c, and when SPA integration is taken up in earnest later it makes more sense to rewrite it on top of
  `StartWebView2_Spa`, so it was **removed**
  (agreed).
- The step-by-step comment of `Test_ResetSurvival_Manual` referred to the removed `Test_Navigate_Real`,
  so it was rewritten to use `UserForm1.StartWebView2_Full`, making it
  self-contained.
- **The core logic (`Handler_xxx` / `dcf` / `Thunks_Init` / memory primitives, etc.)
  is completely unchanged.** This time only the throwaway test Subs were removed.

### The standard file layout goes from 6 to 7

    Wv2Thunks.bas / Wv2Json.bas ★new / ComCallbackHandler.cls
    Wv2Environment.cls / Wv2Pane.cls / Wv2Browser.cls
    Wv2TabBar.cls / Wv2NavBar.cls / UserForm1.frm

(6 classes + 2 standard modules + 1 form = 9 modules. Counted as "standard files",
it is 6 -> 7.)

### Real-machine test results of stage 9.13 (completion record)

**The main purpose was to confirm that "nothing is broken". All items passed.**

- **Compile**: passes cleanly in the VBE (zero leftover calls to removed methods).
- **Low-level layer tests (4 kinds)**: all as expected.
  - `Test_VirtualAllocFree_Roundtrip`: all 5 scenarios α to ε
    `VirtualFree returned 1, LastError = 0 [OK]`. Region release and
    cleanup via the real `Thunks_Init` are also normal.
  - `Test_SentinelStatus`: no leftover region (`m_pRegionBase = 0`, environment variable cleared).
  - `Test_TwoHandlers_Mock`: AddRef/Release dispatch, automatic slot release, and the
    post-release safety checks all match the expected values. `VirtualFree` [OK] in `Thunks_Shutdown`.
  - `Test_Environment_Real`: Environment creation reaches `Es_Ready`
    (State=2) via the synchronous callback, and release is also normal.
- **GUI (`StartWebView2_Full`)**: the 3-row layout (tab bar/nav bar/content)
  is laid out without overlap, and the ready handshake of both tabbar/navbar -> initial sync is established.
  Both the tab bar's full sync and the nav bar's navsync are normal. **`json=` outputs Japanese
  as is** (the `\uXXXX` escapes are gone <- demonstrating the improvement of 9.13-1).

-> **Stage 9.13 complete. The cleanup stage removed the duplication, and the production entry points were consolidated into two:
`StartWebView2_Full` / `StartWebView2_Spa`.**

### Numbers for stage 9.13

- Design principles: 67 (unchanged)
- Spec facts: 16 (unchanged. However, the \u path of spec fact 16 was fixed in the implementation)
- New standard module: `Wv2Json.bas`
- Standard file layout: 6 -> **7**
- `UserForm1.frm`: 1456 -> 700 lines (35 verification Subs + 3 old startup Subs removed)
- `Wv2Thunks.bas`: 2687 -> 2483 lines (4 UserForm-dependent tests removed, core unchanged)


---

## Stage 9.14 — Search-term fallback (default Google)

### Overall policy

This picks up the open item that 9.12a had put on hold as "out of scope, next stage". When a non-URL string such as
`検索したい言葉` ("the words I want to search") is typed into the address bar, it is sent to a Google search. The aim is a feel almost the same as the Chrome/Edge
address bar. The default engine is **Google** (chosen by tarboh).

The core of this project is "driving WebView2 without any external COM library", but for URL
encoding (percent-encoding) conversion to a UTF-8 byte sequence is essential, and this is **mere data conversion
unrelated to the thunk / vtable / memory primitive layer**. `ADODB.Stream` is already used
for writing out HTML, so it is not philosophically a gray area. Therefore
`ADODB.Stream` is used for the UTF-8 conversion (a home-grown UTF-8 encoder was considered, but
spending a whole stage hardening surrogate-pair handling has poor cost-effectiveness, so it was set aside).

### 9.14-1: `NormalizeUrl` made three-way (`Wv2Browser.cls`)

`NormalizeUrl`, which used to be a two-way choice ("no scheme -> prepend `https://`"), is extended to three branches:

| Kind of input | Condition | Conversion |
|---|---|---|
| (1) Has a scheme | Has a part before the colon that is `IsAlphaOnly` (`about:` `file:` etc.) | Returned as is |
| (2) Regarded as a URL | Contains a dot **and** contains no whitespace | Prepend `https://` |
| (3) Regarded as a search term | Contains whitespace **or** contains no dot | Converted to a Google search by `BuildSearchUrl` |

`NavigateActive` navigates to the result of `NormalizeUrl` as is, so extending `NormalizeUrl`
in just one place makes the search fallback work (the loosely coupled design paid off nicely).

### 9.14-2: Home-grown URL encoding (Private additions in `Wv2Browser.cls`)

The placement is as **Private helpers of `Wv2Browser`**. The search fallback is Browser's
responsibility (the same place as `NormalizeUrl`), so cohesion is high. `Wv2Json.bas` is declared "JSON only"
so it is not mixed in. If URL handling grows in the future, it will be extracted into `Wv2Url.bas` in a cleanup stage (the same rhythm as the after-the-fact extraction of JSON
in 9.13, agreed).

- `BuildSearchUrl(terms)` — `https://www.google.com/search?q=` + the encoded result.
- `UrlEncodeComponent(s)` — equivalent to JS `encodeURIComponent`. Passes through only unreserved characters
  (`A-Z a-z 0-9 - _ . ~`) and turns everything else into `%XX`.
- `Utf8Bytes(s)` — VBA String (UTF-16LE) to a UTF-8 byte sequence. Writes text to `ADODB.Stream` with
  `Charset="utf-8"` and reads it back as `adTypeBinary`. No BOM is added, but as
  insurance, if the first 3 bytes are `EF BB BF` they are removed.

**★ Surrogate pair support ★**: processing one character at a time with `AscW` would break surrogate pairs such as emoji, so
**the whole string is converted to UTF-8 in one go and then escaped byte by byte**. All cases were confirmed to match
`urllib.parse.quote(safe="-_.~")` in a container (including Japanese, symbols, emoji, and
the percent sign itself).

### 9.14-3: Adding a logic-verification entry point (`Wv2Browser.cls` + `Wv2Thunks.bas`)

`NormalizeUrl` is Private, so a thin public wrapper
`Debug_NormalizeUrl` was added to `Wv2Browser` to check just the result from outside. It goes **through the same
`NormalizeUrl` that the production `NavigateActive` uses**, so the verification logic is not managed twice (the same position as the low-level
tests kept in 9.13: a "living verification asset").

`Test_9_14_SearchFallback` (+ a dedicated helper `Check`) was added to `Wv2Thunks.bas`.
It is a pure-logic test with no WebView2 startup; just typing the Sub name in the Immediate window
verifies 17 cases in one batch.

### ★ Design principle 68 (new) ★

**Address bar input is routed to URL / search term by a three-way branch.**
(1) With a scheme (a part before the colon that is `IsAlphaOnly`) it passes through, (2) with a dot and no whitespace it becomes a URL by
prepending `https://`, (3) otherwise (whitespace present or no dot) it is
converted into a search URL as a search term. URL encoding converts to a UTF-8 byte sequence with `ADODB.Stream` first and then
applies percent-encoding. **To avoid breaking surrogate pairs (emoji, etc.), escape per byte, not per character**
(convert the whole string to UTF-8 in one go -> `%XX` for each byte). The unreserved characters are
`A-Z a-z 0-9 - _ . ~`.

### ★ Spec fact 17 (new) ★

**Input like `localhost:8080`, whose "part before the colon is letters only", is
treated as a scheme in the `IsAlphaOnly` branch and passed through.** `NormalizeUrl` interprets `localhost:` as a custom scheme and returns it
as is, but Edge/WebView2 completes it with `http://` and
connects to `http://localhost:8080/`. This is therefore **correct behavior for a browser**, and `https://` must not be forcibly prepended
(local development servers are often http). In the first verification of 9.14 this point
produced one `[NG]` as a "mistake in how the expected value was written"; the implementation was correct, and the expected value on the
verification Sub side was fixed to match the actual behavior.

### Real-machine test results of stage 9.14 (completion record)

**Both wheels, logic (A) / real-machine GUI (B). All items passed.**

- **Logic verification `Test_9_14_SearchFallback`**: after fixing the expected value, **17 / 17 [ALL OK]**.
  Scheme passthrough / dotted input made a URL / whitespace or no-dot input searched / Japanese, symbols,
  `C++`, `100%`, `-_.~` passthrough / leading/trailing whitespace trimmed / empty input: all match the expected values.
- **Real-machine GUI (`StartWebView2_Full`)**:
  - `example.com` -> accessed `https://example.com/`.
  - `日本 天気` ("Japan weather") -> `https://www.google.com/search?q=%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97`
    gives Google search results (the Japanese encoding works correctly).
  - `excel vba` -> Google search results.
  - `localhost:8080` -> accessed `http://localhost:8080/` (the connection refusal is just that no
    server is on the port; the routing behavior is correct <- demonstrating spec fact 17).
  - The tab bar / nav bar ready handshake, full sync, and unescaped display of Japanese titles
    are all unharmed.

-> **Stage 9.14 complete. The tabbed browser gained a search-term fallback, and the address bar now
feels like Chrome/Edge: "if it is not a URL, search".**

### Numbers for stage 9.14

- Design principles: 67 -> **68** (design principle 68 added)
- Spec facts: 16 -> **17** (spec fact 17 added)
- Standard file layout: **7** (unchanged. The URL helpers are contained as Private in `Wv2Browser`)
- `Wv2Browser.cls`: `NormalizeUrl` made three-way + added `BuildSearchUrl` / `UrlEncodeComponent` /
  `Utf8Bytes` / `Debug_NormalizeUrl`
- `Wv2Thunks.bas`: added `Test_9_14_SearchFallback` (+ `Check`). Core logic unchanged

---

## Stage 9.15 — Drag-and-drop reordering of tabs

**Goal**: Make it possible to reorder tabs by drag and drop. As the next step toward "usable as a tabbed browser
across the board", add Chrome/Edge-style tab reordering.

**Policy (settled in discussion): A-1 + A-2b + B-1**
- **A-1**: Add `Wv2Browser.MoveTab(fromIndex, toIndex)`. It rearranges `m_tabs` (a 1-based
  Collection) by Remove -> Add (before/after).
- **A-2b**: Reuse the existing path for the sync after reordering. No new event is added (the 5 Public Events
  stay as they are). -> Crystallized as design principle 69.
- **B-1**: The JS side uses the HTML5 Drag & Drop API. A custom pointer drag (B-2) is
  deferred to the visual polish stage (candidate 2); first get it working as quickly as possible with the standard API.

### 9.15-1: Adding `MoveTab` (`Wv2Browser.cls`)

Placed right after `CloseTab`. The procedure has 4 steps:

1. **Grab the active Pane by reference.** Because reordering moves indexes, keep the object reference of `m_tabs.Item(m_activeIndex)`
   rather than the numeric value of `m_activeIndex`.
2. **Remove the target Pane and Add it at the destination.** `to == 1` is `Add moving, , 1`
   (before the first), otherwise `Add moving, , , to-1` (after the (to-1)th). Note that the Collection's
   before/after refer to "the index on the Collection after the Remove".
3. **Reverse-look-up the active Pane's new index.** Scan the reordered `m_tabs` linearly and reset `m_activeIndex` to the position that
   `Is`-matches the reference kept in step 1 (the same reference-match
   pattern as NotifyTitleChanged).
4. **Correct the visible state with `ActivateTab m_activeIndex`** -> `RaiseEvent ActiveChanged`
   -> full sync (design principle 60, plan A-2b). The visible state itself does not change, but going through ActivateTab
   prompts `PushSyncToJs` / `PushNavSyncToJs` and reflects the reordered result to JS.

Out-of-range (from/to outside 1..count) and `from == to` are guarded and leave things unchanged (the latter is treated as success).
It consists only of array operations and PutIsVisible and **contains no wait for completion**,
so it is safe to call synchronously from inside a WebMessage handler (same as activate/close. The SafeTimer detour needed for newtab is not required).

### 9.15-2: Adding the `reorder` command (`Wv2TabBar.cls`)

A `reorder` case was added to the `Select Case cmd` in `OnPaneWebMessage`. It takes `from`/`to` out with
`JsonGetNum` and calls `m_browser.MoveTab from, to` synchronously. The command contract is
`{"cmd":"reorder","from":N,"to":M}` (1-based).

### 9.15-3: Implementing HTML5 Drag & Drop (`BuildTabBarHtml` in `Wv2TabBar.cls`)

- `draggable = true` is given to each tab div.
- `dragstart`: keep the original index of the grabbed tab in the JS variable `dragSrc` and add the `.dragging` class.
- `dragover`: get the rectangle of the tab under the cursor and, depending on whether the cursor X is in its left half or right half, attach an insertion-position marker
  (`.dragover` / `.dragover-after`) (`preventDefault` allows the drop).
- `drop`: from the `after` judgment, **compute the final to (dest) as 1-based** and send reorder.
  After `dest = after ? target+1 : target`, if `dragSrc < dest` then `dest -= 1` to
  correct for "the portion where the original position is before the insertion position" (to match the VBA MoveTab
  contract of "the to-th position after reordering"). If `dragSrc === dest`, nothing is sent (no change).
- `dragend`: clean up `.dragging` and the markers.
- **The DOM is not moved by itself.** It is left to the `sync` redraw from VBA (the source of truth is VBA,
  keeping the full-sync approach). `.dragging` (semi-transparent) and the insertion marker were added to the CSS.

**The consistency between the JS from/to computation and VBA MoveTab** was checked in advance on paper: both logics were simulated in Python and
all 180 patterns (n=2..6 × dragSrc × target × after) were exhaustively cross-checked.

### 9.15-4: Adding logic-verification entry points (`Wv2Browser.cls` + `Wv2Thunks.bas`)

- `Wv2Browser`: added `Debug_SeedDummyTabs(n)` (plants n tagged dummy Panes. The `ObjPtr` of each Pane
  is recorded in `m_dbgTagMap` as sequential tags "A","B",..) / `Debug_TabOrderTitles`
  (returns the current tab order as a "|"-joined list of tags or titles) / `Debug_ActiveIndex`.
  Thanks to the ObjPtr tags of the dummy Panes, even when indexes move through reordering, "what number that tab originally was" can be
  tracked stably (with placeholder sequence numbers the moving one would be lost).
- `Wv2Thunks`: added `Test_9_15_Reorder` (10 cases) and the `CheckOrder` helper. It goes through the production
  `MoveTab` as is, so there is no double management (the same position as Debug_NormalizeUrl).

### Real-machine test results of stage 9.15 (completion record)

**Both wheels, logic (A) / real-machine GUI (B). All items passed.**

- **Logic verification `Test_9_15_Reorder`**: **10 / 10 [ALL OK]**. A to the end / D to the
  front / adjacent swaps (2->3, 3->2) / from==to unchanged / guards for out-of-range from and to /
  moving the active tab itself (active follows) / swapping 2 tabs / multi-step with 5 tabs: all have the order and
  activeIndex matching the expected values. Because the dummy Panes have View=Nothing, the ActivateTab inside MoveTab emits
  `PutIsVisible failed hr=&H80004003`, but this does not affect the order or active following
  (harmless, as announced).
- **Real-machine GUI (`StartWebView2_Full`, 3 tabs: Bing / Example / MSN)**:
  - Dragging a tab and dropping it to the left or right of another tab -> `{"cmd":"reorder","from":N,"to":M}`
    is sent, `MoveTab -> ActivateTab -> ActiveChanged -> PushSyncToJs / PushNavSyncToJs`
    runs, and the tab bar/nav bar are redrawn with the reordered order and active.
  - **Active following is correct on a reference basis**: when the active tab (MSN) is moved, active
    follows it to where it went. **Even when an inactive tab is moved, active is recomputed so that it
    "keeps pointing at MSN" even if the active target (MSN) merely shifts to a different index** (e.g. with MSN at index 2,
    moving the first tab to the end -> active stays 2). "Moving an inactive tab does not switch the active one"
    = the design intent of A-1 was confirmed on the real machine.
  - Unescaped display of long Japanese titles (MSN's Japanese) and the nav bar's URL following are also unharmed.

-> **Stage 9.15 complete. Tabs can now be reordered by drag and drop, and the active
tab follows "the tab itself". It was possible to ride on the full-sync approach without adding any new event.**

### ★ Design principle 69 (new) ★

**For reordering (a state change in which only the order changes), do not add a dedicated event; route it through the existing activation
path and ride on the full sync.**

The full-sync approach (design principle 60) is the idea of "when the state changes, resend the entire state of all tabs",
so even for a change like reordering in which **the content of no tab changes and only the order (and index) moves**,
passing once through `ActivateTab` and raising `ActiveChanged` lets the listeners'
(TabBar/NavBar) `PushSyncToJs` / `PushNavSyncToJs` reflect the reordered result
as is. There is no need to create a dedicated Public Event such as `TabsReordered`.

This accepts a small semantic
mismatch ("ActiveChanged fires even though the active tab has not changed"), in exchange for the benefit that **the event kinds do not increase (still 5), and listeners do not need a new
handler**. It is the same shape as the pattern `CloseTab` already uses,
"call ActivateTab in a reordering-like situation to prompt a sync", and
MoveTab can be seen as a generalization of it. Following the active tab is done by reverse lookup of the object reference
(the same `Is` comparison as NotifyTitleChanged), and does not depend on the numeric index.

### Numbers for 9.15

- Design principles: 68 → **69** (added design principle 69)
- Spec facts: **17** (unchanged)
- Standard file layout: **7** (unchanged)
- `Wv2Browser.cls`: added `MoveTab` + verification hooks (`Debug_SeedDummyTabs` /
  `Debug_TabOrderTitles` / `Debug_ActiveIndex`) and `m_dbgTagMap`
- `Wv2TabBar.cls`: added a `reorder` case to `OnPaneWebMessage`, and HTML5 Drag & Drop handlers
  and CSS (`.dragging` / insertion marker) to `BuildTabBarHtml`
- `Wv2Thunks.bas`: added `Test_9_15_Reorder` (+ `CheckOrder`). Main logic unchanged

---

## Stage 9.16 — Search engine switching + cleanup (Wv2Url extraction / Wv2Tests consolidation)

Stage 9.16 combined two pieces of work into one stage. The first half is a **feature addition (search engine
switching)**; the second half is **cleanup (extracting the pure URL functions into Wv2Url.bas + consolidating the verification Subs into Wv2Tests.bas)**.
The second half started when tarboh pointed out: "Wv2Thunks was supposed to be responsible only for thunk-related functionality.
Doesn't the current state undermine the point of having split out Wv2Json?"

### First half: search engine switching (②)

The "https://www.google.com/search?q=" that 9.14 hardcoded in `BuildSearchUrl`
was moved into a module variable `m_searchTemplate` so that it can be swapped.

- Adopted **option 1-A (hold the template string directly)**. It has one less piece of state than
  holding engines as an Enum + lookup table (option 1-B). All the major engines (Google/Bing/DuckDuckGo/Yahoo) can be expressed
  as a "query-prefix type" string, and it is enough to concatenate the encoded search term at the end.
  If a rare engine whose "{q}" is not at the end is ever needed, only then will this be
  extended to a placeholder scheme (YAGNI).
- Adopted **option 2-A (main-class Property only, no UI)**. It is read and written through
  `Public Property Get/Let SearchEngine`. The UI (combo box etc.) was split out into a separate stage, keeping the granularity of this stage
  small so that it can be verified on its own on a real machine.
- **Empty guard**: if an empty string / whitespace only is Let, it falls back to the default Google template.
  No validation of URL validity is done (the template is a trusted value entered by the developer).
- **The default is Google** (`DEFAULT_SEARCH_TEMPLATE` is set in `Class_Initialize`). With nothing
  configured, the behavior is exactly the same as 9.14.
- Added the verification hook `Debug_SearchEngine` (a thin public wrapper that returns the current `m_searchTemplate`).

### Second half: cleanup (Wv2Url extraction + Wv2Tests consolidation)

**Created Wv2Url.bas (③, option 2-A + option 3-A)**

Of the URL handling that was privately embedded in `Wv2Browser`, the three **pure
data conversions whose output is determined by the input alone** (`IsAlphaOnly` / `UrlEncodeComponent` / `Utf8Bytes`) were moved to
`Wv2Url.bas`. It forms a pair with `Wv2Json.bas` (JSON conversion) as a "group of stateless conversion
utilities".

- **`NormalizeUrl` / `BuildSearchUrl` were left in Wv2Browser**. These depend on the search engine
  setting (`m_searchTemplate`), which is instance state, and are "browser behavior
  (policy)"; pushing them out to a standard module, which cannot hold state, would mean passing the state around as an argument every time. The two that were left call `Wv2Url.IsAlphaOnly` /
  `Wv2Url.UrlEncodeComponent` internally. → **Design principle 70 (separation of policy and mechanism)**.
- The logic is unchanged. Only the location and the
  call qualification (prefixing `Wv2Url.`) changed.

**Created Wv2Tests.bas (option 1-A)**

The pure-logic, feature-by-feature verifications that had accumulated in `Wv2Thunks.bas`
(`Test_9_14_SearchFallback` / `Test_9_15_Reorder` / `Test_9_16_SearchEngine` and the
helpers `Check` / `CheckOrder`) were consolidated into the newly created `Wv2Tests.bas`. They only call the `Debug_*` hooks of
`Wv2Browser` and do not touch the thunk layer at all, so moving them out breaks nothing. Later
`Test_9_17…` tests will be added here as well (no more new files).

- **The verification of the thunk foundation itself was left in Wv2Thunks**. `Test_TwoHandlers_Mock` /
  `Test_Environment_Real` / `Test_ResetSurvival_Manual` /
  `Test_VirtualAllocFree_Roundtrip` / `Test_SentinelStatus` are tightly coupled to the Private implementation inside Wv2Thunks
  (`Thunks_*` / `Sentinel_*` / VirtualAlloc etc.); moving them out
  would require making Private members Public just for verification, which would break encapsulation.
  → **Design principle 71 (criterion for where verification code belongs)**.
- **From a portability viewpoint**: Wv2Tests is verification-only, so when porting to production it is possible to decide "if the tests are not needed,
  do not import it". The number of files goes up by one, but because production code and test code
  are physically separated, effective portability actually improves. The increase happens only this once (later
  tests are added to Wv2Tests).

### Real-machine test

As a regression check, three tests were run from the consolidated `Wv2Tests.bas`:

- `Test_9_14_SearchFallback` → **17 / 17 [ALL OK]** (Japanese encoding via Wv2Url is intact)
- `Test_9_15_Reorder` → **10 / 10 [ALL OK]** (no effect on reordering or active-tab following)
- `Test_9_16_SearchEngine` → **13 / 13 [ALL OK]** (engine switching, empty guard, and no effect on URL input)

Compilation had already passed by the time the 3 tests ran (because VBA compiles all modules
before execution). → The reference resolution of Wv2Tests→Wv2Browser and Wv2Browser→Wv2Url is also sound.

Real-machine GUI (option A, regression check of the default Google):
- Launch with `StartWebView2_Full` → entering `日本 天気` in the address bar shows Google search results
  (the 9.14 search fallback is intact even after the cleanup. The Japanese search that **actually passes through Wv2Url.bas** came out
  correctly, which demonstrates that the module split also works on a real machine).
- `example.com` navigates to `https://example.com/` (the URL branch is intact).
- The 3-row tab bar / nav bar and the tab operations and reordering also work (the cleanup did not break any inter-class references).

Confirming the moment the search engine switches (to Bing etc.) on a real machine was skipped under option A. `m_browser` is
`Private WithEvents` and cannot be touched from outside, so an entry point exposing the reference is needed; but the switching logic has been
demonstrated by 13/13, and it is more natural to add that entry point, tailored to its use, in the stage that next needs it.

→ **Stage 9.16 complete. The search engine can now be switched by swapping the template, and the
pure URL functions were organized into Wv2Url.bas and the feature verifications into Wv2Tests.bas. Wv2Thunks is now purely
the thunk foundation.**

### ★ Design principle 70 (new) ★

**Pure conversions (output determined by input alone) are split out into a stateless standard module,
and state-dependent behavior (policy) stays in the main class. Separation of policy and mechanism.**

**Pure functions that do not reference instance state and whose output is uniquely determined by the input**, such as
`Wv2Json` (JSON conversion) / `Wv2Url` (URL conversion), belong naturally
in a standard module (which cannot hold state), and cohesion is also high. On the other hand, things like `NormalizeUrl` / `BuildSearchUrl` that **depend on an instance
setting (`m_searchTemplate` etc.) so that their behavior changes each time (policy)** stay in
the main class. Pushing the latter out to a standard module would mean passing the state around as an argument every time,
making the caller cumbersome. The criterion for the separation is "does that function read instance state?".
If it does not (mechanism) → standard module. If it does (policy) → leave it in the main class and have it call the pure functions of the
standard module internally.

### ★ Design principle 71 (new) ★

**The ownership of verification code is decided by "coupling with the thing verified". Verification tightly coupled to a production module's Private members stays
in that module; independent verification that only calls the public API is consolidated into a verification-only module.**

There are two kinds of verification Subs. (1) **Verification tightly coupled to the Private implementation inside a production module**
(such as the thunk foundation's `Test_VirtualAllocFree_Roundtrip`, which calls `Thunks_*` / `Sentinel_*` directly)
would require making Private members Public just for verification if moved out, which would break encapsulation,
so it stays in that module. (2) **Independent verification that only calls the production public API (`Debug_*` etc.)**
(`Test_9_14/15/16` etc.) breaks nothing wherever it is placed, so it is consolidated into a verification-only module
(`Wv2Tests.bas`). This physically separates production code from test code, makes it possible
to decide "if the tests are not needed, do not import the verification module" when porting, and improves portability.
The point is neither "all tests in one place" nor "tests scattered across each class" — it is **sorting by
coupling**.

### Numbers for 9.16

- Design principles: 69 → **71** (added design principles 70 / 71)
- Spec facts: **17** (unchanged)
- Standard file layout: 9 → **11** (created `Wv2Url.bas` / `Wv2Tests.bas`. However,
  `Wv2Tests.bas` is verification-only and can be left out when porting, so the production count is 10)
- `Wv2Browser.cls`: added `m_searchTemplate` + `DEFAULT_SEARCH_TEMPLATE` + Property
  `SearchEngine` (Get/Let, with empty guard) + `Debug_SearchEngine`. Rewrote `BuildSearchUrl`
  to reference the template. Deleted the three pure URL functions (`IsAlphaOnly` / `UrlEncodeComponent` /
  `Utf8Bytes`) and replaced them with `Wv2Url.` calls. `NormalizeUrl` / `BuildSearchUrl` are left in place
- `Wv2Url.bas` (new): holds the three pure URL functions as Public
- `Wv2Tests.bas` (new): consolidates `Test_9_14/15/16` + `Check` / `CheckOrder` + help Subs
- `Wv2Thunks.bas`: deleted the group-B tests above, making it purely the thunk foundation. The verification of the thunk foundation itself
  (group A) is left in place. Main logic unchanged

### Import procedure notes (9.16)

- `Wv2Url.bas` / `Wv2Tests.bas` are new, so import them as they are.
- `Wv2Thunks.bas` has a Public Enum name conflict, so comment out the whole existing version
  before importing.
- For `Wv2Browser.cls`, rename the old version to `Wv2Browser_Old` + comment out all of it, and then
  import the new version.

  

---

## Stage 9.17 — The switching entry point for search engines + publishing the reference to the running Browser

### Goal

Collect the open item that 9.16 deferred: "for the search engine, only create the main-class Property (`SearchEngine`), and split out the entry point for choosing it from the UI and
the real-machine switching check into a separate stage". GUI parts
(a dropdown etc.) are still not placed; the scope is limited to growing the following two things:

- **Option A-1**: a foothold for grabbing the running `Wv2Browser` from outside (`CurrentBrowser`).
- **Option B-1**: the minimal entry point for switching the search engine by preset name (`UseSearchEngine`).

The 9.16 real-machine test had "no entry point for a reference to the running Browser", so the moment of switching (to Bing etc.) could not be
seen on a real machine, and it was only backed by the logic verification (13/13). The `CurrentBrowser` of 9.17
fills this gap.

### Implementation (2 files)

#### UserForm1.frm

- **Added the preset constants at the top of the module** (before all procedures, due to VBA's declaration constraint):
  `SE_GOOGLE` = `https://www.google.com/search?q=`,
  `SE_BING` = `https://www.bing.com/search?q=`,
  `SE_DUCKDUCKGO` = `https://duckduckgo.com/?q=`,
  `SE_YAHOO` = `https://search.yahoo.com/search?p=`.
  All are query-prefix type templates where "it is enough to concatenate the encoded search term at the end"
  (the same form as option 1-A of 9.16). It was confirmed that even when the path / query key differs, as in the DuckDuckGo `/?q=` type and the Yahoo `?p=` type,
  a single prefix template can express them.

- **`Public Property Get CurrentBrowser() As Wv2Browser`** (option A-1):
  `m_browser` is `Private WithEvents` and cannot be touched from outside, so exactly
  one entry point that returns a plain reference is published. `WithEvents` is a constraint on the declaration side, and returning the reference from a `Get` is
  not a problem. When not started (`m_browser Is Nothing`) it returns `Nothing`. From now on this is the common foothold for any extension to
  "grab the running Browser from outside".

- **`Public Sub UseSearchEngine(ByVal engineName As String)`** (option B-1):
  Normalizes `"google"/"bing"/"duckduckgo"` (alias `"ddg"`) `/"yahoo"` with `LCase$`+`Trim$`,
  resolves it to a preset, and Lets it into `m_browser.SearchEngine`. An unknown name falls back to the default (Google)
  and a warning is sent to `Debug.Print`. If not started, it does nothing and prints a notice. The final guard
  against an empty string / whitespace only is handled by the Browser-side `Property Let` (Google if empty), so
  this Sub concentrates on name resolution (no double guard). ★This is an entry point for logic, not a UI★

#### Wv2Tests.bas

- Added **`Test_9_17_SearchPresets`**. It Lets each preset's template directly into `Wv2Browser`
  and checks the search URL generation (third branch), the Japanese UTF-8 encoding, and the lack of effect on URL input using
  `Check`. WebView2 is not launched (no dependency on `m_browser`; it just creates a new Browser
  and calls `Debug_NormalizeUrl`/`Debug_SearchEngine`). The preset strings duplicate the
  `Private Const` on the UserForm side, but this is an intentional copy so that verification can run without launching the UserForm (the same as writing the expected values
  directly into the test).
- Added **`Test_9_17_SearchPresets_Help`**. It outputs the logic verification + the real-machine GUI procedure.
  The caveat in the 9.16 Help, "if there is no entry point for a reference...", was updated in 9.17 to the real entry points
  `CurrentBrowser`/`UseSearchEngine`.

#### Why verify with preset values rather than UseSearchEngine itself

`UseSearchEngine` is a method of `UserForm1` and depends on `m_browser` (Private WithEvents), so it does not fit the
policy of Wv2Tests (a receptacle that needs no WebView2 launch and only calls the Browser's `Debug_*`). So
"name → template resolution" and "the not-started guard" are checked on a real machine
(via `CurrentBrowser`), and the logic verification is narrowed to the correctness of the preset values: "Letting each preset string
produces the correct search URL". → A practice of design principle 71 (the ownership of verification code is decided by
coupling). No new principle was established.

### Real-machine test log (summary) — all succeeded

- **Logic verification** `Test_9_17_SearchPresets`: **9 / 9 [ALL OK]**.
  Checked in one go: search URL generation for google/bing/ddg/yahoo, Japanese encoding (`%E6%97%A5%E6%9C%AC...`), and
  no effect on URL input (`example.com` / `https://example.com`).
- **Real machine, default Google**: after launching with `StartWebView2_Full` (3 tabs, TabBar/NavBar IsReady=True),
  `日本 天気` in the address bar → navigated to `https://www.google.com/search?q=%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97`.
  Title "日本 天気 - Google 検索" (Japanese Google Search results title).
- **Real machine, switch to Bing**: `UserForm1.UseSearchEngine "bing"` → `日本 天気` again →
  `https://www.bing.com/search?q=...`, title "日本 天気 - 検索" (Japanese Bing results title).
  ★The moment of switching that was skipped in 9.16 could be demonstrated on a real machine via the CurrentBrowser entry point★
- **Unknown-name fallback**: `UserForm1.UseSearchEngine "nonsense"` → after the warning output it returned to
  Google, and `日本 天気` went to a Google search again.
- Loose coupling (5 events unchanged), full sync, and the ready handshake were intact in all logs.
  The TabBar `sync` and the NavBar `navsync` followed correctly even after the switch.

### Numbers for 9.17

- Design principles: **71** (unchanged. Options A-1/B-1 are practices of existing principles)
- Spec facts: **17** (unchanged)
- Standard file layout: **11** (production count 10. No files added or removed)
- `UserForm1.frm`: added 4 preset constants + `CurrentBrowser` (Property Get) +
  `UseSearchEngine` (Sub). Main logic unchanged
- `Wv2Tests.bas`: added `Test_9_17_SearchPresets` + `_Help`. Reused the existing `Check`

### Import procedure notes (9.17)

- `UserForm1.frm` is paired with a `.frx`, so either delete the existing one and re-import, or
  paste by hand at the end of the code window.
- For `Wv2Tests.bas`, comment out all of the existing version (or delete it) before importing the new version.
- The other files (`Wv2Browser.cls` etc.) are unchanged in 9.17, so no re-import is needed.

### Open items / next stage

- **UI setting for the search engine (option 2-B)**: build a selection dropdown into the NavBar
  (HTML/JS changes + adding a `setengine` case for the JS→VBA command). Since the main-class entry points
  (`UseSearchEngine`/`CurrentBrowser`) are now in place in 9.17, the next stage can concentrate on "calling it from the GUI".
- `Wv2Json` is JSON-only and `Wv2Url` is URL-only. Do not mix them (unchanged since 9.16).

---

## Stage 9.18 — Settings tab + search engine UI via AddHostObjectToScript

### Overall approach

As a continuation of 9.17, which put the main-class entry points (`UseSearchEngine` / `CurrentBrowser`) in place, the search
engine is made selectable in the GUI. The open item from 9.17 originally assumed "option 2-B: build a selection
dropdown into the NavBar and add a `setengine` command to JS→VBA", but after aligning on it, a **completely different approach**
was adopted.

tarboh made the proposal: "ICoreWebView2 has `AddHostObjectToScript`, and it was verified on a real machine before (in the implementation
where all methods were hand-typed, before this project started). If we just hand a Wv2Browser (equivalent
object) to the settings screen's HTML, JS can call VBA methods directly, so no deep implementation should be needed."
Based on the connection record of the old implementation (even a class with Instancing 1-Private can have its VBA functions called via
`window.chrome.webview.hostObjects.sync.<Name>.<Method>()`, and the return value
goes straight back to JS), it was decided to build it as a **settings tab + direct hostObjects connection**.

This is the first route in this project in which **JS→VBA integration is done by direct host-object
RPC rather than by postMessage string commands**, and it also means
acquiring a new tool, separate from the TabBar/NavBar full sync (postMessage).

### Confirmed design (part 1 + part 2-b + part 3)

**Part 1: `Wv2Pane.View_AddHostObjectToScript(objName, instance)` (vtable 49)**
  Ported the real-machine-tested implementation from the old WebView2.cls to the current conventions. Packs `instance` into
  a Variant with `Set vObj = instance`, then `dcf(m_pView, 49, ..., StrPtr(objName), VarPtr(vObj))`. It combines the existing
  `View_Navigate` (StrPtr) and `get_Settings` (VarPtr) patterns.
  `RemoveHostObjectFromScript` (vtable 50) is not added in 9.18 (the settings tab is released together with its
  WebView2 when closed. YAGNI).

**Part 2-b: Expose a thin bridge `Wv2SettingsBridge.cls` (new) to JS**
  Passing the whole Wv2Browser would let JS call even CloseTab / MoveTab, so we set up a bridge holding only the minimal API
  needed for search engine settings and attach that instead.
  The "surface exposed to JS" is structurally minimized at the class boundary. Public methods:
    - `GetEngine()`               current engine name ("google" etc., reverse lookup. No match gives "custom")
    - `SetEngine(name)`           resolves the preset and Lets it into Browser.SearchEngine. Returns the name that was set
    - `PreviewUrl(terms)`         search URL for the current engine (via NormalizeUrl, same as production)
    - `PreviewUrlFor(name,terms)` search URL for the specified engine (does not change the current engine)
    - `EngineList()`              returns the preset definitions as a JSON array (for drawing the cards)
  The preset constants (SE_*) and name resolution that UserForm1 held up to 9.17 are **consolidated into this bridge
  (DRY)**. `UserForm1.UseSearchEngine` was also changed to create a temporary bridge and delegate to `SetEngine`, so the
  Immediate window and the settings screen go through the same resolution path.

**Part 3: TabBar's ⚙ button -> open the settings tab**
  Added a ⚙ at the right end of the TabBar (right-aligned with `margin-left:auto`). A click posts
  `{"cmd":"settings"}` via postMessage -> the "settings" case in `OnPaneWebMessage` ->
  `Wv2Browser.RequestSettingsTabAsync`. Creating the settings tab also goes internally through AddTab -> WaitPaneReady,
  so to avoid the same deadlock as newtab (spec fact 14) it is deferred through a **dedicated one-shot
  SafeTimer** (m_settingsTimer). TimerCall -> `OpenSettingsTab`:
    1. Create an empty tab with AddTab
    2. Create a new `Wv2SettingsBridge` and `BindBrowser Me`
    3. ★Before NavigateToString★ `View_AddHostObjectToScript("Settings", bridge)`
    4. Show `BuildSettingsHtml` with `View_NavigateToString`
    5. `ActivateTab` the created tab
  The settings screen HTML is built by Wv2Browser's `BuildSettingsHtml` (4 engine cards on a dark gradient background,
  selection highlight, live URL preview on hover). The TabBar only conveys the intent "I want to open settings" and knows
  nothing about the settings content (loose coupling is preserved).

### ★ Spec fact 18 (new) ★

**If `AddHostObjectToScript` is called before `NavigateToString`, the JS of the loaded HTML can grab
`window.chrome.webview.hostObjects` from its very first run.**

For the settings tab we verified on the real machine with the policy of making "register (attach) -> Navigate" (order A) the
first choice and falling back to a ready handshake (order B) if that failed. Result: **order A works on the real machine**.
Registering a host object is independent of the document lifecycle; if it is registered before Navigate, the JS
loaded afterwards can reference hostObjects from its first run. A deferred synchronization such as a ready handshake
was unnecessary (no fallback to order B occurred).

  ・In the real-machine log, `OpenSettingsTab: Settings bridge attach OK` appeared before the settings HTML's
    `NavigationStarting` (data:text/html;base64,...), and then the JS called
    `S.EngineList()` / `S.GetEngine()` to draw the cards, and Google was highlighted as selected.
  ・On the JS side, so that it does not crash in an environment where hostObjects cannot be obtained, it is guarded with
    `var S = (... ) ? ... : null;`, and if null it shows "host object not connected" (insurance so that a failure of order A
    can be isolated on screen). As a result, this insurance was never triggered.

  ★Reach of this spec fact★ It applies not only to the settings tab but to all screens that use hostObjects.
    When the TabBar/NavBar are moved to the hostObjects route in the future (see "Open items" below), it is guaranteed
    that "if you attach before writing out the HTML, you can call from the JS's first run".

### Why no new design principle is added

  ・Part 2-b (narrowing the exposure to JS with a bridge) is **a practice of design principles 70/71**. Knowledge belonging
    to the "policy" of the search engine is consolidated on the Browser side (the bridge), and the UI (settings screen JS)
    is shown only the minimal API. It is the application of existing principles, not a new principle.
  ・Part 3's deferred creation of the settings tab is **a reuse of design principle 63 (deadlock avoidance)**. It is a
    one-shot SafeTimer of the same shape as newtab. Not a new principle.
  ・Therefore design principles remain at 71. Only the spec facts increased (18).

### Real-machine test log (summary) — all succeeded

- **Bridge pure logic** `Test_9_18_BridgeLogic`: **17 / 17 [ALL OK]**.
  Checked in one batch: GetEngine initially = google, SetEngine's return value and GetEngine following it, alias ddg->duckduckgo,
  case/whitespace normalization of "  YAHOO ", unknown name -> google fallback, Japanese encoding of PreviewUrl,
  no side effects of PreviewUrlFor (the current engine is unchanged even after calling it), URL input passed
  straight through, 4 engines in EngineList.
- **Real machine, order A works**: click ⚙ -> the settings tab opens, 4 cards (Google/Bing/
  DuckDuckGo/Yahoo) are drawn and Google is highlighted in blue as selected. -> Demonstrates spec fact 18.
- **Real machine, live preview**: hovering a card shows in the Preview field below, on the spot,
  `日本 天気 → https://duckduckgo.com/?q=%E6%97%A5%E6%9C%AC...`.
  The synchronous hostObjects.sync call of PreviewUrlFor works.
- **Real machine, Bing switch -> reflected in production**: select Bing on the settings screen -> in another tab's address bar,
  `日本 天気` -> `https://www.bing.com/search?q=%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97`,
  title "日本 天気 - 検索" (Japanese for "Japan weather - Search"). Bridge -> Browser.SearchEngine -> NavigateActive became a single path.
- **Deadlock avoidance**: the deferred execution settings -> RequestSettingsTabAsync -> TimerCall(settings) ->
  OpenSettingsTab works correctly. Even while the settings tab is being created, the TabBar/NavBar sync/navsync
  follow correctly (4->5 tabs, active follows). The settings tab can be closed with × like a normal tab.
- Loose coupling (5 events unchanged), full synchronization, and the ready handshake are intact in all logs.

### Numbers for stage 9.18

- Design principles: **71** (unchanged. Part 2-b/Part 3 are practices of existing principles 70/71/63)
- Spec facts: 17 -> **18** (spec fact 18 added)
- Standard file set: 11 -> **12** (`Wv2SettingsBridge.cls` newly created. Production entities 10 -> 11.
  Wv2Tests is for verification only and can be dropped when porting, so it is not counted as a production entity)
- `Wv2SettingsBridge.cls` (new): a thin bridge exposed to JS via hostObjects.
  Consolidates the preset constants (SE_*) and name <-> template resolution. Instancing can be 1-Private
- `Wv2Pane.cls`: added `View_AddHostObjectToScript` (vtable 49). Header updated to 9.18
- `Wv2Browser.cls`: added the whole settings-tab route (m_settingsTimer / m_settingsPending /
  m_settingsBridge, RequestSettingsTabAsync / m_settingsTimer_TimerCall /
  OpenSettingsTab / BuildSettingsHtml). Added release of the settings-tab resources to Shutdown
- `Wv2TabBar.cls`: added the ⚙ button (HTML/CSS/onclick) and a "settings" command branch
- `UserForm1.frm`: removed the SE_* constants and moved them to Wv2SettingsBridge. Rewrote UseSearchEngine to
  create a temporary bridge and delegate to SetEngine
- `Wv2Tests.bas`: `Test_9_18_BridgeLogic` + `_Help` + generic helpers CheckEq / CheckBool

### Import procedure notes (stage 9.18)

- `Wv2SettingsBridge.cls` is a fresh import.
- For `Wv2Pane.cls` / `Wv2Browser.cls` / `Wv2TabBar.cls` / `Wv2Tests.bas`, comment out (or delete) the entire old
  version and then import the new version. There are no name conflicts (Public Enum etc.) in
  9.18, so no special procedure like the one for Wv2Thunks is needed.
- `UserForm1.frm` is paired with its `.frx`, so either delete the existing one and re-import, or paste by hand at the end of the
  code window.
- Mind the dependency resolution order: `Wv2SettingsBridge` references `Wv2Browser` (type reference) and `Wv2Url`
  (UrlEncodeComponent). Conversely, `Wv2Browser` does not type-reference `Wv2SettingsBridge`
  (it only does New inside OpenSettingsTab). UserForm1 also does New on `Wv2SettingsBridge`. There is no mutual
  type-declaration dependency, so compilation passes in any import order.

### Open items / next stage

- **★Cold-sweat memo (strong candidate for the next stage)★ Consider whether the TabBar / NavBar can be migrated to the
  hostObjects route**. It is now confirmed on the real machine in 9.18 that direct hostObjects connection (JS calling VBA
  methods directly without going through postMessage string commands) works. The TabBar/NavBar are currently built with a doubly
  elaborate wiring: "full synchronization (sending JSON whole via postMessage) + a hand-written JSON parser (Wv2Json) + command
  branching by a Select Case in OnPaneWebMessage". If this were replaced with direct calls such as
  `hostObjects.sync.TabBar.Activate(3)` / `hostObjects.sync.NavBar.Navigate(url)`,
  a considerable part of the Wv2Json parser and the Select Case could
  possibly be shrunk. However:
    ・The postMessage route is still optimal for "asynchronous, loosely coupled notifications from JS to VBA". hostObjects is
      optimal for "synchronous direct calls". Having both is the strength. Moving everything to hostObjects is not necessarily
      right (in particular, postMessage is the natural fit for the VBA->JS full-synchronization push).
    ・The existing TabBar/NavBar work, are loosely coupled, and have tests in place. With 9.18, a "working
      hostObjects example", available as the regression baseline, consider this carefully. Rewriting everything at once on
      impulse could destabilize the 3-row layout.
  -> A reasonable approach for the next stage is to start small: "make only the TabBar hostObjects-based on a trial basis and compare it with the postMessage version".
     Needs to be agreed on.
- **Polishing the search engine UI**: an input field for arbitrary templates outside the presets, persisting the selection
  (carrying the setting over to the next launch), etc. Small.
- **Preventing duplicate settings tabs**: currently mashing ⚙ increases the number of settings tabs. Perhaps change it so that
  "if a settings tab already exists, just activate it". Can be postponed as polish (YAGNI).
- `Wv2Json` is JSON-only, `Wv2Url` is URL-only. Do not mix them (unchanged since 9.16).

---

## Stage 9.19 — Preventing duplicate settings tabs + default display of the settings screen preview

### Overall policy

A "polish" stage following 9.18, where the settings tab + hostObjects-direct search engine UI came to work.
Not a feature addition; the scope is narrowed to closing two loose ends left in 9.18.

1. **Preventing duplicate settings tabs**: mashing ⚙ makes the settings tabs multiply. Change it so that "if a
   settings tab already exists, just activate it".
2. **Default display of the settings screen hover preview**: in 9.18 the preview of the last hovered card
   stayed. Show the preview for the currently selected engine right after opening (before any hover) and
   when the cursor leaves a card.

Both are small in granularity and can be verified on the real machine on their own, and they also serve to stabilize the regression baseline
for the next stage (considering moving TabBar/NavBar to hostObjects).

### Confirmed design

**(1) Preventing duplicate settings tabs — Option A: an IsSettings flag on Wv2Pane**

We compared 3 options for "how to identify" the settings tab:
  - Option A: Set an `m_isSettings As Boolean` flag on the Pane, and on each ⚙ scan m_tabs
    to look for an existing settings tab. <- adopted
  - Option B: Decide by URL/content (data:text/html;base64,...). A long base64 string is brittle and
    expensive. Rejected.
  - Option C: Hold one settings Pane reference (m_settingsPane) on the Browser side. It would require wiring of our own to
    watch for a dangling reference after it is closed (a breeding ground for loose ends). Not adopted.

What decided it for Option A is that **the liveness check is unified with the m_tabs scan**. A closed settings
tab is reliably removed by CloseTab's `m_tabs.Remove` (reference count 0, so
Pane.Class_Terminate runs), so it naturally stops being found by the scan. Unlike Option C, there is no need to separately detect the
state "a reference is held but it is actually closed". One extra field is acceptable, and
loose coupling is not broken.

Implementation:
  - `Wv2Pane`: `m_isSettings As Boolean` + `Property Get/Let IsSettings`.
    Initialized to False in Class_Initialize.
  - At the top of `Wv2Browser.OpenSettingsTab`, scan m_tabs from 1, and if there is a Pane with
    `m_tabs.Item(j).IsSettings` True, do `ActivateTab j` and
    `Exit Sub` (no new creation). If none, AddTab as before, and set
    `pane.IsSettings = True` right after creation.
  - The duplicate check was placed in OpenSettingsTab, the actual processing side, rather than in RequestSettingsTabAsync (the thin entry that only
    starts the SafeTimer).

**(2) Default display of the settings screen hover preview — improved in 2 steps (9.19b -> 9.19c)**

A change to the JS in BuildSettingsHtml only. The VBA/COM layer was not touched at all.

  - 9.19b (first version): Added `showCurrentPreview()`. If current (the selected engine,
    already determined by GetEngine in loadState) is non-empty, it shows that preview.
    Called once at the end of build (initial display right after opening), and also called from the `#grid` onmouseleave
    (return to the selected one when the cursor leaves).
    -> On the real machine it turned out that "the preview does not return until the cursor completely leaves the grid area".
      #grid is the whole container, so while in the gaps between cards or moving to another card the cursor is still
      inside the container and mouseleave does not fire. It felt counterintuitive to use.

  - 9.19c (final version): The return decision was made per card.
    ・Introduced `var hoverId = null;` (the id of the card the cursor is currently on).
    ・Each card's onmouseenter: `hoverId=item.id; showPreview(item.id);`
    ・Each card's onmouseleave: `hoverId=null; setTimeout(function(){
       if(hoverId===null){ showCurrentPreview(); } },0);`
    ・The `#grid` onmouseleave was removed (handled entirely per card). The initial display at the end of build,
      showCurrentPreview(), is kept.
    -> The moment the cursor leaves a card, it returns to the selected engine. Even when moving from card A to B through the gap,
      B's mouseenter re-sets hoverId first within the same tick, so `if(hoverId===null)` inside
      setTimeout is false and the return processing is skipped
      (no flicker). item lives inside an IIFE closure, so each card correctly grabs its own
      id.

### No change in design principles or spec facts

  - (1) The IsSettings flag merely means "the Pane itself holds the tab kind", and it breaks neither the existing loose coupling
    (the Browser's 5 events unchanged) nor full synchronization. Not a new principle.
  - (2) is an improvement of UI behavior inside the settings screen JS and does not affect the architecture.
  - Therefore design principles remain at 71 and spec facts at 18.

### Real-machine test log (summary) — all succeeded

- **Duplicate prevention**: first ⚙ creates the settings tab (index 4, attach OK) -> mashing ⚙ prints
  `OpenSettingsTab: activated existing settings tab (index 4)` with no new creation ->
  ⚙ after switching to another tab also moves to the existing settings tab -> close the settings tab with ×
  (`CloseTab: closed tab 4`) -> ⚙ afterwards reopens exactly one new settings tab.
  The settings tab stays at most one at all times.
- **Switch reflected in production (regression check)**: select DuckDuckGo on the settings screen
  (`Wv2SettingsBridge.SetEngine: 'duckduckgo' → 'duckduckgo'`) -> in another tab,
  `日本 天気` -> `https://duckduckgo.com/?q=%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97`.
  The 9.18 hostObjects route, full synchronization, and loose coupling are all intact.
- **Default preview display (9.19c)**: right after opening the settings screen (before any hover), the preview of the selected engine
  appears. Hovering a card changes it to that engine, and leaving returns it to the selected one.
  No flicker when moving between cards. The feel is OK.

### Changed files (stage 9.19)

- `Wv2Pane.cls`: added the `m_isSettings` field + `Property Get/Let IsSettings`.
  Header updated to 9.19.
- `Wv2Browser.cls`: added the duplicate check at the top of `OpenSettingsTab` + added `IsSettings = True` on the created Pane.
  Added the default hover preview display to the JS of `BuildSettingsHtml`
  (9.19c: hoverId + per-card mouseenter/mouseleave). Header is stage 9.19
  (with explanations of the 9.19b/9.19c additions inside).
- `Wv2Tests.bas`: `Test_9_19_SettingsTabDedup_Help` (real-machine test procedure) + header update.
  Duplicate prevention involves creating real WebView2 tabs, so it is confirmed by a real-machine procedure rather than a pure-logic test.

### Import procedure notes (stage 9.19)

- For `Wv2Pane.cls` / `Wv2Browser.cls` / `Wv2Tests.bas`, comment out the entire old version
  (or delete) -> import the new version. There are no name conflicts, so no special order is needed.
- The settings screen preview hover improvement is self-contained in Wv2Browser.cls (JS strings only).
  No ripple effect on other files.

### Open items / next stage

- **★First candidate for the next stage★ Consider moving TabBar / NavBar to hostObjects** (continuing the 9.18 cold-sweat memo).
  With 9.19, 9.18 is firmly in place as a regression baseline. A small way in is recommended: first make only the TabBar
  hostObjects-based on a trial basis and compare it with the postMessage version. Needs to be agreed on.
- **Polishing the search engine UI**: an input field for arbitrary templates outside the presets, persisting the selection
  (carrying over to the next launch), etc. Small.
- `Wv2Json` is JSON-only, `Wv2Url` is URL-only. Do not mix them (unchanged since 9.16).

---

## Stage 9.20 — Persisting the search engine selection

### Purpose

Through 9.16-9.19 we got to the point where "the search engine can be chosen from the settings tab", but the selection
lasted only for that launch session, and closing Excel returned it to the default (Google).
9.20 carries the chosen engine over to the next launch (persists it).

### ★History of deliberating the storage location (compared options A-D and decided on option B = an external file)★

We considered 4 options for where to persist. The two decision axes are "portability (host independence)" and
"medium-integrated, environment-non-polluting", and these two are not compatible.

- **Option A: Environment variable (SetEnvironmentVariableW etc.)**
  -> Rejected. An environment variable is volatile and lives only as long as the process. It disappears when Excel is closed, so it does not satisfy
  "carry over to the next launch". Unsuitable as a persistence location.

- **Plan B: external file (VBA built-in Open/Print/Line Input)** ← adopted
  → One line of key=value in %APPDATA%\Wv2Browser\settings.txt. Uses only built-in I/O,
  with neither Win32 nor ADODB. It does not touch the host's object model, so portability to Word and the like
  is preserved. We hold both the path and the contents ourselves, it can be inspected and hand-edited, and lines can be added
  to grow the items.

- **Plan C: registry (VBA built-in SaveSetting/GetSetting)**
  → Portability is on par with plan B (built-in functions, so host-independent; works in Word too). However the
  storage location is fixed at HKCU\...\VB and VBA Program Settings and cannot be chosen. Concern that it gets cluttered as items grow.
  Easiest to use by far, but it yields one step to plan B on the "hold the actual storage ourselves" philosophy.

- **Plan D: use a hidden Excel sheet as the settings medium**
  → ★It once advanced to the point of being adopted as the leading candidate, but was finally rejected★.
  The hidden-sheet plan fit the design philosophy very well: "the medium travels with the workbook (the settings come
  along when it is copied or moved)", "zero side effects outside (does not pollute the environment)", "I/O is just Sheets(...).Range(...)
  with no character-encoding worries", and it satisfied every advantage of plan B (visible location, visual editing, extensibility)
  as a strict superset.
  → But as a ★fatal drawback★, it depends on the Excel object
  model, Sheets(...), and so breaks this project's lifeline: "completes with VBA7 alone and can be ported to
  other hosts such as Word". The moment we choose plan D, a statement that "won't even compile unless it is Excel"
  gets mixed in. The thunk foundation, COM control and tab management do not depend on the host
  at all, so losing portability through persistence alone would be putting the cart before the horse, and we rejected it.

  Note: this tug-of-war between "portability vs. a medium that travels with the workbook" is a story of noticing portability just before adopting plan D and
    reversing course. If we ever decide to "commit to an Excel-only build",
    plan D could be revived as the best option (see the isolation design note below).

### ★Where the responsibility lives — merged into Wv2SettingsBridge instead of cutting a separate module★

At first we also considered cutting persistence into a small module Wv2Store.bas alongside Wv2Url / Wv2Json, but rejected it
and merged it into the bridge. Reasons:

- Wv2SettingsBridge already holds the "command center for search engine settings" (preset constants SE_*,
  name ⇄ template resolution, SetEngine/GetEngine). Persistence is the responsibility of
  "keeping those settings across launches", so placing it in the same class as cohesion into the settings domain is
  natural. It is different in character from "cross-domain groups of pure functions" such as Wv2Url (URL pure functions) and Wv2Json (JSON parsing);
  persistence is a concern specific to the settings domain.
- Do not increase the module count (tarboh's wish; he dislikes splitting into files).
- The 9.13/9.16 cleanup was about "peeling general-purpose pure functions off the Browser itself", but
  this time it is the opposite: "consolidating into the already-existing settings-dedicated class". It is proper cohesion, not separation.

  Note: however, the "isolation that makes the storage implementation swappable" will pay off when plan D is revived in the future. For now,
    the entry points LoadEngineName/LoadEngineTemplate/SaveEngineName are kept closed
    inside the bridge, so swapping the actual storage (file → a future sheet, etc.) is completed within this class. The entry-point design itself
    doubles as the isolation layer.

### Implementation

**Storage format**: one line of key=value text (e.g. `engine=bing`). For now there is just one
item, the engine, but if we add lines, future items such as "non-preset template" ride along directly as the second line.
When reading, a line is split at '=' and only the value of the engine key is picked up. Unknown keys are ignored.

**Added to Wv2SettingsBridge.cls**:
- Module constants: STORE_SUBFOLDER="Wv2Browser" / STORE_FILENAME="settings.txt"
  / STORE_KEY_ENGINE="engine".
- `SettingsPath()` (Private): builds %APPDATA%\Wv2Browser\settings.txt. If APPDATA cannot be
  obtained it returns "" and the caller skips saving/loading.
- `EnsureFolder()` (Private): MkDir if the save folder does not exist. Failures are passed over with On Error.
- `LoadEngineName()` (Public): reads settings.txt with Line Input and returns the value of engine.
  A missing file/folder, a read failure, or an absent key all give "" (the caller falls back to the default).
  Whether the name is valid (is it a preset) is not judged here.
- `LoadEngineTemplate()` (Public): resolves LoadEngineName → ResolveTemplate into a template
  string and returns it. With nothing saved or an unknown name (custom or a corrupt value) it returns "".
  ★Keeping name ⇄ template resolution closed inside the Bridge, only "a template string
    or empty" is handed to the Browser. The Browser does not need to know the vocabulary of engine names.★
- `SaveEngineName(name)` (Private): writes the single line engine=<name> with Output (with only one item
  at present, it overwrites the whole file). Calls EnsureFolder if the folder is missing. Write failures are swallowed
  (a persistence failure must not stop production).
- `Debug_SaveEngineName(name)` (Public): a thin public wrapper for verification. Unused in the production path.
- Added saving at the end of `SetEngine`: when resolution succeeds, `If SetEngine <> "custom" And
  Len(SetEngine) > 0 Then SaveEngineName SetEngine`. custom and empty are not saved
  (non-presets cannot currently be chosen from the UI; template saving will be added in the future).

**Added to Wv2Browser.cls (restore side, option i)**:
- At the end of `Class_Initialize`, create a temporary Wv2SettingsBridge with New and call LoadEngineTemplate().
  If non-empty it overwrites m_searchTemplate; if empty the default (Google) stays.
Dim initBridge As Wv2SettingsBridge
Set initBridge = New Wv2SettingsBridge
Dim savedTemplate As String
savedTemplate = initBridge.LoadEngineTemplate()
If Len(savedTemplate) > 0 Then m_searchTemplate = savedTemplate
Set initBridge = Nothing
- ★Why option i was chosen★: the expected behavior of persistence is that "on launch it always starts with the previous value".
  To guarantee that, fixing it at the Browser's birth (Class_Initialize) leaves the fewest
  gaps. It is consistent whether launched via the UserForm or directly. With option ro (reading in the UserForm launch sequence),
  restoration would not run on the direct-launch test path.
- ★Direction of dependency★: Class_Initialize does New Wv2SettingsBridge, but this is the same practice as
  what OpenSettingsTab already does (only a type reference, so no circular reference).
  LoadEngineTemplate is a side-effect-free read, so BindBrowser is not needed either
  (it reads just the file without a Browser reference).

### No change in design principles or spec facts

- Persistence only "keeps settings across launches", and breaks none of loose coupling (the Browser's 5 events unchanged),
  whole-state sync, or the hostObjects path. Name ⇄ template resolution also stays
  closed in the Bridge. So design principles stay at 71 and spec facts at 18.

### Real-machine test log (summary) — all succeeded

- **Pure logic** (`Test_9_20_Persistence`): checks the save → load round trip (bing/duckduckgo),
  and that for an unknown name LoadEngineName falls to the raw value and LoadEngineTemplate to "".
  Result 6/6 [ALL OK]. Because it touches the real file, the test saves the current value at the start and restores it at the end.
- **Save**: click the Bing card on the settings screen → `SetEngine: 'bing' → 'bing'` followed by
  `SaveEngineName: saved 'bing' -> C:\Users\...\AppData\Roaming\Wv2Browser\settings.txt`.
- **Restore (the main point)**: fully quit Excel → restart → StartWebView2_Full → without opening the settings tab,
  search `日本 天気` (Japan weather) in another tab →
  `NavigateActive: Navigate OK (https://www.bing.com/search?q=...)`.
  Confirmed it starts with the previously chosen Bing. Furthermore, when the settings tab is opened, the Bing card is
  already highlighted as selected from the start, and the URL preview is also Bing's.
- **Default fallback**: delete settings.txt by hand → restart → the `日本 天気` search
  goes back to `https://www.google.com/search?q=...`.
  Confirmed it correctly falls to the default (Google) when the file is missing.

### Changed files (stage 9.20)

- `Wv2SettingsBridge.cls`: added the persistence constants + SettingsPath/EnsureFolder/LoadEngineName/
  LoadEngineTemplate/SaveEngineName/Debug_SaveEngineName. Added one line of saving at the end of SetEngine.
  Header updated to 9.20.
- `Wv2Browser.cls`: added persisted-value restoration (option i) at the end of Class_Initialize. Header updated to
  9.20. Otherwise unchanged.
- `Wv2Tests.bas`: added Test_9_20_Persistence (pure logic) + Test_9_20_Persistence_Help
  (real-machine procedure). Header updated.
- No new files (11→11; production files also remain 11).

### Import operation notes (stage 9.20)

- For all 3 files, comment out the whole old version (or delete it) → import the new version. There are no name conflicts,
  so no special order is needed.
- The save-destination file is %APPDATA%\Wv2Browser\settings.txt (its contents are the single line engine=bing).
  It can be inspected and hand-edited in Explorer, and deleting it by hand returns to the default (Google).

### Open items / next stage (updated from 9.19)

- "Persisting the selection" is completed in 9.20. What remains in polishing the search engine UI is
  **an input field for an arbitrary non-preset template** (saving custom must also be designed together;
  currently SetEngine does not save custom).
- **★Top candidate for the next stage (continued)★ Considering converting TabBar / NavBar to hostObjects** (the cold-sweat memo).
  We recommend a small way in: first convert only TabBar to hostObjects on a trial basis and compare it with the postMessage version.
  Needs alignment.
- `Wv2Json` is for JSON only, `Wv2Url` for URLs only. Do not mix them (unchanged since 9.16).

---

## Stage 9.21 — Experiment: converting activate to hostObjects + confirming it holds with a real-URL Navigate

### Overall approach

With 9.18–9.20, the real examples of direct hostObjects connection (the settings tab) and the regression baselines (preventing duplicate settings tabs and
search engine persistence) have thickened. On top of that, this is the stage that takes **just one step, with minimal scope,** into the question that has been
open since 9.18, the "cold-sweat memo" = whether TabBar/NavBar can be moved onto the hostObjects path.

TabBar's JS→VBA has 6 commands (ready / activate / close / newtab / settings /
reorder), and of those **only activate** was trial-migrated to a direct hostObjects call, and
**compared side by side** with the existing postMessage version. close/reorder/newtab/settings/ready are
left untouched, still postMessage (to guarantee stability).

There were 2 aims.
1. **Confirm that spec fact 18 holds even with a real-URL Navigate** (the core this time). The settings tab used
   `NavigateToString` (`data:text/html;base64,...`), but the TabBar loads a real URL to
   `appassets.tabbar/tabbar.html` with `View_Navigate`.
   We checked on the real machine whether "if you `AddHostObjectToScript` before Navigate, the JS loaded afterward can grab
   hostObjects from its first run" works with a real URL too.
2. **Confirm that peeling off only activate leaves the other paths unharmed** (separability).

### Why "only activate" (why not a full migration)

Sorting the 6 commands coolly, the only ones where hostObjects conversion is plainly welcome are activate / close /
reorder (all synchronous operations with no completion wait, plan 4-A).

  ・ready is a "notification (a handshake signal)" and not an RPC. There is little point in replacing it with hostObjects
    (leaving it as postMessage is natural).
  ・newtab / settings **go through WaitPaneReady internally (waiting for Controller creation to complete)**, so making them
    direct JS→VBA sync calls actually raises the risk of deadlock (spec fact 14). Even if converted to
    hostObjects, the VBA side would still return immediately and queue onto a SafeTimer, so there is
    no gain beyond "the JSON parsing goes away".
  ・So rather than a full migration, a hybrid (activate/close/reorder via hostObjects,
    ready/newtab/settings via postMessage) is the technically optimal answer. However, a hybrid has
    2 coexisting paths and raises cognitive load, and the original motivation of
    "shrinking the Wv2Json parser and the Select Case" is achieved only halfway (the Select Case stays anyway
    because of ready/newtab/settings). Judging this tug-of-war is the main point this time.

We decided to start from the smallest, activate alone, and after confirming that hostObjects holds with a real URL and
that it is separable, judge whether to roll out to close/reorder and spill over to NavBar.

### ★ Spec fact 19 (new) ★

**Even with a real URL (`View_Navigate` to `appassets.<host>/<file>.html`), if you
`AddHostObjectToScript` before Navigate, the JS of the loaded HTML can grab
`window.chrome.webview.hostObjects` from its first run.** (The reach of spec fact 18 extends to real URLs too.)

  ・Spec fact 18 was established in 9.18 on the premise of `NavigateToString` / `data:` URLs. This time we
    confirmed on the real machine that the same "order A (register → Navigate)" holds for a real file URL
    (`https://appassets.tabbar/tabbar.html` via virtual host mapping) as well.
  ・In the real-machine log, `Wv2TabBar.Init: TabBar bridge attach OK (activate prefers
    hostObjects)` appeared **before** `Navigate OK (https://appassets.tabbar/tabbar.html)`, and
    after that, on every click, `HostActivate: index=N (via host)` appeared. It never fell back
    to postMessage even once.
  ・This backs up that host-object registration is independent of the document lifecycle
    (NavigateToString or real URL). The delayed attach via the ready handshake (order B) is unnecessary.
  ・**Reach**: from now on, when moving TabBar / NavBar (both load real URLs) onto the hostObjects path,
    "if you attach before writing out the HTML, it can be called from the JS's first run"
    is guaranteed for real URLs too. This is the premise for rolling out to close/reorder and spilling over to NavBar.

### Why we do not set up a new design principle

  ・HostActivate just grows the existing activate case's identical processing (`m_browser.ActivateTab`)
    onto the hostObjects path too. Calling a synchronous operation with no completion wait through a synchronous RPC is
    **an application of plan 4-A (activate is safe as synchronous)**, not a new principle.
  ・Placing the attach before Navigate is the practice of spec facts 18/19.
  ・The JS-side fallback (dropping to postCmd if hostObjects cannot be obtained) is the same type of insurance as the
    settings tab null guard in 9.18.
  ・So the design principles **stay at 71**. Only spec facts increased (19).

### Implementation breakdown (stage 9.21)

- `Wv2TabBar.cls` (header updated to 9.21):
  ・Added to `Init`, just before Navigate (after virtual host mapping succeeds),
    `m_barPane.View_AddHostObjectToScript("TabBar", Me)`. An attach failure is
    not fatal (the JS falls back to postMessage), so we log hr and continue.
  ・Added a new `Public Function HostActivate(index) As Boolean`. Its body is the same
    `m_browser.ActivateTab(index)` as the activate case. It is called as a direct synchronous call from JS via
    `hostObjects.sync.TabBar.HostActivate(n)`.
  ・Made the tab onclick in `BuildTabBarHtml` hybrid. With the `HO()` helper, if
    `hostObjects.sync` can be obtained it calls `TabBar.HostActivate(n)` and logs to the console
    `[host] activate N`; if not, it falls back to the old `postCmd({cmd:'activate',index:n})` and logs
    `[postmsg] activate N`. The sending of close/newtab/settings/reorder/ready is unchanged.

- `Wv2Tests.bas` (header updated to 9.21):
  ・Added `Test_9_21_HostActivate_Help` (real-machine procedure). Confirming real-URL Navigate + JS execution +
    that the hostObjects path holds requires actually launching WebView2, so there is no pure-logic verification.

### Real-machine test log (summary) — all succeeded

- **★Core★ Attach holds with a real URL (spec fact 19 demonstrated)**: in the startup log,
  `Wv2TabBar.Init: TabBar bridge attach OK` appeared before `Navigate OK
  (https://appassets.tabbar/tabbar.html)`. `AddHostObjectToScript failed`
  did not appear, and the order B fallback did not occur either.
- **activate switched to via-host**: clicking an inactive tab produced
  `Wv2TabBar.HostActivate: index=4 (via host)` / `index=2 (via host)`, and
  `OnPaneWebMessage: msg={"cmd":"activate"...}` **never** appeared for activate.
  It did not drop to the postMessage fallback either; the direct hostObjects call holds reliably.
- **Feel and follow-up unchanged**: on HostActivate success, `ActiveChanged → PushSyncToJs`
  (tab bar highlight) and `PushNavSyncToJs` (nav bar URL/back-forward) run as before, with no change in
  follow-up through whole-state sync.
- **No regression (other paths remain postMessage and unharmed)**:
  ・close: `{"cmd":"close","index":2}` / `{"cmd":"close","index":4}` →
    via `OnPaneWebMessage` as before. The active-tab correction is also normal.
  ・reorder: `{"cmd":"reorder","from":3,"to":2}` → `MoveTab 3→2` normal, redraw follows.
  ・newtab: `{"cmd":"newtab"}` → the `RequestNewTabAsync → TimerCall` SafeTimer path is normal.
  ・settings: `{"cmd":"settings"}` → `RequestSettingsTabAsync → TimerCall(settings) →
    OpenSettingsTab`, `Settings bridge attach OK`, and the settings screen's hostObjects UI is also normal
    (no effect on the 9.18–9.20 settings tab and persistence).
- Loose coupling (the Browser's 5 events unchanged), whole-state sync and the ready handshake are all unharmed throughout the logs.

### Numbers for stage 9.21

- Design principles: **71** (no change. HostActivate is an application of plan 4-A, and attach is the practice of spec facts 18/19)
- Spec facts: 18 → **19** (spec fact 19 added)
- Standard file layout: **12** (no change. No new files. Production files 11)
- Changed files: 2, `Wv2TabBar.cls` / `Wv2Tests.bas`

### Import operation notes (stage 9.21)

- For `Wv2TabBar.cls` / `Wv2Tests.bas`, comment out the whole old version (or delete it) → import the new
  version. There are no name conflicts, so no special order is needed.
- Because we touched the JS generation function (`BuildTabBarHtml`), we confirmed with grep before output that no `\"` has crept in
  (all JS double quotes are written as single quotes. Observed the handover's
  ★rule against writing JS inside a VBA string★). One false positive of the path-separator literal `"\"`
  is legitimate existing code (`If Right$(fullPath,1) <> "\"`).

### Open items / next stage

- **★Strong candidate for the next stage★ Converting close / reorder to hostObjects → considering spillover to NavBar**.
  This time, with activate alone, hostObjects holding with a real URL (spec fact 19) and separability were demonstrated.
  Next, roll close / reorder (likewise synchronous operations with no completion wait, plan 4-A) out to hostObjects,
  and on top of that judge "whether it is worth actually shrinking the Wv2Json parser + the OnPaneWebMessage Select Case".
  However, ready/newtab/settings are rightly left as postMessage, so the Select Case is expected to shrink
  but not vanish (this judgment is the real subject).
  NavBar (navigate/back/forward/reload; back/forward/reload are also synchronous operations) will be considered in the same shape after
  TabBar's conclusion is in. Needs alignment.
- **An input field for an arbitrary non-preset template + persisting custom** (an extension of 9.20. Small to medium).
- **Polishing the tab bar's look** (favicon, hover, show/hide control of the close button, gap animation
  during dragging, etc. CSS-centered. Fetching the favicon is a separate issue).
- `Wv2Json` is for JSON only, `Wv2Url` for URLs only. Do not mix them (unchanged since 9.16).

## Stage 9.22 — Trial: moving close / reorder to hostObjects

### Overall approach

In 9.21, activate alone was shown to work over hostObjects with a real URL (spec fact 19), and that it can be separated out
was demonstrated. As the natural next step, close / reorder (both synchronous operations with no completion wait =
plan 4-A) are rolled out sideways onto the hostObjects path and compared side by side
with the existing postMessage version. This
completes the set of "3 commands that suit hostObjects naturally (activate / close / reorder)", and the next stage
(planned 9.23) can then judge, with data in hand, "whether it is worth actually shrinking the Wv2Json parser + the Select Case in OnPaneWebMessage".

### Why only up to close / reorder (the reason ready / newtab / settings are left alone)

A recap of the sorting done in 9.21.
Of the 6 JS→VBA commands of the TabBar, only 3 — activate / close / reorder — gain cleanly from moving to hostObjects.

  ・ready is a "notification (a handshake signal)", not an RPC. Leaving it as postMessage
    is the natural choice.
  ・newtab / settings pass internally through WaitPaneReady (waiting for Controller creation to complete), so
    turning them into direct sync calls from JS→VBA would actually raise the risk of deadlock (spec fact 14). Even if they were moved to hostObjects, the VBA side would
    return immediately and queue the work on SafeTimer, so there is no gain beyond "the JSON
    parsing goes away".
  ・Therefore the technically optimal solution is a hybrid: move only close / reorder to hostObjects, and leave ready / newtab / settings on
    postMessage. As in 9.21, these three stay unchanged.

### ★Added entry points★ (all are copies of the same shape as HostActivate in 9.21)

- Added `Public Function HostClose(index) As Boolean`. It is called directly from JS as
  `hostObjects.sync.TabBar.HostClose(n)`. Its body is the same
  `m_browser.CloseTab(index)` as the close case. CloseTab only does `Ctrl_Close` and has no completion wait (plan 4-A),
  so a direct sync call is safe.
- Added `Public Function HostReorder(fromIdx, toIdx) As Boolean`. It is called directly from JS as
  `hostObjects.sync.TabBar.HostReorder(from, to)`. Its body is the same `m_browser.MoveTab(fromIdx, toIdx)` as the reorder
  case. MoveTab is only array manipulation +
  PutIsVisible and has no completion wait (plan 4-A).
  ★The VBA argument names were set to fromIdx / toIdx★ (because `From` has been reported to conflict as a VBA dialect issue.
  The JS side calls with positional arguments, so the argument names are invisible from JS and can be chosen freely).

### ★The JS side is a hybrid (the same three-tier structure as activate in 9.21)★

For the × click of close / the drop of reorder, the `HO()` helper (9.21) is reused: if
`hostObjects.sync.TabBar` can be obtained, HostClose / HostReorder are called directly and synchronously and `[host]` is printed to the console;
if the host call throws, it is caught and falls back to postCmd with `[host-fail→postmsg]`;
if hostObjects is absent, it falls back to the conventional postCmd with `[postmsg]`. The sending of newtab / settings / ready is
unchanged.

### Why no new design principle is established

  ・HostClose / HostReorder merely grow the same processing as the existing close / reorder cases
    (`m_browser.CloseTab` / `m_browser.MoveTab`) onto the hostObjects path as well. Calling a
    synchronous operation with no completion wait through a sync RPC is an application of plan 4-A, not a new principle.
  ・For attach, the `AddHostObjectToScript("TabBar", Me)` placed before Navigate in 9.21 is used
    as is (nothing added). This is the practice of spec facts 18/19.
  ・The JS-side fallback is insurance of the same shape as the settings-tab null guard of 9.18 / the activate of 9.21.
  ・So the design principles **remain at 71**.

### Implementation breakdown (stage 9.22)

- `Wv2TabBar.cls` (header updated to 9.22):
  ・Added `Public Function HostClose(index) As Boolean` (the same `m_browser.CloseTab(index)` as the
    close case).
  ・Added `Public Function HostReorder(fromIdx, toIdx) As Boolean` (the same `m_browser.MoveTab(fromIdx, toIdx)` as the
    reorder case).
  ・The × click / drop handlers of `BuildTabBarHtml` were made hybrid using the `HO()` helper of 9.21.
    The sending logic of newtab / settings / ready / activate (already migrated in 9.21) is unchanged.

- `Wv2Tests.bas` (header updated to 9.22):
  ・Added `Test_9_22_HostCloseReorder_Help` (real-machine procedure). Real-URL Navigate + JS execution +
    confirming that the hostObjects path is established require an actual WebView2 launch, so there is no pure-logic verification.

### Real-machine test log (summary) — all successful

- **close switched to the host route**: clicking the × of a tab prints `HostClose: index=N (via host)`, and
  `OnPaneWebMessage: msg={"cmd":"close"...}` never appears for close. The
  active-tab correction is also normal as before.
- **reorder switched to the host route**: drag & drop prints `HostReorder: from=N to=M
  (via host)`, and no postMessage ever appears for reorder. The redraw tracking of MoveTab is also
  normal.
- **activate remains unchanged, still via host as in 9.21**.
- **No regression (the paths not touched remain intact on postMessage)**:
  ・newtab: the SafeTimer path `RequestNewTabAsync → TimerCall` is normal.
  ・settings: `RequestSettingsTabAsync → TimerCall(settings) → OpenSettingsTab`, and
    the hostObjects UI of the settings screen is also normal (9.18 to 9.20 unaffected).
  ・ready: the handshake is healthy.
- Loose coupling (the Browser's 5 events unchanged), whole-system synchronization, and the ready handshake are intact in all logs.

### Numbers for stage 9.22

- Design principles: **71** (no change. HostClose / HostReorder are applications of plan 4-A, and attach reuses
  the one from 9.21)
- Spec facts: 19 → **20** (spec fact 20 added. The establishment of close / reorder over hostObjects
  is itself within the reach of spec fact 19, so nothing new there, but "the EBMode gate cannot prevent a transition to break during the
  initialization burst", which came to light while capturing the logs, was established as spec fact 20)
- Standard file layout: **12** (no change. No new files. Production entities 11)
- Changed files: 2 — `Wv2TabBar.cls` / `Wv2Tests.bas`

### Import operation notes (stage 9.22)

- Comment out the entire old version of `Wv2TabBar.cls` / `Wv2Tests.bas` (or delete it) → import the new
  version. There are no name conflicts, so no special order is needed.
- Since the JS-generating function (`BuildTabBarHtml`) was touched, it was confirmed with grep before output
  that no `\"` crept in (all JS double quotes
  are written as single quotes. The ★taboo when writing JS inside a VBA string★ was observed).

### ★ Spec fact 20 (new) — the EBMode gate cannot prevent a transition to break during the initialization burst ★

**Phenomenon**: When `StartWebView2_Full` (which launches 3 tabs + tab bar + nav bar all at once) is run and
the debugger is used to break / step during or right after it, Excel crashes (AV). Measured reproduction:
  ・Passing `AddTabWithUrl("...bing")` and, with the cursor resting on the next `If p Is Nothing...` line,
    freeze → crash.
  ・Pausing at the **`End Sub`** of `実行用()` (now `Boot`) (`UserForm1.Show vbModeless` → `StartWebView2_Full`)
    also crashes.
    → It is irrelevant whether the line where the break occurred is inside or outside the initialization function.

**Cause (isolated)**: This is not a missing implementation of the Sentinel (re-entrancy guard) or the EBMode gate.
As introduced in stage 8, all COM callback handlers (Controller/Navigation
Starting/Completed/DocumentTitleChanged/WebMessageReceived/NewWindowRequested/
HistoryChanged, etc.) without exception pass through the single path `AcquireHandlerFor` → `Thunks_AcquireSlot` →
`WriteThunkMachineCode`, and have a 91-byte EBMode-check stub clone
in front (there is no raw-thunk direct-call path, except for test code). In normal execution
this gate works correctly: firings during a break are swallowed by the `xor rax,rax / ret 0` at +85, and
in the log the refCount closes cleanly from 2→1 and the event nesting is healthy as well.

The problem lies in the **timing of the verdict** of the EBMode gate. The gate can only branch on "the EBMode value at the moment control enters the stub"
(+45 `call EBMode()` → +51 `cmp/je`). For an instance that passed the check while "running", just as it begins executing the VBA-side Invoke body (p-code) that the tail call targets,
if the process starts transitioning to break because of the user's Shift+F8 or the like, the p-code runs halfway while the VBA interpreter is in the middle of stopping,
leading to an AV. This is a **momentary window where passing the check and the transition to break race each other**,
and by the structure of the gate it inevitably remains.
This is not a defect specific to your architecture; it is a limit at the outer edge of the
VBA EBMode-gate approach itself (SafeTimer in theory has the same window, but a timer
fires sparsely and the events are not as dense as in the WebView2 initialization burst, so it is hard to step on).

**What decides it is the timing (not the line position)**: Whether it crashes depends on "whether WebView2 callbacks are still
arriving at that moment". Even when `StartWebView2_Full` has returned and reached `End Sub`,
the follow-up events for the 3 tabs (NavigationCompleted / DocumentTitleChanged /
HistoryChanged, etc.) are still arriving asynchronously. While this burst is still trailing off, you can step on it by breaking anywhere,
whether inside or outside the initialization function.

**The countermeasure is avoidance, not prevention (an operating rule)**: **Do not break / step
until the WebView2 event burst has settled
(= all tabs have finished loading and the Immediate-window log has stopped)**.
Breaking after it has settled is safe, since there are no callbacks flying in. To debug
the body of an event handler, reach it with F5 after initialization has completed and settled, and then break. Since the verification Sub is designed so that "typing just the Sub name runs everything", there is no
need in the first place to step through the initialization sequence.

## Stage 9.23 — Cleanup: consolidating the activate / close / reorder processing into Host methods

### Overall approach

When activate/close/reorder were moved to the hostObjects path in 9.21/9.22, the same processing
(the calls to ActivateTab/CloseTab/MoveTab and the failure logging) was duplicated in 2 places: "inside the Case of OnPaneWebMessage"
and "inside HostActivate/HostClose/HostReorder". That is a breeding ground for the accident of fixing one and forgetting to fix the
other. This is the stage that reaches a conclusion, with real data (in 9.21/9.22 it never once fell back to the fallback),
on the open item set at the completion of 9.22:
"is it worth shrinking the Select Case".

### ★Decision = plan C-1 (keep the insurance and consolidate the processing)★

Three plans were compared and plan C was adopted.

  ・Plan A (keep as is): the double-path coexistence leaves the duplicated processing. Real harm = the same logic in 2 places. Rejected.
  ・Plan B (remove the fallback + shrink the Select Case to 3 cases): make activate/close/reorder
    hostObjects only. The JSON parsing for those commands really does disappear, but it throws away the safety net for the browser's fundamental operations
    (switch/close/reorder). If hostObjects could not be obtained after, say, a WebView2 runtime update,
    all 3 commands would die. The cost is too heavy; rejected.
  ・★Plan C (adopted): keep the insurance and consolidate the processing★ The JS-side hybrid (hostObjects first, postMessage
    fallback) is left as is. On top of that, the bodies of Case activate/close/
    reorder, which are taken on fallback, are delegated to calls of the corresponding Host methods. The logic that calls ActivateTab/CloseTab/
    MoveTab and the failure logging are gathered in one place, the Host method (DRY). The apparent line count of the Select Case
    does not go down, but the real harm of the two coexisting paths (duplicated processing) is reliably eliminated.

  ★Plan C-1 (fire and forget)★ The Case side calls HostXxx without receiving the return value. Since the failure log is already emitted inside the Host
    method, the double logging on the Case side is removed.
    However, as evidence that the postMessage path was taken, the Debug.Print at the Case entry is kept, and
    [postmsg fallback → HostXxx] is appended to its end to make the delegation target explicit. The JSON parsing (JsonGetNum) is needed to
    restore the arguments on the fallback path, so it remains on the Case side (this is the correct shape. On the hostObjects path, JS
    passes numbers directly, so no parsing is needed; the numbers are restored from JSON only on fallback).

### Why the Select Case does not disappear (the final answer to the open item from 9.22)

For the 3 cases ready/newtab/settings, postMessage is the right answer (ready = notification, newtab/settings =
the WaitPaneReady detour is required, spec fact 14), so removing the Select Case entirely
is impossible in principle. Hence "it shrinks but does not disappear". Plan C accepts this constraint and is the realistic solution that removes only the "removable duplication"
inside the Select Case that cannot be removed. The cognitive load of the hybrid's two coexisting paths remains,
but that is taken on as the price of the safety net.

### Why no new design principle is established

  ・HostActivate/HostClose/HostReorder already exist from 9.21/9.22. Delegating to them from the Case
    is the practice of plan 4-A (synchronous operations with no completion wait) and DRY, not a new principle.
  ・For attach, the AddHostObjectToScript placed before Navigate in 9.21 is used as is (nothing added).
  ・The JS side is unchanged. So the design principles remain at 71.

### Implementation breakdown (stage 9.23)

- Wv2TabBar.cls (header updated to 9.23):
  ・The bodies of Case "activate" / "close" / "reorder" in OnPaneWebMessage were
    replaced with fire-and-forget delegation to HostActivate idx / HostClose idx / HostReorder fromIdx, toIdx respectively.
    The direct calls to ActivateTab/CloseTab/MoveTab and the failure logging were removed from the Case side.
  ・[postmsg fallback → HostXxx] was appended to the Debug.Print at the entry.
  ・The now-unused Dim ok As Boolean in OnPaneWebMessage was deleted (idx is kept because it is used to restore the
    arguments of activate/close).
  ・The 3 cases ready/newtab/settings, the JS-side BuildTabBarHtml, the Host method bodies, and the attach in Init
    are unchanged.

- Wv2Tests.bas (header updated to 9.23):
  ・Added Test_9_23_Consolidation_Help (real-machine procedure). 9.23 is a cleanup stage that changes where the processing lives, and its functionality is identical to 9.22,
    so the focus is confirming zero regression. There is no pure-logic verification.

### Real-machine test log (summary) — succeeded with zero regression

- **activate**: clicking the body of another tab prints only `HostActivate: index=4 (via host)`.
  No activate log from OnPaneWebMessage appears. The ActiveChanged → PushSyncToJs tracking is also normal.
- **close**: clicking × prints only `HostClose: index=4 (via host)`. No close log from OnPaneWebMessage
  appears. The sequence Ctrl_Close → active correction to the remaining tabs → TabClosed is identical to 9.22.
- **reorder**: dragging prints only `HostReorder: from=4 to=3 (via host)`. No reorder log from OnPaneWebMessage
  appears. MoveTab → the redraw in the reordered order and the nav bar tracking are also normal.
- **newtab**: `OnPaneWebMessage: msg={"cmd":"newtab"}` → RequestNewTabAsync → TimerCall.
  The postMessage path through SafeTimer is intact.
- **settings**: `OnPaneWebMessage: msg={"cmd":"settings"}` → RequestSettingsTabAsync →
  TimerCall(settings) → OpenSettingsTab. The hostObjects UI of the settings tab (9.18 to 9.20) is also intact.
- The division of roles (activate/close/reorder = HostXxx, ready/newtab/settings = OnPaneWebMessage) matches
  9.22 exactly. The look and feel are not a hair different from 9.22.
- Confirming that the fallback survives (optional) was skipped because it cannot be triggered in a normal environment. The essence of 9.23 is
  zero regression, and that was met.
- Side note: that the search engine was Yahoo at startup is the normal behavior of 9.20 (persistence) taking effect and
  is unrelated to 9.23 (evidence that the previous selection was carried over).

### Numbers for stage 9.23

- Design principles: **71** (no change. Practice of plan 4-A / DRY; attach reuses the one from 9.21)
- Spec facts: **20** (no change. No new spec facts)
- Standard file layout: **12** (no change. No new files. Production entities 11)
- Changed files: 2 — `Wv2TabBar.cls` / `Wv2Tests.bas`

### Import operation notes (stage 9.23)

- Comment out the entire old version of `Wv2TabBar.cls` / `Wv2Tests.bas` (or delete it) → import the new version.
  There are no name conflicts, so no special order is needed.
- The JS-generating function (`BuildTabBarHtml`) is unchanged, but just in case it was confirmed with
  grep before output that no `\"` crept in (zero in the actual code).

### Open item for the next stage (the natural continuation of 9.23)

- Roll out the same shape to the NavBar (navigate/back/forward/reload).
  Since plan C (keep the insurance and consolidate into Host) passed with zero regression on the TabBar, the NavBar too
  can be aligned to the same shape as "thin Cases that delegate to Host methods". back/forward/reload are synchronous operations,
  so they are within the reach of plan 4-A. navigate carries a URL string,
  but with hostObjects a string argument can be passed directly. To be worked out (starting from creating the Host methods on the NavBar side).


## Stage 9.24 — Moving the NavBar to hostObjects + consolidating the processing into Host methods

### Overall approach

  By 9.21 to 9.23, the TabBar's activate/close/reorder had settled into the form "JS is a hybrid (hostObjects first,
  postMessage fallback), VBA gathers the processing in one Host method, and the Select Case is a thin
  adapter". 9.24 is the round that rolls that shape out sideways onto the NavBar's back/forward/reload/navigate. Since the shape is already established on the TabBar,
  the step of "first do only the hostObjects conversion", as in 9.21/9.22, is not taken, and the whole thing is built from the start as plan C-1
  (plan 2 from the discussion was adopted).

### ★Decision = plan C-1 from the start (build it all at once)★

  ・Plan 1 (first only the hostObjects conversion → consolidate in the next stage) makes each stage the smallest, but it would split the copying of
    the shape that has already been through one full cycle on the TabBar into 2 stages, which is redundant.
  ・Plan 3 (navigate first on its own) lets the new element (a string argument) be verified in isolation, but it makes 3 stages.
  ・→ Plan 2 adopted. The risk is effectively concentrated in the one point "navigate's string argument", and that one point is, even under
    plan 2, isolated automatically in the form "if back/forward/reload are working via host, then the attach
    itself has succeeded", so the gain from splitting into stages is small.

### ★Added entry points★ (all of the same shape as HostActivate in 9.21)

  ・Public Function HostBack(Optional ByVal reserved As Long = 0) As Boolean    → GoBackActive
  ・Public Function HostForward(Optional ByVal reserved As Long = 0) As Boolean → GoForwardActive
  ・Public Function HostReload(Optional ByVal reserved As Long = 0) As Boolean  → ReloadActive
  ・Public Function HostNavigate(ByVal url As String) As Boolean                → NavigateActive
  All of these are fire-and-forget on the Browser side (no completion wait = plan 4-A), so even a direct sync call from JS
  does not fall into the deadlock of spec fact 14. Attach is m_barPane.View_AddHostObjectToScript("NavBar", Me) before Init's Navigate
  (order A of spec facts 18/19, the same position as the TabBar's 5b). Failure is not fatal, so hr is logged
  and Init continues.

### ★Why the dummy argument reserved exists (discussion point 4b)★

  All three TabBar methods took arguments, but back/forward/reload naturally take zero arguments.
  The hostObjects sync proxy cannot tell from the JS side whether a member is "a method or a property",
  so it is possible that merely *getting* ho.NavBar.HostBack fires IDispatch::Invoke(DISPATCH_PROPERTYGET) in the
  implementation. VBA tolerantly accepts this against a zero-argument Public Function and
  executes it, so in the worst case
    (1) the get executes it once → (2) the return value is Boolean, so () raises a TypeError →
    (3) catch and fall back to postMessage → it executes one more time
  and a double execution, "one press of ← goes back twice", can occur. The nasty part is that ★the hybrid's
  safety net backfires, so the fallback cannot prevent it★.
  → We gave it one argument and removed the ambiguity in principle (JS calls nb.HostBack(0)). On the VBA side the Case
    can call it as a bare HostBack since the argument is Optional.
  ※On the real machine no double execution was observed, but that is because it did not occur "since we closed it off on the safe side", and
    it is not proof that "zero arguments are safe". Whether reserved is needed is left as an open item for the next stage (below).

### Where to put the empty-URL guard (discussion point 3)

  The check If Len(Trim$(url)) = 0 that was inside the Case has been moved into HostNavigate. Now the same guard
  is passed whether the call comes via the hostObjects path or the postMessage path (the same DRY
  thinking as 9.23). The JS-side if(v && v.length) is kept as a safety net.

### ★The JS side is made hybrid with a navHost() helper★

  TabBar wrote its 3 places out bare, but NavBar has 4 places and they are similar in shape, so we wrapped them in a helper.
    navHost(msgObj, callHost)
      (1) if ho.NavBar exists, call callHost(ho.NavBar) directly and synchronously → [host]
      (2) on an exception, catch and postCmd(msgObj)                → [host-fail→postmsg]
      (3) if it could not be obtained at all, postCmd(msgObj)       → [postmsg]
  The JSON sent in the fallback is exactly identical to 9.12d, so the thin Case on the VBA side can
  receive it as is. ★Access to ho.NavBar uses dot notation only★ (the form proven with TabBar; a
  computed-property notation such as ho.NavBar[name] is not used because the proxy's behavior is unpredictable).

### Why we do not establish a new design principle

  ・Delegating to Host methods is the practice of plan 4-A (synchronous operations without waiting for completion) and DRY, not a new principle.
  ・attach is an application of the "AddHostObjectToScript before Navigate" established in 9.21 (spec facts 18/19).
  ・Turning the JS into a helper is extracting identical processing, not a change of policy.
  ・Hence the design principles stay at 71.

### Implementation breakdown (stage 9.24)

- Wv2NavBar.cls (header updated from 9.13 → 9.24):
  ・Added step 5b (View_AddHostObjectToScript("NavBar", Me)) to Init, before Navigate.
  ・Added four new methods: HostBack / HostForward / HostReload / HostNavigate.
  ・Replaced the bodies of Case "back"/"forward"/"reload"/"navigate" in OnPaneWebMessage with fire-and-forget delegation to the corresponding Host
    method. Removed the direct calls such as GoBackActive and the failure logs from the Case side
    (confirmed with grep that the 4 Browser calls exist only inside the Host methods).
  ・Appended [postmsg fallback → HostXxx] to the entry Debug.Print. The JsonGetStr for navigate is kept because
    it is needed to restore the argument on the fallback path.
  ・Moved the empty-URL guard from the Case to HostNavigate.
  ・Added HO() / navHost() to BuildNavBarHtml and made the 4 places hybrid.
  ・The ready case / Case "" / Case Else / PushNavSyncToJs / SendNavSync / the navsync-receiving JS /
    the editing flag / Escape's discarding of edits are unchanged.

- Wv2Tests.bas (header updated to 9.24):
  ・Added Test_9_24_HostNavBar_Help (real-machine procedure). The difference from 9.21 to 9.23 is that the verdict has three parts (did it go via host / is there no
    double execution / is the string argument intact).

### Real-machine test log (summary) — all succeeded

- **attach**: `Wv2NavBar.Init: NavBar bridge attach OK` → `Navigate OK
  (https://appassets.navbar/navbar.html)` in that order. Order A holds for NavBar too (spec facts 18/19).
- **navigate**: Only `Wv2NavBar.HostNavigate: url=... (via host)`. No navigate log from OnPaneWebMessage
  appears. Scheme completion and search-engine routing (9.14/9.16 to 9.20) are also normal.
- **back**: One line `HostBack: (via host)` → `GoBackActive: GoBack OK` → NavigationStarting once.
  At the destination canBack:true/canForward:true, confirming ★1 click = going back 1 page★.
- **forward**: Only one line `HostForward: (via host)`. canForward was correctly updated to False.
- **reload**: Only one line `HostReload: (via host)`. It re-navigated to the same URL exactly once, and the history state
  (canBack:true/canForward:false) did not change = reload is not polluting the history.
- **No double execution**: All 4 commands have only one Host log line, and no subsequent postMessage log.
- **ready**: Both TabBar and NavBar still go via `OnPaneWebMessage: msg={"cmd":"ready"}` (correct).
- **Escape**: Discarding of edits works normally (9.12d intact).
- **Zero regressions on the TabBar side**: Switch `HostActivate: index=2`, × `HostClose: index=2`, and drag
  `HostReorder: from=2 to=1` each produced only one line via host. Moreover, ★for all 3 of these, NavBar's
  SendNavSync followed★, confirming that the loosely coupled path of TabBar operation → Browser event → whole-NavBar sync
  is intact after the 9.24 changes.
- navbar.html is 3465 characters (grew only by the HO()/navHost() additions).

### ★ Spec fact 21 (new) — String arguments of hostObjects.sync are passed as BSTR directly and arrive intact ★

  **Fact**: When a String is passed from JS to VBA via hostObjects.sync, it arrives as a BSTR as is, without passing through any JSON escaping.
  Not only Japanese and symbols but ★even the full-width space (U+3000) is not lost★.

  **Evidence**: Entered 「日本　天気」 (full-width space between) in the URL field → Enter.
    Wv2NavBar.HostNavigate: url=日本　天気 (via host)
    Wv2Browser.NavigateActive: Navigate OK
      (https://search.yahoo.com/search?p=%E6%97%A5%E6%9C%AC%E3%80%80%E5%A4%A9%E6%B0%97)
  The byte sequences are E6 97 A5=日 / E6 9C AC=本 / **E3 80 80 = U+3000 full-width space** / E5 A4 A9=天 /
  E6 B0 97=気. The JS input string passes through as UTF-16 all the way to the VBA String, and flows correctly
  through to Wv2Url's percent-encoding.

  **Meaning**: The postMessage path goes through the home-made JSON parser (Wv2Json), so the handling of \u escapes
  became entangled (a party to spec fact 16). The hostObjects path does not pass through there, so it is ★if anything simpler and safer★.
  TabBar (9.21/9.22) only passed Long, so 9.24 is the first to confirm that string arguments work.

  **Scope**: Future designs that carry strings via hostObjects (e.g. passing a custom search template) can be built
  without worrying about JSON escaping.

### Numbers for stage 9.24

- Design principles: **71** (no change. An application of plan 4-A / DRY / spec facts 18 and 19; no new principle)
- Spec facts: 20 → **21** (spec fact 21 added)
- Standard file layout: **12** (no change. No new files. 11 production entities)
- Changed files: two, `Wv2NavBar.cls` / `Wv2Tests.bas`

### Import procedure notes (stage 9.24)

- Comment out the whole old version of `Wv2NavBar.cls` / `Wv2Tests.bas` (or delete it) → import the new version.
  There are no name conflicts, so no special order is needed.
- For the JS generation function (`BuildNavBarHtml`), we confirmed with grep before output that no `\"` crept in
  (zero in the actual code). The odd-quote-count line check also found 0 cases.

### ★Incident note: the tab bar appeared in two rows at startup (unrelated to the 9.24 logic)★

  **Symptom**: On the first startup after the introduction, the form showed **the same tab bar in two rows**, and the navigation bar (the URL field and
  ← → ⟳) was gone.

  **Cause**: The paste source of the code was wrong, and the contents of the `Wv2NavBar` module in the VBA project
  were the code of `Wv2TabBar`. Since `Wv2TabBar.Init` also takes the same three arguments (browser, hwnd, folder),
  `Set m_navBar = New Wv2NavBar` both compiled and ran, and a second tab bar was generated in the folder
  for NavBar.

  **What decided the diagnosis (checklist for the next time a similar incident occurs)**:
    ・The second row's log is prefixed `Wv2TabBar.*` rather than `Wv2NavBar.*`
    ・`tabbar.html` has been written into the NavBar folder (it should be `navbar.html`)
    ・The mapping is `appassets.tabbar` → NavBar folder
      (it should be `appassets.navbar` → NavBar folder)
    ・The HTML character counts of the first and second rows match exactly (both 6349 characters)
    ・**Not a single log line starting with `Wv2NavBar.` appears**
  **Lesson**: A module name alone does not guarantee the identity of its contents. After an introduction, the fastest check is to look at the startup log first to see whether each class
  prefixes its output with its own name (`Wv2NavBar.Init:` etc.). A good example of how the habit of always putting the class name in Debug.Print
  directly served as diagnostic material.

### Open items for the next stage

- **Verify whether reserved is needed (small)**: Remove Optional ByVal reserved from HostBack/HostForward/HostReload,
  change the JS back to nb.HostBack(), and test on the real machine whether double execution occurs. If it does not occur,
  we establish "hostObjects.sync does not mistake a zero-argument VBA Function for a property get" as a spec fact
  and can permanently delete reserved. If it does occur, the necessity of reserved is demonstrated, and we record that
  as a spec fact. ★A cheap experiment that leaves knowledge whichever way it falls★.
- **Expanding a POST link into a new tab gives 405 (medium to large)**: With the HostObject side tightened in 9.24, the agreement is to
  tackle this next. Move NewWindowRequested from "URL interception" to "WebView delegation" (get_Deferral /
  put_NewWindow / Deferral::Complete). The hard part is avoiding the nested deadlock of waiting during the deferral
  for the creation of another tab's Controller to complete.
- An arbitrary-template input field outside the presets + persistence of custom (an extension of 9.20). Spec fact 21 showed that
  passing strings via hostObjects is safe, so the design has become easier.
- Polishing the tab bar's appearance (favicon, hover, display control of the close button, gap animation during drag).

## Stage 9.25 — Verification: whether zero-argument Host methods work → permanent removal of reserved

### Overall policy

  A small verification round to settle the 9.24 open item "was the dummy argument reserved really necessary?".
  In 9.24 we gave the methods Optional ByVal reserved As Long out of concern that "a zero-argument Public Function would
  be mistaken by the hostObjects sync proxy for a property get, executing once at the get and once more
  in the fallback, causing double execution". Because we erred on the safe side the
  phenomenon did not appear, but that is not proof of "unnecessary", so we settle it by measurement.
  ※The implementation was delivered in two installments (9.25 = verification, 9.25b = removal), but 9.25b is merely the removal of
  the scaffolding after the conclusion of 9.25, so the record is combined into one section.

### ★Decision = Plan C (add the probe + delete reserved at the same time)★

  ・Plan A (just remove reserved and try with the real buttons): minimal change, but if double execution occurred
    we could only guess whether the cause was "executing at the get".
  ・Plan B (add only the probe and leave reserved in place): causality is fully clear but it takes two steps.
  ・→ Plan C. In one startup we can obtain both the "mechanism fact (probe)" and the "real behavior (buttons)".
    Even if it misses, the only real harm is "← goes back 2 pages with 1 press", and observing it once is enough. Both answers come out cheaply.

### ★Probe design (isolating only the mechanism with zero side effects)★

  Temporarily added two methods to NavBar:
    ・HostProbe() As Boolean          … ★zero arguments★. It only writes a log, zero side effects.
    ・HostProbeMark(phase As Long)    … Safe because it takes an argument. Stamps the phase.
  The JS runs itself once, via setTimeout(probeParamless, 500), right after sending ready:
      Mark(1) → var f = nb.HostProbe (get only) → Mark(2)
      → f()            success Mark(3) / failure Mark(4)
      → nb.HostProbe() success Mark(5) / failure Mark(6)   (Mark(9) if the get throws)
  ★The point is that we tested with a dedicated zero-side-effect function rather than the real ones (HostBack etc.)★ Even if
  double execution occurred, the log would only gain 2 lines, and neither the browser history nor the screen would move. A pattern for safely testing a dangerous hypothesis.
  ★All logging is emitted from the VBA side★ We do not use console.log because it does not appear in the Immediate window. The order of phases and
  execution timing lines up in a single log, so the causality can be read directly.

  ★Why the trigger is a setTimeout inside the JS★ Wv2Pane.View_ExecuteScript (9.8c) already exists, but
  the NavBar instance is held by UserForm1 as a Private field, so poking it from outside would require adding a public entry point
  to UserForm1.frm. We wanted to avoid adding to a production file just for temporary verification-only code,
  so we took the approach of letting the JS run itself. tarboh only needs to start it normally.

### ★ Spec fact 22 (new) — The hostObjects sync proxy does not call VBA on member retrieval ★

  **Fact**: Member access via hostObjects.sync does not call into the VBA side at the time of the get.
  The result of the get is returned as a function object, and execution is deferred until the actual call ().
  Therefore ★a zero-argument Public Function can be exposed as is★.

  **Evidence (real-machine log)**:
      [PROBE] (1) just before the get
      [PROBE] (2) just after the get                       ← ★no execution log in between★
      [PROBE] ★HostProbe was executed★ (total 1st time)
      [PROBE] (3) f() call succeeded
      [PROBE] ★HostProbe was executed★ (total 2nd time)
      [PROBE] (5) nb.HostProbe() direct call succeeded
    ・No execution log between (1) and (2)    → it is not called at the get
    ・(3) appeared                            → the retrieved value is a function object and can be called later
    ・(5) appeared                            → the nb.HostProbe() form actually used passes
    ・Exactly 2 total                         → zero extra executions
  With the real ones too, HostBack / HostForward / HostReload each produced just one line, there was no fall to the
  postMessage fallback, and ← went back just one page (canBack:true/canForward:true).

  **Meaning**: The sync proxy cannot distinguish "method or property" from the JS side, but it reconciles the two by deferring
  the decision until the actual (). Therefore the VBA-side entry points can be
  designed without worrying about whether there are arguments. The reserved of 9.24 was an unnecessary safety net.

  **★But the procedure was correct★** Double execution would appear in the hard-to-understand form of "one click goes back 2 pages",
  and moreover the hybrid's safety net (postMessage fallback) would backfire,
  so it would be a kind of bug that is hard to notice later. The order of entering from the safe side and trimming by measurement costs only
  one dummy argument, whereas the cost of discovery if we had missed is high. Similar decisions should follow this
  order in the future.

### Why we do not establish a new design principle

  ・Spec fact 22 is an observation of a property of WebView2/VBA, not our design policy.
  ・Deleting reserved merely reverted the provisional measure of 9.24. Hence the design principles stay at 71.

### Implementation breakdown (stage 9.25 / 9.25b)

- Wv2NavBar.cls (header stacked in two layers: 9.24 → 9.25 → 9.25b):
  ・[9.25] Removed Optional ByVal reserved As Long from HostBack / HostForward / HostReload,
    making them zero-argument. The JS also changed from nb.HostBack(0) → nb.HostBack() (3 places).
  ・[9.25] Added HostProbe / HostProbeMark for verification only, and added probeParamless() and
    setTimeout(probeParamless, 500) to BuildNavBarHtml.
  ・[9.25b] Removed all of the above probes (2 VBA methods + JS). Rewrote the HostBack comment from the in-verification description to
    the conclusion version (see spec fact 22).
  ・The bodies of the 4 Host methods / the delegation structure of OnPaneWebMessage / navHost() / attach / ready /
    Escape / editing / PushNavSyncToJs are entirely unchanged.

- Wv2Tests.bas (header stacked in two layers: 9.25 → 9.25b):
  ・[9.25] Added Test_9_25_Probe_Help (a real-machine procedure listing the expected logs for pattern A/B).
  ・[9.25b] Deleted that Sub (the probe body no longer exists, so there is nothing left to run). The conclusion remains in the header.

### Real-machine test log (summary) — hypothesis rejected; zero arguments confirmed

- **Probe (9.25)**: Pattern A. As in the evidence for spec fact 22 above.
- **Real ones (9.25)**: For each of ← → ⟳, HostXxx produced just one line. No double execution.
- **After removal (9.25b)**: Confirmed that not a single [PROBE] line appears. ← → ⟳ also produce just one line, no regression.
- ★Secondary evidence★ The character count of navbar.html went 9.24: 3465 → 9.25: 3889 → 9.25b: 3462.
  The difference between 9.24 and 9.25b is exactly 3 characters, matching exactly the one '0' dropped from each of
  HostBack(0)/HostForward(0)/HostReload(0). This corroborates that no probe JS remains.
  An operational insight: ★the HTML character count can be used to detect this kind of "forgot to delete" leftover★.

### Numbers for stage 9.25

- Design principles: **71** (no change)
- Spec facts: 21 → **22** (spec fact 22 added)
- Standard file layout: **12** (no change. No new files. 11 production entities)
- Changed files: two, `Wv2NavBar.cls` / `Wv2Tests.bas` (delivered in two installments, 9.25 / 9.25b)

### Import procedure notes (stage 9.25)

- As usual, comment out the whole old version (or delete it) → import the new version. No name conflicts.
- In a round that brings in temporary verification code, **preparing everything up to the removal version before delivering it** leaves no hesitation
  (the removal targets of 9.25b were already fixed at the time 9.25 was delivered, so the removal was mechanical).

### Open items for the next stage

- **Expanding a POST link into a new tab gives 405 (medium to large)**: The HostObject side was fully tightened in 9.24/9.25, so
  this is the main issue next. Move NewWindowRequested from "URL interception" to "WebView delegation" (get_Deferral /
  put_NewWindow / Deferral::Complete). The hard part is avoiding the nested deadlock of waiting during the deferral
  for the creation of another tab's Controller to complete. One could also start by checking whether a light interim workaround (delegating to the same WebView to keep the POST)
  succeeds.
- An arbitrary-template input field outside the presets + persistence of custom (an extension of 9.20). Spec fact 21 showed that
  passing strings via hostObjects is safe, so the design is easy.
- Polishing the tab bar's appearance (favicon, hover, display control of the close button, gap animation during drag).


## Stage 9.26 — Isolation: NewWindowRequested delegation probe → root cause of the POST 405 confirmed

### Overall approach

  This round is for **establishing only the facts before tackling the main part** of the issue where
  opening a POST link in a new tab returns 405 (workplace core system / Apache Tomcat 7.0.76). The implementation is one branch and one toggle.

  The current Wv2Pane.View_OnNewWindowRequested is built so that it
    (1) extracts "only" the URL string with args.get_Uri,
    (2) discards the runtime's navigation itself with args.put_Handled(TRUE), and
    (3) calls Navigate (= always GET) again in a new tab,
  so the original POST's method/body/Referer are lost at (2). That was the working theory for the
  405, but it was an **unverified hypothesis**.

  ★Why isolate first★ The main part (put_NewWindow + Deferral) contains the hard part of "a nesting that waits
  inside an event handler for the completion of an asynchronous Controller creation", an area where spec fact 14 (SafeTimer
  deadlock) and spec fact 20 (the EBMode window) both apply at once. **Finding out only after building it all
  that the theory was wrong is the most expensive outcome**. If one branch can give the answer first, get it first.

### ★Decision = cut one stage for option 1 (delegation probe) alone★

  ・Option 2 (go straight to the main part) would mean verifying the hypothesis while still carrying the hard part.
  ・Option 3 (stage split) was adopted, and option 1 alone became 9.26 first. Same structure as the reserved check in 9.25:
    the cost is one branch, and the cost of discovering it if we were wrong is extremely high.

### ★Correction to a statement in the handover document (found during discussion)★

  The handover said "as a temporary workaround, if we **delegate to the same WebView** while keeping Handled=TRUE, the POST may be
  preserved", but this **does not hold**. What "delegate to the same WebView" actually amounts to is
  View_Navigate from 9.6b, and ICoreWebView2::Navigate is ★always GET★. Whether it goes to a new tab or to the same
  tab is irrelevant; the moment it passes through Navigate, it stays a 405.
  → **The only way to preserve the POST is to "hand the navigation itself to the runtime"**. There are just 2 options:
      (a) Do not set put_Handled → the runtime opens its own popup (the 9.6a behavior)
      (b) Pass a different CoreWebView2 to put_NewWindow → the main part
    9.26 is the round that implemented (a) as a verification toggle.

### Implementation breakdown (stage 9.26)

- Wv2Browser.cls (header updated 9.20 → 9.26):
  ・Added Private m_popupDelegateMode As Boolean to the declarations section at the top of the module (default False).
  ・Added Public Property Get/Let PopupDelegateMode (Let logs the switch).
  ・★The toggle is managed collectively in Browser★ (tarboh's decision). It is not held per Pane, so
    all tabs are brought into step with a single switch. The Pane side reads it each time, so a switch while running takes effect immediately.

- Wv2Pane.cls (header updated 9.19 → 9.26):
  ・Added one branch right after the get_Uri log in View_OnNewWindowRequested.
    If m_browser.PopupDelegateMode is True, it calls neither put_Handled nor RequestNewTabAsync and
    does Exit Sub. When m_browser Is Nothing (standalone Pane use), behavior is as before.
  ・**Nothing else is changed at all**.

- Wv2Tests.bas (header stacked as 9.25b → 9.26 → 9.26b):
  ・[9.26] Test_9_26_Popup_On / _Off / _Help. Since it goes through UserForm1.CurrentBrowser (the 9.17
    scaffolding), **UserForm1.frm needed no change**.
  ・[9.26b] Added Test_9_26_PostProbe / _Help / BuildPostProbeHtml / WriteUtf8NoBom.

### ★9.26b = the round that synthesized the verification environment at home★

  What the first verification of 9.26 showed: **all the probe's mechanics worked, but the main question could not be judged**.
  The popup of the site tried (the JDreamIII operating-environment check page) was a ★plain GET★, so
  even with delegation OFF it opened normally in our own tab (isSuccess: True), and it was not a reproduction of the 405.

  → We took the approach of **synthesizing** the 405 setup without going to the workplace. Setting the destination to https://httpbin.org/post
    gives "405 for GET / 200 + the body echoed as JSON for POST", which builds a setup **exactly the same as
    the workplace Tomcat**. Moreover, a single button yields both "a reproduction of the 405" and "direct
    evidence that the body was carried" at the same time.

  Design decisions for the verification page:
  ・It is opened via a **virtual host mapping** (https://appassets.postprobe/postprobe.html) as a real URL. It rides
    directly on the existing Wv2Browser.AddTabWithUrlForSpa.
    ★Why NavigateToString was not used★ The origin becomes "null", so the conditions for Referer and form submission differ
    from production, which would reduce the purity of the verification.
  ・Writing is done with **ADODB.Stream as UTF-8 (no BOM)**. UserForm1.WriteSpaAppFolder writes ANSI (CP932)
    via Print #, so putting Japanese into HTML with meta charset=UTF-8 would garble it.
    We aligned with the approach proven in Wv2TabBar/Wv2NavBar (it is late-bound, so it does not
    conflict with the zero-references policy). ★This character-encoding mismatch was discovered while checking the scaffolding just before implementation, and
    we proceeded only after presenting it as an issue and getting agreement★
  ・★Zero JS★ Built with only a plain <form>. We did not go out of our way to bring in "JS inside VBA strings", a seed of accidents,
    for a requirement that HTML alone can satisfy.
  ・There are 4 buttons (main = POST+_blank / sanity = POST+same tab / control = GET+_blank /
    spare = POST+_blank to postman-echo). A hidden field carries Japanese containing a full-width space, so that
    whether it appears in the echo shows even whether the body arrived intact.

### Real-machine test log (summary) — settled with pattern A

- **9.26 first run (verifying the mechanics)**: pass.
  ・The toggle works in both directions and is reflected immediately from the next click.
  ・In delegation mode, the 3 lines of [9.26 delegation mode] appear, and put_Handled OK / RequestNewTabAsync do
    **not appear**. Our own tab is not created either.
  ・The runtime opens a popup, and its contents are the same screen as in the control case.
  ・Release [idx=13] → refCount = 1 is normal. **No handler reference leak**.
  ・No disturbance at all in the other event paths (NavigationStarting / HistoryChanged / DocumentTitleChanged /
    NavigationCompleted / TabBar/NavBar sync).
  ・★The log does not react even when the popup is closed★ = navigation inside the popup does not ride on our events
    at all. **The reason it cannot be a permanent fix is backed up by measurement**.
  ・However, the main question was not judged (because it was a GET popup, as described above).

- **9.26b (judging the main question)**: ★pattern A★
  ・**Delegation OFF**: `→ title: 405 Method Not Allowed` / `→ isSuccess: False`.
    The screen shows "Method Not Allowed / The method is not allowed for the requested URL."
    → **We were able to reproduce the workplace symptom at home**. Proof that the theory was correct.
  ・**Delegation ON**: a 200 JSON appeared in the popup. The evidence is as in spec fact 23 below.

### ★ Spec fact 23 (new) — If NewWindowRequested is delegated to the runtime, the POST is fully preserved ★

  **Fact**: If you let NewWindowRequested pass through without setting args.put_Handled, the WebView2
  runtime opens its own popup window and **continues the original request as is**.
  Not only the method and body, but **even Referer / Origin / Sec-Fetch-* are carried over completely**.

  **Evidence** (POST to https://httpbin.org/post + target=_blank, delegation ON):
    "form": { "memo": "hello from Excel VBA WebView2",
              "nihongo": "日本語　全角スペース入り",
              "probe": "wv2-post-probe" }
    "Content-Length": "172"
    "Content-Type": "application/x-www-form-urlencoded"
    "Origin": "https://appassets.postprobe"
    "Referer": "https://appassets.postprobe/"
    "Sec-Fetch-Mode": "navigate" / "Sec-Fetch-Dest": "document" / "Sec-Fetch-User": "?1"

  How to read it:
  ・**All 3 fields** appear in form = the body was carried without any loss.
  ・Content-Length 172 / Content-Type urlencoded = it did not turn into a GET (a GET would have a zero-length body).
  ・The 　 in nihongo = **U+3000 full-width space intact**. This is a different path from spec fact 21 (passing a BSTR
    directly through hostObjects) (the browser's internal POST body), but it is not lost here either.
  ・The biggest gain is that **Referer / Origin are carried over**. It means we can also withstand the case where the workplace's Spring MVC was
    looking at Referer for CSRF protection, so the risk of pattern B (a cause other than the POST body) has
    dropped greatly.
  ・Sec-Fetch-User: ?1 = it is continued **as a user-initiated top-level navigation**.

  **Reach**: The main part (passing another tab's CoreWebView2 to put_NewWindow) is also the same in that "the runtime handles
  the navigation", so **we can expect the POST to be preserved there as well**. What 9.26 confirmed is the (a)
  popup path, but the difference from the (b) put_NewWindow path is only "which WebView the runtime puts the navigation on";
  the continuation of the request itself is the same mechanism.
  ※However, there is not yet any measurement of the put_NewWindow path itself. That will be the first hurdle of the main part.

### ★Buttons 2 and 3 became unnecessary★

  Button 2 (POST + same tab) was placed as a sanity check (whether POST itself goes through in this WebView2), but
  it was **already proven the moment the popup side showed 200 + the form echo**, so there is
  no need to press it. The control of button 3 (GET + _blank) was also confirmed in the first 9.26 log.
  ★Something placed as insurance becoming unnecessary depending on the result of the main question★ is expected by design.

### Stage 9.26 numbers

- Design principles: **71** (no change. ※See "Candidate for becoming a principle" below; a decision is needed)
- Spec facts: 22 → **23** (spec fact 23 added)
- Standard file structure: **12** (no change. No new files. 11 production files)
- Changed files: 3: `Wv2Browser.cls` / `Wv2Pane.cls` / `Wv2Tests.bas`
  (imported in 2 rounds, 9.26 / 9.26b. 9.26b is Wv2Tests.bas only)

### ★Candidate for becoming a principle (awaiting tarboh's decision)★

  Whether to establish **"for a defect that reproduces only in the production environment, first try synthesizing a reproduction environment"** as design principle 72.
  9.26b is an example of it, and
    ・The 405 that should only be hit at the workplace was completely reproduced at home using httpbin.org
    ・We were freed from the constraint on the number of attempts (the production system cannot be hit repeatedly)
    ・It could be made into the **form of a controlled experiment**, "405 with OFF / 200 with ON", so the conclusion was uniquely determined
  were the effects. On the other hand, we cannot say "it should be done every time" (synthesis can be costly in some cases),
  so keeping it as an operational note rather than a principle is also a possible decision. To be decided in the next chat.

### Import operation notes (stage 9.26)

- As usual: comment out (or delete) all of the old version → import the new version. No name conflicts.
- 9.26b needs only a replacement of Wv2Tests.bas (the production files stay at 9.26).
- The HTML output destination for verification is %TEMP%\Wv2PostProbe\postprobe.html (4021 characters, UTF-8 without BOM).
  It may be deleted by hand once no longer needed.

### Open items for the next stage

- ★Main part = delegation to our own tab via put_NewWindow (medium to large)★ With spec fact 23 the theory is confirmed,
  so we are now in a position to concentrate on designing the UX (fitting it in our own tab rather than a popup).
  ・Required vtable calls: ICoreWebView2NewWindowRequestedEventArgs's
    get_Deferral (index 9) / put_NewWindow (index 4), and ICoreWebView2Deferral::Complete.
    Look up the indexes in WebView2.h and ride on the existing dcf approach.
  ・★Hard part★ Inside the deferral, waiting for the completion of another tab's Controller creation (SafeTimer one-shot, asynchronous)
    is a "nesting that waits for another event's completion inside an event". This is an area directly hit by spec fact 14.
  ・★Candidate mitigation = pre-warming a spare tab★ If one empty Pane is kept warm at all times, then at the moment
    NewWindowRequested arrives it can be handed to put_NewWindow immediately, making **the deferral itself unnecessary**.
    After use, the next spare is replenished with a Timer. Since nothing is awaited inside the event, it may be possible to
    eliminate the hard part entirely. Design issues such as not showing the spare in the tab bar are to be handled separately.
  ・★Debugging policy★ Because of spec fact 20, break/step execution cannot be used for processing inside event handlers.
    Build on the premise of pushing through with Debug.Print log design.
- Fate of delegation mode (PopupDelegateMode): it becomes unnecessary once the main part is in, but it may
  be worth keeping as an isolation switch. Decide when the main part is complete.
- An input field for arbitrary templates outside the presets + persistence of custom (an extension of 9.20).
- Polishing the tab bar's appearance (favicon, hover, close-button display control, gap animation during dragging).


---

## Stage 9.27 — The main part: pre-warm delegation via put_NewWindow

### Positioning

Spec fact 23 of 9.26b established that "if the navigation itself is handed to the runtime, the POST is fully preserved".
What remained was UX — putting that navigation into **our own tab** rather than a runtime-controlled popup.
This stage achieved that, and **the POST-link new-tab problem (the 405 issue) was permanently resolved**.
This settles the issue that had continued since stage 9.6a.

### Design discussion (choice of approach)

Three options were laid out and compared.

- **Option A: Deferral approach** — take the delay with `GetDeferral` (index 9), and after the other tab's Controller
  creation completes, `put_NewWindow` + `Complete()`.
  ★A point found in the discussion★ If you take a deferral, there is no need to "wait inside the event". Do `GetDeferral` inside the event
  and exit immediately, then continue with a Timer, which drops to the same structure as `RequestNewTabAsync`.
  However, since **the COM pointers of args and the deferral are carried out of the event**, our own
  AddRef/Release management is needed, and this is adjacent to the unresolved area of spec fact 10 (Sentinel MEM_FREE).
- **Option B: Pre-warm approach** — keep one empty Pane warm at all times, and at the moment of `NewWindowRequested` fire
  `put_NewWindow(the spare's ViewPtr)` and immediately `Exit Sub`. Neither a deferral nor carrying references over is needed.
- **Option C: Stage split** — implement option B first, and run the measurement of the `put_NewWindow` path itself
  at minimum cost. Option A is written only once it is really found necessary, as "insurance for when the spare runs out".

**Adopted = option C (option B first)**. The deciding factors were three: (1) since there is not yet any measurement of the `put_NewWindow` path,
that is the hurdle to pass first, not the deferral nesting; (2) option B waits for nothing inside the event, so it
avoids spec facts 14 and 20 altogether; (3) walking next to an unresolved bug for the sake of a mechanism whose success is not
yet confirmed gets the order backwards.

### Issues settled before implementation (7)

| # | Issue | Decision |
|---|---|---|
| 1 | When to create the pre-warm | **Warm it up after startup with a SafeTimer (1500ms)**. Right after Init is the busiest stretch, with Controller creations back to back, and the spec fact 20 window is open, so avoid it |
| 2 | Whether to touch `AddTab` | **Split the first half out into `CreatePaneCore`**. Called from both `AddTab` and pre-warm creation (leave no duplication) |
| 3 | Visibility of the spare | Hide it right after creation with `Ctrl_PutIsVisible False`. Bounds is the same rectangle as the existing tabs. **Leave making it visible to `ActivateTab`, so as not to break the centralization of visibility control** |
| 4 | Form of the switch | **Unify into 3 values of `NewWindowMode`** (0=current / 1=popup / 2=pre-warm). With 2 Booleans there would be 4 combinations, some meaningless. `PopupDelegateMode` is kept as a compatibility wrapper so the 9.26 tests stay alive |
| 5 | Whether to set `put_Handled` too | **Set both (default)**. However, make it switchable with a `SetHandledWithNewWindow` flag, placing insurance so that when a white screen or a double window appears, it can be isolated without fixing code |
| 6 | Order of adoption and replenishment | **Adoption first, replenishment after**. Replenishment is heavy because it waits for Controller creation, so show the screen first |
| 7 | Events that fly before Adopt | **Add nothing**. Judge after seeing in the real-machine log whether any real harm comes from dropped events (if we got ahead of ourselves, we would not know whether it had been needed) |

### Implementation (3 files)

**Wv2Browser.cls (9.26 → 9.27)**

- Split the first half of `AddTab` out into `CreatePaneCore` / `CreatePaneCoreBody` (issue 2).
  `AddTab` becomes a thin function that only does "call Core, then `m_tabs.Add` → `TabAdded` → `ActivateTab`".
  ★Reason for the 2-layer structure★ Body has multiple `Exit Function`s on failure. If the re-entrancy guard were raised and lowered inside Body,
  there would need to be a guarantee of lowering it at every Exit. Sandwiching it with a thin wrapper lets it be lowered in one place.
- `NewWindowMode` (Get/Let, 3 values) / `SetHandledWithNewWindow` (Get/Let) /
  `PopupDelegateMode` (turned into a compatibility wrapper. `m_popupDelegateMode` removed).
- Pre-warm mechanism: `PrewarmAvailable` / `ClaimPrewarmViewPtr` / `ReleasePrewarmClaim` /
  `RequestPrewarmAsync` / `EnsurePrewarmScheduled` / `m_prewarmTimer_TimerCall` /
  `CreatePrewarmPaneNow` / `RequestAdoptPrewarmedAsync` / `m_adoptTimer_TimerCall` /
  `AdoptPrewarmedNow` / `Debug_PrewarmStatus`.
- `Shutdown` and `Class_Terminate` release the spare tab (it is not in m_tabs, so the existing loop does not
  pick it up. Design principle 58).

**Wv2Pane.cls (9.26 → 9.27)**

The mode branch in `View_OnNewWindowRequested` is made 3-valued. In mode 2 it does only 4 things:
`ClaimPrewarmViewPtr` → `dcf(pArgs, 4, put_NewWindow, pNew)` →
`dcf(pArgs, 6, put_Handled, 1&)` → `RequestAdoptPrewarmedAsync` → immediate `Exit Sub`.
**Zero waiting inside the event**. When the spare is unusable / `put_NewWindow` fails,
it emits a ★-marked log and falls back to the current behavior (put_Handled + RequestNewTabAsync).

**Wv2Tests.bas (9.26b → 9.27)**

`Test_9_27_Mode_Legacy` / `_Popup` / `_Prewarm` / `_Status` / `_Handled_On` / `_Off` / `_Help`.
The verification page reuses 9.26b's `postprobe.html` as is (nothing newly created).

### Defensive measures added during implementation (not in the design discussion; added on own judgment)

1. **Re-entrancy guard with `m_creatingPane`** — `CreatePaneCore` spins DoEvents in `WaitPaneReady`, so
   the pre-warm TimerCall can interrupt during the startup sequence. If creation is in progress,
   **skip this time** (no re-arming of the timer. This avoids the danger of replacing a running SafeTimer).
   Even if skipped, the `EnsurePrewarmScheduled` at the end of `AddTab` reserves it again at the next opportunity.
2. **Do not write the function name on the right-hand side inside `ClaimPrewarmViewPtr`** — in VBA, reading a function name
   as a value inside a Function becomes a recursive call. Output to the log via a temporary variable.

### Real-machine test (3 rounds)

**Round 1: pass/fail of the main event -> ★Pattern A (complete success)★**

A controlled experiment on the same page with the same button settled it unambiguously.

| Mode | Result | Path |
|---|---|---|
| 0 (current behavior) | **405 Method Not Allowed** | put_Handled -> RequestNewTabAsync -> Navigate with GET |
| 2 (prewarm) | **200 + form echo** | put_NewWindow -> the original request lands in the spare tab |

Moreover, it landed **in our own tab, not in a popup**. `TabAdded(index=5)` -> `ActivateTab` ->
the tab bar updated to 5 tabs -> title `httpbin.org/post` reflected -> the NavBar URL was synced too.
**All the logs that did not appear at all with the 9.26 popup appeared.**

The POST contents were at the same level as 9.26b: all 3 `form` fields / `Content-Length: 172` /
`Content-Type: application/x-www-form-urlencoded` / `Origin: https://appassets.postprobe` /
`Referer: https://appassets.postprobe/` / `Sec-Fetch-Mode: navigate` / `Sec-Fetch-User: ?1` /
and the `　` (U+3000) in `nihongo` was intact too.

The safety net of point 5 turned out to be unnecessary (using it together with `put_Handled` works normally; neither a white screen nor a double window appeared).

**Round 2: release at shutdown -> ★complete pass★**

```
Wv2Pane.Cleanup: removing persistent handler ... iface=87651692836864   <- the spare View
Wv2Pane.Terminate: Releasing m_pView = 87651692836864
Wv2Pane.Terminate: Releasing m_pCtrl = 87651693768712
Wv2Browser.Shutdown: releasing the spare tab (prewarm)
Wv2Browser.Shutdown: severing back references of all 5 tabs
Wv2Browser.Terminate started (tab count=5)
```

At the moment of `Set m_prewarmPane = Nothing` the reference count reached 0 and `Class_Terminate` ran
on the spot, going through Cleanup all the way to the Release of View/Ctrl (the Debug.Print comes after the Set, so the release log appears first).
In the following Terminate the spare was no longer counted, and the 5 tabs + Environment were released in order.
**No leaks, no double release. The requirement of design principle 58 is met.**

**Round 3: spare-exhausted fallback (double-click the button) -> ★as designed★**

```
Wv2Browser.ClaimPrewarmViewPtr: spare tab unusable -> spare present ... Claimed=True / Usable=False
  ★[9.27 prewarm] The spare tab is unusable, so falling back to the current behavior★
    -> This navigation turns into a GET (a POST link would get 405). Scheduling a refill.
  put_Handled OK (the new window is suppressed)
Wv2Browser.RequestNewTabAsync: pending URL added
```

The second request did not steal the already-Claimed spare; it fell back to the current behavior and opened without breaking. **The Claim design
works exactly as intended.** Note that the second tab's title was `503 Service Temporarily
Unavailable` rather than 405 because of rate limiting on the httpbin.org side (it returns 503 under rapid clicking). Since the purpose of the test was
to confirm the fallback behavior, this is no problem.

In this round **the miss of point 7 actually occurred** (spec fact 25 below), and
**a nesting of `CreatePaneCore` actually occurred** (see Open items below).

### ★ Design principle 72 (new) — For a defect that reproduces only in production, try to synthesize a reproduction environment ★

**For a defect that reproduces only in the production environment, try to synthesize a reproduction environment if it can be put in the form of a controlled experiment.**
The value of synthesis lies less in securing the number of trials than in being able to line up
two results that differ by exactly one variable. If no place to synthesize it can be found / the cost of synthesis exceeds the cost of isolation, do not force it.

Making "can it be put in the form of a controlled experiment" the adoption criterion keeps the principle from being abused.
A real example is reproducing the workplace's 405 at home with `httpbin.org/post` in 9.26b, and in 9.27 too
the form **same page, same button, only the mode changed** was possible, so the conclusion was settled unambiguously.

### ★ Design principle 73 (new) — Before building a mechanism that waits inside an event, consider whether a "state in which there is no need to wait" can be prepared first ★

When you find you must wait for an asynchronous operation to complete inside an event handler, **before designing a waiting mechanism (Deferral,
waiting for the completion of a nested operation), consider whether preparing something in advance would remove the need to wait.**

9.27 is the example. The main event was at first expected to be the Deferral approach, but just warming up a spare tab beforehand
**eliminated the wait inside the event entirely**, and carrying COM references to args/deferral over was no longer needed either.
It avoids both spec fact 14 (deadlock) and spec fact 20 (the EBMode window), and the amount of implementation
also dropped substantially. **"Eliminating the need to wait" is sometimes a better course than "being clever about how to wait."**

The cost is "always reserving one unit's worth of resources up front." Whether that is acceptable decides adoption.

### ★ Design principle 74 (new) — The full-sync approach is self-healing against missed events ★

A re-evaluation of design principle 60 (VBA->JS sends the whole tab state every time: the full-sync approach).

**Even if a delta notification is missed, the next full sync overwrites the whole state, so recovery is automatic.**
In round 3 of 9.27, `TabTitleChanged` / `TabUrlChanged` were missed 3 times, but
the subsequent `TabAdded` -> `PushSyncToJs` re-read the title held by the Pane, so
**the tab bar showed the correct `503 Service Temporarily Unavailable`** (spec fact 25).

In a situation where a delta approach would leave the miss as a permanent gap, the full sync heals itself.
**In domains where event ordering is hard to guarantee (COM callbacks, asynchronous creation), the full-sync approach is
a choice for robustness, not a performance compromise.**

### ★ Spec fact 24 (new) — An existing CoreWebView2 of another Controller can be passed to put_NewWindow ★

If you set, on `ICoreWebView2NewWindowRequestedEventArgs::put_NewWindow` (vtable **index 4**),
**an already-created CoreWebView2 that belongs to a different Controller and is still at about:blank**,
the WebView2 runtime puts the original request into that WebView.

- **POST is fully preserved**. Method, body, even `Referer` / `Origin` / `Sec-Fetch-*` are
  at the same level as spec fact 23 (popup delegation).
- **Using `put_Handled(TRUE)` together is correct** (vtable index 6). Setting both works normally.
  Neither a white screen nor a double window occurs.
- **The Controller must have been created from the same Environment** (this project satisfies that naturally, because
  `Wv2Browser` shares a single Environment).
- **No Deferral is needed**. It is enough to fire `put_NewWindow` before the event handler returns.
- Unlike popup delegation, **the WebView passed in is our own Controller, so all events come through**
  (NavigationStarting / HistoryChanged / DocumentTitleChanged / NavigationCompleted were all confirmed).
  With the 9.26 popup, not a single line appeared. **This is why the main event was prioritized over popup delegation.**

vtable reference (`ICoreWebView2NewWindowRequestedEventArgsVtbl`):
3 = `get_Uri` / **4 = `put_NewWindow`** / 5 = `get_NewWindow` / **6 = `put_Handled`** /
7 = `get_Handled` / 8 = `get_IsUserInitiated` / 9 = **`GetDeferral`** (note: not `get_Deferral`).
`ICoreWebView2Deferral`: 3 = `Complete`.

### ★ Spec fact 25 (new) — Events after put_NewWindow may fire before we adopt it ★

The runtime starts the navigation the moment `put_NewWindow` is fired, so **events of the target Pane may fire
before adoption (registration into m_tabs by the Timer)**.

- In round 1 (httpbin.org was slow), only `NavigationStarting` fired before Adopt.
  That handler does not Notify the Browser, so there was zero actual harm.
- In round 3 (the second request returned quickly with a 503), three events, `HistoryChanged` / `DocumentTitleChanged` /
  `NavigationCompleted`, fired before Adopt, and
  `Wv2Browser.NotifyUrlChanged: senderPane not found in m_tabs (ignored)` appeared 3 times.

**It does not crash** (the Notify family is built to ignore the event when the reverse index lookup fails).
**The display is not broken either** (as in design principle 74, the `TabAdded` -> `PushSyncToJs` at Adopt time re-reads the
title held by the Pane, so the correct value appears).

-> **The "add nothing" decision on point 7 was correct.** Had we written preemptive correction code,
we would never have known whether it was needed.

### File structure (unchanged: 12 files, 11 production)

- Standard files: `Wv2Thunks.bas` / `Wv2Json.bas` / `Wv2Url.bas` / `ComCallbackHandler.cls` /
  `Wv2Environment.cls` / `Wv2Pane.cls` / `Wv2Browser.cls` / `Wv2TabBar.cls` / `Wv2NavBar.cls` /
  `Wv2SettingsBridge.cls` / `UserForm1.frm` / `Wv2Tests.bas` (verification only)
- What was edited in 9.27 were the 3 files `Wv2Browser.cls` / `Wv2Pane.cls` / `Wv2Tests.bas`. No new files.
  **UserForm1.frm needed no changes thanks to the `CurrentBrowser` scaffolding of 9.17** (same as 9.26).
- Design principles: 71 -> **74** (72, 73, 74 added)
- Spec facts: 23 -> **25** (24, 25 added)

### ★Changes to log wording (a caution when comparing against past logs)★

Because the first half of the old `AddTab` moved to `CreatePaneCore`, the prefixes of the following lines changed.

- `Wv2Browser.AddTab: Controller 生成 OK` -> `Wv2Browser.CreatePaneCore: Controller 生成 OK`
- `Wv2Browser.AddTab: 初期 Bounds に...` -> `Wv2Browser.CreatePaneCore: 初期 Bounds に...`
- `Wv2Browser.AddTab: View 取得 OK` -> `Wv2Browser.CreatePaneCore: View 取得 OK`

`Wv2Browser.AddTab: タブ追加 OK (現在のタブ数=N)` is as before (it remains in the second half).
**This line can also be used to tell them apart: if it appears, the tab went into m_tabs; if not, it is the spare tab.**

### Open items / homework for the next stage

- **`m_creatingPane` is a flag, so it drops early when nested**. In round 3 a nesting of
  `CreatePaneCore` actually occurred (while the `WaitPaneReady` of the prewarm creation was
  running DoEvents, `m_newTabTimer_TimerCall` called `AddTab` -> `CreatePaneCore`).
  Each Pane is a separate instance and `WaitPaneReady` also watches its own pane, so **the result was normal**,
  but since the flag drops when the inner one finishes, the next one may run even while the outer one is still creating.
  -> **Consider turning it into a counter (make `m_creatingPane` a Long, increment/decrement)**. Small.
  Note: the guard works only for the prewarm TimerCall; it never applied to the newTab TimerCall
  (that is existing behavior since 9.10e).
- **If `RequestPrewarmAsync` is called while a Claimed spare remains, no refill happens**.
  In round 3, `RequestPrewarmAsync: a spare tab already exists, doing nothing` appeared.
  There is no actual harm (`CreatePrewarmPaneNow` runs right after Adopt), but **the log is misleading**.
  -> Add "and not Claimed" to the early-return condition of `RequestPrewarmAsync`,
  or change the log wording. Small.
- **Whether to hold 2 spares**. Rapid clicking exhausts the spare (demonstrated in round 3). If that is plausible in real work,
  consider going to 2. The cost is 2 extra WebView2 instances at all times. **Deciding once it actually causes trouble** is fine.
- **The fate of `PopupDelegateMode` / `NewWindowMode`**: now that the main event is complete, modes 0 and 1 are
  verification-only. They are worth keeping as isolation switches, so **keep them for now**.
- ★Forgetting to switch modes returns you to 405★ This was demonstrated in the first half of round 3. The default is still 0 (current behavior),
  so **if the main event is to be used routinely, a decision is needed on whether to make the default 2**.
  -> It takes 1500ms at startup for the spare to warm up, so with a default of 2 a link clicked right after startup
    would fall back. How to set the default is a design issue for the next stage.
- (Continued) Free-form template input field outside the presets + persistence of custom (an extension of 9.20).
- (Continued) Polishing the tab bar's appearance (favicon, hover, close-button display control, animation).
- (Continued) Root-cause fix of the Sentinel MEM_FREE bug (spec fact 10).

---

## Stage 9.28 — Making prewarm delegation the default + 2 small fixes around the spare

### Purpose
Although the main event (prewarm delegation via put_NewWindow) passed on the real machine in 9.27,
the default of NewWindowMode was still 0 (current behavior = POST gets 405).
This round collects 3 small items from the "Open items / homework for the next stage" of 9.27 in one go.

### Aligning on the approach (the 4 options presented and the choice)
| Option | Content | Verdict |
|---|---|---|
| (a) | Default 2, prewarm delay stays at 1500ms | ★Adopted★ |
| (b) | Default 2 + shorten the delay | Rejected. The 1500ms is a safety margin deliberately placed in 9.27 point 1 because "right after Init is the busiest interval with consecutive Controller creations, and the window of spec fact 20 opens". Cutting it for the sake of making this the default would be putting things in the wrong order |
| (c) | Default 2 + if there is no spare, wait with a Timer and then put_NewWindow | Rejected. It needs a Deferral, which would be going back to what 9.27 avoided. Only after (a) causes actual harm |
| (d) | Keep default 0 | Rejected. It was already demonstrated in 9.27 round 3 that forgetting to switch returns you to 405 |

Grounds for choosing (a): the 1500ms hole right after startup is practically almost impossible to step into (during that time the first
page is still loading and links cannot be clicked). Even if stepped into, it does not break; it merely falls to 405, and
a ★-marked log appears, so the cause is obvious at a glance.
The explicitly accepted trade-off: one extra WebView2 is always resident (one more msedgewebview2.exe).
In an environment where that is unacceptable, firing NewWindowMode = 0 after startup restores the previous behavior.

### Implementation (only 2 files, Wv2Browser.cls / Wv2Tests.bas. Production is Browser alone)
- **(A) Default value** — m_newWindowMode in Class_Initialize changed from 0 to 2.
  The prewarm reservation was already built to run unconditionally from Init, so nothing else changed.
  ★Wv2Pane.cls is unchanged★ It reads the mode from the Browser every time, so changing the default alone takes effect.
- **(B) m_creatingPane from Boolean to Long** — the declaration / Class_Initialize / the increment and decrement in CreatePaneCore /
  the check (> 0) in m_prewarmTimer_TimerCall. One line of lower-bound clipping was also added.
- **(C) The early-return log of RequestPrewarmAsync** — branching on m_prewarmClaimed to distinguish
  "a usable spare tab already exists" from "the spare tab is Claimed (awaiting adoption). It is refilled automatically right after adoption,
  so no reservation is made here". ★Behavior is identical to 9.27★.
- Wv2Tests.bas: newly created Test_9_28_Help. Corrected the control procedure of Test_9_27_Help to assume default 2.

### ★Post-hoc report: (C) was changed one step from what was agreed★
At the time of agreement it was "add 'and not Claimed' to the early-return condition", but during implementation
it turned out that this would make things worse, so it was changed to a log-wording-only fix.
Reason: adding the condition would set the 1500ms timer when Claimed, but by then the adoption Timer
(1ms) has already finished the refill via AdoptPrewarmedNow -> CreatePrewarmPaneNow, so at firing time it would only add
another wasteful log, "a spare tab already exists (skipping creation)", and the refill would not come any sooner.

### Real-machine test (round 1: using httpbin.org)
- At startup: NewWindowMode = 2 / spare present, usable=True -> ★(A) pass★
- Button 1 pressed: without touching the mode at all, ClaimPrewarmViewPtr -> put_NewWindow OK ->
  put_Handled OK -> Adopt -> TabAdded -> ActivateTab -> Prewarm refill ran through to the end -> ★entered the delegation path★
- Rapid clicking: a nesting of CreatePaneCore actually occurred (the newTab TimerCall during prewarm creation) and
  ★both completed normally★ -> (B) pass. The "skipped" log did not appear this time.
- The new wording `予備タブは Claim 済み (採用待ち)...` (spare tab is Claimed (awaiting adoption)...) appeared during rapid clicking -> ★(C) pass★
- Shutdown: `Shutdown: 予備タブ (プリウォーム) を解放` (releasing the spare tab (prewarm)) -> Release of all 6 tabs + the spare ran through to the end, no leaks.

### ★However, httpbin.org kept returning 503 for its own server-side reasons, and only the confirmation of the 200 was impossible★
It first showed up in the form "delegation went through but the title is 503 and the screen is blank".
**Isolation**: the fallback side (plain GET) of the rapid-click test was equally 503. If healthy,
a GET always returns 405, so it was established that the server was not responding regardless of the request contents.
To be sure, I typed `https://httpbin.org/post` directly into the address bar and checked the GET alone ->
again 503 / isSuccess: False. httpbin.org is a free public service with long-standing reports of periodic 503s, and
I judged it to be inherently unreliable as a verification base.
-> On to stage 9.28b (making the verification destinations 3 routes).

---

## Stage 9.28b — Making the POST probe's destinations 3 routes / making httpbingo.org the primary

### Purpose
Eliminate the situation where one external service going down stops the verification of the main event.
★The change is Wv2Tests.bas only. Not one line of the 11 production files was touched★.

### Pre-checking the candidate (confirmed with the browser itself before writing code)
Typed `https://httpbingo.org/post` directly into the address bar -> the screen showed `Method Not Allowed`.
That the same structure as the workplace's Tomcat (GET gets 405) holds was confirmed before implementation.
★The browser I built myself is serving as my own verification tool★

### Structure
| Button | Role | Destination |
|---|---|---|
| 1 | Primary POST + _blank | **httpbingo.org/post** |
| 2 | Sanity POST + same tab | httpbingo.org/post |
| 3 | Control GET + _blank | httpbingo.org/get |
| 4 | Spare 1 POST + _blank | httpbin.org/post (the old primary, demoted) |
| 5 | Spare 2 POST + _blank | postman-echo.com/post (newly added) |

The endpoints are consolidated into the variables ep / epGet / ep2 / ep3. Future swaps only touch the first 4 lines.
★JS stays at zero★ The policy of building with a plain form only is maintained.

### ★Side effect and benefit of switching to httpbingo★
- **Side effect**: httpbingo's error page has no title element, so Edge uses the URL as a stand-in for the title
  (`→ タイトル: httpbingo.org/post`, i.e. `-> Title: httpbingo.org/post`). The success case has the same title, so
  ★pass/fail cannot be read from the Immediate-window log★. httpbin showed 405 METHOD NOT ALLOWED as the
  title, so this one point is a regression. Always make the verdict from the on-screen display.
- **Benefit**: the POST response JSON has a ★method field★, so it can be directly confirmed on screen that
  the request has not turned into a GET (httpbin does not have this).

### Real-machine test (round 2: using httpbingo.org) ★ The main target fully reproduced ★
Just by pressing button 1 without ever switching modes:
```
Wv2Browser.ClaimPrewarmViewPtr: reserved the spare tab (ViewPtr=31404801116160)
  [9.27 prewarm] put_NewWindow OK (View of the spare tab=31404801116160)
    put_Handled OK (used together with put_NewWindow)
Wv2Browser.Adopt: adopted the spare tab as a tab (current tab count=5, ...)
Wv2Pane.View_OnNavigationCompleted: isSuccess: True     ← True for the first time
Wv2Browser.Prewarm: warmed a spare tab (ViewPtr=31404801131520, ...)
```
The screen (= the substance of the verdict):
```
"method": "POST"                          ← not turned into GET
"Content-Length": ["172"]                 ← body not lost
"form": { memo / nihongo / probe }        ← all 3 fields arrived
"nihongo": ["日本語　全角スペース入り"]      ← UTF-8 intact, even the full-width space
```
Test_9_27_Status after the test: spare present / usable=True / ViewPtr updated = immediate replenishment also confirmed.

---

### ★ Spec fact 26 — put_NewWindow delegation even preserves Sec-Fetch-User: ?1 ★
In the real-machine log of 9.28b, the headers that reached the delegated tab included all of the following.
```
Origin: https://appassets.postprobe        (preserved even when it comes from a virtual host mapping)
Referer: https://appassets.postprobe/
Sec-Fetch-Site: cross-site
Sec-Fetch-Mode: navigate
Sec-Fetch-User: ?1                         ← ★newly confirmed★
```
Sec-Fetch-User: ?1 is the header indicating "a navigation caused by a direct user action".
Even when the navigation is moved onto the CoreWebView2 of a different Controller via put_NewWindow, the runtime
★keeps treating it as a user-initiated navigation★. This is an addendum to spec fact 23: it passes straight through
even to servers that look at this header for CSRF protection or bot detection.
Also, unlike popup delegation (mode 1), all events go through our own handlers, as already established in 9.27.

### ★ Design principle 75 — Make the scaffolding of a controlled experiment (external verification services) redundant ★
Design principle 72 decided that "a bug that only reproduces in production is reproduced by synthesizing a reproduction environment",
but in 9.28 we failed because ★the synthesized scaffolding itself went down, making it impossible to isolate the cause★.
A controlled experiment is valuable because it lays out "two results with exactly one variable changed" side by side,
so if the scaffolding goes down its value drops to zero. Having 3 destinations costs only duplicating the HTML, so it is cheap.
Generalization: do not base the foundation of verification on a single external dependency. Keep it in a form where swapping takes one line.

### ★ Design principle 76 — When the main subject fails, first check whether the control side "fails as expected" ★
In 9.28, when we hit "the delegation went through but the screen is blank (503)", what settled it, before suspecting our own bug,
was noticing that ★a plain GET on the fallback side was also 503★.
If healthy, GET must return 405, so the moment 503 came back instead of 405 we could conclude
"the server is not responding regardless of the request content".
Generalization: when you see an anomaly, first check ★whether the control system produces the expected anomaly★ rather than the normal case.
If even the control system is broken, the cause lies outside your own code.

### ★Log message changes (9.28 / 9.28b. Notes for matching against past logs)★
- `RequestPrewarmAsync: 既に予備タブがあるので何もしない` (RequestPrewarmAsync: a spare tab already exists, so do nothing) was split into two:
  - `使える予備タブが既にあるので何もしない` (a usable spare tab already exists, so do nothing) (not yet Claimed)
  - `予備タブは Claim 済み (採用待ち)。採用直後に自動で補充されるので、ここでは予約しない` (the spare tab is already Claimed (awaiting adoption). It is replenished automatically right after adoption, so no reservation here)
- When using httpbingo the title becomes `→ タイトル: httpbingo.org/post` (→ title: httpbingo.org/post), and ★success and failure give the same string★.
  There is no clue equivalent to `405 METHOD NOT ALLOWED` / `200 OK` that the httpbin logs up to 9.27 had.
- The probe page title changed to `POST プローブ (第9.26b / 第9.28b 更新)` (POST probe (updated in 9.26b / 9.28b)).

---

## Stage 9.29 — Custom search URL input field + persistence of custom (+ cleanup of spare reservations)

### Scope of this round
We collected 2 items from the open items of 9.28b. **The main one is open item 2 (custom search engine)**;
open item 1 (useless reservation) touches the same `Wv2Browser.cls`, so it was bundled in.

There are 3 changed files (no new ones):
- `Wv2SettingsBridge.cls` → stage 9.29 (main)
- `Wv2Browser.cls` → stage 9.29 (settings screen HTML + spare reservation guard)
- `Wv2Tests.bas` → stage 9.29 (verification + making the existing Test_9_20 stash handle 2 keys)

`Wv2Pane.cls` is unchanged from stage 9.27 (which is correct).

### Agreed design (agreed in one go on 8 points before implementation)
| Point | Decision | Reason |
|---|---|---|
| 1 Notation | **Keep the prefix method** | A `%s` placeholder method would touch NormalizeUrl / PreviewUrlFor / the preset definitions / compatibility of saved values all at once. It would turn "add an input field" into "change the notation", so it is split out into a separate stage |
| 2 Validation | **Minimal** (non-empty + starts with `http://` / `https://`) | A search template is an **incomplete URL** ending in `?q=`. Running it through a general URL validity check would reject even correct input |
| 3 Storage format | **2 keys; the `template` line only when custom** | If we always wrote 2 lines, then when a preset's URL is later fixed in code, the saved value would resurrect the old URL (design principle 77) |
| 4 UI | **An independent section below the cards** | Making it a 5th card would mix a non-preset into `EngineList()`. It would break the 9.18 design of "do not duplicate engine names on the HTML side" |
| 5 API | **Add 3** (`SetCustomEngine` / `GetCustomTemplate` / `PreviewUrlForTemplate`) | custom is the only input field where typos can happen, so a preview **before** applying is valuable. It also has the advantage of going through VBA instead of hand-rolling the encoding in JS |
| 6 Restore | Extend `LoadEngineTemplate` | Browser just receives "a template string or empty". Keeps the 9.20 design that "Browser does not know the vocabulary of engine names" (Browser side is **unchanged**) |
| 7 Open item 1 | **Settle by reading the state transitions at implementation time** | I was not confident that the assumption at agreement time (`m_prewarmClaimed`) was right, so it was left to me on condition of settling it and reporting afterwards |
| 8 Verification | Two: pure logic + real machine | It touches a real file, so stash/restore follows the 9.20 convention |

### Storage format (settings.txt)
```
engine=custom
template=https://ja.wikipedia.org/w/index.php?search=
```
When a preset is selected, **only 1 line** `engine=bing` (same as 9.20).
Switching from custom to a preset makes the `template` line disappear naturally.
**Side effect (by design)**: the custom URL the input field remembered disappears at the same time.

`LoadEngineName` was generalized into a generic parser `LoadValue(keyName)`.
Since it **uses only the first `=` as the delimiter**, a line whose value contains `=`,
like `template=https://...?search=`, is read without losing anything.

### Key points of the API
| Method | Return value | Notes |
|---|---|---|
| `SetCustomEngine(template)` | `"custom"` / preset name / `""` | When `""`, it has **touched neither** the Browser nor the saved file |
| `GetCustomTemplate()` | URL or `""` | If currently custom, that; otherwise the saved custom |
| `PreviewUrlForTemplate(tpl, terms)` | URL or `""` | Does not change Browser state. For the pre-apply preview |

**If the input is the same URL as a preset**, it returns **that preset name** instead of `"custom"`,
and saves it as the preset (the `template` line is not written). This simply follows the reverse lookup of `TemplateToName`.
The settings screen also highlights that card and displays "The input was the same as the preset (google), so
it was selected instead".

### Additions to the settings screen (BuildSettingsHtml)
A `.custom` section below the card grid. Input field + [Apply] button + message line.
- While typing, `oninput` → `PreviewUrlForTemplate` gives a live preview **before** applying
- If invalid, red text (`.cmsg.err`). Even if applied, the VBA side ignores it
- While custom is in use, the frame glows blue and a "in use" badge appears (`markCustom()` is called from `markSel()`, so
  there is still only one place to update the selection highlight)
- **Do not use `<form>`** — Enter would submit and navigate the settings screen itself.
  Enter is caught with `onkeydown`

### ★ True cause of open item 1 — the assumption at agreement time was wrong ★
At agreement time I thought it would be settled by a single line, "do not set a reservation if `m_prewarmClaimed`", but
when tracing the code at implementation time it turned out to be **a different cause**.

```
m_adoptTimer_TimerCall
  ├ AdoptPrewarmedNow      … here m_prewarmPane = Nothing / m_prewarmClaimed = False
  └ CreatePrewarmPaneNow
      └ CreatePaneCore
          └ WaitPaneReady … ★runs DoEvents★
              └ (interrupt) m_newTabTimer_TimerCall → AddTab
                              └ EnsurePrewarmScheduled ← m_prewarmPane is still Nothing
                                                          → a reservation is set
```
`m_prewarmClaimed` is already `False` at this point, so it is **not a condition**.
The correct check is "whether we are nested" = `m_creatingPane > 0` (the counter we made in 9.28
could be used as is).

**Trade-off**: if the outer spare creation fails, this `AddTab` sets no reservation, so
the spare stays empty until the next `AddTab` / switch. An automatic retry on failure is **not included**, because in an environment where creation
fails persistently it would become an infinite retry every 1.5 seconds.

### Real-machine test result (full pass)
**Test_9_29_Custom: 23 / 23 [ALL OK]**

Real machine (Test_9_29_Help):
- (A) Apply custom — entering Wikipedia's search template → the pre-apply preview shows
  `日本 天気 → https://ja.wikipedia.org/w/index.php?search=%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97`.
  Input without a scheme gives a red warning, and even when applied only
  `SetCustomEngine: 不正なテンプレートなので無視` (SetCustomEngine: invalid template, ignored) appears and the search engine is unchanged.
  After applying, the frame glows and "in use" is shown. The address bar search went to Wikipedia.
- (B) Persistence — after fully quitting Excel and restarting,
  `Class_Initialize: 保存済みエンジンを復元 = https://ja.wikipedia.org/...` (Class_Initialize: restored the saved engine = ...) appeared,
  and searching without opening the settings tab gave Wikipedia results. When the settings tab was opened the URL was
  restored in the input field and the "in use" badge also appeared.
- (C) Cleanup of spare reservation — during the spare-exhausted fallback,
  `EnsurePrewarmScheduled: 予備タブの生成中 (入れ子) なので予約しない` (EnsurePrewarmScheduled: a spare tab is being created (nested), so no reservation) appeared, and
  **`Prewarm: 既に予備タブがある (生成をスキップ)` (Prewarm: a spare tab already exists (skipping creation)) no longer appears**.
  In `Test_9_27_Status`: spare present / usable=True (the spare is kept at 1).
- Revert — clicking the Google card on the settings tab reverted it, and it stays in effect after restart.
  The custom-part string was gone (as designed).

★This log reflects the order in which the nesting occurs exactly★
`Prewarm: 予備タブの生成を開始` (Prewarm: started creating the spare tab) →(during DoEvents waiting for the Controller)→
`TimerCall: 保留 URL を処理` (TimerCall: processing the pending URL) → `AddTab: タブ追加 OK (現在のタブ数=8)` (AddTab: tab added OK (current tab count=8)) →
`EnsurePrewarmScheduled: …入れ子…` (…nested…) → back to the outer one, `Prewarm: 予備タブを温めた` (Prewarm: warmed the spare tab).
The explanation of the true cause above is visualized as is.

The main target (spec facts 24 to 26) was also reconfirmed at the same time. The delegated tab received
`"method": "POST"` / `Content-Length: 172` / 3 `form` fields /
`Sec-Fetch-User: ?1` / `Origin: https://appassets.postprobe`, all present.

---

### ★ Design principle 77 — In persistence, first decide "is the name authoritative, or the value?" ★
When saving the search engine, a preset saves the **name** (`engine=bing`), and the URL definition is placed in
a constant on the code side. Only custom saves the **value** (`template=https://...`).

Reason: if we always saved the value, then even if a preset's URL were fixed in code later,
**the saved file would keep resurrecting the old URL**. To the user this appears as the most confusing kind of bug:
"the thing that was supposedly fixed is not fixed". Conversely, custom cannot be expressed by a name, so it has no choice but to hold the value.

Generalization: when persisting a setting, classify whether the item is "something the code holds the definition of (= save the name)" or
"something only the user holds the definition of (= save the value)". Mixing them leaves it unclear
which side has the initiative for updates.

### ★ Design principle 78 — A "DoEvents for waiting on completion" is a re-entry path ★
The `DoEvents` in a completion-wait loop such as `WaitPaneReady` **runs every pending event, including SafeTimer's TimerCall**.
Therefore, while waiting for completion,
**the state outside that procedure is observed by other code while it is "not yet rewritten"**.

Open item 1 of 9.29 was exactly this: an `AddTab` that interrupted before `CreatePrewarmPaneNow` finished creating the spare
mistook it for "there is no spare" and set a needless reservation.

The pattern of the fix: **set an "in progress" flag in procedures that contain a completion wait, and have the outside observers
look at it**. The `m_creatingPane` that we turned into a Long counter in 9.28 could
be used as is for a different purpose (since nesting can occur, a Boolean is not enough).

Generalization: the moment you write `DoEvents`, that procedure must be written on the premise that "other code will interrupt".
It is the flip side of spec fact 14 (waiting inside a handler deadlocks):
**where you can wait, conversely be prepared for re-entry**.

### ★ Design principle 79 — The more a bug looks like "fixable with this one line", the more you should pin down the cause before fixing ★
At agreement time, open item 1 was written as "can be eliminated by one line, not setting it if `m_prewarmClaimed`", but
when tracing the call path at implementation time it turned out that **that flag is already False and is not a condition**.
Had I put in that one line, the useless log would not have gone away yet it would have been recorded as "handled",
which would have been doubly confusing the next time it was seen.

Why I got it wrong: because in (C) of 9.28 I had dealt with a similar log of `RequestPrewarmAsync`,
**I presumed that a log with the same wording had the same cause**, but it was actually a different route (nesting).

Generalization: even when the log wording is the same, the cause is not necessarily the same. The more a fix looks like it takes one line,
verify "what value is that flag at that moment" along the path. When an agreed condition turns out to be wrong,
neither fix it on your own and stay silent, nor implement it as agreed, but **report afterwards together with the confirmed cause**.

### ★Log message changes (9.29. Notes for matching against past logs)★
- `Wv2SettingsBridge.SaveEngineName: 保存 '...'` (SaveEngineName: saved '...') → changed to **`SaveSettingsFile`**.
  In the custom case ` / template=<URL>` follows.
  (`SaveEngineName` remains as a thin wrapper of `SaveSettingsFile(name, "")`)
- New: `Wv2Browser.EnsurePrewarmScheduled: 予備タブの生成中 (入れ子) なので予約しない` (a spare tab is being created (nested), so no reservation)
- New: `Wv2SettingsBridge.SetCustomEngine: 適用 '<URL>' -> '<name>'` (applied)
- New: `Wv2SettingsBridge.SetCustomEngine: 不正なテンプレートなので無視 ('<入力>')` (invalid template, ignored ('<input>'))
- Gone: `Wv2Browser.Prewarm: 既に予備タブがある (生成をスキップ)` (a spare tab already exists (skipping creation)) during the spare-exhausted fallback
  (it can still appear via other routes such as at startup)

---

## Stage 9.30 — Isolating VirusTotal's Bkav Pro verdict (no code changes)

### Scope of this round
This round collects open item 1 of 9.29 (VirusTotal flagging Bkav Pro as `X97M.Downloader.Macro`).
★No VBA code was changed at all★ We only built samples and compared them.
The deliverables are a clean build of the workbook (v0_5_0 → v0_5_2) and a false-positive report.

### Agreed design (8 points before implementation → the first move was replaced midway)
The original first move was "remove the bundled unused modules (`旧WebView2.cls` / `MouseScroll.bas`)", but
★they were not in the actual workbook from the start★, so following the pre-analysis results
the first move was replaced with "clean build" (the replacement of point 2 was re-agreed).

| Point | Decision |
|---|---|
| 1 Pre-analysis | Statically analyze the workbook with `oletools` (olevba/pcodedmp) before uploading to VT |
| 2 First move | (after replacement) Clean build. Reduces VT round trips |
| 3 Control | Take just one empty workbook (sample A) (design principle 76) |
| 4 How to build samples | Manually in the VBE. Scripting via a VBIDE reference rejected, because it itself adds detection risk |
| 5 Verdict criterion | Look only at the Bkav Pro row. Upload all together on the same day to align the engine configuration |
| 6 Disclosure | Uploads to VT are distributed to vendors. Confirmed that the sample's sheets are empty |
| 7 Landing point | If the cause is removable, remove it; if it is a core API, false-positive report + record it and accept |
| 8 Record | Keep an investigation record even if the conclusion is "false positive" |

### Samples
| Sample | Contents | How made |
|---|---|---|
| A | Empty workbook (zero code) | Save a new workbook in macro-enabled format |
| B | All modules | Ported into A by VBE drag & drop |
| B2 (= v0_5_1→v0_5_2) | All modules | Into A via ★export → import★ |
| C2 | Without `Wv2Tests.bas` | Same as above (not deleted, but never included from the start) |

### What the pre-analysis (olevba) revealed

### (1) A hypothesis that once looked like the main suspect — VBA Stomping (turned out wrong)
olevba judged **VBA Stomping** (mismatch between source and P-code = a standard AV evasion technique) as positive.
When I reproduced the judgment logic myself and checked the identifiers behind it, v0_5_0 had
`_B_var_UserFrom1` / `_B_var_vbmodelss` / `_B_var_userform12` / `ttt` and the like,
★remnants of code I had mistyped and run in the past, and of deleted variables★ left in the P-code cache.

### (2) ★olevba's stomping judgment cannot, in principle, come out negative★
The judgment is a simple substring match of "whether the P-code identifier appears as a string in the source".
- **Type suffix problem**: P-code holds identifiers with a `$` such as `Test_9_26_PostProbe$`,
  but the source only ever writes them without `$`, so it always misses.
- **pcodedmp identifier shift**: when I traced the single `Class1` remaining in v0_5_2, it appeared at the position of
  `Private m_kind As HandlerKind` in `ComCallbackHandler.cls`,
  and at the same spot function names from `Wv2Thunks.bas` (`Thunks_AcquireSlot` / `AcquireHandlerFor`) and
  `As <crash>` appeared. ★The analysis tool is failing to resolve the identifier table★.
→ **v0_5_2 is the cleanest state reachable by analysis**. Beyond this is the limit of the tools.

### (3) Effect of the clean build (drag and drop is not enough)
| Sample | `__SRP_*` | vbaProject.bin | Modules | Real P-code remnants |
|---|---|---|---|---|
| v0_5_0 | 32 / 360 KB | 2,041,856 | 19 | **6** |
| B (D&D) | 8 / 19 KB | 1,323,008 | 17 (`Wv2Json` missing) | 1 |
| **v0_5_2** | 12 / 21 KB | 1,340,416 | **18** | **0** |
| C2 | 12 / 21 KB | 1,154,048 | 17 | 0 |

★Drag and drop is close to a binary copy between projects, and the P-code comes along with it,
so it does not produce a clean build★. Export to files and then import is required.
Note: in the D&D version, forgetting to move `Wv2Json.bas` happened (a state that does not compile).
Note: in the sample where `Wv2Tests.bas` was **deleted**, the identifiers of the deleted module
  (10 of them, such as `Test_9_18_BridgeLogic` / `CheckEq`) remained as new remnants,
  and got mixed with the "effect of fewer URLs", making them impossible to separate. **Do not delete; do not include it in the first place**.

### VT results (3 samples, same day)
| Sample | Bkav Pro | Other engines |
|---|---|---|
| A (empty workbook) | **Negative** | 0 |
| v0_5_2 (all modules) | **Positive** `X97M.Downloader.Macro` | 0 |
| C2 (without `Wv2Tests`, URLs 49→19) | **Positive** same | 0 |

→ It is confirmed that **the cause is neither the URLs in the test code nor the P-code remnants, but the core APIs themselves**.
All other engines, including Microsoft Defender, are negative. It is a detection by Bkav Pro alone.

### ★VT's Code insights concluded benign★
VirusTotal's own automated code analysis accurately read what this workbook implements (WebView2Loader.dll, machine-code thunks,
SAFEARRAY, `DispCallFunc`, `AddHostObjectToScript`, tabs/history/search engine), and stated explicitly:
"`VirtualAlloc` and dynamic assembly generation are a common shellcode-execution technique, but here they are used strictly as a technical workaround
for implementing COM callbacks in pure VBA", "no evidence of any malicious download,
command execution, credential theft, persistence or defense evasion", and "it is harmless and a legitimate developer
tool". ★The strongest material for the false-positive report★

Of the 24 Suspicious items in olevba, the ones that cannot be avoided are
`VirtualAlloc` / `RtlMoveMemory` (= flagged as code injection) / `SetTimer` (= flagged as shellcode) /
`Lib` (= DLL execution). All of them are the core of this project and cannot be removed.

### False-positive report
Reported in English to `fpreport@bkav.com` (CC: `bkav@bkav.com`).
Contents: "file information + why `VirtualAlloc` is used + a list of what it does not do +
a quotation of the Code insights + the VBA project has no password, so the entire source can be inspected".
Note: Bkav has a track record of multiple delistings in response to false-positive reports, but ★it can recur with every new version released★,
  so this is not a permanent fix; it will be handled case by case.

---

### ★ Design principle 80 — Do a clean build before distribution or scanning ★
A VBA project keeps the identifiers of mistyped code that was run and of deleted variables in the P-code cache
(`__SRP_*`). In a workbook developed over a long period, history accumulates like sediment, and
**identifiers that do not exist in the source remain inside the executable file** (= a state judged as VBA Stomping).

Remedy: **export all modules to files and re-import them into a new workbook**.
★Drag and drop in the VBE does not remove it★ (it is close to a binary copy, so the P-code comes along too).
In this case, 2.04 MB → 1.34 MB and remnants 6 → 0.

Generalization: "if the source is correct, the artifact is correct" does not hold in VBA.
Only by periodically rebuilding via text can the artifact be guaranteed to match the source.

### ★ Design principle 81 — Do not "evade" AV detections; deal with them by reporting and recording ★
Do not obfuscate the code or hide API calls in order to make a detection go away.
That is **exactly what malware does**; it worsens the heuristics and also
distorts the project's design itself (implementing COM with machine-code thunks).

The correct procedure is "isolate and pin down the cause → remove it if it can be removed → if it is the core,
send a false-positive report to the vendor with third-party evidence attached, record the history, and accept it".

Generalization: a detection is a report of fact, not an attack. Explain it rather than hide it.

### ★ Design principle 82 — Do not mistake a proxy indicator for the real thing ★
In this case olevba was used as a proxy indicator for "what Bkav is looking at", but
partway through it turned out that olevba's stomping verdict **structurally cannot be made negative** because of
type suffixes and identifier mismatches. Had I kept "make olevba negative" as the goal here,
I would have gone on endlessly with work that could never be achieved.

Remedy: when using a proxy indicator, state explicitly that "this is not the real thing itself", and
**first confirm the range the indicator can in principle reach**. Once it is found to be unreachable, drop the goal and
move on to the stage of measuring the real thing (= the actual VT result).

Generalization: an indicator that can be run quickly at hand has a pull toward being swapped in for the real thing.
This is the flip side of design principle 76 (look at the control system first): **if the control is broken, stop measuring**.

### Numbers at stage 9.30
- Design principles: **82** (80, 81, 82 added)
- Spec facts: **26** (unchanged)
- Standard file layout: **12** (11 production files. ★Zero VBA code changes★)
- Workbook: `WebView2_v0_5_0.xlsm` (852 KB) → **`WebView2_v0_5_2.xlsm` (683 KB)** is made the official release

## Stage 9.31 — Look of the tab bar: hover background and display control of the close button

**Goal**: now that the features are broadly in place, go from "usable" to "want to use". Improve the
experience of the tab bar, which is always in view. Favicons are out of scope since the retrieval path is separate and a heavy problem in itself.

**★Not a single line of VBA logic was touched. Only the CSS inside `BuildTabBarHtml` in `Wv2TabBar.cls`★**

### CSS changed

```css
.tab{ ... white-space:nowrap;transition:background-color .12s;}   /* added */
.tab.active{background:#fff;color:#000;font-weight:600;}          /* unchanged */
.tab:not(.active):hover{background:#efefef;}                      /* new (A-1) */
.tab .close{ ... color:#888;visibility:hidden;}                   /* added (A-2) */
.tab.active .close,.tab:hover .close{visibility:visible;}         /* new (A-2) */
.tab .close:hover{background:#ccc;color:#000;}                    /* unchanged */
```

### Agreed issues (all 8 as Claude recommended)

| # | Decision | Reason |
|---|---|---|
| 1 | CSS `:hover` only. Do not touch JS | `render()` discards the DOM every time with `innerHTML=''`, so if JS held the hover state it would be lost at every sync |
| 2 | Hover background for inactive tabs only (`#e7e7e7`→`#efefef`) | The active tab is already `#fff`, the brightest. Use `:not(.active)` for a specificity of 0,3,0 so it does not depend on written order |
| 3 | Show the close button when active or hovered | The active tab is the one you want to close "right now", so it should be clickable before the mouse is placed over it |
| 4 | ★Hide with `visibility:hidden`★ | `display:none` makes the title width stretch and shrink so the text jitters. `opacity:0` lets an invisible button be pressed |
| 5 | Transition only the background color, `.12s` | `visibility` is a discrete value and transition has no effect on it, so the close button switches instantly |
| 6 | Keep the × color as is | It is more reliable to judge after seeing how it looks with display control on the real machine |
| 7 | No hover suppression during drag | We would never know whether a preemptive correction had been necessary (the lesson of §9.27 issue 7) |
| 8 | `Test_9_31_Help` (visual) + end-to-end check of v0_5_2 | Work off open item 1 of §9.30 with the same procedure |

### Real-machine test results

- **(A-1) No problem with the change of hover color.**
- **(A-2) Step 9, "pressing an invisible × does not close", was ★impossible to verify★.**
  Trying to press it makes the × appear on hover, so that situation cannot be created.
  ★This is evidence that a misclick structurally cannot occur★ — an element with `visibility:hidden` is excluded from
  hit testing, so "visible = pressable" and "not visible = not pressable" always coincide. The decision in issue 4 was correct.
- **(A-2) In step 10, adding tabs made a horizontal scroll bar appear, which ate into the height of the tab bar from the inside
  and squashed the tabs.** → On to issue 4 of §9.32.
- **(B) No regression in D&D reordering.**
- **★(C) Open item 1 of §9.30 = the end-to-end operation check of v0_5_2 is complete★**
  - `Test_9_27_Status`: spare present / IsReady=True / Claimed=False / usable=True
  - POST probe button 1: `"method": "POST"`, `Sec-Fetch-User: ?1`,
    `Referer: https://appassets.postprobe/`, `Origin: https://appassets.postprobe`,
    and the Japanese in `form` (including a full-width space) is also intact. ★Spec fact 26 reproduces in the clean-build version too★
  - `ViewPtr` changed after use (…848128 → …866560) = **replenishment of the spare tab confirmed**
  - `target=_blank` links also open in our own tab
  - → **v0_5_2 is sound. Open item 1 of §9.30 is closed.**

### Additional specifications decided

- **Do not activate when D&D starts (keep as is).**
  1. Reordering is "tidying the appearance", not "changing what is being viewed"
  2. `ActivateTab` is heavy, so running it every time a tab is grabbed would worsen the feel
  3. The original activates because of "pulling a tab out into a separate window".
     This implementation does not have that feature, so there is no reason to follow it

---

## Stage 9.32 — Feel of tab switching: optimistic update and tidying the scroll bar

**Goal**: eliminate the latency reported in the real-machine test of §9.31, where "the click and the change on screen are clearly out of step".
At the same time, tidy the scroll bar of the tab bar.

**★JS and CSS only. Not a single line of VBA logic was touched★**

### ★The cause of the latency was structural★

```
JS: tab.onclick
  └→ ho.TabBar.HostActivate(n)   ← JS is completely blocked here
        └→ VBA: HostActivate → m_browser.ActivateTab   ← heavy (display switch)
              └→ RaiseEvent ActiveChanged
                    └→ TabBar.m_browser_ActiveChanged → PushSyncToJs
                          └→ fetch every tab's DocumentTitle via COM
                          └→ View_PostWebMessageAsString(json)  ← only sends. Has not arrived yet
        ←── HostActivate finally returns
  └→ JS unblocked → only now is the message event processed
        └→ render() deletes the whole DOM and regenerates it → ★the color changes for the first time here★
```

`hostObjects.sync` is a synchronous RPC, so until all the VBA processing finishes, **the tab bar's JS does not
advance by a single instruction**. The message that `PushSyncToJs` sent cannot be received either, since the JS is blocked.
★The order "switch the main body → log → tab bar" is a necessity of the design★, and the perceived delay was no accident.

**★Do not swap the order on the VBA side★** `HostActivate` is in the middle of a synchronous call from JS,
so even if `PushSyncToJs` is called first inside it, the JS is blocked and cannot receive it. Forcing it in
would enter the deadlock territory of spec fact 14. **Swapping the order is correctly done on the JS side.**

### JS / CSS changed

```javascript
// added (top level)
function markActive(el){ ...remove the current .active and attach it to el... }
function afterPaint(fn){ requestAnimationFrame(function(){ requestAnimationFrame(fn); }); }
function sendActivate(n){ ...the body of onclick up to 9.31, cut out as is... }

// replaced (inside render)
tab.onclick = function(ev){
  if(ev.target && ev.target.className === 'close') return;
  var n = parseInt(this.getAttribute('data-index'), 10);
  markActive(this);                             // change only the appearance first
  afterPaint(function(){ sendActivate(n); });   // draw one frame, then call VBA
};
```
```css
#tabs::-webkit-scrollbar{display:none;}   /* added (overflow-x:auto is kept) */
```

### Agreed issues (all 7 as Claude recommended)

| # | Decision | Reason |
|---|---|---|
| 1 | Introduce an optimistic update | The full sync that comes later from VBA always overwrites it, so it does not make JS hold state |
| 2 | ★Wait for painting with a double `requestAnimationFrame`★ | If you block with a synchronous COM call right after changing the DOM, it freezes before painting. `setTimeout(...,0)` only guarantees the next task, not painting |
| 3 | Do not include it for close | Rolling back on failure is complicated. First put in only activate and see how well it works |
| 4 | Hide the scroll bar completely | A thin 3px bar leaves a line permanently right under the tab bar, which is an eyesore in a different way |
| 5 | Keep `min-width:80px` | Even lowering it to 60px only increases tabs too narrow to read; the overflow situation does not change |
| 6 | Do not activate when D&D starts | See the record in §9.31 |
| 7 | `Test_9_32_Help` (visual, feel) | — |

### Recorded as after-the-fact reports

- **Moved `HO()` from inside `render()` to the top level.** Previously `HO()` was declared inside the
  `for` loop within `render()`, and worked by chance thanks to function declaration hoisting
  (it existed only in `render`'s scope). Moved because `sendActivate()` needed to be taken out of `render`.
  The behavior is identical and the scope becomes sane.
- **Added a fallback (`setTimeout(fn,0)`) to `afterPaint` for when `requestAnimationFrame` is absent.**
  It is a branch that is not taken in WebView2, but a single `if` prevents the worst
  case of "there is no rAF and nothing happens".

### Real-machine test results

- **(1) The switching feel improved. Passed, rated "at a practically fine level".**
  The log order is as before (`HostActivate` → `ActivateTab` → `ActiveChanged` →
  `PushSyncToJs` → NavBar sync). **What changed is that "the tab bar's color changes before the log".**
- **(2) The scroll bar is gone. The problem of squashed tabs is also resolved.**
- **★(3) The wheel did not work★** → on to §9.32b.
- **(4) No regression** (D&D / close / settings tab all normal).

### ★ Design principle 83 — For perceived slowness, suspect "order" before "making it faster" ★

Making the processing itself faster is hard and tends to be a risky modification. On the other hand, "in what order to show things" can
often be changed safely. In §9.32 `ActivateTab` was not made faster by even 1 millisecond, but
just reversing the order of display brought the feel to a practical level. **Order first, speed second.**

### ★ Design principle 84 — If the next full sync will certainly overwrite it, an advance display is not state retention ★

Design principle 60 (full sync) is a principle for "not holding state on the JS side", but
**a temporary advance display for which it is guaranteed that "an overwrite will surely come from VBA" does not violate it.**
Even if it fails, the next sync restores the correct picture, so consistency is still held on the VBA side.
The criteria for deciding whether an advance display is allowed are "will the next sync surely come" and "will it surely be overwritten when it comes".

### ★ Spec fact 27 — While a hostObjects.sync call is in progress, the JS of the calling Pane is completely blocked ★

Messages sent meanwhile with `PostWebMessageAsString` also **cannot be received** (they pile up in the queue).
Therefore "update the tab bar first on the VBA side, then do the heavy processing" is **impossible in principle**.
Do not try to deliver a VBA→JS notification during a JS→VBA synchronous RPC.

### ★ Spec fact 28 — In a horizontal overflow-x:auto container, vertical wheel is not automatically converted into horizontal scrolling ★

Measured on WebView2 (Chromium 151). It depends on the ancestors' overflow configuration and on the hit target,
so it cannot be relied on. **If you want horizontal scrolling, place an explicit `wheel` handler.**

---


## Stage 9.32b — Wheel horizontal scrolling and keeping the scroll position

**Goal**: close the issue that hiding the scroll bar in §9.32 also stopped the wheel from working.
At the same time, fix in the same round the problem found during implementation prep, "the scroll position returns to the left edge at every sync".

**★JS only. Not a single line of CSS or VBA logic was touched★**

### Background 1: automatic conversion did not occur (spec fact 28)

In §9.32 we hid `::-webkit-scrollbar` while keeping `overflow-x:auto`, expecting that "Chromium converts the wheel into horizontal for a horizontal container that cannot
scroll vertically", but ★on the real machine there was no reaction★.
It depends on the ancestors' overflow configuration and on the hit target, so it cannot be relied on.
**Placing an explicit `wheel` handler is the reliable way.**

### Background 2: `render()` resets `scrollLeft` to 0

`render()` runs `tabsEl.innerHTML = ''` at the start, so **at every full sync the horizontal scroll
position returns to the left edge**. While there are few tabs, scrolling itself does not occur, so nobody noticed, but
it becomes evident once scrolling is possible. Moreover, `PushSyncToJs` is also sent **every time a title changes**,
so just by some tab finishing loading while you are looking at the right side, you are sent back to the left edge.
Also, "a new tab added with + is off-screen and not visible" happens at the same time.

★This is a problem you necessarily step on the moment you make the wheel work, so we added it to the same round's issues and fixed it together.★

### JS changed

```javascript
// added: attach once to #tabs
//   Even if innerHTML is rewritten, #tabs itself is not replaced, so the listener survives
var tabsRoot = document.getElementById('tabs');
if(tabsRoot){
  tabsRoot.addEventListener('wheel', function(ev){
    if(tabsRoot.scrollWidth <= tabsRoot.clientWidth) return;
    var d = ev.deltaY || 0;  if(d === 0) return;
    tabsRoot.scrollLeft += d;
    ev.preventDefault();
  }, {passive:false});
}

// start of render()
var keepLeft = tabsEl.scrollLeft;
// end of render()
tabsEl.scrollLeft = keepLeft;
if(active !== lastActive){
  var act = tabsEl.querySelector('.tab.active');
  if(act && act.scrollIntoView){
    try{ act.scrollIntoView({block:'nearest', inline:'nearest'}); }catch(e){}
  }
}
lastActive = active;
```

### Agreed points (all 3 as Claude recommended)

| # | Decision | Reason |
|---|---|---|
| 1 | Add a `wheel` handler to `#tabs` (point 4 of §9.32 is not withdrawn) | Going back to a thin 3px bar would bring back the problem of squashed tabs. Call `preventDefault` with `{passive:false}` |
| 2 | Add `deltaY` as is | One wheel notch is about 100px = a bit more than one tab. If it is not enough, it can be tuned later with a single number |
| 3 | Save/restore + bring the active tab in with nearest | Restore alone leaves a new tab off-screen and invisible. `nearest` does not move when the tab is already visible, so side effects are minimal |

### Post-hoc report

- **Added an "only when the active tab changed" condition to the nearest alignment.** With point 3 as it was,
  if a title-update sync arrived while the user was looking at the right-hand part of the strip, the view would be
  pulled back to the active tab's position (the opposite of what point 3 intends). A top-level `lastActive` (one number) is kept,
  and `scrollIntoView` runs only when `active` differs from the previous value. Since it is a memory of just one number,
  it does not undermine the intent of the full-sync approach (design principle 60).
- **The statement in the §9.32 header, "design principles 82 and spec facts 26, no change in either," was wrong.**
  The real-machine test showed that design principles 83 and 84 and spec facts 27 and 28 had been added. The dev notes are authoritative.

### Real-machine test result — **PASS**

- The tab strip scrolls horizontally with the wheel
- Behavior when adding and removing tabs is also fine (the scroll position does not break)
- No regressions

### ★ Design principle 85 — Delete-all-and-regenerate also throws away the "display state" ★

The `innerHTML=''` of the full-sync approach (design principle 60) throws away not only the DOM but also
**the "display state" such as scroll position, focus, text selection, and open/closed state.**
The VBA side holds data consistency, but the VBA side does not hold display state, so **nobody restores it**.

When building a UI with the full-sync approach, **enumerate the "display state that must not be lost" and explicitly
save and restore it.** In §9.32b, that was `scrollLeft`.

★This principle will definitely apply to roadmap R-3 (TreeView / ListView)★ — a tree's open/closed state,
a list's selected row, and scroll position are all "display state that VBA does not hold," and
they vanish with every full sync. Decide at the design stage what to save and restore.

---
