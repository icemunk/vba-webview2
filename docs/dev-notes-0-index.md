# VBA-WebView2 Development Notes — Index

This file is the overall index. For the per-stage details, see the separate volumes.

---

## Stage numbering scheme (introduced in D-1)

The "9.xx" numbers up to stage 9.32b were milestones along a single axis: "turning WebView2 into a browser".
Stage 9.32b brought that axis to a natural close, and since axes of a different nature were expected to run
in parallel from then on, we moved to a scheme that assigns a symbol to each axis.

| Symbol | Axis | Status | File |
|---|---|---|---|
| (no symbol, 1 to 9.2) | Establishing the thunk foundation | **Frozen** | `dev-notes-1-foundation.ja.md` |
| (no symbol, 9.3a to 9.32b) | Browser feature work | Closed out at stage 9.32b. **Frozen** | `dev-notes-2-browser.ja.md` |
| **D** | **DOM / automatic page control (MSHTML equivalent)** | D-7 complete. **Active** | `dev-notes-3-D-axis.md` |
| **E** | **Development environment and tooling (repository / round trip with the workbook / inspectors)** | E-4 complete. **Active** | `dev-notes-4-E-axis.md` |
| **N** | **Network investigation (equivalent of F12's Network tab + PowerShell generation)** | N-4 complete. **Active** | `dev-notes-6-N-axis.md` |
| C | Form controls (TreeView / ListView equivalents) | Not started | (to be created when work starts) |
| **K** | **Foundation and debt repayment (Sentinel, thunk layer, logging infrastructure, existing bugs)** | K-5 complete. **Active** | `dev-notes-5-K-axis.md` |

- Notation examples: heading `## Stage D-1 — ...` / file header `' --- Wv2Pane.cls  Stage D-1 (...) ---`
  / verification Sub `Test_D1_Eval`. Follow-up rounds get a `b` as before, like `D-1b`.
- ★Past 9.xx entries are not rewritten.★ They stay as they are as the history of a completed axis.
- The roadmap's R numbers (R-1 to R-5) are a "list of things we want to do", not stage numbers.
  Read the correspondence as: the result of starting on R-4 is D-1.

### Organization of the development notes (after refactoring)

| File | Scope | Nature |
|---|---|---|
| **`dev-notes-0-index.md`** (this file) | Stage numbering scheme / table of contents of all stages / full list of design principles / full list of spec facts / concept / roadmap / open items / current workbook and file layout | **Always start reading here** |
| `dev-notes-1-foundation.ja.md` | Preface to stage 9.2 | Frozen |
| `dev-notes-2-browser.ja.md` | Stage 9.3a to stage 9.32b | Frozen |
| `dev-notes-3-D-axis.md` | Stage D-1 onward | **Active, where to append (product code)** |
| `dev-notes-4-E-axis.md` | Stage E-1 onward | **Active, where to append (development environment and tooling)** |
| `dev-notes-5-K-axis.md` | Stage K-1 onward | **Active, where to append (foundation and debt repayment)** |
| `dev-notes-6-N-axis.md` | Stage N-1 onward | **Active, where to append (network investigation)** |

★When you write a new stage, you touch only two files: the **volume for that axis** (the body) and `_0_index.md`
(1 line in the table of contents + 1 line each in the principle/fact tables + replace the open items at the end).★
Rounds that change product code go in `_3_D-axis.md`; rounds that change only the development environment or tools go in `_4_E-axis.md`.
**The rule is that an E-axis round does not change a single line of the `Wv2*` product code.** If a change turns out
to be necessary, split it out as a D-axis or K-axis stage.

---

## The concept of VBA-WebView2 (stated explicitly by tarboh, 2026-08)

**The goal is not to reproduce an ordinary browser completely.** A high degree of polish is a good thing,
but the purpose is not to surf the web leisurely and comfortably from VBA. The aim is these two points:

1. **To be excellent and easy to use for automatically controlling web pages**
2. **To provide a modern GUI by laying out the controls on a UserForm in HTML**

**These two points are the criteria for every later decision about what to keep and what to drop.**


---

## Table of contents of all stages

| Stage | Content | File |
|---|---|---|
| Preface | Creating the WebView2Environment | `1_foundation` |
| Stage 1 | A scheme in which the corresponding functions are defined in a standard module and an array of their function pointers is laid out in memory as a virtual function table with `CopyMemory` | `1_foundation` |
| Stage 2 | Staying in a standard module, abolish CopyMemory and build the virtual function table with memory operations via PointerAccesser | `1_foundation` |
| Stage 3 | Same as stage 2, relaying from the standard module's Invoke to the class module's Invoke | `1_foundation` |
| Stage 4 | Abolish the Invoke relay in the standard module and have WebView2 call the class's Invoke directly | `1_foundation` |
| Stage 5 | Modify the stage 4 thunk so that the class's Invoke can be called as a `Function` | `1_foundation` |
| Stage 6 | Change the thunk writes from CopyMemory to the PointerAccesser scheme | `1_foundation` |
| Stage 7 | Switch the thunk region allocation from VirtualAlloc to the AddressOf acquisition scheme | `1_foundation` |
| Stage 8 | Introduce a two-step jump structure via an EntryPoint stub and make use of the EBMode check | `1_foundation` |
| 8.5 | Reconnaissance phase: analysis of what the EntryPoint stub really is | `1_foundation` |
| 8.6 | Reconnaissance phase: verifying the executability of the module-level array region | `1_foundation` |
| 8.7 | Reconnaissance phase: reconnaissance of a strategy for allocating a huge Sub procedure region | `1_foundation` |
| 8.8 | Reconnaissance phase: verifying the executability of the free space of a Sub procedure | `1_foundation` |
| 8.9 | Minimal implementation of the VirtualAlloc + stub-copy scheme (1 handler) | `1_foundation` |
| 9.0 | Verifying the slot pool + alive flag scheme (1 handler) | `1_foundation` |
| 9.1 | GenericHandler-ization and a real reference count implementation (1 handler) | `1_foundation` |
| 9.2 | Multiple handler support + embedded vtableObj + full implementation of the this reverse lookup | `1_foundation` |
| 9.3a | Introducing the upper-level class (Class8 Environment) and dynamic binding | `2_browser` |
| 9.3b | Controller class (Class9) + dcf unification + reconnaissance log | `2_browser` |
| 9.3c | WebView2 display milestone (Bing appears on screen) | `2_browser` |
| 9.4a | Persistent handler mechanism + NavigationStarting reconnaissance | `2_browser` |
| 9.4b | Rolling out NavigationCompleted + DocumentTitleChanged + navigation control | `2_browser` |
| 9.4c | Two-way WebMessage communication | `2_browser` |
| Summary of 9.3 to 9.4 | Overall wrap-up and handover to stage 9.5 onward | `2_browser` |
| 9.5 | Sentinel mechanism (complete elimination of leaks on reset) + finding the true cause of the VirtualFree problem | `2_browser` |
| 9.6 | Rolling out NewWindowRequested + establishing the naming convention by origin prefix | `2_browser` |
| 9.7 | Introducing the IID check + UserForm mouse resizing | `2_browser` |
| 9.8a | Formally fixing the class names | `2_browser` |
| 9.8b | Full coverage of the View base IFs + Settings family + Ctrl_Close | `2_browser` |
| 9.8c | View_ExecuteScript + improved safety of Ctrl_Close | `2_browser` |
| Summary of 9.8 | Summary of stage 9.8 | `2_browser` |
| 9.9 | SPA support (extracting ICoreWebView2_2 + HistoryChanged + DOMContentLoaded + Exe… | `2_browser` |
| Summary of 9.9 | Summary of stage 9.9 | `2_browser` |
| 9.10 | Wv2Browser umbrella class (tab management layer) + dealing with the Sentinel bug | `2_browser` |
| 9.11 | Implementing the tab UI (tab bar) | `2_browser` |
| 9.11b | Skeleton of the WebView2-based tab bar Wv2TabBar + layout split | `2_browser` |
| 9.11c | VBA→JS direction: rendering the tab list (full sync + ready handshake) | `2_browser` |
| 9.11d | JS→VBA direction: enabling clicks (Layer 2 complete) | `2_browser` |
| 9.12 | Address bar and navigation bar (Layer 3 complete) | `2_browser` |
| 9.13 | Cleanup: unifying the JSON helpers + removing throwaway tests | `2_browser` |
| 9.14 | Search-term fallback (default Google) | `2_browser` |
| 9.15 | Drag reordering of tabs | `2_browser` |
| 9.16 | Search engine switching + cleanup (extracting Wv2Url / consolidating Wv2Tests) | `2_browser` |
| 9.17 | The switching entry point for search engines + publishing the reference to the running Browser | `2_browser` |
| 9.18 | Settings tab + search engine UI via AddHostObjectToScript | `2_browser` |
| 9.19 | Preventing duplicate settings tabs + default display of the settings screen preview | `2_browser` |
| 9.20 | Persisting the search engine selection | `2_browser` |
| 9.21 | Experiment: converting activate to hostObjects + confirming it holds with a real URL Navigate | `2_browser` |
| 9.22 | Experiment: converting close / reorder to hostObjects | `2_browser` |
| 9.23 | Cleanup: consolidating the activate / close / reorder processing into Host methods | `2_browser` |
| 9.24 | Converting the NavBar to hostObjects + consolidating processing into Host methods | `2_browser` |
| 9.25 | Verification: whether zero-argument Host methods work → permanent removal of reserved | `2_browser` |
| 9.26 | Isolation: NewWindowRequested delegation probe → determining the cause of the POST 405 | `2_browser` |
| 9.27 | The main event: prewarm delegation via put_NewWindow | `2_browser` |
| 9.28 | Making prewarm delegation the default + 2 small fixes around the spare | `2_browser` |
| 9.28b | Making the POST probe's destinations 3 routes / making httpbingo.org the primary | `2_browser` |
| 9.29 | Input field for a custom search URL + persisting custom (+ cleaning up the spare reservation) | `2_browser` |
| 9.30 | Isolating VirusTotal's Bkav Pro verdict (no code changes) | `2_browser` |
| 9.31 | Look of the tab bar: hover background and display control of the close button | `2_browser` |
| 9.32 | Perceived tab switching: optimistic update and tidying the scroll bar | `2_browser` |
| 9.32b | Wheel horizontal scrolling and keeping the scroll position | `2_browser` |
| **D-1** | **Synchronous retrieval primitive for ExecuteScript (EvalSync)** | `3_D-axis` |
| **D-2** | **Element registry and reading with Wv2Element** | `3_D-axis` |
| **D-3** | **Writing and operating (`.Value =` / `.Click` / `.SetAttribute`) + waiting on DOM conditions** | `3_D-axis` |
| **D-4** | **Settle waiting (in-page probe / explicit signals / diagnostic log / Google Maps proof)** | `3_D-axis` |
| **D-5** | **Practical use: address → coordinates (`Wv2Maps`) + worksheet integration** | `3_D-axis` |
| **D-6** | **`QuerySelectorAll` and lifetime management + choosing from a candidate list** | `3_D-axis` |
| **D-7** | **Progress display and cancellation (stop with Esc / status bar / resume)** | `3_D-axis` |
| **E-1** | **Moving development to Claude Code and building the round-trip foundation with the workbook** | `4_E-axis` |
| **E-2** | **Closing the round trip with the workbook using `import.ps1`** | `4_E-axis` |
| **E-3** | **Adding new modules (`-AllowNew` of `import.ps1`)** | `4_E-axis` |
| **K-1** | **Debug log file output (Wv2Log)** | `5_K-axis` |
| **K-2** | **Fixing the bug where the settings tab gets crushed by a search** | `5_K-axis` |
| **K-3** | **Replacing the 76 characters lost to CP932** | `5_K-axis` |
| **K-4** | **Slow release when the form is closed (measurement and the `about:blank` countermeasure)** | `5_K-axis` |
| **K-5** | **Starting WebView2 crashes the VBE on IME input** | `5_K-axis` |
| **N-1** | **Wiring up `WebResourceRequested` (list of URL / method / type)** | `6_N-axis` |
| **N-2** | **Request headers and body (`IStream` reading, masking secrets)** | `6_N-axis` |
| **N-3** | **Response status and headers (the result is attached to one row of the list)** | `6_N-axis` |
| **N-4** | **Response body (`GetContent` = the first asynchronous one. Settling how compression is handled)** | `6_N-axis` |
| **E-4** | **Escaping OneDrive, and the doubled attribute lines** | `4_E-axis` |
| (outside stages) | Proposal from outside — the verdict method for "waiting for completion" in an SPA | `3_D-axis` |

---

## Full list of design principles (1 to 122)

The full text of each item is in the section of its "source stage". When you establish a new principle, add one line at the end of this table.

| # | Heading | Source stage |
|---|---|---|
| 1 | A one-shot handler is released automatically right after the completion notification (plan D1-α) | 9.3-9.4 |
| 2 | A persistent handler holds its token and is kept alive until it is explicitly removed (plan P1) | 9.3-9.4 |
| 3 | Automate rolling out handler kinds with HandlerKind-linked arrays (plan r) | 9.3-9.4 |
| 4 | Methods called through dynamic binding are written as Public Sub | 9.3-9.4 |
| 5 | Set m_state = Es_Waiting before the API call (countermeasure for synchronous callbacks) | 9.3-9.4 |
| 6 | dcf does not judge HRESULTs automatically; it returns the return value as is | 9.3-9.4 |
| 7 | Merge the Controller and CoreWebView2 into one class, and show the origin with the Ctrl_ / View_ prefixes | 9.3-9.4 |
| 8 | Pass structure arguments as pointers with VarPtr | 9.3-9.4 |
| 9 | Get the HWND size directly in pixels with GetClientRect | 9.3-9.4 |
| 10 | On a VBA reset, avoid crashes with the alive flag and tolerate leaks (plan R4. Resolved by the Sentinel mechanism of stage 9.5) | 9.3-9.4 |
| 11 | Always read a handler's (this, arg1, arg2) as WebView2's (C++ this, first argument, second argument) | 9.3-9.4 |
| 12 | NavigationStarting / NewWindowRequested / FrameNavigationStarting sit at different levels of the event hierarchy, so use them accordingly | 9.3-9.4 |
| 13 | A persistent handler completes with one AddRef / one Release per event | 9.3-9.4 |
| 14 | Release COM references from child to parent (View → Ctrl → Environment) | 9.3-9.4 |
| 15 | Bulk removal of persistent handlers is rolled out automatically with a HandlerKind-linked loop | 9.3-9.4 |
| 16 | The firing order of events is NavigationStarting → DocumentTitleChanged → NavigationCompleted | 9.3-9.4 |
| 17 | The args of DocumentTitleChanged is always NULL, so take the value via the sender | 9.3-9.4 |
| 18 | NavigateToString is really a data: scheme URL, and the source URL becomes about:blank | 9.3-9.4 |
| 19 | TryGetWebMessageAsString succeeds only when JS sent a string | 9.3-9.4 |
| 20 | The page-load-complete flag is provisionally held by the Pane and will move to the umbrella class in the future | 9.3-9.4 |
| 21 | UTF-16 two-way communication passes through without garbling (Japanese too) | 9.3-9.4 |
| 22 | Write AddXxx / RemoveXxx / OnXxx separately for each kind (readability over generalization, plan i) | 9.3-9.4 |
| 23 | Navigation control methods are named with the View_ prefix + the SDK name | 9.3-9.4 |
| 24 | Mouse resizing of the UserForm is done with the SetWindowLongPtr + GA_ROOT scheme (implemented in 9.7b) | 9.3-9.4 |
| 25 | On a VBA reset, UserForm_Terminate is not called, but the Class_Terminate chain of the fields does fire | 9.3-9.4 |
| 26 | Version-extension interfaces are added at the point they become necessary (implement only the base, plan 4) | 9.3-9.4 |
| 27 | Persist the region address in an environment variable as a decimal string (no workbook-modified flag and no registry pollution) | 9.3-9.4 |
| 28 | Reclaim the region on reset in the order "alive flag to 0 → VirtualFree → clear the environment variable" (Sentinel plan β) | 9.3-9.4 |
| 29 | When an &H literal exceeds &H7FFF, always add the & suffix (the trap of it being evaluated as an Integer) | 9.3-9.4 |
| 30 | Call important Win32 APIs in Function form and always check the return value | 9.3-9.4 |
| 31 | Name COM-derived methods in the form `<origin><version>_<method name>` (Ctrl_ / View_ / Env_) | 9.6 |
| 32 | Unless NewWindowRequested is suppressed with put_Handled, a new WebView2 window is created inside the Excel process | 9.6 |
| 33 | Implement QueryInterface strictly and return E_NOINTERFACE for an unexpected IID | 9.7 |
| 34 | UserForm resize tracking is done by "saving the design-time gap and reproducing it" | 9.7 |
| 35 | UserForm_Resize also fires falsely during Initialize, so return early with a ready flag | 9.7 |
| 36 | Decide the UserForm size in the designer, and do not assign to Me.Width / Me.Height in Initialize | 9.7 |
| 37 | For WebView2 resize tracking it suffices to fire Ctrl_PutBounds from UserForm_Resize (no subclassing needed) | 9.7 |
| 38 | Classes that map 1 to 1 to the SDK get the Wv2 prefix; foundation classes are named without a prefix | 9.8a |
| 39 | Make a rename-only sub-stage, where "zero logic change" is guaranteed, independent | 9.8a |
| 40 | A renamed file states its "former name" in the header comment | 9.8a |
| 41 | Treat Settings as not taking effect, with a few exceptions, until the next top-level navigation | 9.8b |
| 42 | Do not fetch Settings at Init; fetch it lazily on first access and cache it | 9.8b |
| 43 | Release order in Class_Terminate is from child to parent (Settings → View → Ctrl → Environment) | 9.8b |
| 44 | BOOL getters are written in the style of handling the HRESULT internally and returning As Boolean | 9.8b |
| 45 | When passing a thunk slot as a COM object, always pass `h.Slot + VTABLE_OBJ_OFFSET` | 9.8c |
| 46 | For one-shot callbacks that may run concurrently, the Pane does not hold the handler reference (plan a) | 9.8c |
| 47 | The callback ID is held as a member on the handler side and set through a Friend Property right after Acquire | 9.8c |
| 48 | Safety after Ctrl_Close is guaranteed by a hybrid of "clearing members + the m_closed flag" | 9.8c |
| 49 | The token of a persistent handler is also cleared at Ctrl_Close (prevents removing from an already-released handler) | 9.8c |
| 50 | Derived IFs are fetched lazily with `m_pXxxN` + `EnsureXxxN`, and the IID is assembled at the call site each time | 9.9 |
| 51 | Release a derived IF before the base (the pointer values are the same but the refcounts are managed separately) | 9.9 |
| 52 | For events of a derived IF, both registration and removal are fired with the derived IF's pointer | 9.9 |
| 53 | Leave a "32-bit-ization point" comment at places that matter for going 32-bit, and do the actual work together in a separate stage | 9.9 |
| 54 | ExecuteScript-family helpers assume an "expression", and statements must be wrapped in an immediately invoked function | 9.9 |
| 55 | Memory primitives that involve SEH are called only after confirming 3 conditions with VirtualQuery (On Error cannot protect against them) | 9.10 |
| 56 | Always obtain the View from the Controller before calling a Navigate-family method (completion of Controller creation alone leaves m_pView at 0) | 9.10 |
| 57 | Consolidate visibility control in the one place ActivateTab, and make add / remove merely delegate to it | 9.10 |
| 58 | In structures where circular references arise, provide an explicit Shutdown at the upper level and cut the references before release | 9.10 |
| 59 | A child's initial geometry reuses the rectangle most recently applied by the upper level, cached | 9.11c |
| 60 | VBA→JS state sync defaults to full sync (declarative); differential sync is not used | 9.11c |
| 61 | The first VBA→JS send is fired only after receiving the ready signal from JS (ready handshake) | 9.11c |
| 62 | Notification paths from a Pane to the upper level each hold a separate reference per purpose (do not mix the back reference with the WebMessage forwarding destination) | 9.11c |
| 63 | For WebView2 operations called from inside an event handler, choose synchronous / asynchronous by "whether it includes waiting for completion" | 9.11d |
| 64 | When communicating only with the JS of our own HTML, do not build a general-purpose JSON parser; a minimal extraction function suffices | 9.11d |
| 65 | Do not use focus state as a proxy for "is being edited"; hold it in an explicit flag | 9.12 |
| 66 | Toggling the disabled attribute knocks the focus away, so assign only when the value changes | 9.12 |
| 67 | Events needed to display a UI are subscribed to by the display side itself (do not change a lower-level event contract for the convenience of the upper level) | 9.12 |
| 68 | Address bar input is sorted into URL / search term by a three-way branch, and encoding is done byte by byte | 9.14 |
| 69 | Do not add a dedicated event for reordering; route it through the existing activation path and ride it on the full sync | 9.15 |
| 70 | Pure conversions that do not read state go to a standard module; state-dependent policy stays in the main class | 9.16 |
| 71 | The ownership of verification code is decided by "coupling with the thing verified" (if tightly coupled to Private members, it stays in that module) | 9.16 |
| 72 | For a defect that reproduces only in production, try to synthesize a reproduction environment if it can be put in the form of a controlled experiment | 9.27 |
| 73 | Before building a mechanism that waits inside an event, consider whether a "state in which there is no need to wait" can be prepared first | 9.27 |
| 74 | The full-sync approach is self-healing against missed events | 9.27 |
| 75 | Make the scaffolding of a controlled experiment (external services for verification) redundant | 9.28b |
| 76 | When the primary fails, first see whether the control side "fails as expected" | 9.28b |
| 77 | In persistence, decide per item beforehand whether to "save the name or save the value" | 9.29 |
| 78 | A "DoEvents that waits for completion" is a re-entrancy path (hold the in-progress flag as a counter) | 9.29 |
| 79 | The more a defect looks like "this one line fixes it", the more you should pin down the cause before fixing | 9.29 |
| 80 | Do a clean build before distribution or scanning (export → import into a new workbook) | 9.30 |
| 81 | Do not "evade" AV detections; deal with them by reporting and recording | 9.30 |
| 82 | Do not mistake a proxy indicator for the real thing (first check the range the indicator can reach) | 9.30 |
| 83 | For perceived slowness, suspect "order" before "making it faster" | 9.32 |
| 84 | If the next full sync will certainly overwrite it, an advance display is not state retention | 9.32 |
| 85 | Delete-all-and-regenerate also throws away the "display state", so enumerate what to save and restore | 9.32b |
| 86 | Stage numbers are assigned as "axis symbol + serial number", and past axes are frozen and not rewritten | D-1 |
| 87 | When using an asynchronous API with a synchronous face, receive the result not in a single variable but in a pending table keyed by callbackId | D-1 |
| 88 | In areas where the debugger cannot be used, place a guard that turns "freezing" into an "immediate error" | D-1 |
| 89 | When adding pre- and post-processing to an existing procedure that has an Exit in the middle, split the body out into `_Core` and wrap it in a wrapper | D-1 |
| 90 | Check new identifiers against the VBA reserved word list, "ignoring case" | D-1 |
| 91 | Judge the validity of state on the side that actually holds that state (a local cache is limited to a fast path that errs only toward the safe side) | D-2 |
| 92 | For a mechanism placed where the execution environment disappears, do not inject it; make it self-healing by having "the expression that uses it prepare it itself" | D-2 |
| 93 | For an API that returns the same type on failure as on success, always provide a paired entry point to ask about success or failure | D-2 |
| 94 | For a round trip that involves a conversion, judge by "whether what came back matches the original" (the success of one direction alone misses best-fit substitution) | E-1 |
| 95 | Do not fix a constraint of the development environment on the spot; make it a permanent inspector (trust an inspector only after confirming it fires on known bad data) | E-1 |
| 96 | A destructive tool defaults to "preview" and requires an explicit flag to execute (a backup is after the accident, a preview is before it) | E-2 |
| 97 | The acceptance condition for a round-trip tool is "no diff on the second run" (idempotence) | E-2 |
| 98 | Verify through the same path as production (observation within the same process lies through the cache. Calling an existing tool as a child process is shorter and more correct) | E-3 |
| 99 | Do "confirming there is nothing" only after confirming the target could be read (a search on an empty string answers "nothing" to everything) | K-1 |
| 100 | When a verification fails, first suspect whether the "way of measuring" can actually measure the target (do not mistake a limitation of the tool for a defect of the target) | K-1 |
| 101 | Before choosing a symptomatic treatment, count the paths by which the symptom appears (eliminating the true cause fixes them all without enumerating the paths) | K-2 |
| 102 | Information that an API hides cannot be verified with that API (if there is a hiding layer, prepare a tool that looks from outside it) | E-4 |
| 103 | Before measuring, confirm that the property you want to measure is preserved along the measurement path (check the probe itself) | D-3 |
| 104 | A waiting API distinguishes "not yet" from "failure", and stops waiting the moment it knows it is a failure | D-3 |
| 105 | Put verification verdicts in a place that can be searched later, not a place where they flow past and vanish | D-3 |
| 106 | A setting handed to the other side is sent every time, even when empty (if not sent, the previous state remains on the other side) | D-4 |
| 107 | Identifiers for grabbing an external target are decided by observation, not written by guesswork. Build the observing tool first | D-4 |
| 108 | Provide both a general-purpose compromise and an exact specification, and default to the general-purpose side (write down what is being discarded) | D-4 |
| 109 | "It has settled" can be judged only within the observable range. State plainly that what cannot be observed cannot be waited for | D-4 |
| 110 | Test a stateful counterpart on the "second time" (verifying with a single item misses bugs where the previous state remains) | D-5 |
| 111 | Do not write down a value that cannot be judged; leave the reason instead (a plausible-looking value is believed by later stages) | D-5b |
| 112 | Do not use an indirect indicator to judge state. If it can be counted directly, count it | D-6b |
| 113 | Evaluate intermittent observation not by "count" but by "coverage of time" | D-7d |
| 114 | Make verification that a human observes visible to the human (stop it so it can be seen / wait before pressing) | D-7e |
| 115 | ★Do not write a COM vtable index by guesswork; settle it with the SDK header★ (a mismatch crashes) | N-1 |
| 116 | A self-contained verification says nothing when the closed world itself is the cause. ★Mix in one control★ | N-1 |
| 117 | Do not decide the classification the other side returns in your own words. ★Write out the observed values before judging★ | N-1 |
| 118 | When an observation touches the target, count "that you have not broken it" ★via a separate path★ | N-2 |
| 119 | Mask secrets not "when saving" but "when emitting" (leaving an unmasking entry point) | N-2 |
| 120 | Do not wait for "it is ready" in place of "it is done". ★The condition to wait for is the very event you want to see★ | N-3 |
| 121 | ★A verdict that did not run is more dangerous than a FAIL★ (it vanishes from the numbers). Always log that it was skipped | N-3 |
| 122 | When there can be two answers, write ★a verdict that always resolves to one of them★ (turn the unknown into a "verdict" rather than something to "look into") | N-4 |

---

## Full list of spec facts (1 to 80)

The full text of each item is in the section of its "source stage". When you establish a new fact, add one line at the end of this table.

| # | Heading | Source stage |
|---|---|---|
| 1 | WebView2 Settings do not take effect, with a few exceptions, until the next top-level navigation | 9.8b |
| 2 | Ctrl_Close has the side effect of releasing all registered event handlers | 9.8b |
| 3 | A throw in ExecuteScript comes back as HRESULT=S_OK + result="null" (indistinguishable from a normal null) | 9.8c |
| 4 | A VBA UserForm automatically re-Initializes on access to `UserForm1.<member>` | 9.8c |
| 5 | QueryInterface for ICoreWebView2_2 returns the same pointer value as the base (the refcounts are independent) | 9.9 |
| 6 | The origin of a page shown with NavigateToString is about:blank ("null"), and pushState cannot be used | 9.9 |
| 7 | When NavigateToString completes, HistoryChanged fires automatically twice in a row | 9.9 |
| 8 | DOMContentLoadedEventArgs.get_NavigationId returns a UINT64 (a 64-bit-specific getter is needed) | 9.9 |
| 9 | ExecuteScript also returns a JS syntax error as HRESULT=S_OK + result="null", indistinguishable from a runtime exception | 9.9 |
| 10 | On the path after Sentinel detects MEM_FREE, there is a path that does ReadLongPtr on a freed address and causes an AV | 9.9 |
| 11 | The argument of put_IsVisible is a Win32 BOOL (True=1), so normalize VBA's True (-1) to 1/0 | 9.10 |
| 12 | HistoryChanged fires for all of pushState / replaceState / popstate, and fires twice on the first Navigate | 9.10 |
| 13 | SetVirtualHostNameToFolderMapping gives ACCESS_DENIED for the root "/", so spell out the file name | 9.10 |
| 14 | Synchronously waiting from inside an event handler for another asynchronous WebView2 operation to complete deadlocks | 9.10 |
| 15 | PostWebMessageAsString crashes if it arrives before the JS-side listener is registered (it is not queued) | 9.11c |
| 16 | VBA's AscW returns characters of 0x8000 and above as a signed Integer (it becomes negative) | 9.11c |
| 17 | Input whose part before the colon is letters only is treated as a scheme and passed through, which is correct behavior for a browser | 9.14 |
| 18 | If AddHostObjectToScript is called before NavigateToString, JS can grab hostObjects from its first run | 9.18 |
| 19 | Spec fact 18 also holds for Navigate to a real URL (via virtual host mapping) | 9.21 |
| 20 | The EBMode gate cannot prevent a break transition during the initialization burst (a structural window remains) | 9.22 |
| 21 | String arguments of hostObjects.sync are passed as BSTR directly and arrive intact (even full-width spaces are not lost) | 9.24 |
| 22 | The sync proxy of hostObjects does not call VBA on member retrieval; it is deferred until the call () | 9.25 |
| 23 | If NewWindowRequested is delegated to the runtime, POST and headers are fully preserved | 9.26 |
| 24 | An existing CoreWebView2 of another Controller can be passed to put_NewWindow (combining it with put_Handled(TRUE) is correct) | 9.27 |
| 25 | Events after put_NewWindow may fire before we adopt it | 9.27 |
| 26 | With put_NewWindow delegation, even Sec-Fetch-User: ?1 is preserved | 9.28b |
| 27 | While a hostObjects.sync call is in progress, the JS of the calling Pane is completely blocked | 9.32 |
| 28 | In a horizontal overflow-x:auto container, vertical wheel is not automatically converted into horizontal scrolling | 9.32 |
| 29 | The result JSON of ExecuteScript has its keys reordered alphabetically | D-1 |
| 30 | In the result JSON, `<` arrives escaped as `<` | D-1 |
| 31 | JSON.stringify(undefined) returns undefined, so the value key disappears from the result altogether | D-1 |
| 32 | While JS is blocking the thread, ExecuteScript completion notifications arrive all together after the block is lifted | D-1 |
| 33 | In the result JSON only `<` is escaped as `<`, and `>` arrives as is | D-2 |
| 34 | A VBA String is a sequence of UTF-16 code units, so surrogate pairs are restored by adding each half with `ChrW$` | D-2 |
| 35 | VBA source is held in the project's code page (CP932), and characters that cannot be represented are permanently lost on import | E-1 |
| 36 | The VBE Export writes CP932 / CRLF / no BOM (the Import side is settled in spec fact 41) | E-1 |
| 37 | `.frx` changes by a few bytes on every export even if the form has not been changed | E-1 |
| 38 | VBA unifies the case of an identifier to its first declaration and also rewrites the existing occurrences (Value→value / Y→y / ptr→Ptr) | E-1 |
| 39 | `CodeModule.Lines` does not return member attribute lines but `Export` writes them ★The latter half (that AddFromString preserves them) is wrong. Corrected in K-2★ | E-2 |
| 40 | `AddFromString` adds one newline at the end. `CountOfLines` immediately after writing disagrees with the value after saving | E-2 |
| 41 | `VBComponents.Import` reads the file as the ANSI code page (CP932). UTF-8 without BOM becomes garbled | E-3 |
| 42 | A file passed to `Import` must be CRLF, or the header is not recognized and the class is created as a standard module | E-3 |
| 43 | Adding a new module rewrites the spelling of identifiers across the whole project. Moreover, it cannot be observed from the importing process | E-3 |
| 44 | `ADODB.Stream.LoadFromFile` cannot open a file that another handle has open, regardless of the share mode | K-1 |
| 45 | `Debug.Print` displays in CP932 so it shows `?`, but the String itself is intact. If written to a file in UTF-8, it is preserved correctly | K-1 |
| 46 | The VBE drops a redundant `&` on `&H` literals, but keeps it for those in 0x8000 to 0xFFFF (backing up the danger zone of design principle 29) | K-1 |
| 47 | An attribute line passed to `AddFromString` becomes doubled with the one generated by the VBE and is exposed. ★It is visible neither from `CodeModule` nor from `Export`★ (correction of 39) | E-4 |
| 48 | A workbook on OneDrive is opened by Excel with a cloud URL, and AutoSave defaults to ON | E-4 |
| 49 | `ExecuteScript` does not wait for a Promise (even if the return value is a Promise, it returns immediately as `{}`) | D-3 |
| 50 | The value tracker of React and the like overlays `value` on the element instance, so changes go unnoticed unless you `call` the native setter on the `prototype` | D-3 |
| 51 | The `isTrusted` of a synthetic event dispatched from JS is false (the same for `e.click()`) | D-3 |
| 52 | One `EvalSync` round trip measures about 38 ms (the effective interval of 100 ms polling is about 140 ms) | D-3 |
| 53 | `Timer` returns to 0 at midnight (elapsed-time subtraction goes negative and the wait loop never exits) | D-3 |
| 54 | The HWND of a control on a UserForm can only be obtained after `Show` (`hWnd_Frame` becomes 0) | D-3 |
| 55 | If Edge is running, `msedge.exe` hands off to the existing instance and exits immediately (`--headless` cannot be used) | D-3 |
| 56 | The Google Maps map is drawn with canvas / WebGL, and appears neither in the DOM nor in fetch/XHR | D-4 |
| 57 | Delivery of `MutationObserver` records is a microtask, so they cannot be read within the same JS turn | D-4 |
| 58 | Google Maps search keeps streaming results while leaving `/search?tbm=map` open (in flight for over 20 seconds) | D-4 |
| 59 | IDs on real sites can change to auto-generated ones (`#searchboxinput` → `id=ucc-1`) | D-4 |
| 60 | The standard output of `msedge --headless` becomes empty when launched directly from `subprocess` (it can be captured via Start-Process) | D-4 |
| 61 | If Excel has the workbook open, `Workbooks.Open` succeeds as read-only, and write-backs are silently lost | D-4 |
| 62 | A Maps search first changes to `/maps/search/<term>/@<previous center>` and later settles on `/place/...` | D-5 |
| 63 | In a Maps URL, `!3d!4d` are the coordinates of the place itself, and `/@` is the center of the map | D-5 |
| 64 | With `Application.StatusBar = False`, control does not return and `"FALSE"` remains. `Empty` / an empty string does return it | D-7d |
| 65 | The VBA runtime's Esc arrives regardless of focus. Error 18 of `xlErrorHandler` does not trace back to `On Error` when it occurs inside a COM callback | D-7b |
| 66 | `check_idents.py` does not look at labels. Spec fact 38 is triggered even by a case difference in labels | D-7b |
| 67 | ★In an Excel that has launched WebView2, typing into the VBE Immediate window with an IME can crash it. **It does not recover even after closing**★ | K-5 |
| 68 | If there is a running page, COM calls to a ★non-active Pane★ are made to wait 40 times longer (release becomes slow) | K-4 |
| 69 | ★Requests served via `SetVirtualHostNameToFolderMapping` do not appear in `WebResourceRequested`★ (the serving itself is fine) | N-1 |
| 70 | A JS `fetch()` request arrives as `XML_HTTP_REQUEST`(7), not `FETCH`(8) | N-1 |
| 71 | `http://127.0.0.1:<port>` is not stopped by mixed content even when called from an https page (but avoid ports that Chromium blocks) | N-1 |
| 72 | ★The `IStream` of `get_Content` advances its position when read, but resetting with `Seek(0)` does not break the communication★ (measured) | N-2 |
| 73 | When a string body is passed to `fetch()`, the Content-Type is `text/plain;charset=UTF-8` (a CORS simple request) | N-2 |
| 74 | An HTML form POST is `x-www-form-urlencoded`, and ★Japanese is carried percent-encoded★ | N-2 |
| 75 | `MultiByteToWideChar(CP_UTF8, 0, ...)` turns a multibyte sequence cut off midway into U+FFFD and returns the rest (adding `MB_ERR_INVALID_CHARS` makes the whole conversion fail) | N-2 |
| 76 | ★`HttpResponseHeaders::GetIterator` is 7, `HttpRequestHeaders` is 8★ (the response side has no `put_` family, so the ordering is shifted) | N-3 |
| 77 | ★`WebResourceResponseReceived` has no filter★ Images and CSS all come through. The args also carry no type | N-3 |
| 78 | An HTTP/2 response has no reason phrase, so `get_ReasonPhrase` returns an empty string | N-3 |
| 79 | ★`GetContent` returns the already-decoded body★ Even with `content-encoding: gzip/zstd` attached, the contents are plain text | N-4 |
| 80 | For `204 No Content`, the completion of `GetContent` arrives as `0x800700E8` (ERROR_NO_DATA). ★Not a failure, but "there is no body"★ | N-4 |
---

## Roadmap (medium- to long-term development candidates)

### R-1. Implementing methods of the additional ICoreWebView2 interfaces

Select by "what becomes possible" rather than by individual method.

**High value (directly tied to concept 1)**
- ★`CallDevToolsProtocolMethod`★ — on its own opens up the whole CDP. Network interception, DOM operations,
  cookies, screenshots and even PDF are within reach; the range is extraordinarily wide. **Outstanding cost-effectiveness**
- `WebResourceRequested` + `AddWebResourceRequestedFilter` — intercepting and rewriting requests
- `NavigateWithWebResourceRequest` (ICoreWebView2_2) — building and sending a POST ourselves
- `AddScriptToExecuteOnDocumentCreated` — plant hooks before the page loads. The foundation of R-4
- `get_CookieManager` (ICoreWebView2_2) — saving and restoring authentication state
- `PrintToPdf` (ICoreWebView2_7) — report output. A good match with Excel VBA

**Can be given lower priority**: Zoom / Fullscreen / AcceleratorKeyPressed / PermissionRequested.
They are about "being like an ordinary browser" and fall outside the concept. `ProcessFailed` alone is an exception from the viewpoint of robustness.

### R-2. Starting up without WebView2Loader.dll ★High feasibility★

The procedure shown by the C wrapper implementation (`my_webview2.c`, provided by tarboh) has 4 steps, all of which can be done in VBA.

1. `RegOpenKeyExW` (HKLM, with `KEY_WOW64_32KEY`) →
   read the `EBWebView` value of `Software\Microsoft\EdgeUpdate\ClientState\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}`
2. Append `\EBWebView\x64\EmbeddedBrowserWebView.dll` to the end
3. `LoadLibraryW` → `GetProcAddress("CreateWebViewEnvironmentWithOptionsInternal")`
4. Call the obtained function pointer with the existing `dcf`
   (`checkRunningInstance=1, runtimeType=0, userDataFolder, options, handler`)

★**The biggest benefit is not a feature but the distribution form**★ — bundling `WebView2Loader.dll` becomes unnecessary,
which brings us one step closer to **distributing a single xlsm**. It also fits the "zero external libraries" line.

**Notes for when work starts (recorded in advance)**
- ★Do not imitate the C version's `FreeLibrary(module)`★ The callbacks arrive asynchronously later, yet it
  unloads the DLL. It is quite likely that it only happens to stay alive because the Edge main body references the same DLL.
  **The VBA version does not FreeLibrary** (hold it until process exit)
- `CreateWebViewEnvironmentWithOptionsInternal` is a **non-public API**. It can break with an Edge update →
  **keep the existing Loader route as a fallback** (the same idea as design principle 75)
- Currently `hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(path), 0, pVTableObj)` passes
  **NULL** as the third argument (options), so the migration can be done without changing anything there
  ※**Correction of the premise (as of D-1)**: We had evaluated R-2 as "improving the distribution form", but this is wrong.
    **WebView2Loader.dll is bundled with the current Excel from the start** (this project reuses what is
    prepared for Office add-ins), and there was never any need to attach it to the workbook.
    Therefore R-2 gains almost nothing, and the only thing left is the risk of becoming
    vulnerable to Edge updates by moving to a non-public API. **Its priority is lowered to roughly the bottom.**
    Also, we had tied R-5 (extensions) to R-2 as "R-2 provides the scaffolding for environment options", but
    hand-building the options vtable is the same work via the Loader as well, so
    **R-2 is not a prerequisite of R-5**. Evaluate R-5 independently.
    ★For this reason, the diagram under "Dependencies and priority" below and the line "Priority: R-4 > R-2 > ..."
      are invalid as of D-1. Treat "Open items / next-stage homework" at the end of this file as authoritative.★

### R-3. Form controls equivalent to TreeView / ListView

Directly tied to concept 2. **In modern environments where ActiveX is prohibited by default, the real demand is clear.**
The essence is not drawing but "API design for adding items, selecting, and receiving events from VBA", and for that
the **plan C-1** established with `Wv2TabBar` / `Wv2NavBar` / `Wv2SettingsBridge` can be reused as is.
**The tab bar itself is already a "custom control built with WebView2", so this is the most natural extension.**

### R-4. An MSHTML-equivalent library ★In Claude's personal opinion, the most important★

The core of concept 1. It is of a scale that would change the character of this project if done.

★**Already started as the D axis. D-1 (synchronous evaluation) / D-2 (reading) / D-3 (writing and operating + DOM condition waiting) / D-4 (settle waiting) are complete**★ The "hard part is synchronous retrieval" below was resolved in D-1.

★**In D-4, a Google Maps address search was shown to work on the real site, and obtaining coordinates from the URL was demonstrated**★

- Keep an element registry (`__wv2els[]`) on the JS side and return an **integer handle** to VBA. On the VBA side,
  the `Wv2Element` class holds the handle and builds JS on each `.InnerText` / `.Value` / `.Click` access. This
  **reduces escaping hell to "handle + property"**
- ★The hard part is synchronous retrieval★ `ExecuteScript` is asynchronous, so for `el.InnerText` to return a value
  a wait loop is needed. The existing `DoEvents` pattern of `WaitPaneReady` is the basis, but
  this is **territory that spec facts 14 (deadlock) and 20 (EBMode) hit head-on**. The place to spend the most design time
- At first, narrow it to a **minimal set of reading + clicking + input** (`GetElementById` / `QuerySelector` /
  `QuerySelectorAll` / `.InnerText` / `.Value` / `.Click` / `.SetAttribute`)

### R-5. Extensions (Authenticator, etc.)

The real demand is clear, but there is one technical wall. To enable them, at environment creation you must set
`AreBrowserExtensionsEnabled` of `ICoreWebView2EnvironmentOptions` to true and then call
`ICoreWebView2Profile7::AddBrowserExtension`. **The problem is the former: the environment options are a COM
object that is not on the "calling side" but on the "implementing and handing over side"**. For the third argument that we currently pass NULL, we must
**pass an options object whose vtable is hand-built in VBA**. It goes in the same direction as `ComCallbackHandler`, but there are
many properties to implement, and the cost is medium to large.

**★Doing R-2 would give a foothold for touching the third argument, so the order R-2 → R-5 is natural★**

### Dependencies and priority (Claude's personal opinion)

```
R-2 (Loader-independent) ──→ R-5 (extensions)     ※sharing the scaffolding for environment options
R-1 (additional APIs such as CDP) ──→ R-4 (MSHTML equivalent) ──→ R-3 (TreeView/ListView)
                                                ※R-3 can also be started with the current plan C-1 alone
```

**Priority: R-4 > R-2 > R-1 > R-3 > R-5.**
R-4 is the shortest distance to concept 1, and R-2 is a move that changes the distribution form, with a large return for its size.

---


---

## Current state (as of N-4 completion)

### Current workbook

`WebView2_v0.5.2.xlsm` (clean-built at 9.30, **801 KB / 820,491 bytes**,
**20 components**, 17,256 lines of code).
`Wv2Log.bas` was added in K-1. The size decreased because Excel repacks on save.
The D-2 replacement was 4 files: `Wv2Element.cls` (new) / `Wv2Pane.cls` / `Wv2Json.bas` / `Wv2Tests.bas`.
The D-3 replacement was 3 files: `Wv2Pane.cls` / `Wv2Element.cls` / `Wv2Tests.bas` (no new modules).
D-4 was the same 3 files + `tools/import.ps1` (read-only guard).
In D-5, **`Wv2Maps.bas` (new)** was added. The product core is unchanged.
The D-7 replacement was 2 files: `Wv2Maps.bas` / `Wv2Tests.bas` (product core unchanged).
The N-1 replacement was 4 files: `Wv2Thunks.bas` / `ComCallbackHandler.cls` / `Wv2Pane.cls` /
`Wv2Tests.bas` (no new modules). ★The first N-axis stage to touch the product core★.
The N-2 replacement was 3 files: `Wv2Thunks.bas` / `Wv2Pane.cls` / `Wv2Tests.bas`.
The N-3 replacement was the same 4 files as N-1 (including `ComCallbackHandler.cls`).
N-4 was also the same 4 files. ★The N axis is completed with just the 4 files of the product core★.
★Only the spelling of labels changed in `Wv2Thunks.bas`★ (spec facts 38 / 66).

※ Measured and corrected in E-1. The earlier "683 KB, 18 modules" did not reflect the addition of `Wv2Element`
   in D-2. The breakdown is 6 standard + 9 class + 2 form + 2 document = 19.

### Repository (moved out of OneDrive in E-4)

`C:\Users\gugug\GitHub\WebView2-For-Excel-VBA`
(GitHub: `tarboh/WebView2-For-Excel-VBA`)

★**Do not put it back under OneDrive**★ Excel opens the workbook with a cloud URL and
AutoSave defaults to ON (spec fact 48). In E-4 this actually happened: the workbook we had fixed was
overwritten by the AutoSave of an Excel that had grabbed an old copy.

| Location | Contents |
|---|---|
| `src/` | **★Authoritative★** 8 standard + 9 class. Edit here |
| `src/document/` | `ThisWorkbook` / `Sheet1` (for reference. Cannot be replaced by Import) |
| `forms/` | `UserForm1` / `TimerForm` (manual operation) |
| `book/` | The current workbook. Commit only at milestones |
| `docs/` | The 4-volume development notes + `standard-thunk-method.md` + `handover-at-D2.ja.md` |
| `tools/` | `export.ps1` / `import.ps1` / `repair_book.ps1` / `check_cp932.py` / `check_vba.py` / `check_idents.py` / `check_book.py` |
| `Old/` | Past workbooks, modules, and old development notes. **Do not touch** |
| `CLAUDE.md` | Standing instructions that Claude Code reads automatically every session |

- Workbook → `src/`: `powershell -ExecutionPolicy Bypass -File tools/export.ps1 -Force`
- `src/` → workbook: `powershell -ExecutionPolicy Bypass -File tools/import.ps1`
  (★The default is preview. To write back, use `-Apply`★ / pre-gate + backup + diff only + post-verification)
- **When adding a new module, use `-AllowNew`** (E-3). Hand-building in the VBE is no longer necessary.
- **The round trip was closed in E-2, and closed including new additions in E-3.** Idempotence and convergence are both confirmed.
- ★After `-Apply`, "a module we did not touch can change"★ (spec fact 43).
  `import.ps1` reports it, so take it into `src/` with `export.ps1 -Force`.
- ★**After `-Apply`, always run `python tools/check_book.py`**★ (spec fact 47).
  There are anomalies invisible via `CodeModule`, and in E-4 we actually produced a compile error.

### Standard file layout (13 files, 12 in production)

```
Wv2Thunks.bas / Wv2Json.bas / Wv2Url.bas / Wv2Maps.bas / ComCallbackHandler.cls
Wv2Environment.cls / Wv2Pane.cls / Wv2Browser.cls / Wv2Element.cls
Wv2TabBar.cls / Wv2NavBar.cls / Wv2SettingsBridge.cls / UserForm1.frm
Wv2Tests.bas ← for verification only. Can be removed when porting, so 12 in production
(auxiliary: SafeTimer.cls / TimerForm.frm / LibTimers.bas / Module1.bas)
```

Last-updated stage of each file: Pane / Tests = **D-4**, Element = D-4 (1 line), Json = D-2, TabBar = 9.32b,
Browser / SettingsBridge = 9.29, NavBar = 9.25b, UserForm1 = 9.18, Url = 9.16,
Thunks = 9.16, Handler = 9.9b.
※ The D-1-completion entries (9.3a / 9.4a) for Thunks and Handler were wrong. We checked the
   actual file headers and corrected them (corrected in D-2).

### Features verified on the real machine

- Adding, switching, and deleting tabs, **drag-and-drop reordering**, hover / close-button display control,
  improved perceived speed via optimistic update, wheel horizontal scrolling and keeping the scroll position
- URL input, back, forward, reload, Escape, and independent history per tab
- **Opening links in a new tab. ★POST links open correctly in our own tab "with no configuration"★**
- An address bar equivalent to Chrome/Edge. Search engine selection, custom URL, persistence
- **D-1: Synchronous evaluation of JS via EvalSync**
- **D-2: Element retrieval and reading (`QuerySelector` / `Wv2Element`), generation check across page navigations**
- **D-3: Writing and operating (`.Value =` / `.Click` / `.SetAttribute`).
  ★It bypasses the value tracker of React / Vue, so values propagate even on SPAs★**
- **D-3: DOM condition waiting (`WaitFor` / `WaitGone` / `ImplicitWaitSec`).
  It gives up without waiting as soon as it knows it is a failure**
- **D-4: Settle waiting (`WaitSettled`). The in-page probe counts `fetch` / `XHR` / DOM changes, and
  ★if the health check finds it broken, it rebuilds it★. Explicit signals (`ArmContentSignal` / `ArmNetworkSignal` /
  `ArmUrlSignal`) / noise exclusion / cutting off lingering requests / diagnostic log**
- ★**A Google Maps address search works on the real site**★ It has been demonstrated through writing an address, pressing search,
  waiting for completion, and taking the coordinates (`/@35.6824246,139.7682844`) from the URL
- **D-5: `Wv2Maps.MapsGeocode` obtains "address → coordinates + normalized address" in one line.**
  ★3 items can be processed consecutively in the 2-second range★ (the same tab is reused)
- All of it is just Excel VBA with zero external libraries

### Key points of the architecture (unchanged since 9.15)

- **Loose coupling**: `Wv2Browser` fires 5 kinds of Public Event (`TabAdded` / `TabClosed` / `ActiveChanged` /
  `TabTitleChanged` / `TabUrlChanged`), and `UserForm` / `Wv2TabBar` / `Wv2NavBar` subscribe with
  `WithEvents`. **It stays at 5 kinds.**
- **The tab bar / nav bar use WebView2 as a UI component** (one independent WebView2 each).
- **VBA→JS is the full-sync approach** (design principle 60). However, by design principle 84,
  an "advance display that the next sync will certainly overwrite" is tolerated.
- **JS→VBA (hostObjects path = normal operation)**: `Settings.*` / `TabBar.Host*` / `NavBar.Host*`.
  All are plan 4-A (no waiting for completion).
- **Established pattern (plan C-1)**: "JS is a hybrid, VBA consolidates processing in one Host method,
  and the Select Case is a thin adapter".

### Verification resources

| Sub | Target |
|---|---|
| **`Test_K1_Log` / `Test_K1_Help`** | **Wv2Log (K-1). Run, 14/14 passed. Works even with the browser not started** |
| **`Test_K2_Help`** | **The settings tab bug (K-2). Run, all 15 steps normal** |
| **`Test_D4_Probe` / `Test_D4_Settle` / `Test_D4_Signal` / `Test_D4_Log`** | **Settle waiting (D-4).
  Run, all of 28 / 22 / 31 / 21 items passed. Procedure in `Test_D4_Help`** |
| **`Test_D4_Site`** | **★Google Maps reconnaissance★ The only Test_* that actually accesses an external site.
  To change the address, `Test_D4_Site "another address"`** |
| **`Test_D4_Dom`** | **★Enumerates what is on the currently open tab★ A tool for deciding selectors** |
| **`Test_D5_Geocode`** | **Address → coordinates (D-5). Run, all 14 items passed. Actually accesses an external site** |
| **`Test_D5_Sheet`** | **Worksheet integration (D-5b). Run, all 7 items passed. ★Create a new workbook and try it★** |
| **`Test_D6_All`** | **`QuerySelectorAll` and lifetime management (D-6). Run, all 21 items passed** |
| **`Test_D6_Pick`** | **Choosing from the candidate list (D-6b). Run, all 10 items passed. Actually accesses an external site** |
| **`Test_D7_Sheet`** | **Progress and cancellation (D-7). Run, all 13 items passed. ★Press Esc once around the 2nd item★
  Actually accesses an external site. ★Do not mash it (it also takes effect on resume)★** |
| **`Test_D7_Cancel` / `Test_D7_Resume`** | **The cancellation entry point and denominator (15 items) / resume after cancellation (8 items).
  Resume is ★no Esc needed, fully automatic★ so it can be used for regression checks** |
| **`Test_D7_Key` / `Test_D7_StatusBar`** | **★Probes★ Whether Esc is visible to the API / how to restore the status bar.
  Neither needs the network. The purpose is to ★produce facts★ rather than verdicts (design principle 103)** |
| **`Test_D3_Write` / `Test_D3_Framework` / `Test_D3_Wait`** | **Writing, operating and waiting (D-3).
  Run, all of 33 / 7 / 29 items passed. Procedure in `Test_D3_Help`** |
| **`Test_D3_Probe_Promise`** | **Measurement of unknown 1 (D-3). Settled, so normally not needed** |
| `Test_D1_Eval` / `Test_D1_Guard` / `Test_D1_Help` | EvalSync (D-1). Run, passed |
| `Test_D2_Find` / `Test_D2_Stale` / `Test_D2_Help` | Reading elements (D-2). Run, passed |
| `Test_9_32b_Help` / `Test_9_32_Help` / `Test_9_31_Help` | Around the tab bar's perceived feel |
| `Test_9_29_Custom` / `_Help` | Custom search engine |
| `Test_9_28_Help` | Prewarm delegation. Also describes how to isolate the problem when you hit a 503 |
| `Test_9_27_Status` | Dumps the spare tab's state on 1 line |
| `Test_9_27_Mode_Legacy` / `_Popup` / `_Prewarm` | Mode switching (the default is Prewarm) |
| `Test_9_26_PostProbe` / `_Help` | POST probe page (5 buttons). Passed again in the D-1 regression check |
| `Test_9_20_Persistence` / `_Help` | Persistence (2 keys supported as of stage 9.29) |

- ★**The D-1 / D-2 / D-3 tests emit their verdicts to `Wv2Log`**★ (migrated in D-3). The Immediate window
  gets 15 lines of plumbing log per `EvalSync` round trip and scrolls away, so **check pass/fail in the log file**.
  At the end of each test, `★Verdict n items: OK x / FAIL y★` is printed. The location is `?Wv2Log.LogPath`.
- ★**Prepare verification from `UserForm1.Show vbModeless`**★ (spec fact 54). Without Show,
  `hWnd_Frame` stays 0 and `StartWebView2_Full` fails.
- Settings save location: `%APPDATA%\Wv2Browser\settings.txt`. Delete it by hand to return to the default (Google).
- ★External services go down★ If a 503 appears, first type the destination URL directly into the address bar and check whether a 405
  comes back. If it is 503 the server is down, so switch to the spare of button 4 / button 5.
- ★Static analysis of the workbook★ `pip install oletools --break-system-packages` →
  `python3 -m oletools.olevba <book>.xlsm --json`.

### Operating rules (development cycle)

① Align on the approach (present multiple options + state Claude's recommendation) → tarboh chooses →
② Implement → ③ Share real-machine verification logs → ④ Record in the development notes → next stage.

- **Never start implementation before agreeing on the approach.** Keep each stage to a small scope that can be verified independently on the real machine.
- The practice of **agreeing on the issues all at once** has taken hold (stage 9.27 had 7 items, stages 9.29-9.31 had 8,
  stage 9.32 had 7, stage 9.32b had 3, D-1 had 8, D-2 had 9, D-3 had 8).
- **Always report defensive implementations and minor adjustments to the approach after the fact** (design principle 79).
- For verification, Claude prepares `Test_XXX` / `Test_XXX_Help` so that tarboh only has to
  **type the Sub name in the Immediate window**.
- When context gets tight, proactively propose a handover to a new chat before the limit.

### Incident notes (preventing recurrence)

| Stage | Incident | Fastest diagnosis |
|---|---|---|
| 9.24 | The contents of `Wv2NavBar` had become the code of `Wv2TabBar` | Check in the startup log that each class prints its own name as a prefix |
| 9.30 | `Wv2Json.bas` was missing entirely in the D&D port | **Count the modules** (v0_5_2 has **19**. Corrected in E-1) |
| 9.32 | JS `HO()` was declared inside a for loop in `render()` and happened to work thanks to hoisting | Check that function declarations are not inside blocks |
| D-1 | `Dim eNum As Long` was a compile error (`eNum` = reserved word `Enum`) | Make the reserved word check permanent (design principle 90) |
| E-1 | The `.ps1` turned into a pile of syntax errors | Saved without BOM, PowerShell 5.1 read it as CP932. **If the errors concentrate on Japanese comment lines, suspect the character encoding** |
| E-1 | `$raw` collided with the switch parameter `$Raw`, giving a type error | **PowerShell also does not distinguish identifier case.** The same type as VBA's `eNum` / `Enum` |
| E-1 | `git push` was rejected | The remote had commits we had not fetched. **Never force push.** Fetch, compare the contents, then merge |
| E-1 | Diff detection timed out | `difflib.SequenceMatcher` was applied character by character to a 170 KB string. For large files, compare **line by line** |
| K-1 | The surrogate check was about to never hold | Without `&` on `&HD800` it becomes a negative Integer (design principle 29). `check_vba.py` detects it |
| K-1 | The verification failed twice, and both times the cause was on the verification Sub's side | **Directly check the failed item through a path separate from the target** (when I read the log file with Python, the contents were perfect) |
| E-4 | The workbook got a compile error (attribute lines exposed) | `import.ps1` had doubled the attribute lines. ★Visible neither from `CodeModule` nor from `Export`★ Check with `python tools/check_book.py` |
| E-4 | The workbook we had fixed reverted | OneDrive's AutoSave. ★Do not place the workbook under OneDrive★ |
| D-3 | The probe was not measuring unknown 1 | `EvalSync` runs `JSON.stringify` first, so the Promise was crushed midway along the path. **Check the probe itself** (design principle 103) |
| D-3 | 40 verdicts vanished from the Immediate window | They scrolled away under 600 lines of plumbing log. **Verdicts go to the log file** (design principle 105) |
| D-4 | The code we thought we had written back was not in the workbook | Excel was open and `Workbooks.Open` succeeded as read-only (spec fact 61). Added a guard to `import.ps1` |
| D-4 | The real site's selector was not found even after waiting 30 seconds | `#searchboxinput` had changed to `id=ucc-1`. **Do not fix by guessing; build a tool to look into the DOM** (design principle 107) |
| D-4 | Emptying the exclusions had no effect | "Because it was empty, we did not send it", so the previous setting remained on the page side (design principle 106) |
| D-5 | From the 2nd item on, the coordinates were those of the previous one | "Done when the URL changes" was too early (spec fact 62). **It does not show up with a single-item verification** (design principle 110) |
| D-5b | A bogus address got the previous row's coordinates | When the search fails the map does not move, and `/@` leaves a leftover. **If it cannot be judged, do not write it** (design principle 111) |
| D-6b | Running the same search twice did not enter the candidate selection path | The verdict used an indirect indicator, "did the map move". **Count directly** (design principle 112) |
| D-7b | The added `Cleanup:` rewrote `Wv2Thunks` | It collided with the existing label `CleanUp:` differing only in case (spec fact 38). ★`check_idents.py` does not look at labels★ (spec fact 66) |
| D-7d | Cancellation worked but was reported as "not working" | The verification Sub ran a 2nd run right after cancellation, so the moment of stopping was not visible (design principle 114) |
| D-7d | Misdiagnosed as "it can be called 841 times a second, so it is enough" | It should have been judged by ★coverage of time★, not by count. It was actually 2% (design principle 113) |
| N-1 | Back-calculating the vtable index from existing wiring was off by one | `OpenDevToolsWindow` was missed in the count. ★Pull it from the SDK header with `tools/vtable.py`★ (design principle 115) |

---

## Open items / homework for the next stage (as of N-4 completion)

D-1 to D-7, E-1 to E-4, K-1 to K-5, and N-1 to N-4 are complete.
**The D axis now has "read, write, wait, stop" in place, with one practical use (address → coordinates) on top of it.**
**The N axis has reached the point where ★a whole communication, from request to response body, can be obtained★.**
There is no work in progress. ★Next, N-5 (PowerShell code generation) is the main candidate★.
Besides that, item 0 below (slow release when the form is closed).

0. **★The remainder of K-4: release does not become "instant"★** (improved from 5 seconds to 2 seconds in K-4c)
   `Wv2Browser.Shutdown` is called from `UserForm_Terminate` = ★the form has already
   been closed by then★, so even the 1.6 seconds of quiescing happens after the × is pressed. To make it "instant" we would have to
   ★fire it first in `UserForm_QueryClose`★ = touch `forms/` (manual operation).
   ★The cause itself is still unidentified★ (9 losses in a row). The starting point for the next dig is
   ★"only the non-active Pane is slow" (spec fact 68)★.

0b. **★Operating rule from K-5★** ― ★In an Excel that has launched WebView2, typing Japanese
   into the Immediate window with an IME can crash it (spec fact 67)★. **It does not recover even after closing.**
   ★Use the alphanumeric aliases `Boot` / `BootNoQuiet` / `BootMaps` in `Module1`★.
   When Japanese is needed, copy and paste. The reproduction means are `CrashTest1` to `12`.

1. **★Putting it into real work, and the requests that come out of it★**
   `Wv2Maps.MapsGeocodeSheet` has gone all the way through to worksheet integration (D-5b), and ★progress display and
   cancellation have also been added (D-7)★. Next, **the requests that come out of running it on a real address book** take priority.
   ★The measured time is 8 seconds per item (an address that resolves uniquely) / 27 seconds (multiple candidates)★.
   The earlier "1 to 2 seconds per item" written here was wrong. Expect around 15 minutes for 100 items.

1b. **★Esc does not work with a quick "tap" (a long press is needed)★** (remainder of D-7)
   Of the 35.5-second batch, only ★about 2%★ is spent watching for cancellation (design principle 113). The remaining 98% is
   waiting on the `Wv2Pane` side (`EvalSync` / `WaitFor` / `WaitSettled`), which has no detection point.
   ★The 10 to 15 seconds of `MapsOpen` also do not stop for the same reason★ (measured: `cancel detected 2 times / 10.5 seconds`).
   To close this, **wire a detection point into the wait loops of `Wv2Pane` = touch the product core (D-8)**.
   The issue is "where to put the cancel flag" (create a new cross-cutting module, or a public property).

2. **★`WaitVisible` (wait until visible)★** ― In D-3 issue 5 we settled on "look only at existence", and
   in D-6 issue 1 it was also sent to a separate stage. We have to start by
   deciding the definition of visible (`offsetParent` / area / `visibility`). When a real need appears.

3. **★Shadow DOM and iframes are not reached (a hole in the D axis)★**
   `document.querySelector` does not look inside a shadow root, and cannot cross iframes.
   ★The reference product (WebDriver BiDi for SeleniumVBA by sele_chan) has reached the conclusion of
   "do not add Shadow-specific methods; make the existing methods support it"★ (shared on 2026-08-22). We too should
   make the grabbing side (`QuerySelector`) penetrate. **Directly tied to concept 1.**

4. **★The completion of canvas drawing cannot be observed★** (spec fact 56)
   This matters for uses such as saving a map as an image. **Neither printing nor screen capture is implemented**.
   `CapturePreviewAsync` (one completion handler + IStream) is the direct hit, and
   ★completion of drawing can be judged from outside by "take it twice with a short gap; if identical, it has stopped"★.

5. **Routes (travel time and distance between 2 points)** ― More complex than an address search (two input fields, choosing the travel mode).
   After address search has solidified in real work.

6. **K-4: Migrating the remaining `Debug.Print`s** ― There are still over 1,600 in `src/`. Do not do it all at once.

7. **The Bkav Pro false positive (continuing from stage 9.30)** ― Reported, awaiting a reply.
   ★The reference product has achieved 0 detections on VirusTotal (including Bkav)★, so a
   VBA .xlsm can get a clean verdict too.

8. **Roadmap** ― R-4 (MSHTML equivalent) is complete through D-5 as the D axis.
   R-3 (TreeView / ListView) can be started at any time. R-1 / R-5 / R-2 are postponed.

9. **Permanent fix of the Sentinel MEM_FREE bug (spec fact 10) ― the main job of the K axis**
   `Wv2Thunks.bas` = the foundation of all features. An area where the debugger cannot be used, per spec fact 20.

10. **The remainder of the E axis** ― manual operation of `forms/` / deleting the old folder on the OneDrive side /
    ★making the JS verification harness permanent★ (cannot be used while Edge is running; standard output goes through Start-Process) /
    ★`check_idents.py` **does not look at labels**★ (in D-7b `Cleanup:` collided with the existing
    `CleanUp:`, and by spec fact 38 the spelling across the whole project was unified).

11. **Small homework items** ― gap animation while dragging / optimistic update for close /
    making it 2 spare tabs / removing `Wv2Pane.IsSettings`. Low priority.

12. **★Support for 32-bit Excel★** (specified by tarboh on 2026-08-23)
    ★Strengthening the machine-code thunks is required★ Currently it is **explicitly 64-bit only**, and
    the 4 files `Wv2Thunks` / `Wv2Pane` / `Wv2Environment` / `ComCallbackHandler` reject 32-bit at compile time with
    `#If Win64 ... #Else #Error`.
    ★The groundwork exists★ Since the 9-series stages we have adopted the policy of leaving "32-bit-ization point" comments,
    and 3 places are currently marked.
    The issues will be ★the calling convention (x64 has 1 kind, x86 has stdcall / thiscall)★ /
    ★rebuilding the thunk byte sequences to branch by width★ /
    the use of `LongPtr` versus `LongLong` / vtable offsets going from 8 to 4 bytes.
    **Putting it on the K axis (foundation) is the natural choice.** Whether to create a new volume will be discussed when work starts.

13. **★Extracting the observation log★** (from sele_chan's remark on 2026-08-23)
    `SpaProbeDrainLog` emits things raw. The other side ★extracts only what is
    needed on the premise that an AI will read it★. What to keep in order to hit the completion condition can only be
    decided on real sites. For details see the end of `dev-notes-3-D-axis.md`.

14. **★N-5: PowerShell code generation (the main candidate for the N axis)★** ― With N-4 the ingredients are in place.
    URL / method / request headers / request body / status / response headers / response body.
    ★Only once we get this far does it become value that F12 does not have★
    (the real-world flow "operate once by hand in the browser → receive that communication as PowerShell code →
    put it on Task Scheduler" becomes a tool as is).
    Issues likely to come up when work starts:
    - ★Handling of secrets★ We introduced a masking mechanism for Cookie / Authorization in N-2, but
      ★code generation has to put them into "working code"★. Whether to emit them masked,
      use placeholders, or make the user explicitly unmask
    - ★Spec fact 74★ Japanese in a form POST is carried percent-encoded.
      To make code that can be pasted as is, decide between decoding it and passing it raw
    - ★Spec fact 79★ The response can be obtained already decoded, so there is no need to reproduce
      `content-encoding` (`Invoke-WebRequest` also decompresses automatically)
    - The shape of the output (print to the Immediate window / write to a sheet / emit to a file)

15. **The `FETCH`(8) filter may be unused** (spec fact 70)
    `fetch()` arrives as `XML_HTTP_REQUEST`(7). It is unconfirmed whether it is still caught without installing `FETCH`(8).
    ★Installing it does no harm, so leave it as is for now★. Minor.

16. **`multipart/form-data` (file upload) is treated as binary** (N-2)
    Whether to decode the contents separated by the boundary string depends on N-5's needs.
    ★For now it remains as "N bytes of binary + hex of the first 32 bytes"★, so
    it is not that nothing is visible.

17. **The response of `fetch(..., {redirect:'manual'})` cannot be observed** (N-3)
    Using this cross-origin appears not to show up in `WebResourceResponseReceived`
    (in the N-3 verification no response came, both times). ★This is an observation, not an assertion★.
    In N-3 only the verdict on the `Set-Cookie` overlay did not run; the masking mechanism itself
    was verified via another path. Investigate if needed from N-4 onward.

18. **Concurrent requests to the same URL get their responses mixed up** (N-3)
    The matching key is only URI + method (the args carry no clue such as
    NavigationId). ★Mixing them up cannot be detected★. We made it FIFO, so as long as they
    return in order it is correct, but it remains as a hole in principle.
