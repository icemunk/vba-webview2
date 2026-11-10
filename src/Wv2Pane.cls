VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Wv2Pane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  English fork v1.1 (LastNavigationOk) ---
'
'   Added LastNavigationOk so that callers (Wv2.WaitLoaded) can tell a finished
'   load from a failed one. No other behaviour change.
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  Stage N-4 (response body) ---
'
'   *Built differently from N-1 to N-3. This is the first asynchronous one*
'   The full details are collected in the N-4 section (at the end of this file). Key points only:
'
'     *`GetContent` takes a handler and returns immediately. The body arrives later
'     **Unlike N-2, no Seek(0) is needed* we only read a copy of the response
'     **The body has its own switch* (`NetRespBodyOn`, default False)
'     **Do not assume compression; observe it* always keep the encoding and the hex of the first 32 bytes
'     **ResponseView is AddRef'd and held until completion* (it is unclear whether releasing it is safe)
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  Stage N-3 (response status and headers) ---
'
'   Each row of the list now carries a *result*. The detail now carries the *response headers*.
'   The full details are collected in the N-3 section (at the end of this file). Key points only:
'
'     **This event has no filter* images and CSS all arrive too. So we
'       record *only those that match a request already in the list* (the narrowing is not kept in two places)
'     *Matching is by URI + method, against "the oldest request that has no response yet" (FIFO).
'       *Concurrent requests to the same URL get mixed up, and the mix-up cannot be detected*
'     **Rows whose response has not arrived yet stay blank* (writing 0 would be indistinguishable from a real one)
'     **GetIterator on the response side is 7, on the request side 8* reusing one shifts by 1
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  Stage N-2 (request headers and body) ---
'
'   On top of the N-1 list (URL / method / kind), add *what is needed to reproduce a single request*.
'   The full details are collected in the N-2 section (at the end of this file). Key points only:
'
'     **The detail is entirely OFF by default* (ON with NetDetailOn = True)
'     **Always rewind a read IStream with Seek(0)* otherwise an empty body goes out
'     **Cookie / Authorization are hidden by default* (raw on the ring)
'     *The detail uses a ring separate from the list (default 50 entries). The list stays at 500 entries
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  Stage N-1b (reflecting what was learned on the real machine) ---
'
'   # *Spec fact: virtual host requests do not ride on WebResourceRequested*
'     Requests served via SetVirtualHostNameToFolderMapping (all of DOCUMENT / fetch / XHR /
'     images / CSS) do not ride on the event, because the runtime
'     answers directly with an internal resource handler. *This was the cause of the N-1 first verification
'     failing completely*, and the wiring side was correct from the start (on a real site all 10 were caught).
'     The targets of the verification page are "an unreachable local address" and "an external service".
'
'   # *NetLogDrain does not reset the cumulative counters to 0 (issue 3, option Y)*
'     NetLogTotal / NetLogDropped are documented as cumulative "*since the capture started*", but
'     Drain rebuilt the whole ring and reset them to 0. As a result the last line of the real-machine log was
'       NetCaptureStop: capture ended (total fired 0 ...)
'     which left a line that could be read as *"none came" even though 10 had come*.
'     *NetLogDrain ... empties only the ring contents (NetRingEmpty). The cumulative counts are kept
'     *NetLogClear ... resets everything including the cumulative counts (NetRingAlloc)
'     *NetCaptureStart ... goes through NetLogClear, so it counts from 0 each time
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  Stage N-1 (wiring up WebResourceRequested) ---
'
'   *See what the browser is actually hitting, using only Excel VBA*
'   The first step toward the equivalent of the F12 Network tab. N-1 covers only URL / method / kind
'   (headers, body and response come in later stages).
'
'   # Added Public members
'     NetCaptureStart(contexts, uriFilter) As Boolean  start capturing
'     NetCaptureStop()  As Boolean                     stop capturing
'     NetCaptureOn      As Boolean                     whether capturing
'     NetLogDrain()     As Long     flush what has accumulated to Wv2Log and empty the ring
'     NetLogLine(i)     As String   *read a single entry without flushing* (oldest first, 1-based)
'     NetLogCount / NetLogTotal / NetLogDropped As Long
'     NetLogCapacity    As Long     ring capacity (Get/Let)
'     NetLogClear()                 empty the ring
'     View_AddWebResourceRequested / View_RemoveWebResourceRequested
'     View_AddWebResourceRequestedFilter / View_RemoveWebResourceRequestedFilter
'     View_OnWebResourceRequested   (entry point from ComCallbackHandler)
'
'   # *Record only, then return immediately (issue 5)*
'     WebResourceRequested is an event for "stopping / rewriting a request", so
'     *a slow handler jams the traffic itself*. Therefore inside the handler
'     *fire neither Wv2Log nor Debug.Print
'     *of course never call EvalSync (spec facts 14 / 20)
'     just read URL / method / kind from args, put them on the ring, and return.
'
'   # *Records live in an in-memory ring buffer (issue 6)*
'     One page produces tens to hundreds of requests. Writing straight to the log would overflow it, so
'     we used the same pattern as D-4's SpaProbeDrainLog (accumulate -> take out in bulk).
'     *Overflowed entries are not silently discarded* they are visible via NetLogDropped (same judgement as D-6 issue 3).
'
'   # *If no filter is registered, not a single event arrives*
'     Fire AddWebResourceRequestedFilter(uri, ctx) per kind. The default is
'     the three of DOCUMENT(1) / XML_HTTP_REQUEST(7) / FETCH(8) (issue 4).
'     *Remember the registered filters and remove them in NetCaptureStop* (in the spirit of design principle 106.
'     If the previous setting is left on the other side, the next Start unknowingly becomes a wide net).
'
'   # *The vtable index was settled from the SDK header (design principle 115)*
'     Working backwards from the existing wiring, add_WebResourceRequested had been misread as 54
'     (correct is 55; OpenDevToolsWindow had been missed in the count). The tool is tools/vtable.py.
'       ICoreWebView2::add_WebResourceRequested              55
'       ICoreWebView2::remove_WebResourceRequested           56
'       ICoreWebView2::AddWebResourceRequestedFilter         57
'       ICoreWebView2::RemoveWebResourceRequestedFilter      58
'       ...WebResourceRequestedEventArgs::get_Request         3
'       ...WebResourceRequestedEventArgs::get_ResourceContext 7
'       ICoreWebView2WebResourceRequest::get_Uri              3
'       ICoreWebView2WebResourceRequest::get_Method           5
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  Stage K-4a (move the release-path log to Wv2Log) ---
'
'   *The purpose is measurement* We received a report that release when closing the form is slow (a pause of about 0.5 s per item),
'   but this path was still Debug.Print and left nothing in the log, so
'   it was unclear where the time was going.
'
'   Wv2Log has millisecond timestamps, so *just replacing it lets us read the elapsed time*.
'   Furthermore, measuring twice with Wv2Log.LogEcho = False / True separates
'   *"COM release is slow" from "output to the Immediate window is slow"*.
'
'   *This stage does not fix anything. It only measures.* (because guesses missed three times in a row in D-7)
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  Stage D-6 (QuerySelectorAll and lifetime management) ---
'
'   # Added Public members
'     QuerySelectorAll(selector, maxCount) As Collection
'     LastAllTruncated  As Boolean  whether the previous QuerySelectorAll was truncated
'     ElementCount      As Long     number of elements stacked in the JS-side registry
'
'   # *Multiple results are returned as a Collection (issue 2)*
'     You can write For Each el In els, which is the most natural for VBA. Zero results is simply Count = 0.
'     *Why not an array* UBound of an empty array is a runtime error, forcing the caller to
'     guard every time. *Why not a custom class* To make For Each work you need
'     a member attribute line (Attribute Item.VB_UserMemId = 0), but import.ps1 drops
'     attribute lines by design (spec fact 47), so it does not go well with round trips.
'
'   # *Truncation is not done silently (issue 3)*
'     Default 200 entries. If exceeded, set LastAllTruncated = True and also emit a warning log.
'     Do not let anyone assume "the returned count is the real count on the page".
'
'   # *Lifetime management only "shows". It does not discard on its own (issue 4)*
'     QuerySelectorAll can stack 200 at once, so the registry grows faster
'     than D-2 assumed. But if old ones are discarded on our own, *the Wv2Element in hand silently
'     becomes stale* (the judgement of design principle 91 / D-2 issue 3 applies as is).
'     So:
'       *show the current count with ElementCount
'       *notify with LogW once it exceeds a certain number
'       *discard only when ClearElementRegistry is called (present since D-2)
'
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  Stage D-4b (waiting for quiescence) ---
'
'   Use the probe placed in D-4a to *wait until things settle*.
'
'   # Added Public members
'     WaitSettled(timeoutSec, stableMs) As Boolean  wait until quiescent
'     AutoWaitAfterAction  As Boolean  *default False* wait automatically after an action (opt-in)
'     AddIgnoreNetwork(pattern)        substring pattern of URLs to exclude from the quiescence check
'     ClearIgnoreNetwork()             empty the above
'     IgnoreSelectors  As String       selectors of DOM changes to exclude from the quiescence check
'     LastSettleInfo   As String       breakdown of the previous wait (for diagnosis)
'
'   # Quiescence formula (issue 4)
'     Quiet when running fetch + XHR is 0 and at least stableMs has passed since the last DOM change.
'     *However the quiet window is measured from "the moment WaitSettled is called"* even if it was quiet
'     before the call, it does not succeed immediately. Hence *it waits at least stableMs*.
'     Defaults are stableMs = 500 / limit 10 s / polling 100 ms (following a reference product).
'
'   *** Even if WaitSettled returns True, it is not proof that "the app has finished processing" ***
'     It is merely quiet. **It cannot tell whether the app is quiet because it ignored the request
'     or because it finished.** For important operations, combine it with D-4c's explicit signal (arm).
'     *This is why the name is WaitSettled and not WaitIdle* (issue 8)
'     If "idle = nothing is happening" can be read into it, it is mistaken for a guarantee.
'
'   # Return value (follows the D-2 / D-3 convention, design principle 93)
'     True                     ... reached quiescence
'     False + LastEvalOk=True  ... did not quiet down in time (*a normal result*)
'     False + LastEvalOk=False ... failure (in-callback / already Closed, etc.)
'     In either case the breakdown goes into LastSettleInfo (slack / traffic in flight / elapsed).
'
'   # Noise exclusion (issue 5)
'     *The exclusion list is held on the VBA side and pushed to the page on every round trip*
'     because holding it on the page side would lose it on navigation. The push rides on the same
'     single round trip as reading the state, so round trips do not increase.
'     *An exclusion that is too broad produces "too-early quiescence"* the reference product warns about this too.
'     Blocking (stopping the request itself) is not done, because it would change the page's behaviour.
'
'   # *Traffic that started before the probe was placed cannot be counted* (a D-4a specification)
'     This does no harm in the operation of acting right before the wait (D-4c's arm-then-act).
'     To wait for load completion right after opening a page, use the existing
'     NavigationCompleted / DOMContentLoaded.
'
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  Stage D-4a (in-page SPA probe and health check) ---
'
'   The purpose of D-4 is *waiting for quiescence* (wait until traffic and DOM changes settle).
'   D-3's WaitFor / WaitGone are the means when you can write "what to wait for", and
'   **cannot be used when the target of the wait cannot be identified**. D-4 fills that gap.
'
'   D-4a is its foundation. *It goes only as far as placing the observing mechanism (probe) in the page*, and
'   the body of the wait (WaitSettled) and the explicit signal (Arm family) are D-4b / D-4c.
'
'   # Added Public members
'     SpaProbeState()  As String   returns the state JSON after getting the probe into shape
'     SpaProbeReset()  As Boolean  discards the probe and rebuilds it (for diagnosis)
'     SpaProbeHealthy() As Boolean a thin entry that only reads h from the state JSON
'
'   # What the probe counts (window.__wv2p)
'     fetchN / xhrN ... the number of fetch / XHR calls *currently running* (in-flight)
'     lastMut       ... time of the last DOM change counted. The quiet window is measured from here
'     netN / mutN   ... total counts (for diagnosis)
'     skipNet / skipMut ... number excluded (to confirm the noise exclusion is working)
'
'   # *Extension of design principle 92: from "create if missing" to "rebuild if broken"*
'     We wrap fetch and XMLHttpRequest.prototype.open/send in order to count them, but
'     *the page's framework may replace these later*
'     (the reference product WebDriver BiDi for SeleniumVBA implements a health check
'      = it really does break). So every time we check
'       *whether the version number matches
'       *whether window.fetch is still our wrapper
'       *whether XHR open/send are still our wrappers
'       *whether the MutationObserver is alive
'     and if even one is missing we *rebuild*. The number of rebuilds can be read from repairs.
'
'     *When rebuilding, we wrap "the current window.fetch"*
'     Since it wraps the wrapper the framework layered on, the framework's features are not broken.
'     There is also a check so that our own wrapper is not wrapped twice (preventing missed / double counting).
'
'   # *When to set it up (issue 2, option A: extending the self-generating prologue)*
'     AddScriptToExecuteOnDocumentCreated is not used. As in D-2, we use the form
'     "the expression that uses it prepares it itself". Therefore
'     *traffic that started before the probe was placed cannot be counted*
'     This is by specification. Combined with D-4c's practice of "arm right before acting" it does no harm
'     (the probe is set up at the moment of arming, so it always precedes the action).
'
'   # Where the probe is created
'     *It is created only in a Pane that called these members* The Panes of the tab bar / nav bar /
'     settings tab do not call them, so they are untouched (same reason injection was avoided in D-2).
'
''''''''''''''''''''''''''''''''''' --- Wv2Pane.cls  Stage D-3 (writing to and operating on elements) ---
'
'   What D-3 added to Wv2Pane is *just one member, Elem_Exec (Friend)*.
'   The entry points for writing and operating are on the Wv2Element side (Value = / Click / SetAttribute), and
'   Wv2Pane provides the "firing side" path that pairs with D-2's Elem_Read.
'
'     Elem_Read (D-2) ... *reads* one value   takes a JS *expression* that uses variable e
'     Elem_Exec (D-3) ... *fires* one operation   takes a JS *statement* that uses variable e
'
'   # How Elem_Exec returns from the JS side (aligned with the same shape as Elem_Read)
'     {s:'gen', g:...}            generation mismatch = stale
'     {s:'nil', g:...}            no element at the index = treated as stale
'     {s:'ok',  g:..., v:'<rv>'}  success. v is *the name of the path the operation chose*
'
'     *JS exceptions are not caught with try-catch* View_ExecuteScriptSafe already wraps it, so
'     the exception comes up as EvalSync's LastEvalError ("TypeError: ..." etc.).
'     Wrapping it twice would bury the reason.
'
'   # Why the path name (v) is returned
'     Whether the .Value = write went *through the native setter ('setter') or a plain assignment
'     ('direct')* is the only clue to whether it took effect on an SPA.
'     It can be read via Wv2Element.LastInfo (D-3's post-hoc report).
'
'   # DOM condition waits (D-3b, issue 2, option C: *explicit waits first-class, implicit waits opt-in*)
'
'     WaitFor(selector, timeoutSec)  As Wv2Element  wait until it appears
'     WaitGone(selector, timeoutSec) As Boolean     wait until it disappears
'     ImplicitWaitSec                As Single      *default 0 = disabled*
'
'     *Unknown 1 was settled by measurement (2026-08-22)*
'       Even when passing raw ExecuteScript a *Promise that never resolves*, the completion notification
'       came back immediately with resultJson={}. So ExecuteScript does not wait for a Promise.
'       Therefore the wait cannot be written on the JS side and *is polled on the VBA side* (the backbone of issue 3).
'       (In the first place EvalSync fires JSON.stringify(expression), so even if it could wait,
'        a Promise would be flattened to {} on the spot. Waiting on the JS side would mean redoing D-1's wrapping)
'
'     *Do not break / step through while polling* (spec fact 20)
'       EvalSync's DoEvents section stretches by the wait time. Wait 10 seconds and it is a danger zone for the full 10 seconds.
'
'     *The meaning of Nothing / False is split by LastEvalOk (design principle 93)*
'       WaitFor  -> Nothing + LastEvalOk=True  ... did not appear in time (normal)
'                   Nothing + LastEvalOk=False ... failure (bad selector / in-callback, etc.)
'       WaitGone -> False   + LastEvalOk=True  ... did not disappear in time (normal)
'                   False   + LastEvalOk=False ... failure
'       *As soon as it is known to be a failure, it exits without waiting* to avoid freezing for 10 seconds in-callback.
'
'     ImplicitWaitSec applies an implicit retry to QuerySelector / GetElementById
'     (same style as SeleniumVBA). *The default is 0, and only those who set it explicitly are affected*
'     Reason: if it waited by default, confirming "it really is not there" would mean waiting for the timeout every time, and
'     the meaning fixed in D-2, "Nothing + LastEvalOk=True = really does not exist", would be blurred.
'
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  Stage K-1 (moving the debug log to file output) ---
'
'   The K-1 change is *only where the log goes*; not a single line of logic was changed.
'     The Debug.Print calls in EvalSync (9) and ClearElementRegistry (2) were
'     replaced with Wv2Log.LogE / LogW / LogD.
'
'   Wv2Log also fires Debug.Print internally, so how things look in the Immediate window does not change
'   (it can be turned off with Wv2Log.LogEcho = False). The log is kept in %APPDATA%\Wv2Browser\logs\
'   as one file per launch. The path can be found with ?Wv2Log.LogPath.
'
'   *The Debug.Print calls in other procedures are still untouched* We migrate in stages
'   (not rewriting every existing Debug.Print at once was the K-1 agreement).
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  Stage D-2 (element registry + Wv2Element reads) ---
'
'   Scope of D-2 (the JS-side element registry, and reading element handles):
'     GetElementById / QuerySelector ... find one element and return a Wv2Element.
'     ClearElementRegistry           ... manually rebuild the JS-side registry.
'     CurrentDomGen                  ... cache of the registry generation (for the fast path).
'     Elem_Read / Elem_CheckStale    ... Friend paths dedicated to Wv2Element.
'     *Not done* writing (Value =) / Click / SetAttribute are D-3.
'                     QuerySelectorAll (returning multiple) is D-3 or later (issue 7, option a).
'
'   # Added Public members
'     GetElementById(elementId) As Wv2Element
'     QuerySelector(selector)   As Wv2Element
'         Returns a Wv2Element if found, *Nothing* if not.
'         *Nothing has two meanings, so they are split by LastEvalOk (issue 4)*
'           Nothing + LastEvalOk = True  ... really does not exist (normal)
'           Nothing + LastEvalOk = False ... failure. The reason is in LastEvalError
'                                           (in-callback / timeout / selector syntax error, etc.)
'     ClearElementRegistry() As Boolean
'         Discards the JS-side registry and rebuilds it with a new generation. All existing Wv2Element objects
'         become stale. For cleanup when the array has grown from many searches on the same page.
'     CurrentDomGen As String
'         The registry generation last received from JS. An empty string if never touched yet.
'
'   # Shape of the JS-side registry (issue 1, option c: *no injection*)
'     window.__wv2 = { gen: 'g<time>-<random>', els: [element, element, ...] }
'
'     *Neither AddScriptToExecuteOnDocumentCreated nor injection at DOMContentLoaded is used.*
'     Instead we made *the search expression itself "create it itself if missing"*:
'
'       (function(){ if(typeof window.__wv2==='undefined'||!window.__wv2){
'                      window.__wv2={gen:'g'+Date.now()+'-'+Math.random(),els:[]}; }
'                    var R=window.__wv2; ... })()
'
'     Advantages:
'       *No new COM mechanism is needed (the former needs a HandlerKind and vtable for the completion handler)
'       *No extra round trips (check, create and search fit into one expression)
'       *Whether it vanishes by page navigation / SPA rewrite / about:blank,
'         it is rebuilt automatically at the next access (same self-healing as design principle 74)
'       *No wasted injection into the tab bar / nav bar / settings tab Panes
'     Disadvantages:
'       *Cannot be used to plant a hook *before* the page loads.
'         When that is needed, implement AddScriptToExecuteOnDocumentCreated as a separate stage.
'
'   # Generation number (issue 2, option b: *issued by JS*)
'     The generation is a string the JS side builds as 'g' + Date.now() + '-' + Math.random().
'     *The approach of counting navigations on the VBA side (option a) is not adopted.*
'       With an SPA's pushState the DOM stays while the event fires, and conversely even if in-page JS
'       swaps elements the event does not fire, so it drifts from reality.
'     The correspondence "registry vanishes = generation vanishes = a rebuild always gives a different value" is guaranteed, and
'     *the authority for the verdict is always the JS side*.
'
'     m_domGen is a cache of "the generation last heard from JS" and is used only for the fast path:
'       if the generation in hand <> m_domGen, *stale can be determined without a round trip* (safe)
'       if the generation in hand =  m_domGen, JS may still have rebuilt it, so always ask
'     When the cache is stale it always falls to the "ask" side, so no misjudgement occurs.
'
'     *Why Math.random() is mixed in* Date.now() has millisecond precision, so if the
'       registry is rebuilt twice within the same millisecond it could collide. Since the value is only used for comparison,
'       we make it a string and remove the collision worry altogether (an additional agreement on issue 2).
'
'   # Releasing handles (issue 3, option a: *do not release*)
'     Wv2Element.Class_Terminate does not delete the JS-side slot. We cannot control
'     when or in what context Terminate runs, and firing EvalSync from there easily runs into D-1's
'     in-callback guard or a Pane that has already been Ctrl_Close'd.
'     The registry vanishes wholesale on page navigation, so actual harm is small. Manual cleanup is
'     ClearElementRegistry. Proper lifetime management is D-4.
'
'   # How to build JS strings (*do not write backslash + double quote inside VBA strings*)
'     Selectors and attribute names come from the user, not the page, so quotes and
'     backslashes may appear. JsQuote passes them through Wv2Json.JsonEscape and wraps them with
'     Chr$(34), so that sequence never appears in the VBA source.
'     Example: both QuerySelector("input[name='q']") and QuerySelector("a[title=""x""]") work.
'
'   # Post-hoc report (design principle 79)
'     (a) *JsonPickStr was moved to Wv2Json.bas* (Private -> Public)
'         Issue 6 said "add JsonUnescape and rewrite JsonPickStr so that it calls it", but
'         since Wv2Element needs the same decoding, leaving it Private in Wv2Pane
'         would make two implementations. We judged the move to be the form most faithful to the intent of issue 6.
'         Apart from changing the 3 call sites in EvalSync to Wv2Json.JsonPickStr,
'         not one character of logic was changed (the scanning part was ported identically to Wv2Json.JsonUnescapeAt).
'     (b) *LastOk / LastError were added to Wv2Element (issue 5's 8 members + 2)*
'         Since the read family returns an empty string on failure too, it is unusable unless there is a way to
'         tell "truly empty" from "stale / timeout". The Pane-side LastEvalOk is the success of the JS execution itself,
'         so it would be True when stale and cannot substitute.
'     (c) *The return values of Elem_Read / Elem_CheckStale are Boolean + string, not an Enum*
'         Whether a class module's Public Enum is visible unqualified from other modules is
'         unverified in this design, and a compile error would waste a round trip.
'         Aligning with the same style as D-1's LastEvalOk / LastEvalError is also more consistent.
'
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  Stage D-1 (synchronous-fetch primitive for ExecuteScript) ---
'
'   *The stage numbering scheme has changed*
'     The "9.xx" up to stage 9.32b were milestones on a single axis, "turning WebView2 into a browser",
'     and stage 9.32b closed it off. From D-1 on, a symbol is assigned per axis:
'       (plain 9.xx) ... browser feature work (finished, frozen)
'       D          ... DOM / page automation (MSHTML equivalent)  <- this stage is here
'       C          ... form controls (TreeView / ListView equivalent)
'       K          ... foundation / paying down debt (Sentinel etc.)
'
'   Scope of D-1 (throw a JS expression and receive its return value synchronously in VBA):
'     EvalSync ... the deliverable of this stage. Fires one JS expression, waits with DoEvents until completion, and
'                returns the JSON representation of the return value as a String.
'     *Not done* the element registry (__wv2els[]), Wv2Element, the DOM API. D-2 and later.
'
'   # Added Public members
'     EvalSync(jsExpr, [timeoutSec=5]) As String
'         Success: returns the string obtained by JSON.stringify of the JS return value.
'               1+2       -> 3
'               'abc'     -> abc including the quotes (*strings come back with quotes*)
'               undefined -> undefined   (because JSON.stringify returns undefined)
'               null      -> null
'               ({a:1})   -> the JSON representation of {a:1}
'         Failure: returns an empty string. See LastEvalError for the reason.
'     LastEvalOk    As Boolean ... whether the previous EvalSync succeeded
'     LastEvalError As String  ... the reason for the previous failure
'                                 closed / no-view / in-callback /
'                                 execute-failed(n) / timeout(n s) /
'                                 hresult(&Hxxxx) / ReferenceError: xxx /
'                                 syntax-error-or-null (...)
'     ResetCallbackGuard()     ... reset the in-callback guard counter to 0 (emergency use)
'     InCallbackDepth As Long  ... current value of the guard counter (for verification)
'     Debug_SetInCallback(n)   ... injection port for verifying that the guard fires (verification only)
'
'   # Issues and decisions (agreed with tarboh)
'     Issue 1 method name     : EvalSync, with no prefix. View_ / Ctrl_ mean "hit a COM
'                            interface method one-to-one", so they are not
'                            attached to a high-level API with a built-in wait loop.
'     Issue 2 result receiver : a pending table keyed by callbackId (m_evalPending).
'                            Receiving into a single variable would mix up concurrent calls.
'                            The direct call to UserForm1.OnExecuteScriptResult stays (log only).
'     Issue 3 return shape    : String (= the JSON representation of value). Success is obtained with LastEvalOk /
'                            LastEvalError. Type conversion is done in D-2 and later.
'     Issue 4 via Safe        : always go through View_ExecuteScriptSafe. Raw ExecuteScript returns
'                            a JS exception as S_OK + null, indistinguishable from a normal null.
'     Issue 5 timeout         : default 5 seconds. When exceeded, the pending entry is *left behind marked discarded*,
'                            and a completion notice arriving late sees that and discards it (it is kept in the log).
'     Issue 6 concurrency/nesting: allowed. They are separated by callbackId, so no mix-up.
'                            For nesting, only the depth is written to the log.
'     Issue 7 in-callback guard:
'                            *The most valuable part of this stage* If EvalSync is called from inside a WebView2 event
'                            handler, the runtime waits for the handler to return while we wait for the JS to complete, so
'                            it hangs until the timeout (spec fact 14). Moreover, spec fact 20 makes this an area where the debugger
'                            cannot be used, so finding the cause is painful. So
'                            if m_inCallback > 0, fail immediately without firing the JS.
'
'
'   # Post-hoc report on the implementation (design principle 79)
'     (a) *The handlers were split into a Core* At agreement time it was described as "+1 at the entry of each View_On*, -1 at the exit",
'         but reading the actual code, View_OnNewWindowRequested has 5 Exit Sub points in the middle
'         (the others one each), and simply adding 2 lines would create paths that skip the decrement.
'         So the body was renamed to Private ... _Core, and a Public thin wrapper of the same name
'         does +1 -> call Core -> -1.
'         There are 7 targets (NavigationStarting / NavigationCompleted / DocumentTitleChanged /
'         WebMessageReceived / NewWindowRequested / HistoryChanged / DOMContentLoaded).
'     (b) *The counter is restored on error too. But the error is not swallowed* The wrapper
'         restores the counter with On Error GoTo GuardExit and then *re-raises the same error* with Err.Raise.
'         ComCallbackHandler has no error handling, and previously the error
'         propagated as is, so this was done to leave behaviour unchanged (swallowing it would hide
'         defects). Still, ResetCallbackGuard is provided against accidents where the counter is left behind.
'     (c) *No guard was attached to OnExecuteScriptCompleted* This is the receiver of results and
'         no user code runs inside it.
'         Attaching one would raise the counter during EvalSync's DoEvents, which is harmless but confusing.
'     (d) *JSON extraction was made a Private function inside Wv2Pane* Per design principle 64, Wv2Json.JsonGetStr is an implementation that
'         *does not unescape* and returns everything up to the next closing quote
'         as is. EvalSync's return value passes through JSON.stringify twice, so
'         "backslash + quote" and "backslash + uXXXX" are always present.
'         So JsonPickStr (a limited parser that unescapes) was provided as Private.
'         When it is needed elsewhere, promote it to Wv2Json.
'     (e) *It was made to wrap with JSON.stringify* EvalSync fires
'         View_ExecuteScriptSafe("JSON.stringify(" & jsExpr & ")"). That way the value of the structured JSON
'         that Safe returns is *always a string*, and the limited parser of (d) alone
'         can extract it regardless of type. For undefined the value key vanishes entirely, so
'         we detect that and return undefined.
'
'   # Usage notes (to be relaxed in D-2 and later)
'     *jsExpr must be an "expression", not a "statement" (inherits View_ExecuteScriptSafe's
'       constraint as is). To write multiple statements, wrap them in an immediately-invoked function.
'     *The result of an expression that returns a string comes back with quotes. In D-2 the type is decided
'       per Wv2Element property and the quotes are stripped.
'     **Do not break / step through during the DoEvents loop* (spec fact 20)
''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  Stage 9.27 (the main target: put_NewWindow loads the navigation into a spare tab) ---
'
'   Scope of 9.27 (open the new-tab expansion of POST links correctly in "our own tab"):
'     Extended the mode branching in View_OnNewWindowRequested from 2 values to 3.
'       0 = NWM_LEGACY  : current behaviour (put_Handled + RequestNewTabAsync). Navigate is
'                         *always GET*, so a POST link gets a 405.
'       1 = NWM_POPUP   : the 9.26 probe. Does not set put_Handled and passes through, letting the runtime
'                         open a popup. POST is preserved (spec fact 23).
'       2 = NWM_PREWARM : *the real aim of 9.27* Passes the CoreWebView2 of the spare tab that Browser has warmed up
'                         to args.put_NewWindow (vtable index 4). The runtime
'                         loads the original request into that WebView, so POST is preserved and,
'                         moreover, it lands in our own tab.
'
'   *Nothing is waited for inside the event* All that is done here is
'       (1) get the spare ViewPtr with m_browser.ClaimPrewarmViewPtr
'       (2) dcf(pArgs, 4, put_NewWindow, pNew)
'       (3) dcf(pArgs, 6, put_Handled, 1&)   * when SetHandledWithNewWindow is True
'       (4) ask the Timer to adopt it with m_browser.RequestAdoptPrewarmedAsync
'     these 4 only. There is no Controller creation and no Deferral, so neither spec fact 14 (SafeTimer
'     deadlock) nor 20 (the EBMode window) is touched. *This is the key point of the pre-warm approach*
'
'   *Fallback* If the spare is absent / not yet Ready / already Claimed / put_NewWindow fails,
'     in any case it falls straight to the current behaviour below (put_Handled + RequestNewTabAsync).
'     For a POST link it returns to a 405 but does not break. The fact that it fell back is always logged *with a mark.
'
'   *Basis of the vtable index (ICoreWebView2NewWindowRequestedEventArgsVtbl in WebView2.h)*
'       3 = get_Uri / 4 = put_NewWindow / 5 = get_NewWindow /
'       6 = put_Handled / 7 = get_Handled / 8 = get_IsUserInitiated / 9 = GetDeferral
'     put_Handled = 6 is a proven value in use since 9.6b and matches this layout.
'
'   *Untouched* the get_Uri retrieval and log / the empty-URL guard / RequestNewTabAsync (option beta
'     Timer, the workaround for spec fact 14) / all other event handlers.
'
' --- Wv2Pane.cls  Stage 9.26 (NewWindowRequested delegation probe) ---
'
'   Scope of 9.26 (isolating why POST links get a 405. A round that only observes):
'     Just added one branch at the top of View_OnNewWindowRequested.
'     When m_browser.PopupDelegateMode is True,
'       *do not set args.put_Handled
'       *do not call m_browser.RequestNewTabAsync
'     and pass through, letting the WebView2 runtime itself open the popup window
'     (behaviour confirmed in 9.6a: a Chrome_WidgetWin_1 appears inside the Excel process).
'     Since the runtime takes care of the navigation itself, if the original request was POST it continues
'     including the body. This measures "does delegating to the runtime preserve POST?".
'
'     *Why the option 'View_Navigate on the same View' is not adopted*
'       That form, which existed in 9.6b, calls ICoreWebView2::Navigate, so *always GET*.
'       Whether it flows to a new tab or the same tab does not matter; once it goes through Navigate it stays a 405.
'       The starting point of the isolation is that the only way to preserve POST is "hand the navigation to the runtime" (either this delegation or the next stage's
'       put_NewWindow).
'
'     *When m_browser is Nothing* In standalone Pane operation (SetBrowser not run) there is no place for the flag
'       at all, so it falls back to View_Navigate on the same View as before.
'       The behaviour is unchanged since 9.10e.
'
'   *Untouched* the get_Uri retrieval and log / the empty-URL guard / handling when put_Handled fails /
'     RequestNewTabAsync (option beta Timer, the workaround for spec fact 14) / all other event handlers.
'     No design principles or spec facts were added.
'
' --- Wv2Pane.cls Stage 9.19 ---
'
'   Additions in 9.19 (duplicate-prevention flag for the settings tab):
'     *Added the m_isSettings flag (identifies the settings tab)*
'       A boolean flag that holds whether this Pane is the "settings tab". Wv2Browser's
'       OpenSettingsTab sets IsSettings = True on this Pane when it creates the settings tab.
'       When [gear] is pressed again, Browser scans m_tabs from the top and, if there is a Pane with
'       IsSettings = True, just does ActivateTab on it and does not create a new one
'       (preventing the settings tab from multiplying when [gear] is mashed).
'
'     *Why identify by a Pane flag (option A)*
'       The settings tab displays data:text/html;base64,... via NavigateToString,
'       so judging by URL/content is fragile and costly. If the Pane is marked,
'       just scanning m_tabs each time also serves as a "liveness check" (a closed
'       settings tab drops out of m_tabs and so naturally is not found). It has fewer
'       weak points than the option of Browser holding a Pane reference and watching liveness itself
'       (which would need dangling detection).
'
' --- Wv2Pane.cls Stage 9.18 ---
'
'   Additions in 9.18 (exposing a VBA object to JS via AddHostObjectToScript):
'     *Added View_AddHostObjectToScript(objName, instance)* (part 1)
'       Calls ICoreWebView2.AddHostObjectToScript (vtable index 49 on View) and exposes
'       the VBA class instance on the JS side as window.chrome.webview.hostObjects.<name>.
'       JS can then call its methods/properties directly
'       (a direct RPC path that does not go through postMessage string commands).
'
'     *The implementation is a faithful port of the old WebView2.cls (verified on a real machine)*
'       Passing is the same dcf convention as the existing methods:
'         Set vObj = instance      (put the object into a Variant)
'         dcf(m_pView, 49, ..., StrPtr(objName), VarPtr(vObj))
'       The first argument is the host object name (LPCWSTR), the second is a VARIANT* (a pointer to the
'       packed object). The return value is an HRESULT. It works with Instancing left at 1 - Private
'       (confirmed in the old implementation. It suffices that the IDispatch VBA generates can be
'       exposed to the OS).
'
'     *Use (as of 9.18)*
'       When Wv2Browser opens the settings tab, it attaches Wv2SettingsBridge to this Pane and lets
'       the settings screen's JS operate the search engine directly, e.g. hostObjects.sync.Settings.SetEngine("bing").
'       The attach is done before NavigateToString (order A: the hypothesis that if you register and then Navigate,
'       the loaded JS can use hostObjects from the start, to be verified on a real machine. If it fails, fall back
'       to the ready handshake approach = order B).
'
'
'     *RemoveHostObjectFromScript (vtable 50) is not added* (YAGNI)
'       The settings tab is released together with its WebView2 when closed, so explicit detaching
'       is unnecessary in 9.18. Add it in the stage where it becomes needed.
'
' --- Wv2Pane.cls Stage 9.12a (formerly: Class9) ---
'
'   Additions in 9.12a (foundation for the address bar / nav buttons: notification of URL changes):
'     *The structure is exactly the same as 9.11a's NotifyTitleChanged* The Pane tells its owning Browser that
'       "my URL state has changed", and Browser looks up the index in reverse and
'       does RaiseEvent TabUrlChanged. The Pane does not know the nav bar exists (loose coupling).
'     - New Private Sub: NotifyUrlState(pSender)
'         Takes get_Source(4) / get_CanGoBack(38) / get_CanGoForward(39) via sender
'         and calls m_browser.NotifyUrlChanged Me, url, canBack, canFwd.
'     - View_OnNavigationCompleted: calls NotifyUrlState at the end (normal page navigation).
'         *Also call it on the early Exit Sub path where pArgs=0* (route both paths through it)
'     - View_OnHistoryChanged: since url/canBack/canFwd have already been obtained, calls
'         m_browser.NotifyUrlChanged with those values (catches SPA pushState, back/forward).
'     *Multiple firing is harmless* Even if NavigationCompleted and HistoryChanged both fire,
'       it is the whole-sync approach (design principle 60), so sending the same state twice stays consistent.
'
'   Additions in 9.11c (foundation for drawing the VBA->JS tab list: title cache + WebMessage sink):
'     *Title cache (option 3)* Holds the title tied to the Pane object.
'       Even if the tab index shifts, the title follows the correct pane,
'       and the source of truth for whole-sync can be unified. TabBar reads it with Browser.PaneAt(i).DocumentTitle.
'
'       - New field: m_title As String
'       - New Public Property Get DocumentTitle() As String
'       - In View_OnDocumentTitleChanged, save m_title = title after getting the title
'         (just add one line right before the NotifyTitleChanged notification block of 9.11a).
'     *Generic WebMessage sink (option P)* A mechanism that forwards messages posted by JS with postMessage to a
'       registered "sink" (TabBar etc.). It is a separate line from the m_browser back reference
'       (for NewWindowRequested).
'       - New field: m_webMsgSink As Object (the receiving side implements OnPaneWebMessage)
'       - New Public Sub: SetWebMsgSink / ClearWebMsgSink
'       - View_OnWebMessageReceived: added a block that, if there is a sink, calls OnPaneWebMessage(Me, source, msg).
'
'       - View_AddWebMessageReceived uses the existing one (9.4c) as is, and the TabBar side
'         enables reception.
'     *Disconnecting the sink on release* ClearBrowserRef also sets m_webMsgSink to Nothing
'       (design principle 58. Disconnection of back-reference systems is centralized in the Shutdown path).
'
'   Additions in 9.11a (title notification for the loosely coupled tab bar):
'     - Added one block to View_OnDocumentTitleChanged. After getting the title,
'       if m_browser exists, calls m_browser.NotifyTitleChanged Me, title.
'       It reuses the back reference m_browser (introduced in 9.10e for NewWindowRequested) as is.
'       When m_browser is not set (usage without a back reference) it falls back to
'       Debug.Print only, so loose coupling is not broken. There are no new fields/
'       methods on the Pane side (it only uses the existing m_browser).
'
'   Additions in 9.10e (making NewWindowRequested open a new tab, option A + option Y + option beta Timer):
'     *Deadlock avoidance (option beta)* If AddTabWithUrl is called directly inside the NewWindowRequested handler,
'       AddTab's WaitPaneReady waits for ControllerCompleted, but WebView2 does not deliver it until this
'       handler returns, so it deadlocks
'       (timeout) (spec fact 14). So the request is made to m_browser.RequestNewTabAsync, and
'       via SafeTimer (one-shot) AddTabWithUrl is run after the handler completes.
'
'     - New field: m_browser As Wv2Browser (back reference to the owning Browser)
'     - New methods: SetBrowser / ClearBrowserRef (set and disconnect the back reference)
'     - Changed View_OnNewWindowRequested from "Navigate the same View" (9.6b) to
'       "open a new tab with m_browser.AddTabWithUrl". When m_browser is not set it falls back to
'       Navigate on the same View. The return value of window.open becomes null
'       (a simple approach that does not pass a new View via put_NewWindow = option A. The proper way is a separate future stage)
'     *Circular reference caution (design principle 58)* Browser.m_tabs ? Pane.m_browser form a
'       mutual reference. Wv2Browser.Shutdown breaks it by calling each Pane's ClearBrowserRef
'       (via UserForm_Terminate). Class_Terminate does not touch m_browser
'       (in a cycle Class_Terminate itself is not called, so Shutdown is the only path)
'
'   Additions in 9.10d-2a (this session, https origin foundation for SPAs):
'     - New module variable:
'         m_pView3 As LongPtr (cache of the ICoreWebView2_3-derived IF)
'     - Added m_pView3 = 0 to Class_Initialize
'     - Extended the Release order in Class_Terminate / Ctrl_Close:
'         m_pView3 -> m_pView2 -> m_pView -> m_pCtrl (most derived first, design principle 51)
'     - New Private Function:
'         EnsureView3() As Boolean
'         (lazily acquired with QueryInterface(IID_ICoreWebView2_3), same form as EnsureView2)
'     - New Public Function:
'         View3_Ensure() As Boolean (verification wrapper, same form as View2_Ensure)
'         View3_SetVirtualHostNameToFolderMapping(hostName, folderPath, accessKind)
'           As Long
'           (ICoreWebView2_3 vtable index 71. The foundation that maps a local folder to
'            https://<hostName>/ and avoids the about:blank origin's
'            pushState/replaceState restriction (spec fact 6).
'            It is a simple-argument type, so View_Navigate is the model and design principle 45 is not needed)
'     - IID_ICoreWebView2_3 = {A0D6DF20-3B92-416D-AA0C-437A9C727857}
'       (built each time as a local variable in EnsureView3, following option a of 9.9a)
'
'   Additions in 9.10c-1 (visibility control):
'     - New Public Function:
'         Ctrl_PutIsVisible(visible As Boolean) As Long
'         (ICoreWebView2Controller::put_IsVisible, vtable index 4)
'         Used by Wv2Browser's ActivateTab for tab visibility control (only the active tab is visible).
'         It is a Win32 BOOL (True=1/False=0), so normalize to 1/0 before dcf (spec fact 11).
'         It is a simple-argument type, so Ctrl_PutBounds is the model.
'
'   Additions in 9.9c (batch handling of the completion group):
'     - New Public function: View_ExecuteScriptSafe(userCode) As Long
'         An ExecuteScript wrapper that wraps user code in try-catch and returns structured JSON.
'         resultJson is always of the form {ok:true,value:...} or {ok:false,errorName:...,...}
'         (option F: a + X + *).
'     - OnExecuteScriptCompleted calls UserForm1.OnExecuteScriptResult directly
'         (option alpha second implementation, option E1-a-i = Wv2Pane holds no UserForm1 reference.
'          Uses VBA's global UserForm reference, no circular reference).
'         The Debug.Print inside Wv2Pane also coexists (option E2-X, easy to isolate regressions).
'     - Log improvements in View_OnHistoryChanged:
'         (alpha) if about:blank, annotate "normal in a NavigateToString environment".
'         (beta) state in comments and log that pArgs=0 is by specification.
'     - No essential change to logic in the design (extension only).
'
'   Additions in 9.9b (continued, SPA support):
'     - Raised the upper limit of the persistent handler array from HK_NewWindowRequested(7) to HK_DOMContentLoaded(10)
'
'     - New helper RemoveInterfaceFor(kind) As LongPtr
'     - HistoryChanged family (base ICoreWebView2, vtable 13/14):
'         View_AddHistoryChanged / View_RemoveHistoryChanged / View_OnHistoryChanged
'     - DOMContentLoaded family (ICoreWebView2_2, vtable 64/65):
'         View2_AddDOMContentLoaded / View2_RemoveDOMContentLoaded /
'         View2_OnDOMContentLoaded
'
'   * Caution about design principle 45 in 9.9b *
'     Adding 2 persistent handlers = 2 occasions of passing a COM object to WebView2.
'     Based on the lesson of passing h.Slot directly and crashing in 9.8c,
'     View_AddHistoryChanged / View2_AddDOMContentLoaded were both implemented using
'     View_AddNavigationStarting verbatim as the model,
'     always going through pVTableObj = h.Slot + VTABLE_OBJ_OFFSET.
'     Simple-argument types like View_PostWebMessageAsString are not used as the model.
'
'   Additions in 9.9a (continued):
'     - New module variable:
'         m_pView2 As LongPtr (cache of the ICoreWebView2_2-derived IF)
'     - Added m_pView2 = 0 to Class_Initialize
'     - Extended the Release order in Class_Terminate:
'         (Settings -> View2 -> View -> Ctrl, strictly child to parent, design principle 43)
'     - Applied the same Release order to Ctrl_Close too (Release m_pView2 before View)
'     - New Private Function:
'         EnsureView2() As Boolean
'         (if m_pView2 has not been acquired, get it with QueryInterface(IID_ICoreWebView2_2),
'          same pattern as EnsureSettings)
'     - New Public Function:
'         View2_Ensure() As Boolean
'         (verification wrapper, call it from the Immediate window to confirm acquisition succeeded)
'
'   Design decisions of 9.9a (agreed with tarboh):
'     *Option 1: m_pView2 field + EnsureView2 lazy acquisition (following the Settings style)
'     *Option a: IID_ICoreWebView2_2 is built each time as a local variable in EnsureView2
'             (no IID table added to Wv2Thunks, keeping 9.9a to a minimal Wv2Thunks change)
'     *Option X: both EnsureView2 / View2_Ensure are As Boolean (same form as EnsureSettings)
'     *Option i: verification is calling View2_Ensure from Immediate + confirming
'             zero regressions with Test_AllEvents_Real (no new test function)
'     *Option R1: Release order is Settings -> View2 -> View -> Ctrl (strictly child to parent)
'
'   * Significance of 9.9a *
'     Fetching ICoreWebView2_2 is not the goal in itself; this is the stage that prepares the foundation for
'     calling the SPA-support methods of 9.9b and later (DOMContentLoaded etc., methods at vtable 64/65).
'     HistoryChanged is on the base IF (vtable 13/14),
'     so 9.9b could implement it without _2, but DOMContentLoaded
'     requires the _2 IF. It is also the first stage to apply, in practice,
'     the naming convention established in 9.6a (View2_xxx).
'
'   Design principles followed in 9.9a (the main related ones):
'     *Design principle 31: naming convention by origin prefix (View2_xxx)
'     *Design principle 42: lazy acquisition / caching of derived IFs (the EnsureSettings style)
'     *Design principle 43: Release order is child to parent
'     *Design principle 45: add VTABLE_OBJ_OFFSET when passing as a COM object
'                    (9.9a is a case of "receiving a COM object from WebView2",
'                     so VTABLE_OBJ_OFFSET does not appear, but we will be
'                     wary again when registering the DOMContentLoaded handler in 9.9b)
'     *Design principle 48: safety after Close is a "member clear + flag" hybrid
'
'   Additions in 9.8c (continued):
'     - New module variables:
'         m_executeScriptCounter As Long (callback ID numbering for View_ExecuteScript)
'         m_closed As Boolean (flag for Ctrl_Close done)
'     - Added initialization of the above 2 to Class_Initialize
'     - New Public function (View_ family):
'         View_ExecuteScript (jsCode As String) As Long  ' returns callback ID
'     - New Public Sub (receiver for dynamic binding):
'         OnExecuteScriptCompleted (callbackId, errorCode, resultJson)
'     - Extension of Ctrl_Close (option E 1+3 hybrid):
'         Fires Controller::Close + Release of m_pSettings/m_pView/m_pCtrl +
'         clearing of m_persistentTokens / m_persistentHandlers +
'         m_closed = True + m_state = Cs_Failed
'         This makes the double processing in Class_Terminate after Close skip cleanly
'
'   * Initial bug in 9.8c + fix history *
'     In the first version of View_ExecuteScript, passing h.Slot directly to dcf crashed
'     (the Excel process died instantly). In fact, as with the other registration families (View_AddNavigationStarting etc.),
'     h.Slot + VTABLE_OBJ_OFFSET must be passed.
'     -> "When passing a thunk slot to WebView2 as a COM object,
'        always add VTABLE_OBJ_OFFSET (= 176 bytes) before passing"
'        This design principle was reaffirmed (newly established as design principle 45).
'
'   Design decisions of 9.8c (agreed with tarboh):
'     *Option a: concurrency OK, the slot itself manages completion
'     *Option alpha: event callback approach, the first implementation only logs in OnExecuteScriptCompleted
'             (can be extended to UserForm notification in the future)
'     *Option X: returns a sequential callback ID
'     *Option E 1+3: member clear + m_closed flag
'
'   Items of 9.8b (continued):
'     - m_pSettings + EnsureSettings + Settings_PutXxx family (3)
'     - View_Reload / View_Stop / View_GetCanGoBack / View_GetCanGoForward /
'       View_GetBrowserProcessId
'
'   Items of 9.8a (continued, rename):
'     - Class name Class9 -> formally fixed as Wv2Pane
'     - Reference replacement Class7 -> ComCallbackHandler / Class8 -> Wv2Environment
'     - For details see each stage section of the dev notes (dev-notes-*.md under docs)

'   Role:
'     A combined class of ICoreWebView2Controller and ICoreWebView2 (combining policy = option a).
'     We give up one-to-one correspondence with the SDK's hierarchy, but from the VBA user's viewpoint
'     we get the straightforward correspondence "one WebView2 display = one class".
'     It is intuitively closest to think of it as "one display panel" pasted into the Frame control area
'     of a UserForm (= the reason it was named "Pane").
'
'   Implemented persistent handlers:
'     HK_NavigationStarting      (View_AddNavigationStarting   / View_OnNavigationStarting)
'     HK_NavigationCompleted     (View_AddNavigationCompleted  / View_OnNavigationCompleted)
'     HK_WebMessageReceived      (View_AddWebMessageReceived   / View_OnWebMessageReceived)
'     HK_DocumentTitleChanged    (View_AddDocumentTitleChanged / View_OnDocumentTitleChanged)
'     HK_NewWindowRequested      (View_AddNewWindowRequested   / View_OnNewWindowRequested)
'     HK_HistoryChanged          (View_AddHistoryChanged       / View_OnHistoryChanged)        ' Stage 9.9b
'     HK_DOMContentLoaded        (View2_AddDOMContentLoaded    / View2_OnDOMContentLoaded)     ' Stage 9.9b
'
'   Naming convention (design principle 31):
'     *ICoreWebView2Controller methods: Ctrl_xxx (base) / Ctrl2_xxx (extension, future)
'     *ICoreWebView2 methods:           View_xxx (base) / View2_xxx (extension, first introduced in 9.9a)
'     *ICoreWebView2 event receivers: View_OnXxx (base) / View2_OnXxx (extension, planned to be introduced in 9.9b)
'     *ICoreWebView2Environment methods: Env_xxx etc. (to be added when needed in the future)
'     *One-shot callbacks (OnEnvironmentCompleted / OnControllerCompleted):
'       left without a prefix (there is no function that directly fetches a derived-interface version, so
'       there is no need to identify a version)
'
'   Two-way communication flow (established in 9.4c, only the reference names updated to the new names):
'     [VBA -> JS]
'       VBA: m_pane.View_PostWebMessageAsString("hello")
'         -> dcf(m_pView, 33, ..., StrPtr("hello"))
'         -> the WebView2 runtime delivers it to JS
'         -> JS: receives with window.chrome.webview.addEventListener('message', ...)
'     [JS -> VBA]
'       JS: window.chrome.webview.postMessage("hello")
'         -> the WebView2 runtime invokes the handler via HK_WebMessageReceived
'         -> ComCallbackHandler.Invoke[HK_WebMessageReceived]
'         -> Wv2Pane.View_OnWebMessageReceived(pSender, pArgs)
'         -> logs with args.get_Source / args.TryGetWebMessageAsString
''''''''''''''''''''''''''''''''''
Option Explicit

#If Win64 Then
    ' OK
#Else
    #Error "This class requires 64-bit VBA (x64)"
#End If

' --- N-1: time source for elapsed milliseconds ---
'   *Timer is not used* Timer is a Single of "seconds since midnight", so
'   for an afternoon value (in the 50,000s) a Single lacks the significant digits and the ms resolution disappears.
'   It also wraps back when crossing a date. GetTickCount is ms since boot and increases monotonically.
'   The return value is a Long, so it wraps around in about 49.7 days. The difference is computed after raising it to Double,
'   and if it goes negative, 2^32 is added to restore it (NetElapsedMs).
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long


Public Enum ControllerState
    Cs_Created = 0
    Cs_Waiting = 1
    Cs_Ready = 2
    Cs_Failed = 3
End Enum


' --- Internal state ---
' m_handler:     strong reference to the HK_ControllerCompleted handler (one-shot)
'                (set to Nothing after OnControllerCompleted)
' m_pCtrl:       the ICoreWebView2Controller received from WebView2
'                (responsible for Release in Class_Terminate)
' m_pView:       the ICoreWebView2 obtained with Ctrl_GetCoreWebView2
'                (responsible for Release in Class_Terminate)
Private m_handler As ComCallbackHandler
Private m_pCtrl   As LongPtr
Private m_pView   As LongPtr
'   Added in stage 9.8b: cache of ICoreWebView2Settings
'     When Wv2Pane first has a Settings-family method hit,
'     it is obtained with View vtable 3 (get_Settings) and held here.
'     Its lifetime matches m_pView, so it is Released in Class_Terminate.
Private m_pSettings As LongPtr
'   Added in stage 9.9a: cache of the ICoreWebView2_2 derived IF
'     When Wv2Pane first has a View2_xxx-family method hit,
'     it is obtained with QueryInterface(IID_ICoreWebView2_2) and held here.
'     Lazily acquired via EnsureView2 (= same style as EnsureSettings, design principle 42).
'
'     m_pView and m_pView2 are the same COM object seen through different IFs.
'     QueryInterface fires a separate AddRef, so m_pView2 has a reference count
'     independent of m_pView. In Class_Terminate / Ctrl_Close,
'     m_pView2 is Released before m_pView (child to parent, design principle 43).
Private m_pView2 As LongPtr
'   Added in stage 9.10d-2a: cache of the ICoreWebView2_3 derived IF
'     Needed to call SetVirtualHostNameToFolderMapping (vtable 71).
'     Lazily acquired via EnsureView3 (same form as EnsureView2, design principles 42/50).
'     ICoreWebView2_3 : ICoreWebView2_2 : ICoreWebView2 is single inheritance, so
'     the Release order is m_pView3 -> m_pView2 -> m_pView (most derived first, design principle 51).
Private m_pView3 As LongPtr
'   Added in stage 9.10e: back reference to the owning Wv2Browser
'     Lets View_OnNewWindowRequested call Wv2Browser.AddTabWithUrl to open a new tab.
'     Wv2Browser.AddTab sets itself with SetBrowser.
'     *This forms a circular reference (Browser.m_tabs ? Pane.m_browser), so release is
'      explicitly cut via Wv2Browser.Shutdown -> ClearBrowserRef (design principle 58).
Private m_browser As Wv2Browser
Private m_state   As ControllerState
Private m_lastError As Long

' --- Holding persistent event handlers (added in stage 9.4a, upper limit extended in 9.6a/9.9b) ---
' m_persistentHandlers: ComCallbackHandler array indexed by the value range of HandlerKind
' m_persistentTokens:   EventRegistrationToken (LongLong, 8 bytes) at the same index
'
' In stage 9.6a the upper limit was extended from HK_DocumentTitleChanged to HK_NewWindowRequested.
' In stage 9.9b the upper limit was extended from HK_NewWindowRequested(7) to HK_DOMContentLoaded(10).
' In stage N-1 the upper limit was extended from HK_DOMContentLoaded(10) to HK_WebResourceRequested(11).
' In stage N-3 the upper limit was extended from HK_WebResourceRequested(11) to HK_WebResourceResponseReceived(12).
'   This allows HK_HistoryChanged(9) / HK_DOMContentLoaded(10) to be
'   managed in the persistent array too.
'
' Note: HK_ExecuteScriptCompleted(8) is included in the range 3?10, but
'   ExecuteScript is a one-shot family running concurrently, so it is not held in the persistent array
'   (= index 8 always stays Nothing / 0, unused). It is naturally skipped by the
'   "If Not (m_persistentHandlers(k) Is Nothing)" check in CleanupAllPersistentHandlers,
'   so there is no actual harm.
'
' Class_Terminate calls remove_ for every non-Nothing handler
' (the LBound/UBound loop of CleanupAllPersistentHandlers needs no change;
'  this is where "automating horizontal rollout with the HandlerKind-linked array" shows its true value).
Private m_persistentHandlers(HK_NavigationStarting To HK_WebResourceResponseReceived) As ComCallbackHandler
Private m_persistentTokens(HK_NavigationStarting To HK_WebResourceResponseReceived) As LongLong

' --- Page load complete flag (added in stage 9.4c, provisional implementation) ---
'
'   Set to True in View_OnNavigationCompleted.
'   Used by external code (UserForm etc.) for synchronization such as
'   "wait for the page load to complete, then fire PostMessage".
'
'   Note (provisional implementation):
'     This only means "completion of the first navigation". It cannot catch the case where, as in an SPA, the page content
'     is updated via history.pushState or fetch.
'     SPA support (DOMContentLoaded, HistoryChanged, etc.) is planned to be redesigned
'     all together when the comprehensive class is introduced.
Private m_pageLoaded As Boolean
Private m_lastNavOk As Boolean    ' English fork: IsSuccess of the last NavigationCompleted


' Added in 9.8c:
'   m_executeScriptCounter: a sequence counter numbered on each call of
'     View_ExecuteScript (option X adopted). Stored in each ComCallbackHandler.m_callbackId.
'     Initial value 0, becomes 1 on the first call.
'   m_closed: indicates that Ctrl_Close was called, or after Class_Terminate has run.
'     While True, View_xxx / Ctrl_xxx / Settings_xxx return early at their top.
'     This safely rejects calls to the various methods made by mistake after Close.
'     (option E 1+3 hybrid: member clear + m_closed flag)
Private m_executeScriptCounter As Long
Private m_closed As Boolean

' --- Added in stage 9.11c ---
'   m_title: title cache of the document this pane is displaying
'     (option 3). Updated in View_OnDocumentTitleChanged. By tying the title to the Pane
'     object, the title follows the correct pane even if the tab index shifts
'     (unifying the source of truth for whole-sync). TabBar reads it with
'     Browser.PaneAt(i).DocumentTitle.
'   m_webMsgSink: forwarding destination of messages posted by JS with postMessage (option P).
'     TabBar registers itself as the sink. The receiving side must implement OnPaneWebMessage
'     (As Object because of late binding). It is a separate line from m_browser (the
'     back reference for NewWindowRequested). Release is cut with ClearBrowserRef (design principle 58).
Private m_title As String
Private m_webMsgSink As Object
Private m_isSettings As Boolean   ' Stage 9.19 (whether this Pane is the settings tab)

' --- Stage 9.27: how NewWindowRequested is handled (corresponds to Wv2Browser.NewWindowMode) ---
'   A class module cannot have Public Const, so the same values as on the Browser side are held
'   as Private Const. For the meaning of the values, see the comment in the header.
Private Const NWM_LEGACY  As Long = 0
Private Const NWM_POPUP   As Long = 1
Private Const NWM_PREWARM As Long = 2


' --- D-1: module variables and constants for EvalSync (synchronous-fetch primitive) ---
'   m_evalPending : pending table keyed by callbackId (issue 2).
'                   The key is cb & callbackId, the value is Array(state, errorCode, resultJson).
'                   *Items of a Collection cannot be rewritten, so update with Remove -> Add*
'   m_lastEvalOk / m_lastEvalError : success and reason of the previous EvalSync (issue 3).
'   m_inCallback  : depth of being inside a WebView2 event handler (guard of issue 7).
'                   The Public wrapper does +1 / -1. If not 0, EvalSync does not fire.
'   m_evalDepth   : nesting depth of EvalSync. Only written to the log, not used to suppress (issue 6).
Private m_evalPending   As Collection
Private m_lastEvalOk    As Boolean
Private m_lastEvalError As String
Private m_inCallback    As Long
Private m_evalDepth     As Long

' State values of the pending table
Private Const EVAL_WAITING   As Long = 0
Private Const EVAL_DONE      As Long = 1
Private Const EVAL_DISCARDED As Long = 2

' Default timeout (seconds). Make it shorter than READY_TIMEOUT (10 seconds waiting for Controller creation).
' Running one JS expression is originally on the order of milliseconds, so at 5 seconds of waiting it is considered abnormal.
Private Const EVAL_TIMEOUT_DEFAULT As Single = 5

' --- D-3: settings for DOM condition waits (issue 2 option C / issue 4 option A) ---
'   m_implicitWaitSec : implicit wait (seconds) applied to QuerySelector / GetElementById.
'                       *Default 0 = disabled* Set with the ImplicitWaitSec property.
'   POLL_INTERVAL_SEC : polling interval (the fixed 100 ms of issue 4 option A).
'                       *One EvalSync round trip measures 38 ms, so the actual check interval is
'                       about 140 ms* (from the real-machine log of 2026-08-22).
'   WAIT_TIMEOUT_DEFAULT : default upper limit (seconds) of WaitFor / WaitGone.
Private m_implicitWaitSec As Single
Private Const POLL_INTERVAL_SEC    As Single = 0.1
Private Const WAIT_TIMEOUT_DEFAULT As Single = 10


' --- D-4a: version number of the in-page SPA probe ---
'   When the probe's contents change, *always raise it*. Even if an old version remains on the page,
'   a different version fails the health check and it is rebuilt (same contract as the reference product).
Private Const SPA_PROBE_VER As Long = 3


' --- D-6: QuerySelectorAll ---
'   ALL_MAX_DEFAULT    : default upper limit returned at once. *Not cut silently* (LastAllTruncated)
'   ELEMENT_WARN_COUNT : warn when the registry exceeds this (does not discard)
Private Const ALL_MAX_DEFAULT    As Long = 200
Private Const ELEMENT_WARN_COUNT As Long = 5000
Private m_allTruncated As Boolean


' --- D-4b: settings for quiescence waits ---
'   m_ignoreNet / m_ignoreSel : exclusion lists *held on the VBA side* (the page side vanishes on navigation).
'   m_autoWait   : whether to WaitSettled automatically after an action. *Default False*
'   m_lastSettle : breakdown of the previous wait (for diagnosis).
Private m_ignoreNet  As Collection
Private m_ignoreSel  As String
Private m_autoWait   As Boolean
Private m_lastSettle As String

' Default quiet window (ms). Same value as the reference product (WebDriver BiDi for SeleniumVBA).
Private Const SETTLE_STABLE_MS_DEFAULT As Long = 500

' Default for the lingering check (ms). *Requests in flight for more than 10 seconds are not counted*
Private Const STALE_INFLIGHT_MS_DEFAULT As Long = 10000


' --- D-4c: explicit signal (arm) ---
'   *The VBA side also remembers whether it is armed* If only the page side holds it, when the probe vanishes
'   on navigation we cannot notice "it was armed but vanished" (it would silently
'   turn into a wait with no condition). Hold it on both, and if they disagree fail with signal-lost.
Private m_sigDom As Boolean
Private m_sigNet As Boolean
Private m_sigUrl As Boolean

' --- D-4f: exclude lingering requests from the quiescence check ---
'   *Long polling and streaming exist on every site*
'   Google Maps search keeps /search?tbm=map open and streams results,
'   so counting it means quiescence never comes (measured in D-4d).
'   Requests in flight beyond a certain time are regarded as "things that just keep streaming" and not counted.
'   Setting 0 disables it (count everything).
Private m_staleMs As Long


' --- N-2: request detail (headers and body) ---
'   *A ring separate from the list* (issue 5, option S). The list is for "order and the overall picture",
'   the detail is for "the one request I want to reproduce"; the uses differ.
'
'   m_detOn        : whether to take the detail. *Default False* (issue 8, option kappa)
'   m_detCap       : capacity of the detail ring (default 50)
'   m_detHead      : next write position (0-based)
'   m_detCount     : number of entries currently in the ring
'   m_detDropped   : *number dropped on overflow* not cut silently, same as the list
'   m_detBodyMax   : upper limit for reading the body (bytes. Default 65536)
'   m_detRedact    : whether to hide secret headers. *Default True* (issue 3, option M)
'   m_detRedactList: names to hide, lined up as |lowercase|
'   m_detUriFilter : substring match of URLs to take detail for (all if empty)
Private m_detOn         As Boolean
Private m_detCap        As Long
Private m_detHead       As Long
Private m_detCount      As Long
Private m_detDropped    As Long
Private m_detBodyMax    As Long
Private m_detRedact     As Boolean
Private m_detRedactList As String
Private m_detUriFilter  As String
Private m_detSeq()     As Long
Private m_detMethod()  As String
Private m_detUri()     As String
Private m_detHeaders() As String
Private m_detCType()   As String
Private m_detBytes()   As Long
Private m_detTrunc()   As Boolean
Private m_detIsText()  As Boolean
Private m_detBody()    As String
Private m_detHex()     As String
Private m_detRespHeaders() As String   ' N-3: response headers
Private m_detStatus()      As Long     ' N-3: status (0 = not arrived)
Private m_detReason()      As String   ' N-3: reason phrase
Private m_detRespBody()    As String   ' N-4: response body (text only)
Private m_detRespBytes()   As Long     ' N-4: number of body bytes
Private m_detRespTrunc()   As Boolean  ' N-4: whether it was cut at the limit
Private m_detRespIsText()  As Boolean  ' N-4: whether it was judged to be text
Private m_detRespEnc()     As String   ' N-4: content-encoding (*not assumed*)
Private m_detRespHex()     As String   ' N-4: *hex of the first 32 bytes (always kept)*
Private m_detRespGot()     As Boolean  ' N-4: *whether the body arrived* (0 bytes differs from not arrived)

' Default capacity and upper limit of the detail ring. *An order of magnitude smaller than the list* The weight per entry differs.
Private Const NET_DET_CAP_DEFAULT As Long = 50
Private Const NET_DET_CAP_MAX     As Long = 2000

' Upper limit of the body (default 64KB) and the upper limit of that limit (8MB).
'   *It is read synchronously inside the handler* so it is not left unbounded.
Private Const NET_BODY_MAX_DEFAULT As Long = 65536
Private Const NET_BODY_MAX_LIMIT   As Long = 8388608
Private Const NET_BODY_CHUNK       As Long = 16384

' Upper limit on the number of headers read. For one request, COM calls run 3 times x the count.
Private Const NET_HDR_MAX As Long = 40

' Number of leading bytes shown when presenting a binary body.
Private Const NET_HEX_HEAD As Long = 32

' dwOrigin of IStream::Seek (STREAM_SEEK_SET in Windows SDK objidl.h).
Private Const STREAM_SEEK_SET As Long = 0


' --- N-1: capturing network requests (WebResourceRequested) ---
'   m_netOn      : whether capturing (whether add_WebResourceRequested is done)
'   m_netCap     : capacity of the ring (entries). Can be changed with NetLogCapacity.
'   m_netHead    : next write position (0-based)
'   m_netCount   : number of entries currently in the ring (0 to m_netCap)
'   m_netTotal   : *total fired count* since capture started (overflowed ones are counted too)
'   m_netDropped : *number dropped on overflow* not cut silently
'   m_netT0      : GetTickCount value at capture start. The baseline for each entry's elapsed ms.
'   m_netFilters : registered filters (uri & vbTab & ctx). Remembered so they can be removed on Stop.
'   m_netStatus  : status of that row (N-3). *0 = no response has arrived yet*
'   m_netReason  : reason phrase (N-3)
'   m_netRespOn  : whether response events are being picked up (N-3)
'   m_netUnmatched : *number of responses that had no counterpart in the list* (N-3. Mostly normal)
Private m_netOn      As Boolean
Private m_netCap     As Long
Private m_netHead    As Long
Private m_netCount   As Long
Private m_netTotal   As Long
Private m_netDropped As Long
Private m_netT0      As Long
Private m_netRespOn    As Boolean
Private m_netUnmatched As Long

' --- N-4: response body ---
'   m_respBodyOn     : whether to take the body. *Default False* (issue 1, option A)
'   m_respBodyMax    : upper limit of the body (bytes. Default 65536)
'   m_respPending    : *ResponseView held until completion* (issue 5, option beta)
'                      The key is "r" & a serial number. The value is an AddRef'd pointer.
'   m_respBodyOrphan : number of bodies that had no place to go (not silently discarded)
Private m_respBodyOn     As Boolean
Private m_respBodyMax    As Long
Private m_respPending    As Collection
Private m_respBodyOrphan As Long

' Default upper limit of the response body (64KB). The limit on the limit shares NET_BODY_MAX_LIMIT.
Private Const NET_RESP_BODY_MAX_DEFAULT As Long = 65536

' *For 204 No Content, the completion of GetContent arrives with this HRESULT* (measured in N-4).
'   0x800700E8 = HRESULT_FROM_WIN32(ERROR_NO_DATA). *This is not a failure but means
'   "there is no body"*, so it must not be treated the same as a fetch failure
'   (design principle 111: do not leave behind a plausible-looking value).
Private Const E_NO_DATA As Long = &H800700E8
Private m_netFilters As Collection
Private m_netStatus()  As Long
Private m_netReason()  As String
Private m_netSeq()    As Long
Private m_netMs()     As Long
Private m_netCtx()    As Long
Private m_netMethod() As String
Private m_netUri()    As String

' Default capacity and upper limit of the ring. The default of 500 is a guide to hold "one page's worth of operations" entirely.
Private Const NET_CAP_DEFAULT As Long = 500
Private Const NET_CAP_MAX     As Long = 50000

' Kinds captured by default (issue 4). Images, CSS, fonts and scripts are excluded.
'   1 = DOCUMENT (form POST / page navigation) / 7 = XML_HTTP_REQUEST / 8 = FETCH
Private Const NET_CTX_DEFAULT As String = "DOCUMENT,XHR,FETCH"


' --- D-2: generation cache of the element registry ---
'   m_domGen : window.__wv2.gen last received from JS. For the fast path only.
'              An empty string if no element search has been done yet.
'              *Not the authority for the verdict* (the authority is always the JS side. See the header above)
Private m_domGen As String


' ============================================================
' Class_Initialize / Class_Terminate
' ============================================================
Private Sub Class_Initialize()
    m_staleMs = STALE_INFLIGHT_MS_DEFAULT
    m_state = Cs_Created
    m_pCtrl = 0
    m_pView = 0
    m_pView2 = 0    ' Added in 9.9a (cache of the derived IF, not-yet-acquired state)
    m_pView3 = 0    ' Added in 9.10d-2a (ICoreWebView2_3 derived IF, not-yet-acquired state)
    Set m_browser = Nothing   ' 9.10e (back reference, set by SetBrowser at AddTab time)
    m_lastError = 0
    m_pageLoaded = False
    m_executeScriptCounter = 0
    m_closed = False
    m_title = ""                  ' 9.11c (title cache)
    Set m_webMsgSink = Nothing    ' 9.11c (WebMessage sink, not registered)
    m_isSettings = False          ' 9.19 (default is a normal tab)

    ' --- D-1: initialization related to EvalSync ---
    Set m_evalPending = New Collection
    m_lastEvalOk = False
    m_lastEvalError = ""
    m_inCallback = 0
    m_evalDepth = 0

    ' --- D-2: generation cache of the element registry ---
    m_domGen = ""

    ' --- N-1: ring buffer of network requests ---
    m_netOn = False
    m_netCap = NET_CAP_DEFAULT
    Set m_netFilters = New Collection
    NetRingAlloc

    ' --- N-2: detail (headers and body) ---
    m_detOn = False                       ' *Default OFF* (issue 8, option kappa)
    m_detCap = NET_DET_CAP_DEFAULT
    m_detBodyMax = NET_BODY_MAX_DEFAULT
    m_detRedact = True                    ' *Default is to hide* (issue 3, option M)
    m_detRedactList = NetRedactDefaults()
    m_detUriFilter = ""
    NetDetailRingAlloc

    ' --- N-4: response body ---
    m_respBodyOn = False                  ' *Default OFF* (issue 1, option A)
    m_respBodyMax = NET_RESP_BODY_MAX_DEFAULT
    Set m_respPending = New Collection
    m_respBodyOrphan = 0
End Sub

Private Sub Class_Terminate()
    ' COM ownership convention: the received pointer had its VBA-side share of ownership
    ' secured with ComAddRef, so Release it here.
    '
    ' Order (m_pSettings added in 9.8b, m_pView2 added in 9.9a):
    '   1. Unregister persistent event handlers (m_pView must still be alive)
    '   2. Release m_pSettings (Settings is a child object of View)
    '   3. Release m_pView2  <- added in 9.9a
    '                            (View2 is a derived IF obtained from View by QI,
    '                             and has an independent reference count, so before View)
    '   4. Release m_pView
    '   5. Release m_pCtrl
    '
    ' Why this order:
    '   - remove_*Event is a method of m_pView (= ICoreWebView2), so
    '     it must be fired while m_pView is alive
    '   - Settings is a child object that View created. Release it before View
    '   - View2 is a derived IF obtained from View by QueryInterface. It is merely the same COM
    '     object seen through a different IF, but QI fires an independent AddRef,
    '     so View2 and View each have an independent reference count.
    '     Following design principle 43 (child to parent), release View2 first
    '   - View is a child object that Ctrl created. Releasing the Controller side first
    '     would not break WebView2's internal reference-count management, but
    '     to make the lifetime dependencies (Ctrl is the parent, View the child, Settings the grandchild)
    '     explicit at the VBA level too, children are released first.

    ' --- K-4a: *entry point for measuring elapsed time* ---
    '   From here to the next line (the first Cleanup) is the elapsed time of remove_*Event.
    Wv2Log.LogD "Wv2Pane.Terminate: *start* (pCtrl=" & m_pCtrl & _
                " pView=" & m_pView & ")"

    ' --- 1. Unregister all persistent event handlers at once (added in stage 9.4a) ---
    CleanupAllPersistentHandlers
    Wv2Log.LogD "Wv2Pane.Terminate: finished unregistering persistent handlers"

    ' --- 2. Release m_pSettings (added in stage 9.8b) ---
    If m_pSettings <> 0 Then
        Wv2Log.LogD "Wv2Pane.Terminate: Releasing m_pSettings = " & m_pSettings
        ComRelease m_pSettings
        m_pSettings = 0
    End If

    ' --- 3-pre. Release m_pView3 (added in stage 9.10d-2a) ---
    '   View3 is the most derived IF (ICoreWebView2_3) obtained from View by QI.
    '   To release from the derived side, Release it before m_pView2 (design principle 51).
    If m_pView3 <> 0 Then
        Wv2Log.LogD "Wv2Pane.Terminate: Releasing m_pView3 = " & m_pView3
        ComRelease m_pView3
        m_pView3 = 0
    End If

    ' --- 3. Release m_pView2 (added in stage 9.9a) ---
    '   View2 is a derived IF obtained from View by QueryInterface. Release it before View.
    If m_pView2 <> 0 Then
        Wv2Log.LogD "Wv2Pane.Terminate: Releasing m_pView2 = " & m_pView2
        ComRelease m_pView2
        m_pView2 = 0
    End If

    ' --- 4. Release m_pView ---
    If m_pView <> 0 Then
        Wv2Log.LogD "Wv2Pane.Terminate: Releasing m_pView = " & m_pView
        ComRelease m_pView
        m_pView = 0
    End If

    ' --- 5. Release m_pCtrl ---
    If m_pCtrl <> 0 Then
        Wv2Log.LogD "Wv2Pane.Terminate: Releasing m_pCtrl = " & m_pCtrl
        ComRelease m_pCtrl
        m_pCtrl = 0
    End If

    ' m_handler should normally already be Nothing by OnControllerCompleted,
    ' but if it remains, e.g. because an exception occurred in Init, cut it here.
    If Not (m_handler Is Nothing) Then
        Wv2Log.LogD "Wv2Pane.Terminate: m_handler remains, so cutting it (abnormal-termination path)"
        Set m_handler = Nothing
    End If

    Wv2Log.LogD "Wv2Pane.Terminate: *done*"   ' K-4a
End Sub


' ============================================================
' CleanupAllPersistentHandlers (added in stage 9.4a, internal helper)
'
'   For every persistent handler registered in m_persistentHandlers,
'   issue an unregister request to the WebView2 side. Called from Class_Terminate.
'
'   For each HandlerKind, the corresponding remove_ vtable index is used.
'   If unregistering succeeds, WebView2 Releases the ComCallbackHandler -> refcount=0 ->
'   on the auto-release path ComCallbackHandler.ClearOwner; it is cleanest to make m_persistentHandlers(k)
'   Nothing at the point ClearOwner has run.
'
'   However, there is no problem of "if m_persistentHandlers(k) is made Nothing first,
'   the m_owner that ClearOwner touches is already Nothing"
'   (ClearOwner only sets ComCallbackHandler.m_owner to Nothing, and does not touch
'   m_persistentHandlers on the Wv2Pane side).
'
'   Summary of the order:
'     1. Ask WebView2 for View_RemoveEvent (remove_*Event)
'     2. WebView2 Releases the ComCallbackHandler -> the foundation's auto-release path ->
'        ComCallbackHandler.ClearOwner makes ComCallbackHandler.m_owner = Nothing
'     3. Here, make m_persistentHandlers(k) = Nothing on the Wv2Pane side
'
'   In 9.6a the array upper limit was extended to HK_NewWindowRequested, but
'   the loop itself uses LBound/UBound, so no code change is needed
'   (= the true value of design principle 15 "automate horizontal rollout with a HandlerKind-linked loop").
' ============================================================
Private Sub CleanupAllPersistentHandlers()
    If m_pView = 0 Then
        ' If View has already been released, there is no means to fire remove_.
        ' The persistent handlers may still be alive, but nothing can be done (leak accepted).
        Dim k1 As Long
        For k1 = LBound(m_persistentHandlers) To UBound(m_persistentHandlers)
            Set m_persistentHandlers(k1) = Nothing
            m_persistentTokens(k1) = 0^
        Next k1
        Exit Sub
    End If

    Dim k As Long
    For k = LBound(m_persistentHandlers) To UBound(m_persistentHandlers)
        If Not (m_persistentHandlers(k) Is Nothing) Then
            Dim removeIndex As Long
            removeIndex = RemoveVtableIndexFor(CLng(k))
            If removeIndex > 0 Then
                ' 9.9b: decide from kind which interface to fire remove on.
                '   DOMContentLoaded (_2 IF) uses m_pView2, everything else m_pView.
                '   Faithful to design principle 45 (call a derived IF's method through the derived IF's pointer).
                Dim pIfRemove As LongPtr
                pIfRemove = RemoveInterfaceFor(CLng(k))
                If pIfRemove <> 0 Then
                    Wv2Log.LogD "Wv2Pane.Cleanup: removing persistent handler kind=" & k & _
                                " vtable=" & removeIndex & " token=" & m_persistentTokens(k) & _
                                " iface=" & pIfRemove
                    dcf pIfRemove, removeIndex, "View.remove_Event", m_persistentTokens(k)
                Else
                    ' The interface needed for unregistering (e.g. m_pView2) is not acquired / already released.
                    ' There is no means to fire remove, so only cut the reference (leak accepted).
                    Wv2Log.LogD "Wv2Pane.Cleanup: kind=" & k & " has remove IF = 0, " & _
                                "skipping remove (releasing the reference only)"
                End If
            End If
            Set m_persistentHandlers(k) = Nothing
            m_persistentTokens(k) = 0^
        End If
    Next k
End Sub


' ============================================================
' RemoveInterfaceFor (added in stage 9.9b, internal helper)
'
'   Returns the pointer of the interface on which the remove_*Event corresponding to
'   a HandlerKind should be fired. Whereas RemoveVtableIndexFor returns the "vtable index",
'   this returns "through which interface (m_pView / m_pView2 / ...) to call".
'
'   Design principle 45 (call a derived IF's method through the derived IF's pointer):
'     DOMContentLoaded is an event of ICoreWebView2_2 (vtable 65), so
'     remove it through m_pView2. The others are events of the base ICoreWebView2, so
'     remove through m_pView.
'
'     The observation in 9.9a found that "m_pView2 == m_pView (the same value)",
'     but relying on that to hit vtable 65 through m_pView is a dangerous shortcut of
'     "calling a derived IF's method through a base pointer" (= rooted in the same thinking that caused the crash in 9.8c).
'     Using the semantically correct pointer also withstands future SDK changes.
'
'   Return value:
'     The pointer of the corresponding interface.
'     If 0 is returned, it means "there is no means to fire remove" (e.g. DOMContentLoaded but
'     m_pView2 not acquired / already released), so the caller skips remove.
'
'   Future extension:
'     When events of ICoreWebView2_3 and later increase, just add one Case line here
'     to roll it out horizontally (keeping symmetry with RemoveVtableIndexFor).
' ============================================================
Private Function RemoveInterfaceFor(ByVal kind As Long) As LongPtr
    Select Case kind
        Case HK_DOMContentLoaded, HK_WebResourceResponseReceived
            ' Events of the _2 IF. Call through m_pView2 (returns 0 if not acquired).
            '   *N-3's WebResourceResponseReceived is also on the _2 IF* so same handling.
            RemoveInterfaceFor = m_pView2
        Case Else
            ' Events of the base ICoreWebView2. Call through m_pView.
            RemoveInterfaceFor = m_pView
    End Select
End Function


' ============================================================
' RemoveVtableIndexFor (added in stage 9.4a, internal helper)
'
'   Returns the vtable index of the remove_*Event corresponding to a HandlerKind.
'   Paired with add_*Event; the correspondence is fixed by the SDK.
'
'   ICoreWebView2 vtable layout (base):
'     0  QueryInterface
'     1  AddRef
'     2  Release
'     3  get_Settings
'     ...
'     5  Navigate
'     ...
'     7  add_NavigationStarting       8  remove_NavigationStarting
'     ...
'    13  add_HistoryChanged          14  remove_HistoryChanged       <- added in 9.9b
'    15  add_NavigationCompleted     16  remove_NavigationCompleted
'     ...
'    34  add_WebMessageReceived      35  remove_WebMessageReceived
'     ...
'    44  add_NewWindowRequested      45  remove_NewWindowRequested  <- added in 9.6a
'    46  add_DocumentTitleChanged    47  remove_DocumentTitleChanged
'    55  add_WebResourceRequested    56  remove_WebResourceRequested  <- added in N-1
'    (_2 IF) 61  add_WebResourceResponseReceived  62  remove_  <- added in N-3
'
'   ICoreWebView2_2 vtable layout (derived IF following the base's 61 entries):
'    64  add_DOMContentLoaded        65  remove_DOMContentLoaded     <- added in 9.9b
'       * However the vtable indexes of add/remove are offsets on m_pView2 (ICoreWebView2_2).
'         RemoveInterfaceFor returns m_pView2 for DOMContentLoaded.
'
'   In 9.6a, HK_NewWindowRequested = 45 was added.
'   In 9.9b, HK_HistoryChanged = 14 and HK_DOMContentLoaded = 65 were added.
' ============================================================
Private Function RemoveVtableIndexFor(ByVal kind As Long) As Long
    Select Case kind
        Case HK_NavigationStarting:   RemoveVtableIndexFor = 8
        Case HK_NavigationCompleted:  RemoveVtableIndexFor = 16
        Case HK_WebMessageReceived:   RemoveVtableIndexFor = 35
        Case HK_DocumentTitleChanged: RemoveVtableIndexFor = 47
        Case HK_NewWindowRequested:   RemoveVtableIndexFor = 45   ' Added in 9.6a
        Case HK_HistoryChanged:       RemoveVtableIndexFor = 14   ' Added in 9.9b (base IF)
        Case HK_DOMContentLoaded:     RemoveVtableIndexFor = 65   ' Added in 9.9b (_2 IF)
        Case HK_WebResourceRequested: RemoveVtableIndexFor = 56   ' Added in N-1 (base IF)
        Case HK_WebResourceResponseReceived: RemoveVtableIndexFor = 62  ' Added in N-3 (_2 IF)
        Case Else:                    RemoveVtableIndexFor = 0   ' unknown -> cannot unregister
    End Select
End Function


' ============================================================
' Init
'
'   Starts creating the ICoreWebView2Controller.
'   This method returns immediately without waiting for the callback (asynchronous).
'   Completion is notified via OnControllerCompleted.
'
'   The caller waits with DoEvents until ctrl.IsReady becomes True.
'   However WebView2Loader may call back synchronously, in which case
'   it is already Ready / Failed at the point of returning from Init.
' ============================================================
Public Sub Init(ByVal env As Wv2Environment, ByVal hwnd As LongPtr)

    If m_state <> Cs_Created Then
        MsgBox "Wv2Pane.Init cannot be called twice", vbCritical
        Exit Sub
    End If

    If env Is Nothing Then
        MsgBox "Wv2Pane.Init: env is Nothing", vbCritical
        m_state = Cs_Failed
        Exit Sub
    End If

    If Not env.IsReady Then
        MsgBox "Wv2Pane.Init: env is not Ready yet (state=" & env.state & ")", vbCritical
        m_state = Cs_Failed
        Exit Sub
    End If

    Dim pEnv As LongPtr
    pEnv = env.EnvironmentPtr
    If pEnv = 0 Then
        MsgBox "Wv2Pane.Init: env.EnvironmentPtr is 0", vbCritical
        m_state = Cs_Failed
        Exit Sub
    End If

    ' --- 1. Acquire one handler (ComCallbackHandler) ---
    Set m_handler = AcquireHandlerFor(HK_ControllerCompleted, Me)
    If m_handler Is Nothing Then
        MsgBox "Wv2Pane.Init: failed to Acquire the handler", vbCritical
        m_state = Cs_Failed
        Exit Sub
    End If

    ' --- 2. Compute the obj address that WebView2 holds ---
    Dim pVTableObj As LongPtr
    pVTableObj = m_handler.Slot + VTABLE_OBJ_OFFSET

    ' --- 3. Set m_state = Cs_Waiting BEFORE the API call ---
    '   (pattern established in 9.3a: supports synchronous callbacks)
    m_state = Cs_Waiting

    ' --- 4. Call Environment->CreateCoreWebView2Controller(hwnd, handler) ---
    '
    '   ICoreWebView2Environment vtable layout:
    '     0 : QueryInterface
    '     1 : AddRef
    '     2 : Release
    '     3 : CreateCoreWebView2Controller(parentWindow HWND, handler ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*)
    '     4 : ... (after this it depends on the SDK version)
    '
    '   Arguments:
    '     hwnd       : HWND of the Frame control
    '     pVTableObj : vtableObj address of the ComCallbackHandler handler
    Debug.Print "Wv2Pane.Init: dcf(pEnv=" & pEnv & ", 3, ""Env.CreateController"", hwnd=" & _
                hwnd & ", handler=" & pVTableObj & ")"

    Dim hr As Long
    hr = dcf(pEnv, 3, "Env.CreateCoreWebView2Controller", hwnd, pVTableObj)

    If hr <> 0 Then
        MsgBox "CreateCoreWebView2Controller failed. HRESULT: &H" & Hex(hr), vbCritical
        Wv2Thunks.Thunks_DiscardHandler m_handler   ' English fork: return the slot
        Set m_handler = Nothing
        m_state = Cs_Failed
        m_lastError = hr
        ' Leak countermeasures on failure are not implemented in 9.3a / 9.3b (same policy as Wv2Environment).
        ' Relies on the forced cleanup of all slots in Thunks_Shutdown.
        Exit Sub
    End If
End Sub


' ============================================================
' OnControllerCompleted
'
'   From the HK_ControllerCompleted branch of ComCallbackHandler.Handler_Invoke,
'   m_owner.OnControllerCompleted is called via dynamic binding.
'
'   Exactly the same form as Wv2Environment.OnEnvironmentCompleted. It must be a Public Sub
'   (a Friend does not ride on IDispatch).
'
'   Note (9.6a naming convention): one-shot callbacks keep names without an origin prefix.
'   There is no function that directly receives a derived IF (ICoreWebView2Environment2 etc.),
'   so there is no need to identify a version (confirmed with tarboh, stage 9.6a).
'
'   Arguments:
'     errorCode   : HRESULT (S_OK = 0 means success)
'     pController : ICoreWebView2Controller pointer
'                   (passed in a state where WebView2 has temporarily AddRef'd it.
'                    If you keep it, you must AddRef on the VBA side.)
' ============================================================
Public Sub OnControllerCompleted( _
    ByVal errorCode As Long, _
    ByVal pController As LongPtr)

    Debug.Print "Wv2Pane.OnControllerCompleted: errorCode=" & errorCode & _
                ", pController=" & pController

    m_lastError = errorCode

    If errorCode = 0 And pController <> 0 Then
        ' Secure the reference for the share held on the VBA side
        m_pCtrl = pController
        ComAddRef m_pCtrl
        m_state = Cs_Ready
    Else
        m_pCtrl = 0
        m_state = Cs_Failed
    End If

    ' One-shot handler pattern (option D1-alpha): cut the reference to the handler
    Set m_handler = Nothing
End Sub


' ============================================================
' Property group
' ============================================================
Public Property Get state() As ControllerState
    state = m_state
End Property

Public Property Get IsReady() As Boolean
    IsReady = (m_state = Cs_Ready)
End Property

Public Property Get IsFailed() As Boolean
    IsFailed = (m_state = Cs_Failed)
End Property

Public Property Get LastError() As Long
    LastError = m_lastError
End Property

' --- Exposing raw pointers (for internal observation / debugging) ---
Public Property Get CtrlPtr() As LongPtr
    CtrlPtr = m_pCtrl
End Property

Public Property Get ViewPtr() As LongPtr
    ViewPtr = m_pView
End Property

' --- IsPageLoaded (added in 9.4c) ---
'   Set to True in View_OnNavigationCompleted.
'   Used for synchronization by external code (UserForm etc.) waiting for page load completion.
'   When SPA support is added in the future, it is planned to be moved to the comprehensive class side.
Public Property Get IsPageLoaded() As Boolean
    IsPageLoaded = m_pageLoaded
End Property

' --- LastNavigationOk (English fork) ---
'   IsSuccess of the most recent NavigationCompleted. Read it after IsPageLoaded
'   turned True: IsPageLoaded alone is also True for a failed load (DNS error,
'   refused connection, cancelled). An HTTP error page (404, 500) is a success here.
Public Property Get LastNavigationOk() As Boolean
    LastNavigationOk = m_lastNavOk
End Property

' --- DocumentTitle (added in 9.11c, read entry point of the title cache) ---
'   Returns the latest title that View_OnDocumentTitleChanged saved.
'   TabBar's PushSyncToJs reads it with Browser.PaneAt(i).DocumentTitle.
'   A pane whose title has not been obtained yet returns an empty string.
Public Property Get DocumentTitle() As String
    DocumentTitle = m_title
End Property


' --- IsSettings (added in 9.19, settings-tab identification flag) ---
'   Whether this Pane was opened as the settings tab. Wv2Browser.OpenSettingsTab
'   sets True at creation, and it is referenced by the duplicate check when [gear] is clicked again.
Public Property Get IsSettings() As Boolean
    IsSettings = m_isSettings
End Property

Public Property Let IsSettings(ByVal value As Boolean)
    m_isSettings = value
End Property


' ============================================================
' Ctrl_PutBounds (ICoreWebView2Controller::put_Bounds, vtable index 6)
'
'   Pass a RECT corresponding to the HWND client area of a Frame control etc.
'   to set the display rectangle of WebView2.
'
'   The bounds argument is a RECT obtained with GetClientRect. The coordinate system starts at (0, 0)
'   inside the client area of the parent HWND.
'
'   Under the x64 ABI a 16-byte structure is passed by value, but as inherited from the old version,
'   passing a pointer with VarPtr(rc) is interpreted correctly by WebView2
'   (verified on a real machine). Pass it to dcf as a LongPtr.
'
'   Return value: HRESULT (S_OK = 0 means success)
' ============================================================
Public Function Ctrl_PutBounds(ByRef bounds As RECT) As Long
    If m_pCtrl = 0 Then
        Ctrl_PutBounds = &H80004003   ' E_POINTER
        Exit Function
    End If
    Ctrl_PutBounds = dcf(m_pCtrl, 6, "Ctrl.put_Bounds", VarPtr(bounds))
End Function


' ============================================================
' Ctrl_PutIsVisible (ICoreWebView2Controller::put_IsVisible, vtable index 4)
'   Added in stage 9.10c-1.
'
'   Switches the Controller between visible and invisible. Used by Wv2Browser's tab switching
'   (ActivateTab) to realize "only the active tab is visible, the others invisible".
'
'   Argument: visible (True = visible / False = invisible)
'   Return value: HRESULT (S_OK = 0 means success)
'
'   Implementation notes:
'     put_IsVisible is a simple-argument method taking one BOOL (integer).
'     It is not a handler family that passes a COM object, so a simple-argument model
'     such as Ctrl_PutBounds (calling dcf directly) is the correct model.
'     VARIANT_BOOL / Win32 BOOL distinction: the argument of ICoreWebView2Controller::put_IsVisible
'     is a Win32 BOOL (= 4-byte int, True=1/False=0). VBA's True is -1, so
'     normalize to 1 / 0 before passing to dcf (spec fact 11).
' ============================================================
Public Function Ctrl_PutIsVisible(ByVal visible As Boolean) As Long
    If m_pCtrl = 0 Then
        Ctrl_PutIsVisible = &H80004003   ' E_POINTER
        Exit Function
    End If
    Dim b As Long
    If visible Then b = 1 Else b = 0
    Ctrl_PutIsVisible = dcf(m_pCtrl, 4, "Ctrl.put_IsVisible", b)
End Function


' ============================================================
' Ctrl_GetCoreWebView2 (ICoreWebView2Controller::get_CoreWebView2, vtable index 25)
'
'   Gets the ICoreWebView2 pointer as an out argument and stores it in m_pView.
'
'   COM convention: get_xxx-family methods return the obtained pointer already
'   AddRef'd (= the receiver does not need to AddRef again, only Release).
'   m_pView is Released in Class_Terminate.
'
'   Return value: HRESULT (S_OK = 0 means success)
'
'   Note: this method is assumed to be called only once in the Init chain.
'         If m_pView is already non-zero, to prevent double acquisition
'         the existing pointer is Released first (it should originally be called only
'         at Init time, but for safety).
' ============================================================
Public Function Ctrl_GetCoreWebView2() As Long
    If m_pCtrl = 0 Then
        Ctrl_GetCoreWebView2 = &H80004003   ' E_POINTER
        Exit Function
    End If

    ' If there is an existing m_pView, Release it first (prevents double acquisition)
    If m_pView <> 0 Then
        ComRelease m_pView
        m_pView = 0
    End If

    Dim pv As LongPtr
    Dim hr As Long
    hr = dcf(m_pCtrl, 25, "Ctrl.get_CoreWebView2", VarPtr(pv))
    If hr = 0 Then
        m_pView = pv
    End If
    Ctrl_GetCoreWebView2 = hr
End Function


' ============================================================
' View_Navigate (ICoreWebView2::Navigate, vtable index 5)
'
'   Navigates to the specified URL. Navigate itself only returns true synchronously, and
'   the actual completion of the page load is received with HK_NavigationCompleted.
'
'   The URL argument is a VBA String (= UTF-16 LE), and StrPtr gives the LPCWSTR.
'
'   Return value: HRESULT (S_OK = 0 means success)
' ============================================================
Public Function View_Navigate(ByVal url As String) As Long
    If m_pView = 0 Then
        View_Navigate = &H80004003   ' E_POINTER
        Exit Function
    End If
    ' Reset the flag at the start of a new navigation
    m_pageLoaded = False
    View_Navigate = dcf(m_pView, 5, "View.Navigate", StrPtr(url))
End Function


' ============================================================
' View_AddNavigationStarting (new in stage 9.4a, renamed with the View_ prefix in 9.6a)
'
'   Calls ICoreWebView2::add_NavigationStarting (vtable index 7) to
'   register a persistent event handler.
'
'   Flow:
'     1. Acquire a ComCallbackHandler (HK_NavigationStarting) (owner = Me)
'     2. Ask WebView2 to register with dcf(m_pView, 7, handler, &token)
'     3. WebView2 fires AddRef (the ComCallbackHandler's refcount goes up)
'     4. Save the obtained token in m_persistentTokens(HK_NavigationStarting)
'     5. Save the ComCallbackHandler in m_persistentHandlers(HK_NavigationStarting)
'
'   After the call, on link clicks on Bing etc.,
'   ComCallbackHandler.Invoke[HK_NavigationStarting] -> Wv2Pane.View_OnNavigationStarting
'   events can be received via this path.
'
'   If already registered, do nothing (prevents double registration).
'
'   Return value: HRESULT (S_OK = 0 means success)
' ============================================================
Public Function View_AddNavigationStarting() As Long
    If m_pView = 0 Then
        View_AddNavigationStarting = &H80004003   ' E_POINTER
        Exit Function
    End If

    ' If already registered, do nothing
    If Not (m_persistentHandlers(HK_NavigationStarting) Is Nothing) Then
        Debug.Print "Wv2Pane.View_AddNavigationStarting: already registered, skipping"
        View_AddNavigationStarting = 0
        Exit Function
    End If

    ' Acquire a ComCallbackHandler (owner = Me; receives events via dynamic binding)
    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_NavigationStarting, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_AddNavigationStarting: failed to Acquire the handler"
        View_AddNavigationStarting = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If

    ' Fire add_NavigationStarting (vtable index 7)
    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As LongLong
    Dim hr As Long
    hr = dcf(m_pView, 7, "View.add_NavigationStarting", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_AddNavigationStarting: add_NavigationStarting failed hr=&H" & Hex(hr)
        ' Release the ComCallbackHandler (refcount=0 goes to the auto-release path)
        ' h is referenced as a local variable, so just doing Set h = Nothing
        ' leaves the reference from the foundation's m_handlers(idx).
        ' WebView2 may not have been able to fire AddRef, but the ComCallbackHandler slot is handled by
        ' the sentinel mechanism of stage 9.5, so no individual handling is needed here.
        Wv2Thunks.Thunks_DiscardHandler h   ' English fork: return the slot (see Thunks_DiscardHandler)
        Set h = Nothing
        View_AddNavigationStarting = hr
        Exit Function
    End If

    ' Success: save the token and the ComCallbackHandler
    m_persistentTokens(HK_NavigationStarting) = token
    Set m_persistentHandlers(HK_NavigationStarting) = h
    Debug.Print "Wv2Pane.View_AddNavigationStarting OK: token=" & token & ", handler.slot=" & h.Slot

    View_AddNavigationStarting = 0
End Function


' ============================================================
' View_RemoveNavigationStarting (new in stage 9.4a, renamed in 9.6a)
'
'   Unregisters the registered NavigationStarting handler.
'   Call it when the user wants to unregister explicitly. From Class_Terminate,
'   CleanupAllPersistentHandlers does it all at once, so there is no need
'   to call this directly.
'
'   Flow:
'     1. Unregister from the WebView2 side with View_RemoveEvent (vtable index 8)
'     2. WebView2 Releases the ComCallbackHandler -> the foundation's auto-release path
'     3. On the Wv2Pane side, m_persistentHandlers(k) = Nothing, token = 0
'
'   Return value: HRESULT (S_OK = 0 means success; if not registered, S_OK and nothing is done)
' ============================================================
Public Function View_RemoveNavigationStarting() As Long
    If m_pView = 0 Then
        View_RemoveNavigationStarting = &H80004003   ' E_POINTER
        Exit Function
    End If

    If m_persistentHandlers(HK_NavigationStarting) Is Nothing Then
        View_RemoveNavigationStarting = 0   ' not registered is a no-op
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView, 8, "View.remove_NavigationStarting", _
             m_persistentTokens(HK_NavigationStarting))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_RemoveNavigationStarting: unregister failed hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Pane.View_RemoveNavigationStarting OK"
    End If

    ' Cut the Wv2Pane-side reference whether it succeeded or failed
    Set m_persistentHandlers(HK_NavigationStarting) = Nothing
    m_persistentTokens(HK_NavigationStarting) = 0^

    View_RemoveNavigationStarting = hr
End Function


' ============================================================
' View_OnNavigationStarting (new in stage 9.4a, renamed in 9.6a)
'   Called from ComCallbackHandler via dynamic binding.
'
'   From the HK_NavigationStarting branch of ComCallbackHandler.Handler_Invoke,
'   m_owner.View_OnNavigationStarting is called via dynamic binding.
'
'   Public Sub is required (a Friend does not ride on IDispatch::Invoke).
'
'   Arguments:
'     pSender : Invoke's 1st argument from WebView2's viewpoint = ICoreWebView2 pointer
'               (should be the same as m_pView, but received in a separate variable just in case)
'     pArgs   : Invoke's 2nd argument from WebView2's viewpoint =
'               ICoreWebView2NavigationStartingEventArgs pointer
'
'   Behaviour:
'     Calls args.get_Uri (vtable index 3) and writes the URL to the log.
'     get_Uri returns an LPWSTR* as an out argument, so GetStringProperty handles
'     everything through PtrToString + CoTaskMemFree.
'
'   Reference count of args:
'     WebView2 calls Invoke with args temporarily AddRef'd.
'     WebView2 Releases it after returning from the callback, so
'     if args need not be kept inside View_OnNavigationStarting, no AddRef is needed.
'     We only get the URL and write it to the log, so we do not keep it.
' ============================================================
Public Sub View_OnNavigationStarting(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' ============================================================
    ' D-1 issue 7: thin wrapper for the in-callback guard
    '   The body is View_OnNavigationStarting_Core. Here m_inCallback is raised by 1, Core is called, and
    '   it is decremented by 1 *whichever path exits*. Core has Exit Sub in the middle, so
    '   writing +1 / -1 directly inside Core would miss restoring in some paths (post-hoc report (a)).
    '   If Core exits with an error, the counter is restored and the same error is re-raised
    '   (not swallowed = previous behaviour unchanged. Post-hoc report (b)).
    ' ============================================================
    Dim errNum As Long
    Dim errSrc As String
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View_OnNavigationStarting_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errSrc = Err.source
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        Debug.Print "Wv2Pane.View_OnNavigationStarting: error in Core " & errNum & " " & errDsc & _
                    " (in-callback guard already released, error re-raised)"
        Err.Raise errNum, errSrc, errDsc
    End If
End Sub


Private Sub View_OnNavigationStarting_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Debug.Print "Wv2Pane.View_OnNavigationStarting: pSender=" & pSender & ", pArgs=" & pArgs
    If pSender <> m_pView Then
        Debug.Print "  [WARNING] pSender does not match m_pView (pSender=" & pSender & _
                    ", m_pView=" & m_pView & ")"
    End If

    If pArgs = 0 Then
        Debug.Print "  pArgs is 0, so skipping URL retrieval"
        Exit Sub
    End If

    ' Call args.get_Uri to get the URL
    Dim url As String
    url = GetStringProperty(pArgs, 3, "NavStartingArgs.get_Uri")
    Debug.Print "  -> destination URL: " & url

    ' English fork (security): the settings tab carries the "Settings" host
    ' object, which can change the saved search engine. If that tab leaves the
    ' settings page for a real document, drop the host object and the flag so
    ' the next site cannot call it. (NavigateToString of the settings page
    ' itself arrives as a data:/about: URI and is left alone.)
    If m_isSettings Then
        Dim scheme As String
        scheme = LCase$(Left$(url, 6))
        If Left$(scheme, 5) = "http:" Or scheme = "https:" Or Left$(scheme, 5) = "file:" Then
            View_RemoveHostObjectFromScript "Settings"
            m_isSettings = False
            Wv2Log.LogI "Wv2Pane: settings tab navigated to a web page - Settings host object removed"
        End If
    End If
End Sub


' ============================================================
' View_AddNavigationCompleted (new in stage 9.4b, renamed in 9.6a)
'
'   Calls ICoreWebView2::add_NavigationCompleted (vtable index 15) to
'   register a persistent event handler.
'
'   Same pattern as View_AddNavigationStarting. The differences are:
'     - Kind: HK_NavigationCompleted
'     - vtable index: 15 (add) / 16 (remove)
'     - args type: ICoreWebView2NavigationCompletedEventArgs
'                (has get_IsSuccess, get_WebErrorStatus, get_NavigationId)
' ============================================================
Public Function View_AddNavigationCompleted() As Long
    If m_pView = 0 Then
        View_AddNavigationCompleted = &H80004003   ' E_POINTER
        Exit Function
    End If

    If Not (m_persistentHandlers(HK_NavigationCompleted) Is Nothing) Then
        Debug.Print "Wv2Pane.View_AddNavigationCompleted: already registered, skipping"
        View_AddNavigationCompleted = 0
        Exit Function
    End If

    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_NavigationCompleted, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_AddNavigationCompleted: failed to Acquire the handler"
        View_AddNavigationCompleted = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If

    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As LongLong
    Dim hr As Long
    hr = dcf(m_pView, 15, "View.add_NavigationCompleted", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_AddNavigationCompleted: add_NavigationCompleted failed hr=&H" & Hex(hr)
        Wv2Thunks.Thunks_DiscardHandler h   ' English fork: return the slot (see Thunks_DiscardHandler)
        Set h = Nothing
        View_AddNavigationCompleted = hr
        Exit Function
    End If

    m_persistentTokens(HK_NavigationCompleted) = token
    Set m_persistentHandlers(HK_NavigationCompleted) = h
    Debug.Print "Wv2Pane.View_AddNavigationCompleted OK: token=" & token & ", handler.slot=" & h.Slot

    View_AddNavigationCompleted = 0
End Function


' ============================================================
' View_RemoveNavigationCompleted (new in stage 9.4b, renamed in 9.6a)
'   Unregisters with View_RemoveEvent (vtable index 16).
'   If not registered, a no-op (returns S_OK).
' ============================================================
Public Function View_RemoveNavigationCompleted() As Long
    If m_pView = 0 Then
        View_RemoveNavigationCompleted = &H80004003
        Exit Function
    End If

    If m_persistentHandlers(HK_NavigationCompleted) Is Nothing Then
        View_RemoveNavigationCompleted = 0
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView, 16, "View.remove_NavigationCompleted", _
             m_persistentTokens(HK_NavigationCompleted))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_RemoveNavigationCompleted: unregister failed hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Pane.View_RemoveNavigationCompleted OK"
    End If

    Set m_persistentHandlers(HK_NavigationCompleted) = Nothing
    m_persistentTokens(HK_NavigationCompleted) = 0^

    View_RemoveNavigationCompleted = hr
End Function


' ============================================================
' View_OnNavigationCompleted (new in stage 9.4b, renamed in 9.6a)
'   Called from ComCallbackHandler via dynamic binding.
'
'   args = calls get_IsSuccess of ICoreWebView2NavigationCompletedEventArgs
'   (vtable index 3) and writes success/failure to the log.
'
'   Public Sub is required (a Friend does not ride on IDispatch::Invoke).
' ============================================================
Public Sub View_OnNavigationCompleted(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' ============================================================
    ' D-1 issue 7: thin wrapper for the in-callback guard
    '   The body is View_OnNavigationCompleted_Core. Here m_inCallback is raised by 1, Core is called, and
    '   it is decremented by 1 *whichever path exits*. Core has Exit Sub in the middle, so
    '   writing +1 / -1 directly inside Core would miss restoring in some paths (post-hoc report (a)).
    '   If Core exits with an error, the counter is restored and the same error is re-raised
    '   (not swallowed = previous behaviour unchanged. Post-hoc report (b)).
    ' ============================================================
    Dim errNum As Long
    Dim errSrc As String
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View_OnNavigationCompleted_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errSrc = Err.source
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        Debug.Print "Wv2Pane.View_OnNavigationCompleted: error in Core " & errNum & " " & errDsc & _
                    " (in-callback guard already released, error re-raised)"
        Err.Raise errNum, errSrc, errDsc
    End If
End Sub


Private Sub View_OnNavigationCompleted_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Debug.Print "Wv2Pane.View_OnNavigationCompleted: pSender=" & pSender & ", pArgs=" & pArgs
    If pSender <> m_pView Then
        Debug.Print "  [WARNING] pSender does not match m_pView (pSender=" & pSender & _
                    ", m_pView=" & m_pView & ")"
    End If

    If pArgs = 0 Then
        Debug.Print "  pArgs is 0, so skipping IsSuccess retrieval"
        ' Even without pArgs it is valid as a page-load-complete notification
        m_lastNavOk = True
        m_pageLoaded = True
        ' --- 9.12a: notify the URL state on this path too (do not bury it in the early return) ---
        NotifyUrlState pSender
        Exit Sub
    End If

    ' Call args.get_IsSuccess to get success or failure
    Dim isOk As Boolean
    isOk = GetBoolProperty(pArgs, 3, "NavCompletedArgs.get_IsSuccess")
    m_lastNavOk = isOk
    Debug.Print "  -> isSuccess: " & isOk

    ' Set the page-load-complete flag (added in 9.4c)
    ' Provisional implementation: set True whenever NavigationCompleted arrives, regardless of success/failure.
    ' This is for synchronizing on the first load. If additional navigations in an SPA etc.
    ' need to be distinguished, design that separately on the comprehensive class side.
    m_pageLoaded = True

    ' --- 9.12a: notify the owning Browser of the URL state (for updating the address bar) ---
    NotifyUrlState pSender
End Sub


' ============================================================
' NotifyUrlState (9.12a, notify the Browser of the URL state)
'
'   Reads the current URL / can-go-back / can-go-forward from sender (ICoreWebView2)
'   and passes them to the owning Browser's NotifyUrlChanged. The Browser looks up the
'   index in reverse by object-reference match and does RaiseEvent TabUrlChanged (exactly the same structure
'   as 9.11a's NotifyTitleChanged).
'
'   *The Pane does not know the nav bar exists* It just tells the
'   Browser that "my state has changed". Whether there is a listener (the nav bar) is not its concern (loose coupling).
'
'   Callers:
'     - View_OnNavigationCompleted : completion of a normal page navigation
'     - View_OnHistoryChanged      : SPA pushState / back / forward
'   Fired from both, but it is the whole-sync approach (design principle 60), so multiple firing is harmless.
' ============================================================
Private Sub NotifyUrlState(ByVal pSender As LongPtr)
    If m_browser Is Nothing Then Exit Sub
    If pSender = 0 Then Exit Sub

    Dim url As String
    url = GetStringProperty(pSender, 4, "View.get_Source")
    Dim canBack As Boolean
    canBack = GetBoolProperty(pSender, 38, "View.get_CanGoBack")
    Dim canFwd As Boolean
    canFwd = GetBoolProperty(pSender, 39, "View.get_CanGoForward")

    m_browser.NotifyUrlChanged Me, url, canBack, canFwd
End Sub


' ============================================================
' View_AddDocumentTitleChanged (new in stage 9.4b, renamed in 9.6a)
'
'   Calls ICoreWebView2::add_DocumentTitleChanged (vtable index 46) to
'   register a persistent event handler.
'
'   args is an IUnknown* (no information). The title is obtained by firing sender.get_DocumentTitle
'   (vtable 48).
' ============================================================
Public Function View_AddDocumentTitleChanged() As Long
    If m_pView = 0 Then
        View_AddDocumentTitleChanged = &H80004003
        Exit Function
    End If

    If Not (m_persistentHandlers(HK_DocumentTitleChanged) Is Nothing) Then
        Debug.Print "Wv2Pane.View_AddDocumentTitleChanged: already registered, skipping"
        View_AddDocumentTitleChanged = 0
        Exit Function
    End If

    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_DocumentTitleChanged, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_AddDocumentTitleChanged: failed to Acquire the handler"
        View_AddDocumentTitleChanged = &H8000FFFF
        Exit Function
    End If

    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As LongLong
    Dim hr As Long
    hr = dcf(m_pView, 46, "View.add_DocumentTitleChanged", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_AddDocumentTitleChanged: add_DocumentTitleChanged failed hr=&H" & Hex(hr)
        Wv2Thunks.Thunks_DiscardHandler h   ' English fork: return the slot (see Thunks_DiscardHandler)
        Set h = Nothing
        View_AddDocumentTitleChanged = hr
        Exit Function
    End If

    m_persistentTokens(HK_DocumentTitleChanged) = token
    Set m_persistentHandlers(HK_DocumentTitleChanged) = h
    Debug.Print "Wv2Pane.View_AddDocumentTitleChanged OK: token=" & token & ", handler.slot=" & h.Slot

    View_AddDocumentTitleChanged = 0
End Function


' ============================================================
' View_RemoveDocumentTitleChanged (new in stage 9.4b, renamed in 9.6a)
'   Unregisters with View_RemoveEvent (vtable index 47).
' ============================================================
Public Function View_RemoveDocumentTitleChanged() As Long
    If m_pView = 0 Then
        View_RemoveDocumentTitleChanged = &H80004003
        Exit Function
    End If

    If m_persistentHandlers(HK_DocumentTitleChanged) Is Nothing Then
        View_RemoveDocumentTitleChanged = 0
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView, 47, "View.remove_DocumentTitleChanged", _
             m_persistentTokens(HK_DocumentTitleChanged))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_RemoveDocumentTitleChanged: unregister failed hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Pane.View_RemoveDocumentTitleChanged OK"
    End If

    Set m_persistentHandlers(HK_DocumentTitleChanged) = Nothing
    m_persistentTokens(HK_DocumentTitleChanged) = 0^

    View_RemoveDocumentTitleChanged = hr
End Function


' ============================================================
' View_OnDocumentTitleChanged (new in stage 9.4b, renamed in 9.6a)
'   Called from ComCallbackHandler via dynamic binding.
'
'   sender = calls get_DocumentTitle (vtable index 48) of ICoreWebView2 and
'   writes the current title to the log.
'
'   args is an IUnknown* with no information, so it must be fetched via sender.
'
'   Public Sub is required (a Friend does not ride on IDispatch::Invoke).
' ============================================================
Public Sub View_OnDocumentTitleChanged(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' ============================================================
    ' D-1 issue 7: thin wrapper for the in-callback guard
    '   The body is View_OnDocumentTitleChanged_Core. Here m_inCallback is raised by 1, Core is called, and
    '   it is decremented by 1 *whichever path exits*. Core has Exit Sub in the middle, so
    '   writing +1 / -1 directly inside Core would miss restoring in some paths (post-hoc report (a)).
    '   If Core exits with an error, the counter is restored and the same error is re-raised
    '   (not swallowed = previous behaviour unchanged. Post-hoc report (b)).
    ' ============================================================
    Dim errNum As Long
    Dim errSrc As String
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View_OnDocumentTitleChanged_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errSrc = Err.source
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        Debug.Print "Wv2Pane.View_OnDocumentTitleChanged: error in Core " & errNum & " " & errDsc & _
                    " (in-callback guard already released, error re-raised)"
        Err.Raise errNum, errSrc, errDsc
    End If
End Sub


Private Sub View_OnDocumentTitleChanged_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Debug.Print "Wv2Pane.View_OnDocumentTitleChanged: pSender=" & pSender & ", pArgs=" & pArgs
    If pSender <> m_pView Then
        Debug.Print "  [WARNING] pSender does not match m_pView (pSender=" & pSender & _
                    ", m_pView=" & m_pView & ")"
    End If

    ' Call sender (ICoreWebView2).get_DocumentTitle (vtable index 48) to get the title
    Dim title As String
    title = GetStringProperty(pSender, 48, "View.get_DocumentTitle")
    Debug.Print "  -> title: " & title

    ' --- 9.11c: cache the title in this pane (option 3, unifying the source of truth) ---
    m_title = title

    ' --- 9.11a: notify the owning Browser of the title change (via the back reference m_browser) ---
    '   Reuse the back reference with a proven record in NewWindowRequested as is.
    '   When m_browser is not set (usage without a back reference), only Debug.Print is used
    '   as a fallback, so loose coupling is not broken.
    If Not m_browser Is Nothing Then
        m_browser.NotifyTitleChanged Me, title
    End If
End Sub


' ============================================================
' View_AddWebMessageReceived (new in stage 9.4c, renamed in 9.6a)
'
'   Calls ICoreWebView2::add_WebMessageReceived (vtable index 34) to
'   register a persistent event handler.
'
'   Receives messages sent from the JS side with window.chrome.webview.postMessage(text).
'   On receipt, Wv2Pane.View_OnWebMessageReceived is called.
' ============================================================
Public Function View_AddWebMessageReceived() As Long
    If m_pView = 0 Then
        View_AddWebMessageReceived = &H80004003
        Exit Function
    End If

    If Not (m_persistentHandlers(HK_WebMessageReceived) Is Nothing) Then
        Debug.Print "Wv2Pane.View_AddWebMessageReceived: already registered, skipping"
        View_AddWebMessageReceived = 0
        Exit Function
    End If

    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_WebMessageReceived, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_AddWebMessageReceived: failed to Acquire the handler"
        View_AddWebMessageReceived = &H8000FFFF
        Exit Function
    End If

    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As LongLong
    Dim hr As Long
    hr = dcf(m_pView, 34, "View.add_WebMessageReceived", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_AddWebMessageReceived: add_WebMessageReceived failed hr=&H" & Hex(hr)
        Wv2Thunks.Thunks_DiscardHandler h   ' English fork: return the slot (see Thunks_DiscardHandler)
        Set h = Nothing
        View_AddWebMessageReceived = hr
        Exit Function
    End If

    m_persistentTokens(HK_WebMessageReceived) = token
    Set m_persistentHandlers(HK_WebMessageReceived) = h
    Debug.Print "Wv2Pane.View_AddWebMessageReceived OK: token=" & token & ", handler.slot=" & h.Slot

    View_AddWebMessageReceived = 0
End Function


' ============================================================
' View_RemoveWebMessageReceived (new in stage 9.4c, renamed in 9.6a)
'   Unregisters with View_RemoveEvent (vtable index 35).
' ============================================================
Public Function View_RemoveWebMessageReceived() As Long
    If m_pView = 0 Then
        View_RemoveWebMessageReceived = &H80004003
        Exit Function
    End If

    If m_persistentHandlers(HK_WebMessageReceived) Is Nothing Then
        View_RemoveWebMessageReceived = 0
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView, 35, "View.remove_WebMessageReceived", _
             m_persistentTokens(HK_WebMessageReceived))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_RemoveWebMessageReceived: unregister failed hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Pane.View_RemoveWebMessageReceived OK"
    End If

    Set m_persistentHandlers(HK_WebMessageReceived) = Nothing
    m_persistentTokens(HK_WebMessageReceived) = 0^

    View_RemoveWebMessageReceived = hr
End Function


' ============================================================
' View_OnWebMessageReceived (new in stage 9.4c, renamed in 9.6a)
'   Called from ComCallbackHandler via dynamic binding.
'
'   Receives messages that JS sent with window.chrome.webview.postMessage(text).
'   Public Sub is required (a Friend does not ride on IDispatch::Invoke).
'
'   vtable indexes of args = ICoreWebView2WebMessageReceivedEventArgs:
'     3 : get_Source                 (sender URL, used in this implementation)
'     4 : get_WebMessageAsJson       (unused, add if needed)
'     5 : TryGetWebMessageAsString   (succeeds when JS sent a string, used in this implementation)
'
'   TryGetWebMessageAsString returns an HRESULT, but GetStringProperty returns
'   a string only on success regardless of HRESULT (an empty string on failure).
'   If JS sent an object with postMessage(jsonObj), this becomes an empty string
'   (in that case fetch it with get_WebMessageAsJson).
' ============================================================
Public Sub View_OnWebMessageReceived(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' ============================================================
    ' D-1 issue 7: thin wrapper for the in-callback guard
    '   The body is View_OnWebMessageReceived_Core. Here m_inCallback is raised by 1, Core is called, and
    '   it is decremented by 1 *whichever path exits*. Core has Exit Sub in the middle, so
    '   writing +1 / -1 directly inside Core would miss restoring in some paths (post-hoc report (a)).
    '   If Core exits with an error, the counter is restored and the same error is re-raised
    '   (not swallowed = previous behaviour unchanged. Post-hoc report (b)).
    ' ============================================================
    Dim errNum As Long
    Dim errSrc As String
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View_OnWebMessageReceived_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errSrc = Err.source
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        Debug.Print "Wv2Pane.View_OnWebMessageReceived: error in Core " & errNum & " " & errDsc & _
                    " (in-callback guard already released, error re-raised)"
        Err.Raise errNum, errSrc, errDsc
    End If
End Sub


Private Sub View_OnWebMessageReceived_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Debug.Print "Wv2Pane.View_OnWebMessageReceived: pSender=" & pSender & ", pArgs=" & pArgs
    If pSender <> m_pView Then
        Debug.Print "  [WARNING] pSender does not match m_pView (pSender=" & pSender & _
                    ", m_pView=" & m_pView & ")"
    End If

    If pArgs = 0 Then
        Debug.Print "  pArgs is 0, so skipping message retrieval"
        Exit Sub
    End If

    ' Call args.get_Source to get the sender URL
    Dim source As String
    source = GetStringProperty(pArgs, 3, "WebMsgArgs.get_Source")
    Debug.Print "  -> source: " & source

    ' Call args.TryGetWebMessageAsString to get the message body
    ' Succeeds when JS sent a string, an empty string when it sent an object
    Dim msg As String
    msg = GetStringProperty(pArgs, 5, "WebMsgArgs.TryGetWebMessageAsString")
    Debug.Print "  -> message: " & msg

    ' --- 9.11c: forward the message to the registered sink (TabBar etc.) (option P) ---
    '   The sink must implement OnPaneWebMessage(pane, source, msg).
    '   It is a late-bound call, so without an implementation it becomes a runtime error,
    '   but it is the responsibility of the sink-registering side (TabBar) to provide the correct method.
    If Not m_webMsgSink Is Nothing Then
        m_webMsgSink.OnPaneWebMessage Me, source, msg
    End If
End Sub


' ============================================================
' View_AddNewWindowRequested (new in stage 9.6a)
'
'   Calls ICoreWebView2::add_NewWindowRequested (vtable index 44) to
'   register a persistent event handler.
'
'   Fires on an in-page "open a new window" request (target="_blank" link,
'   window.open, Ctrl+click, etc.).
'
'   Behaviour in stage 9.6a:
'     Only writes the URL to the log. Since put_Handled is not set, the browser behaves with the
'     default (= a separate OS window, or opened in the OS default browser).
'     Bing links continue to be in the state of "cannot be opened from VBA" via this path
'     (the previous behaviour verified by tarboh).
'
'   Stage 9.6b is planned to add the behaviour of put_Handled = TRUE + View_Navigate(uri) on the same View,
'   so that Bing links can be opened "on the same screen".
' ============================================================
Public Function View_AddNewWindowRequested() As Long
    If m_pView = 0 Then
        View_AddNewWindowRequested = &H80004003   ' E_POINTER
        Exit Function
    End If

    If Not (m_persistentHandlers(HK_NewWindowRequested) Is Nothing) Then
        Debug.Print "Wv2Pane.View_AddNewWindowRequested: already registered, skipping"
        View_AddNewWindowRequested = 0
        Exit Function
    End If

    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_NewWindowRequested, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_AddNewWindowRequested: failed to Acquire the handler"
        View_AddNewWindowRequested = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If

    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As LongLong
    Dim hr As Long
    hr = dcf(m_pView, 44, "View.add_NewWindowRequested", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_AddNewWindowRequested: add_NewWindowRequested failed hr=&H" & Hex(hr)
        Wv2Thunks.Thunks_DiscardHandler h   ' English fork: return the slot (see Thunks_DiscardHandler)
        Set h = Nothing
        View_AddNewWindowRequested = hr
        Exit Function
    End If

    m_persistentTokens(HK_NewWindowRequested) = token
    Set m_persistentHandlers(HK_NewWindowRequested) = h
    Debug.Print "Wv2Pane.View_AddNewWindowRequested OK: token=" & token & ", handler.slot=" & h.Slot

    View_AddNewWindowRequested = 0
End Function


' ============================================================
' View_RemoveNewWindowRequested (new in stage 9.6a)
'   Unregisters with View_RemoveEvent (vtable index 45).
'   If not registered, a no-op (returns S_OK).
' ============================================================
Public Function View_RemoveNewWindowRequested() As Long
    If m_pView = 0 Then
        View_RemoveNewWindowRequested = &H80004003
        Exit Function
    End If

    If m_persistentHandlers(HK_NewWindowRequested) Is Nothing Then
        View_RemoveNewWindowRequested = 0
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView, 45, "View.remove_NewWindowRequested", _
             m_persistentTokens(HK_NewWindowRequested))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_RemoveNewWindowRequested: unregister failed hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Pane.View_RemoveNewWindowRequested OK"
    End If

    Set m_persistentHandlers(HK_NewWindowRequested) = Nothing
    m_persistentTokens(HK_NewWindowRequested) = 0^

    View_RemoveNewWindowRequested = hr
End Function


' ============================================================
' View_OnNewWindowRequested (new in 9.6a, handling added in 9.6b)
'
'   vtable layout of args = ICoreWebView2NewWindowRequestedEventArgs
'   (from the IDL comment of the old WebView2.cls):
'      3 : get_Uri(LPWSTR* uri)            <- obtained in 9.6a
'      4 : put_NewWindow(ICoreWebView2*)
'      5 : get_NewWindow(ICoreWebView2**)
'      6 : put_Handled(BOOL)               <- used in 9.6b
'      7 : get_Handled(BOOL*)
'      8 : get_IsUserInitiated(BOOL*)
'      9 : GetDeferral(ICoreWebView2Deferral**)
'     10 : get_WindowFeatures(ICoreWebView2WindowFeatures**)
'
'   Behaviour in stage 9.6b:
'     1. Get the URL with get_Uri and write it to the log (inherited from 9.6a)
'     2. Declare "the VBA side handles it" with args.put_Handled(1)
'     3. Fire View_Navigate(url) on the same View to navigate "on the same screen"
'
'   Behaviour when put_Handled is not set (confirmed in 9.6a):
'     The WebView2 runtime creates a new WebView2 window inside the Excel process.
'     Result confirmed with UIA Inspect:
'       - Class name: Chrome_WidgetWin_1
'       - Name: "Bing Images" (= title only, no browser name)
'       - In Task Manager, a WebView2 process in the "Utility" group
'     That is, it does not hand the baton to Edge; WebView2 itself opens
'     a separate window. put_Handled = TRUE can suppress this.
'
'   Behaviour on failure (option a + option c, agreed with tarboh):
'     When put_Handled fails: WebView2 opens a new window on its own (returns to the 9.6a behaviour)
'     When View_Navigate fails: stays on the current page, the link does not open (no hang)
'     Neither is fatal, so only Debug.Print and processing continues.
'
'   Future (tabbed browser, equivalent to 9.6c):
'     Instead of View_Navigate, the policy is to pass another Wv2Pane instance with
'     args.put_NewWindow. put_Handled stays as is. To be handled by the TabWebView2 comprehensive class.
'
'   Public Sub is required (a Friend does not ride on IDispatch::Invoke).
' ============================================================
Public Sub View_OnNewWindowRequested(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' ============================================================
    ' D-1 issue 7: thin wrapper for the in-callback guard
    '   The body is View_OnNewWindowRequested_Core. Here m_inCallback is raised by 1, Core is called, and
    '   it is decremented by 1 *whichever path exits*. Core has Exit Sub in the middle, so
    '   writing +1 / -1 directly inside Core would miss restoring in some paths (post-hoc report (a)).
    '   If Core exits with an error, the counter is restored and the same error is re-raised
    '   (not swallowed = previous behaviour unchanged. Post-hoc report (b)).
    ' ============================================================
    Dim errNum As Long
    Dim errSrc As String
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View_OnNewWindowRequested_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errSrc = Err.source
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        Debug.Print "Wv2Pane.View_OnNewWindowRequested: error in Core " & errNum & " " & errDsc & _
                    " (in-callback guard already released, error re-raised)"
        Err.Raise errNum, errSrc, errDsc
    End If
End Sub


Private Sub View_OnNewWindowRequested_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Debug.Print "Wv2Pane.View_OnNewWindowRequested: pSender=" & pSender & ", pArgs=" & pArgs
    If pSender <> m_pView Then
        Debug.Print "  [WARNING] pSender does not match m_pView (pSender=" & pSender & _
                    ", m_pView=" & m_pView & ")"
    End If

    If pArgs = 0 Then
        Debug.Print "  pArgs is 0, so skipping URL retrieval"
        Exit Sub
    End If

    ' --- 1. Get the destination URL with args.get_Uri (vtable index 3) (inherited from 9.6a) ---
    Dim url As String
    url = GetStringProperty(pArgs, 3, "NewWindowReqArgs.get_Uri")
    Debug.Print "  -> new window request URL: " & url

    ' --- 9.26/9.27: switch how NewWindowRequested is handled by mode ---
    '   When m_browser is Nothing (standalone Pane operation = SetBrowser not run), there is no place for the
    '   mode itself, so it is always LEGACY-equivalent (the current behaviour below).
    Dim nwMode As Long
    nwMode = NWM_LEGACY
    If Not (m_browser Is Nothing) Then nwMode = m_browser.NewWindowMode

    If nwMode = NWM_POPUP Then
        ' *Mode 1 (the 9.26 probe)* Without setting put_Handled and without issuing a new-tab request,
        '   Exit. The handling after that is left to the runtime, and the original request
        '   (with the body, if POST) continues in its own popup window = spec fact 23.
        '   *For verification* the opened window is outside our management (close it manually).
        Debug.Print "  [9.26 delegation mode] delegating to the runtime without setting put_Handled"
        Debug.Print "    (if a popup opens, the original request is continued = POST survived)"
        Debug.Print "    * No own new tab is created. Close the popup manually."
        Exit Sub

    ElseIf nwMode = NWM_PREWARM Then
        ' *Mode 2 (the real aim of 9.27)* Passes the CoreWebView2 of the warmed-up spare tab to
        '   args.put_NewWindow. The runtime loads the original request into that WebView,
        '   so it opens in our own tab with the POST method/body/Referer preserved.
        Dim pNew As LongPtr
        pNew = m_browser.ClaimPrewarmViewPtr()

        If pNew = 0 Then
            ' No spare / not Ready / already Claimed. Fall to the current behaviour without breaking anything.
            Debug.Print "  *[9.27 pre-warm] the spare tab cannot be used, so falling back to the current behaviour*"
            Debug.Print "    -> this navigation turns into GET (a 405 for a POST link). Reserving a refill."
            m_browser.RequestPrewarmAsync
            ' v Falls to the current behaviour below (put_Handled + RequestNewTabAsync)
        Else
            Dim hrNw As Long
            hrNw = dcf(pArgs, 4, "NewWindowReqArgs.put_NewWindow", pNew)

            If hrNw <> 0 Then
                ' put_NewWindow failed. Return the reservation, then fall to the current behaviour.
                Debug.Print "  *[9.27 pre-warm] put_NewWindow failed hr=&H" & Hex(hrNw) & _
                            " -> falling back to the current behaviour*"
                m_browser.ReleasePrewarmClaim
                ' v Falls to the current behaviour below
            Else
                Debug.Print "  [9.27 pre-warm] put_NewWindow OK (View of the spare tab=" & pNew & ")"

                ' Whether to also use put_Handled is switched by a Browser-side switch (insurance for issue 5).
                If m_browser.SetHandledWithNewWindow Then
                    hrNw = dcf(pArgs, 6, "NewWindowReqArgs.put_Handled", 1&)
                    If hrNw <> 0 Then
                        Debug.Print "    put_Handled failed hr=&H" & Hex(hrNw)
                    Else
                        Debug.Print "    put_Handled OK (used together with put_NewWindow)"
                    End If
                Else
                    Debug.Print "    not setting put_Handled (SetHandledWithNewWindow = False)"
                End If

                ' Adoption (register in m_tabs -> TabAdded -> ActivateTab) and refill are left to the Timer.
                ' *Not waiting here* is the key point of the pre-warm approach (avoids spec fact 14).
                m_browser.RequestAdoptPrewarmedAsync url
                Exit Sub
            End If
        End If
    End If

    ' --- Added in 9.6b: handling by opening on the same View ---
    If url = "" Then
        Debug.Print "  url is empty, so skipping put_Handled / View_Navigate"
        Exit Sub
    End If

    ' --- 2. put_Handled = TRUE ---
    '   Declare to the WebView2 runtime that "the VBA side handles it".
    '   The argument 1& is the BOOL TRUE (in Win32 BOOL = Long; the & suffix
    '   explicitly guarantees it is a Long, in line with design principle 29).
    Dim hr As Long
    hr = dcf(pArgs, 6, "NewWindowReqArgs.put_Handled", 1&)
    If hr <> 0 Then
        Debug.Print "  put_Handled failed hr=&H" & Hex(hr) & _
                    " (WebView2 will open a new window)"
    Else
        Debug.Print "  put_Handled OK (the new window is suppressed)"
    End If

    ' --- 3. Open in a new tab (9.10e, option A + option Y) ---
    '   Changed from 9.6b's "View_Navigate on the same View" to "open in a Wv2Browser new tab".
    '   AddTabWithUrl is called via the m_browser back reference (option Y).
    '   *The return value of window.open becomes null* because this is the simple approach (option A)
    '     that does not pass a new View via put_NewWindow. The proper way (put_NewWindow + Deferral) is a separate future stage.
    '   When m_browser is not set (SetBrowser not run), to avoid breaking, fall back to
    '     View_Navigate on the same View (so it works in standalone Pane operation too).
    If m_browser Is Nothing Then
        Debug.Print "  [WARNING] m_browser not set (SetBrowser not run). Falling back to Navigate on the same View" & _
                    "."
        hr = View_Navigate(url)
        If hr <> 0 Then Debug.Print "  View_Navigate failed hr=&H" & Hex(hr)
        Exit Sub
    End If

    ' *Deadlock avoidance (option beta Timer)* If m_browser.AddTabWithUrl is called directly here,
    '   AddTab's WaitPaneReady waits for the new Controller's ControllerCompleted, but
    '   WebView2 does not deliver it until this handler (NewWindowRequested) returns,
    '   so it deadlocks (timeout) (spec fact 14). So
    '   the request is made to RequestNewTabAsync, and AddTabWithUrl is run after the handler completes
    '   via a one-shot Timer.
    m_browser.RequestNewTabAsync url
    Debug.Print "  Requesting new tab creation from the Timer (RequestNewTabAsync): " & url
End Sub


' ============================================================
' SetBrowser / ClearBrowserRef (9.10e, set and disconnect the Wv2Browser back reference)
'
'   SetBrowser: Wv2Browser.AddTab passes itself (Me) when creating a tab. This lets
'     Wv2Pane call m_browser.AddTabWithUrl from View_OnNewWindowRequested.
'   ClearBrowserRef: called by Wv2Browser.Shutdown to break the circular reference (Browser.m_tabs ? Pane.m_browser).
'     If the UserForm sets m_browser=Nothing without calling this,
'     neither Browser nor Pane is released because of the circular reference, and it leaks (design principle 58).
'
'   *Design principle 58* VBA's reference-counting GC does not automatically release mutual references.
'     When holding a back reference, prepare a "Shutdown that explicitly cuts it" at the upper level.
' ============================================================
Public Sub SetBrowser(ByVal browser As Wv2Browser)
    Set m_browser = browser
End Sub

Public Sub ClearBrowserRef()
    Set m_browser = Nothing
    Set m_webMsgSink = Nothing   ' 9.11c: cut the sink back reference too (design principle 58)
End Sub


' ============================================================
' SetWebMsgSink / ClearWebMsgSink (9.11c, set and disconnect the WebMessage sink, option P)
'
'   SetWebMsgSink: when this pane receives a postMessage from JS, forwards it to
'     OnPaneWebMessage(pane, source, msg) of the specified sink object.
'     TabBar registers itself as the sink at Init time. To actually enable reception,
'     separately call View_AddWebMessageReceived (the registration has existed since 9.4c).
'   ClearWebMsgSink: cuts the sink. Also called from ClearBrowserRef (at release).
'
'   *A separate line from m_browser (the back reference for NewWindowRequested)* This one is
'     dedicated to WebMessage forwarding. Not mixing the two separates the responsibilities.
' ============================================================
Public Sub SetWebMsgSink(ByVal sink As Object)
    Set m_webMsgSink = sink
End Sub

Public Sub ClearWebMsgSink()
    Set m_webMsgSink = Nothing
End Sub


' ============================================================
' View_AddHistoryChanged (new in stage 9.9b, SPA support)
'
'   Calls ICoreWebView2::add_HistoryChanged (vtable index 13) to
'   register a persistent event handler. It is an event of the base ICoreWebView2, so
'   m_pView is used (m_pView2 is not needed).
'
'   Firing timing:
'     When an SPA calls history.pushState / replaceState / back / forward
'     and the history state changes. It also fires for a normal page navigation (Navigate)
'     if it is pushed onto the history. The main means of detecting SPA route changes.
'
'   * The implementation uses View_AddNavigationStarting as a verbatim model *
'     (design principle 45: when passing a COM object to WebView2,
'      always go through pVTableObj = h.Slot + VTABLE_OBJ_OFFSET.
'      Simple-argument types like View_PostWebMessageAsString are not used as the model)
'
'   Differences:
'     - Kind: HK_HistoryChanged
'     - vtable index: 13 (add) / 14 (remove)
'     - args type: IUnknown* (no information. Get get_Source etc. via sender)
'
'   Return value: HRESULT (S_OK = 0 means success)
' ============================================================
Public Function View_AddHistoryChanged() As Long
    If m_pView = 0 Then
        View_AddHistoryChanged = &H80004003   ' E_POINTER
        Exit Function
    End If

    ' If already registered, do nothing
    If Not (m_persistentHandlers(HK_HistoryChanged) Is Nothing) Then
        Debug.Print "Wv2Pane.View_AddHistoryChanged: already registered, skipping"
        View_AddHistoryChanged = 0
        Exit Function
    End If

    ' Acquire a ComCallbackHandler (owner = Me; receives events via dynamic binding)
    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_HistoryChanged, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_AddHistoryChanged: failed to Acquire the handler"
        View_AddHistoryChanged = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If

    ' Fire add_HistoryChanged (vtable index 13)
    '   Important: the address of the handler COM object passed to WebView2 is not
    '   h.Slot itself but h.Slot + VTABLE_OBJ_OFFSET (design principle 45).
    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As LongLong
    Dim hr As Long
    hr = dcf(m_pView, 13, "View.add_HistoryChanged", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_AddHistoryChanged: add_HistoryChanged failed hr=&H" & Hex(hr)
        Wv2Thunks.Thunks_DiscardHandler h   ' English fork: return the slot (see Thunks_DiscardHandler)
        Set h = Nothing
        View_AddHistoryChanged = hr
        Exit Function
    End If

    ' Success: save the token and the ComCallbackHandler
    m_persistentTokens(HK_HistoryChanged) = token
    Set m_persistentHandlers(HK_HistoryChanged) = h
    Debug.Print "Wv2Pane.View_AddHistoryChanged OK: token=" & token & ", handler.slot=" & h.Slot

    View_AddHistoryChanged = 0
End Function


' ============================================================
' View_RemoveHistoryChanged (new in stage 9.9b)
'   Unregisters with View_RemoveEvent (vtable index 14). It is a base IF, so m_pView.
'   If not registered, a no-op (returns S_OK).
' ============================================================
Public Function View_RemoveHistoryChanged() As Long
    If m_pView = 0 Then
        View_RemoveHistoryChanged = &H80004003
        Exit Function
    End If

    If m_persistentHandlers(HK_HistoryChanged) Is Nothing Then
        View_RemoveHistoryChanged = 0   ' not registered is a no-op
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView, 14, "View.remove_HistoryChanged", _
             m_persistentTokens(HK_HistoryChanged))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_RemoveHistoryChanged: unregister failed hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Pane.View_RemoveHistoryChanged OK"
    End If

    Set m_persistentHandlers(HK_HistoryChanged) = Nothing
    m_persistentTokens(HK_HistoryChanged) = 0^

    View_RemoveHistoryChanged = hr
End Function


' ============================================================
' View_OnHistoryChanged (new in stage 9.9b)
'   Called from ComCallbackHandler via dynamic binding.
'
'   args is an IUnknown* with no information, so the current state is fetched
'   via sender (= ICoreWebView2) (same pattern as DocumentTitleChanged).
'
'   Information obtained (option N2-a):
'     - get_Source        (vtable 4)  : current URL
'     - get_CanGoBack     (vtable 38) : whether "back" is possible
'     - get_CanGoForward  (vtable 39) : whether "forward" is possible
'   The essential SPA behaviour that pushState on an SPA changes the URL and makes CanGoBack True
'   can be made visible with one event.
'
'   Public Sub is required (a Friend does not ride on IDispatch::Invoke).
' ============================================================
Public Sub View_OnHistoryChanged(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' ============================================================
    ' D-1 issue 7: thin wrapper for the in-callback guard
    '   The body is View_OnHistoryChanged_Core. Here m_inCallback is raised by 1, Core is called, and
    '   it is decremented by 1 *whichever path exits*. Core has Exit Sub in the middle, so
    '   writing +1 / -1 directly inside Core would miss restoring in some paths (post-hoc report (a)).
    '   If Core exits with an error, the counter is restored and the same error is re-raised
    '   (not swallowed = previous behaviour unchanged. Post-hoc report (b)).
    ' ============================================================
    Dim errNum As Long
    Dim errSrc As String
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View_OnHistoryChanged_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errSrc = Err.source
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        Debug.Print "Wv2Pane.View_OnHistoryChanged: error in Core " & errNum & " " & errDsc & _
                    " (in-callback guard already released, error re-raised)"
        Err.Raise errNum, errSrc, errDsc
    End If
End Sub


Private Sub View_OnHistoryChanged_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' 9.9c (beta): pArgs=0 is the specification of HistoryChanged (args is an IUnknown* with no information).
    ' Go through sender to fetch get_Source / get_CanGoBack / get_CanGoForward.
    Debug.Print "Wv2Pane.View_OnHistoryChanged: pSender=" & pSender & _
                ", pArgs=" & pArgs & " (args=0 is by specification)"
    If pSender <> m_pView Then
        Debug.Print "  [WARNING] pSender does not match m_pView (pSender=" & pSender & _
                    ", m_pView=" & m_pView & ")"
    End If

    ' Get the current state via sender (ICoreWebView2)
    Dim url As String
    url = GetStringProperty(pSender, 4, "View.get_Source")
    Dim canBack As Boolean
    canBack = GetBoolProperty(pSender, 38, "View.get_CanGoBack")
    Dim canFwd As Boolean
    canFwd = GetBoolProperty(pSender, 39, "View.get_CanGoForward")

    ' 9.9c (alpha): about:blank is normal in a NavigateToString environment.
    ' A note is added so that users are not confused by "the URL looks empty".
    If url = "about:blank" Then
        Debug.Print "  -> URL: about:blank (normal in a NavigateToString environment, 9.9b spec fact 6)"
    Else
        Debug.Print "  -> URL: " & url
    End If
    Debug.Print "  -> CanGoBack: " & canBack & ", CanGoForward: " & canFwd

    ' --- 9.12a: notify the owning Browser of the URL state (SPA pushState / back / forward) ---
    '   Here url / canBack / canFwd have already been obtained, so they are passed directly
    '   without going through NotifyUrlState (avoiding reading get_* twice).
    If Not m_browser Is Nothing Then
        m_browser.NotifyUrlChanged Me, url, canBack, canFwd
    End If
End Sub


' ============================================================
' View2_AddDOMContentLoaded (new in stage 9.9b, SPA support, ICoreWebView2_2)
'
'   Calls ICoreWebView2_2::add_DOMContentLoaded (vtable index 64) to
'   register a persistent event handler. This is an event of the derived IF (ICoreWebView2_2),
'   so m_pView2 is used rather than m_pView.
'   (= the first real-world use of EnsureView2 / m_pView2 implemented in 9.9a)
'
'   Firing timing:
'     When the page's DOM has finished being built (DOMContentLoaded, once at initial load).
'     Fires before NavigationCompleted, at the point the DOM becomes touchable.
'     Fires at initial load on an SPA too (does not fire on pushState = that is
'     HistoryChanged's job).
'
'   * The implementation uses View_AddHistoryChanged / View_AddNavigationStarting
'      as a verbatim model (design principle 45) *
'     The only difference is "use EnsureView2() + m_pView2 instead of m_pView".
'     pVTableObj = h.Slot + VTABLE_OBJ_OFFSET is always gone through in the same way.
'
'   Naming: derived from a derived IF, so the View2_ prefix (9.6a/9.9b naming convention).
'
'   Differences (compared with the base version):
'     - Kind: HK_DOMContentLoaded
'     - Interface: m_pView2 (lazily acquired with EnsureView2)
'     - vtable index: 64 (add) / 65 (remove)  * offsets on m_pView2
'     - args type: ICoreWebView2DOMContentLoadedEventArgs
'                (get_NavigationId = UINT64, vtable 3)
'
'   Return value: HRESULT (S_OK = 0 means success)
' ============================================================
Public Function View2_AddDOMContentLoaded() As Long
    ' The _2 IF is required. Get m_pView2 with EnsureView2 (lazily acquired if not yet).
    If Not EnsureView2() Then
        Debug.Print "Wv2Pane.View2_AddDOMContentLoaded: EnsureView2 failed (cannot get m_pView2)"
        View2_AddDOMContentLoaded = &H80004003   ' E_POINTER
        Exit Function
    End If

    ' If already registered, do nothing
    If Not (m_persistentHandlers(HK_DOMContentLoaded) Is Nothing) Then
        Debug.Print "Wv2Pane.View2_AddDOMContentLoaded: already registered, skipping"
        View2_AddDOMContentLoaded = 0
        Exit Function
    End If

    ' Acquire a ComCallbackHandler (owner = Me; receives events via dynamic binding)
    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_DOMContentLoaded, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View2_AddDOMContentLoaded: failed to Acquire the handler"
        View2_AddDOMContentLoaded = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If

    ' Fire add_DOMContentLoaded (vtable index 64, on m_pView2)
    '   Important: the address of the handler COM object passed to WebView2 is not
    '   h.Slot itself but h.Slot + VTABLE_OBJ_OFFSET (design principle 45).
    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As LongLong
    Dim hr As Long
    hr = dcf(m_pView2, 64, "View2.add_DOMContentLoaded", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View2_AddDOMContentLoaded: add_DOMContentLoaded failed hr=&H" & Hex(hr)
        Wv2Thunks.Thunks_DiscardHandler h   ' English fork: return the slot (see Thunks_DiscardHandler)
        Set h = Nothing
        View2_AddDOMContentLoaded = hr
        Exit Function
    End If

    ' Success: save the token and the ComCallbackHandler
    m_persistentTokens(HK_DOMContentLoaded) = token
    Set m_persistentHandlers(HK_DOMContentLoaded) = h
    Debug.Print "Wv2Pane.View2_AddDOMContentLoaded OK: token=" & token & ", handler.slot=" & h.Slot

    View2_AddDOMContentLoaded = 0
End Function


' ============================================================
' View2_RemoveDOMContentLoaded (new in stage 9.9b, ICoreWebView2_2)
'   Unregisters with remove_DOMContentLoaded (vtable index 65). It is a _2 IF, so m_pView2.
'   If not registered, a no-op (returns S_OK).
' ============================================================
Public Function View2_RemoveDOMContentLoaded() As Long
    ' The _2 IF is required. But this is unregistering, so it should already hold m_pView2.
    ' Go through EnsureView2 just in case (if not acquired, returns 0 and is rejected).
    If Not EnsureView2() Then
        Debug.Print "Wv2Pane.View2_RemoveDOMContentLoaded: m_pView2 is not acquired"
        View2_RemoveDOMContentLoaded = &H80004003
        Exit Function
    End If

    If m_persistentHandlers(HK_DOMContentLoaded) Is Nothing Then
        View2_RemoveDOMContentLoaded = 0   ' not registered is a no-op
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView2, 65, "View2.remove_DOMContentLoaded", _
             m_persistentTokens(HK_DOMContentLoaded))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View2_RemoveDOMContentLoaded: unregister failed hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Pane.View2_RemoveDOMContentLoaded OK"
    End If

    Set m_persistentHandlers(HK_DOMContentLoaded) = Nothing
    m_persistentTokens(HK_DOMContentLoaded) = 0^

    View2_RemoveDOMContentLoaded = hr
End Function


' ============================================================
' View2_OnDOMContentLoaded (new in stage 9.9b, ICoreWebView2_2)
'   Called from ComCallbackHandler via dynamic binding.
'
'   args = calls get_NavigationId of ICoreWebView2DOMContentLoadedEventArgs
'   (vtable index 3) and writes the NavigationId to the log.
'   NavigationId is a UINT64, so it is fetched with GetLongLongProperty (64-bit only).
'
'   Naming: derived from a derived IF, so the View2_On prefix.
'
'   Public Sub is required (a Friend does not ride on IDispatch::Invoke).
'
'   * 32-bit point *
'     GetLongLongProperty is a 64-bit-only helper enclosed in #If Win64.
'     In the 32-bit support stage, how to get NavigationId will be designed separately.
' ============================================================
Public Sub View2_OnDOMContentLoaded(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' ============================================================
    ' D-1 issue 7: thin wrapper for the in-callback guard
    '   The body is View2_OnDOMContentLoaded_Core. Here m_inCallback is raised by 1, Core is called, and
    '   it is decremented by 1 *whichever path exits*. Core has Exit Sub in the middle, so
    '   writing +1 / -1 directly inside Core would miss restoring in some paths (post-hoc report (a)).
    '   If Core exits with an error, the counter is restored and the same error is re-raised
    '   (not swallowed = previous behaviour unchanged. Post-hoc report (b)).
    ' ============================================================
    Dim errNum As Long
    Dim errSrc As String
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View2_OnDOMContentLoaded_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errSrc = Err.source
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        Debug.Print "Wv2Pane.View2_OnDOMContentLoaded: error in Core " & errNum & " " & errDsc & _
                    " (in-callback guard already released, error re-raised)"
        Err.Raise errNum, errSrc, errDsc
    End If
End Sub


Private Sub View2_OnDOMContentLoaded_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Debug.Print "Wv2Pane.View2_OnDOMContentLoaded: pSender=" & pSender & ", pArgs=" & pArgs
    If pSender <> m_pView Then
        Debug.Print "  [WARNING] pSender does not match m_pView (pSender=" & pSender & _
                    ", m_pView=" & m_pView & ")"
    End If

    If pArgs = 0 Then
        Debug.Print "  pArgs is 0, so skipping NavigationId retrieval"
        Exit Sub
    End If

    ' Get args.get_NavigationId (vtable index 3, UINT64)
    Dim navId As LongLong
    navId = GetLongLongProperty(pArgs, 3, "DOMContentLoadedArgs.get_NavigationId")
    Debug.Print "  -> navigationId: " & navId
End Sub


' ============================================================
' View_PostWebMessageAsString (new in stage 9.4c)
'
'   Calls ICoreWebView2::PostWebMessageAsString (vtable index 33) to
'   send a message to JS.
'
'   To receive it on the JS side, register the following listener in the HTML:
'     window.chrome.webview.addEventListener('message', event => {
'         console.log(event.data);   // received content
'     });
'
'   Note: WebView2 isolates messages "per page".
'         If PostMessage is fired before the page's NavigationCompleted (= the listener
'         is not yet registered), the message reaches nowhere.
'         Always wait for NavigationCompleted before firing.
' ============================================================
Public Function View_PostWebMessageAsString(ByVal webMessage As String) As Long
    If m_pView = 0 Then
        View_PostWebMessageAsString = &H80004003
        Exit Function
    End If
    View_PostWebMessageAsString = dcf(m_pView, 33, _
                                     "View.PostWebMessageAsString", _
                                     StrPtr(webMessage))
End Function


' ============================================================
' Navigation control methods (new in stage 9.4b)
'
'   Exposes the navigation-related methods of ICoreWebView2 as Public.
'   Users can call them from a UserForm etc. as m_pane.View_xxx().
'
'   The vtable index of each method was verified on a real machine in the old version:
'     4  : get_Source
'     6  : NavigateToString
'     40 : GoBack
'     41 : GoForward
'     48 : get_DocumentTitle
' ============================================================

' --- Display an HTML string directly in WebView2 ---
Public Function View_NavigateToString(ByVal htmlContent As String) As Long
    If m_pView = 0 Then
        View_NavigateToString = &H80004003
        Exit Function
    End If
    ' Reset the flag at the start of a new navigation
    m_pageLoaded = False
    View_NavigateToString = dcf(m_pView, 6, "View.NavigateToString", StrPtr(htmlContent))
End Function

' --- Expose a VBA object to JS (9.18, AddHostObjectToScript, vtable 49) ---
'   objName  : the name that becomes window.chrome.webview.hostObjects.<objName> on the JS side
'   instance : the VBA class instance to expose (Instancing 1-Private is fine)
'
'   *The old WebView2.cls implementation verified on a real machine, ported to the current convention*
'     Put the object in a Variant and pass its VarPtr as the second argument (VARIANT*).
'     The first argument is the host object name as an LPCWSTR. A form combining the existing View_Navigate (StrPtr) and
'     get_Settings (VarPtr).
'
'   *Timing of the call (order A)* In the settings tab, call it "before" NavigateToString.
'     The hypothesis is that if registration comes before Navigate, the loaded JS can use
'     hostObjects from the start. If it turns out undefined on a real machine, fall back to the
'     ready handshake approach (order B).
Public Function View_AddHostObjectToScript(ByVal objName As String, ByVal instance As Object) As Long
    If m_pView = 0 Then
        View_AddHostObjectToScript = &H80004003   ' E_POINTER
        Exit Function
    End If
    Dim vObj As Variant
    Set vObj = instance
    View_AddHostObjectToScript = dcf(m_pView, 49, "View.AddHostObjectToScript", _
                                     StrPtr(objName), VarPtr(vObj))
End Function

' --- Go back one step in history ---
' English fork: ICoreWebView2::RemoveHostObjectFromScript (vtable 50, checked against WebView2.h)
Public Function View_RemoveHostObjectFromScript(ByVal objName As String) As Long
    If m_pView = 0 Then
        View_RemoveHostObjectFromScript = &H80004003
        Exit Function
    End If
    View_RemoveHostObjectFromScript = dcf(m_pView, 50, "View.RemoveHostObjectFromScript", StrPtr(objName))
End Function

Public Function View_GoBack() As Long
    If m_pView = 0 Then
        View_GoBack = &H80004003
        Exit Function
    End If
    View_GoBack = dcf(m_pView, 40, "View.GoBack")
End Function

' --- Go forward one step in history ---
Public Function View_GoForward() As Long
    If m_pView = 0 Then
        View_GoForward = &H80004003
        Exit Function
    End If
    View_GoForward = dcf(m_pView, 41, "View.GoForward")
End Function

' --- Get the URL of the currently displayed page ---
Public Function View_GetSource() As String
    If m_pView = 0 Then Exit Function
    View_GetSource = GetStringProperty(m_pView, 4, "View.get_Source")
End Function

' --- Get the DocumentTitle of the currently displayed page ---
Public Function View_GetDocumentTitle() As String
    If m_pView = 0 Then Exit Function
    View_GetDocumentTitle = GetStringProperty(m_pView, 48, "View.get_DocumentTitle")
End Function


' ============================================================
' Added in stage 9.8b: full coverage of the ICoreWebView2 base IF methods (View_ family)
'                  + Ctrl_Close of ICoreWebView2Controller
'                  + helper-style implementation of ICoreWebView2Settings (Settings_ family)
'
'   Design principles:
'     *The View_ family follows the existing pattern (via m_pView + GetXxxProperty / dcf)
'     *The Settings family caches m_pSettings with EnsureSettings and
'       is self-contained in Wv2Pane (policy of not adding a Wv2Settings class, 9.8b-A adopted)
'     *BOOL getters (View_GetCanGoBack etc.) are in the style of returning As Boolean
'       (HRESULT checked internally, consistent with IsPageLoaded, 9.8b-C adopted)
'     *void ones (View_Reload, View_Stop, Ctrl_Close) return the HRESULT As Long
'       (consistent with View_Navigate)
' ============================================================

' --- Reload the page (equivalent to F5, View vtable 31) ---
Public Function View_Reload() As Long
    If m_pView = 0 Then
        View_Reload = &H80004003
        Exit Function
    End If
    ' This is a reload, so reset the page-load-complete flag too
    m_pageLoaded = False
    View_Reload = dcf(m_pView, 31, "View.Reload")
End Function

' --- Stop navigation / loading (equivalent to Esc, View vtable 43) ---
Public Function View_Stop() As Long
    If m_pView = 0 Then
        View_Stop = &H80004003
        Exit Function
    End If
    View_Stop = dcf(m_pView, 43, "View.Stop")
End Function

' --- Whether "back" is possible (View vtable 38, for enabling the UI Back button) ---
'   Return value: True if View_GoBack can be called
'   HRESULT is handled internally; returns False on failure
Public Function View_GetCanGoBack() As Boolean
    If m_pView = 0 Then Exit Function
    View_GetCanGoBack = GetBoolProperty(m_pView, 38, "View.get_CanGoBack")
End Function

' --- Whether "forward" is possible (View vtable 39, for enabling the UI Forward button) ---
'   Return value: True if View_GoForward can be called
'   HRESULT is handled internally; returns False on failure
Public Function View_GetCanGoForward() As Boolean
    If m_pView = 0 Then Exit Function
    View_GetCanGoForward = GetBoolProperty(m_pView, 39, "View.get_CanGoForward")
End Function

' --- Get the PID of the WebView2 browser process (View vtable 37, for diagnosis) ---
'   Return value: Process ID (Long) of the browser process
'   Useful for identifying the WebView2 process by matching against Task Manager
'   HRESULT is handled internally; returns 0 on failure
Public Function View_GetBrowserProcessId() As Long
    If m_pView = 0 Then Exit Function
    View_GetBrowserProcessId = GetLongProperty(m_pView, 37, "View.get_BrowserProcessId")
End Function


' ============================================================
' EnsureSettings (Private, internal helper)
'
'   If m_pSettings is not acquired (= 0), calls View vtable 3 (get_Settings) to
'   get ICoreWebView2Settings* and caches it in m_pSettings.
'   If already acquired, uses it as is.
'
'   COM ownership:
'     get_Settings returns the pointer as [out, retval]. By the time it returns
'     the refcount is already raised, so no separate AddRef is needed on the VBA side.
'     Responsibility for releasing m_pSettings lies with Class_Terminate.
'
'   Return value: True if m_pSettings is valid, False if acquisition failed
' ============================================================
Private Function EnsureSettings() As Boolean
    If m_pSettings <> 0 Then
        EnsureSettings = True
        Exit Function
    End If
    If m_pView = 0 Then
        Debug.Print "Wv2Pane.EnsureSettings: m_pView is not acquired"
        Exit Function
    End If

    Dim pSettings As LongPtr
    Dim hr As Long
    hr = dcf(m_pView, 3, "View.get_Settings", VarPtr(pSettings))
    If hr <> 0 Then
        Debug.Print "Wv2Pane.EnsureSettings: get_Settings failed hr=&H" & Hex(hr)
        Exit Function
    End If
    If pSettings = 0 Then
        Debug.Print "Wv2Pane.EnsureSettings: pSettings is 0"
        Exit Function
    End If

    m_pSettings = pSettings
    Debug.Print "Wv2Pane.EnsureSettings: m_pSettings acquired = " & m_pSettings
    EnsureSettings = True
End Function

' --- Set whether JS may run (Settings vtable 4 = put_IsScriptEnabled) ---
'   Default True. Setting False stops JS from running at all.
'   IsScriptEnabled is originally to be set after WebView2 starts and before navigation,
'   but it takes effect after navigation too (valid from the next load).
Public Function Settings_PutIsScriptEnabled(ByVal v As Boolean) As Long
    If Not EnsureSettings() Then
        Settings_PutIsScriptEnabled = &H80004003
        Exit Function
    End If
    Dim bv As Long
    If v Then bv = 1 Else bv = 0
    Settings_PutIsScriptEnabled = dcf(m_pSettings, 4, "Settings.put_IsScriptEnabled", bv)
End Function

' --- Set whether the right-click menu is shown (Settings vtable 14 = put_AreDefaultContextMenusEnabled) ---
'   Default True. Setting False stops all right-click-family menus, including Ctrl+click,
'   from appearing. Useful when you want to suppress right-click in a business app.
Public Function Settings_PutAreDefaultContextMenusEnabled(ByVal v As Boolean) As Long
    If Not EnsureSettings() Then
        Settings_PutAreDefaultContextMenusEnabled = &H80004003
        Exit Function
    End If
    Dim bv As Long
    If v Then bv = 1 Else bv = 0
    Settings_PutAreDefaultContextMenusEnabled = dcf(m_pSettings, 14, _
                                              "Settings.put_AreDefaultContextMenusEnabled", bv)
End Function

' --- Set whether DevTools (F12) is available (Settings vtable 12 = put_AreDevToolsEnabled) ---
'   Default True. Setting False means DevTools does not open even when F12 is pressed.
'   Used when you want to seal DevTools in a distributed app.
Public Function Settings_PutAreDevToolsEnabled(ByVal v As Boolean) As Long
    If Not EnsureSettings() Then
        Settings_PutAreDevToolsEnabled = &H80004003
        Exit Function
    End If
    Dim bv As Long
    If v Then bv = 1 Else bv = 0
    Settings_PutAreDevToolsEnabled = dcf(m_pSettings, 12, "Settings.put_AreDevToolsEnabled", bv)
End Function


' ============================================================
' Added in stage 9.9a: lazy acquisition of the ICoreWebView2_2 derived IF (foundation of the View2_ family)
'
'   Design decisions (agreed with tarboh):
'     *Option 1: m_pView2 field + EnsureView2 lazy acquisition (following the Settings style)
'     *Option a: IID_ICoreWebView2_2 is built each time as a local variable in EnsureView2
'             (no IID table added to Wv2Thunks, keeping 9.9a to a minimal Wv2Thunks change)
'     *Option X: return value is As Boolean (same form as EnsureSettings)
'
'   Source of IID_ICoreWebView2_2:
'     WebView2.h L5078:
'       EXTERN_C __declspec(selectany) const IID IID_ICoreWebView2_2 =
'           {0x9E8F0CF8,0xE670,0x4B5E,{0xB2,0xBC,0x73,0xE0,0x61,0xE3,0x18,0x4C}};
'     MIDL_INTERFACE string: "9E8F0CF8-E670-4B5E-B2BC-73E061E3184C"
'
'   Methods added by ICoreWebView2_2 (61?67 of the base ICoreWebView2):
'     61 : add_WebResourceResponseReceived
'     62 : remove_WebResourceResponseReceived
'     63 : NavigateWithWebResourceRequest
'     64 : add_DOMContentLoaded         <- planned for use in 9.9b
'     65 : remove_DOMContentLoaded      <- planned for use in 9.9b
'     66 : get_CookieManager
'     67 : get_Environment
' ============================================================


' ============================================================
' EnsureView2 (Private, internal helper)
'
'   If m_pView2 is not acquired (= 0), calls QueryInterface(IID_ICoreWebView2_2) to
'   get ICoreWebView2_2* and caches it in m_pView2.
'   If already acquired, uses it as is.
'
'   COM ownership:
'     QueryInterface returns the pointer already AddRef'd on success.
'     m_pView2 has a reference count independent of m_pView
'     (= merely the same COM object seen through a different IF, but the reference count is separate).
'     Responsibility for releasing m_pView2 lies with Class_Terminate / Ctrl_Close (children first).
'
'   How to call QueryInterface:
'     IUnknown::QueryInterface (vtable 0):
'       HRESULT QueryInterface(REFIID riid, void** ppvObject);
'     The first argument riid is a pointer to a GUID structure.
'     The second argument ppvObject is a LongPtr out argument.
'     Via dcf, pass them as VarPtr(localIID) and VarPtr(pv).
'
'   The IID is built each time as a local variable of EnsureView2 (option a adopted):
'     Not added to Wv2Thunks' m_iidTable; generated where and when it is needed.
'     It is called only once the first time (if already cached it exits immediately with True), so
'     the overhead is minimal. Call Wv2Thunks' FillGUID as Public.
'
'   Return value: True if m_pView2 is valid, False if acquisition failed (option X, same form as EnsureSettings)
' ============================================================
Private Function EnsureView2() As Boolean
    If m_pView2 <> 0 Then
        EnsureView2 = True
        Exit Function
    End If
    If m_pView = 0 Then
        Debug.Print "Wv2Pane.EnsureView2: m_pView is not acquired"
        Exit Function
    End If

    ' --- Assemble IID_ICoreWebView2_2 into a local GUID structure ---
    '   {9E8F0CF8-E670-4B5E-B2BC-73E061E3184C}
    '   FillGUID is a Public Sub of Wv2Thunks (promoted from Private to Public in 9.9a)
    Dim iidView2 As GUID
    FillGUID iidView2, "9E8F0CF8-E670-4B5E-B2BC-73E061E3184C"

    ' --- Fire QueryInterface(riid, &ppv) ---
    '   vtable index 0 is QueryInterface (IUnknown base).
    '   Arguments:
    '     #1: REFIID riid     -> VarPtr(iidView2)  (reference to the GUID structure)
    '     #2: void** ppvObject -> VarPtr(pv)        (reference to the output LongPtr)
    '   Return value: HRESULT (S_OK = 0 means success)
    Dim pv As LongPtr
    Dim hr As Long
    hr = dcf(m_pView, 0, "View.QueryInterface(ICoreWebView2_2)", _
             VarPtr(iidView2), VarPtr(pv))
    If hr <> 0 Then
        Debug.Print "Wv2Pane.EnsureView2: QueryInterface failed hr=&H" & Hex(hr)
        Exit Function
    End If
    If pv = 0 Then
        Debug.Print "Wv2Pane.EnsureView2: QueryInterface succeeded but pv is 0"
        Exit Function
    End If

    m_pView2 = pv
    Debug.Print "Wv2Pane.EnsureView2: m_pView2 acquired = " & m_pView2
    EnsureView2 = True
End Function


' ============================================================
' View2_Ensure (new in 9.9a, Public verification wrapper)
'
'   A thin wrapper for calling ?UserForm1.GetActivePane.View2_Ensure
'   from the Immediate window to confirm that m_pView2 was acquired.
'
'   From 9.9b on, each method such as View2_AddDOMContentLoaded calls
'   EnsureView2 internally, so direct use of this function is basically for verification only.
'
'   Option i adopted: verification is done with this function + confirming zero regressions with Test_AllEvents_Real.
'             No test function (Test_View2_Ensure_Real) is added.
'
'   Return value: True if m_pView2 is valid (acquired successfully or already held)
' ============================================================
Public Function View2_Ensure() As Boolean
    View2_Ensure = EnsureView2()
End Function


' ============================================================
' EnsureView3 (new in 9.10d-2a, lazy acquisition of the ICoreWebView2_3 derived IF)
'
'   Takes ICoreWebView2_3 out with QueryInterface and caches it in m_pView3.
'   Exactly the same form as EnsureView2 (9.9a). QI from the base m_pView
'   (ICoreWebView2_3 : ICoreWebView2_2 : ICoreWebView2 is single inheritance, so
'    it can be taken out of the base in one step).
'
'   The foundation for calling SetVirtualHostNameToFolderMapping (vtable 71).
'
'   IID_ICoreWebView2_3 = {A0D6DF20-3B92-416D-AA0C-437A9C727857}
'     (WebView2.h: 0xA0D6DF20,0x3B92,0x416D,{0xAA,0x0C,0x43,0x7A,0x9C,0x72,0x78,0x57})
'
'   Return value: True if m_pView3 is valid (acquired successfully or already held)
' ============================================================
Private Function EnsureView3() As Boolean
    If m_pView3 <> 0 Then
        EnsureView3 = True
        Exit Function
    End If
    If m_pView = 0 Then
        Debug.Print "Wv2Pane.EnsureView3: m_pView is not acquired"
        Exit Function
    End If

    ' --- Assemble IID_ICoreWebView2_3 into a local GUID structure ---
    '   FillGUID is a Public Sub of Wv2Thunks (promoted to Public in 9.9a)
    Dim iidView3 As GUID
    FillGUID iidView3, "A0D6DF20-3B92-416D-AA0C-437A9C727857"

    ' --- Fire QueryInterface(riid, &ppv) (vtable index 0) ---
    Dim pv As LongPtr
    Dim hr As Long
    hr = dcf(m_pView, 0, "View.QueryInterface(ICoreWebView2_3)", _
             VarPtr(iidView3), VarPtr(pv))
    If hr <> 0 Then
        Debug.Print "Wv2Pane.EnsureView3: QueryInterface failed hr=&H" & Hex(hr)
        Exit Function
    End If
    If pv = 0 Then
        Debug.Print "Wv2Pane.EnsureView3: QueryInterface succeeded but pv is 0"
        Exit Function
    End If

    m_pView3 = pv
    Debug.Print "Wv2Pane.EnsureView3: m_pView3 acquired = " & m_pView3
    EnsureView3 = True
End Function


' ============================================================
' View3_Ensure (new in 9.10d-2a, Public verification wrapper)
'   Confirms m_pView3 acquisition from Immediate with ?UserForm1.GetActivePane.View3_Ensure.
'   Same form as View2_Ensure (9.9a).
' ============================================================
Public Function View3_Ensure() As Boolean
    View3_Ensure = EnsureView3()
End Function


' ============================================================
' View3_SetVirtualHostNameToFolderMapping (new in 9.10d-2a, ICoreWebView2_3)
'
'   Calls ICoreWebView2_3::SetVirtualHostNameToFolderMapping (vtable index 71) to
'   map a local folder to a virtual host name. This allows
'   serving local files from a secure origin https://<hostName>/, avoiding
'   the pushState/replaceState restriction of the about:blank origin (spec fact 6).
'
'   *This is a simple-argument method* It does not pass a COM object (handler) to WebView2,
'   so design principle 45 (VTABLE_OBJ_OFFSET) is not needed. A string-argument type such as View_Navigate
'   is the model (handler types such as View_AddNavigationStarting are not
'   the model).
'
'   Signature (WebView2.h):
'     HRESULT SetVirtualHostNameToFolderMapping(
'       LPCWSTR hostName,     ' virtual host name (e.g. "appassets.example")
'       LPCWSTR folderPath,   ' absolute path of the local folder to map to
'       COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND accessKind)  ' DENY=0/ALLOW=1/DENY_CORS=2
'
'   *Timing of the call* Must be set before Navigate. And since View3 is
'   needed, go through EnsureView3 first.
'
'   accessKind is a Win32 enum (4-byte int). It is accepted as an argument so the user can choose
'   (normally pass 1 = ALLOW). Pass it to dcf as a Long as is
'   (a simple integer argument, same as put_Handled 1&).
'
'   Return value: HRESULT (S_OK = 0 means success)
' ============================================================
Public Function View3_SetVirtualHostNameToFolderMapping( _
    ByVal hostName As String, _
    ByVal folderPath As String, _
    ByVal accessKind As Long) As Long

    If Not EnsureView3() Then
        Debug.Print "Wv2Pane.View3_SetVirtualHostNameToFolderMapping: EnsureView3 failed"
        View3_SetVirtualHostNameToFolderMapping = &H80004003   ' E_POINTER
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView3, 71, "View3.SetVirtualHostNameToFolderMapping", _
             StrPtr(hostName), StrPtr(folderPath), accessKind)

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View3_SetVirtualHostNameToFolderMapping: failed hr=&H" & _
                    Hex(hr) & " (host=" & hostName & ", folder=" & folderPath & ")"
    Else
        Debug.Print "Wv2Pane.View3_SetVirtualHostNameToFolderMapping OK: https://" & _
                    hostName & "/ -> " & folderPath & " (accessKind=" & accessKind & ")"
    End If

    View3_SetVirtualHostNameToFolderMapping = hr
End Function


' ============================================================
' Ctrl_Close (Controller vtable 24)
'
'   Explicitly closes the WebView2 controller. Used when you want to end the Controller's lifetime
'   immediately from the VBA side.
'
'   Extension in 9.8c (option E 1+3 hybrid):
'     To improve safety after Close, the following is done:
'       (1) Fire Controller::Close as usual (vtable 24)
'       (2) Cache m_pSettings / m_pView2 / m_pView / m_pCtrl locally,
'           Release them on the VBA side too, and clear the members to 0
'           (m_pView2 added in 9.9a; Release order strictly child to parent = design principle 43)
'       (3) Clear m_persistentTokens to 0 (prevents double remove in Cleanup)
'           * Persistent handlers are Released by WebView2 itself on Close, so
'              the VBA side must not fire remove_xxx (= firing remove_xxx at a
'              slot whose refCount is already 0 risks an AV)
'       (4) Set m_closed = True (early return in View_xxx / Settings_xxx)
'       (5) m_state = Cs_Failed (to show it as a state different from right after Init)
'
'   As a result:
'     - Even if Class_Terminate runs after Close, CleanupAllPersistentHandlers skips
'       because m_persistentTokens are all 0 (= safe)
'     - Release is also skipped because m_pSettings/m_pView2/m_pView/m_pCtrl are all 0
'     - Even if the user mistakenly calls View_xxx after Close, m_closed causes an early return
'
'   Notes:
'     - The return value of Ctrl_Close is the HRESULT of Controller::Close itself
'       (= even if the later VBA-side clear processing fails, the return value is unaffected)
'     - Calling Init again on the same Wv2Pane instance is not an intended use
'       (= to reuse, create a separate instance with a new New Wv2Pane)
' ============================================================
Public Function Ctrl_Close() As Long
    If m_closed Then
        Debug.Print "Wv2Pane.Ctrl_Close: already Closed, ignoring"
        Ctrl_Close = 0  ' a double Close is harmless, so return S_OK
        Exit Function
    End If
    If m_pCtrl = 0 Then
        Ctrl_Close = &H80004003
        Exit Function
    End If

    Debug.Print "Wv2Pane.Ctrl_Close: m_pCtrl = " & m_pCtrl
    Dim hr As Long
    hr = dcf(m_pCtrl, 24, "Ctrl.Close")

    ' --- (2) (3) (4) (5): transition to the post-Close state ---
    '
    ' At the point Close is fired, the WebView2 runtime side should already have Released
    ' the registered handlers (observed on a real machine, the finding of 9.8b "side effects of Ctrl_Close").
    ' In that state, the VBA side also fires Release on m_pSettings/m_pView2/m_pView/m_pCtrl
    ' to settle the VBA-side AddRef share.
    '
    ' The order is the same as Class_Terminate: Settings -> View2 -> View -> Ctrl (strictly child to parent,
    ' m_pView2 added in 9.9a, design principle 43).
    If m_pSettings <> 0 Then
        Debug.Print "Wv2Pane.Ctrl_Close: Releasing m_pSettings = " & m_pSettings
        ComRelease m_pSettings
        m_pSettings = 0
    End If
    If m_pView3 <> 0 Then
        ' Added in 9.10d-2a: View3 is the most derived IF, Release it before m_pView2 (design principle 51).
        Debug.Print "Wv2Pane.Ctrl_Close: Releasing m_pView3 = " & m_pView3
        ComRelease m_pView3
        m_pView3 = 0
    End If
    If m_pView2 <> 0 Then
        ' Added in 9.9a: View2 is a derived IF obtained from View by QI, Release it before View.
        Debug.Print "Wv2Pane.Ctrl_Close: Releasing m_pView2 = " & m_pView2
        ComRelease m_pView2
        m_pView2 = 0
    End If
    If m_pView <> 0 Then
        Debug.Print "Wv2Pane.Ctrl_Close: Releasing m_pView = " & m_pView
        ComRelease m_pView
        m_pView = 0
    End If
    If m_pCtrl <> 0 Then
        Debug.Print "Wv2Pane.Ctrl_Close: Releasing m_pCtrl = " & m_pCtrl
        ComRelease m_pCtrl
        m_pCtrl = 0
    End If

    ' (3) Clear the persistent handlers' tokens to 0:
    '   WebView2 fires Release on its own at Close,
    '   so firing remove_xxx here would mean "firing remove at an already released handler",
    '   which is dangerous. Setting token to 0 makes CleanupAllPersistentHandlers
    '   skip it (it is implemented not to process elements with token=0, 9.4a).
    '
    '   9.9b: changed the loop range from a fixed value to LBound/UBound.
    '   It follows automatically even if the array upper limit is extended to HK_DOMContentLoaded
    '   (= the same "automated horizontal rollout" style as CleanupAllPersistentHandlers).
    Dim k As Long
    For k = LBound(m_persistentHandlers) To UBound(m_persistentHandlers)
        m_persistentTokens(k) = 0^
        Set m_persistentHandlers(k) = Nothing
    Next k

    ' (4) Set the m_closed flag
    m_closed = True
    m_state = Cs_Failed   ' regarded as an invalid state after Close

    Ctrl_Close = hr
End Function


' ============================================================
' Added in stage 9.8c: View_ExecuteScript + OnExecuteScriptCompleted
'                  (option a = concurrency OK, option alpha = event callback approach,
'                   option X = returns a sequential callback ID)
' ============================================================


' --- Run arbitrary JavaScript inside WebView2 (View vtable 29) ---
'
'   Arguments:
'     jsCode  : the JS code to run (String)
'
'   Return value (As Long):
'     >= 1    : the assigned callback ID (the call started normally)
'     <= 0    : error (m_pView = 0, already Closed, dcf failure, etc.)
'                 -1 -> m_closed (already Closed)
'                 -2 -> m_pView = 0 (View not acquired)
'                 -3 -> Acquire failed
'                 -4 -> dcf failed (the absolute value of the return is the low part of the HRESULT,
'                       details are written to Debug.Print)
'
'   Flow:
'     1. Assign a callback ID (increment m_executeScriptCounter by 1)
'     2. Get a new handler with AcquireHandlerFor(HK_ExecuteScriptCompleted, Me)
'     3. Set callbackId on the handler (CallbackId property Let)
'     4. Fire View::ExecuteScript(jsCode, handler.Slot)
'     5. Set the local h to Nothing (= Wv2Pane does not hold the ComCallbackHandler;
'        only the WebView2 runtime side holds it via AddRef)
'     6. Return the callback ID
'
'   Afterwards:
'     The WebView2 runtime runs the JS -> on completion fires handler.Invoke
'     -> ComCallbackHandler.Invoke[HK_ExecuteScriptCompleted] runs
'     -> Wv2Pane.OnExecuteScriptCompleted is called via dynamic binding
'     -> WebView2 Releases the handler -> refCount=0 -> auto-released
'
'   Concurrency:
'     Even if the user calls View_ExecuteScript 10 times in a row, each has an independent
'     slot + an independent callback ID, so they do not interfere (the core of option a).
Public Function View_ExecuteScript(ByVal jsCode As String) As Long
    If m_closed Then
        Debug.Print "Wv2Pane.View_ExecuteScript: rejected because already Closed"
        View_ExecuteScript = -1
        Exit Function
    End If
    If m_pView = 0 Then
        Debug.Print "Wv2Pane.View_ExecuteScript: m_pView is not acquired"
        View_ExecuteScript = -2
        Exit Function
    End If

    ' --- 1. Assign a callback ID ---
    m_executeScriptCounter = m_executeScriptCounter + 1
    Dim cbid As Long
    cbid = m_executeScriptCounter

    ' --- 2. Acquire a handler ---
    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_ExecuteScriptCompleted, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_ExecuteScript: handler Acquire failed"
        View_ExecuteScript = -3
        Exit Function
    End If

    ' --- 3. Set the callback ID ---
    h.callbackId = cbid

    ' --- 4. Fire ExecuteScript (vtable 29) ---
    '   Signature: HRESULT ExecuteScript(LPCWSTR javaScript, IExecuteScriptCompletedHandler*)
    '
    '   Important: the pointer to the "handler COM object" passed to WebView2 is not
    '   h.Slot itself but h.Slot + VTABLE_OBJ_OFFSET.
    '   We must pass an address that points not at the head of the thunk header but at the VTable object part
    '   inside the thunk (= the same pattern as the existing add_NavigationStarting etc.,
    '   design principle 8 family).
    '
    '   If this is forgotten, WebView2 follows a bogus address as a vtable and
    '   crashes immediately.
    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim hr As Long
    hr = dcf(m_pView, 29, "View.ExecuteScript", StrPtr(jsCode), pVTableObj)

    ' --- 5. Discard the local reference (Wv2Pane does not keep it) ---
    '   WebView2 fires AddRef and holds it itself, so it is fine for us to drop it.
    '   It is called via m_owner (= Me Wv2Pane) on the completion notice.
    If hr <> 0 Then Wv2Thunks.Thunks_DiscardHandler h   ' English fork: no completion will come
    Set h = Nothing

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_ExecuteScript: dcf failed hr=&H" & Hex(hr) & _
                    " (callbackId=" & cbid & " discarded, the handler is expected to be released immediately)"
        View_ExecuteScript = -4
        Exit Function
    End If

    Debug.Print "Wv2Pane.View_ExecuteScript OK: callbackId=" & cbid & _
                ", jsLen=" & Len(jsCode)
    View_ExecuteScript = cbid
End Function


' ============================================================
' View_ExecuteScriptSafe (new in stage 9.9c, helper for structured results of JS exceptions)
'
'   Wraps the JS expression specified by the user in try-catch and then passes it to ExecuteScript.
'   The completion notice's resultJson is always one of the following structured JSONs:
'     On success: {"ok":true, "value":<JS return value>}
'     On failure: {"ok":false, "errorName":"<exception name>", "errorMessage":"<message>",
'              "stack":"<stack trace>"}
'
'   This makes it possible to structurally distinguish on the VBA side the problem found in 9.8c,
'   where "on a JS exception HRESULT=S_OK + result=null" was indistinguishable.
'
'   Design decisions (agreed with tarboh):
'     *F1: the name is View_ExecuteScriptSafe (concise; Safe = wrapped in try-catch)
'     *F2: on the VBA side, flow the JSON string as is to OnExecuteScriptCompleted
'           (64-bit VBA has no standard JSON parsing. Users do simple analysis with InStr etc.)
'     *F3: simple concatenation (escaping of " in user code is the user's responsibility)
'
'   * Constraint (note for users) *
'     The user code must be an "expression", not a "statement".
'     (= it assumes a single expression assignable as var __r = (USER_CODE))
'
'     Concrete NG / OK examples:
'       NG: "throw new Error('x')"   <- throw is a statement. It cannot be placed inside
'                                       the parentheses of an expression, so wrappedJs
'                                       becomes a syntax error at the parse stage, and ExecuteScript
'                                       returns resultJson = "null"
'                                       (= it never reaches try-catch)
'       OK: "(()=>{ throw new Error('x') })()"
'                                     <- wrapping in an immediately-invoked function makes it treatable as an "expression"
'                                       (the return value is not determined because throw runs, but
'                                        the catch clause definitely fires)
'       OK: "(function(){ var x=1; return x+2; })()"
'                                     <- multiple statements + return wrapped in an immediately-invoked function
'       OK: "nonexistentFunction()"   <- an expression (function call). A runtime
'                                       ReferenceError can be caught by catch
'
'     * Fact found by real-machine verification (9.9c) *
'       Hard to write in standard VBA documentation, but a JS syntax error
'       comes back from ExecuteScript as HRESULT=S_OK + result="null".
'       It cannot be distinguished from a runtime exception (ReferenceError etc.). This helper
'       handles runtime exceptions but cannot handle syntax errors.
'       (The proper way to catch syntax errors is to use eval(), but that is excessive)
'
'     To include " in a string literal, the user side writes \"
'       (JS's \" escape, not VBA's "" escape).
'
'   Usage example:
'     cbid = m_pane.View_ExecuteScriptSafe("1 + 2")
'       -> resultJson = "{\"ok\":true,\"value\":3}"
'     cbid = m_pane.View_ExecuteScriptSafe("(()=>{throw new Error('boom')})()")
'       -> resultJson = "{\"ok\":false,\"errorName\":\"Error\",
'                         \"errorMessage\":\"boom\",\"stack\":\"...\"}"
'       * Passing "throw new Error('boom')" directly gives resultJson="null"
'         (see the constraint above)
'     cbid = m_pane.View_ExecuteScriptSafe("nonexistentFunction()")
'       -> resultJson = "{\"ok\":false,\"errorName\":\"ReferenceError\",...}"
'
'   Return value: same as View_ExecuteScript (a callback ID or a negative error code)
' ============================================================
Public Function View_ExecuteScriptSafe(ByVal userCode As String) As Long
    ' Assemble the user code into a JS expression that is wrapped in try-catch and returns structured JSON.
    ' Build it on one line with no line breaks (source maps are meaningless for ExecuteScript, so readability < simplicity).
    Dim wrappedJs As String
    wrappedJs = "(function(){try{var __r=(" & userCode & ");" & _
                "return {ok:true,value:__r};}" & _
                "catch(e){return {ok:false,errorName:e.name," & _
                "errorMessage:e.message,stack:e.stack};}})()"

    Debug.Print "Wv2Pane.View_ExecuteScriptSafe: wrappedJsLen=" & Len(wrappedJs)
    View_ExecuteScriptSafe = View_ExecuteScript(wrappedJs)
End Function


' --- Receive the completion notice of ExecuteScript (dynamic-binding receiver) ---
'
'   Called from ComCallbackHandler.Handler_Invoke[HK_ExecuteScriptCompleted] via
'   m_owner.OnExecuteScriptCompleted. It cannot be called by dynamic binding unless Public,
'   so Public is required (= design principle 14 family).
'
'   Arguments:
'     callbackId  : the sequential ID that View_ExecuteScript returned
'     errorCode   : HRESULT (0 = S_OK)
'     resultJson  : the JS return value as a JSON-encoded string
'                   Examples:
'                     ExecuteScript("1 + 2")     -> "3"
'                     ExecuteScript("'foo'")     -> "\"foo\""
'                     ExecuteScript("document.title")
'                                                -> "\"Search - Microsoft Bing\""
'                     ExecuteScript("undefined") -> "null" (* undefined becomes null when JSON-ized)
'                     ExecuteScript("throw 1")   -> "" (errorCode is non-0)
'
'   Option alpha second implementation (extended in 9.9c):
'     1. Log output inside Wv2Pane (= the 9.8c first implementation, kept for isolating regressions)
'     2. Notification to UserForm1.OnExecuteScriptResult (= added in 9.9c)
'
'   * Design decision on notifying UserForm1 (option E1-a-i) *
'     Wv2Pane does **not hold** a reference to UserForm1. A VBA UserForm is
'     globally accessible at module level (= writing UserForm1 auto-creates
'     it), so calling UserForm1.OnExecuteScriptResult directly here
'     notifies without a circular reference.
'
'     This is somewhat crude hierarchically in the sense that "Wv2Pane knows UserForm1 exists"
'     (= normally it is cleaner for the lower level not to know the upper), but
'     when TabWebView2 comes in at 9.10, the "path that carries result notifications upward" itself
'     will be redesigned, so 9.9c makes it a direct call as a provisional measure.
'
'     No defenses such as On Error Resume Next are put on the UserForm1 side, so
'     UserForm1.OnExecuteScriptResult is to be implemented so that it throws no errors.
Public Sub OnExecuteScriptCompleted( _
    ByVal callbackId As Long, _
    ByVal errorCode As Long, _
    ByVal resultJson As String)

    ' --- (1) Log inside Wv2Pane (implemented in 9.8c, option E2-X = keep) ---
    Debug.Print "Wv2Pane.OnExecuteScriptCompleted: callbackId=" & callbackId & _
                ", errorCode=" & errorCode & _
                ", resultJson=" & resultJson

    ' --- (1b) D-1: update the pending table (issues 2 and 5) ---
    '   If it is a callbackId that EvalSync is waiting on, write the result.
    '   If already discarded (late arrival after a timeout), leave a log and discard.
    '   Results of direct View_ExecuteScript calls that do not go through EvalSync are not
    '   in the table, so they just pass through.
    Dim k As String
    Dim ent As Variant
    Dim found As Boolean

    k = "cb" & callbackId
    If Not m_evalPending Is Nothing Then
        found = False
        On Error Resume Next
        ent = m_evalPending(k)
        found = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0

        If found Then
            If ent(0) = EVAL_DISCARDED Then
                Debug.Print "Wv2Pane.OnExecuteScriptCompleted: *discarded* " & k & _
                            " arrived late, so discarding it (completion after a timeout)"
                m_evalPending.Remove k
            Else
                m_evalPending.Remove k
                m_evalPending.Add Array(EVAL_DONE, errorCode, resultJson), k
            End If
        End If
    End If

    ' --- (2) Notification to UserForm1 (added in 9.9c, option alpha second implementation) ---
    '   Use VBA's global UserForm reference (no circular reference, option E1-a-i).
    '   Regarding the case where UserForm1 does not exist (not launched, or already Terminated):
    '   VBA auto-creates UserForm1 merely by referencing it, so
    '   the call itself goes through even if it is not launched. If already Terminated,
    '   it is auto-created + Initialized again.
    '   However, the completion notice of ExecuteScript originally only arrives when Wv2Pane is alive
    '   = UserForm1 should be on display, so
    '   this auto-creation path is realistically never taken.
    ' English fork: no longer calls UserForm1.OnExecuteScriptResult. That made this class
    ' uncompilable without UserForm1, and a late completion after the form was unloaded
    ' silently re-created a hidden UserForm1 (running UserForm_Initialize). The form's
    ' handler only printed the same values that (1) above already prints.
End Sub


' ============================================================
' *The body of D-1: EvalSync and its surroundings*
' ============================================================


' ============================================================
' EvalSync (D-1) - fire a JS expression and receive the return value synchronously
'
'   Arguments:
'     jsExpr     : the JS *expression* to evaluate (statements not allowed. Wrap multiple statements in an immediately-invoked function)
'     timeoutSec : upper limit of the wait (default 5 seconds). If 0 or less, rounded to the default.
'
'   Return value (As String):
'     Success : the string obtained by JSON.stringify of the JS return value. LastEvalOk = True
'     Failure : an empty string. The reason goes into LastEvalError. LastEvalOk = False
'
'   Flow:
'     1. Check rejection conditions (already Closed / View not acquired / *in-callback*)
'     2. Fire JSON.stringify(expression) with View_ExecuteScriptSafe and get a callbackId
'     3. Register a waiting entry in the pending table
'     4. Spin DoEvents to wait for completion (same form as WaitPaneReady)
'     5. Completed -> unpack the JSON and return it / exceeded -> mark as discarded and return an empty string
'
'   *Why wrap with JSON.stringify (post-hoc report (e))*
'     The Safe wrapper returns ok and value, but if the type of value varies among number, object and string,
'     extraction on the VBA side would branch per type. Passing through JSON.stringify
'     makes value *always a string*, and a single limited parser suffices.
'
'   *Do not break / step through during the DoEvents loop (spec fact 20)*
' ============================================================
Public Function EvalSync(ByVal jsExpr As String, _
                         Optional ByVal timeoutSec As Single = 5) As String
    Dim res As String
    Dim cbid As Long
    Dim k As String
    Dim t As Single
    Dim ent As Variant
    Dim timedOut As Boolean
    Dim errCode As Long
    Dim resultJson As String
    Dim v As String
    Dim eName As String
    Dim eMsg As String

    res = ""
    m_lastEvalOk = False
    m_lastEvalError = ""

    ' --- 1. Rejection conditions ---
    If m_closed Then
        m_lastEvalError = "closed"
        Wv2Log.LogW "Wv2Pane.EvalSync: rejected because already Closed"
        EvalSync = ""
        Exit Function
    End If

    If m_pView = 0 Then
        m_lastEvalError = "no-view"
        Wv2Log.LogW "Wv2Pane.EvalSync: m_pView is not acquired"
        EvalSync = ""
        Exit Function
    End If

    If m_inCallback > 0 Then
        ' *Guard of issue 7* Firing here would hang until the timeout because of spec fact 14.
        m_lastEvalError = "in-callback"
        Wv2Log.LogW "Wv2Pane.EvalSync: *rejected because it was called from inside an event handler* " & _
                    "(depth=" & m_inCallback & "). When you need the JS result inside a handler, " & _
                    "call it after a beat via SafeTimer (spec fact 14). " & _
                    "If this appears even though you are outside a handler, suspect a leftover counter and " & _
                    "run ResetCallbackGuard."
        EvalSync = ""
        Exit Function
    End If

    If timeoutSec <= 0 Then timeoutSec = EVAL_TIMEOUT_DEFAULT
    If m_evalPending Is Nothing Then Set m_evalPending = New Collection

    ' --- 2. Fire via the Safe wrapper (issue 4) ---
    cbid = View_ExecuteScriptSafe("JSON.stringify(" & jsExpr & ")")
    If cbid <= 0 Then
        m_lastEvalError = "execute-failed(" & cbid & ")"
        Wv2Log.LogE "Wv2Pane.EvalSync: failed to issue ExecuteScript rc=" & cbid
        EvalSync = ""
        Exit Function
    End If

    ' --- 3. Register in the pending table ---
    k = "cb" & cbid
    m_evalPending.Add Array(EVAL_WAITING, 0&, ""), k

    m_evalDepth = m_evalDepth + 1
    If m_evalDepth > 1 Then
        Wv2Log.LogD "Wv2Pane.EvalSync: nested execution depth=" & m_evalDepth & " (" & k & ")"
    End If

    ' --- 4. Wait for completion (same form as WaitPaneReady) ---
    timedOut = False
    t = Timer
    Do
        DoEvents
        ent = m_evalPending(k)
        If ent(0) = EVAL_DONE Then Exit Do
        If ElapsedSec(t) > timeoutSec Then
            timedOut = True
            ' *Issue 5* Do not delete it; leave it marked discarded. To discard late arrivals.
            m_evalPending.Remove k
            m_evalPending.Add Array(EVAL_DISCARDED, 0&, ""), k
            Exit Do
        End If
    Loop

    m_evalDepth = m_evalDepth - 1

    If timedOut Then
        m_lastEvalError = "timeout(" & timeoutSec & "s)"
        Wv2Log.LogE "Wv2Pane.EvalSync: *timeout* marked " & k & " as discarded " & _
                    "(a completion notice arriving late will be discarded)"
        EvalSync = ""
        Exit Function
    End If

    ' --- 5. Unpack the result ---
    errCode = ent(1)
    resultJson = ent(2)
    m_evalPending.Remove k

    If errCode <> 0 Then
        m_lastEvalError = "hresult(&H" & Hex$(errCode) & ")"
        Wv2Log.LogE "Wv2Pane.EvalSync: ExecuteScript failed " & m_lastEvalError
        EvalSync = ""
        Exit Function
    End If

    If InStr(1, resultJson, """ok"":true") > 0 Then
        If Wv2Json.JsonPickStr(resultJson, "value", v) Then
            m_lastEvalOk = True
            res = v
        Else
            ' JSON.stringify(undefined) returns undefined, and the outer stringify drops
            ' the value key entirely. This is the signal that the return value was undefined.
            m_lastEvalOk = True
            res = "undefined"
        End If

    ElseIf InStr(1, resultJson, """ok"":false") > 0 Then
        Wv2Json.JsonPickStr resultJson, "errorName", eName
        Wv2Json.JsonPickStr resultJson, "errorMessage", eMsg
        m_lastEvalError = eName & ": " & eMsg
        Wv2Log.LogW "Wv2Pane.EvalSync: JS exception " & m_lastEvalError

    Else
        ' The Safe wrapper always returns ok, so reaching here means a JS syntax error
        ' (the case where ExecuteScript returns S_OK + null. Measured in 9.9c).
        m_lastEvalError = "syntax-error-or-null (resultJson=" & Left$(resultJson, 60) & ")"
        Wv2Log.LogE "Wv2Pane.EvalSync: " & m_lastEvalError
    End If

    EvalSync = res
End Function


' ============================================================
' LastEvalOk / LastEvalError (D-1) - the result of the previous EvalSync
'   Used to tell, when EvalSync returned an empty string, whether it is "JS returned an empty string"
'   or "it failed".
' ============================================================
Public Property Get LastEvalOk() As Boolean
    LastEvalOk = m_lastEvalOk
End Property

Public Property Get LastEvalError() As String
    LastEvalError = m_lastEvalError
End Property


' ============================================================
' InCallbackDepth / ResetCallbackGuard / Debug_SetInCallback (D-1 issue 7)
'
'   ResetCallbackGuard:
'     An emergency exit for when an event handler exited in an unexpected way and the counter was left over.
'     The wrapper always restores it so it is normally unnecessary, but when EvalSync keeps being
'     rejected as in-callback, call this.
'
'   Debug_SetInCallback:
'     *For verification only* An injection port for reproducing the guard firing from Test_D1_Guard.
'     Not called from production code (the Debug_ prefix follows the same convention as Wv2Browser's Debug_*).
' ============================================================
Public Property Get InCallbackDepth() As Long
    InCallbackDepth = m_inCallback
End Property

Public Sub ResetCallbackGuard()
    Debug.Print "Wv2Pane.ResetCallbackGuard: m_inCallback = " & m_inCallback & " -> 0"
    m_inCallback = 0
End Sub

Public Sub Debug_SetInCallback(ByVal n As Long)
    Debug.Print "Wv2Pane.Debug_SetInCallback: " & m_inCallback & " -> " & n & " (verification only)"
    m_inCallback = n
End Sub


' ============================================================
' JsQuote (D-2, Friend) - turn a VBA string into a JS string literal
'
'   *The only entry that avoids writing backslash + double quote in VBA source*
'     Wv2Json.JsonEscape escapes quotes, backslashes and control characters per the JSON spec,
'     and it is wrapped with Chr$(34). The escape sequences are *assembled at run time*
'     so they do not appear in the VBA source (conforms to the project's JS-in-VBA rule).
'
'   Examples:
'     JsQuote("input[name='q']")   -> "input[name='q']"          (as is)
'     JsQuote("a[title=""x""]")    -> the quotes become escaped per the JSON spec
'                                    (the escape sequences are *assembled at run time*)
'     JsQuote("C:\temp")           -> "C:\\temp"
'
'   It is Friend because Wv2Element.GetAttribute uses it to wrap attribute names.
' ============================================================
Friend Function JsQuote(ByVal s As String) As String
    JsQuote = Chr$(34) & Wv2Json.JsonEscape(s) & Chr$(34)
End Function


' ============================================================
' JsRegistryPrologue (D-2, Private) - self-generating code for the registry
'
'   The common part placed at the head of every D-2 JS expression (issue 1 option c).
'   Creates it if missing and binds the registry to variable R.
'   *If it was created here, the generation becomes a new value* = all existing handles become stale.
'
'   *JS string literals use single quotes* (project rule).
' ============================================================
Private Function JsRegistryPrologue() As String
    JsRegistryPrologue = _
        "if(typeof window.__wv2==='undefined'||!window.__wv2){" & _
        "window.__wv2={gen:'g'+Date.now()+'-'+Math.random(),els:[]};}" & _
        "var R=window.__wv2;"
End Function


' ============================================================
' GetElementById / QuerySelector (D-2) - find one element
'
'   Returns a Wv2Element if found, Nothing if not (issue 4 option a).
'
'   *Nothing has two meanings*
'     Nothing + LastEvalOk = True  ... really does not exist (normal)
'     Nothing + LastEvalOk = False ... failure. The reason is in LastEvalError
'
'   Usage example:
'     Dim p As Wv2Pane, el As Wv2Element
'     Set p = UserForm1.GetActivePane
'     Set el = p.QuerySelector("input[name='q']")
'     If el Is Nothing Then
'         Debug.Print "not found (ok=" & p.LastEvalOk & " err=" & p.LastEvalError & ")"
'     Else
'         Debug.Print el.TagName & " / " & el.Value
'     End If
' ============================================================
'   *From D-3, if ImplicitWaitSec > 0 it retries until found*
'     The default is 0, so unless set, it remains the same as D-2: "answers in one round trip".
Public Function GetElementById(ByVal elementId As String) As Wv2Element
    Set GetElementById = FindWithWait("document.getElementById(" & JsQuote(elementId) & ")", _
                                      "GetElementById", elementId, m_implicitWaitSec)
End Function

Public Function QuerySelector(ByVal selector As String) As Wv2Element
    Set QuerySelector = FindWithWait("document.querySelector(" & JsQuote(selector) & ")", _
                                     "QuerySelector", selector, m_implicitWaitSec)
End Function


' ============================================================
' FindOne (D-2, Private) - common implementation of search
'
'   finderExpr : a JS expression that returns one element (must return null if not found)
'   apiName / arg : for logging
'
'   The JS side returns the following:
'     {g:'<generation>', found:true,  h:<index>}   found
'     {g:'<generation>', found:false, h:0}        not found
'
'   *Why h does not use negative values* Wv2Json.JsonGetNum does not handle negative numbers (design principle 64).
'     Since found separates the cases, h can always be 0 or more.
' ============================================================
'   quiet (D-3): if True, do not write the "not found" log. During polling,
'                writing every round trip would fill up the Immediate window.
Private Function FindOne(ByVal finderExpr As String, _
                         ByVal apiName As String, _
                         ByVal arg As String, _
                         Optional ByVal quiet As Boolean = False) As Wv2Element
    Dim js As String
    Dim res As String
    Dim gen As String
    Dim h As Long
    Dim el As Wv2Element

    Set FindOne = Nothing

    js = "(function(){" & JsRegistryPrologue() & _
         "var e=" & finderExpr & ";" & _
         "if(!e){return {g:R.gen,found:false,h:0};}" & _
         "R.els.push(e);" & _
         "return {g:R.gen,found:true,h:R.els.length-1};})()"

    res = EvalSync(js)

    If Not m_lastEvalOk Then
        Debug.Print "Wv2Pane." & apiName & ": failed (" & arg & ") err=" & m_lastEvalError
        Exit Function
    End If

    ' Update the generation cache (only on success)
    If Wv2Json.JsonPickStr(res, "g", gen) Then
        m_domGen = gen
    End If

    If InStr(1, res, """found"":true") = 0 Then
        ' Not found. *This is normal* so m_lastEvalOk stays True.
        If Not quiet Then
            Debug.Print "Wv2Pane." & apiName & ": not found (" & arg & ")"
        End If
        Exit Function
    End If

    h = Wv2Json.JsonGetNum(res, "h")

    Set el = New Wv2Element
    el.InitElement Me, h, gen
    Set FindOne = el

    Debug.Print "Wv2Pane." & apiName & ": found (" & arg & ") handle=" & h & " gen=" & gen
End Function


' ============================================================
' CurrentDomGen (D-2) - read entry point of the generation cache
'
'   *This is not an authority* It is only "the value last heard from JS".
'   Used for the fast path of the stale check, and for verification / debug display.
' ============================================================
Public Property Get CurrentDomGen() As String
    CurrentDomGen = m_domGen
End Property


' ============================================================
' ClearElementRegistry (D-2, issue 3) - manually rebuild the registry
'
'   Rebuilds the JS-side registry with a new generation.
'   *All existing Wv2Element objects become stale.*
'
'   For cleanup when els has grown from many searches within the same page. Normally unnecessary
'   (because it vanishes wholesale on page navigation).
'
'   Return value: True = success. False = failure (reason in LastEvalError)
' ============================================================
Public Function ClearElementRegistry() As Boolean
    Dim js As String
    Dim res As String
    Dim gen As String

    js = "(function(){" & _
         "window.__wv2={gen:'g'+Date.now()+'-'+Math.random(),els:[]};" & _
         "return window.__wv2.gen;})()"

    res = EvalSync(js)

    If Not m_lastEvalOk Then
        Wv2Log.LogE "Wv2Pane.ClearElementRegistry: failed err=" & m_lastEvalError
        ClearElementRegistry = False
        Exit Function
    End If

    gen = Wv2Json.JsonUnescape(res)
    m_domGen = gen
    Wv2Log.LogD "Wv2Pane.ClearElementRegistry: rebuilt the registry (new generation=" & gen & ")"
    ClearElementRegistry = True
End Function


' ============================================================
' Elem_Read (D-2, Friend) - read one property of a Wv2Element
'
'   Arguments:
'     elemHandle : index into R.els[] on the JS side
'     elemGen    : the generation at the time the element was obtained
'     valueExpr  : a JS expression using variable e (the target element). Example "e.innerText"
'     outVal     : the value read (escapes already decoded)
'     outStatus  : failure reason. An empty string on success
'   Return value:
'     True  = read succeeded (True also when the value was null / undefined. outVal is an empty string)
'     False = stale or failure. The reason is in outStatus
'
'   The JS side returns the following:
'     {s:'gen', g:'<generation>'}                 generation differs = stale
'     {s:'nil', g:'<generation>'}                 no element at the index = treated as stale
'     {s:'ok',  g:'<generation>', has:false, v:''} value is null / undefined
'     {s:'ok',  g:'<generation>', has:true,  v:'<value>'}
'
'   *The value is always normalized to a string with String(x) before returning*
'     Returning it in the JSON type would require per-type extraction on the VBA side. The
'     return values of the read APIs are all String, so aligning on the JS side is simpler (and it also
'     avoids the undefined-vanishing problem of spec fact 31).
' ============================================================
Friend Function Elem_Read(ByVal elemHandle As Long, _
                          ByVal elemGen As String, _
                          ByVal valueExpr As String, _
                          ByRef outVal As String, _
                          ByRef outStatus As String) As Boolean
    Dim js As String
    Dim res As String
    Dim gen As String

    outVal = ""
    outStatus = ""
    Elem_Read = False

    ' --- Fast path: if the generation in hand differs from the cache, stale can be determined without a round trip ---
    '     (when the cache is stale it falls to the matching side and always asks, so no misjudgement)
    If Len(m_domGen) > 0 Then
        If elemGen <> m_domGen Then
            outStatus = "stale"
            Exit Function
        End If
    End If

    js = "(function(){" & JsRegistryPrologue() & _
         "if(R.gen!==" & JsQuote(elemGen) & "){return {s:'gen',g:R.gen};}" & _
         "var e=R.els[" & CStr(elemHandle) & "];" & _
         "if(!e){return {s:'nil',g:R.gen};}" & _
         "var x=" & valueExpr & ";" & _
         "if(x===null||x===undefined){return {s:'ok',g:R.gen,has:false,v:''};}" & _
         "return {s:'ok',g:R.gen,has:true,v:String(x)};})()"

    res = EvalSync(js)

    If Not m_lastEvalOk Then
        outStatus = m_lastEvalError
        Exit Function
    End If

    If Wv2Json.JsonPickStr(res, "g", gen) Then
        m_domGen = gen
    End If

    If InStr(1, res, """s"":""ok""") = 0 Then
        outStatus = "stale"
        Exit Function
    End If

    If InStr(1, res, """has"":true") = 0 Then
        ' The value is null / undefined. *Not an error* (no such attribute / an element without value)
        outVal = ""
        Elem_Read = True
        Exit Function
    End If

    Wv2Json.JsonPickStr res, "v", outVal
    Elem_Read = True
End Function


' ============================================================
' Elem_CheckStale (D-2, Friend) - whether the handle has become invalid
'
'   Return value: True = stale (no longer usable)
'   outStatus: the reason if the check failed. An empty string on success
'
'   *True (the safe side) is also returned when it could not be determined.*
'     Callers who need to distinguish should look at outStatus.
' ============================================================
Friend Function Elem_CheckStale(ByVal elemHandle As Long, _
                                ByVal elemGen As String, _
                                ByRef outStatus As String) As Boolean
    Dim js As String
    Dim res As String
    Dim gen As String

    outStatus = ""
    Elem_CheckStale = True

    ' Fast path (same logic as Elem_Read)
    If Len(m_domGen) > 0 Then
        If elemGen <> m_domGen Then Exit Function
    End If

    js = "(function(){" & JsRegistryPrologue() & _
         "var ok=(R.gen===" & JsQuote(elemGen) & ")&&(!!R.els[" & CStr(elemHandle) & "]);" & _
         "return {s:ok?'ok':'stale',g:R.gen};})()"

    res = EvalSync(js)

    If Not m_lastEvalOk Then
        outStatus = m_lastEvalError
        Exit Function
    End If

    If Wv2Json.JsonPickStr(res, "g", gen) Then
        m_domGen = gen
    End If

    Elem_CheckStale = (InStr(1, res, """s"":""ok""") = 0)
End Function


' ============================================================
' Elem_Exec (D-3, Friend) - fire one write / operation on an element
'
'   Arguments:
'     elemHandle  : index into R.els[] on the JS side
'     elemGen     : the generation at the time the element was obtained
'     actionStmts : a JS *statement* using variable e (the target element) (including the trailing semicolon).
'                   To return a clue about the result, put a string in variable rv.
'     outVal      : the content of rv (the path name). An empty string for operations that do not use it
'     outStatus   : failure reason. An empty string on success
'   Return value:
'     True  = the operation ran to the end on the JS side
'     False = stale or failure. The reason is in outStatus
'
'   *Success = "the JS ran without an exception", not "the intended effect occurred"*
'     For example, Click on a disabled button raises no exception but nothing happens.
'     To confirm the effect, check with a read (Elem_Read).
'
'   The handling of the fast path and generation cache is the same as Elem_Read (it only falls to the safe side).
' ============================================================
Friend Function Elem_Exec(ByVal elemHandle As Long, _
                          ByVal elemGen As String, _
                          ByVal actionStmts As String, _
                          ByRef outVal As String, _
                          ByRef outStatus As String) As Boolean
    Dim js As String
    Dim res As String
    Dim gen As String

    outVal = ""
    outStatus = ""
    Elem_Exec = False

    ' --- Fast path: if the generation in hand differs from the cache, stale can be determined without a round trip ---
    If Len(m_domGen) > 0 Then
        If elemGen <> m_domGen Then
            outStatus = "stale"
            Exit Function
        End If
    End If

    js = "(function(){" & JsRegistryPrologue() & _
         "if(R.gen!==" & JsQuote(elemGen) & "){return {s:'gen',g:R.gen};}" & _
         "var e=R.els[" & CStr(elemHandle) & "];" & _
         "if(!e){return {s:'nil',g:R.gen};}" & _
         "var rv='';" & actionStmts & _
         "return {s:'ok',g:R.gen,v:String(rv)};})()"

    res = EvalSync(js)

    If Not m_lastEvalOk Then
        outStatus = m_lastEvalError
        Exit Function
    End If

    If Wv2Json.JsonPickStr(res, "g", gen) Then
        m_domGen = gen
    End If

    If InStr(1, res, """s"":""ok""") = 0 Then
        outStatus = "stale"
        Exit Function
    End If

    Wv2Json.JsonPickStr res, "v", outVal
    Elem_Exec = True
End Function

' ============================================================
' *The body of D-3b: DOM condition waits*
'
'   The measurement of unknown 1 (2026-08-22) settled that "ExecuteScript does not wait for a Promise", so
'   the wait is implemented as *polling on the VBA side*, following the backbone of issue 3.
'   The option of placing a mechanism on the JS side (notifying completion by postMessage) is not adopted. It would
'   increase the pending tables, and it departs from the EvalSync pattern built in D-1.
' ============================================================


' ============================================================
' ImplicitWaitSec (D-3, the opt-in part of issue 2 option C)
'
'   When set above 0, QuerySelector / GetElementById retry *until found*
'   (the same style as SeleniumVBA's ImplicitMaxWait).
'
'   *The default is 0 (disabled)* For why it does not wait by default, see the header.
'   WaitFor / WaitGone are explicit waits, so they are *not affected by this setting*.
'
'   Usage example:
'     p.ImplicitWaitSec = 5      ' later searches persist for up to 5 seconds
'     p.ImplicitWaitSec = 0      ' restore
' ============================================================
Public Property Get ImplicitWaitSec() As Single
    ImplicitWaitSec = m_implicitWaitSec
End Property

Public Property Let ImplicitWaitSec(ByVal waitSec As Single)
    If waitSec < 0 Then waitSec = 0
    m_implicitWaitSec = waitSec
    Wv2Log.LogD "Wv2Pane.ImplicitWaitSec = " & waitSec & " sec"
End Property


' ============================================================
' WaitFor (D-3, the first-class API of issue 2 option C)
'   Waits *until an element matching the selector appears*.
'
'   Arguments:
'     selector   : CSS selector
'     timeoutSec : upper limit (seconds). Default 10. If 0 or less, rounded to the default
'   Return value:
'     Found                          -> Wv2Element
'     Did not appear in time         -> *Nothing + LastEvalOk = True* (normal)
'     Failure (bad selector etc.)    -> *Nothing + LastEvalOk = False* (reason in LastEvalError)
'
'   *Looks only at existence (issue 5, option A)*
'     Even display:none counts as "appeared". On an SPA, being in the DOM but invisible is
'     perfectly common. If you want to wait until visible, prepare a separate WaitVisible.
'
'   *On failure it returns immediately without waiting*
'     in-callback (spec fact 14) and selector syntax errors do not get fixed by waiting.
'     If it persisted here, calling from inside an event handler would freeze for timeoutSec.
'
'   *Do not break / step through while waiting* (spec fact 20)
'
'   *Do not call it right after opening a tab*
'     While there is no View yet, EvalSync fails with no-view, and by the policy above
'     *it returns Nothing immediately without waiting*. Right after opening a page, first
'     wait until EvalSync goes through (D4WaitPane in the verification is a working example).
'
'   Usage example:
'     Dim el As Wv2Element
'     Set el = p.WaitFor("#result", 5)
'     If el Is Nothing Then
'         Debug.Print IIf(p.LastEvalOk, "did not appear", "failed: " & p.LastEvalError)
'     End If
' ============================================================
Public Function WaitFor(ByVal selector As String, _
                        Optional ByVal timeoutSec As Single = WAIT_TIMEOUT_DEFAULT) As Wv2Element
    If timeoutSec <= 0 Then timeoutSec = WAIT_TIMEOUT_DEFAULT

    Set WaitFor = FindWithWait("document.querySelector(" & JsQuote(selector) & ")", _
                               "WaitFor", selector, timeoutSec)
End Function


' ============================================================
' WaitGone (D-3)
'   Waits *until no element matching the selector remains*.
'
'   Return value:
'     True                  ... it disappeared (or there was none from the start)
'     False + LastEvalOk=True  ... did not disappear in time (a normal "still there")
'     False + LastEvalOk=False ... failure (reason in LastEvalError)
'
'   *The registry is not used for the existence check* (D-3 post-hoc report)
'     Running FindOne would push the not-yet-vanished element into R.els
'     at every poll (about 70 pile up after waiting 10 seconds). WaitGone does not need
'     to grab it, so it uses a dedicated expression (ExistsNow) that returns only whether found is true.
'
'   Usage example:
'     If Not p.WaitGone(".spinner", 10) Then Debug.Print "still spinning"
' ============================================================
Public Function WaitGone(ByVal selector As String, _
                         Optional ByVal timeoutSec As Single = WAIT_TIMEOUT_DEFAULT) As Boolean
    Dim t0 As Single
    Dim found As Boolean
    Dim ok As Boolean

    WaitGone = False
    If timeoutSec <= 0 Then timeoutSec = WAIT_TIMEOUT_DEFAULT

    t0 = Timer
    Do
        found = ExistsNow(selector, ok)

        If Not ok Then
            ' A failure does not get fixed by waiting. Return with LastEvalOk = False as is.
            Debug.Print "Wv2Pane.WaitGone: failed (" & selector & ") err=" & m_lastEvalError
            Exit Function
        End If

        If Not found Then
            WaitGone = True
            Exit Function
        End If

        If ElapsedSec(t0) >= timeoutSec Then
            Debug.Print "Wv2Pane.WaitGone: timed out (" & selector & ") " & _
                        timeoutSec & " seconds passed and it did not disappear"
            Exit Function
        End If

        PollPause
    Loop
End Function


' ============================================================
' FindWithWait (D-3, Private) - repeat FindOne until found
'
'   If waitSec <= 0, try only once (= the behaviour up to D-2 exactly).
'   *Persist only when not found (Nothing + LastEvalOk=True)*
'   Give up immediately on failure (LastEvalOk=False).
' ============================================================
Private Function FindWithWait(ByVal finderExpr As String, _
                              ByVal apiName As String, _
                              ByVal arg As String, _
                              ByVal waitSec As Single) As Wv2Element
    Dim el As Wv2Element
    Dim t0 As Single

    t0 = Timer
    Do
        Set el = FindOne(finderExpr, apiName, arg, (waitSec > 0))

        If Not el Is Nothing Then
            Set FindWithWait = el
            Exit Function
        End If

        If Not m_lastEvalOk Then Exit Function
        If waitSec <= 0 Then Exit Function
        If ElapsedSec(t0) >= waitSec Then
            Debug.Print "Wv2Pane." & apiName & ": timed out (" & arg & ") " & _
                        waitSec & " seconds passed and it did not appear"
            Exit Function
        End If

        PollPause
    Loop
End Function


' ============================================================
' ExistsNow (D-3, Private) - check only whether an element exists, in one round trip
'
'   *Not registered in the registry* For WaitGone's polling (see the comment above).
'   outOk : whether the JS execution itself succeeded. If False the return value is meaningless.
' ============================================================
Private Function ExistsNow(ByVal selector As String, ByRef outOk As Boolean) As Boolean
    Dim js As String
    Dim res As String
    Dim gen As String

    ExistsNow = False
    outOk = False

    js = "(function(){" & JsRegistryPrologue() & _
         "var e=document.querySelector(" & JsQuote(selector) & ");" & _
         "return {g:R.gen,found:!!e};})()"

    res = EvalSync(js)
    If Not m_lastEvalOk Then Exit Function

    If Wv2Json.JsonPickStr(res, "g", gen) Then
        m_domGen = gen
    End If

    outOk = True
    ExistsNow = (InStr(1, res, """found"":true") > 0)
End Function


' ============================================================
' PollPause / ElapsedSec (D-3, Private) - polling pause and elapsed time
'
'   PollPause : spin DoEvents for POLL_INTERVAL_SEC (100ms).
'               *Sleep is not used* In order to keep WebView2's completion notices flowing,
'               the message pump must not be stopped during this time either.
'
'   ElapsedSec: *Timer returns to 0 at midnight* so a plain subtraction goes negative and
'               the wait loop never exits. Correct by adding the amount crossed
'               (D-3 post-hoc report. The real-machine verification ran at 1:53, so this is not someone else's problem).
' ============================================================
Private Sub PollPause()
    Dim t0 As Single

    t0 = Timer
    Do
        DoEvents
        If ElapsedSec(t0) >= POLL_INTERVAL_SEC Then Exit Do
    Loop
End Sub

Private Function ElapsedSec(ByVal sinceTimer As Single) As Single
    Dim d As Single

    d = Timer - sinceTimer
    If d < 0 Then d = d + 86400
    ElapsedSec = d
End Function

' ============================================================
' *The body of D-4a: in-page SPA probe*
' ============================================================


' ============================================================
' SpaProbeState (D-4a) - get the probe into shape, then read the state
'
'   Return value: the state JSON (on success). On failure an empty string + LastEvalOk = False
'
'   Contents of the JSON:
'     v  version number  h  healthy (True/False)   g  installation generation  rp number of rebuilds
'     f  running fetch   x  running XHR
'     q  ms since the last DOM change   qn ms since the last traffic completion
'     m  DOM changes counted  sm DOM changes excluded
'     n  traffic counted      sn traffic excluded
'     up ms since the probe was installed      ob whether the MutationObserver is alive
'     sda whether the DOM signal is armed      sdh its hit count
'     sna whether the traffic signal is armed  snh its hit count
'     sua whether the URL signal is armed      suh whether it matches now
'     lgN number of diagnostic log entries     lgOn whether the diagnostic log is enabled
'     il  *requests in flight now* (up to 5. In the form "xhr 1234ms https://...")
'     ifn the number of *new ones* among the requests in flight (the quiescence check looks at this)
'     ifo likewise the number of *lingering ones* (those that exceeded StaleInflightMs)
'
'   *A health check runs on every call, and if broken it is rebuilt*
'   Therefore the 2 round trips of "first Ensure, then Read" are unnecessary.
'
'   Usage example:
'     Debug.Print p.SpaProbeState
'     -> {""v"":1,""h"":true,""g"":1,...}
' ============================================================
Public Function SpaProbeState() As String
    Dim js As String

    js = "(function(){" & SpaProbeJs() & SpaProbeApplyJs() & _
         "return {v:P.v,h:HP(P),g:P.gen,rp:P.repairs," & _
         "f:P.fetchN,x:P.xhrN," & _
         "q:Math.round(N()-P.lastMut),qn:Math.round(N()-P.lastNet)," & _
         "m:P.mutN,sm:P.skipMut,n:P.netN,sn:P.skipNet," & _
         "up:Math.round(N()-P.since),ob:!!P.ob," & _
         "sda:!!P.sdArmed,sdh:P.sdHits,sna:!!P.snArmed,snh:P.snHits," & _
         "sua:!!P.suArmed," & _
         "suh:(P.suArmed&&String(location.href).indexOf(P.suPat)>=0)?1:0," & _
         "lgN:P.log.length,lgOn:!!P.logOn," & _
         "il:P.ifl.map(function(o){" & _
         "return o.w+' '+Math.round(N()-o.t)+'ms '+o.u;}).slice(0,5).join(' # ')," & _
         "ifn:P.ifl.filter(function(o){return (N()-o.t)<" & StaleMsJs() & ";}).length," & _
         "ifo:P.ifl.filter(function(o){return (N()-o.t)>=" & StaleMsJs() & ";}).length};})()"

    SpaProbeState = EvalSync(js)

    If Not m_lastEvalOk Then
        Wv2Log.LogW "Wv2Pane.SpaProbeState: failed err=" & m_lastEvalError
        SpaProbeState = ""
    End If
End Function


' ============================================================
' SpaProbeHealthy (D-4a) - a thin entry that looks only at whether it is healthy
'   *Returns False when it could not be determined* (safe side). If you need to distinguish, use LastEvalOk.
' ============================================================
Public Function SpaProbeHealthy() As Boolean
    Dim st As String

    st = SpaProbeState()
    If Not m_lastEvalOk Then Exit Function

    SpaProbeHealthy = (InStr(1, st, """h"":true") > 0)
End Function


' ============================================================
' SpaProbeReset (D-4a) - discard the probe and rebuild it
'
'   *Normally unnecessary* because the health check detects breakage and rebuilds automatically.
'   Use it when you want to return the counted values to zero, and when you want to verify the health check itself.
'
'   *It does not restore the original fetch / XHR* Removing the wrapper would break things if the framework
'   had grabbed our wrapper. Immediately afterwards it wraps the current window.fetch again.
'
'   *The implementation is "break the version number, then have it rebuilt"* (it does not set
'   window.__wv2p to null). That way the generation and rebuild count keep increasing, and the exclusion list is carried over.
'   Setting it to null would return the generation to 1, and we could no longer trace when it was rebuilt.
' ============================================================
Public Function SpaProbeReset() As Boolean
    Dim js As String

    js = "(function(){var o=window.__wv2p;if(o){o.v=-1;}" & SpaProbeJs() & _
         "return {g:P.gen,h:HP(P)};})()"

    EvalSync js
    SpaProbeReset = m_lastEvalOk

    If Not m_lastEvalOk Then
        Wv2Log.LogW "Wv2Pane.SpaProbeReset: failed err=" & m_lastEvalError
    Else
        Wv2Log.LogD "Wv2Pane.SpaProbeReset: rebuilt the probe"
    End If
End Function


' ============================================================
' SpaProbeJs (D-4a, Private) - JS that gets the probe into shape (extension of the self-generating prologue)
'
'   When run, the probe ends up bound to variable P.
'   Creates it if missing and *rebuilds it if broken*. The health check can be called with HP(P).
'
'   *JS string literals use single quotes* (project rule).
'   *What is counted*
'     fetch          ... +1 on call, -1 on resolve/failure
'     XHR            ... +1 on send, -1 on loadend (double subtraction is prevented with __wv2d)
'     MutationObserver ... childList / characterData / subtree / attributes
'   The exclusion lists (ignoreNet / ignoreSel) are *read by reference each time* so
'   they can be swapped without rebuilding (the entry for that is added in D-4b).
' ============================================================
Private Function SpaProbeJs() As String
    Dim js As String

    js = "var V=" & CStr(SPA_PROBE_VER) & ",N=function(){return performance.now();};"
    js = js & "var XP=(window.XMLHttpRequest?XMLHttpRequest.prototype:null);"
    js = js & "var HP=function(p){return !!p&&p.v===V" & _
                      "&&(typeof window.fetch!=='function'||window.fetch===p.fw)" & _
                      "&&(!XP||(XP.open===p.xo&&XP.send===p.xs))&&!!p.ob;};"
    js = js & "var P=window.__wv2p;"
    js = js & "if(!HP(P)){var prev=P;"

    ' --- Container of the new probe (the exclusion lists are carried over) ---
    js = js & "var q={v:V,gen:(prev?((prev.gen||0)+1):1)," & _
                      "repairs:(prev?((prev.repairs||0)+1):0)," & _
                      "since:N(),lastMut:N(),lastNet:N()," & _
                      "fetchN:0,xhrN:0,mutN:0,netN:0,skipMut:0,skipNet:0," & _
                      "sdArmed:false,sdSel:'',sdHits:0,sdObs:null," & _
                      "snArmed:false,snPat:'',snHits:0," & _
                      "suArmed:false,suPat:''," & _
                      "ifl:[],log:[],logOn:(prev&&prev.logOn?true:false)," & _
                      "ignoreNet:(prev&&prev.ignoreNet?prev.ignoreNet:[])," & _
                      "ignoreSel:(prev&&prev.ignoreSel?prev.ignoreSel:'')};"

    ' --- Exclusion check (read from q each time, so it can be swapped later) ---
    js = js & "var ig=function(u){var a=q.ignoreNet,i;" & _
                      "for(i=0;i<a.length;i++){if(u.indexOf(a[i])>=0){return true;}}return false;};"
    js = js & "var sk=function(t){if(!q.ignoreSel){return false;}" & _
                      "try{var e=(t&&t.nodeType===1)?t:(t?t.parentElement:null);" & _
                      "return !!(e&&e.closest(q.ignoreSel));}catch(x){return false;}};"

    ' --- D-4e: ledger of in-flight requests and the diagnostic log ---
    '   ifl is the list of "requests in flight now". *To see the identity of lingering requests*
    '   log is a ring buffer of recent events (disabled by default. Enable with logOn).
    js = js & "var lg=function(k,d){if(!q.logOn){return;}" & _
                      "q.log.push(Math.round(N())+' '+k+' '+String(d).slice(0,120));" & _
                      "if(q.log.length>200){q.log.shift();}};q.lg=lg;"
    js = js & "var ifA=function(u,w){var o={u:String(u).slice(0,150),t:N(),w:w};" & _
                      "q.ifl.push(o);lg('NET+',w+' '+o.u);return o;};"
    js = js & "var ifD=function(o){var i=q.ifl.indexOf(o);" & _
                      "if(i>=0){q.ifl.splice(i,1);}" & _
                      "lg('NET-',Math.round(N()-o.t)+'ms '+o.u);};"

    ' --- D-4c: hit detection of explicit signals and the arm family ---
    '   *On a hit, set lastMut to now* = take a whole quiet window again after the evidence
    '   (same idea as the reference product. Do not be satisfied with one poll).
    js = js & "var sn=function(u){if(q.snArmed&&q.snPat&&" & _
                      "String(u).indexOf(q.snPat)>=0){q.snHits++;q.lastMut=N();" & _
                      "lg('SIG','net hit '+String(u).slice(0,80));}};q.sn=sn;"
    js = js & "q.armUrl=function(pat){q.suPat=pat;q.suArmed=true;return 0;};"
    js = js & "q.armDom=function(sel){var n;" & _
                      "try{n=document.querySelector(sel);}catch(e){return -2;}" & _
                      "if(!n){return -1;}" & _
                      "if(q.sdObs){try{q.sdObs.disconnect();}catch(e){}}" & _
                      "q.sdSel=sel;q.sdHits=0;q.sdArmed=true;" & _
                      "q.sdObs=new MutationObserver(function(ms){" & _
                      "q.sdHits+=ms.length;q.lastMut=N();" & _
                      "lg('SIG','dom hit x'+ms.length);});" & _
                      "q.sdObs.observe(n," & _
                      "{childList:true,characterData:true,subtree:true,attributes:true});" & _
                      "return 0;};"
    js = js & "q.armNet=function(pat){q.snPat=pat;q.snHits=0;q.snArmed=true;return 0;};"
    js = js & "q.disarm=function(){if(q.sdObs){try{q.sdObs.disconnect();}catch(e){}}" & _
                      "q.sdObs=null;q.sdArmed=false;q.snArmed=false;q.suArmed=false;" & _
                      "return 0;};"

    ' --- Wrapping fetch (*wrap the current window.fetch*) ---
    js = js & "if(typeof window.fetch==='function'){"
    js = js & "var of=(prev&&prev.of&&prev.fw===window.fetch)?prev.of:window.fetch;q.of=of;"
    '   *Even for excluded requests, the explicit-signal check is done*
    '   Exclusion only means "remove from the quiescence check", and the armed marks are treated separately.
    js = js & "q.fw=function(u,o){" & _
                      "var r=(typeof u==='string')?u:((u&&u.url)?u.url:'');" & _
                      "var sp=ig(r);if(sp){q.skipNet++;}else{q.netN++;q.fetchN++;}" & _
                      "var en=sp?null:ifA(r,'fetch');" & _
                      "var d=function(){if(!sp){q.fetchN--;q.lastNet=N();ifD(en);}sn(r);};" & _
                      "return of.apply(this,arguments).then(" & _
                      "function(v){d();return v;},function(e){d();throw e;});};"
    js = js & "window.fetch=q.fw;}"

    ' --- Wrapping XHR ---
    js = js & "if(XP){"
    js = js & "var oo=(prev&&prev.oo&&prev.xo===XP.open)?prev.oo:XP.open;"
    js = js & "var os=(prev&&prev.os&&prev.xs===XP.send)?prev.os:XP.send;q.oo=oo;q.os=os;"
    js = js & "q.xo=function(m,u){try{this.__wv2u=String(u);}catch(e){}" & _
                      "return oo.apply(this,arguments);};"
    js = js & "q.xs=function(){var s=this;try{var u=s.__wv2u||'';" & _
                      "var sp=ig(u);if(sp){q.skipNet++;}else{q.netN++;q.xhrN++;}" & _
                      "var en=sp?null:ifA(u,'xhr');" & _
                      "s.addEventListener('loadend',function(){" & _
                      "if(!s.__wv2d){s.__wv2d=true;" & _
                      "if(!sp){q.xhrN--;q.lastNet=N();ifD(en);}sn(u);}});}catch(e){}" & _
                      "return os.apply(this,arguments);};"
    js = js & "XP.open=q.xo;XP.send=q.xs;}"

    ' --- Observing DOM changes ---
    js = js & "try{q.ob=new MutationObserver(function(ms){" & _
                      "var i,c=0,t;for(i=0;i<ms.length;i++){t=ms[i].target;" & _
                      "if(ms[i].type==='characterData'){t=(t&&t.parentElement)?t.parentElement:t;}" & _
                      "if(sk(t)){q.skipMut++;}else{c++;}}" & _
                      "if(c>0){q.mutN+=c;q.lastMut=N();lg('DOM','x'+c);}});"
    js = js & "q.ob.observe(document," & _
                      "{childList:true,characterData:true,subtree:true,attributes:true});" & _
                      "}catch(e){q.ob=null;}"

    js = js & "window.__wv2p=q;P=q;}"

    SpaProbeJs = js
End Function

' ============================================================
' *The body of D-4b: waiting for quiescence*
' ============================================================


' ============================================================
' WaitSettled (D-4b) - wait until traffic and DOM changes settle
'
'   Arguments:
'     timeoutSec : upper limit (seconds). Default 10. If 0 or less, rounded to the default
'     stableMs   : quiet window (ms). Default 500. If the DOM does not move for this long it is regarded as quiescent
'   Return value:
'     True                     ... reached quiescence
'     False + LastEvalOk=True  ... did not quiet down in time (normal)
'     False + LastEvalOk=False ... failure (reason in LastEvalError)
'
'   *** True is not proof that "the app has finished processing" ***
'     It is merely quiet. It cannot be told whether it is quiet because it was ignored or because it finished.
'     For important operations, combine it with D-4c's explicit signal.
'
'   *Do not break / step through while waiting* (spec fact 20)
'     Wait 10 seconds and you are in the DoEvents section for the full 10 seconds.
'
'   Usage example:
'     el.Click
'     If p.WaitSettled(10) Then
'         Set r = p.QuerySelector("#result")
'     Else
'         Debug.Print "did not settle: " & p.LastSettleInfo
'     End If
' ============================================================
Public Function WaitSettled(Optional ByVal timeoutSec As Single = WAIT_TIMEOUT_DEFAULT, _
                            Optional ByVal stableMs As Long = SETTLE_STABLE_MS_DEFAULT) As Boolean
    Dim t0 As Single
    Dim st As String
    Dim inflight As Long
    Dim quietMs As Long
    Dim quietEff As Long
    Dim elapsedMs As Long
    Dim polls As Long
    Dim sigOk As Boolean
    Dim inflightList As String
    Dim stale As Long

    WaitSettled = False
    m_lastSettle = ""
    If timeoutSec <= 0 Then timeoutSec = WAIT_TIMEOUT_DEFAULT
    If stableMs < 0 Then stableMs = SETTLE_STABLE_MS_DEFAULT

    t0 = Timer
    Do
        polls = polls + 1
        st = SpaProbeState()

        If Not m_lastEvalOk Then
            ' A failure does not get fixed by waiting (design principle 104).
            m_lastSettle = "error " & m_lastEvalError
            Wv2Log.LogW "Wv2Pane.WaitSettled: failed err=" & m_lastEvalError
            m_sigDom = False
            m_sigNet = False
            m_sigUrl = False
            Exit Function
        End If

        ' *Lingering requests are not counted* (D-4f)
        '   ifn = the number of requests in flight that started within StaleInflightMs.
        '   A realistic compromise so as not to keep waiting on long polling or streaming, and
        '   it does not contradict the meaning of WaitSettled, which is that
        '   *it judges "looks quiet", not "everything is finished"*.
        '   If accuracy is needed, use AddIgnoreNetwork or an explicit signal.
        inflight = Wv2Json.JsonGetNum(st, "ifn")
        stale = Wv2Json.JsonGetNum(st, "ifo")
        quietMs = Wv2Json.JsonGetNum(st, "q")

        ' --- D-4c: look at the state of the armed signals ---
        '   *Should be armed but the page side is not armed = it vanished, e.g. through navigation*
        '   Neither silently re-arm nor ignore; return it as a failure (fail-visible).
        If m_sigDom Or m_sigNet Or m_sigUrl Then
            If (m_sigDom And InStr(1, st, """sda"":true") = 0) Or _
               (m_sigNet And InStr(1, st, """sna"":true") = 0) Or _
               (m_sigUrl And InStr(1, st, """sua"":true") = 0) Then
                m_lastSettle = "signal-lost (the arm vanished, e.g. through navigation)"
                Wv2Log.LogW "Wv2Pane.WaitSettled: *signal-lost* " & _
                            "The armed signal has vanished from the page side. " & _
                            "It may have crossed a navigation (arm again)"
                DisarmSignals
                Exit Function
            End If
        End If

        sigOk = True
        If m_sigDom Then
            If Wv2Json.JsonGetNum(st, "sdh") = 0 Then sigOk = False
        End If
        If m_sigNet Then
            If Wv2Json.JsonGetNum(st, "snh") = 0 Then sigOk = False
        End If
        If m_sigUrl Then
            If Wv2Json.JsonGetNum(st, "suh") = 0 Then sigOk = False
        End If

        ' --- *The quiet window is measured from "the moment of the call"* (D-4b post-hoc report) ---
        '   The page-side q is "the time since the last DOM change", so if it was
        '   already quiet before the call, the condition is met immediately on the first round trip. Then
        '   *"has not started yet" and "already finished" cannot be distinguished*
        '   (it is perfectly common for an operation's effect to start tens of ms late).
        '   So we cap it by the time elapsed since the call, and *watch for at least stableMs*.
        '   If anything moves during this time q resets, so even a late start can be caught.
        '   However, processing that starts after stableMs is in principle missed.
        '   That is the job of D-4c's explicit signal (arm).
        elapsedMs = CLng(ElapsedSec(t0) * 1000)
        quietEff = quietMs
        If quietEff > elapsedMs Then quietEff = elapsedMs

        If inflight = 0 And quietEff >= stableMs And sigOk Then
            m_lastSettle = "stable q=" & quietMs & "ms(eff=" & quietEff & ") slack=" & _
                           (quietEff - stableMs) & "ms polls=" & polls & _
                           " elapsed=" & Format$(ElapsedSec(t0), "0.00") & "s" & _
                           IIf(stale > 0, " (" & stale & " lingering excluded)", "") & _
                           SigInfoOf(st)
            WaitSettled = True
            DisarmSignals
            Exit Function
        End If

        If ElapsedSec(t0) >= timeoutSec Then
            ' *If the cause of the timeout is traffic, show what is in flight* (lesson of D-4d)
            '   On Maps, one endless XHR lingered and quiescence never came.
            '   If the URL is known, it can be removed with AddIgnoreNetwork.
            Wv2Json.JsonPickStr st, "il", inflightList

            m_lastSettle = "timeout inflight=" & inflight & _
                           IIf(stale > 0, "(" & stale & " lingering excluded)", "") & _
                           " q=" & quietMs & _
                           "ms (stableMs=" & stableMs & ") polls=" & polls & _
                           SigInfoOf(st)
            If Len(inflightList) > 0 Then
                m_lastSettle = m_lastSettle & " | in flight: " & inflightList
            End If
            Wv2Log.LogD "Wv2Pane.WaitSettled: timed out " & m_lastSettle
            DisarmSignals
            Exit Function
        End If

        PollPause
    Loop
End Function


' ============================================================
' LastSettleInfo (D-4b) - breakdown of the previous wait
'
'   On success: stable q=<quiet ms> slack=<spare ms> polls=<round trips> elapsed=<seconds>
'   On failure: timeout inflight=<traffic in flight> q=<quiet ms> ... / error <reason>
'
'   *A wait with small slack (under 100 ms) is "borderline"* Knowledge from the reference product.
'   It may just have slipped into the quiet window by luck, so consider raising stableMs or
'   combining with an explicit signal (D-4c).
' ============================================================
Public Property Get LastSettleInfo() As String
    LastSettleInfo = m_lastSettle
End Property


' ============================================================
' AutoWaitAfterAction (D-4b, the opt-in part of issue 7 option C)
'
'   When True, WaitSettled is called automatically *after Wv2Element's .Value = / .Click / .SetAttribute succeeds*.
'   Default False.
'
'   *The D-3 contract does not change* The operation's return value stays "the JS ran without an exception", and
'   the success of the wait is checked by yourself via p.LastSettleInfo / p.WaitSettled.
'   The same "explicit is first-class, automatic is opt-in" style as ImplicitWaitSec (D-3).
' ============================================================
Public Property Get AutoWaitAfterAction() As Boolean
    AutoWaitAfterAction = m_autoWait
End Property

Public Property Let AutoWaitAfterAction(ByVal onOff As Boolean)
    m_autoWait = onOff
    Wv2Log.LogD "Wv2Pane.AutoWaitAfterAction = " & onOff
End Property


' ============================================================
' AddIgnoreNetwork / ClearIgnoreNetwork / IgnoreSelectors (D-4b, issue 5)
'
'   Specify what to remove from the quiescence check. *This is not blocking* Requests fly as normal.
'
'   AddIgnoreNetwork : *substring-match* pattern of URL. For measurement systems that are always chattering
'                      Example: p.AddIgnoreNetwork "google-analytics"
'   IgnoreSelectors  : CSS selectors (comma-separated) of elements whose DOM changes are ignored
'                      Example: p.IgnoreSelectors = ".clock, .ticker, #ads"
'                      *Judged with closest* so changes in descendants of the specified element are also excluded
'
'   *The lists are held on the VBA side* because placing them on the page side would lose them on navigation.
'   They are pushed along with each state read, so round trips do not increase.
'
'   *** An exclusion that is too broad produces "too-early quiescence" ***
'     If you exclude even meaningful traffic, it judges quiescence while processing is still going on.
'     Keep exclusions to "only those you looked at in the log and confirmed to be unrelated".
' ============================================================
Public Function AddIgnoreNetwork(ByVal urlPattern As String) As Boolean
    Dim v As Variant

    If Len(urlPattern) = 0 Then Exit Function
    If m_ignoreNet Is Nothing Then Set m_ignoreNet = New Collection

    For Each v In m_ignoreNet
        If CStr(v) = urlPattern Then
            AddIgnoreNetwork = True
            Exit Function
        End If
    Next v

    m_ignoreNet.Add urlPattern
    Wv2Log.LogD "Wv2Pane.AddIgnoreNetwork: " & urlPattern & _
                " (total " & m_ignoreNet.Count & " entries)"
    AddIgnoreNetwork = True
End Function

Public Sub ClearIgnoreNetwork()
    Set m_ignoreNet = New Collection
    Wv2Log.LogD "Wv2Pane.ClearIgnoreNetwork: emptied the exclusion list"
End Sub

Public Property Get IgnoreSelectors() As String
    IgnoreSelectors = m_ignoreSel
End Property

Public Property Let IgnoreSelectors(ByVal cssSelectors As String)
    m_ignoreSel = cssSelectors
    Wv2Log.LogD "Wv2Pane.IgnoreSelectors = " & cssSelectors
End Property


' ============================================================
' SpaProbeApplyJs (D-4b, Private) - JS that pushes the exclusion lists to the page
'
'   Rides on the same round trip as reading the state. It assumes variable P (the probe)
'   is already bound. *Always sent even when empty* (if not sent, the previous setting remains on the page side).
' ============================================================
Private Function SpaProbeApplyJs() As String
    Dim js As String
    Dim arr As String
    Dim v As Variant

    If Not m_ignoreNet Is Nothing Then
        For Each v In m_ignoreNet
            If Len(arr) > 0 Then arr = arr & ","
            arr = arr & JsQuote(CStr(v))
        Next v
    End If

    ' *Always send even when empty* (stepped on during D-4b's real-machine verification)
    '   If we "do not send because it is empty", *the previous setting keeps remaining on the page side*.
    '   IgnoreSelectors = "" and ClearIgnoreNetwork stop taking effect,
    '   and something excluded once never comes back. Overwriting every time makes the VBA side the sole authority.
    js = "P.ignoreNet=[" & arr & "];P.ignoreSel=" & _
         JsQuote(m_ignoreSel) & ";"

    SpaProbeApplyJs = js
End Function

' ============================================================
' *D-4c: explicit signals (arm, then act)*
'
'   *Quiescence alone cannot distinguish "the app finished" from "the app ignored it"*
'   because both are quiet. For important operations, *arm a completion marker right before the operation*,
'   and wait until it is observed at the next WaitSettled.
'
'   # One-shot
'     An armed signal is *consumed by the next single WaitSettled*. If you do the same operation twice,
'     arm again twice. Consumption happens on success and on timeout alike.
'
'   # Re-measure the quiet window on a hit
'     Do not be satisfied with one poll right after the evidence arrives; *take a whole
'     quiet window again after that* (same idea as the reference product).
'
'   # *If the arm vanishes, do not silently fix it (fail-visible)*
'     When the probe vanishes on page navigation, the arm vanishes with it. Waiting as is
'     would turn into a "wait with no condition" and return success the moment it goes quiet.
'     The VBA side also remembers that it is armed, and if they disagree it fails with signal-lost.
'
'   # Relation to exclusions
'     *Even for excluded requests, the explicit-signal check is done* Exclusion only means
'     "remove from the quiescence check", not "pretend we did not see it".
'
'   Usage example (arm, then act):
'     p.ArmContentSignal "#result-table"     ' wait for the existing container to be rewritten
'     el.Click
'     If Not p.WaitSettled(10) Then Debug.Print p.LastSettleInfo
' ============================================================


' ============================================================
' ArmContentSignal (D-4c) - wait for an existing subtree to be rewritten
'
'   *False if the target does not exist at the moment of arming* (fail-fast)
'   Lets you notice on the spot, rather than after waiting, an accident such as a misspelled selector
'   or specifying an element that has not been created yet.
'   *An element that will be created after the operation cannot be specified* Specify an existing parent.
'
'   Return value: True = armed / False = no target, invalid selector, or communication failure
'           The reason is in LastEvalError (not-found / bad-selector / other)
' ============================================================
Public Function ArmContentSignal(ByVal selector As String) As Boolean
    Dim js As String
    Dim res As String

    js = "(function(){" & SpaProbeJs() & SpaProbeApplyJs() & _
         "return {r:P.armDom(" & JsQuote(selector) & ")};})()"

    res = EvalSync(js)
    If Not m_lastEvalOk Then
        Wv2Log.LogW "Wv2Pane.ArmContentSignal: failed err=" & m_lastEvalError
        Exit Function
    End If

    ' *JsonGetNum does not handle negative numbers (design principle 64)* so read the return value as a string.
    If InStr(1, res, """r"":0") = 0 Then
        If InStr(1, res, "-2") > 0 Then
            m_lastEvalError = "bad-selector"
        Else
            m_lastEvalError = "not-found"
        End If
        m_lastEvalOk = False
        Wv2Log.LogW "Wv2Pane.ArmContentSignal: cannot arm (" & selector & _
                    ") reason=" & m_lastEvalError
        Exit Function
    End If

    m_sigDom = True
    Wv2Log.LogD "Wv2Pane.ArmContentSignal: armed (" & selector & ")"
    ArmContentSignal = True
End Function


' ============================================================
' ArmNetworkSignal (D-4c) - wait for the completion of a specific request
'
'   urlPattern is a *substring match* of the URL. A hit is when the response comes back.
'   *Existence cannot be checked* (because it is a request that has yet to fly). If you misspell it,
'   it times out without a hit. net:miss appears in LastSettleInfo, and
'   in that case look at the actual URL in D-4d's diagnostic log and fix it.
' ============================================================
Public Function ArmNetworkSignal(ByVal urlPattern As String) As Boolean
    Dim js As String

    If Len(urlPattern) = 0 Then Exit Function

    js = "(function(){" & SpaProbeJs() & SpaProbeApplyJs() & _
         "return {r:P.armNet(" & JsQuote(urlPattern) & ")};})()"

    EvalSync js
    If Not m_lastEvalOk Then
        Wv2Log.LogW "Wv2Pane.ArmNetworkSignal: failed err=" & m_lastEvalError
        Exit Function
    End If

    m_sigNet = True
    Wv2Log.LogD "Wv2Pane.ArmNetworkSignal: armed (" & urlPattern & ")"
    ArmNetworkSignal = True
End Function


' ============================================================
' DisarmSignals (D-4c) - cancel the arm
'   Called automatically when WaitSettled ends (one-shot).
'   Call it by hand if you call off the operation after arming.
' ============================================================
Public Function DisarmSignals() As Boolean
    m_sigDom = False
    m_sigNet = False
    m_sigUrl = False

    EvalSync "(function(){var p=window.__wv2p;" & _
             "if(p&&p.disarm){p.disarm();}return 1;})()"

    DisarmSignals = m_lastEvalOk
End Function


' ============================================================
' SigInfoOf (D-4c, Private) - hit or miss of the signals to put in the wait breakdown
' ============================================================
Private Function SigInfoOf(ByVal st As String) As String
    Dim r As String

    If m_sigDom Then
        r = r & " dom:"
        If Wv2Json.JsonGetNum(st, "sdh") > 0 Then r = r & "hit" Else r = r & "miss"
    End If

    If m_sigNet Then
        r = r & " net:"
        If Wv2Json.JsonGetNum(st, "snh") > 0 Then r = r & "hit" Else r = r & "miss"
    End If

    If m_sigUrl Then
        r = r & " url:"
        If Wv2Json.JsonGetNum(st, "suh") > 0 Then r = r & "hit" Else r = r & "miss"
    End If

    If Len(r) > 0 Then r = " signal=" & Trim$(r)
    SigInfoOf = r
End Function

' ============================================================
' ArmUrlSignal (D-4e) - wait until the URL contains the specified string
'
'   *Found necessary by D-4d's measurement* The completion of an SPA may appear neither in the DOM nor in traffic but
'   *in the URL*. Google Maps address search is exactly that:
'   once the search goes through, the URL changes to /place/... and /@latitude,longitude,zoom.
'   This is both a "completion marker" and "the very coordinates we wanted".
'
'   Changes by pushState cannot be observed as DOM changes, so
'   it is made so that *location.href is looked at on every state read*.
'
'   One-shot. Consumed by the next single WaitSettled.
'   *Even if it already matches at the moment of arming, it counts as a hit* so
'   specify a string that is not contained in the pre-navigation URL.
'
'   Usage example:
'     p.ArmUrlSignal "/place/"
'     el.Click
'     If p.WaitSettled(15) Then Debug.Print p.View_GetSource()
' ============================================================
Public Function ArmUrlSignal(ByVal urlFragment As String) As Boolean
    Dim js As String

    If Len(urlFragment) = 0 Then Exit Function

    js = "(function(){" & SpaProbeJs() & SpaProbeApplyJs() & _
         "return {r:P.armUrl(" & JsQuote(urlFragment) & ")};})()"

    EvalSync js
    If Not m_lastEvalOk Then
        Wv2Log.LogW "Wv2Pane.ArmUrlSignal: failed err=" & m_lastEvalError
        Exit Function
    End If

    m_sigUrl = True
    Wv2Log.LogD "Wv2Pane.ArmUrlSignal: armed (" & urlFragment & ")"
    ArmUrlSignal = True
End Function


' ============================================================
' SpaProbeLogging (D-4e) - diagnostic log ON / OFF
'
'   *Default is False* When enabled, the page side accumulates the most recent 200 events:
'     NET+ <fetch|xhr> <URL>     a request started
'     NET- <elapsed ms> <URL>    a request finished
'     DOM  x<count>              a burst of DOM changes
'     SIG  <kind> hit ...        an armed signal hit
'
'   *A tool for deciding what to wait for by observation rather than guesswork*
'   It corresponds to the Discovery Log of the reference product (WebDriver BiDi for SeleniumVBA).
'   Run once without arming, look at the URLs and DOM movements that appeared in the log, and decide what to arm next /
'   what to exclude.
'
'   *Do not leave it ON all the time* The page side gets heavier by the amount accumulated.
' ============================================================
Public Property Get SpaProbeLogging() As Boolean
    Dim st As String

    st = SpaProbeState()
    If Not m_lastEvalOk Then Exit Property

    SpaProbeLogging = (InStr(1, st, """lgOn"":true") > 0)
End Property

Public Property Let SpaProbeLogging(ByVal onOff As Boolean)
    Dim js As String

    js = "(function(){" & SpaProbeJs() & _
         "P.logOn=" & IIf(onOff, "true", "false") & ";" & _
         "if(!P.logOn){P.log=[];}return {r:1};})()"

    EvalSync js
    Wv2Log.LogD "Wv2Pane.SpaProbeLogging = " & onOff & _
                IIf(m_lastEvalOk, "", " (failed: " & m_lastEvalError & ")")
End Property


' ============================================================
' SpaProbeDrainLog (D-4e) - take out the accumulated diagnostic log and flush it to Wv2Log
'
'   *After taking out, the page side becomes empty* (does not keep accumulating).
'   Return value: the number flushed. 0 means nothing had accumulated.
'
'   There is a limit to the amount one EvalSync can carry, so *up to 60 per call*.
'   The rest can be taken at the next call (if the return value is 60, more may remain).
' ============================================================
Public Function SpaProbeDrainLog() As Long
    Dim js As String
    Dim res As String
    Dim txt As String
    Dim lines As Variant
    Dim i As Long

    js = "(function(){" & SpaProbeJs() & _
         "var a=P.log.splice(0,60);" & _
         "return {n:a.length,s:a.join(String.fromCharCode(1))};})()"

    res = EvalSync(js, 10)
    If Not m_lastEvalOk Then
        Wv2Log.LogW "Wv2Pane.SpaProbeDrainLog: failed err=" & m_lastEvalError
        Exit Function
    End If

    SpaProbeDrainLog = Wv2Json.JsonGetNum(res, "n")
    If SpaProbeDrainLog = 0 Then Exit Function

    Wv2Json.JsonPickStr res, "s", txt
    lines = Split(txt, ChrW$(1))

    Wv2Log.LogI "  --- diagnostic log (" & SpaProbeDrainLog & " entries) ---"
    For i = LBound(lines) To UBound(lines)
        Wv2Log.LogI "    " & lines(i)
    Next i
End Function

' ============================================================
' StaleInflightMs (D-4f) - the boundary (ms) for excluding lingering requests from the quiescence check
'
'   *Default 10000 (10 seconds)* Setting 0 disables it (count all requests).
'
'   # Why it is needed (D-4d's measurement)
'     Google Maps address search keeps /search?tbm=map *open* and keeps
'     streaming results. The search itself finished in 2 seconds, the DOM went quiet and the URL changed,
'     yet because of this one request inflight never reached 0, and the quiescence wait timed out.
'     Long polling and streaming exist on every site, so relying only on individual exclusions
'     (AddIgnoreNetwork) means stepping on it at least once on any unfamiliar site.
'
'   # What is being discarded (honestly)
'     *Even a request that is truly just slow stops being counted once it exceeds this time*
'     But at that point the DOM has not moved during the quiet window either, so
'     it does not contradict the meaning of WaitSettled, "looks settled"
'     (keeping the premise that True is not proof that "the app has finished processing").
'     In situations that need accuracy, exclude individually with AddIgnoreNetwork, or
'     write "what to wait for" with D-4c's explicit signal.
' ============================================================
Public Property Get StaleInflightMs() As Long
    StaleInflightMs = m_staleMs
End Property

Public Property Let StaleInflightMs(ByVal ms As Long)
    If ms < 0 Then ms = 0
    m_staleMs = ms
    Wv2Log.LogD "Wv2Pane.StaleInflightMs = " & ms & _
                IIf(ms = 0, " (disabled: count all)", "")
End Property


' ============================================================
' StaleMsJs (D-4f, Private) - turn the lingering boundary into a JS numeric literal
'   When 0 (disabled), use a huge value so that "everything is new".
' ============================================================
Private Function StaleMsJs() As String
    If m_staleMs <= 0 Then
        StaleMsJs = "1e15"
    Else
        StaleMsJs = CStr(m_staleMs)
    End If
End Function

' ============================================================
' *The body of D-6: QuerySelectorAll and lifetime management*
' ============================================================


' ============================================================
' QuerySelectorAll (D-6) - grab all matching elements at once
'
'   Arguments:
'     selector : CSS selector
'     maxCount : upper limit to return (default 200). If 0 or less, rounded to the default
'   Return value:
'     A Collection of Wv2Element. *If none is found, an empty Collection* (not Nothing)
'     It returns an empty Collection on failure too. Distinguish with LastEvalOk (design principle 93).
'
'   *If the count reaches the upper limit, LastAllTruncated = True*
'     So that no one assumes "the returned count = the number on the page".
'
'   *Returned in document order* Exactly the order of querySelectorAll.
'
'   Usage example:
'     Dim els As Collection, el As Wv2Element
'     Set els = p.QuerySelectorAll("a[href]")
'     Debug.Print els.Count & " items"
'     For Each el In els
'         Debug.Print el.GetAttribute("href")
'     Next el
'
'   *One round trip per property* (the D-2 compromise), so reading all attributes from
'   200 elements takes 200 round trips = about 8 seconds. When you want to fetch them in bulk, writing
'   a single JS expression with EvalSync is faster.
' ============================================================
Public Function QuerySelectorAll(ByVal selector As String, _
                                 Optional ByVal maxCount As Long = ALL_MAX_DEFAULT) As Collection
    Dim js As String
    Dim res As String
    Dim gen As String
    Dim handles As String
    Dim parts As Variant
    Dim i As Long
    Dim el As Wv2Element
    Dim outCol As Collection

    Set outCol = New Collection
    Set QuerySelectorAll = outCol
    m_allTruncated = False

    If maxCount <= 0 Then maxCount = ALL_MAX_DEFAULT

    ' *Just changed D-2's FindOne from push to a loop*
    '   The handles are returned as a comma-separated string (because Wv2Json cannot handle arrays.
    '   Split is smaller than extending the parser - issue 5).
    js = "(function(){" & JsRegistryPrologue() & _
         "var q=document.querySelectorAll(" & JsQuote(selector) & ");" & _
         "var mx=" & CStr(maxCount) & ",h=[],i;" & _
         "for(i=0;i<q.length&&i<mx;i++){R.els.push(q[i]);h.push(R.els.length-1);}" & _
         "return {g:R.gen,n:q.length,k:h.length,cut:(q.length>mx)," & _
         "h:h.join(','),tot:R.els.length};})()"

    res = EvalSync(js)

    If Not m_lastEvalOk Then
        Debug.Print "Wv2Pane.QuerySelectorAll: failed (" & selector & ") err=" & m_lastEvalError
        Exit Function
    End If

    If Wv2Json.JsonPickStr(res, "g", gen) Then
        m_domGen = gen
    End If

    m_allTruncated = (InStr(1, res, """cut"":true") > 0)
    If m_allTruncated Then
        Wv2Log.LogW "Wv2Pane.QuerySelectorAll: *truncated* (" & selector & _
                    ") the page has " & Wv2Json.JsonGetNum(res, "n") & _
                    " items but stopped at " & maxCount & " items"
    End If

    Wv2Json.JsonPickStr res, "h", handles
    If Len(handles) = 0 Then
        Debug.Print "Wv2Pane.QuerySelectorAll: 0 items (" & selector & ")"
        WarnIfElementsPiled Wv2Json.JsonGetNum(res, "tot")
        Exit Function
    End If

    parts = Split(handles, ",")
    For i = LBound(parts) To UBound(parts)
        Set el = New Wv2Element
        el.InitElement Me, CLng(parts(i)), gen
        outCol.Add el
    Next i

    WarnIfElementsPiled Wv2Json.JsonGetNum(res, "tot")

    Debug.Print "Wv2Pane.QuerySelectorAll: " & outCol.Count & " items (" & _
                selector & ") gen=" & gen
End Function


' ============================================================
' LastAllTruncated (D-6) - whether the previous QuerySelectorAll was truncated at the limit
'   *If True, "the returned count = the number on the page" does not hold*
'   If you need all of them, raise maxCount or narrow the selector.
' ============================================================
Public Property Get LastAllTruncated() As Boolean
    LastAllTruncated = m_allTruncated
End Property


' ============================================================
' ElementCount (D-6) - the number of elements stacked in the JS-side registry
'
'   *Grabbed elements are not released* (D-2 issue 3). The registry vanishes wholesale on page navigation, so
'   actual harm is small, but QuerySelectorAll can stack 200 at a time, so it grows fast.
'   If concerned, look at the count with this and rebuild with ClearElementRegistry.
'   *Rebuilding makes all Wv2Element objects in hand stale*
'
'   Return value: the number stacked. -1 on failure (reason in LastEvalError)
' ============================================================
Public Property Get ElementCount() As Long
    Dim res As String

    res = EvalSync("(function(){" & JsRegistryPrologue() & _
                  "return {n:R.els.length};})()")

    If Not m_lastEvalOk Then
        ElementCount = -1
        Exit Property
    End If

    ElementCount = Wv2Json.JsonGetNum(res, "n")
End Property


' ============================================================
' WarnIfElementsPiled (D-6, Private) - report the pile-up
'   *Does not discard* Discarding on its own would silently make the Wv2Element in hand stale.
'   It only reports, and the caller decides whether to discard.
' ============================================================
Private Sub WarnIfElementsPiled(ByVal total As Long)
    If total < ELEMENT_WARN_COUNT Then Exit Sub

    Wv2Log.LogW "Wv2Pane: *" & total & " items are piled up in the element registry* " & _
                "(by design grabbed elements are not released). If not needed, it can be rebuilt with ClearElementRegistry" & _
                ". However, all Wv2Element objects in hand become stale."
End Sub



' ============================================================
' *N-1 : capturing network requests (WebResourceRequested)*
'
'   Everything from here down is the set added in stage N-1. The order is
'     (1) entries for the user   NetCaptureStart / NetCaptureStop / NetLogDrain etc.
'     (2) COM wiring   View_AddWebResourceRequested and 4 others
'     (3) receiver       View_OnWebResourceRequested (+ in-callback guard)
'     (4) ring       NetRingAlloc / NetPush / NetElapsedMs / NetRingIndex
'     (5) kind names   NetCtxName / NetCtxFromName
'
'   Usage (minimal):
'     p.NetCaptureStart          ' *If called before Navigate, the first DOCUMENT is caught too*
'     ... operate by hand once ...
'     p.NetLogDrain              ' flush what has accumulated to the log file and empty it
'     p.NetCaptureStop
' ============================================================


' ============================================================
' NetCaptureStart (N-1) - start capturing
'
'   contexts  : specify the kinds to catch, comma-separated. If empty, DOCUMENT,XHR,FETCH.
'               Names or numbers both work. "ALL" means all kinds (images and CSS come too).
'                 DOCUMENT / STYLESHEET / IMAGE / MEDIA / FONT / SCRIPT /
'                 XHR (= XMLHTTPREQUEST) / FETCH / TEXTTRACK / EVENTSOURCE /
'                 WEBSOCKET / MANIFEST / SIGNEDEXCHANGE / PING /
'                 CSPVIOLATIONREPORT / OTHER / ALL
'               *Numbers are also accepted* so that it can be used without fixing this even if the SDK adds kinds.
'   uriFilter : URL filter. Default "*" (all). * and ? can be used.
'
'   *If already capturing, stop once and re-set up*
'     This prevents the accident where you thought you changed the kinds but the previous net also remains (in the spirit of design principle 106).
'
'   *If not a single filter could be set, remove the handler too and return False*
'     With zero filters not one event arrives, so we do not create the worst state of
'     "capturing but nothing shows up".
'
'   Return value: True if capturing.
' ============================================================
Public Function NetCaptureStart(Optional ByVal contexts As String = "", _
                                Optional ByVal uriFilter As String = "*") As Boolean
    Dim parts As Variant
    Dim i     As Long
    Dim ctx   As Long
    Dim added As Long
    Dim hr    As Long
    Dim spec  As String

    If m_closed Then
        Wv2Log.LogW "Wv2Pane.NetCaptureStart: this pane is already closed"
        Exit Function
    End If
    If m_pView = 0 Then
        Wv2Log.LogW "Wv2Pane.NetCaptureStart: m_pView is not acquired (create a tab first)"
        Exit Function
    End If

    If m_netOn Then NetCaptureStop

    spec = Trim$(contexts)
    If LenB(spec) = 0 Then spec = NET_CTX_DEFAULT
    If LenB(Trim$(uriFilter)) = 0 Then uriFilter = "*"

    ' --- (1) Register the handler ---
    hr = View_AddWebResourceRequested()
    If hr <> 0 Then
        Wv2Log.LogE "Wv2Pane.NetCaptureStart: add_WebResourceRequested failed hr=&H" & Hex(hr)
        Exit Function
    End If

    ' --- (2) Set up the filters (*if none is set, not one arrives*) ---
    parts = Split(spec, ",")
    For i = LBound(parts) To UBound(parts)
        ctx = NetCtxFromName(CStr(parts(i)))
        If ctx < 0 Then
            Wv2Log.LogW "Wv2Pane.NetCaptureStart: unknown kind [" & Trim$(CStr(parts(i))) & _
                        "], skipping"
        Else
            hr = View_AddWebResourceRequestedFilter(uriFilter, ctx)
            If hr <> 0 Then
                Wv2Log.LogW "Wv2Pane.NetCaptureStart: filter registration failed " & _
                            NetCtxName(ctx) & " hr=&H" & Hex(hr)
            Else
                m_netFilters.Add uriFilter & vbTab & ctx
                added = added + 1
            End If
        End If
    Next i

    If added = 0 Then
        Wv2Log.LogE "Wv2Pane.NetCaptureStart: *not a single filter could be set* " & _
                    "In this state not one event arrives, so the handler is removed too and it is aborted."
        View_RemoveWebResourceRequested
        Exit Function
    End If

    ' --- (3) Also set up the response event (N-3) ---
    '   *This event has no filter* so everything arrives. Only those that match a request
    '   in the list are recorded (issue 1, option A). In an environment where the _2 IF cannot be obtained,
    '   *the status column stays blank forever* so do not proceed silently; report it.
    m_netRespOn = False
    hr = View2_AddWebResourceResponseReceived()
    If hr = 0 Then
        m_netRespOn = True
    Else
        Wv2Log.LogW "Wv2Pane.NetCaptureStart: *could not set up the response event* hr=&H" & _
                    Hex(hr) & " the status column of the list stays blank."
    End If

    NetLogClear
    m_netOn = True
    NetCaptureStart = True

    Wv2Log.LogI "Wv2Pane.NetCaptureStart: capture started uri=" & uriFilter & _
                " kinds=" & spec & " (" & added & " filters)"
End Function


' ============================================================
' NetCaptureStop (N-1) - stop capturing
'
'   *Always remove the filters that were set* (in the spirit of design principle 106)
'   Removing just the handler also stops events from arriving, but if the net is left on the other side,
'   the next NetCaptureStart would unknowingly run with a wide net.
'
'   Return value: True if everything including cleanup succeeded.
' ============================================================
Public Function NetCaptureStop() As Boolean
    Dim v     As Variant
    Dim parts As Variant
    Dim hr    As Long
    Dim allOk As Boolean

    If Not m_netOn Then
        NetCaptureStop = True
        Exit Function
    End If

    allOk = True

    ' --- (1) Remove the filters ---
    '   *When already closed / there is no View, COM cannot be hit*
    '   The other side is gone, so no net remains either. Just discard what we remember.
    If m_closed Or m_pView = 0 Then
        Wv2Log.LogD "Wv2Pane.NetCaptureStop: there is no View, so skipping filter removal"
    Else
        For Each v In m_netFilters
            parts = Split(CStr(v), vbTab)
            hr = View_RemoveWebResourceRequestedFilter(CStr(parts(0)), CLng(parts(1)))
            If hr <> 0 Then
                Wv2Log.LogW "Wv2Pane.NetCaptureStop: filter removal failed " & _
                            CStr(parts(0)) & " / " & NetCtxName(CLng(parts(1))) & _
                            " hr=&H" & Hex(hr)
                allOk = False
            End If
        Next v
    End If
    Set m_netFilters = New Collection

    ' --- (2) Remove the handler ---
    hr = View_RemoveWebResourceRequested()
    If hr <> 0 Then allOk = False

    ' --- (3) Remove the response event (N-3) ---
    hr = View2_RemoveWebResourceResponseReceived()
    If hr <> 0 Then allOk = False
    m_netRespOn = False

    ' --- (4) N-4: release the held ResponseViews ---
    '   *It may end without the completion ever arriving* so do not leak.
    NetRespPendingClear

    m_netOn = False
    NetCaptureStop = allOk

    Wv2Log.LogI "Wv2Pane.NetCaptureStop: capture ended (total fired " & m_netTotal & " / " & _
                "in hand " & m_netCount & " / overflowed " & m_netDropped & ")"
End Function


' ============================================================
' NetCaptureOn (N-1) - whether capturing
' ============================================================
Public Property Get NetCaptureOn() As Boolean
    NetCaptureOn = m_netOn
End Property


' ============================================================
' NetLogCount / NetLogTotal / NetLogDropped (N-1)
'
'   NetLogCount   : number of entries currently in the ring (the number that can be taken out)
'   NetLogTotal   : *total fired count* since capture started (overflowed ones are counted too)
'   NetLogDropped : *number dropped by overflowing from the ring*
'
'   *Both are "cumulative since capture started" and do not return to 0 on NetLogDrain* (N-1b)
'   Only the contents were taken out. To return to 0, call NetLogClear.
'
'   *If NetLogCount and NetLogTotal disagree, some were missed*
'   A pair so as not to cut counts silently (same judgement as D-6's LastAllTruncated).
' ============================================================
Public Property Get NetLogCount() As Long
    NetLogCount = m_netCount
End Property

Public Property Get NetLogTotal() As Long
    NetLogTotal = m_netTotal
End Property

Public Property Get NetLogDropped() As Long
    NetLogDropped = m_netDropped
End Property


' ============================================================
' NetLogCapacity (N-1) - capacity of the ring (entries)
'   *Changing it discards the contents* (because it is rebuilt). Decide before starting capture.
' ============================================================
Public Property Get NetLogCapacity() As Long
    NetLogCapacity = m_netCap
End Property

Public Property Let NetLogCapacity(ByVal n As Long)
    If n < 1 Then n = 1
    If n > NET_CAP_MAX Then
        Wv2Log.LogW "Wv2Pane.NetLogCapacity: " & n & " exceeds the upper limit " & NET_CAP_MAX & ", so rounding"
        n = NET_CAP_MAX
    End If
    m_netCap = n
    NetRingAlloc
    Wv2Log.LogD "Wv2Pane.NetLogCapacity = " & m_netCap & " (contents were rebuilt, so empty)"
End Property


' ============================================================
' NetLogClear (N-1) - empty the ring
'   *The baseline of elapsed ms is reset here too* (the next entry starts from +0ms)
'   *The cumulative counters (NetLogTotal / NetLogDropped) also return to 0*
'   If you only want to take out the contents, use NetLogDrain (the cumulative counts are kept).
' ============================================================
Public Sub NetLogClear()
    NetRingAlloc
End Sub


' ============================================================
' NetLogLine (N-1) - read only one entry without flushing
'
'   idx is *1-based, oldest first*. An empty string if out of range.
'   One line is 7 items separated by vbTab (5 -> 7 in N-3):
'     serial number / elapsed ms / method / kind name / URL / status / reason phrase
'   *Status was added at the end* so as not to move the existing indexes (0 to 4).
'   *Blank if no response has arrived yet* (issue 7, option alpha)
'
'   *Draining removes them, so when using it for judgement, read with this*
'   (this will also be the entry point when emitting PowerShell code in N-5)
' ============================================================
Public Property Get NetLogLine(ByVal idx As Long) As String
    Dim k As Long
    k = NetRingIndex(idx)
    If k < 0 Then Exit Property

    NetLogLine = m_netSeq(k) & vbTab & m_netMs(k) & vbTab & _
                 m_netMethod(k) & vbTab & NetCtxName(m_netCtx(k)) & vbTab & m_netUri(k) & _
                 vbTab & NetStatusText(m_netStatus(k)) & vbTab & m_netReason(k)
End Property


' ============================================================
' NetLogDrain (N-1) - flush what has accumulated to Wv2Log and empty it
'
'   *After taking out, the ring becomes empty* (same contract as D-4's SpaProbeDrainLog).
'   *Note for N-3* Responses are written to the list rows, so if you drain a request whose
'   response has not arrived yet, that response loses its place and is counted in NetRespUnmatched.
'   *Read the verdict from the log file* (design principle 105). The Immediate window is flooded by plumbing logs.
'
'   Return value: the number flushed. 0 means nothing had accumulated.
' ============================================================
Public Function NetLogDrain() As Long
    Dim i As Long
    Dim n As Long

    n = m_netCount
    If n = 0 Then
        Wv2Log.LogI "  --- network requests: none accumulated (total fired " & m_netTotal & ") ---"
        Exit Function
    End If

    Wv2Log.LogI "  --- network requests " & n & " (total fired " & m_netTotal & ") ---"
    Wv2Log.LogI "        #   elapsed  method   kind       status  URL"
    For i = 1 To n
        Wv2Log.LogI "      " & NetLogFormat(i)
    Next i

    If m_netDropped > 0 Then
        ' *Counted cumulatively* so if you drain repeatedly while overflowed, it appears every time.
        '   It is not that it keeps appearing and is annoying; *the fact that we are still missing entries*
        '   does not go away. If you want it gone, recount with NetLogClear.
        Wv2Log.LogW "  *" & m_netDropped & " entries have overflowed and been discarded since capture started* " & _
                    "Increase NetLogCapacity or run NetLogDrain more often."
    End If

    ' --- N-3: responses that had no counterpart in the list ---
    '   *A large number is not abnormal* This event has no filter, so
    '   responses for images and CSS all come here too. The indicator of misses is NetLogDropped.
    If m_netUnmatched > 0 Then
        Wv2Log.LogI "  (" & m_netUnmatched & " responses had no counterpart in the list - " & _
                    "mostly things outside the capture target such as images and CSS, and this is normal)"
    End If

    NetLogDrain = n
    ' *Empty only the contents (the cumulative counts are kept)* issue 3 option Y
    NetRingEmpty
End Function


' ============================================================
' View_AddWebResourceRequested (N-1, ICoreWebView2 vtable 55)
'
'   *Index 55 is the value settled from WebView2.h with tools/vtable.py* (design principle 115).
'   Working backwards from the existing wiring it had come out as 54 (OpenDevToolsWindow was missed in the count).
'
'   The implementation uses View2_AddDOMContentLoaded as a verbatim model. The only differences are
'   "m_pView instead of m_pView2" and "55/56 instead of vtable 64/65".
'   pVTableObj = h.Slot + VTABLE_OBJ_OFFSET is always gone through in the same way (design principle 45).
'
'   Return value: HRESULT (S_OK = 0 means success)
' ============================================================
Public Function View_AddWebResourceRequested() As Long
    If m_closed Then
        View_AddWebResourceRequested = &H80004003   ' E_POINTER
        Exit Function
    End If
    If m_pView = 0 Then
        Debug.Print "Wv2Pane.View_AddWebResourceRequested: m_pView is not acquired"
        View_AddWebResourceRequested = &H80004003
        Exit Function
    End If

    ' If already registered, do nothing
    If Not (m_persistentHandlers(HK_WebResourceRequested) Is Nothing) Then
        Debug.Print "Wv2Pane.View_AddWebResourceRequested: already registered, skipping"
        View_AddWebResourceRequested = 0
        Exit Function
    End If

    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_WebResourceRequested, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_AddWebResourceRequested: failed to Acquire the handler"
        View_AddWebResourceRequested = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If

    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As LongLong
    Dim hr As Long
    hr = dcf(m_pView, 55, "View.add_WebResourceRequested", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_AddWebResourceRequested: failed hr=&H" & Hex(hr)
        Wv2Thunks.Thunks_DiscardHandler h   ' English fork: return the slot (see Thunks_DiscardHandler)
        Set h = Nothing
        View_AddWebResourceRequested = hr
        Exit Function
    End If

    m_persistentTokens(HK_WebResourceRequested) = token
    Set m_persistentHandlers(HK_WebResourceRequested) = h
    Debug.Print "Wv2Pane.View_AddWebResourceRequested OK: token=" & token & _
                ", handler.slot=" & h.Slot

    View_AddWebResourceRequested = 0
End Function


' ============================================================
' View_RemoveWebResourceRequested (N-1, ICoreWebView2 vtable 56)
'   If not registered, a no-op (returns S_OK).
' ============================================================
Public Function View_RemoveWebResourceRequested() As Long
    If m_persistentHandlers(HK_WebResourceRequested) Is Nothing Then
        View_RemoveWebResourceRequested = 0   ' not registered is a no-op
        Exit Function
    End If

    Dim hr As Long
    If m_pView <> 0 Then
        hr = dcf(m_pView, 56, "View.remove_WebResourceRequested", _
                 m_persistentTokens(HK_WebResourceRequested))
        If hr <> 0 Then
            Debug.Print "Wv2Pane.View_RemoveWebResourceRequested: unregister failed hr=&H" & Hex(hr)
        Else
            Debug.Print "Wv2Pane.View_RemoveWebResourceRequested OK"
        End If
    Else
        ' There is no m_pView, so there is no means to fire remove. Just cut the reference.
        Debug.Print "Wv2Pane.View_RemoveWebResourceRequested: m_pView is 0, so releasing the reference only"
    End If

    Set m_persistentHandlers(HK_WebResourceRequested) = Nothing
    m_persistentTokens(HK_WebResourceRequested) = 0^

    View_RemoveWebResourceRequested = hr
End Function


' ============================================================
' View_AddWebResourceRequestedFilter (N-1, ICoreWebView2 vtable 57)
'
'   *If no filter is registered, not a single event arrives*
'   Call once per kind. NetCaptureStart fires it for the 3 default kinds.
'
'   uriFilter       : URL pattern. * and ? can be used ("*" for all)
'   resourceContext : value of COREWEBVIEW2_WEB_RESOURCE_CONTEXT (0=ALL, 1=DOCUMENT, ...)
'
'   Return value: HRESULT (S_OK = 0 means success)
' ============================================================
Public Function View_AddWebResourceRequestedFilter( _
    ByVal uriFilter As String, _
    ByVal resourceContext As Long) As Long

    If m_pView = 0 Then
        View_AddWebResourceRequestedFilter = &H80004003
        Exit Function
    End If
    View_AddWebResourceRequestedFilter = dcf(m_pView, 57, _
                                             "View.AddWebResourceRequestedFilter", _
                                             StrPtr(uriFilter), resourceContext)
End Function


' ============================================================
' View_RemoveWebResourceRequestedFilter (N-1, ICoreWebView2 vtable 58)
'   *It is not removed unless called with the same (uri, context) pair as when it was registered*
' ============================================================
Public Function View_RemoveWebResourceRequestedFilter( _
    ByVal uriFilter As String, _
    ByVal resourceContext As Long) As Long

    If m_pView = 0 Then
        View_RemoveWebResourceRequestedFilter = &H80004003
        Exit Function
    End If
    View_RemoveWebResourceRequestedFilter = dcf(m_pView, 58, _
                                                "View.RemoveWebResourceRequestedFilter", _
                                                StrPtr(uriFilter), resourceContext)
End Function


' ============================================================
' View_OnWebResourceRequested (N-1)
'   Called from ComCallbackHandler via dynamic binding. Public Sub is required.
'
'   *Unlike the other receivers, the error is not re-raised* (post-hoc report)
'     WebResourceRequested is called synchronously in the middle of traffic, so if Err.Raise
'     is done here, *an error dialog appears in the middle of the traffic*. Moreover, a single page produces
'     tens to hundreds of them, so a single broken one makes it impossible to continue operating.
'     So instead of swallowing it, *one "(err)" entry is left in the ring*.
'     The fact that it was swallowed always shows up in the output of NetLogDrain.
' ============================================================
Public Sub View_OnWebResourceRequested(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Dim errNum As Long
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View_OnWebResourceRequested_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        On Error Resume Next
        NetPush "(err)", -1, "*error in the receiver " & errNum & " " & errDsc & "*"
    End If
End Sub


' ============================================================
' View_OnWebResourceRequested_Core (N-1, Private)
'
'   *Record only, then return immediately (issue 5)*
'     Fire neither Wv2Log nor Debug.Print here. EvalSync is out of the question (spec facts 14 / 20).
'     The funcName passed to dcf / GetStringProperty is also left empty, so that even on failure
'     nothing is shown in the Immediate window (failures are written to the ring, so they are not lost).
'
'   vtable of args (settled with tools/vtable.py):
'     3 = get_Request(ICoreWebView2WebResourceRequest**)
'     7 = get_ResourceContext(COREWEBVIEW2_WEB_RESOURCE_CONTEXT*)
'   vtable of request:
'     3 = get_Uri(LPWSTR*)   5 = get_Method(LPWSTR*)
'
'   *get_Request returns AddRef'd* so always Release it (COM convention).
'   This path runs hundreds of times per page, so a leak here visibly leaks.
' ============================================================
Private Sub View_OnWebResourceRequested_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Dim pReq As LongPtr
    Dim hr   As Long
    Dim uri  As String
    Dim mth  As String
    Dim ctx  As Long

    If pArgs = 0 Then
        NetPush "(err)", -1, "*args is 0*"
        Exit Sub
    End If

    ' Kind (if it cannot be obtained, 0 = ALL is returned. get_ResourceContext does not originally
    ' return ALL, so if ALL lines up in the log it is a sign that this retrieval failed)
    ctx = GetLongProperty(pArgs, 7, "")

    hr = dcf(pArgs, 3, "", VarPtr(pReq))
    If hr <> 0 Or pReq = 0 Then
        NetPush "(err)", ctx, "*get_Request failed hr=&H" & Hex(hr) & "*"
        Exit Sub
    End If

    uri = GetStringProperty(pReq, 3, "")
    mth = GetStringProperty(pReq, 5, "")

    NetPush mth, ctx, uri

    ' --- N-2: detail (headers and body) ---
    '   *Default OFF* so this runs only when NetDetailOn = True.
    '   Read before releasing pReq (both args and pReq are valid only during Invoke).
    If m_detOn Then NetReadDetail pReq, mth, uri

    ComRelease pReq
End Sub


' ============================================================
' NetReadDetail (N-2, Private) - read the detail of one request and put it on the ring
'
'   *Runs inside the handler* (issue 1, option A). args is valid only during Invoke, so
'   there is no choice but to read synchronously here. Hence it is always cut at the limit.
'
'   *Bodies are read only for non-GET (issue 7, option alpha)*
'     A GET almost never has a body. Excluding just this by default removes most of the waste.
'     To narrow further by URL, use NetDetailUriFilter.
' ============================================================
Private Sub NetReadDetail(ByVal pReq As LongPtr, _
                          ByVal mth As String, _
                          ByVal uri As String)
    Dim headers  As String
    Dim cType    As String
    Dim nBytes   As Long
    Dim trunc    As Boolean
    Dim isText   As Boolean
    Dim bodyText As String
    Dim hexHead  As String
    Dim buf()    As Byte

    ' Narrow by URL (do not narrow if empty)
    If LenB(m_detUriFilter) > 0 Then
        If InStr(1, uri, m_detUriFilter, vbTextCompare) = 0 Then Exit Sub
    End If

    headers = NetReadHeaders(pReq)
    cType = NetHeaderValue(headers, "content-type")

    ' *Do not read the body of a GET*
    If UCase$(mth) <> "GET" Then
        If NetReadBody(pReq, nBytes, trunc, buf) Then
            If nBytes > 0 Then
                isText = NetIsTextType(cType)
                If isText Then
                    bodyText = Utf8BytesToString(buf, nBytes)
                Else
                    hexHead = NetHexHead(buf, nBytes)
                End If
            End If
        End If
    End If

    NetDetailPush m_netTotal, mth, uri, headers, cType, _
                  nBytes, trunc, isText, bodyText, hexHead
End Sub


' ============================================================
' NetRingAlloc (N-1, Private) - rebuild the ring and empty it
'   Passed through when the capacity is changed and when capture is restarted.
' ============================================================
Private Sub NetRingAlloc()
    If m_netCap < 1 Then m_netCap = NET_CAP_DEFAULT

    ReDim m_netSeq(0 To m_netCap - 1)
    ReDim m_netMs(0 To m_netCap - 1)
    ReDim m_netCtx(0 To m_netCap - 1)
    ReDim m_netMethod(0 To m_netCap - 1)
    ReDim m_netUri(0 To m_netCap - 1)
    ReDim m_netStatus(0 To m_netCap - 1)
    ReDim m_netReason(0 To m_netCap - 1)

    m_netHead = 0
    m_netCount = 0
    m_netTotal = 0
    m_netDropped = 0
    m_netUnmatched = 0
    m_netT0 = GetTickCount()
End Sub


' ============================================================
' NetRingEmpty (N-1b, Private) - empty *only the contents*
'   The cumulative counters (m_netTotal / m_netDropped) and the baseline of elapsed ms are not touched.
'   Used by NetLogDrain to "discard what was taken out".
' ============================================================
Private Sub NetRingEmpty()
    m_netHead = 0
    m_netCount = 0
End Sub


' ============================================================
' NetPush (N-1, Private) - put one entry on the ring
'   *If it overflows, overwrite the oldest and count the number discarded* (not cut silently)
' ============================================================
Private Sub NetPush(ByVal mth As String, ByVal ctx As Long, ByVal uri As String)
    m_netTotal = m_netTotal + 1

    m_netSeq(m_netHead) = m_netTotal
    m_netMs(m_netHead) = NetElapsedMs()
    m_netCtx(m_netHead) = ctx
    m_netMethod(m_netHead) = mth
    m_netUri(m_netHead) = uri
    ' *This slot is reused, so always clear it* Leaving the previous request's response would be a lie (N-3)
    m_netStatus(m_netHead) = 0
    m_netReason(m_netHead) = ""

    m_netHead = m_netHead + 1
    If m_netHead >= m_netCap Then m_netHead = 0

    If m_netCount < m_netCap Then
        m_netCount = m_netCount + 1
    Else
        m_netDropped = m_netDropped + 1
    End If
End Sub


' ============================================================
' NetElapsedMs (N-1, Private) - elapsed ms since capture started
'   GetTickCount is a Long, so it wraps around in about 49.7 days. The difference is raised to Double
'   and computed, and if it goes negative, 2^32 is added to restore it.
' ============================================================
Private Function NetElapsedMs() As Long
    Dim d As Double
    d = CDbl(GetTickCount()) - CDbl(m_netT0)
    If d < 0 Then d = d + 4294967296#
    If d > 2147483647# Then d = 2147483647#
    NetElapsedMs = CLng(d)
End Function


' ============================================================
' NetRingIndex (N-1, Private) - oldest-first 1-based -> position on the ring
'   -1 if out of range.
' ============================================================
Private Function NetRingIndex(ByVal idx As Long) As Long
    Dim k As Long

    If idx < 1 Or idx > m_netCount Then
        NetRingIndex = -1
        Exit Function
    End If

    k = m_netHead - m_netCount + (idx - 1)
    Do While k < 0
        k = k + m_netCap
    Loop
    NetRingIndex = k Mod m_netCap
End Function


' ============================================================
' NetLogFormat (N-1, Private) - make one entry a single column-aligned line
' ============================================================
Private Function NetLogFormat(ByVal idx As Long) As String
    Dim k As Long
    k = NetRingIndex(idx)
    If k < 0 Then Exit Function

    NetLogFormat = Right$("     " & m_netSeq(k), 5) & "  " & _
                   Right$("       " & m_netMs(k), 7) & "  " & _
                   Left$(m_netMethod(k) & "        ", 8) & " " & _
                   Left$(NetCtxName(m_netCtx(k)) & "          ", 10) & " " & _
                   Right$("     " & NetStatusText(m_netStatus(k)), 5) & " " & _
                   m_netUri(k)
End Function


' ============================================================
' NetStatusText (N-3, Private) - how to present the status
'   *0 means "no response has arrived yet", so make it blank* (issue 7, option alpha)
'   Writing 0 is indistinguishable from a real status code, and later stages would believe it (design principle 111).
' ============================================================
Private Function NetStatusText(ByVal st As Long) As String
    If st = 0 Then Exit Function
    If st < 0 Then NetStatusText = "(unreadable)": Exit Function
    NetStatusText = CStr(st)
End Function


' ============================================================
' NetCtxName (N-1, Private) - the name of COREWEBVIEW2_WEB_RESOURCE_CONTEXT
'   *Values settled with tools/vtable.py --enum COREWEBVIEW2_WEB_RESOURCE_CONTEXT*
' ============================================================
Private Function NetCtxName(ByVal ctx As Long) As String
    Select Case ctx
        Case 0:  NetCtxName = "ALL"
        Case 1:  NetCtxName = "DOCUMENT"
        Case 2:  NetCtxName = "STYLESHEET"
        Case 3:  NetCtxName = "IMAGE"
        Case 4:  NetCtxName = "MEDIA"
        Case 5:  NetCtxName = "FONT"
        Case 6:  NetCtxName = "SCRIPT"
        Case 7:  NetCtxName = "XHR"
        Case 8:  NetCtxName = "FETCH"
        Case 9:  NetCtxName = "TEXTTRACK"
        Case 10: NetCtxName = "EVENTSRC"
        Case 11: NetCtxName = "WEBSOCKET"
        Case 12: NetCtxName = "MANIFEST"
        Case 13: NetCtxName = "SIGNEDEX"
        Case 14: NetCtxName = "PING"
        Case 15: NetCtxName = "CSPREPORT"
        Case 16: NetCtxName = "OTHER"
        Case Else: NetCtxName = "CTX" & ctx
    End Select
End Function


' ============================================================
' NetCtxFromName (N-1, Private) - name (or number) -> kind value
'   -1 if unknown. *Numbers are also accepted* so that no fix is needed even if the SDK adds kinds.
' ============================================================
Private Function NetCtxFromName(ByVal ctxToken As String) As Long
    Dim t As String

    t = UCase$(Trim$(ctxToken))
    t = Replace$(t, "_", "")
    t = Replace$(t, " ", "")

    If LenB(t) = 0 Then
        NetCtxFromName = -1
        Exit Function
    End If

    If IsNumeric(t) Then
        NetCtxFromName = CLng(t)
        Exit Function
    End If

    Select Case t
        Case "ALL":                             NetCtxFromName = 0
        Case "DOCUMENT", "DOC":                 NetCtxFromName = 1
        Case "STYLESHEET", "CSS":               NetCtxFromName = 2
        Case "IMAGE", "IMG":                    NetCtxFromName = 3
        Case "MEDIA":                           NetCtxFromName = 4
        Case "FONT":                            NetCtxFromName = 5
        Case "SCRIPT", "JS":                    NetCtxFromName = 6
        Case "XHR", "XMLHTTPREQUEST":           NetCtxFromName = 7
        Case "FETCH":                           NetCtxFromName = 8
        Case "TEXTTRACK":                       NetCtxFromName = 9
        Case "EVENTSOURCE", "EVENTSRC":         NetCtxFromName = 10
        Case "WEBSOCKET", "WS":                 NetCtxFromName = 11
        Case "MANIFEST":                        NetCtxFromName = 12
        Case "SIGNEDEXCHANGE", "SIGNEDEX":      NetCtxFromName = 13
        Case "PING":                            NetCtxFromName = 14
        Case "CSPVIOLATIONREPORT", "CSPREPORT": NetCtxFromName = 15
        Case "OTHER":                           NetCtxFromName = 16
        Case Else:                              NetCtxFromName = -1
    End Select
End Function


' ============================================================
' *N-2 : request headers and body*
'
'   On top of the N-1 list (URL / method / kind), add *what is needed to reproduce a single request*.
'
'   # *The detail is entirely OFF by default (issue 8, option kappa)*
'     The N-1 list is complete as a "light tool that can always be used". Adding
'     header iteration (about 60 COM calls per request) and body reading
'     by default would spoil that quality. It works only when `NetDetailOn = True`.
'
'   # *Reading can break the traffic (issue 2)*
'     The `IStream` returned by `get_Content` has a position. Reading advances the position,
'     so when WebView2 later sends from the same stream, *an empty body goes out*.
'     After reading, always rewind to the start with `Seek(0, STREAM_SEEK_SET)`.
'     *Do not settle this with "it's probably fine"* -- in the verification, we compare against the body that the
'     POST target returns and count that the body arrived unbroken (design principle 112).
'
'   # *Secrets are hidden by default (issue 3, option M)*
'     `Cookie` / `Authorization` and the like are not left in plain text in the log file.
'     *They are held raw on the ring and hidden when output*, so if necessary
'     they can be seen with `NetRedact = False`. The default is the safe side.
'
'   # Added Public members
'     NetDetailOn        As Boolean  *default False* whether to take the detail
'     NetDetailCapacity  As Long     capacity of the detail ring (default 50)
'     NetBodyMaxBytes    As Long     upper limit of the body (default 65536)
'     NetRedact          As Boolean  *default True* whether to hide secrets
'     AddRedactHeader(name) / ClearRedactHeaders()
'     NetDetailUriFilter As String   substring match of URLs to take detail for (all if empty)
'     NetDetailCount     As Long
'     NetDetailDropped   As Long     *number dropped on overflow* (not cut silently)
'     NetDetailLine(i)   As String   seq/method/URL/bytes/truncated/kind
'     NetDetailHeaders(i) As String  headers (returned with redaction applied)
'     NetDetailBody(i)   As String   body text (contents only when it is text)
'     NetDetail(i)       As Boolean  output one entry to Wv2Log
'     NetDetailAll()     As Long     output all
'     NetDetailClear()
'
'   # vtable index (*settled with tools/vtable.py*, design principle 115)
'       ICoreWebView2WebResourceRequest::get_Content  ->IStream**   7
'       ICoreWebView2WebResourceRequest::get_Headers                9
'       ICoreWebView2HttpRequestHeaders::GetIterator                8
'       ...HttpHeadersCollectionIterator::GetCurrentHeader          3
'       ...::get_HasCurrentHeader / ::MoveNext                    4 / 5
'       IStream::Read / Seek        (from Windows SDK objidl.h)   3 / 5
' ============================================================


' ============================================================
' NetDetailOn (N-2) - whether to take the detail (headers and body)
'   *Default False* With the list alone it stays as light as N-1.
' ============================================================
Public Property Get NetDetailOn() As Boolean
    NetDetailOn = m_detOn
End Property

Public Property Let NetDetailOn(ByVal onOff As Boolean)
    m_detOn = onOff
    Wv2Log.LogD "Wv2Pane.NetDetailOn = " & onOff & _
                IIf(onOff, " (headers and body are read too)", " (list only)")
End Property


' ============================================================
' NetDetailCapacity (N-2) - capacity of the detail ring (entries)
'   *A ring separate from the list (issue 5, option S)* The list is for "order and the overall picture", the detail is
'   for "the one request I want to reproduce"; the uses differ. The list is not cut down for the detail.
'   *Changing it discards the contents*
' ============================================================
Public Property Get NetDetailCapacity() As Long
    NetDetailCapacity = m_detCap
End Property

Public Property Let NetDetailCapacity(ByVal n As Long)
    If n < 1 Then n = 1
    If n > NET_DET_CAP_MAX Then
        Wv2Log.LogW "Wv2Pane.NetDetailCapacity: " & n & " exceeds the upper limit " & _
                    NET_DET_CAP_MAX & ", so rounding"
        n = NET_DET_CAP_MAX
    End If
    m_detCap = n
    NetDetailRingAlloc
    Wv2Log.LogD "Wv2Pane.NetDetailCapacity = " & m_detCap & " (contents were rebuilt, so empty)"
End Property


' ============================================================
' NetBodyMaxBytes (N-2) - upper limit for reading the body (bytes)
'   *It is read synchronously inside the handler* so leaving it unbounded would jam the traffic.
'   That it was cut at the limit appears in the "truncated" column of NetDetailLine (not cut silently).
' ============================================================
Public Property Get NetBodyMaxBytes() As Long
    NetBodyMaxBytes = m_detBodyMax
End Property

Public Property Let NetBodyMaxBytes(ByVal n As Long)
    If n < 0 Then n = 0
    If n > NET_BODY_MAX_LIMIT Then n = NET_BODY_MAX_LIMIT
    m_detBodyMax = n
    Wv2Log.LogD "Wv2Pane.NetBodyMaxBytes = " & m_detBodyMax
End Property


' ============================================================
' NetRedact (N-2) - whether to hide secret headers
'   *Default True* Cookie / Authorization and the like are not left in plain text in the log.
'   *They are held raw on the ring* so setting False lets you see them on the spot.
' ============================================================
Public Property Get NetRedact() As Boolean
    NetRedact = m_detRedact
End Property

Public Property Let NetRedact(ByVal onOff As Boolean)
    m_detRedact = onOff
    If Not onOff Then
        Wv2Log.LogW "Wv2Pane.NetRedact = False *Cookie etc. will be left in plain text in the log* " & _
                    "Always check before handing the log to anyone."
    Else
        Wv2Log.LogD "Wv2Pane.NetRedact = True (secret headers are hidden)"
    End If
End Property


' ============================================================
' AddRedactHeader / ClearRedactHeaders (N-2)
'   Add to the targets to hide / empty including the default list. Names are case-insensitive.
'   For the default list see NetRedactDefaults.
' ============================================================
Public Sub AddRedactHeader(ByVal headerName As String)
    Dim t As String
    t = LCase$(Trim$(headerName))
    If LenB(t) = 0 Then Exit Sub
    If InStr(1, m_detRedactList, "|" & t & "|") > 0 Then Exit Sub
    m_detRedactList = m_detRedactList & t & "|"
    Wv2Log.LogD "Wv2Pane.AddRedactHeader: " & t
End Sub

Public Sub ClearRedactHeaders()
    m_detRedactList = "|"
    Wv2Log.LogW "Wv2Pane.ClearRedactHeaders: *there are no targets left to hide* " & _
                "Cookie will also be output in plain text."
End Sub


' ============================================================
' NetDetailUriFilter (N-2) - substring match of URLs to take detail for
'   All if empty. *Header iteration is not cheap* so if the target is decided,
'   narrowing here keeps it light.
' ============================================================
Public Property Get NetDetailUriFilter() As String
    NetDetailUriFilter = m_detUriFilter
End Property

Public Property Let NetDetailUriFilter(ByVal pattern As String)
    m_detUriFilter = pattern
    Wv2Log.LogD "Wv2Pane.NetDetailUriFilter = [" & pattern & "]" & _
                IIf(LenB(pattern) = 0, " (not narrowed)", "")
End Property


' ============================================================
' NetDetailCount (N-2) - number of details currently in the ring
' ============================================================
Public Property Get NetDetailCount() As Long
    NetDetailCount = m_detCount
End Property


' ============================================================
' NetDetailDropped (N-2) - number discarded by overflowing the detail ring
'   *Not cut silently* The same role as the list's NetLogDropped.
'   Returns to 0 with NetDetailClear / a change of NetDetailCapacity.
' ============================================================
Public Property Get NetDetailDropped() As Long
    NetDetailDropped = m_detDropped
End Property


' ============================================================
' NetDetailLine (N-2) - read the heading of one detail (oldest first, 1-based)
'   6 items separated by vbTab:
'     serial number / method / URL / body bytes / truncated or not / Content-Type
'   *The contents of headers and body are not included* (get them with NetDetailHeaders / NetDetailBody)
' ============================================================
Public Property Get NetDetailLine(ByVal idx As Long) As String
    Dim k As Long
    k = NetDetailRingIndex(idx)
    If k < 0 Then Exit Property

    NetDetailLine = m_detSeq(k) & vbTab & m_detMethod(k) & vbTab & m_detUri(k) & _
                    vbTab & m_detBytes(k) & vbTab & CStr(m_detTrunc(k)) & _
                    vbTab & m_detCType(k)
End Property


' ============================================================
' NetDetailHeaders (N-2) - read the headers of one detail
'   One header per line, separated by vbLf. *If NetRedact = True, secrets are hidden in the return*
' ============================================================
Public Property Get NetDetailHeaders(ByVal idx As Long) As String
    Dim k As Long
    k = NetDetailRingIndex(idx)
    If k < 0 Then Exit Property

    NetDetailHeaders = NetRedactApply(m_detHeaders(k))
End Property


' ============================================================
' NetDetailBody (N-2) - read the body of one detail
'   *Contents are filled in only when judged to be text* (issue 4, option P).
'   An empty string for binary. For the byte count, look at the 4th item of NetDetailLine.
' ============================================================
Public Property Get NetDetailBody(ByVal idx As Long) As String
    Dim k As Long
    k = NetDetailRingIndex(idx)
    If k < 0 Then Exit Property

    NetDetailBody = m_detBody(k)
End Property


' ============================================================
' NetDetail (N-2) - output one detail to Wv2Log
'   *Not mixed into the list (NetLogDrain) (issue 6, option V)* One request becomes tens of lines, so
'   mixing would make the list unreadable (in the spirit of design principle 105).
'   Return value: True if output.
' ============================================================
Public Function NetDetail(ByVal idx As Long) As Boolean
    Dim k As Long
    Dim i As Long
    Dim lines As Variant

    k = NetDetailRingIndex(idx)
    If k < 0 Then Exit Function

    Wv2Log.LogI "  --- detail #" & m_detSeq(k) & " ---"
    Wv2Log.LogI "      " & m_detMethod(k) & " " & m_detUri(k)

    If LenB(m_detHeaders(k)) > 0 Then
        lines = Split(NetRedactApply(m_detHeaders(k)), vbLf)
        Wv2Log.LogI "      [headers: " & (UBound(lines) - LBound(lines) + 1) & "]"
        For i = LBound(lines) To UBound(lines)
            Wv2Log.LogI "        " & lines(i)
        Next i
    Else
        Wv2Log.LogI "      [headers] could not be obtained"
    End If

    If m_detBytes(k) = 0 Then
        Wv2Log.LogI "      [body] none"
    ElseIf m_detIsText(k) Then
        Wv2Log.LogI "      [body] " & m_detBytes(k) & " bytes (text" & _
                    IIf(m_detTrunc(k), ", *cut at the limit*", "") & ")"
        lines = Split(Replace$(m_detBody(k), vbCrLf, vbLf), vbLf)
        For i = LBound(lines) To UBound(lines)
            Wv2Log.LogI "        " & lines(i)
        Next i
    Else
        Wv2Log.LogI "      [body] " & m_detBytes(k) & " bytes (binary" & _
                    IIf(m_detTrunc(k), ", *cut at the limit*", "") & _
                    ") head: " & m_detHex(k)
    End If

    ' --- N-3 / N-4: response ---
    If m_detStatus(k) = 0 Then
        Wv2Log.LogI "      [response] not arrived yet"
    Else
        Wv2Log.LogI "      [response] " & NetStatusText(m_detStatus(k)) & " " & m_detReason(k)
        If LenB(m_detRespHeaders(k)) > 0 Then
            lines = Split(NetRedactApply(m_detRespHeaders(k)), vbLf)
            Wv2Log.LogI "      [response headers: " & (UBound(lines) - LBound(lines) + 1) & "]"
            For i = LBound(lines) To UBound(lines)
                Wv2Log.LogI "        " & lines(i)
            Next i
        End If

        ' --- N-4: response body ---
        '   *Always output the encoding and the leading hex* (issue 2, option P')
        '   If it is compressed, the head of the hex shows it (1F 8B for gzip).
        If m_detRespBytes(k) > 0 Or LenB(m_detRespEnc(k)) > 0 Then
            Wv2Log.LogI "      [response body] " & m_detRespBytes(k) & " bytes" & _
                        IIf(m_detRespTrunc(k), " *cut at the limit*", "") & _
                        IIf(LenB(m_detRespEnc(k)) > 0, " encoding=" & m_detRespEnc(k), "") & _
                        IIf(m_detRespIsText(k), " (text)", " (binary)")
            Wv2Log.LogI "        head: " & m_detRespHex(k)
            If LenB(m_detRespBody(k)) > 0 Then
                lines = Split(Replace$(Left$(m_detRespBody(k), 2000), vbCrLf, vbLf), vbLf)
                For i = LBound(lines) To UBound(lines)
                    Wv2Log.LogI "        " & lines(i)
                Next i
            End If
        End If
    End If

    NetDetail = True
End Function


' ============================================================
' NetDetailAll (N-2) - output all accumulated details to Wv2Log
'   Output *only when called explicitly*. The return value is the number output.
'   *They do not disappear after output* (unlike NetLogDrain. Kept because we want to use them for judgement)
' ============================================================
Public Function NetDetailAll() As Long
    Dim i As Long

    If m_detCount = 0 Then
        Wv2Log.LogI "  --- details: none accumulated ---"
        Exit Function
    End If

    Wv2Log.LogI "  --- details " & m_detCount & " ---"
    If m_detDropped > 0 Then
        Wv2Log.LogW "  *" & m_detDropped & " entries overflowed from the detail ring and were discarded* " & _
                    "Increase NetDetailCapacity or narrow with NetDetailUriFilter."
    End If
    For i = 1 To m_detCount
        NetDetail i
    Next i
    NetDetailAll = m_detCount
End Function


' ============================================================
' NetDetailClear (N-2) - empty the detail ring
' ============================================================
Public Sub NetDetailClear()
    NetDetailRingAlloc
End Sub


' ============================================================
' NetReadHeaders (N-2, Private) - make the request headers into one string
'
'   Gets the iterator with ICoreWebView2HttpRequestHeaders::GetIterator (8)
'   and iterates with GetCurrentHeader (3) / get_HasCurrentHeader (4) / MoveNext (5).
'
'   *Runs inside the handler* so it is cut at the limit (NET_HDR_MAX).
'   *GetCurrentHeader returns two LPWSTRs* so CoTaskMemFree both.
'
'   Return value: "name: value" joined with vbLf. An empty string if none could be obtained.
' ============================================================
Private Function NetReadHeaders(ByVal pReq As LongPtr) As String
    Dim pHeaders As LongPtr
    Dim hr       As Long

    If pReq = 0 Then Exit Function

    hr = dcf(pReq, 9, "", VarPtr(pHeaders))          ' request.get_Headers
    If hr <> 0 Or pHeaders = 0 Then Exit Function

    ' *GetIterator on the request side is 8* (7 on the response side. The order differs -- design principle 115)
    NetReadHeaders = NetReadHeadersIter(pHeaders, 8)
    ComRelease pHeaders
End Function


' ============================================================
' NetReadHeadersIter (N-3, Private) - make a collection of headers into one string
'
'   The common content for the request side and the response side. *The only difference is the GetIterator index* so
'   that was made an argument (request side 8 / response side 7).
'   The iterator type is ICoreWebView2HttpHeadersCollectionIterator for both.
'     3 = GetCurrentHeader(LPWSTR* name, LPWSTR* value)
'     4 = get_HasCurrentHeader(BOOL*)   5 = MoveNext(BOOL*)
'
'   *Runs inside the handler* so it is cut at the limit (NET_HDR_MAX).
'   *GetCurrentHeader returns two LPWSTRs* so CoTaskMemFree both.
' ============================================================
Private Function NetReadHeadersIter(ByVal pHeaders As LongPtr, _
                                    ByVal iterIndex As Long) As String
    Dim pIter    As LongPtr
    Dim hr       As Long
    Dim n        As Long
    Dim pName    As LongPtr
    Dim pValue   As LongPtr
    Dim moreFlag As Long
    Dim sb       As String

    ' *Do not return empty; leave the reason (N-3b)*
    '   Since we are inside the handler, we cannot fire the log. If we do not know
    '   where it became empty, a real-machine round trip is wasted (design principle 111).
    If pHeaders = 0 Then
        NetReadHeadersIter = "(diagnosis: the header collection is 0)"
        Exit Function
    End If

    hr = dcf(pHeaders, iterIndex, "", VarPtr(pIter))   ' GetIterator
    If hr <> 0 Then
        NetReadHeadersIter = "(diagnosis: GetIterator failed hr=&H" & Hex(hr) & _
                             " index=" & iterIndex & ")"
        Exit Function
    End If
    If pIter = 0 Then
        NetReadHeadersIter = "(diagnosis: GetIterator is S_OK but iter is 0 index=" & _
                             iterIndex & ")"
        Exit Function
    End If

    Do While n < NET_HDR_MAX
        If Not GetBoolProperty(pIter, 4, "") Then Exit Do   ' get_HasCurrentHeader

        pName = 0
        pValue = 0
        hr = dcf(pIter, 3, "", VarPtr(pName), VarPtr(pValue))   ' GetCurrentHeader
        If hr = 0 Then
            If LenB(sb) > 0 Then sb = sb & vbLf
            sb = sb & PtrToString(pName) & ": " & PtrToString(pValue)
            n = n + 1
        End If
        If pName <> 0 Then CoTaskMemFree pName
        If pValue <> 0 Then CoTaskMemFree pValue
        If hr <> 0 Then Exit Do

        moreFlag = 0
        hr = dcf(pIter, 5, "", VarPtr(moreFlag))     ' MoveNext(BOOL*)
        If hr <> 0 Or moreFlag = 0 Then Exit Do
    Loop

    If n >= NET_HDR_MAX Then
        sb = sb & vbLf & "(*cut off at " & NET_HDR_MAX & "*)"
    End If
    If n = 0 Then
        ' *The iterator was obtained but there is not even one*
        sb = "(diagnosis: 0 headers. HasCurrentHeader is False from the start" & _
             " index=" & iterIndex & ")"
    End If

    ComRelease pIter
    NetReadHeadersIter = sb
End Function


' ============================================================
' NetReadBody (N-2, Private) - read the request body
'
'   Reads the IStream returned by ICoreWebView2WebResourceRequest::get_Content (7)
'   with IStream::Read (3), and *after reading rewinds to the start with Seek (5)*.
'
'   *If not rewound, the traffic breaks (issue 2)*
'     IStream has a position. If left as read, when WebView2 later sends from
'     the same stream, *an empty body goes out*.
'
'   Arguments (all outputs):
'     outBytes  : number of bytes read (the limit value if cut at the limit)
'     outTrunc  : whether cut at the limit
'     outBuf()  : the byte sequence read (0-based, outBytes entries valid)
'
'   Return value: True if there was a stream and it was read to the end (True even for 0 bytes).
'           False means "there is no body" (GET etc.) or the retrieval failed.
' ============================================================
Private Function NetReadBody(ByVal pReq As LongPtr, _
                             ByRef outBytes As Long, _
                             ByRef outTrunc As Boolean, _
                             ByRef outBuf() As Byte) As Boolean
    Dim pStream As LongPtr
    Dim hr      As Long
    Dim cap     As Long
    Dim got     As Long
    Dim chunk   As Long
    Dim cbRead  As Long

    outBytes = 0
    outTrunc = False

    If pReq = 0 Then Exit Function

    hr = dcf(pReq, 7, "", VarPtr(pStream))           ' get_Content
    If hr <> 0 Or pStream = 0 Then Exit Function     ' no body (GET etc.)

    cap = m_detBodyMax
    If cap < 1 Then
        ' Limit 0 = do not read. Even so the position has not been moved, so there is no need to rewind.
        ComRelease pStream
        NetReadBody = True
        Exit Function
    End If

    ' *The reading part was split out into NetReadStream (N-4)*
    '   The response body is read the same way, so it was shared. The only difference is
    '   *this side must rewind with Seek(0) after reading*.
    If Not NetReadStream(pStream, cap, got, outTrunc, outBuf) Then
        ComRelease pStream
        Exit Function
    End If

    ' ** Always rewind what was read ** (issue 2, option X)
    '   Seek(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER* plibNewPosition)
    '   dlibMove is an 8-byte by-value, dwOrigin is 4 bytes, and the last may be NULL.
    hr = dcf(pStream, 5, "", 0^, STREAM_SEEK_SET, 0^)
    If hr <> 0 Then
        ' *Could not rewind* This may have broken the traffic, so do not proceed silently.
        ' Since we are inside the handler we do not fire the log; leave a mark in the body contents.
        outTrunc = True
        outBytes = got
        ComRelease pStream
        NetReadBody = True
        Exit Function
    End If

    ComRelease pStream
    outBytes = got
    NetReadBody = True
End Function


' ============================================================
' NetReadStream (N-4, Private) - read an IStream up to the limit
'
'   *Returned with the position left moved* Whether to rewind is the caller's concern.
'     *request body (N-2) ... *the traffic breaks if not rewound* so the caller does Seek(0)
'     *response body (N-4) ... it is a copy of the response so no need to rewind
'
'   *When the limit is reached, read one extra byte to check "whether more remains"*
'   Without this, "exactly at the limit" and "was cut" cannot be distinguished.
'
'   Return value: True if read to the end (True even for 0 bytes).
' ============================================================
Private Function NetReadStream(ByVal pStream As LongPtr, _
                               ByVal cap As Long, _
                               ByRef outBytes As Long, _
                               ByRef outTrunc As Boolean, _
                               ByRef outBuf() As Byte) As Boolean
    Dim hr     As Long
    Dim got    As Long
    Dim chunk  As Long
    Dim cbRead As Long

    outBytes = 0
    outTrunc = False
    If pStream = 0 Or cap < 1 Then Exit Function

    ReDim outBuf(0 To cap - 1)

    Do While got < cap
        chunk = cap - got
        If chunk > NET_BODY_CHUNK Then chunk = NET_BODY_CHUNK
        cbRead = 0
        hr = dcf(pStream, 3, "", VarPtr(outBuf(got)), chunk, VarPtr(cbRead))  ' Read
        If hr <> 0 Then Exit Do
        If cbRead <= 0 Then Exit Do
        got = got + cbRead
    Loop

    ' Try a single byte to confirm *whether more remains* (so as not to cut silently)
    If got >= cap Then
        Dim probe(0 To 0) As Byte
        cbRead = 0
        hr = dcf(pStream, 3, "", VarPtr(probe(0)), 1&, VarPtr(cbRead))
        If hr = 0 And cbRead > 0 Then outTrunc = True
    End If

    outBytes = got
    NetReadStream = True
End Function


' ============================================================
' NetDetailPush (N-2, Private) - put one detail on the ring
'   *If it overflows, overwrite the oldest* (same convention as the list).
' ============================================================
Private Sub NetDetailPush(ByVal seqNo As Long, _
                          ByVal mth As String, _
                          ByVal uri As String, _
                          ByVal headers As String, _
                          ByVal cType As String, _
                          ByVal nBytes As Long, _
                          ByVal trunc As Boolean, _
                          ByVal isText As Boolean, _
                          ByVal bodyText As String, _
                          ByVal hexHead As String)
    m_detSeq(m_detHead) = seqNo
    m_detMethod(m_detHead) = mth
    m_detUri(m_detHead) = uri
    m_detHeaders(m_detHead) = headers
    m_detCType(m_detHead) = cType
    m_detBytes(m_detHead) = nBytes
    m_detTrunc(m_detHead) = trunc
    m_detIsText(m_detHead) = isText
    m_detBody(m_detHead) = bodyText
    m_detHex(m_detHead) = hexHead
    ' N-3: the response has not arrived yet (this slot is reused, so always clear it)
    m_detStatus(m_detHead) = 0
    m_detReason(m_detHead) = ""
    m_detRespHeaders(m_detHead) = ""
    ' N-4: the body has not arrived yet either (this slot is reused, so always clear it)
    m_detRespBody(m_detHead) = ""
    m_detRespBytes(m_detHead) = 0
    m_detRespTrunc(m_detHead) = False
    m_detRespIsText(m_detHead) = False
    m_detRespEnc(m_detHead) = ""
    m_detRespHex(m_detHead) = ""
    m_detRespGot(m_detHead) = False

    m_detHead = m_detHead + 1
    If m_detHead >= m_detCap Then m_detHead = 0
    If m_detCount < m_detCap Then
        m_detCount = m_detCount + 1
    Else
        m_detDropped = m_detDropped + 1
    End If
End Sub


' ============================================================
' NetDetailRingAlloc / NetDetailRingIndex (N-2, Private)
' ============================================================
Private Sub NetDetailRingAlloc()
    If m_detCap < 1 Then m_detCap = NET_DET_CAP_DEFAULT

    ReDim m_detSeq(0 To m_detCap - 1)
    ReDim m_detMethod(0 To m_detCap - 1)
    ReDim m_detUri(0 To m_detCap - 1)
    ReDim m_detHeaders(0 To m_detCap - 1)
    ReDim m_detCType(0 To m_detCap - 1)
    ReDim m_detBytes(0 To m_detCap - 1)
    ReDim m_detTrunc(0 To m_detCap - 1)
    ReDim m_detIsText(0 To m_detCap - 1)
    ReDim m_detBody(0 To m_detCap - 1)
    ReDim m_detHex(0 To m_detCap - 1)
    ReDim m_detRespHeaders(0 To m_detCap - 1)
    ReDim m_detStatus(0 To m_detCap - 1)
    ReDim m_detReason(0 To m_detCap - 1)
    ReDim m_detRespBody(0 To m_detCap - 1)
    ReDim m_detRespBytes(0 To m_detCap - 1)
    ReDim m_detRespTrunc(0 To m_detCap - 1)
    ReDim m_detRespIsText(0 To m_detCap - 1)
    ReDim m_detRespEnc(0 To m_detCap - 1)
    ReDim m_detRespHex(0 To m_detCap - 1)
    ReDim m_detRespGot(0 To m_detCap - 1)

    m_detHead = 0
    m_detCount = 0
    m_detDropped = 0
End Sub

Private Function NetDetailRingIndex(ByVal idx As Long) As Long
    Dim k As Long

    If idx < 1 Or idx > m_detCount Then
        NetDetailRingIndex = -1
        Exit Function
    End If

    k = m_detHead - m_detCount + (idx - 1)
    Do While k < 0
        k = k + m_detCap
    Loop
    NetDetailRingIndex = k Mod m_detCap
End Function


' ============================================================
' NetRedactApply (N-2, Private) - hide secret headers
'   *Hidden at output time* The ring holds them raw, so
'   setting NetRedact = False lets you see them on the spot (issue 3, option M).
' ============================================================
Private Function NetRedactApply(ByVal headers As String) As String
    Dim lines As Variant
    Dim i     As Long
    Dim p     As Long
    Dim nm    As String
    Dim out   As String

    If LenB(headers) = 0 Then Exit Function
    If Not m_detRedact Then
        NetRedactApply = headers
        Exit Function
    End If

    lines = Split(headers, vbLf)
    For i = LBound(lines) To UBound(lines)
        p = InStr(1, CStr(lines(i)), ":")
        If p > 1 Then
            nm = LCase$(Trim$(Left$(CStr(lines(i)), p - 1)))
            If InStr(1, m_detRedactList, "|" & nm & "|") > 0 Then
                lines(i) = Left$(CStr(lines(i)), p) & " <redacted " & _
                           (Len(CStr(lines(i))) - p - 1) & " chars>"
            End If
        End If
        If LenB(out) > 0 Then out = out & vbLf
        out = out & CStr(lines(i))
    Next i

    NetRedactApply = out
End Function


' ============================================================
' NetIsTextType (N-2, Private) - whether Content-Type is text-like
'   *False (treated as binary) when it cannot be determined* The safe side.
'   When there is no Content-Type, it is not regarded as the form-submission default
'   application/x-www-form-urlencoded but False
'   (do not leave behind a plausible-looking value = design principle 111).
' ============================================================
Private Function NetIsTextType(ByVal cType As String) As Boolean
    Dim t As String
    t = LCase$(cType)
    If LenB(t) = 0 Then Exit Function

    If InStr(1, t, "text/") > 0 Then NetIsTextType = True: Exit Function
    If InStr(1, t, "json") > 0 Then NetIsTextType = True: Exit Function
    If InStr(1, t, "xml") > 0 Then NetIsTextType = True: Exit Function
    If InStr(1, t, "x-www-form-urlencoded") > 0 Then NetIsTextType = True: Exit Function
    If InStr(1, t, "javascript") > 0 Then NetIsTextType = True: Exit Function
    If InStr(1, t, "charset=") > 0 Then NetIsTextType = True: Exit Function
End Function


' ============================================================
' NetHexHead (N-2, Private) - show the head of binary data in hex
' ============================================================
Private Function NetHexHead(ByRef b() As Byte, ByVal cb As Long) As String
    Dim i As Long
    Dim n As Long
    Dim s As String

    n = cb
    If n > NET_HEX_HEAD Then n = NET_HEX_HEAD
    For i = 0 To n - 1
        s = s & Right$("0" & Hex$(b(i)), 2) & " "
    Next i
    If cb > n Then s = s & "..."
    NetHexHead = RTrim$(s)
End Function


' ============================================================
' NetHeaderValue (N-2, Private) - pick one out of the headers read
'   Names are case-insensitive. An empty string if absent.
' ============================================================
Private Function NetHeaderValue(ByVal headers As String, ByVal headerName As String) As String
    Dim lines As Variant
    Dim i     As Long
    Dim p     As Long
    Dim nm    As String

    If LenB(headers) = 0 Then Exit Function

    lines = Split(headers, vbLf)
    For i = LBound(lines) To UBound(lines)
        p = InStr(1, CStr(lines(i)), ":")
        If p > 1 Then
            nm = LCase$(Trim$(Left$(CStr(lines(i)), p - 1)))
            If nm = LCase$(headerName) Then
                NetHeaderValue = Trim$(Mid$(CStr(lines(i)), p + 1))
                Exit Function
            End If
        End If
    Next i
End Function


' ============================================================
' NetRedactDefaults (N-2, Private) - headers hidden by default
'   *Anything not here is output in plain text* To add, use AddRedactHeader.
' ============================================================
Private Function NetRedactDefaults() As String
    NetRedactDefaults = "|cookie|set-cookie|authorization|proxy-authorization|" & _
                        "x-api-key|x-auth-token|x-csrf-token|api-key|"
End Function

' ============================================================
' *N-3 : response status and headers*
'
'   Add the *result* to the N-1 list and the *response headers* to the N-2 detail.
'   One row of the list becomes self-contained as "when, what, how it was hit, and what happened".
'
'   # *This event has no filter (the central problem of N-3)*
'     AddWebResourceResponseReceivedFilter does not exist in the vtable of
'     ICoreWebView2_2 (after 61/62 comes NavigateWithWebResourceRequest). That is,
'     *images, CSS and fonts all come*. Moreover args holds only Request and
'     Response, and *the kind (COREWEBVIEW2_WEB_RESOURCE_CONTEXT) is not included*
'     so it cannot be narrowed by kind either.
'
'     So *we record only those that match a request already in the list (issue 1, option A)*.
'     N-1's net already defines "what we want to see", so we ride on it.
'     If you want everything, widen the list side with NetCaptureStart "ALL" and
'     the responses follow automatically. *The narrowing is not kept in two places*
'
'   # Matching (issue 2, option P / issue 6, option Y)
'     args has no handle for association such as a NavigationId, so
'     *we attach by URI + method to "the *oldest* request that has no response yet"*.
'     *Concurrent requests to the same URL are in principle mixed up. Moreover, the mix-up
'     cannot be detected* -- we state this honestly. A redirect (301 -> 200)
'     fires the request itself twice, so it attaches one to one.
'
'   # *Blank when no response has arrived yet (issue 7, option alpha)*
'     Putting 0 is indistinguishable from a real status code, and later stages (whether a human or
'     N-5's code generation) would believe it (design principle 111).
'
'   # Added Public members
'     NetLogStatus(i)        As Long    status of that row (0 = not arrived yet)
'     NetDetailRespHeaders(i) As String response headers (returned with redaction applied)
'     NetRespOn              As Boolean whether response events are being picked up
'     NetRespUnmatched       As Long    *number of responses that had no counterpart in the list*
'
'   # vtable index (*settled with tools/vtable.py*, design principle 115)
'       ICoreWebView2_2::add_WebResourceResponseReceived            61
'       ICoreWebView2_2::remove_WebResourceResponseReceived         62
'       ...ResponseReceivedEventArgs::get_Request / get_Response   3 / 4
'       ICoreWebView2WebResourceResponseView::get_Headers            3
'       ...::get_StatusCode (int*) / get_ReasonPhrase              4 / 5
'       *ICoreWebView2HttpResponseHeaders::GetIterator = 7*
'
'   *Trap* GetIterator of the request-side ICoreWebView2HttpRequestHeaders is 8, but
'   GetIterator of the response-side ICoreWebView2HttpResponseHeaders is *7*.
'   The response side has no put_ family and AppendHeader comes first, so the order is shifted.
'   Reusing it with "they must be the same" shifts by one and calls GetHeaders (design principle 115).
' ============================================================


' ============================================================
' NetLogStatus (N-3) - status code of that row
'   *0 means "no response has arrived yet"* There is no real 0, so it can be used as a sentinel.
'   In the display (NetLogLine / NetLogFormat) 0 is made blank (issue 7, option alpha).
' ============================================================
Public Property Get NetLogStatus(ByVal idx As Long) As Long
    Dim k As Long
    k = NetRingIndex(idx)
    If k < 0 Then Exit Property
    NetLogStatus = m_netStatus(k)
End Property


' ============================================================
' NetRespOn (N-3) - whether response events are being picked up
'   *If False, the status column of the list stays blank forever*
'   That happens when ICoreWebView2_2 could not be obtained.
' ============================================================
Public Property Get NetRespOn() As Boolean
    NetRespOn = m_netRespOn
End Property


' ============================================================
' NetRespUnmatched (N-3) - number of responses that had no counterpart in the list
'
'   *Mostly "things outside the capture target", which is normal*
'     This event has no filter, so responses for images and CSS all come too.
'     Since the list is narrowed to DOCUMENT / XHR / FETCH, other responses
'     are always counted here. *A large number is not abnormal*
'
'   *The indicator of misses is NetLogDropped*
'     Responses for requests that overflowed from the list and vanished also get mixed in here, but the two cannot be distinguished.
'     When you suspect misses, look at NetLogDropped.
' ============================================================
Public Property Get NetRespUnmatched() As Long
    NetRespUnmatched = m_netUnmatched
End Property


' ============================================================
' NetDetailRespHeaders (N-3) - read the response headers of one detail
'   One header per line, separated by vbLf. *If NetRedact = True, secrets are hidden in the return*
'   (Set-Cookie is in the default list)
' ============================================================
Public Property Get NetDetailRespHeaders(ByVal idx As Long) As String
    Dim k As Long
    k = NetDetailRingIndex(idx)
    If k < 0 Then Exit Property

    NetDetailRespHeaders = NetRedactApply(m_detRespHeaders(k))
End Property


' ============================================================
' View2_AddWebResourceResponseReceived (N-3, ICoreWebView2_2 vtable 61)
'
'   *An event of the _2 IF, so call through m_pView2* (design principle 45).
'   The implementation uses View2_AddDOMContentLoaded as a verbatim model.
'
'   Return value: HRESULT (S_OK = 0 means success)
' ============================================================
Public Function View2_AddWebResourceResponseReceived() As Long
    If m_closed Then
        View2_AddWebResourceResponseReceived = &H80004003   ' E_POINTER
        Exit Function
    End If
    If Not EnsureView2() Then
        Debug.Print "Wv2Pane.View2_AddWebResourceResponseReceived: EnsureView2 failed"
        View2_AddWebResourceResponseReceived = &H80004003
        Exit Function
    End If

    If Not (m_persistentHandlers(HK_WebResourceResponseReceived) Is Nothing) Then
        Debug.Print "Wv2Pane.View2_AddWebResourceResponseReceived: already registered"
        View2_AddWebResourceResponseReceived = 0
        Exit Function
    End If

    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_WebResourceResponseReceived, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View2_AddWebResourceResponseReceived: Acquire failed"
        View2_AddWebResourceResponseReceived = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If

    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As LongLong
    Dim hr As Long
    hr = dcf(m_pView2, 61, "View2.add_WebResourceResponseReceived", _
             pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View2_AddWebResourceResponseReceived: failed hr=&H" & Hex(hr)
        Wv2Thunks.Thunks_DiscardHandler h   ' English fork: return the slot (see Thunks_DiscardHandler)
        Set h = Nothing
        View2_AddWebResourceResponseReceived = hr
        Exit Function
    End If

    m_persistentTokens(HK_WebResourceResponseReceived) = token
    Set m_persistentHandlers(HK_WebResourceResponseReceived) = h
    Debug.Print "Wv2Pane.View2_AddWebResourceResponseReceived OK: token=" & token

    View2_AddWebResourceResponseReceived = 0
End Function


' ============================================================
' View2_RemoveWebResourceResponseReceived (N-3, vtable 62, m_pView2)
'   If not registered, a no-op (returns S_OK).
' ============================================================
Public Function View2_RemoveWebResourceResponseReceived() As Long
    If m_persistentHandlers(HK_WebResourceResponseReceived) Is Nothing Then
        View2_RemoveWebResourceResponseReceived = 0
        Exit Function
    End If

    Dim hr As Long
    If m_pView2 <> 0 Then
        hr = dcf(m_pView2, 62, "View2.remove_WebResourceResponseReceived", _
                 m_persistentTokens(HK_WebResourceResponseReceived))
        If hr <> 0 Then
            Debug.Print "Wv2Pane.View2_RemoveWebResourceResponseReceived: unregister failed hr=&H" & Hex(hr)
        End If
    Else
        Debug.Print "Wv2Pane.View2_RemoveWebResourceResponseReceived: m_pView2 is 0, so releasing the reference only"
    End If

    Set m_persistentHandlers(HK_WebResourceResponseReceived) = Nothing
    m_persistentTokens(HK_WebResourceResponseReceived) = 0^

    View2_RemoveWebResourceResponseReceived = hr
End Function


' ============================================================
' View2_OnWebResourceResponseReceived (N-3)
'   Called from ComCallbackHandler via dynamic binding. Public Sub is required.
'   *Like N-1's receiver, the error is not re-raised* (do not show a dialog in the middle of traffic)
' ============================================================
Public Sub View2_OnWebResourceResponseReceived(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Dim errNum As Long
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View2_OnWebResourceResponseReceived_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        On Error Resume Next
        m_netUnmatched = m_netUnmatched + 1
    End If
End Sub


' ============================================================
' View2_OnWebResourceResponseReceived_Core (N-3, Private)
'
'   *Record only, then return immediately* (same as N-1 / N-2). Fire neither the log nor Debug.Print.
'   *Do not touch the traffic itself* The body (GetContent) is asynchronous, so that is N-4's matter.
'   What is read here is only the status, reason phrase and response headers.
'
'   vtable of args: 3 = get_Request, 4 = get_Response
'   vtable of ResponseView: 3 = get_Headers, 4 = get_StatusCode, 5 = get_ReasonPhrase
' ============================================================
Private Sub View2_OnWebResourceResponseReceived_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Dim pReq  As LongPtr
    Dim pResp As LongPtr
    Dim hr    As Long
    Dim uri   As String
    Dim mth   As String
    Dim st    As Long
    Dim reason As String
    Dim k     As Long

    If pArgs = 0 Then Exit Sub

    ' --- Read the Request to learn which request this is the response to ---
    hr = dcf(pArgs, 3, "", VarPtr(pReq))             ' get_Request
    If hr <> 0 Or pReq = 0 Then Exit Sub
    uri = GetStringProperty(pReq, 3, "")             ' request.get_Uri
    mth = GetStringProperty(pReq, 5, "")             ' request.get_Method
    ComRelease pReq

    ' --- Look in the list for "the oldest request that has no response yet" ---
    '   *If not found, count it and finish (issue 1 option A / issue 5 option W)*
    '   This event has no filter, so responses for images and CSS come here too.
    '   They are not in the list, so naturally they are not found = normal.
    k = NetMatchRequest(uri, mth)
    If k < 0 Then
        m_netUnmatched = m_netUnmatched + 1
        Exit Sub
    End If

    ' --- Read the Response ---
    hr = dcf(pArgs, 4, "", VarPtr(pResp))            ' get_Response
    If hr <> 0 Or pResp = 0 Then
        m_netStatus(k) = -1                          ' *the response came but could not be read*
        m_netReason(k) = "(failed to get the response)"
        Exit Sub
    End If

    st = GetLongProperty(pResp, 4, "")               ' get_StatusCode
    reason = GetStringProperty(pResp, 5, "")         ' get_ReasonPhrase

    m_netStatus(k) = st
    m_netReason(k) = reason

    ' --- Response headers only when the detail is ON (issue 4, option U) ---
    If m_detOn Then NetAttachRespHeaders m_netSeq(k), pResp, st, reason

    ' --- N-4: start fetching the body (*this is the only place to fire it*) ---
    '   ResponseView is tied to the lifetime of args, so GetContent is fired inside the handler.
    '   The call itself returns immediately, and the body reaches the completion handler later.
    If m_respBodyOn And m_detOn Then NetStartRespContent m_netSeq(k), pResp

    ComRelease pResp
End Sub


' ============================================================
' NetMatchRequest (N-3, Private) - find from the list the one to attach a response to
'
'   *Find "the *oldest* request by URI + method that has no response yet"*
'   (issue 2 option P / issue 6 option Y). Scan oldest first and return the first one found.
'
'   *At agreement time we said "newest", but a check before implementation showed "oldest" is correct*
'   In a scene where two requests to the same URL fly and then the responses come back (Req1 -> Req2 -> Resp -> Resp),
'   searching from the newest *swaps the statuses*. Searching from the oldest gives FIFO,
'   and in the one-round-trip-at-a-time scene the result is the same (so oldest is always equal or better).
'
'   *Honest about the weakness* Concurrent requests to the same URL are mixed up. Moreover,
'   the mix-up cannot be detected. args has no handle for association such as a
'   NavigationId, so it cannot in principle be pinned down any further.
'
'   Return value: position on the ring. -1 if not found.
' ============================================================
Private Function NetMatchRequest(ByVal uri As String, ByVal mth As String) As Long
    Dim i As Long
    Dim k As Long

    NetMatchRequest = -1
    If m_netCount = 0 Then Exit Function

    ' *Oldest first (i = 1 -> m_netCount)* = attach the response that arrived first to the request that went out first
    For i = 1 To m_netCount
        k = NetRingIndex(i)
        If k >= 0 Then
            If m_netStatus(k) = 0 Then                       ' no response attached yet
                If m_netUri(k) = uri Then
                    If LenB(mth) = 0 Or m_netMethod(k) = mth Then
                        NetMatchRequest = k
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function


' ============================================================
' NetAttachRespHeaders (N-3, Private) - add response headers to the detail entry
'
'   Find the entry with the same serial number in the detail ring and write to it.
'   *If not found, do nothing* (it is the response of a request whose detail was not taken, so normal).
' ============================================================
Private Sub NetAttachRespHeaders(ByVal seqNo As Long, _
                                 ByVal pResp As LongPtr, _
                                 ByVal st As Long, _
                                 ByVal reason As String)
    Dim i As Long
    Dim k As Long
    Dim pHeaders As LongPtr
    Dim hr As Long

    For i = m_detCount To 1 Step -1
        k = NetDetailRingIndex(i)
        If k >= 0 Then
            If m_detSeq(k) = seqNo Then
                m_detStatus(k) = st
                m_detReason(k) = reason

                ' ResponseView.get_Headers (3) -> ICoreWebView2HttpResponseHeaders
                hr = dcf(pResp, 3, "", VarPtr(pHeaders))
                If hr <> 0 Then
                    m_detRespHeaders(k) = "(diagnosis: get_Headers failed hr=&H" & Hex(hr) & ")"
                ElseIf pHeaders = 0 Then
                    m_detRespHeaders(k) = "(diagnosis: get_Headers was S_OK but 0 was returned)"
                Else
                    ' *GetIterator on the response side is 7* (8 on the request side. The order differs)
                    m_detRespHeaders(k) = NetReadHeadersIter(pHeaders, 7)
                    ComRelease pHeaders
                End If
                Exit Sub
            End If
        End If
    Next i
End Sub


' ============================================================
' *N-4 : response body*
'
'   N-1 to N-3 were "read synchronously inside the handler".
'   *N-4 is the first to be asynchronous* -- `GetContent` takes a handler and returns immediately, and
'   the body reaches `OnResponseContentCompleted` later.
'
'   # *Unlike N-2, there is no need to worry about Seek(0)*
'     This only reads a copy of the response and does not touch the traffic itself.
'     There is no occasion for design principle 118 (count by a separate path when observation touches the target).
'
'   # *The body has its own switch (issue 1, option A)*
'     Headers are a few KB, bodies a few hundred KB; the orders of magnitude differ. Putting them on the same switch
'     would make NetDetailOn hard to use casually. It is taken only with `NetRespBodyOn = True`.
'     *NetDetailOn must also be True or there is nowhere to put it* (the body goes into the detail entry).
'
'   # *Do not assume compression; observe it (issue 2, option P')*
'     Responses carry content-encoding: gzip / br / zstd. Whether what GetContent returns is
'     already decoded is not written in the SDK header. So we keep all three:
'     *record content-encoding
'     **always keep the hex of the first 32 bytes* (for gzip, 1F 8B shows it at a glance)
'     *if judged to be text, also try UTF-8 decoding
'     One real-machine run gives the answer (design principle 117).
'
'   # *ResponseView is held until completion (issue 5, option beta)*
'     Whether "ResponseView may be released after calling GetContent" is
'     not written in the header. If we guess wrong, it shows up as *access to a released object
'     = crash*, and spec fact 20 means the debugger cannot be used either.
'     So we AddRef it, hold it in m_respPending, and Release it on completion.
'
'   # Added Public members
'     NetRespBodyOn        As Boolean  *default False*
'     NetRespBodyMaxBytes  As Long     upper limit of the body (default 65536)
'     NetRespBodyOrphan    As Long     number of bodies that had no place to go
'     NetDetailRespBody(i) As String   body (contents only when it is text)
'     NetDetailRespLine(i) As String   bytes / truncated / kind / encoding / head hex
'
'   # vtable index (*settled with tools/vtable.py*)
'       ICoreWebView2WebResourceResponseView::GetContent(handler)   6
'       completion handler Invoke(HRESULT, IStream*)                       3
'       IStream::Read                                                3
' ============================================================


' ============================================================
' NetRespBodyOn (N-4) - whether to take the response body
'   *Default False* A body is an order of magnitude different from headers (a few hundred KB and up).
'   *NetDetailOn must also be True or there is nowhere to put it* so we report that in that case.
' ============================================================
Public Property Get NetRespBodyOn() As Boolean
    NetRespBodyOn = m_respBodyOn
End Property

Public Property Let NetRespBodyOn(ByVal onOff As Boolean)
    m_respBodyOn = onOff
    If onOff And Not m_detOn Then
        Wv2Log.LogW "Wv2Pane.NetRespBodyOn = True but *NetDetailOn is False* so " & _
                    "there is nowhere to put the body and not a single one can be taken. NetDetailOn = True is needed too."
    Else
        Wv2Log.LogD "Wv2Pane.NetRespBodyOn = " & onOff
    End If
End Property


' ============================================================
' NetRespBodyMaxBytes (N-4) - upper limit of the response body (bytes)
'   *Held separately from the request body (issue 3, option S)* The appropriate values differ.
' ============================================================
Public Property Get NetRespBodyMaxBytes() As Long
    NetRespBodyMaxBytes = m_respBodyMax
End Property

Public Property Let NetRespBodyMaxBytes(ByVal n As Long)
    If n < 0 Then n = 0
    If n > NET_BODY_MAX_LIMIT Then n = NET_BODY_MAX_LIMIT
    m_respBodyMax = n
    Wv2Log.LogD "Wv2Pane.NetRespBodyMaxBytes = " & m_respBodyMax
End Property


' ============================================================
' NetRespBodyOrphan (N-4) - number of bodies that had no place to go
'   *Increases when the detail entry had vanished by the time the completion arrived*
'   (it overflowed from the ring / arrived after NetDetailClear). Not silently discarded.
' ============================================================
Public Property Get NetRespBodyOrphan() As Long
    NetRespBodyOrphan = m_respBodyOrphan
End Property


' ============================================================
' NetDetailRespBody (N-4) - response body of one detail
'   *Contents are filled in only when judged to be text* (same convention as N-2's body).
'   If it was compressed, the contents cannot be relied on, so look at the encoding column and the
'   head hex of NetDetailRespLine.
' ============================================================
Public Property Get NetDetailRespBody(ByVal idx As Long) As String
    Dim k As Long
    k = NetDetailRingIndex(idx)
    If k < 0 Then Exit Property
    NetDetailRespBody = m_detRespBody(k)
End Property


' ============================================================
' NetDetailRespLine (N-4) - the profile of the response body
'   6 items separated by vbTab:
'     bytes / truncated or not / text or not / content-encoding / hex of the first 32 bytes /
'     *arrived or not* (a 0-byte response and "not arrived yet" are different things)
'   *Always keeping the encoding and hex* is the point of N-4 (issue 2, option P').
' ============================================================
Public Property Get NetDetailRespLine(ByVal idx As Long) As String
    Dim k As Long
    k = NetDetailRingIndex(idx)
    If k < 0 Then Exit Property

    NetDetailRespLine = m_detRespBytes(k) & vbTab & CStr(m_detRespTrunc(k)) & _
                        vbTab & CStr(m_detRespIsText(k)) & vbTab & m_detRespEnc(k) & _
                        vbTab & m_detRespHex(k) & vbTab & CStr(m_detRespGot(k))
End Property


' ============================================================
' NetStartRespContent (N-4, Private) - start fetching the body
'
'   *Called from inside the WebResourceResponseReceived handler*
'   ResponseView is tied to the lifetime of args, so GetContent is fired here and nowhere else.
'   GetContent itself returns immediately, and the body reaches the completion handler later.
'
'   *AddRef and hold the ResponseView (issue 5, option beta)*
'     Whether it may be released before completion is not written in the SDK. If we guess wrong it shows up as
'     a crash, so we fall to the safe side. It is always Released on completion / cleanup.
' ============================================================
Private Sub NetStartRespContent(ByVal seqNo As Long, ByVal pResp As LongPtr)
    Dim h  As ComCallbackHandler
    Dim hr As Long
    Dim ky As String

    If pResp = 0 Then Exit Sub

    ' *Do not fire if there is no place to put it* The body goes into the detail entry, so
    '   taking the body of a request whose detail was not taken would just be thrown away.
    If NetFindDetailBySeq(seqNo) < 0 Then Exit Sub

    ky = "r" & seqNo
    ' Do not fire twice for the same seq (it should be one response per request, but just in case)
    On Error Resume Next
    Dim dummy As Variant
    dummy = m_respPending(ky)
    If Err.Number = 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0

    Set h = AcquireHandlerFor(HK_ResponseGetContent, Me)
    If h Is Nothing Then Exit Sub
    h.callbackId = seqNo

    ' *Keep the ResponseView alive until completion*
    ComAddRef pResp
    m_respPending.Add pResp, ky

    hr = dcf(pResp, 6, "", h.Slot + VTABLE_OBJ_OFFSET)   ' GetContent
    If hr <> 0 Then
        ' Could not fire. Undo the hold.
        ComRelease pResp
        m_respPending.Remove ky
        Wv2Thunks.Thunks_DiscardHandler h   ' English fork
        Set h = Nothing
    End If
End Sub


' ============================================================
' OnResponseContentCompleted (N-4)
'   Called from ComCallbackHandler via dynamic binding. Public Sub is required.
'   A one-shot family, so no origin prefix (9.6a convention).
'
'   *IStream is valid only during Invoke* so read it to the end synchronously.
'   *Like the receivers since N-1, the error is not re-raised*
' ============================================================
Public Sub OnResponseContentCompleted(ByVal cbId As Long, _
                                      ByVal errorCode As Long, _
                                      ByVal pStream As LongPtr)
    Dim errNum As Long

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    OnResponseContentCompleted_Core cbId, errorCode, pStream

GuardExit:
    errNum = Err.Number
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        On Error Resume Next
        m_respBodyOrphan = m_respBodyOrphan + 1
    End If
End Sub


Private Sub OnResponseContentCompleted_Core(ByVal cbId As Long, _
                                            ByVal errorCode As Long, _
                                            ByVal pStream As LongPtr)
    Dim k       As Long
    Dim nBytes  As Long
    Dim trunc   As Boolean
    Dim buf()   As Byte
    Dim enc     As String
    Dim cType   As String
    Dim ky      As String

    ' --- Release the held ResponseView (cleanup for issue 5, option beta) ---
    ky = "r" & cbId
    On Error Resume Next
    Dim pKept As LongPtr
    pKept = CLngPtr(m_respPending(ky))
    If pKept <> 0 Then
        ComRelease pKept
        m_respPending.Remove ky
    End If
    Err.Clear
    On Error GoTo 0

    ' --- Look for the place to put it ---
    k = NetFindDetailBySeq(cbId)
    If k < 0 Then
        ' *The detail entry had vanished by completion* (overflowed / cleared)
        m_respBodyOrphan = m_respBodyOrphan + 1
        Exit Sub
    End If

    ' *If we got this far, it "arrived"* Even for 0 bytes or a failure, distinguish it from not arrived
    m_detRespGot(k) = True

    If errorCode = E_NO_DATA Then
        ' *A response with no body (204 etc.)* is not a failure. Leave it as 0 bytes.
        m_detRespBody(k) = ""
        m_detRespBytes(k) = 0
        Exit Sub
    ElseIf errorCode <> 0 Then
        m_detRespBody(k) = ""
        m_detRespEnc(k) = "(fetch failed hr=&H" & Hex(errorCode) & ")"
        Exit Sub
    End If

    If pStream = 0 Then
        ' *A response with no body* (204 / 304 etc.). Leave it as 0 bytes.
        m_detRespBytes(k) = 0
        Exit Sub
    End If

    ' --- Read (*no need to rewind with Seek* since it is a copy of the response) ---
    If Not NetReadStream(pStream, m_respBodyMax, nBytes, trunc, buf) Then Exit Sub

    m_detRespBytes(k) = nBytes
    m_detRespTrunc(k) = trunc
    If nBytes <= 0 Then Exit Sub

    ' --- *Do not assume what came; keep everything (issue 2, option P')* ---
    enc = NetHeaderValue(m_detRespHeaders(k), "content-encoding")
    cType = NetHeaderValue(m_detRespHeaders(k), "content-type")
    m_detRespEnc(k) = enc
    m_detRespHex(k) = NetHexHead(buf, nBytes)     ' *always kept*
    m_detRespIsText(k) = NetIsTextType(cType)
    If m_detRespIsText(k) Then
        m_detRespBody(k) = Utf8BytesToString(buf, nBytes)
    End If
End Sub


' ============================================================
' NetFindDetailBySeq (N-4, Private) - find the position in the detail ring from a serial number
'   -1 if not found. The same lookup as N-3's NetAttachRespHeaders.
' ============================================================
Private Function NetFindDetailBySeq(ByVal seqNo As Long) As Long
    Dim i As Long
    Dim k As Long

    NetFindDetailBySeq = -1
    For i = m_detCount To 1 Step -1
        k = NetDetailRingIndex(i)
        If k >= 0 Then
            If m_detSeq(k) = seqNo Then
                NetFindDetailBySeq = k
                Exit Function
            End If
        End If
    Next i
End Function


' ============================================================
' NetRespPendingClear (N-4, Private) - release all held ResponseViews
'   *It may end without the completion ever arriving* (capture stopped / Pane closed).
'   To avoid leaking, Release them all together here.
' ============================================================
Private Sub NetRespPendingClear()
    Dim v As Variant

    If m_respPending Is Nothing Then Exit Sub

    On Error Resume Next
    For Each v In m_respPending
        If CLngPtr(v) <> 0 Then ComRelease CLngPtr(v)
    Next v
    Err.Clear
    On Error GoTo 0

    Set m_respPending = New Collection
End Sub
