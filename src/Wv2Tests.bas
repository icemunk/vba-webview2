Attribute VB_Name = "Wv2Tests"
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  Stage N-4 (response body) ---
'
'   Test_N4_Body ... regression test
'   Test_N4_Help ... procedure
'
'   *N-4 is the first stage that is asynchronous*
'     GetContent takes a handler and returns immediately; the body arrives later.
'     Verdicts must always wait with *N4Wait (until the body arrives)*. In N-3 we
'     waited for "until the detail appears", read an empty value, and burned three real-machine runs
'     (design principle 120). N4Wait is a dedicated helper so we do not repeat the same mistake.
'
'   *The biggest unknown is compression*
'     The SDK does not say whether GetContent returns the decoded body. httpbingo's
'     /gzip contains "gzipped": true in its payload, so exactly one of *the body
'     reads as text / it starts with 1F 8B* must hold. That is what the verdict uses.
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  Stage N-3 (response status and headers) ---
'
'   Test_N3_Status ... regression test. Targets are the same 2 routes as N-1 / N-2
'   Test_N3_Help   ... procedure
'
'   *(1) to (4) need an external host*
'     An unreachable local target (plan F) gets *no response*, so verifying the status
'     needs a real destination. Conversely, plan F serves as the control for
'     "no response = blank" (that is (5)).
'
'   *This event has no filter*
'     Images and CSS all arrive. Responses not in the list are only counted in
'     NetRespUnmatched and discarded. *Seeing large numbers is normal*.
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  Stage N-2 (request headers and body) ---
'
'   Test_N2_Detail ... regression test. Same 2 routes as N-1 (plan F / plan D)
'   Test_N2_Help   ... procedure
'
'   *The crux is "reading does not break the traffic"*
'     Reading an IStream advances its position, so failing to Seek(0) back sends an
'     empty body. "The POST seemed to succeed" is not evidence, so *count by matching
'     against the body the destination (httpbingo) echoes back* (design principle 112). This one case needs an external host.
'
'   *A Cookie cannot be produced when the target is unreachable*, so the masking mechanism itself is
'   checked by adding an "always present header" = User-Agent to the masked set.
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  Stage N-1c (stop assuming the type) ---
'
'   *What the N-1b real-machine run revealed*
'     *WebView2 reports fetch() requests as XML_HTTP_REQUEST(7), not FETCH(8)*
'     (we fired 10 fetches on the real machine and all 10 arrived as XHR).
'     The N-1b test had assumed "for fetch the type should be FETCH", so
'     *requests were captured yet 3 FAILs were reported*. The implementation is unharmed.
'
'   * What was fixed (test side only)
'     1. The verdict is now "was it captured", and *the type is not assumed*
'        As a bonus, the false alarm on localOk / netOk ("nothing arrived") is gone too
'     2. *Added a verdict that counts "fetch arrives as XHR type"*
'        The spec fact we found becomes a regression test as is
'     3. *One drain is inserted right after (3)*
'        In N-1b, dropping the capacity to 3 at (4) wiped the contents of (2), so
'        we had to *fill in by guesswork* "seq 1/2 should have been f-local and f-net".
'        From now on no guessing is needed (design principle 112).
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  Stage N-1b (move the targets outside the virtual host) ---
'
'   *What the first N-1 real-machine run revealed*
'     Requests served through a virtual host (SetVirtualHostNameToFolderMapping)
'     do not show up in WebResourceRequested. The first Test_N1_Capture failed completely
'     because *every single target was on the virtual host*, and the wiring was
'     correct from the start (Test_N1_Site captured all 10).
'
'   *Targets are now 2 routes + 1 control route (issue 1, plan G)*
'     Plan F  http://127.0.0.1:59999/...  unreachable local (zero external dependency)
'     Plan D  https://httpbingo.org/...   external service (needs the net)
'     Control data.json on the virtual host   *count that it is NOT captured*
'   One real-machine run checks both to see which targets are usable.
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  Stage N-1 (network request capture) ---
'
'   Test_N1_Capture ... *a regression test that is self-contained with its own HTML* (issue 8)
'   Test_N1_Watch   ... *start capturing the traffic of the currently open tab* (practical entry point)
'   Test_N1_Drain   ... flush the accumulated entries to the log and empty the buffer
'   Test_N1_Stop    ... stop capturing (also removes the filter)
'   Test_N1_Help    ... procedure
'
'   *Do not rely on real sites* The verification page is written out to %TEMP%\Wv2NetProbe
'   and opened on the virtual host https://appassets.netprobe/ (same technique as 9.26b).
'   Form POST, fetch and XHR all complete inside this folder.
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  Stage K-4a (isolating the cause of slow release) ---
'
'   Test_K4_Quiet ... *quiet all tabs before closing the form*
'                   Measures whether release is slow "because the pages keep running".
'   Test_K4_Help  ... procedure.
'
'   *What the measurements have shown so far (K-4a)*
'     Bare state (start -> close at once)  ... total release 0.336 s
'     After opening one Maps page          ... total release *6.718 s* (20x)
'     Difference LogEcho = True / False    ... *zero* (Debug.Print is not the culprit)
'     The slowness is not in a specific step; *every COM round trip takes 300-500 ms*
'     (8-16 ms in the bare state). One Ctrl_Close call is even heavier at 1.4 s.
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  Stage D-7e (verifying progress and abort) ---
'
'   Additions in D-7b:
'     Test_D7_StatusBar ... *probe* measures Excel's raw behavior without going through
'                         product code. Separates whether the StatusBar verdict that FAILed in D-7
'                         was a problem on the verdict side or Excel's specification (principle 103).
'     Test_D7_Resume    ... *automatic* artificially builds the state after an abort and checks
'                         without Esc that we can resume from it (for regression checks).
'
'   Carried over from D-7:
'     Test_D7_Cancel ... *no network needed* the abort entry point, clearing at entry, how the
'                      denominator is counted, and that no tab is opened when there are 0 items.
'     Test_D7_Sheet  ... *real site + press Esc by hand* from abort through to resume.
'     Test_D7_Help   ... procedure.
'
'   * *D-7e: make verification that a human observes visible to the human*
'     In the D-7d real-machine run the abort did work, yet to tarboh it looked
'     as if "it had not stopped". *Because the second run was started automatically right after the abort*.
'     Moreover the repeatedly pressed Esc also hit the second run, and 3 verdicts FAILed.
'     -> After an abort, *announce it clearly, pause for a beat, and wait until Esc is released*
'       before trying the resume (TestWaitEscReleased).
'     Test_D7_Key was rebuilt for the same reason. Even when it said "10 seconds from now",
'     *it finished before it could be read*, so it now waits until the key is pressed.
'
'   *Verification of Esc itself is not automated* There was a plan to set a flag with SafeTimer
'   instead, but SafeTimer needs WithEvents = a new class module would be required,
'   so it was put off. Instead *resuming from the post-abort state* is automated in Test_D7_Resume,
'   and that covers every regression check. The Esc path is checked once by a human in Test_D7_Sheet.
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  Stage D-4a (verifying the in-page SPA probe) ---
'
'   Additions in D-4a:
'     Test_D4_Probe ... *the first move of D-4* verifies in one batch the probe's installation,
'                     health check, rebuild, enumeration and round-trip cost.
'     Test_D4_Help  ... procedure.
'
'   *What is being measured (unknowns 1-4 of the D-4 issues)*
'     Unknown 1 Does the wrap survive?       ... (6) breaks it on purpose and confirms the automatic repair
'     Unknown 2 Load of observation          ... (2) compares round-trip time before and after installation
'     Unknown 3 Does it vanish on navigation? ... (8) confirms the generation restarts from scratch after page navigation
'     Unknown 4 Does it avoid attaching to extra Panes ... (1) confirms it is "not created until called"
'
'   The verification page (BuildD4ProbeHtml) is made *completely quiet right after loading*.
'   Noise (periodic DOM rewriting) starts only when startNoise() is fired.
'   The same page is used for verifying the D-4b quiescence wait as well.
'''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  Stage D-3 (verifying writes and operations) ---
'
'   Additions in D-3:
'     Test_D3_Probe_Promise ... *run this first* measures unknown 1.
'                             Checks whether ExecuteScript waits for a Promise.
'                             This result decides "where to run the wait loop" (issue 3).
'     Test_D3_Write         ... verifies Value = / Click / SetAttribute in one batch with its own HTML.
'     Test_D3_Framework     ... checks *whether writes go through the native setter* using
'                             monitoring that imitates React's value tracker (unknown 3).
'     Test_D3_Help          ... procedure for the 3 above, and an explanation of the log to look at.
'
'   The verdict helpers (TestEq / TestBool / D2El / D2WaitTitle) are reused from D-2.
'   *In D-3 the output target of verdicts was changed to Wv2Log, and the D-1 / D-2 tests were moved onto it too*
'   (the Immediate window is flooded by the ExecuteScript plumbing log and cannot be read).
'   Not a single existing Test_* was changed.
'
'   *The verification page (BuildD3ProbeHtml) is its own HTML, as in D-2*
'     The difference from the D-2 page is that *monitoring JS* is planted on the page side:
'       window.__p = {inputs, changes, clicks, clickInfo,
'                     trackedSet, notified, ignored}
'     - number of input / change firings (confirms "fire both" from issue 6)
'     - number of clicks and the identity of the event (type / bubbles / isTrusted)
'     - *a React-like value tracker* is laid over #react, and
'       it counts assignments through the tracker (trackedSet), the number of times the framework
'       noticed the change (notified) / failed to notice (ignored).
'''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  Stage D-2 (element registry + verifying Wv2Element) ---
'
'   Additions in D-2:
'     Test_D2_Find  ... opens its own verification page and verifies GetElementById / QuerySelector and
'                     the 8 read members of Wv2Element in one batch, with expected values.
'     Test_D2_Stale ... navigates the page while holding an element and confirms that the generation
'                     changes and IsStale becomes True, that reads fail when stale, and
'                     that it can be re-acquired on the new page.
'     Test_D2_Help  ... procedure for the 2 above, and an explanation of the log to look at.
'
'   *The verification page is its own HTML (issue 8, plan b)*
'     External sites go down and their DOM changes without notice (design principle 75, the lesson of 9.28).
'     It is fed in with NavigateToString, so there is neither an external dependency nor file I/O.
'     What it covers: id / class / arbitrary attributes / Japanese / symbols / line breaks /
'     innerHTML containing tags (spec fact 30) / value of input, textarea and select /
'     empty elements / elements with no attributes.
'
'   Not a single existing Test_* was changed.
'
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  Stage D-1 (verifying EvalSync) ---
'
'   *The stage numbering scheme has changed* Up to 9.32b, the 9.xx numbers were milestones on the axis
'     of "browser feature work". From D-1 on, a symbol per axis is used (D = DOM / page automation).
'
'   Additions in D-1:
'     Test_D1_Eval  ... batch verification of EvalSync. Runs types, Japanese, escaping, exceptions, syntax errors,
'                     timeouts and recovery after a timeout in one go.
'     Test_D1_Guard ... checks that the in-callback guard of issue 7 fires and that ResetCallbackGuard works.
'     Test_D1_Help  ... procedure for the 2 above, and an explanation of the log to look at.
'
'   Not a single existing Test_* was changed.
'
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  Stage 9.32b (verification procedure for horizontal wheel scrolling) ---
'
'   Additions in 9.32b:
'     *Test_9_32b_Help (real-machine procedure)* Only JS changed, so there is no pure-logic test.
'       (1) Does the tab row scroll horizontally with the wheel?
'       (2) Does the scroll position stay put (not jump back to the left edge) when the title update syncs?
'       (3) Does a new tab added with + come into a visible position automatically?
'       (4) Regression: switching feel / D&D / closing
'
'   Not a single existing Test_* was changed.
'

''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  Stage 9.32 (verification procedure for tab-switch feel and scrollbar) ---
'
'   Additions in 9.32:
'     *Test_9_32_Help (real-machine procedure)*
'       9.32 also changes only JS and CSS, so *there is no pure-logic test*
'       (judged by eye and by feel). What to look at:
'         (1) Does the tab color change the instant it is clicked, with the body switch catching up afterwards?
'         (2) Even with more tabs, no horizontal scrollbar appears and the tabs do not get squashed?
'         (3) Can the tab row be scrolled horizontally with the wheel?
'         (4) Regression: are D&D reordering / closing / the settings tab still intact?
'
'   Not a single existing Test_* was changed.
'

''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  Stage 9.31 (verification procedure for the tab bar's appearance) ---
'
'   Additions in 9.31:
'     *Test_9_31_Help (real-machine procedure)*
'       9.31 changes only CSS, so *there is no pure-logic test*
'       (judged by eye). The procedure has 2 parts:
'         Part 1 = visual check of (A-1) hover background / (A-2) show/hide control of the close button
'         Part 2 = *open item 1 from 9.30 = end-to-end check of v0_5_2*
'                   new-tab expansion of POST links / pre-warm / drag reordering.
'                   The clean-built v0_5_2 of 9.30 only had "passes compilation" and
'                   "search engine switch from the settings tab" verified,
'                   so this round, which touches the tab bar, steps through the rest together.
'
'   Not a single existing Test_* was changed.
'

''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  Stage 9.29 (verifying custom search engines) ---
'
'   Additions in 9.29:
'     *Test_9_29_Custom (pure logic)*
'       Checks SetCustomEngine / GetCustomTemplate of Wv2SettingsBridge,
'       PreviewUrlForTemplate and the saving / loading of custom (2 lines: engine + template)
'       against expected values. WebView2 is not started.
'       It touches the real file (%APPDATA%\Wv2Browser\settings.txt), so at the top it
'       saves *engine and template as a pair* and writes them back at the end.
'
'     *Test_9_29_Help (real-machine procedure)*
'       Custom input box on the settings tab -> apply -> search -> restore on restart, end to end.
'
'     *Test_9_20_Persistence's save/restore was made custom-aware*
'       In 9.29 settings.txt became 2 keys, so writing back with Debug_SaveEngineName would
'       "erase the template line of someone whose original was custom".
'       Changed to save and restore with Debug_SaveSettings (engine + template).
'
' --- Wv2Tests.bas  Stage 9.28b (POST probe destinations made 3 routes / main one moved to httpbingo.org) ---
'
'   In the 9.28 real-machine run httpbin.org *kept returning 503 for server-side reasons*, so the main
'   re-check could not be done. A plain GET (typed directly into the address bar) also gave 503, which
'   confirmed that the 405 that should have come back from a healthy server was masked by server downtime.
'   httpbin.org is a free public service with long-standing reports of periodic 503s.
'   -> *Do not make the verification base depend on a single external service*: increase destinations to 3.
'
'   *Changes (BuildPostProbeHtml and the explanatory text only. Not one line of production files is touched)*
'     Button 1 (main / POST + _blank)        : httpbin.org -> *httpbingo.org*
'     Button 2 (sanity / POST same tab)      : httpbin.org -> *httpbingo.org*
'     Button 3 (control / GET + _blank)      : httpbin.org/get -> *httpbingo.org/get*
'                                        (aligned to the same server as the main one to narrow the variables to one)
'     Button 4 (spare 1)                     : postman-echo.com -> *httpbin.org* (former main one demoted)
'     Button 5 (spare 2)                     : *newly added* postman-echo.com
'
'   *Why httpbingo.org was made the main one, and a caution*
'     - GET returns 405 Method Not Allowed (confirmed on the real machine). The same picture as the
'       workplace Tomcat holds as is.
'     - The POST response JSON has a *method field*, so you can confirm directly on screen
'       that it has not turned into a GET (an advantage httpbin lacks).
'     - *Caution* httpbingo's error page has no title element, and Edge uses the URL as
'       a substitute title. Hence *pass/fail cannot be read from the Immediate log*
'       (httpbin showed 405 METHOD NOT ALLOWED in the title).
'       Always judge by *what is shown on screen*.
'
'   *JS stays at zero* The policy of building with a plain form only is kept (JS inside VBA strings invites accidents).
'
' --- Wv2Tests.bas  Stage 9.28 (updating verification procedures for pre-warm delegation becoming the default) ---
'
'   Additions / changes in 9.28:
'     *Test_9_28_Help (new)*
'       Real-machine procedure with the default now 2 (pre-warm delegation). The main aim is to confirm *without switching modes*
'       that POST links give 200 in our own tab.
'     *The wording of Test_9_27_Help / Test_9_27_Status was fixed to assume default 2*
'       The 9.27 procedure was written on the premise "default 0 -> take the control, then switch to 2".
'       With the default now 2, to see the control (405) you must first run Test_9_27_Mode_Legacy.
'       Only that one point was fixed.
'
'   *Test_9_27_Mode_Legacy / _Popup / _Prewarm / _Status / _Handled_On / _Off are
'     unchanged from 9.27* Only the default changed; the meaning of the entry points did not.
'
' --- Wv2Tests.bas  Stage 9.27 (main target: verifying pre-warm delegation via put_NewWindow) ---
'
'   Additions in 9.27:
'     *Test_9_27_Mode_Legacy / _Popup / _Prewarm (mode switching)*
'       Thin entry points that switch Wv2Browser.NewWindowMode between 0 / 1 / 2. Effective immediately even while running.
'     *Test_9_27_Status (state dump)*
'       Confirms in one line whether the spare tab (pre-warm) is warmed up.
'     *Test_9_27_Handled_On / _Off (insurance switch for issue 5)*
'       Toggles whether put_Handled(TRUE) is also set after put_NewWindow.
'       For isolating symptoms such as the spare tab staying white / opening twice.
'     *Test_9_27_Help (real-machine procedure)*
'       Reuses as is the POST probe page of 9.26b (Test_9_26_PostProbe).
'       The pass condition is *delegation OFF (mode 0) gives 405 / mode 2 gives 200 + form echo*.
'
'   *No new verification page is made* The postprobe.html of 9.26b already has
'     the same picture as the workplace, "GET gives 405 / POST gives 200 + echoes the body",
'     so it is exactly adequate for judging the main target as well.
'
' --- Wv2Tests.bas  Stage 9.26b (adding the POST probe page) ---
'
'   What the 9.26 real-machine verification revealed:
'     The probe mechanism (toggle / branching / runtime delegation / no reference leak) all held.
'     However, the popups of the sites tried were *plain GETs*, so even with delegation OFF they opened normally in
'     our own tab (isSuccess: True), and the 405 was not reproduced.
'     In other words, the evidence for the main question "does delegation preserve the POST?" has not been obtained yet.
'
'   *Aim of this step* Synthesize the 405 picture at home without going to the workplace, and judge the main question.
'     With https://httpbin.org/post as the destination
'       - hitting it with GET  -> 405 METHOD NOT ALLOWED   (= exactly the same picture as the workplace Tomcat)
'       - hitting it with POST -> 200 + the sent body echoed as JSON
'     so a single button yields both "reproduction of 405" and "direct evidence that the body was carried".
'     If delegation OFF gives 405 / ON gives 200, pattern A is confirmed.
'
'   *What was added (all self-contained in Wv2Tests.bas. Not one line of production files is touched)*
'     - Test_9_26_PostProbe      : writes out the verification page and opens it in a new tab
'     - Test_9_26_PostProbe_Help : the procedure and how to read the results
'     - BuildPostProbeHtml       : HTML of the verification page (private)
'     - WriteUtf8NoBom           : UTF-8 (no BOM) write-out helper (private)
'
'   *Design decision notes*
'     - It is opened via a virtual host mapping (https://appassets.postprobe/postprobe.html).
'       NavigateToString makes the origin "null", changing the Referer and form-submission
'       conditions from production, so it is not used, to keep the verification pure.
'       It plugs straight into the existing Wv2Browser.AddTabWithUrlForSpa.
'     - Writing uses ADODB.Stream with UTF-8 (no BOM). UserForm1.WriteSpaAppFolder writes
'       ANSI (CP932) with Print #, so putting Japanese into an HTML with meta charset=UTF-8
'       garbles it. Aligned with the approach proven in Wv2TabBar/Wv2NavBar.
'       (ADODB.Stream is late-bound, so it does not violate the zero-references policy)
'     - *Zero JS* Built with a plain <form> only. JS inside VBA strings invites accidents, so
'       JS is not brought in needlessly for a requirement that HTML alone satisfies.
'
' --- Wv2Tests.bas  Stage 9.26 (NewWindowRequested delegation probe) ---
'
'   Additions in 9.26:
'     *Test_9_26_Popup_On / Test_9_26_Popup_Off (real-machine switches)*
'       Thin entry points that only toggle PopupDelegateMode of the running Wv2Browser.
'       They go through UserForm1.CurrentBrowser (the 9.17 scaffold), so UserForm1.frm is unchanged.
'     *Test_9_26_Popup_Help (real-machine procedure)*
'       Outputs the procedure for stepping on a POST link in the workplace core system to see whether the 405 disappears, and
'       how to read the results (pattern A/B).
'
'   *This round has no pure-logic verification* What we want to see is the real-machine behavior itself:
'     "does the WebView2 runtime continue the navigation?", so it cannot be made a
'     verification that does not start WebView2 (same judgment as 9.21-9.25).
'
' --- Wv2Tests.bas  Stage 9.25b (removing the verification probe / making zero arguments permanent) ---
'
'   Since the 9.25 verification concluded with "pattern A = safe", Test_9_25_Probe_Help was deleted.
'   The probe itself (HostProbe / HostProbeMark of Wv2NavBar and probeParamless in JS) was also
'   removed at the same time, so leaving only the procedure Sub would leave nothing runnable.
'
'   *Conclusion (spec fact 22. Details in the development notes)*
'     The sync proxy of hostObjects does not call into the VBA side at member-fetch time.
'     The call is deferred until the actual (). Hence a Public Function with zero arguments can
'     be exposed as is, and the dummy argument "reserved" of 9.24 was unnecessary.
'     Measured log: (1) just before fetch -> (2) just after fetch -> *execution* -> (3) f() succeeded ->
'     *execution* -> (5) direct call succeeded. No execution between (1) and (2); exactly 2 in total.

''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  Stage 9.25 (verification: whether a zero-argument Host method is possible + removing reserved) ---
'
'   Additions in 9.25:
'     *Test_9_25_Probe_Help (real-machine procedure)*
'       The round that settles the 9.24 open item "was the dummy argument reserved really necessary?".
'       Removed reserved from NavBar to return to zero arguments, and at the same time planted a verification-only
'       probe with zero side effects (HostProbe / HostProbeMark) (plan C).
'
'       *In this round the answer about the mechanism comes out just by starting up* The probe is run by NavBar's JS on its own
'       500ms after the ready message is sent, so tarboh only has to start normally and
'       read the sequence of [PROBE] lines. Then press <- -> [Reload] once each to check that
'       the real thing is not executed twice.

''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  Stage 9.24 (verification of NavBar moving to hostObjects + Host consolidation) ---
'
'   Additions in 9.24:
'     *Test_9_24_HostNavBar_Help (real-machine procedure)*
'       Real-machine procedure for the round that moved NavBar's back/forward/reload/navigate to the hostObjects route and consolidated the processing into
'       HostBack/HostForward/HostReload/HostNavigate.
'       It checks a real URL Navigate + JS execution + the hostObjects route working, so there is no pure-logic
'       verification (same judgment as 9.21-9.23).
'
'       *The verdict of this round has three parts*
'         (1) Did the 4 commands go via host? (HostXxx log appears and the corresponding cmd log
'             of OnPaneWebMessage does not)
'         (2) *Is there no double execution? (issue 4b)* Press <- once and check HostBack
'             appears on only 1 line. If 2 lines appear, or the back of OnPaneWebMessage follows,
'             even the dummy argument reserved did not block it (report it)
'         (3) *Are string arguments intact? (spec fact 21 candidate)* Check whether a Japanese search term and a URL with
'             symbols appear in the HostNavigate log without garbling
'
'       For regression checks: the ready handshake / discarding edits on Escape / the editing flag /
'       the TabBar side (9.21-9.23) must be intact.

''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  Stage 9.23 (cleanup: verifying the Host consolidation of activate/close/reorder) ---
'
'   Additions in 9.23:
'     *Test_9_23_Consolidation_Help (real-machine procedure)*
'       9.23 is a cleanup stage (plan C-1) that consolidated the bodies of the activate/close/reorder cases of OnPaneWebMessage into
'       fire-and-forget delegation to HostActivate/HostClose/HostReorder. The JS-side hybrid is unchanged, so in
'       normal operation the hostObjects route works and the Case (= fallback route) is not taken at all.
'       Hence what to look at on the real machine is "no regression" = activate/close/reorder
'       still work via host as before (HostXxx log) with the UX unchanged, and
'       still work via host as before (HostXxx log) with the UX unchanged, and that
'       ready/newtab/settings are intact. There is no pure-logic verification because a real WebView2 start is
'       required (same judgment as 9.21/9.22).
'

''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  Stage 9.22 (trial: verifying close / reorder moved to hostObjects) ---
'
'   Additions in 9.22:
'     *Test_9_22_HostCloseReorder_Help (real-machine procedure)*
'       Real-machine procedure for the round that moved the TabBar's close / reorder to go through hostObjects.sync.TabBar.HostClose /
'       HostReorder. Confirms that the close x and reorder drag
'       produce the via-host log and no OnPaneWebMessage, that activate (9.21) is
'       still via host, and that ready/newtab/settings are still
'       postMessage (no regression). A real URL Navigate + JS execution +
'       the hostObjects route working require a real start, so there is no pure-logic verification.
'
' --- Wv2Tests.bas  Stage 9.21 (trial: verifying activate moved to hostObjects) ---
'
'   Additions in 9.21:
'     *Test_9_21_HostActivate_Help (real-machine procedure)*
'       Procedure for checking on the real machine the result of the trial migration of the TabBar's activate to go through
'       hostObjects.sync.TabBar.HostActivate. There is no pure-logic verification (it checks a real URL
'       Navigate + JS execution + the hostObjects route working, so a real
'       WebView2 start is required). Open several tabs and switch by clicking, and confirm that the Immediate window shows
'       "HostActivate: index=N (via host)" and does not show the former
'       "OnPaneWebMessage: msg={"cmd":"activate"...}".
'
' --- Wv2Tests.bas  Stage 9.20 (verifying persistence of the search engine selection) ---
'
'   Additions in 9.20:
'     *Test_9_20_Persistence (pure logic)*
'       Directly calls Debug_SaveEngineName / LoadEngineName /
'       LoadEngineTemplate of Wv2SettingsBridge, and checks the save -> load round trip and the fallback
'       for a missing file / an unknown engine name.
'       WebView2 is not started. It touches the real file (%APPDATA%\Wv2Browser\
'       settings.txt), so at the top it saves the current stored value aside, and
'       writes it back at the end so as not to break the user's settings.
'
'     *Test_9_20_Persistence_Help (real-machine procedure)*
'       Start -> change the engine in settings -> fully quit Excel -> restart and confirm that search starts with
'       the previous engine (restoring involves running Class_Initialize,
'       so it is confirmed by the real-machine procedure).
'
' --- Wv2Tests.bas  Stage 9.19 (verification procedure for preventing duplicate settings tabs) ---
'
'   Additions in 9.19:
'     *Added Test_9_19_SettingsTabDedup_Help*
'       Preventing duplicate settings tabs (mashing [gear] does not add settings tabs) involves WebView2's actual
'       tab creation and activation, so it is confirmed by a real-machine procedure, not pure-logic
'       verification. Typing Test_9_19_SettingsTabDedup_Help in the Immediate window shows the
'       procedure. The crux of the logic (setting the IsSettings flag + avoiding duplicates by scanning m_tabs)
'       is covered by code review of Wv2Pane / Wv2Browser.
'
' --- Wv2Tests.bas  Stage 9.18 (pure-logic verification of Wv2SettingsBridge) ---
'
'   Additions in 9.18: added Test_9_18_BridgeLogic. Creates a Wv2SettingsBridge, binds a
'   real Wv2Browser with BindBrowser, then directly calls the same methods
'   the settings screen's JS calls (SetEngine / GetEngine / PreviewUrl / PreviewUrlFor / EngineList)
'   and checks name <-> template resolution, preview URL generation, and absence of side effects.
'   WebView2 is not started (it only calls the bridge + Browser's Debug_*).
'   NOTE: whether the hostObjects route works on the real machine (order A) is confirmed by the real-machine procedure in Help.
'   Added the generic helpers CheckEq (string match) / CheckBool (boolean).
'
' --- Wv2Tests.bas  Stage 9.17 (verifying search engine preset resolution) ---
'
'   Additions in 9.17: added Test_9_17_SearchPresets. Checks in pure logic whether the search URL
'   follows that engine correctly when the template strings that UserForm1.UseSearchEngine resolves
'   preset names ("google"/"bing"/"duckduckgo"/"yahoo") to are Let directly into Wv2Browser.
'   Does not start the UserForm (independent of m_browser),
'   it only news a Browser and calls Debug_NormalizeUrl / Debug_SearchEngine.
'   NOTE: UseSearchEngine itself (name -> template resolution + not-started guard) is confirmed on the real machine
'     (via CurrentBrowser). Here we only guarantee the correctness of the preset values.
'
' --- Wv2Tests.bas  Stage 9.16 (cleanup: consolidating feature verification Subs) ---
'
'   A standard module that collects the "pure-logic verification Subs" that do not start WebView2.
'   Only calls Wv2Browser's public verification entry points (Debug_NormalizeUrl / Debug_SearchEngine /
'   Debug_TabOrderTitles / Debug_SeedDummyTabs / Debug_ActiveIndex / MoveTab)
'   and never touches the thunk foundation or WebView2 itself.
'
'   *Why it was separated (plan 1-A)*
'     These verifications had originally piled up in Wv2Thunks.bas, but Wv2Thunks is the heart module of
'     "machine-code thunks, vtables and memory primitives", and having feature verification unrelated to
'     thunks live there was an encroachment on its responsibility
'     (inconsistent with the principle of "splitting by concern" that split out Wv2Json). So this module was created as a
'     receptacle dedicated to verification, and everything from Test_9_14 on was consolidated here.
'     Later Test_9_17, Test_9_18... are added here too (no more new files).
'
'   *Portability* This module is verification code only, so at production porting one can decide "if tests are
'     not needed, do not import Wv2Tests.bas". Production code and
'     test code are physically separated, so portability actually goes up.
'
'   *What goes here / what does not*
'     Goes here  : pure-logic verification that only calls Wv2Browser's Debug_* (no WebView2 start needed).
'     Does not   : verification of the thunk foundation itself (Test_VirtualAllocFree_Roundtrip etc.) stays in
'              Wv2Thunks because it is tightly coupled to its Private members.
'              Verification involving a real start (StartWebView2_Full etc.) stays in UserForm1
'              (because it holds the UserForm's instance state m_browser etc.).
''''''''''''''''''''''''''''''''''

Option Explicit

' --- Verdict counters (counted by TestBool / TestEq / D1Case, printed by TestCountPrint) ---
'     *The Immediate window gets flooded by the ExecuteScript plumbing log,
'     so verdicts are also kept in the log file (via K-1's Wv2Log).*
Private m_okCount As Long
Private m_ngCount As Long

' --- N-1: virtual host name and folder name for placing the verification page ---
Private Const N1_HOST   As String = "appassets.netprobe"
Private Const N1_FOLDER As String = "Wv2NetProbe"

' --- N-1b: verification targets (*both outside the virtual host*) ---
'   N1_LOCAL ... an unreachable local address. The connection always fails, but the request is sent.
'     *Pick a port that avoids the numbers Chromium blocks* (1/7/9/11/13/.../10080 etc.)
'     *It is not stopped as mixed content even with http* 127.0.0.1 is by spec a
'     "potentially trustworthy origin", so it can be called from an https page.
'   N1_NET   ... an external service. It needs the net, but in exchange it arrives cleanly.
Private Const N1_LOCAL  As String = "http://127.0.0.1:59999"
Private Const N1_NET    As String = "https://httpbingo.org"


' ============================================================
' Test_9_14_SearchFallback  (Stage 9.14 logic verification of the search-term fallback)
'
'   *Just type  Test_9_14_SearchFallback  in the Immediate window*
'
'   WebView2 is not started. Creates a Wv2Browser and calls Debug_NormalizeUrl to check in one batch
'   that input is correctly routed to "URL" vs "Google search", and that percent-encoding of
'   Japanese and symbols is correct.
'   Prints [OK] when it matches the expected value and [NG] when not. A pass/fail summary at the end.
' ============================================================
Public Sub Test_9_14_SearchFallback()
    Dim b As Wv2Browser
    Set b = New Wv2Browser
    ' NOTE: Initialize of Wv2Browser is not needed (Debug_NormalizeUrl only calls NormalizeUrl,
    '    and never touches tabs, the Environment or windows).

    Dim total As Long, pass As Long
    total = 0: pass = 0

    Debug.Print "==== Test_9_14_SearchFallback start ===="

    ' --- (1) has a scheme -> as is ---
    Check b, "https://example.com", "https://example.com", total, pass
    Check b, "http://example.com/a?b=1", "http://example.com/a?b=1", total, pass
    Check b, "about:blank", "about:blank", total, pass
    Check b, "file:///C:/tmp/x.html", "file:///C:/tmp/x.html", total, pass

    ' --- (2) has a dot, no whitespace -> URL (https:// prefixed) ---
    Check b, "example.com", "https://example.com", total, pass
    Check b, "www.google.co.jp", "https://www.google.co.jp", total, pass
    Check b, "192.168.0.1", "https://192.168.0.1", total, pass
    Check b, "localhost:8080", "localhost:8080", total, pass          ' only letters before the colon -> treated as a scheme and passed through (Edge adds http://)

    ' --- (3) has whitespace or no dot -> Google search ---
    Check b, "hello world", "https://www.google.com/search?q=hello%20world", total, pass
    Check b, "localhost", "https://www.google.com/search?q=localhost", total, pass
    Check b, "excel vba webview2", "https://www.google.com/search?q=excel%20vba%20webview2", total, pass

    ' --- (4) encoding of Japanese and symbols (equivalent to encodeURIComponent) ---
    Check b, (ChrW(&H65E5&) & ChrW(&H672C&) & " " & ChrW(&H5929&) & ChrW(&H6C17&)), "https://www.google.com/search?q=%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97", total, pass
    Check b, ("C++ " & ChrW(&H5165&) & ChrW(&H9580&)), "https://www.google.com/search?q=C%2B%2B%20%E5%85%A5%E9%96%80", total, pass
    Check b, "100% pure", "https://www.google.com/search?q=100%25%20pure", total, pass
    Check b, "a.b-c_d~e f", "https://www.google.com/search?q=a.b-c_d~e%20f", total, pass  ' search because of the space; -_.~ pass through

    ' --- (5) trimming leading/trailing whitespace ---
    Check b, "  example.com  ", "https://example.com", total, pass
    Check b, "", "", total, pass   ' empty input stays empty (ignored on the NavigateActive side)

    Debug.Print "==== Result: " & pass & " / " & total & _
                IIf(pass = total, "  [ALL OK]", "  [!! FAILED !!]") & " ===="

    Set b = Nothing
End Sub

' Verification helper: compares with the expected value and Debug.Prints (dedicated to Test_9_14_SearchFallback)
Private Sub Check(ByRef b As Wv2Browser, ByVal inp As String, _
                  ByVal expected As String, ByRef total As Long, ByRef pass As Long)
    total = total + 1
    Dim got As String
    got = b.Debug_NormalizeUrl(inp)
    If got = expected Then
        pass = pass + 1
        Debug.Print "[OK] """ & inp & """ -> " & got
    Else
        Debug.Print "[NG] """ & inp & """" & vbCrLf & _
                    "      got     = " & got & vbCrLf & _
                    "      expected= " & expected
    End If
End Sub

' ============================================================
' Test_9_15_Reorder  (Stage 9.15 logic verification of tab reordering)
'
'   *Just type  Test_9_15_Reorder  in the Immediate window*
'
'   WebView2 is not started. Creates a Wv2Browser, seeds tagged dummy tabs ("A","B","C",...) with Debug_SeedDummyTabs,
'   then calls MoveTab with various from/to values and checks the tab order after reordering (Debug_TabOrderTitles) and
'   how the active index follows (Debug_ActiveIndex)
'   against expected values.
'   A pass/fail summary at the end.
'
'   *Checking that the active tab is followed* Right after the seed, activeIndex=1 (= tab "A"). Even as MoveTab
'   is repeated, we confirm that activeIndex follows "the position where A is"
'   (because MoveTab looks up the active tab backwards by object reference).
'
'   NOTE: Dummy Panes have no View, so ActivateTab inside MoveTab
'     writes a PutIsVisible failure log, but this does not affect the order or the activeIndex following.
' ============================================================
Public Sub Test_9_15_Reorder()
    Dim total As Long, pass As Long
    total = 0: pass = 0
    Debug.Print "==== Test_9_15_Reorder start ===="

    Dim b As Wv2Browser

    ' --- Case 1: in [A,B,C,D] move A(1) to the end (4) -> [B,C,D,A], active follows A = 4 ---
    Set b = New Wv2Browser
    b.Debug_SeedDummyTabs 4
    b.MoveTab 1, 4
    CheckOrder b, "case1 A->end", "B|C|D|A", 4, total, pass

    ' --- Case 2: in [A,B,C,D] move D(4) to the front (1) -> [D,A,B,C], active (A) is 2 ---
    Set b = New Wv2Browser
    b.Debug_SeedDummyTabs 4
    b.MoveTab 4, 1
    CheckOrder b, "case2 D->front", "D|A|B|C", 2, total, pass

    ' --- Case 3: in [A,B,C,D] move B(2) behind C (3) -> [A,C,B,D], active (A) is 1 ---
    Set b = New Wv2Browser
    b.Debug_SeedDummyTabs 4
    b.MoveTab 2, 3
    CheckOrder b, "case3 B->3", "A|C|B|D", 1, total, pass

    ' --- Case 4: in [A,B,C,D] move C(3) before B (2) -> [A,C,B,D], active (A) is 1 ---
    Set b = New Wv2Browser
    b.Debug_SeedDummyTabs 4
    b.MoveTab 3, 2
    CheckOrder b, "case4 C->2", "A|C|B|D", 1, total, pass

    ' --- Case 5: from==to is no change (treated as success) -> [A,B,C,D], active 1 ---
    Set b = New Wv2Browser
    b.Debug_SeedDummyTabs 4
    b.MoveTab 2, 2
    CheckOrder b, "case5 same", "A|B|C|D", 1, total, pass

    ' --- Case 6: out-of-range from=0 is no change -> [A,B,C], active 1 ---
    Set b = New Wv2Browser
    b.Debug_SeedDummyTabs 3
    b.MoveTab 0, 2
    CheckOrder b, "case6 from out of range", "A|B|C", 1, total, pass

    ' --- Case 7: out-of-range to=99 is no change -> [A,B,C], active 1 ---
    Set b = New Wv2Browser
    b.Debug_SeedDummyTabs 3
    b.MoveTab 2, 99
    CheckOrder b, "case7 to out of range", "A|B|C", 1, total, pass

    ' --- Case 8: move the active tab itself. [A,B,C] active=1(A), move A(1) to 3 ---
    '   -> [B,C,A], active follows A to 3 ---
    Set b = New Wv2Browser
    b.Debug_SeedDummyTabs 3
    b.MoveTab 1, 3
    CheckOrder b, "case8 move active", "B|C|A", 3, total, pass

    ' --- Case 9: swap 2 tabs [A,B] -> move B(2) to 1 -> [B,A], active(A)=2 ---
    Set b = New Wv2Browser
    b.Debug_SeedDummyTabs 2
    b.MoveTab 2, 1
    CheckOrder b, "case9 swap 2 tabs", "B|A", 2, total, pass

    ' --- Case 10: 5 tabs, multi-step. in [A,B,C,D,E] move C(3) to 1 -> [C,A,B,D,E], active(A)=2 ---
    Set b = New Wv2Browser
    b.Debug_SeedDummyTabs 5
    b.MoveTab 3, 1
    CheckOrder b, "case10 C->front", "C|A|B|D|E", 2, total, pass

    Debug.Print "==== Result: " & pass & " / " & total & _
                IIf(pass = total, "  [ALL OK]", "  [!! FAILED !!]")
End Sub


' ------------------------------------------------------------
' CheckOrder - checks the tab order and activeIndex after MoveTab against expected values (Stage 9.15)
'   b            : target Browser
'   label        : case name (for log display)
'   expectOrder  : expected tab order (tags joined with "|", like "B|C|D|A")
'   expectActive : expected activeIndex
' ------------------------------------------------------------
Private Sub CheckOrder(ByRef b As Wv2Browser, ByVal label As String, _
                       ByVal expectOrder As String, ByVal expectActive As Long, _
                       ByRef total As Long, ByRef pass As Long)
    total = total + 1
    Dim gotOrder As String
    Dim gotActive As Long
    gotOrder = b.Debug_TabOrderTitles()
    gotActive = b.Debug_ActiveIndex()
    If gotOrder = expectOrder And gotActive = expectActive Then
        pass = pass + 1
        Debug.Print "[OK] " & label & " -> order=" & gotOrder & " active=" & gotActive
    Else
        Debug.Print "[NG] " & label & vbCrLf & _
                    "      got     : order=" & gotOrder & " active=" & gotActive & vbCrLf & _
                    "      expected: order=" & expectOrder & " active=" & expectActive
    End If
End Sub

' ============================================================
' Test_9_16_SearchEngine  (Stage 9.16 logic verification of search engine switching)
'
'   *Just type  Test_9_16_SearchEngine  in the Immediate window*
'
'   WebView2 is not started. Creates a Wv2Browser and, while switching the search template with SearchEngine (Property Let),
'   checks with Check that search-term input is correctly converted to each engine's
'   search URL (the 3rd branch of Debug_NormalizeUrl).
'   Also confirms the current template with Debug_SearchEngine. A pass/fail summary at the end.
'
'   *Axes being checked*
'     (1) The default (no setting) stays Google (same behavior as 9.14).
'     (2) Switching to Bing / DuckDuckGo makes the search URL follow that engine.
'     (3) Letting an empty string falls back to the default (Google) (empty guard).
'     (4) Japanese encoding (UTF-8 percent-encoding) is still correct after switching engines.
'     (5) URL / scheme input falls into branches (1)(2), so it is unaffected by the engine setting.
' ============================================================
Public Sub Test_9_16_SearchEngine()
    Dim b As Wv2Browser
    Set b = New Wv2Browser
    ' NOTE: Init is not needed (Debug_NormalizeUrl / SearchEngine only touch NormalizeUrl and
    '    m_searchTemplate, and never touch tabs or the Environment).

    Dim total As Long, pass As Long
    total = 0: pass = 0

    Debug.Print "==== Test_9_16_SearchEngine start ===="

    ' --- (1) the default is Google (set in Class_Initialize) ---
    Debug.Print "  Current engine = " & b.Debug_SearchEngine()
    Check b, "hello world", "https://www.google.com/search?q=hello%20world", total, pass
    Check b, (ChrW(&H65E5&) & ChrW(&H672C&) & " " & ChrW(&H5929&) & ChrW(&H6C17&)), "https://www.google.com/search?q=%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97", total, pass

    ' --- (2) switch to Bing ---
    b.SearchEngine = "https://www.bing.com/search?q="
    Debug.Print "  Current engine = " & b.Debug_SearchEngine()
    Check b, "hello world", "https://www.bing.com/search?q=hello%20world", total, pass
    Check b, "excel vba", "https://www.bing.com/search?q=excel%20vba", total, pass
    ' Japanese encoding must still be correct after the switch
    Check b, (ChrW(&H65E5&) & ChrW(&H672C&) & " " & ChrW(&H5929&) & ChrW(&H6C17&)), "https://www.bing.com/search?q=%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97", total, pass
    ' Debug_SearchEngine must return the current value
    total = total + 1
    If b.Debug_SearchEngine() = "https://www.bing.com/search?q=" Then
        pass = pass + 1
        Debug.Print "[OK] Debug_SearchEngine = Bing template"
    Else
        Debug.Print "[NG] Debug_SearchEngine = " & b.Debug_SearchEngine()
    End If

    ' --- (3) switch to DuckDuckGo ---
    b.SearchEngine = "https://duckduckgo.com/?q="
    Check b, "privacy", "https://duckduckgo.com/?q=privacy", total, pass

    ' --- (4) Let an empty string -> fall back to the default (Google) (empty guard) ---
    b.SearchEngine = ""
    Debug.Print "  Engine after empty Let = " & b.Debug_SearchEngine()
    Check b, "hello world", "https://www.google.com/search?q=hello%20world", total, pass
    total = total + 1
    If b.Debug_SearchEngine() = "https://www.google.com/search?q=" Then
        pass = pass + 1
        Debug.Print "[OK] empty Let falls back to Google"
    Else
        Debug.Print "[NG] after empty Let = " & b.Debug_SearchEngine()
    End If

    ' --- (5) Let whitespace only -> the empty guard also returns to Google ---
    b.SearchEngine = "https://www.bing.com/search?q="   ' switch to Bing once
    b.SearchEngine = "   "                              ' whitespace only -> should return to Google
    total = total + 1
    If b.Debug_SearchEngine() = "https://www.google.com/search?q=" Then
        pass = pass + 1
        Debug.Print "[OK] whitespace-only Let falls back to Google"
    Else
        Debug.Print "[NG] after whitespace-only Let = " & b.Debug_SearchEngine()
    End If

    ' --- (6) the engine setting does not affect URL / scheme input (branches 1 and 2) ---
    b.SearchEngine = "https://duckduckgo.com/?q="   ' set search to DDG
    Check b, "example.com", "https://example.com", total, pass          ' has a dot -> URL (no effect)
    Check b, "https://example.com", "https://example.com", total, pass  ' has a scheme -> as is (no effect)
    Check b, "about:blank", "about:blank", total, pass                  ' has a scheme -> as is (no effect)

    Debug.Print "==== Result: " & pass & " / " & total & _
                IIf(pass = total, "  [ALL OK]", "  [!! FAILED !!]")
End Sub


' ============================================================
' Test_9_16_SearchEngine_Help  (Stage 9.16 help for the verification procedure)
'
'   *Type  Test_9_16_SearchEngine_Help  in the Immediate window to show the procedure*
' ============================================================
Public Sub Test_9_16_SearchEngine_Help()
    Debug.Print "==== Stage 9.16 search engine switching: verification procedure ===="
    Debug.Print ""
    Debug.Print "[A. Logic verification (no WebView2 start needed)]"
    Debug.Print "  Just type the following in the Immediate window:"
    Debug.Print "    Test_9_16_SearchEngine"
    Debug.Print "  -> Checks default Google / Bing / DuckDuckGo switching, the empty guard, Japanese encoding,"
    Debug.Print "     and no effect on URL input, all in one batch. Pass if [ALL OK] appears at the end."
    Debug.Print ""
    Debug.Print "[B. Real-machine GUI verification (WebView2 started)]"
    Debug.Print "  1) Start normally with StartWebView2_Full."
    Debug.Print ("  2) Type  " & ChrW(&H65E5&) & ChrW(&H672C&) & " " & ChrW(&H5929&) & ChrW(&H6C17&) & "  in the address bar and press Enter.")
    Debug.Print "     -> Google search results should appear with the default unchanged."
    Debug.Print "  3) Switch the search engine to Bing in the Immediate window:"
    Debug.Print "       Set gBrowser = <reference to the running Wv2Browser>   ' reference exposed by the startup Sub"
    Debug.Print "       gBrowser.SearchEngine = ""https://www.bing.com/search?q="""
    Debug.Print "     NOTE: how to obtain the running Browser reference depends on the startup Sub's implementation."
    Debug.Print "       If you do not hold a reference, you can also check by rewriting the default of BuildSearchUrl (DEFAULT_"
    Debug.Print "       SEARCH_TEMPLATE) and restarting."
    Debug.Print ("  4) Type  " & ChrW(&H65E5&) & ChrW(&H672C&) & " " & ChrW(&H5929&) & ChrW(&H6C17&) & "  again and press Enter.")
    Debug.Print "     -> If Bing search results appear this time, switching succeeded on the real machine too."
    Debug.Print ""
    Debug.Print "  NOTE: if there is no way to get a global reference in step 3) of B, the logic verification in A"
    Debug.Print "     guarantees that switching is correct. On the real machine, just checking the default Google behavior is enough."
End Sub

' ============================================================
' Test_9_17_SearchPresets  (Stage 9.17 verifying search engine preset resolution)
'
'   *Just type  Test_9_17_SearchPresets  in the Immediate window*
'
'   WebView2 is not started. Lets the preset templates (same strings as SE_GOOGLE etc.) that UserForm1.UseSearchEngine resolves names to
'   directly into Wv2Browser, and checks with Check whether search-term input is converted to that engine's search URL
'   (the 3rd branch of Debug_NormalizeUrl)
'   and whether Debug_SearchEngine returns the set value.
'
'   *What this guarantees* That each preset's URL prefix template is correct and that simply
'   Letting it generates a correct search URL. The name resolution of UseSearchEngine
'   ("bing" -> SE_BING) and the not-started guard are confirmed on the real machine (via CurrentBrowser).
'
'   NOTE: the preset strings duplicate UserForm1's Private Const, but this is an intentional copy so that
'   verification can run without starting the UserForm (the same as hard-coding expected values on the test side).
' ============================================================
Public Sub Test_9_17_SearchPresets()
    Dim b As Wv2Browser
    Set b = New Wv2Browser
    ' NOTE: Init is not needed (same as Test_9_16. Only touches NormalizeUrl / m_searchTemplate)

    Dim total As Long, pass As Long
    total = 0: pass = 0

    Debug.Print "==== Test_9_17_SearchPresets start ===="

    ' --- google preset ---
    b.SearchEngine = "https://www.google.com/search?q="
    Debug.Print "  google  = " & b.Debug_SearchEngine()
    Check b, "hello world", "https://www.google.com/search?q=hello%20world", total, pass
    Check b, (ChrW(&H65E5&) & ChrW(&H672C&) & " " & ChrW(&H5929&) & ChrW(&H6C17&)), "https://www.google.com/search?q=%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97", total, pass

    ' --- bing preset ---
    b.SearchEngine = "https://www.bing.com/search?q="
    Debug.Print "  bing    = " & b.Debug_SearchEngine()
    Check b, "excel vba", "https://www.bing.com/search?q=excel%20vba", total, pass
    Check b, (ChrW(&H65E5&) & ChrW(&H672C&) & " " & ChrW(&H5929&) & ChrW(&H6C17&)), "https://www.bing.com/search?q=%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97", total, pass

    ' --- duckduckgo preset (must concatenate correctly even when the path is of the /?q= type) ---
    b.SearchEngine = "https://duckduckgo.com/?q="
    Debug.Print "  ddg     = " & b.Debug_SearchEngine()
    Check b, "privacy", "https://duckduckgo.com/?q=privacy", total, pass
    Check b, (ChrW(&H65E5&) & ChrW(&H672C&) & " " & ChrW(&H5929&) & ChrW(&H6C17&)), "https://duckduckgo.com/?q=%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97", total, pass

    ' --- yahoo preset (a ?p= type query key can also be expressed with a prefix template) ---
    b.SearchEngine = "https://search.yahoo.com/search?p="
    Debug.Print "  yahoo   = " & b.Debug_SearchEngine()
    Check b, "excel vba", "https://search.yahoo.com/search?p=excel%20vba", total, pass

    ' --- switching presets does not affect URL / scheme input (branches 1 and 2) ---
    b.SearchEngine = "https://search.yahoo.com/search?p="   ' set search to Yahoo
    Check b, "example.com", "https://example.com", total, pass          ' has a dot -> URL
    Check b, "https://example.com", "https://example.com", total, pass  ' has a scheme -> as is

    Debug.Print "==== Result: " & pass & " / " & total & _
                IIf(pass = total, "  [ALL OK]", "  [!! FAILED !!]")
End Sub


' ============================================================
' Test_9_17_SearchPresets_Help  (Stage 9.17 help for the verification procedure)
'
'   *Type  Test_9_17_SearchPresets_Help  in the Immediate window to show the procedure*
' ============================================================
Public Sub Test_9_17_SearchPresets_Help()
    Debug.Print "==== Stage 9.17 search engine switching entry + CurrentBrowser exposure: verification procedure ===="
    Debug.Print ""
    Debug.Print "[A. Logic verification (no WebView2 start needed)]"
    Debug.Print "  Just type the following in the Immediate window:"
    Debug.Print "    Test_9_17_SearchPresets"
    Debug.Print "  -> Checks in one batch that the search URL is generated correctly for each of the google / bing / duckduckgo / yahoo presets,"
    Debug.Print "     that Japanese encoding is preserved, and that URL input is"
    Debug.Print "     unaffected. Pass if [ALL OK] appears at the end."
    Debug.Print ""
    Debug.Print "[B. Real-machine GUI verification (WebView2 started)]  *9.17 made a reference entry point available*"
    Debug.Print "  1) Start normally with StartWebView2_Full."
    Debug.Print ("  2) Type  " & ChrW(&H65E5&) & ChrW(&H672C&) & " " & ChrW(&H5929&) & ChrW(&H6C17&) & "  in the address bar and press Enter.")
    Debug.Print "     -> Google search results should appear with the default unchanged."
    Debug.Print "  3) Switch the search engine to Bing in the Immediate window (either way works):"
    Debug.Print "       Method 1 (preset name):  UserForm1.UseSearchEngine ""bing"""
    Debug.Print "       Method 2 (grab a reference)  :  Set b = UserForm1.CurrentBrowser"
    Debug.Print "                              b.SearchEngine = ""https://www.bing.com/search?q="""
    Debug.Print ("  4) Type  " & ChrW(&H65E5&) & ChrW(&H672C&) & " " & ChrW(&H5929&) & ChrW(&H6C17&) & "  again and press Enter.")
    Debug.Print "     -> If Bing search results appear this time, switching succeeded on the real machine too."
    Debug.Print "  5) Name resolution check:  UserForm1.UseSearchEngine ""duckduckgo""  -> DDG,"
    Debug.Print "       UserForm1.UseSearchEngine ""nonsense""  -> a warning appears and it returns to Google."
    Debug.Print "  6) Not-started guard:  before running StartWebView2_Full, run"
    Debug.Print "       UserForm1.UseSearchEngine ""bing""  -> a notice 'Browser is not started' appears."
End Sub


' ============================================================
' Test_9_18_BridgeLogic  (Stage 9.18 pure-logic verification of Wv2SettingsBridge)
'
'   *Just type  Test_9_18_BridgeLogic  in the Immediate window*
'
'   WebView2 is not started. Creates a Wv2SettingsBridge, binds a real Wv2Browser with
'   BindBrowser, then directly calls the same methods the settings screen's JS calls
'   (SetEngine / GetEngine / PreviewUrl / PreviewUrlFor / EngineList)
'   and checks name <-> template resolution and preview URL generation.
'
'   *What this guarantees* The settings screen's JS only calls these methods via hostObjects.sync.Settings,
'     so if the methods' return values are correct here, the remaining uncertainty narrows to just
'     "whether the hostObjects route works on the real machine (order A)".
'     That real-machine connectivity is confirmed with the procedure in Test_9_18_BridgeLogic_Help.
' ============================================================
Public Sub Test_9_18_BridgeLogic()
    Dim total As Long, pass As Long
    total = 0: pass = 0
    Debug.Print "==== Test_9_18_BridgeLogic (Wv2SettingsBridge pure logic) ===="

    Dim b As Wv2Browser
    Set b = New Wv2Browser          ' initialized with the default Google

    Dim br As Wv2SettingsBridge
    Set br = New Wv2SettingsBridge
    br.BindBrowser b

    ' --- initial state: the default is google ---
    CheckEq "GetEngine initial", br.GetEngine(), "google", total, pass

    ' --- return value of SetEngine and GetEngine following it ---
    CheckEq "SetEngine bing return value", br.SetEngine("bing"), "bing", total, pass
    CheckEq "GetEngine after bing", br.GetEngine(), "bing", total, pass

    CheckEq "SetEngine DDG alias ddg", br.SetEngine("ddg"), "duckduckgo", total, pass
    CheckEq "GetEngine after ddg", br.GetEngine(), "duckduckgo", total, pass

    ' --- normalization of case and leading/trailing whitespace ---
    CheckEq "SetEngine '  YAHOO '", br.SetEngine("  YAHOO "), "yahoo", total, pass

    ' --- an unknown name falls back to Google ---
    CheckEq "SetEngine unknown name", br.SetEngine("nonsense"), "google", total, pass
    CheckEq "GetEngine after unknown name", br.GetEngine(), "google", total, pass

    ' --- PreviewUrl is generated with "the current engine (google)", with Japanese encoding ---
    CheckEq "PreviewUrl (google)", br.PreviewUrl((ChrW(&H65E5&) & ChrW(&H672C&) & " " & ChrW(&H5929&) & ChrW(&H6C17&))), _
            "https://www.google.com/search?q=%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97", total, pass

    ' --- PreviewUrlFor is generated with "the specified engine" (does not change the current engine) ---
    CheckEq "PreviewUrlFor bing", br.PreviewUrlFor("bing", "excel vba"), _
            "https://www.bing.com/search?q=excel%20vba", total, pass
    CheckEq "PreviewUrlFor yahoo", br.PreviewUrlFor("yahoo", "privacy"), _
            "https://search.yahoo.com/search?p=privacy", total, pass
    ' The current engine stays google even after PreviewUrlFor (no side effects)
    CheckEq "current after PreviewUrlFor", br.GetEngine(), "google", total, pass

    ' --- PreviewUrl does not turn URL input into a search (branches 1 and 2 pass through) ---
    CheckEq "PreviewUrl URL input", br.PreviewUrl("example.com"), "https://example.com", total, pass

    ' --- EngineList returns JSON containing the 4 engines ---
    Dim js As String
    js = br.EngineList()
    CheckBool "EngineList has google", (InStr(js, """id"":""google""") > 0), total, pass
    CheckBool "EngineList has bing", (InStr(js, """id"":""bing""") > 0), total, pass
    CheckBool "EngineList has duckduckgo", (InStr(js, """id"":""duckduckgo""") > 0), total, pass
    CheckBool "EngineList has yahoo", (InStr(js, """id"":""yahoo""") > 0), total, pass

    Debug.Print "==== Result: " & pass & " / " & total & _
                IIf(pass = total, "  [ALL OK]", "  [!! FAILED !!]")
End Sub


' --- Stage 9.18: string equality check (generic) ---
Private Sub CheckEq(ByVal label As String, ByVal got As String, _
                    ByVal expected As String, ByRef total As Long, ByRef pass As Long)
    total = total + 1
    If got = expected Then
        pass = pass + 1
        Debug.Print "[OK] " & label & " -> " & got
    Else
        Debug.Print "[NG] " & label & vbCrLf & _
                    "      got     = " & got & vbCrLf & _
                    "      expected= " & expected
    End If
End Sub


' --- Stage 9.18: boolean check (generic) ---
Private Sub CheckBool(ByVal label As String, ByVal cond As Boolean, _
                      ByRef total As Long, ByRef pass As Long)
    total = total + 1
    If cond Then
        pass = pass + 1
        Debug.Print "[OK] " & label
    Else
        Debug.Print "[NG] " & label & "  (condition is False)"
    End If
End Sub


' ============================================================
' Test_9_18_BridgeLogic_Help  (Stage 9.18 help for the verification procedure)
'
'   *Type  Test_9_18_BridgeLogic_Help  in the Immediate window to show the procedure*
' ============================================================
Public Sub Test_9_18_BridgeLogic_Help()
    Debug.Print "==== Stage 9.18 settings tab + AddHostObjectToScript: verification procedure ===="
    Debug.Print ""
    Debug.Print "[A. Bridge pure-logic verification (no WebView2 start needed)]"
    Debug.Print "  Just type the following in the Immediate window:"
    Debug.Print "    Test_9_18_BridgeLogic"
    Debug.Print "  -> Directly calls SetEngine/GetEngine/PreviewUrl(For)/EngineList of Wv2SettingsBridge"
    Debug.Print "     and checks name resolution, preview URL and absence of side effects in one batch."
    Debug.Print "     If [ALL OK] appears at the end, the only remaining uncertainty is hostObjects connectivity on the real machine."
    Debug.Print ""
    Debug.Print "[B. Real-machine hostObjects connectivity + settings screen (WebView2 started)]  *the main point of 9.18*"
    Debug.Print "  1) Start normally with StartWebView2_Full."
    Debug.Print "  2) Click the [gear] button at the right end of the tab bar."
    Debug.Print "     -> A new tab opens; if the search-engine selection card screen appears, settings tab creation is OK."
    Debug.Print "  3) *Real-machine verdict for order A* 4 cards (Google/Bing/DuckDuckGo/Yahoo) are drawn, and"
    Debug.Print "     the card of the current engine (default Google) is highlighted as selected in blue."
    Debug.Print "     -> If this appears, 'attach before Navigate makes hostObjects available"
    Debug.Print "        when the JS loads (order A)' holds on the real machine. If the card frame shows 'Host object not connected',"
    Debug.Print "        order A failed -> consider switching to order B (ready handshake)."
    Debug.Print "  4) Hover the mouse over a card -> the Preview field below should show"
    Debug.Print "     'Japanese search -> generated URL' live (synchronous call of PreviewUrlFor)."
    Debug.Print "  5) Click the Bing card -> the Bing card should change to the selected highlight."
    Debug.Print ("  6) Move to a normal tab other than the settings tab and type  " & ChrW(&H65E5&) & ChrW(&H672C&) & " " & ChrW(&H5929&) & ChrW(&H6C17&) & "  in the address bar.")
    Debug.Print "     -> If Bing search results appear, the switch made on the settings screen is reflected in production."
    Debug.Print "  7) The settings tab can be closed normally with x (it behaves as a normal tab)."
    Debug.Print ""
    Debug.Print "  NOTE: the settings tab opens even when you press [gear] with no normal tab (even as the last 1 tab)."
End Sub


' ============================================================
' Test_9_19_SettingsTabDedup_Help  (Stage 9.19 verification procedure for preventing duplicate settings tabs)
'
'   *Type  Test_9_19_SettingsTabDedup_Help  in the Immediate window to show the procedure*
'
'   Key points of the logic (covered by code review):
'     - At the top of OpenSettingsTab, scan m_tabs, and if there is a Pane with IsSettings=True,
'       ActivateTab it and Exit Sub (do not create a new one).
'     - On creation, set pane.IsSettings = True.
'     - A closed settings tab is removed from m_tabs so the scan does not find it, and the next [gear]
'       can open a new one (the aliveness check is integrated with the scan).
' ============================================================
Public Sub Test_9_19_SettingsTabDedup_Help()
    Debug.Print "==== Stage 9.19 preventing duplicate settings tabs: verification procedure (WebView2 started) ===="
    Debug.Print ""
    Debug.Print "  1) Start normally with StartWebView2_Full."
    Debug.Print "  2) Press the [gear] at the right end of the tab bar once."
    Debug.Print "     -> One settings tab should open and become active."
    Debug.Print "  3) *The main point of duplicate prevention* Next, mash [gear] 2-3 more times."
    Debug.Print "     -> No more settings tabs should be added; the existing settings tab should merely become active."
    Debug.Print "       (The Immediate window shows 'activated existing settings tab (index N)')"
    Debug.Print "  4) Click another normal tab to leave the settings tab."
    Debug.Print "     -> Pressing [gear] in that state switches to the existing settings tab (no new one is opened)."
    Debug.Print "  5) *Close, then reopen* Press x on the settings tab to close it."
    Debug.Print "     -> Pressing [gear] once more after that should open 1 new settings tab."
    Debug.Print "       (Once closed it is removed from m_tabs, so the scan does not find it and it can be reopened)"
    Debug.Print "  6) The settings screen's card operations (engine selection, hover preview, reflecting the switch in production)"
    Debug.Print "     should work the same as in 9.18 (confirms duplicate prevention did not break the settings feature)."
    Debug.Print ""
    Debug.Print "  NOTE: expected behavior in summary: there is always at most 1 settings tab. [gear] means 'open if absent / move to it if present'."
End Sub


' ============================================================
' Test_9_20_Persistence  (Stage 9.20 pure-logic verification of search engine persistence)
'
'   *Just type  Test_9_20_Persistence  in the Immediate window*
'
'   Directly calls Wv2SettingsBridge's save/load. WebView2 is not started.
'   Because it touches the real file, it saves the current stored value aside at the top and, at the end,
'   writes it back so as not to break the user's real settings.
' ============================================================
Public Sub Test_9_20_Persistence()
    Dim total As Long, pass As Long
    total = 0: pass = 0
    Debug.Print "==== Test_9_20_Persistence (persistence, pure logic) ===="

    Dim br As Wv2SettingsBridge
    Set br = New Wv2SettingsBridge

    ' --- save the current stored value aside (restored after the test) ---
    '   *9.29: save the template as a pair, not only the engine* If the original was custom,
    '     writing back only the engine line would erase the template line.
    Dim savedBefore As String, savedTplBefore As String
    savedBefore = br.LoadEngineName()
    savedTplBefore = br.LoadCustomTemplate()
    Debug.Print "  (saved aside) current stored value = '" & savedBefore & "'" & _
                IIf(Len(savedTplBefore) > 0, " / template = '" & savedTplBefore & "'", "")

    ' --- save -> load round trip (bing) ---
    br.Debug_SaveEngineName "bing"
    CheckEq "save bing -> LoadEngineName", br.LoadEngineName(), "bing", total, pass
    CheckEq "save bing -> LoadEngineTemplate", br.LoadEngineTemplate(), _
            "https://www.bing.com/search?q=", total, pass

    ' --- overwrite with another engine (duckduckgo) ---
    br.Debug_SaveEngineName "duckduckgo"
    CheckEq "overwrite ddg -> LoadEngineName", br.LoadEngineName(), "duckduckgo", total, pass
    CheckEq "overwrite ddg -> LoadEngineTemplate", br.LoadEngineTemplate(), _
            "https://duckduckgo.com/?q=", total, pass

    ' --- after saving an unknown name, LoadEngineTemplate falls back to "" ---
    '   (LoadEngineName returns the raw value, but template resolution makes it "")
    br.Debug_SaveEngineName "nonsense"
    CheckEq "unknown name -> LoadEngineName (raw value)", br.LoadEngineName(), "nonsense", total, pass
    CheckEq "unknown name -> LoadEngineTemplate (empty)", br.LoadEngineTemplate(), "", total, pass

    ' --- write back the saved-aside value (restoring the user settings) ---
    If Len(savedBefore) > 0 Then
        br.Debug_SaveSettings savedBefore, savedTplBefore
        Debug.Print "  (restored) stored value put back to '" & savedBefore & "'."
    Else
        Debug.Print "  (restored) The stored value was originally empty, so the test's final value"
        Debug.Print "         (nonsense) remains in the file, but LoadEngineTemplate falls back to '', so"
        Debug.Print "         the next start begins with the default (Google) and no harm is done. If it bothers you,"
        Debug.Print "         choosing any engine on the settings screen will overwrite it."
    End If

    Debug.Print "==== Result: " & pass & " / " & total & _
                IIf(pass = total, "  [ALL OK]", "  [!! FAILED !!]")
End Sub


' ============================================================
' Test_9_20_Persistence_Help  (Stage 9.20 real-machine verification procedure for persistence)
'
'   *Type  Test_9_20_Persistence_Help  in the Immediate window to show the procedure*
'
'   Restoring involves running Wv2Browser.Class_Initialize, so it is confirmed by a
'   real-machine procedure, not pure logic (the key is fully quitting Excel -> restarting).
' ============================================================
Public Sub Test_9_20_Persistence_Help()
    Debug.Print "==== Stage 9.20 persistence of search engine selection: real-machine verification procedure ===="
    Debug.Print ""
    Debug.Print "  1) Start normally with StartWebView2_Full."
    Debug.Print "  2) Press [gear] to open the settings tab and click the Bing card."
    Debug.Print "     -> The following 2 lines should appear in the Immediate window:"
    Debug.Print "         Wv2SettingsBridge.SetEngine: 'bing' -> 'bing'"
    Debug.Print "         Wv2SettingsBridge.SaveSettingsFile: saved 'bing' -> ...\settings.txt"
    Debug.Print ("  3) In another normal tab type   " & ChrW(&H65E5&) & ChrW(&H672C&) & " " & ChrW(&H5929&) & ChrW(&H6C17&) & "   in the address bar and confirm that Bing search results")
    Debug.Print "     appear (up to here it is the same as 9.18)."
    Debug.Print ""
    Debug.Print "  --- *From here is the main point of 9.20* ---"
    Debug.Print "  4) Fully quit Excel (not just close the workbook, but Excel itself)."
    Debug.Print "  5) Restart Excel, open the workbook and run StartWebView2_Full."
    Debug.Print "     -> The following 1 line should appear in the Immediate window (proof of restoring):"
    Debug.Print "         Wv2Browser.Class_Initialize: restored saved engine = https://www.bing.com/search?q="
    Debug.Print ("  6) Without opening the settings tab, immediately type   " & ChrW(&H65E5&) & ChrW(&H672C&) & " " & ChrW(&H5929&) & ChrW(&H6C17&) & "   in a normal tab.")
    Debug.Print "     -> If Bing search results from the previously chosen engine appear right away, persistence succeeded."
    Debug.Print "  7) (Optional) Opening the settings tab, also confirm that the Bing card is highlighted as selected."
    Debug.Print ""
    Debug.Print "  NOTE: storage location: %APPDATA%\Wv2Browser\settings.txt (the content is the 1 line  engine=bing )."
    Debug.Print "     If you want to see the contents in Notepad or by hand, you can open it in Explorer."
    Debug.Print "  NOTE: you can also confirm that deleting the file by hand and restarting returns to the default (Google)."
End Sub


' ============================================================
' Test_9_21_HostActivate_Help  (Stage 9.21 real-machine procedure for moving activate to hostObjects)
'
'   *Type  Test_9_21_HostActivate_Help  in the Immediate window to show the procedure*
'
'   Check on the real machine whether spec fact 18 also holds for View_Navigate to a real URL (appassets.tabbar)
'   (= whether attach before Navigate lets the JS grab hostObjects from the very first time).
' ============================================================
Public Sub Test_9_21_HostActivate_Help()
    Debug.Print "==== Stage 9.21 moving activate to hostObjects: real-machine verification procedure ===="
    Debug.Print ""
    Debug.Print "  1) Start normally with StartWebView2_Full."
    Debug.Print "     -> The following 1 line should appear in the startup log (proof the attach succeeded):"
    Debug.Print "         Wv2TabBar.Init: TabBar bridge attached OK (activate prefers hostObjects)"
    Debug.Print "       NOTE: if 'AddHostObjectToScript(TabBar) failed' appears instead,"
    Debug.Print "         the attach did not go through, and activate falls back to postMessage from then on."
    Debug.Print "  2) Increase the tabs to 3-4 with the + button."
    Debug.Print "  3) Click the body of another tab (one that is not active now) to switch."
    Debug.Print ""
    Debug.Print "  --- *This is the 9.21 verdict* ---"
    Debug.Print "  4) If the following 1 line appears in the Immediate window on every click, the hostObjects route holds:"
    Debug.Print "         Wv2TabBar.HostActivate: index=N (via host)"
    Debug.Print "     and confirm that the former postMessage-route log"
    Debug.Print "         Wv2TabBar.OnPaneWebMessage: msg={""cmd"":""activate"",...}"
    Debug.Print "     does not appear for activate (for close/newtab/reorder/settings"
    Debug.Print "     OnPaneWebMessage still appears as before = those remain postMessage)."
    Debug.Print "  5) On every switch, confirm that the active highlight and the nav bar (URL box, back/forward) follow"
    Debug.Print "     that tab (the whole-state sync is unchanged, so it should feel the same)."
    Debug.Print ""
    Debug.Print "  --- Regression check (9.18/9.19/9.20 must be unaffected) ---"
    Debug.Print "  6) Closing a tab with x / opening with + / reordering by drag all work as before."
    Debug.Print "  7) The settings tab opens with [gear], changing the engine is reflected in search, and"
    Debug.Print "     confirm that the engine carries over after restarting Excel (9.20 persistence)."
    Debug.Print ""
    Debug.Print "  NOTE: if you can open the browser's DevTools console, on a tab click"
    Debug.Print "     [host] activate N appears ([postmsg] activate N when falling back to postMessage)."
    Debug.Print "     The VBA-side log alone also tells whether it went via host or via postMessage."
End Sub


' ============================================================
' Test_9_22_HostCloseReorder_Help  (Stage 9.22 real-machine procedure for moving close/reorder to hostObjects)
'
'   *Type  Test_9_22_HostCloseReorder_Help  in the Immediate window to show the procedure*
'
'   Following activate, close / reorder were also moved to direct hostObjects calls. Check on the real machine that spec fact 19
'   (attach before Navigate) reaches close/reorder
'   even with a real URL (appassets.tabbar), and that activate/ready/newtab/settings are intact.
' ============================================================
Public Sub Test_9_22_HostCloseReorder_Help()
    Debug.Print "==== Stage 9.22 moving close / reorder to hostObjects: real-machine verification procedure ===="
    Debug.Print ""
    Debug.Print "  1) Start normally with StartWebView2_Full."
    Debug.Print "     -> The same attach-success line as 9.21 should appear in the startup log:"
    Debug.Print "         Wv2TabBar.Init: TabBar bridge attached OK (activate prefers hostObjects)"
    Debug.Print "       NOTE: the attach is a single registration shared by the whole TabBar (activate/close/reorder)."
    Debug.Print "  2) Increase the tabs to 4-5 with the + button."
    Debug.Print ""
    Debug.Print "  --- *This is the 9.22 verdict (close)* ---"
    Debug.Print "  3) Click the x of any tab to close it. If the following 1 line appears, via-host holds:"
    Debug.Print "         Wv2TabBar.HostClose: index=N (via host)"
    Debug.Print "     and confirm that the former postMessage-route log"
    Debug.Print "         Wv2TabBar.OnPaneWebMessage: msg={""cmd"":""close"",...}"
    Debug.Print "     does not appear for close. Also confirm that the active correction after closing (the highlight moving to a remaining"
    Debug.Print "     tab) and the nav bar following are as before."
    Debug.Print ""
    Debug.Print "  --- *This is the 9.22 verdict (reorder)* ---"
    Debug.Print "  4) Drag a tab to reorder. If the following 1 line appears, via-host holds:"
    Debug.Print "         Wv2TabBar.HostReorder: from=A to=B (via host)"
    Debug.Print "     and confirm that the former postMessage-route log"
    Debug.Print "         Wv2TabBar.OnPaneWebMessage: msg={""cmd"":""reorder"",...}"
    Debug.Print "     does not appear for reorder. The tab bar should be redrawn in the reordered order, and"
    Debug.Print "     the grabbed tab should follow with its active highlight kept."
    Debug.Print ""
    Debug.Print "  --- Activate of 9.21 must still be via host ---"
    Debug.Print "  5) Click the body of another tab to switch. As before, the following should appear:"
    Debug.Print "         Wv2TabBar.HostActivate: index=N (via host)"
    Debug.Print ""
    Debug.Print "  --- Regression check (ready/newtab/settings remain postMessage) ---"
    Debug.Print "  6) + for a new tab -> OnPaneWebMessage: msg={""cmd"":""newtab""} appears"
    Debug.Print "     (the former behavior via SafeTimer. Not via host = correct)."
    Debug.Print "  7) [gear] for the settings tab -> OnPaneWebMessage: msg={""cmd"":""settings""} appears and"
    Debug.Print "     the settings tab opens. Engine change -> search reflection -> carry-over across restart (9.20) is also as before."
    Debug.Print "  8) In the initial sync right after startup, OnPaneWebMessage: msg={""cmd"":""ready""} appears"
    Debug.Print "     (ready is a notification, so it remains postMessage = correct)."
    Debug.Print ""
    Debug.Print "  NOTE: expected: activate/close/reorder go via host (HostXxx log),"
    Debug.Print "     ready/newtab/settings go via postMessage (OnPaneWebMessage log)."
    Debug.Print "     This division is the aim of 9.22. If DevTools can be opened, clicking x gives"
    Debug.Print "     [host] close N, and dragging gives [host] reorder A->B."
End Sub


' ============================================================
' Test_9_23_Consolidation_Help  (Stage 9.23 real-machine procedure for Host consolidation of activate/close/reorder)
'
'   *Type  Test_9_23_Consolidation_Help  in the Immediate window to show the procedure*
'
'   9.23 is a cleanup stage that changed only "where the processing lives". The ActivateTab/CloseTab/MoveTab calls and failure logs
'   of activate/close/reorder were gathered into one place in the Host methods, and
'   the Cases of OnPaneWebMessage became thin adapters that delegate to the Host methods when falling back.
'   The functionality should be exactly the same as 9.22, and the aim is to confirm zero regressions.
' ============================================================
Public Sub Test_9_23_Consolidation_Help()
    Debug.Print "==== Stage 9.23 Host consolidation of activate/close/reorder: real-machine verification procedure ===="
    Debug.Print ""
    Debug.Print "  [Premise] 9.23 only changed where the processing lives. The functionality is the same as 9.22."
    Debug.Print "     In normal operation hostObjects holds, so the Case activate/close/reorder"
    Debug.Print "     (= fallback route) is normally not taken. So the main point is confirming zero regressions."
    Debug.Print ""
    Debug.Print "  1) Start normally with StartWebView2_Full."
    Debug.Print "     -> The same attach-success line as 9.21/9.22 should appear in the startup log:"
    Debug.Print "         Wv2TabBar.Init: TabBar bridge attached OK (activate prefers hostObjects)"
    Debug.Print "  2) Increase the tabs to 4-5 with the + button."
    Debug.Print ""
    Debug.Print "  --- *Regression check: the 3 commands still go via host as before* ---"
    Debug.Print "  3) Click the body of another tab -> Wv2TabBar.HostActivate: index=N (via host)"
    Debug.Print "  4) Click x on a tab    -> Wv2TabBar.HostClose: index=N (via host)"
    Debug.Print "  5) Drag a tab to reorder -> Wv2TabBar.HostReorder: from=A to=B (via host)"
    Debug.Print "     All should work as before, and the UX (active correction / redraw / nav bar following) should be"
    Debug.Print "     exactly the same as 9.22. That the OnPaneWebMessage"
    Debug.Print "     activate/close/reorder logs do not appear with these 3 is also the same as 9.22 (because they go via host)."
    Debug.Print ""
    Debug.Print "  --- Regression check: ready/newtab/settings remain postMessage ---"
    Debug.Print "  6) + for a new tab -> OnPaneWebMessage: msg={""cmd"":""newtab""} (via SafeTimer)"
    Debug.Print "  7) [gear] for the settings tab -> OnPaneWebMessage: msg={""cmd"":""settings""} -> the settings tab opens"
    Debug.Print "  8) Initial sync right after startup -> OnPaneWebMessage: msg={""cmd"":""ready""}"
    Debug.Print ""
    Debug.Print "  --- (Optional) Checking the fallback route is alive ---"
    Debug.Print "  9) In a normal environment the fallback cannot be reached. Only if you can temporarily disable"
    Debug.Print "     window.chrome.webview.hostObjects in DevTools,"
    Debug.Print "     if the following 2 lines appear in succession on clicking x, the delegation is alive:"
    Debug.Print "         Wv2TabBar.OnPaneWebMessage: msg={""cmd"":""close"",...}"
    Debug.Print "         -> close(index=N) [postmsg fallback -> HostClose]"
    Debug.Print "         Wv2TabBar.HostClose: index=N (via host)   NOTE: evidence of delegation to Host"
    Debug.Print "     (In an environment where it cannot be disabled, this step may be skipped. The essence of 9.23 is zero regressions.)"
    Debug.Print ""
    Debug.Print "  NOTE: expected: appearance and feel are exactly the same as 9.22. The log division"
    Debug.Print "     (activate/close/reorder=HostXxx, ready/newtab/settings=OnPaneWebMessage)"
    Debug.Print "     is also the same as 9.22. If nothing looks changed, 9.23 is a success."
End Sub



' ============================================================
' Test_9_24_HostNavBar_Help  (Stage 9.24 real-machine procedure for moving NavBar to hostObjects)
'
'   *Type  Test_9_24_HostNavBar_Help  in the Immediate window to show the procedure*
'
'   9.24 is the round that rolled out to NavBar, in the same form, the pattern established with the TabBar (9.21-9.23).
'   The JS is a hybrid (hostObjects first, postMessage fallback), VBA consolidated the
'   processing into 4 Host methods, and the Select Case is a thin adapter.
' ============================================================
Public Sub Test_9_24_HostNavBar_Help()
    Debug.Print "==== Stage 9.24 moving NavBar to hostObjects: real-machine verification procedure ===="
    Debug.Print ""
    Debug.Print "  [Premise] Start normally with StartWebView2_Full."
    Debug.Print "     -> The startup log should show attach lines for both TabBar and NavBar:"
    Debug.Print "         Wv2TabBar.Init: TabBar bridge attached OK (...)"
    Debug.Print "         Wv2NavBar.Init: NavBar bridge attached OK (back/forward/reload/navigate prefer hostObjects)"
    Debug.Print "       *If the NavBar line does not appear, everything from here on goes via postMessage"
    Debug.Print "         (it works, but the aim of 9.24 is not met). Note the hr and report it."
    Debug.Print ""
    Debug.Print "  --- *Verdict 1: does navigate go via host? (first time a string argument works)* ---"
    Debug.Print "  1) Type  example.com  in the URL box and press Enter. If the following 1 line appears, it holds:"
    Debug.Print "         Wv2NavBar.HostNavigate: url=example.com (via host)"
    Debug.Print "     and confirm that the former postMessage-route log"
    Debug.Print "         Wv2NavBar.OnPaneWebMessage: msg={""cmd"":""navigate"",...}"
    Debug.Print "     does not appear. Also confirm that scheme completion (https://) works and it navigates."
    Debug.Print ""
    Debug.Print "  --- *Verdict 2: do back/forward/reload go via host? + double-execution check* ---"
    Debug.Print "  2) Navigate one more page in the same tab, then click <- 'only once'."
    Debug.Print "     The expectation is the following 1 line 'only':"
    Debug.Print "         Wv2NavBar.HostBack: (via host)"
    Debug.Print "     **This is the most important point this time** If any of the following happens, it is double execution:"
    Debug.Print "         - the HostBack line appears twice"
    Debug.Print "         - OnPaneWebMessage: msg={""cmd"":""back""} follows after HostBack"
    Debug.Print "         - one click goes back 2 pages"
    Debug.Print "       (A phenomenon that the dummy argument reserved was supposed to block. Report it immediately if it appears.)"
    Debug.Print "  3) Click -> once -> Wv2NavBar.HostForward: (via host) appears on only 1 line."
    Debug.Print "  4) Click the reload button once -> Wv2NavBar.HostReload: (via host) appears on"
    Debug.Print "     only 1 line. The page should reload only once."
    Debug.Print ""
    Debug.Print "  --- *Verdict 3: are string arguments intact? (spec fact 21 candidate)* ---"
    Debug.Print ("  5) Type Japanese in the URL box and press Enter (e.g. " & ChrW(&H5317&) & ChrW(&H6D77&) & ChrW(&H9053&) & " " & ChrW(&H5929&) & ChrW(&H6C17&) & ").")
    Debug.Print "     -> The Japanese should appear as is in the HostNavigate log (no garbling or loss)."
    Debug.Print ("         Wv2NavBar.HostNavigate: url=" & ChrW(&H5317&) & ChrW(&H6D77&) & ChrW(&H9053&) & " " & ChrW(&H5929&) & ChrW(&H6C17&) & " (via host)")
    Debug.Print "       It should be searched with the currently set search engine (the route of 9.14/9.16-9.20)."
    Debug.Print "  6) Also try a URL with symbols (e.g. https://www.google.com/search?q=a&b=c#x )."
    Debug.Print "     -> ? & # should appear as is in the log, and it should navigate to that page."
    Debug.Print "     NOTE: this was the spot that went through JSON escaping on the postMessage route. With hostObjects"
    Debug.Print "        the BSTR is passed directly, so it should actually be safer."
    Debug.Print ""
    Debug.Print "  --- Regression check (routes not touched) ---"
    Debug.Print "  7) Right after startup OnPaneWebMessage: msg={""cmd"":""ready""} appears and the URL box and"
    Debug.Print "     the button enabled states are initially synced (ready remains postMessage = correct)."
    Debug.Print "  8) Type something in the URL box and press Escape -> the typed characters vanish and the original URL returns"
    Debug.Print "     (immediate restore from lastUrl. The behavior of 9.12d must be intact)."
    Debug.Print "  9) Even if a navigation happens in the background while editing, the string being typed must not vanish (editing flag)."
    Debug.Print " 10) Tab switching / x / drag reordering must still work as in 9.23"
    Debug.Print "     (HostActivate / HostClose / HostReorder logs). The nav bar's URL box must"
    Debug.Print "     follow tab switching (ActiveChanged -> PushNavSyncToJs)."
    Debug.Print ""
    Debug.Print "  --- (Optional) Checking the fallback route is alive ---"
    Debug.Print " 11) Cannot be reached in a normal environment. Only if you can temporarily disable hostObjects in DevTools,"
    Debug.Print "     if the following 2 lines appear in succession on clicking <-, the delegation is alive:"
    Debug.Print "         Wv2NavBar.OnPaneWebMessage: msg={""cmd"":""back""}"
    Debug.Print "         -> back [postmsg fallback -> HostBack]"
    Debug.Print "         Wv2NavBar.HostBack: (via host)   NOTE: evidence of delegation to Host"
    Debug.Print ""
    Debug.Print "  NOTE: expected: appearance and feel are exactly the same as 9.23. The only difference is the log division"
    Debug.Print "     (back/forward/reload/navigate=HostXxx, ready=OnPaneWebMessage)."
End Sub
' ============================================================
' Test_9_26_Popup_On / Test_9_26_Popup_Off (Stage 9.26, real-machine switches)
'
'   Switches the delegation mode of the running Browser. After StartWebView2_Full,
'   just type the Sub name in the Immediate window.
'   - On  : delegate NewWindowRequested to the runtime (opens as a popup)
'   - Off : current behavior (put_Handled + our own new tab)
'   There is no need to reopen tabs. It takes effect immediately from the next link clicked.
' ============================================================
Public Sub Test_9_26_Popup_On()
    SetPopupDelegate True
End Sub

Public Sub Test_9_26_Popup_Off()
    SetPopupDelegate False
End Sub

Private Sub SetPopupDelegate(ByVal onOff As Boolean)
    Dim b As Wv2Browser
    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Debug.Print "[9.26] Browser is not started. Run UserForm1.StartWebView2_Full first."
        Exit Sub
    End If
    b.PopupDelegateMode = onOff
    Debug.Print "[9.26] current PopupDelegateMode = " & b.PopupDelegateMode
End Sub


' ============================================================
' Test_9_26_Popup_Help (Stage 9.26, real-machine procedure)
'
'   About the 405 on the new-tab expansion of POST links, this round only checks "if we delegate to the runtime,
'   is the POST preserved?". Note that the implementation change is a single branch and is not a permanent fix
'   (the popup comes out of our own tab UI).
' ============================================================
Public Sub Test_9_26_Popup_Help()
    Debug.Print "==== Stage 9.26 real-machine procedure (NewWindowRequested delegation probe) ===="
    Debug.Print ""
    Debug.Print "  [What we want to learn in this round]"
    Debug.Print "    If we delegate to the WebView2 runtime without setting put_Handled,"
    Debug.Print "    is the original POST (method + body + Referer) preserved and the 405 gone?"
    Debug.Print "    -> If it is gone, it is almost certain that the next stage's main approach (passing another tab's CoreWebView2"
    Debug.Print "       via args.put_NewWindow) will work. If not, the cause is something other than the POST body"
    Debug.Print "       (Referer/Cookie/session), so we change direction before the main approach."
    Debug.Print ""
    Debug.Print "  --- Procedure ---"
    Debug.Print "  1) UserForm1.StartWebView2_Full   ' start as usual"
    Debug.Print "  2) Log in to the core system (at this point it is still the current behavior = delegation OFF)"
    Debug.Print ("  3) *Control* With delegation still OFF, click '" & ChrW(&H30A2&) & ChrW(&H30DD&) & ChrW(&H30A4&) & ChrW(&H30F3&) & ChrW(&H30C8&) & ChrW(&H7BA1&) & ChrW(&H7406&) & "' (Appointment Management)")
    Debug.Print "       -> Expected: our own new tab opens and HTTP 405 is displayed (reproduces the current state)"
    Debug.Print "       -> Log: Wv2Pane.View_OnNewWindowRequested -> put_Handled OK"
    Debug.Print "                -> new tab creation requested of the Timer (RequestNewTabAsync)"
    Debug.Print "  4) Close the 405 tab and return to the original screen"
    Debug.Print "  5) Test_9_26_Popup_On            ' switch to delegation mode"
    Debug.Print ("  6) Click '" & ChrW(&H30A2&) & ChrW(&H30DD&) & ChrW(&H30A4&) & ChrW(&H30F3&) & ChrW(&H30C8&) & ChrW(&H7BA1&) & ChrW(&H7406&) & "' (Appointment Management) once more")
    Debug.Print "       -> Log: [9.26 delegation mode] delegating to the runtime without setting put_Handled"
    Debug.Print "          NOTE: the put_Handled OK / RequestNewTabAsync lines should not appear (report it if they do)"
    Debug.Print "  7) Check the contents of the popup window that opened"
    Debug.Print "  8) When done, close the popup with x and restore with Test_9_26_Popup_Off"
    Debug.Print ""
    Debug.Print "  --- How to read the results ---"
    Debug.Print "  *Pattern A: the business screen appeared normally in the popup*"
    Debug.Print "      = The POST is preserved by runtime delegation. The assumption about the cause was correct."
    Debug.Print "        The next stage is the main approach (put_NewWindow). We can concentrate on the design of the UX (fitting it into our own tab)."
    Debug.Print "  *Pattern B: still 405 even in the popup*"
    Debug.Print "      = Move on to suspecting factors other than the POST body (Referer / Cookie / session / user agent)."
    Debug.Print "        Building the main approach would likely not solve it, so it is good that this was found before starting."
    Debug.Print "        Make a note of the body of the 405 page (the full Tomcat message)."
    Debug.Print "  *Pattern C: the popup does not open / another symptom*"
    Debug.Print "      = Please share the log as is (whether the put_Handled line is absent is the deciding point)."
    Debug.Print ""
    Debug.Print "  --- Cautions ---"
    Debug.Print "  - Delegation mode is for verification. While it stays ON, every target=_blank link becomes a popup and"
    Debug.Print "    does not go onto the tab bar. Return to Test_9_26_Popup_Off once the check is done."
    Debug.Print "  - The popup is outside this Controller's management, so closing it is manual."
    Debug.Print "  - *Spec fact 20* Until the WebView2 event burst settles (all tabs finished loading +"
    Debug.Print "    the Immediate window log stops), do not break/step through the code."
    Debug.Print "    This round looks at branches inside event handlers, so the only evidence is the Debug.Print log."
End Sub
' ============================================================
' Test_9_26_PostProbe (Stage 9.26b, open the POST probe page)
'
'   Writes out %TEMP%\Wv2PostProbe\postprobe.html and opens it in a new tab as the virtual host
'   https://appassets.postprobe/postprobe.html.
'   After StartWebView2_Full, just type the Sub name in the Immediate window.
'
'   The page has 4 buttons (details in Test_9_26_PostProbe_Help).
' ============================================================
Public Sub Test_9_26_PostProbe()
    Dim b As Wv2Browser
    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Debug.Print "[9.26b] Browser is not started. Run UserForm1.StartWebView2_Full first."
        Exit Sub
    End If

    Dim folderPath As String
    folderPath = Environ$("TEMP")
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    folderPath = folderPath & "Wv2PostProbe"

    If Not WriteUtf8NoBom(folderPath, "postprobe.html", BuildPostProbeHtml()) Then
        Debug.Print "[9.26b] Failed to write out the HTML. Aborting."
        Exit Sub
    End If

    Dim pane As Wv2Pane
    Set pane = b.AddTabWithUrlForSpa("appassets.postprobe", folderPath, "postprobe.html")
    If pane Is Nothing Then
        Debug.Print "[9.26b] Failed to create the tab."
        Exit Sub
    End If

    Debug.Print "[9.26b] Opened the probe page. Current NewWindowMode = " & b.NewWindowMode
    Debug.Print "        (0=current/GET interception  1=popup delegation  2=pre-warm delegation)"
    Debug.Print "        For the procedure see Test_9_28_Help (latest) / Test_9_26_PostProbe_Help (roles of the buttons)."
    Debug.Print "        9.28b: the main destination is httpbingo.org. If it is down, use button 4 (httpbin) /"
    Debug.Print "        button 5 (postman-echo). *httpbingo does not show 405 in the title*, so"
    Debug.Print "        judge pass/fail by what is shown on screen, not by the log."
End Sub


' ============================================================
' Test_9_26_PostProbe_Help (Stage 9.26b, real-machine procedure)
' ============================================================
Public Sub Test_9_26_PostProbe_Help()
    Debug.Print "==== Stage 9.26b real-machine procedure (judging the delegation mode with the POST probe) ===="
    Debug.Print ""
    Debug.Print "  [Why this stage is needed]"
    Debug.Print "    The popup of the site tried in 9.26 was a plain GET, so even with delegation OFF"
    Debug.Print "    it opened normally in our own tab and the 405 was not reproduced. So we"
    Debug.Print "    use httpbin.org to synthesize at home the same picture as the workplace (405 for GET / 200 for POST)"
    Debug.Print "    and obtain the verdict on the main question."
    Debug.Print ""
    Debug.Print "  --- Procedure ---"
    Debug.Print "  1) UserForm1.StartWebView2_Full   ' start as usual"
    Debug.Print "  2) Test_9_26_PostProbe            ' the probe page opens in a new tab"
    Debug.Print "  3) *Control (delegation still OFF)* Press button 1 (POST + separate window)"
    Debug.Print "       -> Expected: our own new tab opens and 405 METHOD NOT ALLOWED appears"
    Debug.Print "       -> Log: put_Handled OK -> RequestNewTabAsync"
    Debug.Print "       *If 405 appears here, the workplace symptom has been reproduced at home*"
    Debug.Print "  4) Close that 405 tab and return to the probe page's tab"
    Debug.Print "  5) Test_9_26_Popup_On             ' switch to delegation mode"
    Debug.Print "  6) Press button 1 once more"
    Debug.Print "       -> Log: [9.26 delegation mode] delegating to the runtime without setting put_Handled"
    Debug.Print "       -> Expected: a 200 JSON appears in the popup"
    Debug.Print "  7) Close the popup with x and restore with Test_9_26_Popup_Off"
    Debug.Print ""
    Debug.Print "  --- Roles of the buttons ---"
    Debug.Print "  Button 1 [POST + separate window] *main one*"
    Debug.Print "      Should give 405 with delegation OFF and 200 with ON. Only this one changes behavior."
    Debug.Print "  Button 2 [POST + same tab]     sanity check"
    Debug.Print "      A route that does not go through NewWindowRequested. Should give 200 in either mode."
    Debug.Print "      If this is 200, it is settled that 'POST itself goes through normally in this WebView2'."
    Debug.Print "  Button 3 [GET + separate window]  control"
    Debug.Print "      Should give 200 in either mode. Evidence that the popup route itself is not broken."
    Debug.Print "      (reproduces the 'a GET popup opens even when OFF' seen in 9.26)"
    Debug.Print "  Button 4 [POST + separate window / spare destination 1]"
    Debug.Print "      An escape route (httpbin.org) for when httpbingo.org is down. Same role as button 1."
    Debug.Print "  Button 5 [POST + separate window / spare destination 2]  NOTE: newly added in 9.28b"
    Debug.Print "      A further escape route (postman-echo.com). Same role as button 1."
    Debug.Print ""
    Debug.Print "  --- How to read the results ---"
    Debug.Print "  *Pattern A: button 1 gives 405 when OFF / 200 when ON*"
    Debug.Print "      = The POST is preserved by runtime delegation. The assumption was correct."
    Debug.Print "        You may proceed to the next stage, the main approach (fitting it into our own tab with put_NewWindow)."
    Debug.Print "        Moreover, if the values of probe / nihongo / memo are visible in the form inside the 200 JSON,"
    Debug.Print "        that is direct evidence that the body was carried without loss (Japanese can be confirmed too)."
    Debug.Print "  *Pattern B: still 405 even when ON*"
    Debug.Print "      = The POST is not preserved even when delegating to the runtime. Building the main approach"
    Debug.Print "        would likely not solve it, so it is good that this was found before starting."
    Debug.Print "  *Pattern C: no 405 even when OFF*"
    Debug.Print "      = Reproducing the 405 failed. The destination may have changed, so"
    Debug.Print "        please share what was displayed (status, body) as is."
    Debug.Print "  *If button 2 does not give 200*"
    Debug.Print "      = The POST is not going through at all. The cause is in another layer, so consult."
    Debug.Print ""
    Debug.Print "  --- Cautions ---"
    Debug.Print "  - Run this in an environment that can reach the external network (httpbingo.org / httpbin.org / postman-echo.com)."
    Debug.Print "    Even if one gives 503, the remaining two let you continue verification."
    Debug.Print "  - *Caution from 9.28b* httpbingo's error page has no title, so Edge uses the"
    Debug.Print "    URL as the title. The title: line of the log cannot tell pass from fail, so"
    Debug.Print "    always judge by what is displayed (the text 405 or a 200 JSON)."
    Debug.Print "  - After the check, always restore the delegation mode with Test_9_26_Popup_Off."
    Debug.Print "  - *Spec fact 20* Do not break/step through until the event burst settles."
    Debug.Print "    Push through using only the Debug.Print log and the screen display as evidence."
End Sub


' ============================================================
' BuildPostProbeHtml (Stage 9.26b, private)
'
'   Builds the HTML of the verification page. *No JS at all* A plain <form> is
'   enough for the requirement, so JS is not brought into VBA strings (to avoid planting seeds of accidents).
'
'   The hidden field nihongo holds Japanese containing a full-width space, so
'   if the POST is preserved that string appears as is in httpbin's echo JSON.
'   = Direct evidence that the body was carried without loss.
' ============================================================
Private Function BuildPostProbeHtml() As String
    Dim s As String
    ' 9.28b: destinations made 3 routes. Main one = httpbingo.org (Go implementation of httpbin).
    '   The picture "405 for GET / 200 + echo the body for POST" is the same as httpbin.
    Dim ep As String
    Dim epGet As String
    Dim ep2 As String
    Dim ep3 As String
    ep = "https://httpbingo.org/post"
    epGet = "https://httpbingo.org/get"
    ep2 = "https://httpbin.org/post"
    ep3 = "https://postman-echo.com/post"

    s = "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    s = s & "<title>POST probe (Stage 9.26b / updated in 9.28b)</title>" & vbLf
    s = s & "<style>" & vbLf
    s = s & "  *{box-sizing:border-box;margin:0;padding:0;}" & vbLf
    s = s & "  body{font-family:'Segoe UI','Meiryo',sans-serif;color:#e8eaed;" & _
            "background:radial-gradient(1200px 600px at 20% -10%,#2b3550 0%,#151a26 55%,#0e121b 100%);" & _
            "min-height:100vh;padding:40px 28px;}" & vbLf
    s = s & "  .wrap{max-width:760px;margin:0 auto;}" & vbLf
    s = s & "  .eyebrow{letter-spacing:.22em;font-size:11px;color:#8ea2c8;text-transform:uppercase;}" & vbLf
    s = s & "  h1{font-size:24px;font-weight:650;margin:6px 0 6px;}" & vbLf
    s = s & "  .lead{color:#9aa7bd;font-size:13.5px;line-height:1.7;margin-bottom:24px;}" & vbLf
    s = s & "  .row{border:1px solid rgba(255,255,255,.08);border-radius:12px;padding:16px 18px;" & _
            "margin-bottom:14px;background:rgba(255,255,255,.03);}" & vbLf
    s = s & "  .row.main{border-color:rgba(110,168,254,.5);" & _
            "background:linear-gradient(180deg,rgba(110,168,254,.14),rgba(110,168,254,.04));}" & vbLf
    s = s & "  .tag{font-size:10.5px;letter-spacing:.08em;color:#6ea8fe;" & _
            "border:1px solid rgba(110,168,254,.4);border-radius:999px;padding:2px 9px;}" & vbLf
    s = s & "  .desc{font-size:12.5px;color:#9aa7bd;margin:8px 0 12px;line-height:1.6;}" & vbLf
    s = s & "  button{font-family:inherit;font-size:14px;font-weight:600;color:#0e121b;" & _
            "background:#6ea8fe;border:0;border-radius:8px;padding:9px 18px;cursor:pointer;}" & vbLf
    s = s & "  button:hover{background:#8fbcff;}" & vbLf
    s = s & "  .sub button{background:rgba(255,255,255,.14);color:#e8eaed;}" & vbLf
    s = s & "  .sub button:hover{background:rgba(255,255,255,.22);}" & vbLf
    s = s & "  label{display:block;font-size:11.5px;color:#8ea2c8;margin-bottom:5px;}" & vbLf
    s = s & "  input[type=text]{width:100%;font-family:inherit;font-size:13px;color:#cfe0ff;" & _
            "background:rgba(0,0,0,.28);border:1px solid rgba(255,255,255,.1);" & _
            "border-radius:8px;padding:8px 10px;margin-bottom:14px;}" & vbLf
    s = s & "  .ep{font-family:'Consolas','Courier New',monospace;font-size:11.5px;color:#7d8aa0;}" & vbLf
    s = s & "</style></head><body>" & vbLf
    s = s & "<div class=""wrap"">" & vbLf
    s = s & "  <div class=""eyebrow"">POST Probe</div>" & vbLf
    s = s & "  <h1>Judging the NewWindowRequested delegation mode</h1>" & vbLf
    s = s & "  <div class=""lead"">The destination returns 405 when hit with GET and 200 when hit with POST." & _
            "This is exactly the same picture as the workplace core system. Since 9.28 the default is pre-warm delegation (mode 2), so" & _
            "<b>pressing button 1 without switching anything should show a 200 in our own tab</b>, which is normal." & _
            "To see the control (405), first run Test_9_27_Mode_Legacy.<br>" & _
            "When the main destination is down, use the spare buttons (4 and 5) that have the same role.</div>" & vbLf

    ' --- common hidden fields (the same ones go into every form) ---
    Dim fields As String
    fields = "    <input type=""hidden"" name=""probe"" value=""wv2-post-probe"">" & vbLf & _
             ("    <input type=""hidden"" name=""nihongo"" value=""" & ChrW(&H65E5&) & ChrW(&H672C&) & ChrW(&H8A9E&) & ChrW(&H3000&) & ChrW(&H5168&) & ChrW(&H89D2&) & ChrW(&H30B9&) & ChrW(&H30DA&) & ChrW(&H30FC&) & ChrW(&H30B9&) & ChrW(&H5165&) & ChrW(&H308A&) & """>") & vbLf

    ' --- Button 1: POST + separate window (main) ---
    s = s & "  <div class=""row main"">" & vbLf
    s = s & "    <span class=""tag"">main</span>" & vbLf
    s = s & "    <div class=""desc"">POST + target=_blank. Should give 405 in mode 0 and 200 in mode 2." & _
            "Only this one changes behavior with the mode. If the method field of the response JSON is POST," & _
            "that is direct evidence that it was carried with its body without turning into a GET." & _
            "<br><span class=""ep"">" & ep & "</span></div>" & vbLf
    s = s & "    <form action=""" & ep & """ method=""post"" target=""_blank"">" & vbLf
    s = s & fields
    s = s & "    <label>memo (a string carried in the POST body. Feel free to rewrite it)</label>" & vbLf
    s = s & "    <input type=""text"" name=""memo"" value=""hello from Excel VBA WebView2"">" & vbLf
    s = s & "    <button type=""submit"">Button 1 : POST + separate window</button>" & vbLf
    s = s & "    </form>" & vbLf
    s = s & "  </div>" & vbLf

    ' --- Button 2: POST + same tab (sanity check) ---
    s = s & "  <div class=""row sub"">" & vbLf
    s = s & "    <span class=""tag"">sanity</span>" & vbLf
    s = s & "    <div class=""desc"">POST + same tab. A route that does not go through NewWindowRequested, so" & _
            "it should give 200 in either mode. Confirms that POST itself goes through in this WebView2.</div>" & vbLf
    s = s & "    <form action=""" & ep & """ method=""post"">" & vbLf
    s = s & fields
    s = s & "    <button type=""submit"">Button 2 : POST + same tab</button>" & vbLf
    s = s & "    </form>" & vbLf
    s = s & "  </div>" & vbLf

    ' --- Button 3: GET + separate window (control) ---
    s = s & "  <div class=""row sub"">" & vbLf
    s = s & "    <span class=""tag"">control</span>" & vbLf
    s = s & "    <div class=""desc"">GET + target=_blank. Should give 200 in either mode." & _
            "Confirms that the new-window route itself is not broken." & _
            "<br><span class=""ep"">" & epGet & "</span></div>" & vbLf
    s = s & "    <form action=""" & epGet & """ method=""get"" target=""_blank"">" & vbLf
    s = s & fields
    s = s & "    <button type=""submit"">Button 3 : GET + separate window</button>" & vbLf
    s = s & "    </form>" & vbLf
    s = s & "  </div>" & vbLf

    ' --- Button 4: POST + separate window (spare destination) ---
    s = s & "  <div class=""row sub"">" & vbLf
    s = s & "    <span class=""tag"">spare</span>" & vbLf
    s = s & "    <div class=""desc"">An escape route (1) with the same role as button 1. Use it when httpbingo.org is down." & _
            "<br><span class=""ep"">" & ep2 & "</span></div>" & vbLf
    s = s & "    <form action=""" & ep2 & """ method=""post"" target=""_blank"">" & vbLf
    s = s & fields
    s = s & "    <button type=""submit"">Button 4 : POST + separate window (spare 1 httpbin)</button>" & vbLf
    s = s & "    </form>" & vbLf
    s = s & "  </div>" & vbLf

    ' --- Button 5: POST + separate window (spare destination 2. Newly added in 9.28b) ---
    s = s & "  <div class=""row sub"">" & vbLf
    s = s & "    <span class=""tag"">spare</span>" & vbLf
    s = s & "    <div class=""desc"">An escape route (2) with the same role as button 1. By having 3 destinations," & _
            "it keeps verification from stopping when one external service goes down (we actually got tripped up by this in 9.28)." & _
            "<br><span class=""ep"">" & ep3 & "</span></div>" & vbLf
    s = s & "    <form action=""" & ep3 & """ method=""post"" target=""_blank"">" & vbLf
    s = s & fields
    s = s & "    <button type=""submit"">Button 5 : POST + separate window (spare 2 postman-echo)</button>" & vbLf
    s = s & "    </form>" & vbLf
    s = s & "  </div>" & vbLf

    s = s & "</div>" & vbLf
    s = s & "</body></html>"

    BuildPostProbeHtml = s
End Function


' ============================================================
' WriteUtf8NoBom (Stage 9.26b, private)
'
'   Writes a UTF-8 (no BOM) text file with ADODB.Stream.
'   Uses exactly the same procedure as Wv2TabBar.WriteTabBarHtml (write as text, re-read as binary,
'   discard the leading 3-byte BOM and then save).
'
'   *Why UserForm1.WriteSpaAppFolder is not used*
'     That one writes ANSI (CP932) with Print #, so putting Japanese into an HTML that declares
'     meta charset=UTF-8 garbles it. The method is aligned with TabBar/NavBar.
' ============================================================
Private Function WriteUtf8NoBom(ByVal folderPath As String, _
                                ByVal fileName As String, _
                                ByVal content As String) As Boolean
    On Error GoTo eh

    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
        Debug.Print "[9.26b] creating folder " & folderPath
    End If

    Dim fullPath As String
    fullPath = folderPath
    If Right$(fullPath, 1) <> "\" Then fullPath = fullPath & "\"
    fullPath = fullPath & fileName

    ' Write as text in UTF-8, re-read as binary and discard the BOM (EF BB BF)
    Dim bin As Object
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 2                ' adTypeText
    bin.Charset = "UTF-8"
    bin.Open
    bin.WriteText content
    bin.Position = 0
    bin.Type = 1                ' adTypeBinary
    bin.Position = 3            ' skip EF BB BF
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

    Debug.Print "[9.26b] write OK " & fullPath & " (" & Len(content) & " chars, UTF-8 without BOM)"
    WriteUtf8NoBom = True
    Exit Function
eh:
    Debug.Print "[9.26b] write error " & Err.Number & " " & Err.Description
    WriteUtf8NoBom = False
End Function


' ============================================================
' Stage 9.27 - real-machine switches and procedure for pre-warm delegation (put_NewWindow)
' ============================================================

' ============================================================
' Test_9_27_Mode_Legacy / _Popup / _Prewarm (Stage 9.27, mode switching)
'
'   A thin entry point that only switches Wv2Browser.NewWindowMode. Common to all tabs, and effective
'   immediately from the next link clicked (no need to reopen tabs).
'     0 = current behavior : put_Handled + our own new tab. Navigate is always GET -> POST gives 405
'     1 = popup            : left to the runtime. POST is preserved but it does not go onto the tab bar (9.26)
'     2 = pre-warm         : pass the spare tab's CoreWebView2 to put_NewWindow (main one of 9.27)
' ============================================================
Public Sub Test_9_27_Mode_Legacy()
    SetNewWindowMode 0
End Sub

Public Sub Test_9_27_Mode_Popup()
    SetNewWindowMode 1
End Sub

Public Sub Test_9_27_Mode_Prewarm()
    SetNewWindowMode 2
End Sub

Private Sub SetNewWindowMode(ByVal modeValue As Long)
    Dim b As Wv2Browser
    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Debug.Print "[9.27] Browser is not started. Run UserForm1.StartWebView2_Full first."
        Exit Sub
    End If
    b.NewWindowMode = modeValue
    Debug.Print "[9.27] current NewWindowMode = " & b.NewWindowMode
End Sub


' ============================================================
' Test_9_27_Status (Stage 9.27, state dump of the spare tab)
'
'   An entry point for checking, before opening in mode 2, whether the spare tab is warmed up.
'   Right after startup it shows "no spare (pending=True)", and after about 1.5 seconds changes to "spare present ... usable=True".
'   If it stays at usable=False, share the log.
' ============================================================
Public Sub Test_9_27_Status()
    Dim b As Wv2Browser
    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Debug.Print "[9.27] Browser is not started. Run UserForm1.StartWebView2_Full first."
        Exit Sub
    End If

    Debug.Print "==== Stage 9.27 state dump ===="
    Debug.Print "  NewWindowMode            = " & b.NewWindowMode & _
                "   (0=current / 1=popup / 2=pre-warm. The default since 9.28 is 2)"
    Debug.Print "  SetHandledWithNewWindow  = " & b.SetHandledWithNewWindow
    Debug.Print "  spare tab                = " & b.Debug_PrewarmStatus()
    Debug.Print "  number of visible tabs   = " & b.TabCount & "   (spare not included)"
    Debug.Print "  active index             = " & b.ActiveIndex
End Sub


' ============================================================
' Test_9_27_Handled_On / _Off (Stage 9.27, insurance switch for issue 5)
'
'   Whether to also set put_Handled(TRUE) right after setting the spare tab in put_NewWindow.
'   The default is On (the same form as the official sample). If symptoms such as the spare tab staying white / the window opening twice
'   appear, turn it Off and retry, to isolate which is correct.
'   Insurance so that *it can be isolated without fixing the code*.
' ============================================================
Public Sub Test_9_27_Handled_On()
    SetHandledFlag True
End Sub

Public Sub Test_9_27_Handled_Off()
    SetHandledFlag False
End Sub

Private Sub SetHandledFlag(ByVal onOff As Boolean)
    Dim b As Wv2Browser
    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Debug.Print "[9.27] Browser is not started. Run UserForm1.StartWebView2_Full first."
        Exit Sub
    End If
    b.SetHandledWithNewWindow = onOff
    Debug.Print "[9.27] current SetHandledWithNewWindow = " & b.SetHandledWithNewWindow
End Sub


' ============================================================
' Test_9_27_Help (Stage 9.27, real-machine procedure)
' ============================================================
Public Sub Test_9_27_Help()
    Debug.Print "==== Stage 9.27 real-machine procedure (main target: putting it on the spare tab with put_NewWindow) ===="
    Debug.Print ""
    Debug.Print "  [What we want to learn in this round]"
    Debug.Print "    Spec fact 23 settled that 'handing the navigation to the runtime preserves the POST'."
    Debug.Print "    This time we see whether it can be put on *our own tab*. If we pass the spare tab's (pre-warm)"
    Debug.Print "    CoreWebView2 to args.put_NewWindow and get a 200 + form echo while sitting on the tab bar"
    Debug.Print "    rather than in a popup, it is a success."
    Debug.Print ""
    Debug.Print "  --- Procedure ---"
    Debug.Print "  1) UserForm1.StartWebView2_Full   ' start as usual"
    Debug.Print "  2) After the startup log settles, Test_9_27_Status"
    Debug.Print "       -> Expected: spare tab = spare present ViewPtr=... / IsReady=True / usable=True"
    Debug.Print "       NOTE: right after startup it is no spare (pending=True). It warms up in about 1.5 seconds."
    Debug.Print "       NOTE: *spec fact 20* do not break/step through until the log stops."
    Debug.Print "  3) Test_9_26_PostProbe            ' open the 9.26b verification page in a new tab"
    Debug.Print "  4) *Control* Type Test_9_27_Mode_Legacy, then press button 1 (POST + _blank)"
    Debug.Print "       NOTE: the default became 2 in 9.28, so to take the control, explicitly return it to 0."
    Debug.Print "       -> Expected: 405 METHOD NOT ALLOWED in our own new tab (reproduces the same symptom as the workplace)"
    Debug.Print "       -> Log: put_Handled OK -> new tab creation requested of the Timer"
    Debug.Print "  5) Close the 405 tab and return to the probe page's tab"
    Debug.Print "  6) Test_9_27_Mode_Prewarm         ' to the main mode (restores the 9.28 default)"
    Debug.Print "  7) Press button 1 once more"
    Debug.Print "  8) Test_9_27_Status               ' check whether it was replenished"
    Debug.Print "       -> Expected: spare present / usable=True (what was used has been replenished)"
    Debug.Print ""
    Debug.Print "  --- Expected log (step 7) ---"
    Debug.Print "    Wv2Pane.View_OnNewWindowRequested: pSender=... , pArgs=..."
    Debug.Print "      -> new window request URL: https://httpbin.org/post"
    Debug.Print "    Wv2Browser.ClaimPrewarmViewPtr: reserved the spare tab (ViewPtr=...)"
    Debug.Print "      [9.27 pre-warm] put_NewWindow OK (spare tab's View=...)"
    Debug.Print "        put_Handled OK (used together with put_NewWindow)"
    Debug.Print "    Wv2Browser.RequestAdoptPrewarmedAsync: reserved adoption of the spare tab (...)"
    Debug.Print "    Wv2Browser.Adopt: adopted the spare tab as a tab (current tab count=N, requested URL=...)"
    Debug.Print "    Wv2Browser.Prewarm: started creating the spare tab -> warmed up the spare tab (ViewPtr=...)"
    Debug.Print ""
    Debug.Print "  --- How to read the results ---"
    Debug.Print "  *Pattern A (success): a new tab was added to the tab bar and a 200 JSON appeared in it*"
    Debug.Print "      If the form shows 3 fields / Content-Length: 172, the POST is completely"
    Debug.Print "      preserved. The main approach holds. The 405 problem is solved including the UX."
    Debug.Print "  *Pattern B: a tab was added but its content is blank white / windows open twice*"
    Debug.Print "      Using put_Handled together may be causing trouble. Switch to Test_9_27_Handled_Off and"
    Debug.Print "      redo steps 5-7. If this fixes it, 'do not set Handled once NewWindow is set'"
    Debug.Print "      is confirmed as correct (the insurance of issue 5 worked)."
    Debug.Print "  *Pattern C: a fallback log marked with * appeared*"
    Debug.Print "      The case where the spare tab could not be used. Share it together with the result of Test_9_27_Status."
    Debug.Print "      (the cause, e.g. no spare / not replenished while usable=False, is isolated with the state dump)"
    Debug.Print "  *Pattern D: put_NewWindow failed with an hr*"
    Debug.Print "      Please share the value of hr as is. It will be a vtable index or argument type problem."
    Debug.Print ""
    Debug.Print "  --- Cautions ---"
    Debug.Print "  - The spare tab does not appear on the tab bar (because it is not put in m_tabs). It is normal that Excel's task"
    Debug.Print "    manager shows one more msedgewebview2.exe."
    Debug.Print "  - Test_9_26_Popup_On / _Off still work in 9.27 (internally delegating to mode 1 / 0)."
    Debug.Print "    However _Off also cancels mode 2, so to return to the main mode use Test_9_27_Mode_Prewarm."
End Sub


' ============================================================
' Test_9_28_Help (Stage 9.28, real-machine procedure)
'
'   Confirmation of the state where the default is 2 (pre-warm delegation). The biggest difference from 9.27 is that
'   *a POST link gives 200 without typing any mode switch*.
' ============================================================
Public Sub Test_9_28_Help()
    Debug.Print "==== Stage 9.28 real-machine procedure (making pre-warm delegation the default + 2 small fixes) ===="
    Debug.Print ""
    Debug.Print "  [What we want to learn in this round]"
    Debug.Print "    (A) Does a POST link give 200 in our own tab without typing a mode switch?"
    Debug.Print "    (B) Even if nesting occurs at startup, is the spare eventually warmed up? (counter-ization)"
    Debug.Print "    (C) Is the log for the spare running out worded to match reality?"
    Debug.Print ""
    Debug.Print "  --- Procedure ---"
    Debug.Print "  1) UserForm1.StartWebView2_Full   ' start as usual"
    Debug.Print "       NOTE: *spec fact 20* do not break/step through until the log stops."
    Debug.Print "  2) After the startup log settles, Test_9_27_Status"
    Debug.Print "       -> Expected: NewWindowMode = 2   *if this is 0, the default change was not applied*"
    Debug.Print "       -> Expected: spare tab = spare present ViewPtr=... / IsReady=True / usable=True"
    Debug.Print "  3) Test_9_26_PostProbe            ' open the 9.26b verification page in a new tab"
    Debug.Print "  4) *Main point* Press button 1 (POST + _blank) without touching the mode at all"
    Debug.Print "       -> Expected: a new tab is added to the tab bar and a 200 JSON appears in it"
    Debug.Print "               *Judge on the screen* If the method of the response JSON is POST, and"
    Debug.Print "               the 3 fields probe / nihongo / memo are visible in the form, the POST survived."
    Debug.Print "               (httpbingo does not show 405 in the title, so the log cannot tell)"
    Debug.Print "       -> Expected log:"
    Debug.Print "           Wv2Browser.ClaimPrewarmViewPtr: reserved the spare tab (ViewPtr=...)"
    Debug.Print "             [9.27 pre-warm] put_NewWindow OK"
    Debug.Print "           Wv2Browser.Adopt: adopted the spare tab as a tab (current tab count=N, ...)"
    Debug.Print "           Wv2Browser.Prewarm: warmed up the spare tab (ViewPtr=...)"
    Debug.Print "  5) Test_9_27_Status               ' check whether it was replenished"
    Debug.Print "       -> Expected: spare present / usable=True"
    Debug.Print ""
    Debug.Print "  --- Check of (C) (optional. It is only a log wording fix, so if you have time) ---"
    Debug.Print "  6) Return to the probe page and press button 1 *twice quickly* (to make the spare run out)"
    Debug.Print "       -> Expected: the 2nd one shows *the spare tab cannot be used, so falling back to current behavior*, and"
    Debug.Print "               then the next 1 line appears (the spot whose wording was changed in 9.28):"
    Debug.Print "           Wv2Browser.RequestPrewarmAsync: the spare tab is already Claimed (awaiting adoption)."
    Debug.Print "           It is replenished automatically right after adoption, so no reservation is made here"
    Debug.Print "       NOTE: in 9.27 this spot said 'there is already a spare tab so do nothing', which was confusing."
    Debug.Print "       NOTE: the 2nd one gives 405 (because there is only 1 spare). This is as expected."
    Debug.Print "       NOTE: mashing or a fault in the external service may return 503 (easily confused with 405)."
    Debug.Print "         How to isolate: type the destination URL directly in the address bar and see whether a 405 appears."
    Debug.Print "         If it is 503, the server is down, so switch to the spares of button 4 / button 5."
    Debug.Print ""
    Debug.Print "  --- Check of (B) (just look at the log) ---"
    Debug.Print "  7) The following line may appear in the startup log (in fact it is proof that nesting occurred):"
    Debug.Print "       Wv2Browser.TimerCall(prewarm): a Pane was being created, so skipped this time"
    Debug.Print "     Even then it is eventually warmed up by the re-reservation at the end of AddTab, so in step 2"
    Debug.Print "     it is normal if usable=True."
    Debug.Print ""
    Debug.Print "  --- At shutdown ---"
    Debug.Print "  8) Close the UserForm"
    Debug.Print "       -> Expected: 'Wv2Browser.Shutdown: released the spare tab (pre-warm)' appears"
    Debug.Print ""
    Debug.Print "  --- To return to the former behavior ---"
    Debug.Print "  - Test_9_27_Mode_Legacy   (mode 0 = the default before 9.27. POST goes back to 405)"
    Debug.Print "  - Test_9_27_Mode_Prewarm  (mode 2 = back to the default of 9.28)"
End Sub


' ============================================================
' Test_9_29_Custom  (Stage 9.29 pure-logic verification of custom search engines)
'
'   *Just type  Test_9_29_Custom  in the Immediate window*
'
'   Checks the custom-related members of Wv2SettingsBridge (SetCustomEngine / GetCustomTemplate /
'   PreviewUrlForTemplate) and the persistence of custom (2 lines: engine + template).
'   WebView2 is not started (only the bridge + Browser's Debug_*).
'   Because it touches the real file, it saves engine and template as a pair at the top and restores them at the end.
' ============================================================
Public Sub Test_9_29_Custom()
    Dim total As Long, pass As Long
    total = 0: pass = 0
    Debug.Print "==== Test_9_29_Custom (custom search engine, pure logic) ===="

    Dim br As Wv2SettingsBridge
    Set br = New Wv2SettingsBridge

    ' --- save the current stored values aside (the engine + template pair) ---
    Dim savedName As String, savedTpl As String
    savedName = br.LoadEngineName()
    savedTpl = br.LoadCustomTemplate()
    Debug.Print "  (saved aside) engine = '" & savedName & "' / template = '" & savedTpl & "'"

    Dim b As Wv2Browser
    Set b = New Wv2Browser
    br.BindBrowser b

    Dim before As String
    before = b.Debug_SearchEngine()

    Const TPL As String = "https://example.com/find?query="

    ' --- invalid input: the return value is "", and neither Browser nor the saved file is touched ---
    CheckEq "empty string -> return value is empty", br.SetCustomEngine(""), "", total, pass
    CheckEq "no scheme -> return value is empty", br.SetCustomEngine("example.com/?q="), "", total, pass
    CheckEq "ftp:// -> return value is empty", br.SetCustomEngine("ftp://example.com/?q="), "", total, pass
    CheckEq "Browser unchanged after invalid input", b.Debug_SearchEngine(), before, total, pass

    ' --- normal application ---
    CheckEq "apply -> 'custom'", br.SetCustomEngine(TPL), "custom", total, pass
    CheckEq "reflected in Browser", b.Debug_SearchEngine(), TPL, total, pass
    CheckEq "GetEngine is custom", br.GetEngine(), "custom", total, pass
    CheckEq "GetCustomTemplate", br.GetCustomTemplate(), TPL, total, pass
    CheckEq "search URL generation (production route)", b.Debug_NormalizeUrl("hokkaido"), TPL & "hokkaido", total, pass

    ' --- persistence (custom is the 2 lines engine + template) ---
    CheckEq "save -> LoadEngineName", br.LoadEngineName(), "custom", total, pass
    CheckEq "save -> LoadCustomTemplate", br.LoadCustomTemplate(), TPL, total, pass
    CheckEq "save -> LoadEngineTemplate", br.LoadEngineTemplate(), TPL, total, pass

    ' --- preview before applying (does not change Browser's state) ---
    CheckEq "PreviewUrlForTemplate", _
            br.PreviewUrlForTemplate("https://foo.test/s?k=", "hokkaido"), _
            "https://foo.test/s?k=hokkaido", total, pass
    CheckEq "PreviewUrlForTemplate invalid -> empty", _
            br.PreviewUrlForTemplate("foo.test/s?k=", "hokkaido"), "", total, pass
    CheckEq "preview does not change Browser", b.Debug_SearchEngine(), TPL, total, pass

    ' --- returning to a preset erases the template line (issue 3, plan a-1) ---
    CheckEq "SetEngine bing", br.SetEngine("bing"), "bing", total, pass
    CheckEq "-> LoadEngineName", br.LoadEngineName(), "bing", total, pass
    CheckEq "-> LoadCustomTemplate is empty", br.LoadCustomTemplate(), "", total, pass
    CheckEq "-> LoadEngineTemplate", br.LoadEngineTemplate(), _
            "https://www.bing.com/search?q=", total, pass
    CheckEq "-> GetCustomTemplate is also empty", br.GetCustomTemplate(), "", total, pass

    ' --- if the input is the same URL as a preset, it becomes that preset's name ---
    CheckEq "same URL as preset -> 'google'", _
            br.SetCustomEngine("https://www.google.com/search?q="), "google", total, pass
    CheckEq "-> template line is not written", br.LoadCustomTemplate(), "", total, pass

    ' --- a broken template line falls back to "" on restore (starts with the default Google) ---
    br.Debug_SaveSettings "custom", "example.com/?q="
    CheckEq "broken template -> LoadEngineTemplate is empty", br.LoadEngineTemplate(), "", total, pass

    ' --- write back the saved-aside values (restoring the user settings) ---
    If Len(savedName) > 0 Then
        br.Debug_SaveSettings savedName, savedTpl
        Debug.Print "  (restored) put back to engine = '" & savedName & "' / template = '" & savedTpl & "'."
    Else
        br.Debug_SaveSettings "google", ""
        Debug.Print "  (restored) Nothing had been saved originally, so wrote engine=google (same as the default, so no harm)."
    End If

    Debug.Print "==== Result: " & pass & " / " & total & _
                IIf(pass = total, "  [ALL OK]", "  [!! FAILED !!]")
End Sub


' ============================================================
' Test_9_29_Help  (Stage 9.29 real-machine procedure)
'
'   *Type  Test_9_29_Help  in the Immediate window to show the procedure*
' ============================================================
Public Sub Test_9_29_Help()
    Debug.Print "==== Stage 9.29 real-machine procedure (custom search URL + cleanup of spare reservations) ===="
    Debug.Print ""
    Debug.Print "  [What we want to learn in this round]"
    Debug.Print "    (A) Can an arbitrary search URL be applied from the settings tab's custom input box?"
    Debug.Print "    (B) Is it restored after restarting Excel? (engine=custom + template)"
    Debug.Print "    (C) Has the wasteful reservation log at the spare-runs-out fallback gone?"
    Debug.Print ""
    Debug.Print "  --- Preliminaries (optional) ---"
    Debug.Print "  0) Test_9_29_Custom            ' the pure logic should be [ALL OK]"
    Debug.Print ""
    Debug.Print "  --- (A) Applying a custom engine ---"
    Debug.Print "  1) UserForm1.StartWebView2_Full   ' start as usual"
    Debug.Print "       NOTE: *spec fact 20* do not break/step through until the log stops."
    Debug.Print "  2) Press [gear] to open the settings tab"
    Debug.Print "       -> Expected: below the cards, a frame and input box for *custom (other than presets)* appear"
    Debug.Print "  3) Paste the following into the input box, but do not press Apply yet:"
    Debug.Print "         https://ja.wikipedia.org/w/index.php?search="
    Debug.Print ("       -> Expected: while typing, the line below shows  " & ChrW(&H65E5&) & ChrW(&H672C&) & " " & ChrW(&H5929&) & ChrW(&H6C17&) & " -> https://ja.wikipedia.org/... ")
    Debug.Print "               i.e. *a preview before applying* (VBA is doing the encoding)"
    Debug.Print "  4) Deliberately type something with no scheme such as  example.com/?q= "
    Debug.Print "       -> Expected: a red message 'The URL must start with http:// or https:// ...' appears"
    Debug.Print "               *Even if you press [Apply] in this state, the log only says 'invalid template, ignored'"
    Debug.Print "                 and the search engine does not change*"
    Debug.Print "  5) Go back to the URL from step 3 and press [Apply] (Enter does the same)"
    Debug.Print "       -> Expected log:"
    Debug.Print "           Wv2Browser.SearchEngine: search template set = https://ja.wikipedia.org/..."
    Debug.Print "           Wv2SettingsBridge.SaveSettingsFile: saved 'custom' / template=... -> ...\settings.txt"
    Debug.Print "           Wv2SettingsBridge.SetCustomEngine: applied '...' -> 'custom'"
    Debug.Print "       -> Expected: the custom frame glows blue and an 'in use' badge appears. All cards are unselected."
    Debug.Print ("  6) In another normal tab type   " & ChrW(&H65E5&) & ChrW(&H672C&) & " " & ChrW(&H5929&) & ChrW(&H6C17&) & "   in the address bar")
    Debug.Print "       -> Expected: Wikipedia search results appear"
    Debug.Print ""
    Debug.Print "  --- (B) Persistence ---"
    Debug.Print "  7) Fully quit Excel (not just close the workbook, but Excel itself)"
    Debug.Print "  8) Restart and run StartWebView2_Full"
    Debug.Print "       -> Expected log (proof of restoring):"
    Debug.Print "           Wv2Browser.Class_Initialize: restored saved engine = https://ja.wikipedia.org/..."
    Debug.Print ("  9) Without opening the settings tab, immediately type   " & ChrW(&H65E5&) & ChrW(&H672C&) & " " & ChrW(&H5929&) & ChrW(&H6C17&) & "   in a normal tab")
    Debug.Print "       -> Expected: Wikipedia search results (the custom engine was restored)"
    Debug.Print " 10) Open the settings tab"
    Debug.Print "       -> Expected: the previous URL is in the input box and the 'in use' badge is shown"
    Debug.Print ""
    Debug.Print "  --- (C) Cleanup of spare reservations (just look at the log) ---"
    Debug.Print " 11) Open the probe page with Test_9_26_PostProbe and press button 1 *twice quickly*"
    Debug.Print "       -> Expected: after the 2nd one shows *the spare tab cannot be used, so falling back to current behavior*,"
    Debug.Print "               the following line does *not appear* even after waiting 1.5 seconds (the wasteful log removed in 9.29):"
    Debug.Print "                 Wv2Browser.Prewarm: a spare tab already exists (skipping creation)"
    Debug.Print "       -> Instead the following may appear (this is normal. A sign that nesting was detected):"
    Debug.Print "           Wv2Browser.EnsurePrewarmScheduled: the spare tab is being created (nested), so no reservation is made"
    Debug.Print " 12) Test_9_27_Status               ' is the spare kept at 1?"
    Debug.Print "       -> Expected: spare present / usable=True"
    Debug.Print ""
    Debug.Print "  --- To return to the former state ---"
    Debug.Print "  - Click the Google card on the settings tab to go back"
    Debug.Print "    (at that time the template line of settings.txt is erased = by design)"
    Debug.Print "  - Deleting %APPDATA%\Wv2Browser\settings.txt by hand and restarting also returns to the default."
End Sub
' ============================================================
' Test_9_31_Help  (Stage 9.31 real-machine procedure / visual check)
'
'   *Type  Test_9_31_Help  in the Immediate window to show the procedure*
'
'   9.31 changes only CSS, so there is no pure-logic test.
'   Along with it, open item 1 from 9.30 (end-to-end check of v0_5_2) is worked off here.
' ============================================================
Public Sub Test_9_31_Help()
    Debug.Print "==== Stage 9.31 real-machine procedure (tab hover + show/hide control of the close button) ===="
    Debug.Print ""
    Debug.Print "  [What we want to learn in this round]"
    Debug.Print "    (A-1) Does the background get lighter when the mouse is over a tab? (inactive only)"
    Debug.Print "    (A-2) Is the close button visible 'only when active or hovered'?"
    Debug.Print "    (C)   *Open item from 9.30* Does v0_5_2 work normally end to end?"
    Debug.Print ""
    Debug.Print "  NOTE: this round changes *CSS only*. The log wording is completely unchanged from 9.29."
    Debug.Print "     Make all verdicts *by looking at the screen*."
    Debug.Print ""
    Debug.Print "  --- Preparation ---"
    Debug.Print "  1) UserForm1.StartWebView2_Full   ' start as usual"
    Debug.Print "       NOTE: *spec fact 20* do not break/step through until the log stops."
    Debug.Print "  2) Press + 2-3 times to make 3-4 tabs"
    Debug.Print "       (the aim is to create a state with 1 active tab + several inactive ones)"
    Debug.Print ""
    Debug.Print "  --- (A-1) Hover background ---"
    Debug.Print "  3) Put the mouse over an *inactive* tab -> then move it away"
    Debug.Print "       -> Expected: when over it, the background becomes slightly lighter (#e7e7e7 -> #efefef),"
    Debug.Print "               and returns when moved away. The change is *gradual (0.12 s)*"
    Debug.Print "       -> Expected: neither the tab's *width nor the text position moves* (only the color changes)"
    Debug.Print "  4) Put the mouse over the *active* tab"
    Debug.Print "       -> Expected: *nothing changes* (it is white already = the lightest, so by design it is not changed)"
    Debug.Print "  5) Switch tabs (click another tab)"
    Debug.Print "       -> Expected: the change to white is also smooth at the same 0.12 s"
    Debug.Print "               NOTE: if this looks unnatural, it is material for deciding to return issue 5 to 'no transition'"
    Debug.Print ""
    Debug.Print "  --- (A-2) Show/hide control of the close button ---"
    Debug.Print "  6) Move the mouse completely out of the tab bar (onto the page body, etc.) and look at the tab row"
    Debug.Print "       -> Expected: *the x is visible only on the active tab*"
    Debug.Print "       -> Expected: the x is *not visible* on inactive tabs"
    Debug.Print "  7) Put the mouse over an inactive tab"
    Debug.Print "       -> Expected: the x appears only on that tab (immediately. By design there is no fade)"
    Debug.Print "       -> Expected: *the title text does not move left or right*"
    Debug.Print "               (because visibility:hidden reserves the space. If this moves,"
    Debug.Print "                it has become display:none, so report it)"
    Debug.Print "  8) Put the mouse over the x"
    Debug.Print "       -> Expected: the x's background becomes a gray rounded shape (same as up to 9.29)"
    Debug.Print "  9) Click the 'place where the x should be' on an inactive tab where the x is not visible"
    Debug.Print "       -> Expected: *the tab does not close* (= an invisible button cannot be pressed)."
    Debug.Print "               The correct behavior is that it *switches* to that tab."
    Debug.Print "               NOTE: if it closes, it is equivalent to opacity:0, so report it"
    Debug.Print " 10) Increase the tabs to 8-10 and recheck 6-7 with the tabs narrowed"
    Debug.Print "       -> Expected: it bottoms out at min-width:80px, and the x's space is always reserved"
    Debug.Print "       -> If complaints such as the titles being too unreadable arise here, use them as material for the next stage"
    Debug.Print ""
    Debug.Print "  --- (B) Regression: drag reordering must not be broken ---"
    Debug.Print " 11) Grab a tab, move it left and right, and release it over another tab"
    Debug.Print "       -> Expected: the grabbed tab becomes faint (opacity 0.45),"
    Debug.Print "               and a *blue vertical line* appears on the left or right of the destination tab"
    Debug.Print "       -> Expected log: [host] reorder N->M (on the WebView2 DevTools side)"
    Debug.Print "       -> Expected: on release the order changes and the active tab follows"
    Debug.Print "       NOTE: the hover color may appear during a drag, but this is *by design* (deferred in issue 7)"
    Debug.Print ""
    Debug.Print "  --- (C) *Open item 1 from 9.30: end-to-end check of v0_5_2* ---"
    Debug.Print "  NOTE: this part is unrelated to the 9.31 change. It is a procedure for stepping through, in one go,"
    Debug.Print "    the 3 items left unchecked in the clean-built workbook, taking this opportunity."
    Debug.Print ""
    Debug.Print " 12) Test_9_27_Status               ' is the spare tab warmed up?"
    Debug.Print "       -> Expected: spare present / usable=True"
    Debug.Print " 13) Test_9_26_PostProbe            ' open the POST probe page"
    Debug.Print " 14) Press button 1 (httpbingo POST + _blank)"
    Debug.Print "       -> Expected: *our own new tab* opens, and method: POST and"
    Debug.Print "               the contents of the form appear on screen (neither 405 nor blank)"
    Debug.Print "       -> *Caution* The title is httpbingo.org/post whether it succeeds or fails, so"
    Debug.Print "                  always judge by *what is shown on screen* (from 9.28b)"
    Debug.Print "       -> If a 503 appears, go to the isolation procedure in Test_9_28_Help (a server-side problem)"
    Debug.Print " 15) Test_9_27_Status               ' was the spare replenished?"
    Debug.Print "       -> Expected: spare present / usable=True (replenished right after being used)"
    Debug.Print " 16) Ctrl+click a link on any page, or press a target=_blank link"
    Debug.Print "       -> Expected: it opens in our own tab (not as a popup window)"
    Debug.Print ""
    Debug.Print "  --- Verdict ---"
    Debug.Print "  - 3-11 all as expected -> 9.31 passes"
    Debug.Print "  - 12-16 all as expected -> *the end-to-end check of v0_5_2 (open item 1 from 9.30) is worked off*"
    Debug.Print "  - Anything that bothers you (color too dark/light, transition too slow/fast, x too small, etc.)"
    Debug.Print "    can be handled by readjusting issues 2, 5 and 6, so please do raise it."
End Sub
' ============================================================
' Test_9_32_Help  (Stage 9.32 real-machine procedure / visual and feel check)
'
'   *Type  Test_9_32_Help  in the Immediate window to show the procedure*
'
'   9.32 changes only JS and CSS. There is no pure-logic test.
' ============================================================
Public Sub Test_9_32_Help()
    Debug.Print "==== Stage 9.32 real-machine procedure (optimistic update of tab switching + cleanup of the scrollbar) ===="
    Debug.Print ""
    Debug.Print "  [What we want to learn in this round]"
    Debug.Print "    (1) Does the tab color change the *instant* it is clicked? (the body switch catches up afterwards)"
    Debug.Print "    (2) Even with more tabs, no horizontal scrollbar appears and the tabs do not get squashed?"
    Debug.Print "    (3) Can the tab row be scrolled horizontally with the mouse wheel?"
    Debug.Print "    (4) Regression: are D&D reordering / closing / the settings tab still intact?"
    Debug.Print ""
    Debug.Print "  NOTE: not one line of the VBA logic was changed, so *both the wording and the order of the log are the same as 9.31*."
    Debug.Print "     The only change is that 'the tab bar color changes before the log appears'."
    Debug.Print ""
    Debug.Print "  --- Preparation ---"
    Debug.Print "  1) UserForm1.StartWebView2_Full   ' start as usual"
    Debug.Print "       NOTE: *spec fact 20* do not break/step through until the log stops."
    Debug.Print "  2) Press + to make 4-5 tabs, and open a different page in each"
    Debug.Print "       (pages that take time to switch make the difference easier to see)"
    Debug.Print ""
    Debug.Print "  --- (1) *Main point: the feel of tab switching* ---"
    Debug.Print "  3) Click an inactive tab"
    Debug.Print "       -> Expected: *the tab bar color switches almost at the same time as the click*"
    Debug.Print "               (it turns white the instant it is pressed, and the tab that was held returns to gray)"
    Debug.Print "       -> Expected: the browser body display switches *after* that"
    Debug.Print "       -> Expected: the Immediate window log also appears *after* the tab bar color change"
    Debug.Print "               Wv2TabBar.HostActivate: index=N (via host)"
    Debug.Print "               Wv2TabBar[EVT] ActiveChanged(index=N) -> PushSyncToJs"
    Debug.Print "               Wv2TabBar.PushSyncToJs: sent OK (...)"
    Debug.Print "  4) Switch tabs 5-6 times in a row and compare the feel with 9.31"
    Debug.Print "       -> Verdict: has the feeling that 'the click and the screen change are out of step' decreased?"
    Debug.Print "       *Caution* The switching of the browser body itself has *not become faster*."
    Debug.Print "               Only 'the time until the tab bar reacts' has become faster."
    Debug.Print "               If you feel the body side is still too slow, a separate stage looking at the contents of"
    Debug.Print "               ActivateTab will be needed, so please report it so."
    Debug.Print "  5) Also look at the feel when pressing the close button (x)"
    Debug.Print "       -> Expected: *this is the same as 9.31 (still slow)* which is correct."
    Debug.Print "               Because issue 3 decided not to put the optimistic update into close."
    Debug.Print "               If it bothers you, the same technique can be put into close in the next stage."
    Debug.Print ""
    Debug.Print "  --- (2)(3) Scrollbar and wheel ---"
    Debug.Print "  6) Mash + to increase the tabs to 10-12"
    Debug.Print "       -> Expected: *no horizontal scrollbar appears under the tab row*"
    Debug.Print "       -> Expected: the tab height is *not squashed* as it was in 9.31"
    Debug.Print "               (in 9.31 the bar ate into the height from the inside and the tabs were pushed up)"
    Debug.Print "  7) Put the mouse over the tab row and roll the wheel up and down"
    Debug.Print "       -> Expected: the tab row *scrolls left and right*"
    Debug.Print "               (Chromium converts the vertical wheel into horizontal scrolling)"
    Debug.Print "       -> Expected: the page of the browser body does not scroll"
    Debug.Print "  8) Scroll to the right end and click the rightmost tab"
    Debug.Print "       -> Expected: as in step 3 the color changes immediately and the switch works correctly"
    Debug.Print "       NOTE: if the wheel does not work, please report it (we will return to the alternative of issue 4)"
    Debug.Print ""
    Debug.Print "  --- (4) Regression ---"
    Debug.Print "  9) Grab a tab and reorder it"
    Debug.Print "       -> Expected: faint + blue vertical line + on release the order changes (same as 9.31)"
    Debug.Print "       -> Expected: *grabbing an inactive tab does not make it active*"
    Debug.Print "               (a specification decided in issue 6 to keep as is)"
    Debug.Print " 10) Close several tabs with x"
    Debug.Print "       -> Expected: they close correctly and the active tab is corrected"
    Debug.Print " 11) Open the settings tab with [gear] and click a search engine card"
    Debug.Print "       -> Expected: the same as up to 9.29 (the selection is reflected and the save log appears)"
    Debug.Print " 12) Close the settings tab and try searching in a normal tab's address bar"
    Debug.Print "       -> Expected: searching works with the chosen engine"
    Debug.Print ""
    Debug.Print "  --- Verdict ---"
    Debug.Print "  - 3-4 improve the feel -> the main point of 9.32 succeeded"
    Debug.Print "  - 6-8 as expected -> the scrollbar cleanup also succeeded"
    Debug.Print "  - 9-12 same as 9.31 -> no regression"
    Debug.Print "  - If anything feels off, e.g. the color change is so fast that it seems to change 'before pressing',"
    Debug.Print "    please report it. It can be adjusted by reducing the double rAF to a single one, etc."
End Sub
' ============================================================
' Test_9_32b_Help  (Stage 9.32b real-machine procedure / visual check)
'
'   *Type  Test_9_32b_Help  in the Immediate window to show the procedure*
' ============================================================
Public Sub Test_9_32b_Help()
    Debug.Print "==== Stage 9.32b real-machine procedure (wheel horizontal scrolling + keeping the scroll position) ===="
    Debug.Print ""
    Debug.Print "  [What we want to learn in this round]"
    Debug.Print "    (1) Does the tab row scroll horizontally when the wheel is rolled over it?"
    Debug.Print "    (2) Does the scroll position stay put (not jump back to the left edge) even when the title-update sync arrives?"
    Debug.Print "    (3) Does a new tab added with + come into a visible position automatically?"
    Debug.Print "    (4) Regression: switching feel / D&D / closing"
    Debug.Print ""
    Debug.Print "  NOTE: not one line of VBA was changed, so *the log is exactly the same as 9.32*."
    Debug.Print ""
    Debug.Print "  --- Preparation ---"
    Debug.Print "  1) UserForm1.StartWebView2_Full"
    Debug.Print "       NOTE: *spec fact 20* do not break/step through until the log stops."
    Debug.Print "  2) Mash + to increase the tabs to about 12 (to create an overflowing state)"
    Debug.Print ""
    Debug.Print "  --- (1) Wheel ---"
    Debug.Print "  3) Put the mouse over the tab row and roll the wheel down"
    Debug.Print "       -> Expected: the tab row *scrolls right*"
    Debug.Print "       -> Expected: the page of the browser body *does not scroll*"
    Debug.Print "  4) Roll the wheel up -> it goes back left"
    Debug.Print "  5) Roll the wheel also with only 1-2 tabs (a state with no overflow)"
    Debug.Print "       -> Expected: *nothing happens* (early return with scrollWidth <= clientWidth)"
    Debug.Print "               If the body page scrolls here, that is as designed (because it does not reach preventDefault)."
    Debug.Print "               It happens over the tab bar, so there is no real harm."
    Debug.Print ""
    Debug.Print "  --- (2) *The scroll position must not return on its own* ---"
    Debug.Print "  6) Scroll to near the right end and leave it"
    Debug.Print "       (if a tab is loading, a title-update sync is sent when it finishes)"
    Debug.Print "       -> Expected: even if PushSyncToJs appears in the log, *the displayed position does not move*"
    Debug.Print "       -> Up to 9.32 it *went back to the left edge* here (the point fixed this time)"
    Debug.Print "  7) While looking at the right side, do something that changes the title of one of the tabs"
    Debug.Print "       (e.g. make a visible tab load a page)"
    Debug.Print "       -> Expected: the position does not move. It is not pulled back to the active tab's position"
    Debug.Print ""
    Debug.Print "  --- (3) The new tab must be visible ---"
    Debug.Print "  8) With the view returned to the left end, press +"
    Debug.Print "       -> Expected: the new tab is attached at the right end, but *it automatically moves to the right end and is visible*"
    Debug.Print "  9) Conversely, while looking at the right end, click a tab on the left"
    Debug.Print "       -> Expected: that tab is already visible, so *it does not move much*"
    Debug.Print "               (it is nearest, so things already visible are not shifted)"
    Debug.Print " 10) While looking at the right end, do something like Test_9_27_Status that switches to a left tab"
    Debug.Print "       -> Expected: when the active tab changes, the view moves into the visible range"
    Debug.Print ""
    Debug.Print "  --- (4) Regression ---"
    Debug.Print " 11) Is the feel of tab switching the same as 9.32? (the color changes immediately on click)"
    Debug.Print " 12) Does D&D reordering work? Does the position not jump after dragging?"
    Debug.Print " 13) Can tabs be closed with x? Does the scroll position not break after closing?"
    Debug.Print " 14) Does the settings tab open with [gear]?"
    Debug.Print ""
    Debug.Print "  --- Verdict ---"
    Debug.Print "  - If 3-4 work, issue 1 succeeded (if not, return issue 1 to (b) the 3px bar)"
    Debug.Print "  - If the position does not move in 6-7, the save/restore of issue 3 succeeded"
    Debug.Print "  - If the new tab is visible in 8, the nearest shift also succeeded"
    Debug.Print "  - If the scroll amount is too fast / slow, it can be adjusted with a single number"
End Sub


' ============================================================
' *Verification of D-1*
' ============================================================


' ============================================================
' Test_D1_Eval (D-1)
'   Runs through EvalSync and lines up the results.
'   Precondition: the browser has been started with StartWebView2_Full.
'         The active tab is showing a real page (most pass even with about:blank).
'   *Do not break/step through during the DoEvents loop (spec fact 20)*
' ============================================================
Public Sub Test_D1_Eval()
    Dim p As Wv2Pane
    Set p = UserForm1.GetActivePane

    If p Is Nothing Then
        Wv2Log.LogI "Test_D1_Eval: there is no active Pane." & _
                    "Run StartWebView2_Full first."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D1_Eval start ================"
    Wv2Log.LogI "  target tab: " & p.DocumentTitle
    Wv2Log.LogI "  in-callback depth (expected 0): " & p.InCallbackDepth
    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (1) Return value for each type ---"

    D1Case p, "integer", "1 + 2", 5, True
    D1Case p, "decimal", "1 / 4", 5, True
    D1Case p, "string (returned with quotes is normal)", "'abc'", 5, True
    D1Case p, "Japanese (checks the decoding of uXXXX)", ("'" & ChrW(&H65E5&) & ChrW(&H672C&) & ChrW(&H8A9E&) & ChrW(&H30C6&) & ChrW(&H30B9&) & ChrW(&H30C8&) & "'"), 5, True
    D1Case p, "string containing a quote", "'[' + String.fromCharCode(34) + ']'", 5, True
    D1Case p, "string containing a backslash", "'[' + String.fromCharCode(92) + ']'", 5, True
    D1Case p, "boolean", "1 === 1", 5, True
    D1Case p, "null", "null", 5, True
    D1Case p, "undefined (the case where the value key disappears)", "undefined", 5, True
    D1Case p, "object", "({a:1,b:[1,2,3]})", 5, True
    D1Case p, "array", "[1,'x',null]", 5, True
    D1Case p, "long string (300 chars)", _
              "(function(){var s='';for(var i=0;i<300;i++){s+='x';}return s;})()", 5, True

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) Information of the real page ---"
    D1Case p, "document.title", "document.title", 5, True
    D1Case p, "location.href", "document.location.href", 5, True
    D1Case p, "number of characters in body", "document.body.innerText.length", 5, True

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) Failure cases (normal if OK is shown rather than FAIL) ---"
    D1Case p, "JS exception (ReferenceError)", "nonexistentFunctionForTest()", 5, False
    D1Case p, "JS exception (throw wrapped in an IIFE)", _
              "(function(){throw new Error('boom');})()", 5, False
    D1Case p, "syntax error (the expression is broken)", "1 +", 5, False

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) Timeout and recovery ---"
    Wv2Log.LogI "  NOTE: block JS for 3 seconds and cut it off at 1 second. A few seconds later,"
    Wv2Log.LogI "     if the late-arrival log of *discarded* appears, issue 5 succeeded."
    D1Case p, "timeout (3 s cut off at 1 s)", _
              "(function(){var t=Date.now();while(Date.now()-t<3000){}return 1;})()", 1, False

    ' Wait until the JS thread is free (the late-arrival log appears during this time)
    Dim t0 As Single
    t0 = Timer
    Do While (Timer - t0) < 4
        DoEvents
    Loop

    D1Case p, "recovery after the timeout", "1 + 1", 5, True

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback depth (expected 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D1_Eval end ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' D1Case (run one case for D-1 verification)
'   expectOk = True  ... a case that should succeed
'   expectOk = False ... a case that should fail (OK if an error reason appears)
' ============================================================
Private Sub D1Case(ByVal p As Wv2Pane, _
                   ByVal caseName As String, _
                   ByVal expr As String, _
                   ByVal toSec As Single, _
                   ByVal expectOk As Boolean)
    Dim r As String
    Dim mark As String

    r = p.EvalSync(expr, toSec)

    If p.LastEvalOk = expectOk Then
        mark = "  [OK  ] "
        m_okCount = m_okCount + 1
    Else
        mark = "  [FAIL] "
        m_ngCount = m_ngCount + 1
    End If

    If p.LastEvalOk Then
        Wv2Log.LogI mark & caseName & " -> " & Left$(r, 100)
    Else
        Wv2Log.LogI mark & caseName & " -> failed: " & p.LastEvalError
    End If
End Sub


' ============================================================
' Test_D1_Guard (D-1 issue 7)
'   Confirms that the in-callback guard works and that ResetCallbackGuard can recover.
'   The route of calling from inside an actual event handler is hard to reproduce,
'   so we make the same state with Debug_SetInCallback and check.
' ============================================================
Public Sub Test_D1_Guard()
    Dim p As Wv2Pane
    Dim r As String

    Set p = UserForm1.GetActivePane
    If p Is Nothing Then
        Wv2Log.LogI "Test_D1_Guard: there is no active Pane." & _
                    "Run StartWebView2_Full first."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D1_Guard start ================"

    Wv2Log.LogI "  1) EvalSync in the normal state (depth=" & p.InCallbackDepth & ")"
    r = p.EvalSync("1 + 1", 5)
    TestBool "  2 is returned in the normal state", (p.LastEvalOk And r = "2")
    If Not p.LastEvalOk Then Wv2Log.LogI "         r=" & r & " err=" & p.LastEvalError

    Wv2Log.LogI "  2) Create a state of being inside a handler and EvalSync (being rejected is normal)"
    p.Debug_SetInCallback 1
    r = p.EvalSync("1 + 1", 5)
    TestBool "  *rejected in-callback (returned at once without hanging)*", _
             ((Not p.LastEvalOk) And p.LastEvalError = "in-callback")

    Wv2Log.LogI "  3) Recover with ResetCallbackGuard and run again"
    p.ResetCallbackGuard
    r = p.EvalSync("1 + 1", 5)
    TestBool "  recovers with ResetCallbackGuard", (p.LastEvalOk And r = "2")
    Wv2Log.LogI "         depth=" & p.InCallbackDepth

    TestCountPrint
    Wv2Log.LogI "================ Test_D1_Guard end ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' Test_D1_Help (procedure for D-1)
' ============================================================
Public Sub Test_D1_Help()
    Debug.Print ""
    Debug.Print "=========================================================="
    Debug.Print " D-1 verification procedure (EvalSync = synchronous retrieval from ExecuteScript)"
    Debug.Print "=========================================================="
    Debug.Print ""
    Debug.Print "  --- Preparation ---"
    Debug.Print "  1) Run UserForm1.Show vbModeless, then run StartWebView2_Full"
    Debug.Print "     *Show comes first* Without the form's window the HWND of Frame1 cannot"
    Debug.Print "     be obtained, and Browser.Init fails with hWnd_Frame still 0."
    Debug.Print "  2) Open any real page (Google etc. is fine)"
    Debug.Print "  3) *Wait until the event burst settles* (spec fact 20)"
    Debug.Print "     Go on only after the Immediate window log has stopped."
    Debug.Print ""
    Debug.Print "  --- Run ---"
    Debug.Print "  4) Test_D1_Eval  ... runs types, Japanese, exceptions and timeouts in one batch"
    Debug.Print "  5) Test_D1_Guard ... checks that the in-callback guard fires"
    Debug.Print ""
    Debug.Print "  *Do not break/step through during execution*"
    Debug.Print "    EvalSync waits by running DoEvents, so stopping in the middle of it"
    Debug.Print "    steps into the window of spec fact 20."
    Debug.Print ""
    Debug.Print "  --- *Verdicts remain in the log file* ---"
    Debug.Print "    The Immediate window is flushed away at once by the ExecuteScript plumbing log (15 lines per round trip),"
    Debug.Print "    so check pass/fail in the log file:"
    Debug.Print "      ?Wv2Log.LogPath   ... location of the file"
    Debug.Print "      Just look at the trailing ""*Verdicts n: OK x / FAIL y*"" to know pass/fail"
    Debug.Print ""
    Debug.Print "  --- What to look at (Test_D1_Eval) ---"
    Debug.Print "  - (1) All cases are [OK  ]"
    Debug.Print "      NOTE: strings are normally returned *with quotes* (stripped in D-2)"
    Debug.Print "      NOTE: the Japanese must not be garbled (evidence that uXXXX decoding works)"
    Debug.Print "      NOTE: for undefined, the string undefined is returned, which is normal"
    Debug.Print "  - (2) document.title and location.href match the actual page"
    Debug.Print "  - (3) The failure cases also show [OK  ] (meaning they failed as expected)"
    Debug.Print "      ReferenceError / Error: boom / syntax-error-or-null should be visible"
    Debug.Print "  - (4) The timeout *returns in about 1 second* (does not wait 5 seconds)"
    Debug.Print "      A few seconds later the late-arrival log of *discarded* should appear"
    Debug.Print "      The last recovery should be [OK  ]"
    Debug.Print "  - The final in-callback depth is 0"
    Debug.Print ""
    Debug.Print "  --- What to look at (Test_D1_Guard) ---"
    Debug.Print "  - 2) returns *at once* (does not hang). This is the purpose of issue 7"
    Debug.Print "  - 3) recovers, and the depth returns to 0"
    Debug.Print ""
    Debug.Print "  --- Regression check (because D-1 re-wrapped the callbacks of Wv2Pane) ---"
    Debug.Print "  6) Adding / switching / closing tabs and D&D reordering"
    Debug.Print "  7) URL input, back, forward, reload (the title must be updated)"
    Debug.Print "  8) New-tab expansion of links (including the POST link of Test_9_26_PostProbe)"
    Debug.Print "  9) The settings tab (gear) opens and the search engine can be changed"
    Debug.Print "     NOTE: 6-9 are for seeing the effect of separating View_On* into Core."
    Debug.Print "       If any one of them does not work, suspect the guard wrapper."
    Debug.Print ""
    Debug.Print "  --- If you want to try by hand ---"
    Debug.Print "  ?UserForm1.GetActivePane.EvalSync(""document.title"")"
    Debug.Print "  ?UserForm1.GetActivePane.EvalSync(""document.querySelectorAll('a').length"")"
    Debug.Print "  ?UserForm1.GetActivePane.LastEvalError"
    Debug.Print ""
End Sub


' ============================================================
' Test_D2_Find (verification of the D-2 main part)
'   Opens the verification page in a new tab and runs element acquisition and reading in one batch.
' ============================================================
Public Sub Test_D2_Find()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim el As Wv2Element
    Dim el2 As Wv2Element

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D2_Find: Browser is not started." & _
                    "Run UserForm1.StartWebView2_Full first."
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD2ProbeHtml())
    If p Is Nothing Then
        Wv2Log.LogI "Test_D2_Find: failed to create the tab."
        Exit Sub
    End If

    If Not D2WaitTitle(p, ("D-2 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&)), 10) Then
        Wv2Log.LogI "Test_D2_Find: could not confirm that the verification page loaded. Aborting."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D2_Find start ================"
    Wv2Log.LogI "  generation (an empty string before acquisition is normal): [" & p.CurrentDomGen & "]"
    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (1) Can be acquired ---"

    Set el = p.GetElementById("ttl")
    TestBool "GetElementById(ttl) is not Nothing", Not (el Is Nothing)
    If el Is Nothing Then
        Wv2Log.LogI "  Cannot continue the following verification. Aborting."
        TestCountPrint
        Exit Sub
    End If
    Wv2Log.LogI "        handle=" & el.Handle & " gen=" & el.Generation
    Wv2Log.LogI "        Pane-side generation cache: " & p.CurrentDomGen
    TestBool "not stale right after acquisition", (el.IsStale = False)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) 4 kinds of reads + attributes ---"
    TestEq "TagName (returned in uppercase)", el, el.TagName, "H1"
    TestEq "InnerText (Japanese)", el, el.InnerText, ("D-2 " & ChrW(&H8981&) & ChrW(&H7D20&) & ChrW(&H30EC&) & ChrW(&H30B8&) & ChrW(&H30B9&) & ChrW(&H30C8&) & ChrW(&H30EA&) & ChrW(&H306E&) & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&))

    Set el = D2El(p, "box")
    TestEq "GetAttribute(class)", el, el.GetAttribute("class"), "card"
    TestEq "GetAttribute(data-note) Japanese attribute", el, _
         el.GetAttribute("data-note"), (ChrW(&H5C5E&) & ChrW(&H6027&) & ChrW(&H306E&) & ChrW(&H5024&) & " (" & ChrW(&H65E5&) & ChrW(&H672C&) & ChrW(&H8A9E&) & ")")
    TestEq "InnerHTML (*decoding of spec fact 30*)", el, _
         el.InnerHTML, ("<span class=""tag"">" & ChrW(&H5185&) & ChrW(&H5074&) & "</span>" & ChrW(&H30C6&) & ChrW(&H30AD&) & ChrW(&H30B9&) & ChrW(&H30C8&))

    Set el = D2El(p, "esc")
    TestEq "mix of symbols", el, el.InnerText, (ChrW(&H8A18&) & ChrW(&H53F7&) & ": < > & "" ' \ " & ChrW(&H306E&) & ChrW(&H6DF7&) & ChrW(&H5728&))

    Set el = D2El(p, "pre")
    TestEq "text containing line breaks (decoding of \n)", el, el.InnerText, ("1 " & ChrW(&H884C&) & ChrW(&H76EE&)) & vbLf & ("2 " & ChrW(&H884C&) & ChrW(&H76EE&))

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) value of input elements ---"
    Set el = D2El(p, "txt")
    TestEq "TagName of input", el, el.TagName, "INPUT"
    TestEq "Value of input", el, el.value, (ChrW(&H521D&) & ChrW(&H671F&) & ChrW(&H5024&))
    TestEq "GetAttribute(value) of input", el, el.GetAttribute("value"), (ChrW(&H521D&) & ChrW(&H671F&) & ChrW(&H5024&))

    Set el = D2El(p, "area")
    TestEq "Value of textarea", el, el.value, (ChrW(&H30C6&) & ChrW(&H30AD&) & ChrW(&H30B9&) & ChrW(&H30C8&) & ChrW(&H30A8&) & ChrW(&H30EA&) & ChrW(&H30A2&) & ChrW(&H306E&) & ChrW(&H5024&))

    Set el = D2El(p, "sel")
    TestEq "Value of select (the selected option)", el, el.value, "b"

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) Empty / absent must *succeed with an empty string* (LastOk=True) ---"
    Set el = D2El(p, "empty")
    TestEq "InnerText of an empty element", el, el.InnerText, ""

    Set el = D2El(p, "lnk")
    TestEq "Value of an a element (has no value)", el, el.value, ""
    TestEq "nonexistent attribute", el, el.GetAttribute("data-nothing"), ""
    TestEq "href attribute", el, el.GetAttribute("href"), "https://example.com/path?x=1"

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) QuerySelector (*single quotes inside the selector*) ---"
    Set el2 = p.QuerySelector("input[name='q']")
    TestBool "QuerySelector(input[name='q']) can be obtained", Not (el2 Is Nothing)
    If Not el2 Is Nothing Then
        TestEq "the same element was obtained", el2, el2.value, (ChrW(&H521D&) & ChrW(&H671F&) & ChrW(&H5024&))
    End If

    Set el2 = p.QuerySelector("#box .tag")
    TestBool "descendant selector works", Not (el2 Is Nothing)
    If Not el2 Is Nothing Then
        TestEq "InnerText of the descendant selector", el2, el2.InnerText, (ChrW(&H5185&) & ChrW(&H5074&))
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) Distinguishing not found / failure (issue 4) ---"
    Set el2 = p.QuerySelector("#nothing-here")
    TestBool "nonexistent selector -> Nothing", (el2 Is Nothing)
    TestBool "  and LastEvalOk = True (meaning it really is absent)", (p.LastEvalOk = True)

    Set el2 = p.QuerySelector("###")
    TestBool "invalid selector -> Nothing", (el2 Is Nothing)
    TestBool "  and LastEvalOk = False (meaning it failed)", (p.LastEvalOk = False)
    Wv2Log.LogI "        LastEvalError = " & p.LastEvalError

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (7) ClearElementRegistry (issue 3) ---"
    Set el = D2El(p, "ttl")
    TestBool "not stale before the cleanup", (el.IsStale = False)
    TestBool "ClearElementRegistry succeeds", p.ClearElementRegistry()
    TestBool "becomes stale after the cleanup", (el.IsStale = True)
    Set el = p.GetElementById("ttl")
    TestBool "can be re-acquired even after the cleanup", Not (el Is Nothing)
    If Not el Is Nothing Then
        TestEq "the re-acquired element can be read", el, el.TagName, "H1"
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback depth (expected 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D2_Find end ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' Test_D2_Stale (verification of D-2 issue 2: the generation changes on page navigation)
' ============================================================
Public Sub Test_D2_Stale()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim el As Wv2Element
    Dim genBefore As String
    Dim genAfter As String
    Dim v As String

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D2_Stale: Browser is not started." & _
                    "Run UserForm1.StartWebView2_Full first."
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD2ProbeHtml())
    If p Is Nothing Then
        Wv2Log.LogI "Test_D2_Stale: failed to create the tab."
        Exit Sub
    End If
    If Not D2WaitTitle(p, ("D-2 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&)), 10) Then
        Wv2Log.LogI "Test_D2_Stale: could not confirm that the verification page loaded. Aborting."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D2_Stale start ================"

    Set el = p.GetElementById("ttl")
    If el Is Nothing Then
        Wv2Log.LogI "  [FAIL] cannot obtain the element, so aborting."
        m_ngCount = m_ngCount + 1
        TestCountPrint
        Exit Sub
    End If
    genBefore = el.Generation
    Wv2Log.LogI "  1) Before navigation: handle=" & el.Handle & " gen=" & genBefore
    TestEq "     readable", el, el.TagName, "H1"
    TestBool "     not stale", (el.IsStale = False)

    Wv2Log.LogI "  2) Navigate the same tab to another page"
    p.View_NavigateToString BuildD2SecondHtml()
    If Not D2WaitTitle(p, ("D-2 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&) & " 2 " & ChrW(&H679A&) & ChrW(&H76EE&)), 10) Then
        Wv2Log.LogI "  [FAIL] could not confirm that the 2nd page loaded. Aborting."
        Exit Sub
    End If

    Wv2Log.LogI "  3) After navigation: look at the state of the old handle"
    TestBool "     IsStale = True", (el.IsStale = True)
    v = el.TagName
    TestBool "     read gives an empty string + LastOk=False", (Len(v) = 0 And el.LastOk = False)
    Wv2Log.LogI "        LastError = " & el.LastError
    TestBool "     LastError is stale", (el.LastError = "stale")

    Wv2Log.LogI "  4) Can be re-acquired on the new page"
    Set el = p.GetElementById("second")
    TestBool "     can be acquired", Not (el Is Nothing)
    If Not el Is Nothing Then
        genAfter = el.Generation
        TestEq "     readable", el, el.InnerText, ("2 " & ChrW(&H679A&) & ChrW(&H76EE&) & ChrW(&H306E&) & ChrW(&H30DA&) & ChrW(&H30FC&) & ChrW(&H30B8&))
        Wv2Log.LogI "        new gen=" & genAfter
        TestBool "     the generation has changed", (genAfter <> genBefore)
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback depth (expected 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D2_Stale end ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' D2El (D-2: guarded helper for acquiring an element)
'
'   If it cannot be obtained, it emits a FAIL and returns *a dummy with no Pane set*.
'   The dummy merely returns LastOk=False / LastError="no-pane" when read, so
'   the verification runs through to the end without stopping at runtime error 91.
'   (It directly uses the route taken when a Wv2Element is plainly New-ed)
' ============================================================
Private Function D2El(ByVal p As Wv2Pane, ByVal elementId As String) As Wv2Element
    Dim e As Wv2Element

    Set e = p.GetElementById(elementId)
    If e Is Nothing Then
        Wv2Log.LogI "  [FAIL] element #" & elementId & " cannot be obtained " & _
                    "(LastEvalOk=" & p.LastEvalOk & " err=" & p.LastEvalError & ")"
        Set e = New Wv2Element
    End If

    Set D2El = e
End Function


' ============================================================
' D2WaitTitle (D-2: wait for the page load to finish by its title)
'
'   NavigateToString is asynchronous, so we read document.title with EvalSync and
'   wait until it becomes the expected title. IsPageLoaded is not used because
'   in Test_D2_Stale, which navigates the same Pane twice, it cannot be
'   told apart from "the completion of the previous page".
'
'   *Do not break/step through during this wait loop either* (spec fact 20)
' ============================================================
Private Function D2WaitTitle(ByVal p As Wv2Pane, _
                             ByVal wantTitle As String, _
                             ByVal timeoutSec As Single) As Boolean
    Dim t0 As Single
    Dim res As String
    Dim cur As String

    D2WaitTitle = False
    t0 = Timer

    Do
        DoEvents
        res = p.EvalSync("document.title", 3)
        If p.LastEvalOk Then
            cur = Wv2Json.JsonUnescape(res)
            If cur = wantTitle Then
                D2WaitTitle = True
                Exit Function
            End If
        End If
        If (Timer - t0) > timeoutSec Then
            Wv2Log.LogI "D2WaitTitle: timeout (expected=" & wantTitle & _
                        " actual=" & cur & " err=" & p.LastEvalError & ")"
            Exit Function
        End If
    Loop
End Function


' ============================================================
' BuildD2ProbeHtml (the verification page of D-2, issue 8 plan b)
'
'   *Quotation marks in the static HTML part are written with VBA's ""* (it is not JS, so no problem)
'   *No JS is used at all* (this page only provides a DOM)
' ============================================================
Private Function BuildD2ProbeHtml() As String
    Dim s As String

    s = "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    s = s & ("<title>D-2 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&) & "</title>") & vbLf
    s = s & "<style>" & vbLf
    s = s & "  body{font-family:'Segoe UI','Meiryo',sans-serif;background:#12161f;" & _
            "color:#e8eaed;padding:32px;line-height:1.8;}" & vbLf
    s = s & "  h1{font-size:22px;margin:0 0 18px;}" & vbLf
    s = s & "  .card{border:1px solid rgba(255,255,255,.12);border-radius:10px;" & _
            "padding:14px 16px;margin:12px 0;background:rgba(255,255,255,.04);}" & vbLf
    s = s & "  .tag{color:#6ea8fe;font-weight:600;}" & vbLf
    s = s & "  pre{background:#0b0e15;padding:10px 12px;border-radius:8px;margin:12px 0;}" & vbLf
    s = s & "  input,textarea,select{font-size:14px;padding:6px 8px;margin:4px 0;" & _
            "background:#0b0e15;color:#e8eaed;border:1px solid rgba(255,255,255,.18);" & _
            "border-radius:6px;}" & vbLf
    s = s & "  .note{color:#8ea2c8;font-size:12.5px;}" & vbLf
    s = s & "</style></head><body>" & vbLf
    s = s & ("<h1 id=""ttl"">D-2 " & ChrW(&H8981&) & ChrW(&H7D20&) & ChrW(&H30EC&) & ChrW(&H30B8&) & ChrW(&H30B9&) & ChrW(&H30C8&) & ChrW(&H30EA&) & ChrW(&H306E&) & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&) & "</h1>") & vbLf
    s = s & ("<p class=""note"">" & ChrW(&H3053&) & ChrW(&H306E&) & ChrW(&H30DA&) & ChrW(&H30FC&) & ChrW(&H30B8&) & ChrW(&H306F&) & " Test_D2_Find / Test_D2_Stale " & ChrW(&H5C02&) & ChrW(&H7528&) & ChrW(&H3067&) & ChrW(&H3059&) & ChrW(&H3002&)) & _
            (ChrW(&H5916&) & ChrW(&H90E8&) & ChrW(&H30B5&) & ChrW(&H30A4&) & ChrW(&H30C8&) & ChrW(&H306B&) & ChrW(&H4F9D&) & ChrW(&H5B58&) & ChrW(&H3057&) & ChrW(&H307E&) & ChrW(&H305B&) & ChrW(&H3093&) & ChrW(&H3002&) & "</p>") & vbLf
    s = s & ("<div id=""box"" class=""card"" data-note=""" & ChrW(&H5C5E&) & ChrW(&H6027&) & ChrW(&H306E&) & ChrW(&H5024&) & " (" & ChrW(&H65E5&) & ChrW(&H672C&) & ChrW(&H8A9E&) & ")"">") & _
            ("<span class=""tag"">" & ChrW(&H5185&) & ChrW(&H5074&) & "</span>" & ChrW(&H30C6&) & ChrW(&H30AD&) & ChrW(&H30B9&) & ChrW(&H30C8&) & "</div>") & vbLf
    s = s & ("<p id=""esc"">" & ChrW(&H8A18&) & ChrW(&H53F7&) & ": &lt; &gt; &amp; &quot; ' \ " & ChrW(&H306E&) & ChrW(&H6DF7&) & ChrW(&H5728&) & "</p>") & vbLf
    s = s & ("<pre id=""pre"">1 " & ChrW(&H884C&) & ChrW(&H76EE&)) & vbLf & ("2 " & ChrW(&H884C&) & ChrW(&H76EE&) & "</pre>") & vbLf
    s = s & "<div class=""card"">" & vbLf
    s = s & ("  <input id=""txt"" type=""text"" name=""q"" value=""" & ChrW(&H521D&) & ChrW(&H671F&) & ChrW(&H5024&) & """>") & vbLf
    s = s & ("  <textarea id=""area"" rows=""2"">" & ChrW(&H30C6&) & ChrW(&H30AD&) & ChrW(&H30B9&) & ChrW(&H30C8&) & ChrW(&H30A8&) & ChrW(&H30EA&) & ChrW(&H30A2&) & ChrW(&H306E&) & ChrW(&H5024&) & "</textarea>") & vbLf
    s = s & "  <select id=""sel""><option value=""a"">A</option>" & _
            "<option value=""b"" selected>B</option></select>" & vbLf
    s = s & "</div>" & vbLf
    s = s & ("<p><a id=""lnk"" href=""https://example.com/path?x=1"" title=""" & ChrW(&H30EA&) & ChrW(&H30F3&) & ChrW(&H30AF&) & """>") & _
            (ChrW(&H30EA&) & ChrW(&H30F3&) & ChrW(&H30AF&) & "</a></p>") & vbLf
    s = s & "<div id=""empty""></div>" & vbLf
    s = s & "</body></html>"

    BuildD2ProbeHtml = s
End Function


' ============================================================
' BuildD2SecondHtml (D-2: the 2nd page navigated to. Only to see the generation change)
' ============================================================
Private Function BuildD2SecondHtml() As String
    Dim s As String

    s = "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    s = s & ("<title>D-2 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&) & " 2 " & ChrW(&H679A&) & ChrW(&H76EE&) & "</title>") & vbLf
    s = s & "<style>body{font-family:'Segoe UI','Meiryo',sans-serif;background:#141b14;" & _
            "color:#e8eaed;padding:32px;}</style></head><body>" & vbLf
    s = s & ("<h1 id=""second"">2 " & ChrW(&H679A&) & ChrW(&H76EE&) & ChrW(&H306E&) & ChrW(&H30DA&) & ChrW(&H30FC&) & ChrW(&H30B8&) & "</h1>") & vbLf
    s = s & ("<p>" & ChrW(&H3053&) & ChrW(&H3053&) & ChrW(&H306B&) & ChrW(&H6765&) & ChrW(&H305F&) & ChrW(&H3089&) & ChrW(&H3001&) & "1 " & ChrW(&H679A&) & ChrW(&H76EE&) & ChrW(&H3067&) & ChrW(&H53D6&) & ChrW(&H3063&) & ChrW(&H305F&) & " Wv2Element " & ChrW(&H306F&) & ChrW(&H3059&) & ChrW(&H3079&) & ChrW(&H3066&) & " stale " & ChrW(&H306B&) & ChrW(&H306A&) & ChrW(&H308A&) & ChrW(&H307E&) & ChrW(&H3059&) & ChrW(&H3002&) & "</p>") & vbLf
    s = s & "</body></html>"

    BuildD2SecondHtml = s
End Function


' ============================================================
' Test_D2_Help (procedure for D-2)
' ============================================================
Public Sub Test_D2_Help()
    Debug.Print ""
    Debug.Print "=========================================================="
    Debug.Print " D-2 verification procedure (element registry + reading of Wv2Element)"
    Debug.Print "=========================================================="
    Debug.Print ""
    Debug.Print "  --- Preparation ---"
    Debug.Print "  1) Run UserForm1.Show vbModeless, then run StartWebView2_Full"
    Debug.Print "     *Show comes first* Without the form's window the HWND of Frame1 cannot"
    Debug.Print "     be obtained, and Browser.Init fails with hWnd_Frame still 0."
    Debug.Print "  2) *Wait until the event burst settles* (spec fact 20)"
    Debug.Print "     Go on only after the Immediate window log has stopped."
    Debug.Print "     NOTE: the verification page is its own HTML, so there is no need to open an external site."
    Debug.Print ""
    Debug.Print "  --- Run ---"
    Debug.Print "  3) Test_D2_Find  ... runs acquisition, reading, absence and cleanup in one batch"
    Debug.Print "  4) Test_D2_Stale ... confirms that the generation changes on page navigation"
    Debug.Print "     Both *open one new tab* (you may close it by hand after running)"
    Debug.Print ""
    Debug.Print "  *Do not break/step through during execution*"
    Debug.Print "    EvalSync waits by running DoEvents, so stopping in the middle of it"
    Debug.Print "    steps into the window of spec fact 20."
    Debug.Print ""
    Debug.Print "  --- *Verdicts remain in the log file* ---"
    Debug.Print "    The Immediate window is flushed away at once by the ExecuteScript plumbing log (15 lines per round trip),"
    Debug.Print "    so check pass/fail in the log file:"
    Debug.Print "      ?Wv2Log.LogPath   ... location of the file"
    Debug.Print "      Just look at the trailing ""*Verdicts n: OK x / FAIL y*"" to know pass/fail"
    Debug.Print ""
    Debug.Print "  --- What to look at (Test_D2_Find) ---"
    Debug.Print "  - All lines of (1)-(7) are [OK  ]"
    Debug.Print "  - The InnerHTML of (2) comes out *correctly including the tags*"
    Debug.Print "      Evidence that the decoding of spec fact 30 (< arrives as <) works."
    Debug.Print "      If this FAILs, suspect Wv2Json.JsonUnescapeAt."
    Debug.Print "  - (4) is [OK  ] = 'empty' and 'failure' can be told apart"
    Debug.Print "  - The 2 lines of (6) are [OK  ] *as a pair* (the convention of issue 4):"
    Debug.Print "      nonexistent  -> Nothing + LastEvalOk=True"
    Debug.Print "      invalid specification  -> Nothing + LastEvalOk=False"
    Debug.Print "  - The final in-callback depth is 0"
    Debug.Print ""
    Debug.Print "  --- What to look at (Test_D2_Stale) ---"
    Debug.Print "  - In 3) IsStale=True, and reading gives an empty string + LastError=stale"
    Debug.Print "  - In 4) an element of the new generation can be obtained (gen differs from before)"
    Debug.Print "    This is the heart of issue 2 (the generation is issued by JS and always changes on navigation)."
    Debug.Print ""
    Debug.Print "  --- Regression check (because D-2 touched Wv2Pane and Wv2Json) ---"
    Debug.Print "  5) Test_D1_Eval  ... EvalSync must not be broken"
    Debug.Print "     *JsonPickStr was moved to Wv2Json, so always run this*"
    Debug.Print "  6) Test_D1_Guard ... the in-callback guard works"
    Debug.Print "  7) Adding / switching / closing tabs and D&D reordering"
    Debug.Print "  8) The settings tab (gear) opens and the search engine can be changed"
    Debug.Print "     NOTE: Wv2Json is also used by TabBar / NavBar / SettingsBridge."
    Debug.Print ""
    Debug.Print "  --- If you want to try by hand ---"
    Debug.Print "  Set p = UserForm1.GetActivePane"
    Debug.Print "  Set el = p.QuerySelector(""input[name='q']"")"
    Debug.Print "  ?el.TagName"
    Debug.Print "  ?el.Value"
    Debug.Print "  ?el.IsStale"
    Debug.Print "  ?p.CurrentDomGen"
    Debug.Print ""
End Sub


' ============================================================
' Test_K1_Log  (Stage K-1)
'
'   Verifies Wv2Log in one go from the Immediate window.
'   Works without starting the browser (does not depend on WebView2).
' ============================================================
Public Sub Test_K1_Log()
    Dim k1Total As Long
    Dim k1Pass As Long

    Debug.Print String$(64, "=")
    Debug.Print "K-1 verification: Wv2Log (file output of the debug log)"
    Debug.Print String$(64, "=")

    ' --- Preparation: open a new log ---
    Wv2Log.LogLevel = LOG_DEBUG
    Wv2Log.LogEcho = True
    Wv2Log.LogStart

    Dim k1Path As String
    k1Path = Wv2Log.LogPath
    Debug.Print "  log: " & k1Path
    Debug.Print ""

    K1Bool "(1) LogPath can be obtained", (Len(k1Path) > 0), k1Total, k1Pass
    K1Bool "(2) the file really exists", (Len(Dir(k1Path)) > 0), k1Total, k1Pass

    ' --- write the 4 levels ---
    Wv2Log.LogE ("K-1: ERROR " & ChrW(&H306E&) & ChrW(&H884C&))
    Wv2Log.LogW ("K-1: WARN " & ChrW(&H306E&) & ChrW(&H884C&))
    Wv2Log.LogI ("K-1: INFO " & ChrW(&H306E&) & ChrW(&H884C&))
    Wv2Log.LogD ("K-1: DEBUG " & ChrW(&H306E&) & ChrW(&H884C&))

    ' --- does the threshold cut them off? ---
    Wv2Log.LogLevel = LOG_WARN
    Wv2Log.LogI ("K-1: " & ChrW(&H3053&) & ChrW(&H306E&) & " INFO " & ChrW(&H306F&) & ChrW(&H51FA&) & ChrW(&H306A&) & ChrW(&H3044&) & ChrW(&H306F&) & ChrW(&H305A&))
    Wv2Log.LogD ("K-1: " & ChrW(&H3053&) & ChrW(&H306E&) & " DEBUG " & ChrW(&H306F&) & ChrW(&H51FA&) & ChrW(&H306A&) & ChrW(&H3044&) & ChrW(&H306F&) & ChrW(&H305A&))
    Wv2Log.LogLevel = LOG_DEBUG
    Wv2Log.LogD ("K-1: " & ChrW(&H3057&) & ChrW(&H304D&) & ChrW(&H3044&) & ChrW(&H5024&) & ChrW(&H3092&) & ChrW(&H623B&) & ChrW(&H3057&) & ChrW(&H305F&))

    ' --- Unicode assembled at run time (the source stays CP932) ---
    Dim k1Emoji As String
    k1Emoji = ChrW$(&HD83D&) & ChrW$(&HDE00&)
    Wv2Log.LogD ("K-1: " & ChrW(&H30B5&) & ChrW(&H30ED&) & ChrW(&H30B2&) & ChrW(&H30FC&) & ChrW(&H30C8&) & ChrW(&H30DA&) & ChrW(&H30A2&) & " [") & k1Emoji & "]"
    Wv2Log.LogD ("K-1: " & ChrW(&H65E5&) & ChrW(&H672C&) & ChrW(&H8A9E&) & ChrW(&H3068&) & ChrW(&H8A18&) & ChrW(&H53F7&) & " " & ChrW(&H2605&) & " " & ChrW(&H2190&) & " " & ChrW(&H2192&) & " " & ChrW(&H2460&) & " " & ChrW(&HFF5E&) & " " & ChrW(&H2015&))

    ' --- re-entry guard: are they queued and flushed together later? ---
    Wv2Log.Debug_SetLogDepth 1
    Wv2Log.LogD ("K-1: " & ChrW(&H518D&) & ChrW(&H5165&) & ChrW(&H4E2D&) & ChrW(&H306E&) & " 1 " & ChrW(&H884C&) & ChrW(&H76EE&))
    Wv2Log.LogD ("K-1: " & ChrW(&H518D&) & ChrW(&H5165&) & ChrW(&H4E2D&) & ChrW(&H306E&) & " 2 " & ChrW(&H884C&) & ChrW(&H76EE&))
    Wv2Log.Debug_SetLogDepth 0
    Wv2Log.LogD ("K-1: " & ChrW(&H518D&) & ChrW(&H5165&) & ChrW(&H304B&) & ChrW(&H3089&) & ChrW(&H629C&) & ChrW(&H3051&) & ChrW(&H305F&))

    Wv2Log.LogFlush

    ' --- read back and compare ---
    '   *"Confirming absence" only means something once the file has been read back*
    '   If InStr(empty, x) = 0 passes even though nothing was read, it passes by a miss.
    '   We actually did exactly this in the K-1 real-machine run, so we compare only when the read succeeded.
    Dim k1Text As String

    ' (3) can it be read while running? (checks that it can be opened in an editor)
    '   *It cannot be measured with ADODB.Stream* ADODB cannot open a file
    '   that another handle has open, regardless of the share mode (measured in the K-1 real-machine run).
    '   A reader that states the share mode explicitly can open it, so we check with VBA's own
    '   Open ... For Binary Access Read Shared.
    Debug.Print ""
    K1Bool "(3) the log can be read even while running (shared read)", _
           K1CanReadWhileOpen(k1Path), k1Total, k1Pass

    ' Read the body with ADODB for comparison. It cannot be opened while open, so close it first
    Wv2Log.LogStop
    k1Text = K1ReadUtf8(k1Path)
    K1Bool "(4) the log could be read back", (Len(k1Text) > 0), k1Total, k1Pass

    If Len(k1Text) = 0 Then
        Debug.Print ""
        Debug.Print "  *The log cannot be read back, so the following comparisons are cut off*"
        Debug.Print "     (InStr on an empty string returns 0 for anything, so"
        Debug.Print "      continuing here would pass by a miss)"
    Else
        K1Bool "(5) the ERROR line is present", (InStr(k1Text, ("K-1: ERROR " & ChrW(&H306E&) & ChrW(&H884C&))) > 0), k1Total, k1Pass
        K1Bool "(6) the WARN line is present", (InStr(k1Text, ("K-1: WARN " & ChrW(&H306E&) & ChrW(&H884C&))) > 0), k1Total, k1Pass
        K1Bool "(7) the INFO line is present", (InStr(k1Text, ("K-1: INFO " & ChrW(&H306E&) & ChrW(&H884C&))) > 0), k1Total, k1Pass
        K1Bool "(8) the DEBUG line is present", (InStr(k1Text, ("K-1: DEBUG " & ChrW(&H306E&) & ChrW(&H884C&))) > 0), k1Total, k1Pass
        K1Bool "(9) no INFO that was cut by the threshold", (InStr(k1Text, (ChrW(&H3053&) & ChrW(&H306E&) & " INFO " & ChrW(&H306F&) & ChrW(&H51FA&) & ChrW(&H306A&) & ChrW(&H3044&) & ChrW(&H306F&) & ChrW(&H305A&))) = 0), k1Total, k1Pass
        K1Bool "(10) no DEBUG that was cut by the threshold", (InStr(k1Text, (ChrW(&H3053&) & ChrW(&H306E&) & " DEBUG " & ChrW(&H306F&) & ChrW(&H51FA&) & ChrW(&H306A&) & ChrW(&H3044&) & ChrW(&H306F&) & ChrW(&H305A&))) = 0), k1Total, k1Pass
        K1Bool "(11) the surrogate pair is intact", (InStr(k1Text, k1Emoji) > 0), k1Total, k1Pass
        K1Bool "(12) the Japanese is intact", (InStr(k1Text, (ChrW(&H65E5&) & ChrW(&H672C&) & ChrW(&H8A9E&) & ChrW(&H3068&) & ChrW(&H8A18&) & ChrW(&H53F7&) & " " & ChrW(&H2605&) & " " & ChrW(&H2190&) & " " & ChrW(&H2192&) & " " & ChrW(&H2460&) & " " & ChrW(&HFF5E&) & " " & ChrW(&H2015&))) > 0), k1Total, k1Pass
        K1Bool "(13) the 2 lines during re-entry flow out later", _
               (InStr(k1Text, ("K-1: " & ChrW(&H518D&) & ChrW(&H5165&) & ChrW(&H4E2D&) & ChrW(&H306E&) & " 1 " & ChrW(&H884C&) & ChrW(&H76EE&))) > InStr(k1Text, ("K-1: " & ChrW(&H518D&) & ChrW(&H5165&) & ChrW(&H304B&) & ChrW(&H3089&) & ChrW(&H629C&) & ChrW(&H3051&) & ChrW(&H305F&)))), _
               k1Total, k1Pass
    End If

    ' --- are the sequence numbers free of gaps? (the very motivation of K-1) ---
    Dim k1Lines As Variant
    Dim k1Idx As Long
    Dim k1Num As Long
    Dim k1Prev As Long
    Dim k1Ok As Boolean
    ' *Look at the presence of gaps, not the order*
    '   The deferred queue is flushed later in the re-entry guard test, so the order in the file
    '   is deliberately swapped. The motivation of K-1 is "not noticing that something was dropped", so
    '   what to look at is gaps in the sequence numbers, not the ordering.
    Dim k1Seq As Long
    k1Lines = Split(k1Text, vbCrLf)
    k1Prev = 0
    k1Seq = 0
    For k1Idx = LBound(k1Lines) To UBound(k1Lines)
        If Len(k1Lines(k1Idx)) >= 6 Then
            If IsNumeric(Left$(k1Lines(k1Idx), 6)) Then
                k1Num = CLng(Left$(k1Lines(k1Idx), 6))
                k1Seq = k1Seq + 1
                If k1Num > k1Prev Then k1Prev = k1Num
            End If
        End If
    Next k1Idx
    k1Ok = (k1Seq = k1Prev) And (k1Seq > 0)
    If Len(k1Text) > 0 Then
        K1Bool "(14) no gaps in the sequence numbers (" & k1Seq & " lines / max " & k1Prev & ")", _
               k1Ok, k1Total, k1Pass
    End If

    Debug.Print ""
    Debug.Print "  Result: " & k1Pass & " / " & k1Total & " passed"
    If k1Pass = k1Total Then
        Debug.Print "  *K-1 passed*"
    Else
        Debug.Print "  *There are failures. Look at the [FAIL] above*"
    End If
    Debug.Print "  actual log: " & k1Path
    Debug.Print String$(64, "=")
End Sub


Private Sub K1Bool(ByVal label As String, ByVal cond As Boolean, _
                   ByRef k1Total As Long, ByRef k1Pass As Long)
    k1Total = k1Total + 1
    If cond Then
        k1Pass = k1Pass + 1
        Debug.Print "  [OK  ] " & label
    Else
        Debug.Print "  [FAIL] " & label
    End If
End Sub


' Only looks at whether a file that is kept open can be read in shared mode
'   *ADODB.Stream cannot be used for this purpose* Regardless of the share mode
'   it cannot open a file that another handle has open (measured in K-1).
Private Function K1CanReadWhileOpen(ByVal k1Path As String) As Boolean
    On Error GoTo eh
    Dim k1Handle As Long
    Dim k1Bytes() As Byte
    k1Handle = FreeFile
    Open k1Path For Binary Access Read Shared As #k1Handle
    If LOF(k1Handle) > 0 Then
        ReDim k1Bytes(0 To LOF(k1Handle) - 1)
        Get #k1Handle, 1, k1Bytes
        K1CanReadWhileOpen = True
    End If
    Close #k1Handle
    Exit Function
eh:
    On Error Resume Next
    Close #k1Handle
End Function


' Read a UTF-8 / no-BOM file (verification only. ADODB is fine)
Private Function K1ReadUtf8(ByVal k1Path As String) As String
    On Error GoTo eh
    Dim k1Stream As Object
    Set k1Stream = CreateObject("ADODB.Stream")
    k1Stream.Type = 2
    k1Stream.Charset = "UTF-8"
    k1Stream.Open
    k1Stream.LoadFromFile k1Path
    K1ReadUtf8 = k1Stream.ReadText
    k1Stream.Close
    Exit Function
eh:
    Debug.Print "K1ReadUtf8: failed (" & Err.Number & ") " & Err.Description
End Function


' ============================================================
' Test_K1_Help  (Stage K-1)
' ============================================================
Public Sub Test_K1_Help()
    Debug.Print String$(64, "=")
    Debug.Print "How to use K-1 (file output of the debug log)"
    Debug.Print String$(64, "=")
    Debug.Print ""
    Debug.Print "  1) Test_K1_Log      ... runs the automatic verification in one go"
    Debug.Print "  2) ?Wv2Log.LogPath  ... location of the current log file"
    Debug.Print "  3) Wv2Log.LogStart  ... switches to a new log (a fresh start for verification)"
    Debug.Print ""
    Debug.Print "  *Verification procedure for the D axis*"
    Debug.Print "    If you type Wv2Log.LogStart and then run Test_D2_Find,"
    Debug.Print "    only that one test run is contained in one file."
    Debug.Print "    Even if the Immediate window scrolls away, everything remains in the file."
    Debug.Print ""
    Debug.Print "  *The 6 digits at the start of a line are a sequence number* If there is a gap, the log was dropped."
    Debug.Print "    This is the motivation for building K-1 (not noticing that something was dropped is dangerous)."
    Debug.Print ""
    Debug.Print "  Settings:"
    Debug.Print "    Wv2Log.LogLevel = LOG_ERROR / LOG_WARN / LOG_INFO / LOG_DEBUG"
    Debug.Print "    Wv2Log.LogEcho  = False   ... stops the duplicate output to the Immediate window"
    Debug.Print ""
    Debug.Print "  *The log of a running session can be opened in an editor* (because it is opened with Shared)."
    Debug.Print "    Only ADODB.Stream cannot open it. Regardless of the share mode,"
    Debug.Print "    it cannot open a file that another handle has open (measured in K-1)."
    Debug.Print ""
    Debug.Print "  One log per startup. When there are more than 20, the oldest are deleted first."
    Debug.Print "  Location: %APPDATA%\Wv2Browser\logs\"
    Debug.Print String$(64, "=")
End Sub


' ============================================================
' Test_K2_Help  (Stage K-2)
'
'   The settings tab needs a real start of WebView2, so it cannot be fully automated.
'   This outputs the procedure for checking on screen.
' ============================================================
Public Sub Test_K2_Help()
    Debug.Print String$(64, "=")
    Debug.Print "K-2 verification: the bug where the settings tab gets crushed by a search"
    Debug.Print String$(64, "=")
    Debug.Print ""
    Debug.Print "  *If you type Wv2Log.LogStart first, this one verification run is contained in one file*"
    Debug.Print "    The log location is ?Wv2Log.LogPath"
    Debug.Print ""
    Debug.Print "  [1] The fixed core"
    Debug.Print "    1) Start the browser"
    Debug.Print "    2) Press the gear at the right end of the tab bar to open the settings tab"
    Debug.Print "    3) With that settings tab still active, type an arbitrary search term into the address bar"
    Debug.Print "       -> It navigates to the search results (this is normal. The settings tab is also an ordinary tab)"
    Debug.Print "    4) *Press the gear once more*"
    Debug.Print "       -> *If a new settings tab opens, it is fixed*"
    Debug.Print "       -> If nothing happens / it only switches to the search-results tab, it is not fixed"
    Debug.Print ""
    Debug.Print "  [2] Duplicate prevention must not be broken (regression of 9.19)"
    Debug.Print "    5) With a settings tab open, mash the gear 3-4 times"
    Debug.Print "       -> There stays always exactly 1 settings tab. It must not increase"
    Debug.Print ""
    Debug.Print "  [3] It must also be fixed for routes other than the address bar"
    Debug.Print "    6) Navigate to a page with links from the settings tab, then gear -> a new settings tab opens"
    Debug.Print "    7) Even after reloading on the settings tab, the settings screen must be kept"
    Debug.Print ""
    Debug.Print "  [4] *Values must not diverge even with 2 settings views (the new feature this time)*"
    Debug.Print "    8) Open settings tab A with the gear"
    Debug.Print "    9) Navigate to an arbitrary page in A's address bar"
    Debug.Print "   10) Press the gear to open settings tab B"
    Debug.Print "   11) Switch to A and press *back* -> A also returns to the settings screen (now there are 2)"
    Debug.Print "   12) Switch to B and change the search engine to another one"
    Debug.Print "   13) *Switch to A*"
    Debug.Print "       -> *A's selection display and preview must also be the changed engine*"
    Debug.Print "       -> If A still has the old engine selected, the sync is not working"
    Debug.Print ""
    Debug.Print "  [5] Regression (nothing must be broken)"
    Debug.Print "   14) Change the engine on the settings tab, search in a normal tab, and confirm it is reflected"
    Debug.Print "   15) Entering and applying a custom URL still works as before"
    Debug.Print "       NOTE: by design the contents of the custom input box are not overwritten by the sync (to avoid breaking what is being typed)"
    Debug.Print ""
    Debug.Print "  What appears in the log:"
    Debug.Print "    Wv2Browser.OpenSettingsTab: the tab (index N) showing the settings screen is..."
    Debug.Print "    Wv2Browser.SyncSettingsViewsNow: reflected <engine> to N settings views"
    Debug.Print "    Wv2Browser.IsSettingsView: cannot determine so... (normal while loading, etc.)"
    Debug.Print String$(64, "=")
End Sub

' ============================================================
' Test_D3_Probe_Promise (the first move of D-3: *measuring unknown 1*)
'
'   Checks "does ExecuteScript wait for a Promise?". This result decides
'   issue 3 (whether the wait loop is run on the VBA side or the JS side).
'
'   *(B) cannot be judged automatically* The result of a raw ExecuteScript does not pass through EvalSync's
'   pending table, so VBA cannot receive it. Instead, look by eye at the lines of resultJson= that
'   Wv2Pane.OnExecuteScriptCompleted emits.
'   It is arranged so that the verdict can be made from *between which markers that line appeared*.
' ============================================================
Public Sub Test_D3_Probe_Promise()
    Dim p As Wv2Pane
    Dim res As String
    Dim cb As Long

    Set p = UserForm1.GetActivePane
    If p Is Nothing Then
        Wv2Log.LogI "Test_D3_Probe_Promise: there is no active Pane." & _
                    "Run UserForm1.StartWebView2_Full first."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D3_Probe_Promise start ================"
    Wv2Log.LogI "  *Measuring unknown 1: does ExecuteScript wait for a Promise?*"
    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (A) Via EvalSync (the synchronous primitive of D-1) ---"
    Wv2Log.LogI "      EvalSync fires JSON.stringify(expression), so even if the expression returns a Promise"
    Wv2Log.LogI "      it collapses to {} on the spot. This confirms that *by spec it is so*."

    res = p.EvalSync("1+1")
    Wv2Log.LogI "      1+1                 -> [" & res & "] ok=" & p.LastEvalOk
    res = p.EvalSync("Promise.resolve(42)")
    Wv2Log.LogI "      Promise.resolve(42) -> [" & res & "] ok=" & p.LastEvalOk
    TestBool "EvalSync does not resolve a Promise to a value (becomes {})", (res = "{}")

    res = p.EvalSync("(async function(){return 5;})()")
    Wv2Log.LogI "      return of an async function    -> [" & res & "] ok=" & p.LastEvalOk

    res = p.EvalSync("new Promise(function(r){setTimeout(function(){r(7);},300);})")
    Wv2Log.LogI "      delayed Promise        -> [" & res & "] ok=" & p.LastEvalOk

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (B) raw ExecuteScript (not wrapped in JSON.stringify) ---"
    Wv2Log.LogI "      *Look with your eyes here* Read the lines"
    Wv2Log.LogI "      'OnExecuteScriptCompleted: ... resultJson=' that appear between the markers."

    Wv2Log.LogI ""
    Wv2Log.LogI "  *** (B-1) Promise.resolve(42) starts here ***"
    cb = p.View_ExecuteScript("Promise.resolve(42)")
    D3Pump 2
    Wv2Log.LogI "  *** (B-1) ends here (callbackId=" & cb & ") ***"
    Wv2Log.LogI "      If resultJson={} it *does not wait* / if resultJson=42 it *waits*"

    Wv2Log.LogI ""
    Wv2Log.LogI "  *** (B-2) a Promise that never resolves starts here ***"
    cb = p.View_ExecuteScript("new Promise(function(r){})")
    D3Pump 2
    Wv2Log.LogI "  *** (B-2) ends here (callbackId=" & cb & ") ***"
    Wv2Log.LogI "      resultJson={} appeared in between -> *does not wait* (settled)"
    Wv2Log.LogI "      nothing appeared in between       -> *waits* (settled)"

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- Meaning of the verdict ---"
    Wv2Log.LogI "  - does not wait -> *poll on the VBA side* as in the skeleton of issue 3"
    Wv2Log.LogI "  - waits     -> the wait can be done on the JS side, so redesign starting from how EvalSync wraps"
    Wv2Log.LogI "                (the lesson of 9.30 / D-1: if the premise differs, replace the first move)"
    Wv2Log.LogI "  NOTE: the result of (A) does not change whichever way (B) goes. The current EvalSync"
    Wv2Log.LogI "     cannot wait for a Promise, so putting the wait on the JS side needs work on D-1's wrapping."

    Wv2Log.LogI "  in-callback depth (expected 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D3_Probe_Promise end ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' Test_D3_Write (verification of the D-3 main part: writes and operations)
' ============================================================
Public Sub Test_D3_Write()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim el As Wv2Element
    Dim dummy As Wv2Element
    Dim info As String
    Dim res As String

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D3_Write: Browser is not started." & _
                    "Run UserForm1.StartWebView2_Full first."
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD3ProbeHtml())
    If p Is Nothing Then
        Wv2Log.LogI "Test_D3_Write: failed to create the tab."
        Exit Sub
    End If
    If Not D2WaitTitle(p, ("D-3 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&)), 10) Then
        Wv2Log.LogI "Test_D3_Write: could not confirm that the verification page loaded. Aborting."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D3_Write start ================"
    Wv2Log.LogI "  --- (1) Writing to input ---"

    Set el = D2El(p, "txt")
    el.value = (ChrW(&H66F8&) & ChrW(&H304D&) & ChrW(&H8FBC&) & ChrW(&H3093&) & ChrW(&H3060&) & ChrW(&H5024&))
    info = el.LastInfo
    TestBool "Value = succeeds (LastOk)", el.LastOk
    Wv2Log.LogI "        route (LastInfo) = " & info & "  *setter is the expected value*"
    TestBool "  *goes through the native setter*", (info = "setter")
    TestEq "reading back gives the written value", el, el.value, (ChrW(&H66F8&) & ChrW(&H304D&) & ChrW(&H8FBC&) & ChrW(&H3093&) & ChrW(&H3060&) & ChrW(&H5024&))
    TestEq "*the value attribute does not change* (stays at the initial value)", el, _
         el.GetAttribute("value"), (ChrW(&H521D&) & ChrW(&H671F&) & ChrW(&H5024&))

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) Symbols and Japanese (checking JsQuote) ---"
    el.value = (ChrW(&H8A18&) & ChrW(&H53F7&) & ": < > & "" ' \ " & ChrW(&H306E&) & ChrW(&H6DF7&) & ChrW(&H5728&))
    TestBool "succeeds even with symbols mixed in", el.LastOk
    TestEq "the mix of symbols round-trips", el, el.value, _
         (ChrW(&H8A18&) & ChrW(&H53F7&) & ": < > & "" ' \ " & ChrW(&H306E&) & ChrW(&H6DF7&) & ChrW(&H5728&))
    el.value = ""
    TestEq "an empty string can be written", el, el.value, ""

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) Both input and change must fire (issue 6) ---"
    Wv2Log.LogI "        There have been 3 writes so far."
    TestBool "the input event fired 3 times", (D3Cnt(p, "inputs") = 3)
    TestBool "the change event fired 3 times", (D3Cnt(p, "changes") = 3)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) textarea / select / checkbox ---"
    Set el = D2El(p, "area")
    el.value = (ChrW(&H66F8&) & ChrW(&H304D&) & ChrW(&H63DB&) & ChrW(&H3048&) & ChrW(&H305F&) & ChrW(&H30C6&) & ChrW(&H30AD&) & ChrW(&H30B9&) & ChrW(&H30C8&))
    TestEq "can write to textarea", el, el.value, (ChrW(&H66F8&) & ChrW(&H304D&) & ChrW(&H63DB&) & ChrW(&H3048&) & ChrW(&H305F&) & ChrW(&H30C6&) & ChrW(&H30AD&) & ChrW(&H30B9&) & ChrW(&H30C8&))

    Set el = D2El(p, "sel")
    TestEq "initial value of select", el, el.value, "b"
    el.value = "a"
    TestEq "select can be switched", el, el.value, "a"

    Set el = D2El(p, "chk")
    TestBool "Click on the checkbox succeeds", el.Click()
    res = p.EvalSync("document.getElementById('chk').checked")
    TestBool "  checked became true", (res = "true")

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) Click on a button (issue 7) ---"
    Set el = D2El(p, "btn")
    TestBool "Click returns True", el.Click()
    Wv2Log.LogI "        route (LastInfo) = " & el.LastInfo
    TestBool "  *e.click() was used*", (el.LastInfo = "click")
    TestBool "  it was counted once on the page side", (D3Cnt(p, "clicks") = 1)
    Wv2Log.LogI "        identity of the event (type/bubbles/isTrusted) = " & _
                D3Str(p, "clickInfo")
    Wv2Log.LogI "        NOTE: isTrusted=false is *in principle so because it is a synthetic event*"

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) SetAttribute ---"
    Set el = D2El(p, "dv")
    TestBool "SetAttribute returns True", _
           el.SetAttribute("data-note", (ChrW(&H5C5E&) & ChrW(&H6027&) & ChrW(&H306E&) & ChrW(&H5024&) & " (" & ChrW(&H65E5&) & ChrW(&H672C&) & ChrW(&H8A9E&) & ")"))
    TestEq "the written attribute can be read", el, el.GetAttribute("data-note"), _
         (ChrW(&H5C5E&) & ChrW(&H6027&) & ChrW(&H306E&) & ChrW(&H5024&) & " (" & ChrW(&H65E5&) & ChrW(&H672C&) & ChrW(&H8A9E&) & ")")
    TestBool "an attribute value containing quotes can also be written", _
           el.SetAttribute("data-q", "a""b'c")
    TestEq "  round trip", el, el.GetAttribute("data-q"), "a""b'c"
    TestBool "*an invalid attribute name gives False*", (el.SetAttribute("", "x") = False)
    Wv2Log.LogI "        LastError = " & el.LastError

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (7) Writing to an element that has no value ---"
    Wv2Log.LogI "        div has no value, so it falls back to plain assignment."
    Wv2Log.LogI "        *In JS terms it is not an exception, so LastOk is True*"
    el.value = ("div " & ChrW(&H306B&) & ChrW(&H66F8&) & ChrW(&H3044&) & ChrW(&H3066&) & ChrW(&H307F&) & ChrW(&H308B&))
    info = el.LastInfo
    TestBool "LastOk is True (not an error)", el.LastOk
    TestBool "  the route is direct (there is no native setter)", (info = "direct")
    TestEq "  reading back shows the property that appeared", el, el.value, ("div " & ChrW(&H306B&) & ChrW(&H66F8&) & ChrW(&H3044&) & ChrW(&H3066&) & ChrW(&H307F&) & ChrW(&H308B&))

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (8) Distinguishing failures (following the D-2 convention, issue 8) ---"
    Set dummy = New Wv2Element
    TestBool "Click on an element made with New is False", (dummy.Click() = False)
    TestBool "  LastError = no-pane", (dummy.LastError = "no-pane")
    dummy.value = "x"
    TestBool "  Value = is also LastOk=False / no-pane", _
           (dummy.LastOk = False And dummy.LastError = "no-pane")

    Set el = D2El(p, "txt")
    TestBool "can write before the cleanup", el.SetAttribute("data-a", "1")
    TestBool "ClearElementRegistry succeeds", p.ClearElementRegistry()
    TestBool "*Click after the cleanup is False*", (el.Click() = False)
    TestBool "  LastError = stale", (el.LastError = "stale")
    el.value = (ChrW(&H66F8&) & ChrW(&H3051&) & ChrW(&H306A&) & ChrW(&H3044&) & ChrW(&H306F&) & ChrW(&H305A&))
    TestBool "  Value = also stops at stale", _
           (el.LastOk = False And el.LastError = "stale")

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback depth (expected 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D3_Write end ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' Test_D3_Framework (*measuring unknown 3* does it get through to the framework?)
'
'   The #react on the verification page has *monitoring that imitates React's value tracker*
'   laid over it. The tracker remembers "the value it knows", and when an input
'   event arrives and it differs from the current value, it *notices the change*.
'
'     Plain assignment (e.value = x)  ... it goes through the tracker, so the tracker remembers the value, and
'                               even when input fires it *does not notice* (ignored increases)
'     Via the native setter  ... it bypasses the tracker, so the values differ and
'                               on input it *notices* (notified increases)
'
'   This is the very reason D-3 exists. If it FAILs, suspect the descriptor acquisition of Wv2Element.Value =
'   (e.constructor.prototype).
' ============================================================
Public Sub Test_D3_Framework()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim el As Wv2Element

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D3_Framework: Browser is not started." & _
                    "Run UserForm1.StartWebView2_Full first."
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD3ProbeHtml())
    If p Is Nothing Then
        Wv2Log.LogI "Test_D3_Framework: failed to create the tab."
        Exit Sub
    End If
    If Not D2WaitTitle(p, ("D-3 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&)), 10) Then
        Wv2Log.LogI "Test_D3_Framework: could not confirm that the verification page loaded."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D3_Framework start ================"
    Wv2Log.LogI "  --- (1) *Plain assignment* + firing input (writing without D-3) ---"

    p.EvalSync "(function(){var e=document.getElementById('react');" & _
               ("e.value='" & ChrW(&H7D20&) & ChrW(&H306E&) & ChrW(&H4EE3&) & ChrW(&H5165&) & "';") & _
               "e.dispatchEvent(new Event('input',{bubbles:true}));return 1;})()"
    TestBool "the plain assignment ran", p.LastEvalOk
    TestBool "  it went through the tracker (trackedSet=1)", (D3Cnt(p, "trackedSet") = 1)
    TestBool "  *the framework does not notice (notified=0)*", _
           (D3Cnt(p, "notified") = 0)
    TestBool "  it was counted as a miss (ignored=1)", (D3Cnt(p, "ignored") = 1)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) *Wv2Element.Value =* (the D-3 way of writing) ---"

    Set el = D2El(p, "react")
    el.value = ("D-3 " & ChrW(&H306E&) & ChrW(&H66F8&) & ChrW(&H304D&) & ChrW(&H8FBC&) & ChrW(&H307F&))
    TestBool "the write succeeds", el.LastOk
    TestBool "  via the native setter (LastInfo=setter)", (el.LastInfo = "setter")
    TestBool "  *bypassed the tracker (trackedSet stays at 1)*", _
           (D3Cnt(p, "trackedSet") = 1)
    TestBool "  *the framework noticed the change (notified=1)*", _
           (D3Cnt(p, "notified") = 1)
    TestEq "  the value is also in", el, el.value, ("D-3 " & ChrW(&H306E&) & ChrW(&H66F8&) & ChrW(&H304D&) & ChrW(&H8FBC&) & ChrW(&H307F&))

    Wv2Log.LogI ""
    Wv2Log.LogI "  This is the heart of D-3. If (1) has notified=0 and (2) has notified=1,"
    Wv2Log.LogI "  then .Value = works on React / Vue pages too."
    Wv2Log.LogI "  in-callback depth (expected 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D3_Framework end ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' TestBool / TestEq / TestCountReset / TestCountPrint (verdict helpers)
'
'   *The common verdicts used by all tests of D-1 / D-2 / D-3* (originally D2Bool / D2Eq of D-2.
'   Merged into one pair when D-3 made them also output to the log file)
'
'   The same verdict as D-2's TestBool / TestEq, but the difference is that *the output target is Wv2Log*.
'   Wv2Log internally also fires Debug.Print, so how the Immediate window looks does not change.
'   The aim is *the verdict lines remain in the log file*:
'     - The Immediate window is flushed away by the ExecuteScript plumbing log (15 lines per round trip)
'     - With the log file it is enough to search for ' [FAIL]'
'
'   TestCountPrint outputs "OK n / FAIL m" at the end, so looking at just that tells pass/fail.
' ============================================================
Private Sub TestCountReset()
    m_okCount = 0
    m_ngCount = 0
End Sub

Private Sub TestCountPrint()
    Wv2Log.LogI "  *Verdicts " & (m_okCount + m_ngCount) & ": OK " & m_okCount & _
                " / FAIL " & m_ngCount & "*"
End Sub

Private Sub TestBool(ByVal label As String, ByVal cond As Boolean)
    If cond Then
        m_okCount = m_okCount + 1
        Wv2Log.LogI "  [OK  ] " & label
    Else
        m_ngCount = m_ngCount + 1
        Wv2Log.LogI "  [FAIL] " & label
    End If
End Sub

Private Sub TestEq(ByVal label As String, _
                 ByVal el As Wv2Element, _
                 ByVal got As String, _
                 ByVal want As String)
    If got = want Then
        m_okCount = m_okCount + 1
        Wv2Log.LogI "  [OK  ] " & label
    Else
        m_ngCount = m_ngCount + 1
        Wv2Log.LogI "  [FAIL] " & label
        Wv2Log.LogI "         expected: [" & want & "]"
        Wv2Log.LogI "         actual:   [" & got & "]"
    End If

    If el Is Nothing Then Exit Sub
    If Not el.LastOk Then
        Wv2Log.LogI "         * LastOk=False err=" & el.LastError
    End If
End Sub

' ============================================================
' D3Cnt / D3Str (D-3: read the verification page's monitoring counters)
'   Takes one item from window.__p on the page side. Returns -1 if it cannot be read.
' ============================================================
Private Function D3Cnt(ByVal p As Wv2Pane, ByVal countName As String) As Long
    Dim res As String

    res = p.EvalSync("window.__p." & countName)
    If Not p.LastEvalOk Then
        m_ngCount = m_ngCount + 1
        Wv2Log.LogW "  [FAIL] cannot read counter " & countName & " err=" & p.LastEvalError
        D3Cnt = -1
        Exit Function
    End If

    D3Cnt = CLng(Val(res))
End Function

Private Function D3Str(ByVal p As Wv2Pane, ByVal itemName As String) As String
    Dim res As String

    res = p.EvalSync("window.__p." & itemName)
    If Not p.LastEvalOk Then
        D3Str = "(cannot read: " & p.LastEvalError & ")"
        Exit Function
    End If

    D3Str = Wv2Json.JsonUnescape(res)
End Function


' ============================================================
' D3Pump (D-3: run DoEvents for the specified seconds)
'   Just marking time to wait for the completion notification of a raw ExecuteScript to appear.
'   *Do not break/step through during this either* (spec fact 20)
' ============================================================
Private Sub D3Pump(ByVal waitSec As Single)
    Dim t0 As Single

    t0 = Timer
    Do
        DoEvents
        If (Timer - t0) > waitSec Then Exit Do
    Loop
End Sub


' ============================================================
' BuildD3ProbeHtml (the verification page of D-3)
'
'   The D-2 page with *monitoring JS* added. What it counts:
'     inputs / changes ... the number of input / change events fired at #txt (issue 6)
'     clicks / clickInfo ... the click count of #btn and the identity of the event
'     trackedSet ... the number of assignments that went through the *React-like value tracker*
'     notified   ... the number of times the tracker noticed "the value changed"
'     ignored    ... the number of times an input came but the tracker did not notice
'
'   *Quotation marks in the static HTML part are written with VBA's ""* (it is not JS, so no problem)
'   *JS strings are always single quotes* (project rule)
' ============================================================
Private Function BuildD3ProbeHtml() As String
    Dim s As String

    s = "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    s = s & ("<title>D-3 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&) & "</title>") & vbLf
    s = s & "<style>" & vbLf
    s = s & "  body{font-family:'Segoe UI','Meiryo',sans-serif;background:#12161f;" & _
            "color:#e8eaed;padding:32px;line-height:1.8;}" & vbLf
    s = s & "  h1{font-size:22px;margin:0 0 18px;}" & vbLf
    s = s & "  .card{border:1px solid rgba(255,255,255,.12);border-radius:10px;" & _
            "padding:14px 16px;margin:12px 0;background:rgba(255,255,255,.04);}" & vbLf
    s = s & "  input,textarea,select,button{font-size:14px;padding:6px 8px;margin:4px 4px;" & _
            "background:#0b0e15;color:#e8eaed;border:1px solid rgba(255,255,255,.18);" & _
            "border-radius:6px;}" & vbLf
    s = s & "  pre{background:#0b0e15;padding:10px 12px;border-radius:8px;" & _
            "white-space:pre-wrap;word-break:break-all;}" & vbLf
    s = s & "  .note{color:#8ea2c8;font-size:12.5px;}" & vbLf
    s = s & "</style></head><body>" & vbLf
    s = s & ("<h1 id=""ttl"">D-3 " & ChrW(&H66F8&) & ChrW(&H304D&) & ChrW(&H8FBC&) & ChrW(&H307F&) & ChrW(&H3068&) & ChrW(&H64CD&) & ChrW(&H4F5C&) & ChrW(&H306E&) & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&) & "</h1>") & vbLf
    s = s & ("<p class=""note"">" & ChrW(&H3053&) & ChrW(&H306E&) & ChrW(&H30DA&) & ChrW(&H30FC&) & ChrW(&H30B8&) & ChrW(&H306F&) & " Test_D3_Write / Test_D3_Framework " & ChrW(&H5C02&) & ChrW(&H7528&) & ChrW(&H3067&) & ChrW(&H3059&) & ChrW(&H3002&)) & _
            (ChrW(&H5916&) & ChrW(&H90E8&) & ChrW(&H30B5&) & ChrW(&H30A4&) & ChrW(&H30C8&) & ChrW(&H306B&) & ChrW(&H4F9D&) & ChrW(&H5B58&) & ChrW(&H3057&) & ChrW(&H307E&) & ChrW(&H305B&) & ChrW(&H3093&) & ChrW(&H3002&) & "</p>") & vbLf
    s = s & "<div class=""card"">" & vbLf
    s = s & ("  <input id=""txt"" type=""text"" name=""q"" value=""" & ChrW(&H521D&) & ChrW(&H671F&) & ChrW(&H5024&) & """>") & vbLf
    s = s & ("  <textarea id=""area"" rows=""2"">" & ChrW(&H5143&) & ChrW(&H306E&) & ChrW(&H30C6&) & ChrW(&H30AD&) & ChrW(&H30B9&) & ChrW(&H30C8&) & "</textarea>") & vbLf
    s = s & "  <select id=""sel""><option value=""a"">A</option>" & _
            "<option value=""b"" selected>B</option></select>" & vbLf
    s = s & "  <input id=""chk"" type=""checkbox"">" & vbLf
    s = s & "  <input id=""react"" type=""text"" value="""">" & vbLf
    s = s & "</div>" & vbLf
    s = s & ("<p><button id=""btn"" type=""button"">" & ChrW(&H30AF&) & ChrW(&H30EA&) & ChrW(&H30C3&) & ChrW(&H30AF&) & ChrW(&H3057&) & ChrW(&H3066&) & ChrW(&H306D&) & "</button>") & _
            ("<span id=""btnlog"">" & ChrW(&H672A&) & ChrW(&H30AF&) & ChrW(&H30EA&) & ChrW(&H30C3&) & ChrW(&H30AF&) & "</span></p>") & vbLf
    s = s & ("<div id=""dv"" class=""card"">" & ChrW(&H5C5E&) & ChrW(&H6027&) & ChrW(&H306E&) & ChrW(&H66F8&) & ChrW(&H304D&) & ChrW(&H8FBC&) & ChrW(&H307F&) & ChrW(&H5148&) & "</div>") & vbLf
    s = s & "<pre id=""cnt""></pre>" & vbLf
    s = s & "<script>" & vbLf
    s = s & "(function(){" & vbLf
    s = s & "  var P={inputs:0,changes:0,clicks:0,clickInfo:''," & _
            "trackedSet:0,notified:0,ignored:0};" & vbLf
    s = s & "  window.__p=P;" & vbLf
    s = s & "  function show(){" & _
            "document.getElementById('cnt').textContent=JSON.stringify(P);}" & vbLf
    s = s & "  var t=document.getElementById('txt');" & vbLf
    s = s & "  t.addEventListener('input',function(){P.inputs++;show();});" & vbLf
    s = s & "  t.addEventListener('change',function(){P.changes++;show();});" & vbLf
    s = s & "  var b=document.getElementById('btn');" & vbLf
    s = s & "  b.addEventListener('click',function(ev){P.clicks++;" & vbLf
    s = s & "    P.clickInfo=ev.type+'/'+ev.bubbles+'/'+ev.isTrusted;" & vbLf
    s = s & "    document.getElementById('btnlog').textContent=" & _
            ("'" & ChrW(&H30AF&) & ChrW(&H30EA&) & ChrW(&H30C3&) & ChrW(&H30AF&) & " '+P.clicks+' " & ChrW(&H56DE&) & "';show();});") & vbLf
    s = s & "  var r=document.getElementById('react');" & vbLf
    s = s & "  var d=Object.getOwnPropertyDescriptor(r.constructor.prototype,'value');" & vbLf
    s = s & "  var last=d.get.call(r);" & vbLf
    s = s & "  Object.defineProperty(r,'value',{configurable:true," & vbLf
    s = s & "    get:function(){return d.get.call(this);}," & vbLf
    s = s & "    set:function(v){P.trackedSet++;last=v;d.set.call(this,v);show();}});" & vbLf
    s = s & "  r.addEventListener('input',function(){" & vbLf
    s = s & "    var cur=d.get.call(r);" & vbLf
    s = s & "    if(cur!==last){P.notified++;last=cur;}else{P.ignored++;}" & vbLf
    s = s & "    show();});" & vbLf
    s = s & "  show();" & vbLf
    s = s & "})();" & vbLf
    s = s & "</" & "script>" & vbLf
    s = s & "</body></html>"

    BuildD3ProbeHtml = s
End Function


' ============================================================
' Test_D3_Wait (verification of D-3b: waiting for DOM conditions)
'
'   Makes elements that appear late / disappear late with setTimeout, and checks
'   whether WaitFor / WaitGone / ImplicitWaitSec hang on as expected.
'   *Also outputs the seconds waited* because we want to see not only that it "hung on" but also that it does "not hang on too long".
'
'   *Do not break/step through during execution* (spec fact 20)
'   The whole wait loop becomes a long DoEvents section (as in unknown 2).
' ============================================================
Public Sub Test_D3_Wait()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim el As Wv2Element
    Dim t0 As Single
    Dim took As Single

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D3_Wait: Browser is not started." & _
                    "Run UserForm1.Show vbModeless and StartWebView2_Full first."
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD3ProbeHtml())
    If p Is Nothing Then
        Wv2Log.LogI "Test_D3_Wait: failed to create the tab."
        Exit Sub
    End If
    If Not D2WaitTitle(p, ("D-3 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&)), 10) Then
        Wv2Log.LogI "Test_D3_Wait: could not confirm that the verification page loaded."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D3_Wait start ================"
    Wv2Log.LogI "  --- (1) An element that already exists returns without waiting ---"

    t0 = Timer
    Set el = p.WaitFor("#txt", 5)
    took = D3Took(t0)
    TestBool "an element that already exists can be grabbed with WaitFor", Not (el Is Nothing)
    Wv2Log.LogI "        time waited = " & Format$(took, "0.00") & " s"
    TestBool "  *returns without waiting (under 0.5 s)*", (took < 0.5)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) An element that appears late (added after 800ms) ---"

    p.EvalSync "(function(){setTimeout(function(){" & _
               "var d=document.createElement('input');d.id='late';d.type='text';" & _
               ("d.value='" & ChrW(&H9045&) & ChrW(&H308C&) & ChrW(&H3066&) & ChrW(&H51FA&) & ChrW(&H305F&) & "';document.body.appendChild(d);},800);return 1;})()")
    TestBool "the delayed addition could be set up", p.LastEvalOk

    t0 = Timer
    Set el = p.WaitFor("#late", 5)
    took = D3Took(t0)
    TestBool "*the late-appearing element could be grabbed*", Not (el Is Nothing)
    Wv2Log.LogI "        time waited = " & Format$(took, "0.00") & " s (the setup is 0.80 s)"
    TestBool "  the wait time is reasonable (0.4-3.0 s)", (took > 0.4 And took < 3)

    If Not el Is Nothing Then
        TestEq "  the value of the grabbed element can be read", el, el.value, (ChrW(&H9045&) & ChrW(&H308C&) & ChrW(&H3066&) & ChrW(&H51FA&) & ChrW(&H305F&))
        el.value = (ChrW(&H5F85&) & ChrW(&H3063&) & ChrW(&H3066&) & ChrW(&H304B&) & ChrW(&H3089&) & ChrW(&H66F8&) & ChrW(&H3044&) & ChrW(&H305F&))
        TestBool "  *can write straight to the waited-for element (connection with D-3a)*", el.LastOk
        TestEq "    can be read back", el, el.value, (ChrW(&H5F85&) & ChrW(&H3063&) & ChrW(&H3066&) & ChrW(&H304B&) & ChrW(&H3089&) & ChrW(&H66F8&) & ChrW(&H3044&) & ChrW(&H305F&))
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) Something that does not appear (timeout in 1 second) ---"

    t0 = Timer
    Set el = p.WaitFor("#never-ever", 1)
    took = D3Took(t0)
    TestBool "something that does not appear is Nothing", (el Is Nothing)
    TestBool "  *LastEvalOk=True (meaning it did not appear)*", (p.LastEvalOk = True)
    Wv2Log.LogI "        time waited = " & Format$(took, "0.00") & " s (specified 1.00 s)"
    TestBool "  it really hung on (0.8 s or more)", (took > 0.8)
    TestBool "  it does not hang on too long (under 2.5 s)", (took < 2.5)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) *A failure gives up without waiting* (invalid selector) ---"

    t0 = Timer
    Set el = p.WaitFor("###", 5)
    took = D3Took(t0)
    TestBool "an invalid selector is Nothing", (el Is Nothing)
    TestBool "  *LastEvalOk=False (meaning it failed)*", (p.LastEvalOk = False)
    Wv2Log.LogI "        time waited = " & Format$(took, "0.00") & " s (specified 5.00 s)"
    TestBool "  *returns at once without waiting 5 s (under 0.5 s)*", (took < 0.5)
    Wv2Log.LogI "        LastEvalError = " & p.LastEvalError

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) WaitGone (remove #dv after 800ms) ---"

    p.EvalSync "(function(){setTimeout(function(){" & _
               "var e=document.getElementById('dv');" & _
               "if(e){e.parentNode.removeChild(e);}},800);return 1;})()"
    TestBool "the delayed removal could be set up", p.LastEvalOk

    t0 = Timer
    TestBool "*could wait until it disappeared*", p.WaitGone("#dv", 5)
    took = D3Took(t0)
    Wv2Log.LogI "        time waited = " & Format$(took, "0.00") & " s (the setup is 0.80 s)"
    TestBool "  the wait time is reasonable (0.4-3.0 s)", (took > 0.4 And took < 3)

    t0 = Timer
    TestBool "something absent from the start is True at once", p.WaitGone("#never-ever", 5)
    TestBool "  returns without waiting (under 0.5 s)", (D3Took(t0) < 0.5)

    t0 = Timer
    TestBool "something that does not disappear is False", (p.WaitGone("#ttl", 1) = False)
    took = D3Took(t0)
    TestBool "  *LastEvalOk=True (meaning it is still there)*", (p.LastEvalOk = True)
    TestBool "  it really hung on (0.8 s or more)", (took > 0.8)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) ImplicitWaitSec (opt-in of issue 2 plan C) ---"

    TestBool "*the default is 0 (disabled)*", (p.ImplicitWaitSec = 0)
    t0 = Timer
    Set el = p.QuerySelector("#late2")
    took = D3Took(t0)
    TestBool "by default QuerySelector returns Nothing without waiting", _
           (el Is Nothing And took < 0.5)

    p.EvalSync "(function(){setTimeout(function(){" & _
               "var d=document.createElement('div');d.id='late2';" & _
               ("d.textContent='" & ChrW(&H5F8C&) & ChrW(&H304B&) & ChrW(&H3089&) & ChrW(&H51FA&) & ChrW(&H305F&) & " 2';document.body.appendChild(d);},800);") & _
               "return 1;})()"
    TestBool "the delayed addition could be set up", p.LastEvalOk

    p.ImplicitWaitSec = 5
    t0 = Timer
    Set el = p.QuerySelector("#late2")
    took = D3Took(t0)
    TestBool "*with ImplicitWaitSec=5 QuerySelector hangs on and picks it up*", _
           Not (el Is Nothing)
    Wv2Log.LogI "        time waited = " & Format$(took, "0.00") & " s (the setup is 0.80 s)"
    TestBool "  the wait time is reasonable (0.4-3.0 s)", (took > 0.4 And took < 3)

    p.ImplicitWaitSec = 0
    TestBool "can be returned to 0", (p.ImplicitWaitSec = 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback depth (expected 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D3_Wait end ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' D3Took (D-3b: elapsed seconds. *Timer returns to 0 at midnight*, so correct for it)
' ============================================================
Private Function D3Took(ByVal sinceTimer As Single) As Single
    Dim d As Single

    d = Timer - sinceTimer
    If d < 0 Then d = d + 86400
    D3Took = d
End Function

' ============================================================
' Test_D3_Help (procedure for D-3)
' ============================================================
Public Sub Test_D3_Help()
    Debug.Print ""
    Debug.Print "=========================================================="
    Debug.Print " D-3 verification procedure (writes and operations)"
    Debug.Print "=========================================================="
    Debug.Print ""
    Debug.Print "  --- Preparation ---"
    Debug.Print "  1) Wv2Log.LogStart  ... contain this one test run in one file"
    Debug.Print "  2) Run UserForm1.Show vbModeless, then run StartWebView2_Full"
    Debug.Print "     *Show comes first* Without the form's window the HWND of Frame1 cannot"
    Debug.Print "     be obtained, and Browser.Init fails with hWnd_Frame still 0."
    Debug.Print "  3) *Wait until the event burst settles* (spec fact 20)"
    Debug.Print ""
    Debug.Print "  --- Run (*in this order*) ---"
    Debug.Print "  4) Test_D3_Probe_Promise ... *done* measuring unknown 1 (settled 2026-08-22)"
    Debug.Print "     Measures whether ExecuteScript waits for a Promise. This result"
    Debug.Print "     decides the design of the wait API (WaitFor / WaitGone)."
    Debug.Print "     *An active tab is required* (it does not open a new tab)"
    Debug.Print "  5) Test_D3_Write      ... *done* writes and operations (33 OK)"
    Debug.Print "  6) Test_D3_Framework  ... *done* measuring whether it works on SPAs (7 OK)"
    Debug.Print "  7) Test_D3_Wait       ... *verification of D-3b* the wait API (opens 1 tab)"
    Debug.Print "     NOTE: 4-6 are settled. Only 7 needs to be looked at next."
    Debug.Print ""
    Debug.Print "  *Do not break/step through during execution* (spec fact 20)"
    Debug.Print ""
    Debug.Print "  --- *Verdicts remain in the log file* ---"
    Debug.Print "    The D-3 tests also write their verdicts to Wv2Log. The Immediate window is"
    Debug.Print "    flushed away at once by the ExecuteScript plumbing log (15 lines per round trip),"
    Debug.Print "    so *check pass/fail in the log file*:"
    Debug.Print "      ?Wv2Log.LogPath      ... location of the file"
    Debug.Print "      Just look at the trailing ""*Verdicts n: OK x / FAIL y*"" to know pass/fail"
    Debug.Print "      If there is a FAIL, search for [FAIL]"
    Debug.Print ""
    Debug.Print "  --- What to look at (Test_D3_Probe_Promise) ---"
    Debug.Print "  - Did lines of resultJson= appear between the markers of (B-1)(B-2)?"
    Debug.Print "      {} appeared in (B-2)   -> *does not wait* = poll in VBA as in the skeleton of issue 3"
    Debug.Print "      nothing appeared in (B-2)   -> *waits*   = redo the design"
    Debug.Print ""
    Debug.Print "  --- What to look at (Test_D3_Write) ---"
    Debug.Print "  - All lines of (1)-(8) are [OK  ]"
    Debug.Print "  - LastInfo of (1) is setter (via the native setter)"
    Debug.Print "  - input / change of (3) are both 3 times (firing both, issue 6)"
    Debug.Print "  - (7) is [OK  ] = confirms the difference between 'JS ran' and 'the effect appeared'"
    Debug.Print "  - (8) is [OK  ] = distinguishing no-pane / stale (following the D-2 convention)"
    Debug.Print ""
    Debug.Print "  --- What to look at (Test_D3_Framework) *the heart of D-3* ---"
    Debug.Print "  - (1) is notified=0 / ignored=1  ... a plain assignment is not noticed"
    Debug.Print "  - (2) is notified=1 / trackedSet stays at 1 ... the D-3 way of writing is noticed"
    Debug.Print "    If this FAILs, .Value = does not work on React / Vue pages."
    Debug.Print ""
    Debug.Print "  --- What to look at (Test_D3_Wait) *D-3b* ---"
    Debug.Print "  - All lines of (1)-(6) are [OK  ]"
    Debug.Print "  - (2) an element that appears late is grabbed and can be written to as is"
    Debug.Print "      = evidence that the wait (D-3b) and the write (D-3a) are connected"
    Debug.Print "  - (3) something that does not appear gives Nothing + LastEvalOk=True *after hanging on for 1 s*"
    Debug.Print "  - (4) *an invalid selector gives up at once without waiting*"
    Debug.Print "      If this FAILs, calling from inside a handler would freeze for the wait time"
    Debug.Print "  - (6) ImplicitWaitSec defaults to 0, and with 5 QuerySelector hangs on"
    Debug.Print ""
    Debug.Print "  --- Regression check (because D-3 touched Wv2Pane and Wv2Element) ---"
    Debug.Print "  8) Test_D2_Find   ... reading must not be broken"
    Debug.Print "  9) Test_D2_Stale  ... the handling of generation and stale must not be broken"
    Debug.Print " 10) Test_D1_Eval / Test_D1_Guard ... EvalSync and the guard"
    Debug.Print ""
    Debug.Print "  --- If you want to try by hand ---"
    Debug.Print "  Set p = UserForm1.GetActivePane"
    Debug.Print "  Set el = p.QuerySelector(""input[name='q']"")"
    Debug.Print "  el.Value = ""search term"""
    Debug.Print "  ?el.LastOk : ?el.LastInfo : ?el.Value"
    Debug.Print "  ?p.QuerySelector(""button"").Click"
    Debug.Print ""
    Debug.Print "  --- Waiting (D-3b) ---"
    Debug.Print "  Set el = p.WaitFor(""#result"", 5)   ' up to 5 seconds until it appears"
    Debug.Print "  ?p.WaitGone("".spinner"", 10)        ' up to 10 seconds until it disappears"
    Debug.Print "  p.ImplicitWaitSec = 5               ' default 0. Adds an implicit hang-on to searches"
    Debug.Print ""
    Debug.Print "  *What can be waited for is only 'DOM conditions'*"
    Debug.Print "    WaitFor / WaitGone only look at the presence of the specified element."
    Debug.Print "    The quiescence wait that waits for 'all communication to finish and things to go quiet'"
    Debug.Print "    (MutationObserver / fetch counter) is an open item for D-4."
    Debug.Print ""
End Sub

' ============================================================
' Test_D4_Probe (verification of D-4a: in-page SPA probe)
'
'   *The first move of D-4* Before building the body of the wait, confirm that the observing mechanism stands up healthily
'   (design principle 103: check the probe itself before measuring).
' ============================================================
Public Sub Test_D4_Probe()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim st As String
    Dim before As Single
    Dim after As Single
    Dim n0 As Long
    Dim g0 As Long

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D4_Probe: Browser is not started." & _
                    "Run UserForm1.Show vbModeless and StartWebView2_Full first."
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD4ProbeHtml())
    If p Is Nothing Then
        Wv2Log.LogI "Test_D4_Probe: failed to create the tab."
        Exit Sub
    End If
    If Not D2WaitTitle(p, ("D-4 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&)), 10) Then
        Wv2Log.LogI "Test_D4_Probe: could not confirm that the verification page loaded."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D4_Probe start ================"
    Wv2Log.LogI "  --- (1) The probe is not created until it is called (unknown 4) ---"

    st = p.EvalSync("typeof window.__wv2p")
    TestBool "*undefined before installation*", (Wv2Json.JsonUnescape(st) = "undefined")
    Wv2Log.LogI "        typeof window.__wv2p = " & st

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) Comparison of round-trip cost (unknown 2) ---"

    before = D4RoundTrip(p, 20)
    Wv2Log.LogI "        1 round trip before installation = " & Format$(before, "0.0") & " ms"

    st = p.SpaProbeState()
    TestBool "SpaProbeState succeeds", p.LastEvalOk
    Wv2Log.LogI "        state = " & st

    after = D4RoundTrip(p, 20)
    Wv2Log.LogI "        1 round trip after installation = " & Format$(after, "0.0") & " ms"
    TestBool "  setting up the observation does not make the round trip extremely slow (under 3x)", _
             (after < before * 3 + 10)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) Initial state ---"

    TestBool "*healthy (h=true)*", (InStr(1, st, """h"":true") > 0)
    TestBool "  MutationObserver is alive (ob=true)", _
             (InStr(1, st, """ob"":true") > 0)
    TestBool "  version number is 1", (Wv2Json.JsonGetNum(st, "v") = 1)
    TestBool "  generation is 1", (Wv2Json.JsonGetNum(st, "g") = 1)
    TestBool "  rebuild count is 0", (Wv2Json.JsonGetNum(st, "rp") = 0)
    TestBool "  in-flight communications is 0", _
             (Wv2Json.JsonGetNum(st, "f") = 0 And Wv2Json.JsonGetNum(st, "x") = 0)
    TestBool "  SpaProbeHealthy is also True", p.SpaProbeHealthy()

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) Counting DOM changes ---"

    st = p.SpaProbeState()
    n0 = Wv2Json.JsonGetNum(st, "m")
    p.EvalSync "(function(){document.getElementById('target')" & _
               (".textContent='" & ChrW(&H66F8&) & ChrW(&H304D&) & ChrW(&H63DB&) & ChrW(&H3048&) & ChrW(&H305F&) & " '+Date.now();return 1;})()")
    st = p.SpaProbeState()
    TestBool "*DOM changes were counted*", (Wv2Json.JsonGetNum(st, "m") > n0)
    Wv2Log.LogI "        elapsed since the last change q = " & _
                Wv2Json.JsonGetNum(st, "q") & " ms"
    TestBool "  q is small because it is right after (under 1000ms)", _
             (Wv2Json.JsonGetNum(st, "q") < 1000)
    TestBool "  sm = 0 because nothing is excluded", (Wv2Json.JsonGetNum(st, "sm") = 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) Counting communications (a blob URL, so it completes offline) ---"

    st = p.SpaProbeState()
    n0 = Wv2Json.JsonGetNum(st, "n")

    p.EvalSync "(function(){var u=URL.createObjectURL(new Blob(['hi']));" & _
               "fetch(u).then(function(r){return r.text();});return 1;})()"
    D3Pump 1
    st = p.SpaProbeState()
    TestBool "*fetch was counted*", (Wv2Json.JsonGetNum(st, "n") > n0)
    TestBool "  completed and the in-flight count returned to 0", (Wv2Json.JsonGetNum(st, "f") = 0)

    n0 = Wv2Json.JsonGetNum(st, "n")
    p.EvalSync "(function(){var u=URL.createObjectURL(new Blob(['hi']));" & _
               "var r=new XMLHttpRequest();r.open('GET',u);r.send();return 1;})()"
    D3Pump 1
    st = p.SpaProbeState()
    TestBool "*XHR was counted*", (Wv2Json.JsonGetNum(st, "n") > n0)
    TestBool "  completed and the in-flight count returned to 0", (Wv2Json.JsonGetNum(st, "x") = 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) *Health check and automatic repair* (unknown 1) ---"

    st = p.SpaProbeState()
    g0 = Wv2Json.JsonGetNum(st, "g")
    Wv2Log.LogI "        before breaking: generation=" & g0 & " rebuilds=" & _
                Wv2Json.JsonGetNum(st, "rp")

    Wv2Log.LogI "        Create a situation where the page side replaced window.fetch"
    p.EvalSync "(function(){var o=window.fetch;" & _
               "window.fetch=function(){return o.apply(this,arguments);};return 1;})()"

    st = p.SpaProbeState()
    TestBool "*even if broken, calling it returns it to healthy (h=true)*", _
             (InStr(1, st, """h"":true") > 0)
    TestBool "  *the rebuild was recorded (rp >= 1)*", _
             (Wv2Json.JsonGetNum(st, "rp") >= 1)
    TestBool "  the generation advanced by 1", (Wv2Json.JsonGetNum(st, "g") = g0 + 1)
    Wv2Log.LogI "        after breaking: state = " & st

    n0 = Wv2Json.JsonGetNum(st, "n")
    p.EvalSync "(function(){var u=URL.createObjectURL(new Blob(['hi']));" & _
               "fetch(u).then(function(r){return r.text();});return 1;})()"
    D3Pump 1
    st = p.SpaProbeState()
    TestBool "  *it can count even after the rebuild*", (Wv2Json.JsonGetNum(st, "n") > n0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (7) SpaProbeReset (manual rebuild) ---"

    g0 = Wv2Json.JsonGetNum(st, "g")
    TestBool "SpaProbeReset succeeds", p.SpaProbeReset()
    st = p.SpaProbeState()
    TestBool "  the generation advanced", (Wv2Json.JsonGetNum(st, "g") > g0)
    TestBool "  the counted value returned to 0", (Wv2Json.JsonGetNum(st, "n") = 0)
    TestBool "  healthy", (InStr(1, st, """h"":true") > 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (8) The probe disappears on page navigation (unknown 3) ---"

    p.View_NavigateToString BuildD2SecondHtml()
    If Not D2WaitTitle(p, ("D-2 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&) & " 2 " & ChrW(&H679A&) & ChrW(&H76EE&)), 10) Then
        Wv2Log.LogI "  [FAIL] could not confirm the navigation"
        m_ngCount = m_ngCount + 1
    Else
        st = p.EvalSync("typeof window.__wv2p")
        TestBool "*returns to undefined after navigation*", _
                 (Wv2Json.JsonUnescape(st) = "undefined")
        st = p.SpaProbeState()
        TestBool "  can be set up again on the new page", _
                 (InStr(1, st, """h"":true") > 0)
        TestBool "  the generation is counted again from 1", (Wv2Json.JsonGetNum(st, "g") = 1)
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback depth (expected 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D4_Probe end ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' D4RoundTrip (D-4a: measure the average ms of one EvalSync round trip)
'   *Re-measuring spec fact 52 (about 38ms)* Checks whether setting up the probe makes it heavy.
' ============================================================
Private Function D4RoundTrip(ByVal p As Wv2Pane, ByVal shots As Long) As Single
    Dim i As Long
    Dim t0 As Single

    t0 = Timer
    For i = 1 To shots
        p.EvalSync "1"
    Next i

    D4RoundTrip = D3Took(t0) / shots * 1000
End Function


' ============================================================
' BuildD4ProbeHtml (the verification page of D-4)
'
'   Made *completely quiet right after loading* (needed for verifying the quiescence wait).
'   Page-side entry points:
'     startNoise(ms) ... keeps rewriting #noise every ms (reproduces a page that does not settle)
'     stopNoise()    ... stops it
'     later(ms)      ... adds #late after ms (used in D-4b)
'     chain(ms)      ... fetches a blob after ms and, once that completes, rewrites the DOM further
'                      (*the chain fetch -> DOM update* = reproduces an SPA. Used in D-4b)
' ============================================================
Private Function BuildD4ProbeHtml() As String
    Dim s As String

    s = "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    s = s & ("<title>D-4 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&) & "</title>") & vbLf
    s = s & "<style>" & vbLf
    s = s & "  body{font-family:'Segoe UI','Meiryo',sans-serif;background:#12161f;" & _
            "color:#e8eaed;padding:32px;line-height:1.8;}" & vbLf
    s = s & "  h1{font-size:22px;margin:0 0 18px;}" & vbLf
    s = s & "  .card{border:1px solid rgba(255,255,255,.12);border-radius:10px;" & _
            "padding:14px 16px;margin:12px 0;background:rgba(255,255,255,.04);}" & vbLf
    s = s & "  .note{color:#8ea2c8;font-size:12.5px;}" & vbLf
    s = s & "</style></head><body>" & vbLf
    s = s & ("<h1 id=""ttl"">D-4 " & ChrW(&H9759&) & ChrW(&H7A4F&) & ChrW(&H5F85&) & ChrW(&H3061&) & ChrW(&H306E&) & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&) & "</h1>") & vbLf
    s = s & ("<p class=""note"">" & ChrW(&H8AAD&) & ChrW(&H307F&) & ChrW(&H8FBC&) & ChrW(&H307F&) & ChrW(&H76F4&) & ChrW(&H5F8C&) & ChrW(&H306F&) & ChrW(&H2605&) & ChrW(&H4F55&) & ChrW(&H3082&) & ChrW(&H52D5&) & ChrW(&H3044&) & ChrW(&H3066&) & ChrW(&H3044&) & ChrW(&H306A&) & ChrW(&H3044&) & ChrW(&H2605&) & ChrW(&H72B6&) & ChrW(&H614B&) & ChrW(&H3067&) & ChrW(&H3059&) & ChrW(&H3002&)) & _
            ("startNoise() / later() / chain() " & ChrW(&H3067&) & ChrW(&H52D5&) & ChrW(&H304D&) & ChrW(&H3092&) & ChrW(&H4F5C&) & ChrW(&H308A&) & ChrW(&H307E&) & ChrW(&H3059&) & ChrW(&H3002&) & "</p>") & vbLf
    s = s & ("<div id=""target"" class=""card"">" & ChrW(&H66F8&) & ChrW(&H304D&) & ChrW(&H63DB&) & ChrW(&H3048&) & ChrW(&H5BFE&) & ChrW(&H8C61&) & "</div>") & vbLf
    s = s & ("<div id=""noise"" class=""card"">" & ChrW(&H30CE&) & ChrW(&H30A4&) & ChrW(&H30BA&) & ": " & ChrW(&H505C&) & ChrW(&H6B62&) & ChrW(&H4E2D&) & "</div>") & vbLf
    s = s & "<div id=""slot"" class=""card""></div>" & vbLf
    s = s & ("<p><button id=""btn"" type=""button"">" & ChrW(&H62BC&) & ChrW(&H3059&) & ChrW(&H3068&) & " 400ms " & ChrW(&H5F8C&) & ChrW(&H306B&) & ChrW(&H66F4&) & ChrW(&H65B0&) & "</button></p>") & vbLf
    s = s & "<script>" & vbLf
    s = s & "(function(){" & vbLf
    s = s & "  var t=null;" & vbLf
    s = s & "  window.startNoise=function(ms){" & vbLf
    s = s & "    if(t){clearInterval(t);}" & vbLf
    s = s & "    t=setInterval(function(){" & vbLf
    s = s & ("      document.getElementById('noise').textContent='" & ChrW(&H30CE&) & ChrW(&H30A4&) & ChrW(&H30BA&) & " '+Date.now();") & vbLf
    s = s & "    },ms||200);return 1;};" & vbLf
    s = s & "  window.stopNoise=function(){" & vbLf
    s = s & "    if(t){clearInterval(t);t=null;}" & vbLf
    s = s & ("    document.getElementById('noise').textContent='" & ChrW(&H30CE&) & ChrW(&H30A4&) & ChrW(&H30BA&) & ": " & ChrW(&H505C&) & ChrW(&H6B62&) & ChrW(&H4E2D&) & "';return 1;};") & vbLf
    s = s & "  window.later=function(ms){" & vbLf
    s = s & "    setTimeout(function(){" & vbLf
    s = s & "      var d=document.createElement('div');d.id='late';" & vbLf
    s = s & ("      d.textContent='" & ChrW(&H9045&) & ChrW(&H308C&) & ChrW(&H3066&) & ChrW(&H51FA&) & ChrW(&H305F&) & "';") & vbLf
    s = s & "      document.getElementById('slot').appendChild(d);" & vbLf
    s = s & "    },ms||800);return 1;};" & vbLf
    s = s & "  window.chain=function(ms){" & vbLf
    s = s & "    setTimeout(function(){" & vbLf
    s = s & "      var u=URL.createObjectURL(new Blob(['ok']));" & vbLf
    s = s & "      fetch(u).then(function(r){return r.text();}).then(function(x){" & vbLf
    s = s & "        setTimeout(function(){" & vbLf
    s = s & ("          document.getElementById('target').textContent='" & ChrW(&H9023&) & ChrW(&H9396&) & ChrW(&H306E&) & ChrW(&H7D50&) & ChrW(&H679C&) & " '+x;") & vbLf
    s = s & "        },300);" & vbLf
    s = s & "      });" & vbLf
    s = s & "    },ms||500);return 1;};" & vbLf
    s = s & "  document.getElementById('btn').addEventListener('click',function(){" & vbLf
    s = s & "    setTimeout(function(){" & vbLf
    s = s & "      document.getElementById('target').textContent=" & _
            ("'" & ChrW(&H30AF&) & ChrW(&H30EA&) & ChrW(&H30C3&) & ChrW(&H30AF&) & ChrW(&H306E&) & ChrW(&H7D50&) & ChrW(&H679C&) & " '+Date.now();") & vbLf
    s = s & "    },400);" & vbLf
    s = s & "  });" & vbLf
    s = s & "})();" & vbLf
    s = s & "</" & "script>" & vbLf
    s = s & "</body></html>"

    BuildD4ProbeHtml = s
End Function


' ============================================================
' Test_D4_Settle (verification of D-4b: quiescence wait)
'
'   *The heart of D-4* Deterministically confirms that it returns not at "the fetch finished" but at the point
'   where "even the DOM update after that has finished".
'
'   *Do not break/step through during execution* (spec fact 20)
' ============================================================
Public Sub Test_D4_Settle()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim el As Wv2Element
    Dim t0 As Single
    Dim took As Single
    Dim ok As Boolean
    Dim txt As String

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D4_Settle: Browser is not started." & _
                    "Run UserForm1.Show vbModeless and StartWebView2_Full first."
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD4ProbeHtml())
    If p Is Nothing Then
        Wv2Log.LogI "Test_D4_Settle: failed to create the tab."
        Exit Sub
    End If
    If Not D2WaitTitle(p, ("D-4 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&)), 10) Then
        Wv2Log.LogI "Test_D4_Settle: could not confirm that the verification page loaded."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D4_Settle start ================"
    Wv2Log.LogI "  --- (1) A quiet page ---"

    t0 = Timer
    ok = p.WaitSettled(5)
    took = D3Took(t0)
    TestBool "a quiet page reaches quiescence", ok
    Wv2Log.LogI "        " & p.LastSettleInfo
    Wv2Log.LogI "        time taken = " & Format$(took, "0.00") & " s"
    TestBool "  *watches for at least the quiescence window (0.4 s or more)*", (took > 0.4)
    TestBool "  not needlessly long (under 2 s)", (took < 2)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) *The chain fetch -> DOM update* (the heart of D-4) ---"
    Wv2Log.LogI "        Fetch after 100ms and rewrite #target 300ms after that completes"

    p.EvalSync "(function(){document.getElementById('target')" & _
               (".textContent='" & ChrW(&H307E&) & ChrW(&H3060&) & "';return 1;})()")
    p.EvalSync "window.chain(100)"

    t0 = Timer
    ok = p.WaitSettled(5)
    took = D3Took(t0)
    TestBool "reaches quiescence after the chain", ok
    Wv2Log.LogI "        " & p.LastSettleInfo
    Wv2Log.LogI "        time taken = " & Format$(took, "0.00") & " s"

    Set el = D2El(p, "target")
    txt = el.InnerText
    Wv2Log.LogI "        #target at the point the wait ended = " & txt
    TestBool "** it waited until the DOM update (did not return on the communication finishing alone) **", _
             (InStr(1, txt, (ChrW(&H9023&) & ChrW(&H9396&) & ChrW(&H306E&) & ChrW(&H7D50&) & ChrW(&H679C&))) > 0)
    TestBool "  waited for the chain (0.7 s or more)", (took > 0.7)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) A page that does not settle (keeps rewriting every 50ms) ---"

    p.EvalSync "window.startNoise(50)"
    t0 = Timer
    ok = p.WaitSettled(2)
    took = D3Took(t0)
    TestBool "*False if it does not settle*", (ok = False)
    TestBool "  *LastEvalOk=True (not a failure but a timeout)*", (p.LastEvalOk = True)
    Wv2Log.LogI "        " & p.LastSettleInfo
    TestBool "  the breakdown is timeout", (InStr(1, p.LastSettleInfo, "timeout") > 0)
    TestBool "  it hung on for 2 seconds as specified", (took > 1.7 And took < 4)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) *Noise exclusion works* (issue 5) ---"
    Wv2Log.LogI "        Keep the noise going, and exclude only the changes of #noise from the quiescence judgment"

    p.IgnoreSelectors = "#noise"
    t0 = Timer
    ok = p.WaitSettled(5)
    took = D3Took(t0)
    TestBool "*with the exclusion it reaches quiescence*", ok
    Wv2Log.LogI "        " & p.LastSettleInfo
    Wv2Log.LogI "        time taken = " & Format$(took, "0.00") & " s"

    Wv2Log.LogI "        state = " & p.SpaProbeState()
    TestBool "  the excluded DOM changes are counted (sm > 0)", _
             (Wv2Json.JsonGetNum(p.SpaProbeState(), "sm") > 0)

    p.IgnoreSelectors = ""
    ok = p.WaitSettled(2)
    TestBool "  *removing the exclusion makes it not settle again*", (ok = False)
    p.EvalSync "window.stopNoise()"

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) Changing stableMs lengthens the wait ---"

    t0 = Timer
    ok = p.WaitSettled(6, 2000)
    took = D3Took(t0)
    TestBool "it reaches quiescence even with a 2000ms quiescence window", ok
    Wv2Log.LogI "        " & p.LastSettleInfo
    TestBool "  it takes 2 seconds or more", (took > 1.9)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) *AutoWaitAfterAction* (opt-in of issue 7) ---"

    TestBool "the default is False", (p.AutoWaitAfterAction = False)

    p.EvalSync "(function(){document.getElementById('target')" & _
               (".textContent='" & ChrW(&H62BC&) & ChrW(&H3059&) & ChrW(&H524D&) & "';return 1;})()")
    Set el = D2El(p, "btn")
    el.Click
    Set el = D2El(p, "target")
    txt = el.InnerText
    Wv2Log.LogI "        auto wait OFF: #target right after the click = " & txt
    TestBool "*with OFF the pre-update value is seen (not waiting)*", (txt = (ChrW(&H62BC&) & ChrW(&H3059&) & ChrW(&H524D&)))

    p.AutoWaitAfterAction = True
    p.EvalSync "(function(){document.getElementById('target')" & _
               (".textContent='" & ChrW(&H62BC&) & ChrW(&H3059&) & ChrW(&H524D&) & "';return 1;})()")
    Set el = D2El(p, "btn")
    t0 = Timer
    el.Click
    took = D3Took(t0)
    Set el = D2El(p, "target")
    txt = el.InnerText
    Wv2Log.LogI "        auto wait ON: Click took " & Format$(took, "0.00") & " s"
    Wv2Log.LogI "        #target right after the click = " & txt
    TestBool "*with ON the post-update value is seen (it is waiting)*", _
             (InStr(1, txt, (ChrW(&H30AF&) & ChrW(&H30EA&) & ChrW(&H30C3&) & ChrW(&H30AF&) & ChrW(&H306E&) & ChrW(&H7D50&) & ChrW(&H679C&))) > 0)
    Wv2Log.LogI "        " & p.LastSettleInfo
    p.AutoWaitAfterAction = False
    TestBool "  can be returned to False", (p.AutoWaitAfterAction = False)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (7) Entry point of the exclusion list (communications) ---"

    TestBool "AddIgnoreNetwork succeeds", p.AddIgnoreNetwork("example.invalid")
    TestBool "  adding the same one is still True (no duplicate)", _
             p.AddIgnoreNetwork("example.invalid")
    TestBool "  the quiescence judgment works even with an exclusion in", p.WaitSettled(5)
    p.ClearIgnoreNetwork

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback depth (expected 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D4_Settle end ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_D4_Signal (verification of D-4c: explicit signal)
'
'   *The reason D-4c exists* Quiescence alone cannot distinguish "the app finished" from "the app ignored it".
'   By waiting until the armed marker is observed, the latter can be ruled out.
'   (3) is exactly that check - *it correctly fails in a situation where quiescence alone would succeed*
' ============================================================
Public Sub Test_D4_Signal()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim t0 As Single
    Dim took As Single
    Dim ok As Boolean

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D4_Signal: Browser is not started." & _
                    "Run UserForm1.Show vbModeless and StartWebView2_Full first."
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD4ProbeHtml())
    If p Is Nothing Then
        Wv2Log.LogI "Test_D4_Signal: failed to create the tab."
        Exit Sub
    End If
    If Not D2WaitTitle(p, ("D-4 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&)), 10) Then
        Wv2Log.LogI "Test_D4_Signal: could not confirm that the verification page loaded."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D4_Signal start ================"
    Wv2Log.LogI "  --- (1) arm is fail-fast (you notice on the spot) ---"

    TestBool "*an element that does not exist cannot be armed*", _
             (p.ArmContentSignal("#nope-nope") = False)
    Wv2Log.LogI "        LastEvalError = " & p.LastEvalError
    TestBool "  the reason is not-found", (p.LastEvalError = "not-found")

    TestBool "*an invalid selector cannot be armed either*", (p.ArmContentSignal("###") = False)
    TestBool "  the reason is bad-selector", (p.LastEvalError = "bad-selector")

    TestBool "an existing container can be armed", p.ArmContentSignal("#slot")
    p.DisarmSignals

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) *Arm, then act* ---"
    Wv2Log.LogI "        Arm #slot, and add an element into it after 800ms"

    TestBool "can arm", p.ArmContentSignal("#slot")
    p.EvalSync "window.later(800)"

    t0 = Timer
    ok = p.WaitSettled(5)
    took = D3Took(t0)
    TestBool "*the signal hits and quiescence is reached*", ok
    Wv2Log.LogI "        " & p.LastSettleInfo
    Wv2Log.LogI "        time taken = " & Format$(took, "0.00") & " s"
    TestBool "  the breakdown is dom:hit", (InStr(1, p.LastSettleInfo, "dom:hit") > 0)
    TestBool "  it waited for the setup (0.8 s or more)", (took > 0.8)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) ** Armed but nothing happens ** (the heart of D-4c) ---"
    Wv2Log.LogI "        The page stays quiet. A situation where quiescence alone would succeed"

    TestBool "can arm", p.ArmContentSignal("#slot")
    t0 = Timer
    ok = p.WaitSettled(2)
    took = D3Took(t0)
    TestBool "** False even though quiet (being ignored can be detected) **", (ok = False)
    TestBool "  LastEvalOk=True (not a failure but a timeout)", (p.LastEvalOk = True)
    Wv2Log.LogI "        " & p.LastSettleInfo
    TestBool "  the breakdown is dom:miss", (InStr(1, p.LastSettleInfo, "dom:miss") > 0)
    TestBool "  it hung on for 2 seconds as specified", (took > 1.7)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) One-shot (not carried over to the next wait) ---"

    t0 = Timer
    ok = p.WaitSettled(5)
    took = D3Took(t0)
    TestBool "*the arm is consumed, so quiescence is reached normally*", ok
    TestBool "  the wait time is also normal (under 2 s)", (took < 2)
    Wv2Log.LogI "        " & p.LastSettleInfo
    TestBool "  signal= does not appear in the breakdown", _
             (InStr(1, p.LastSettleInfo, "signal=") = 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) Communication signal ---"

    TestBool "can arm (blob:)", p.ArmNetworkSignal("blob:")
    p.EvalSync "window.chain(100)"
    ok = p.WaitSettled(5)
    TestBool "*the communication signal hits*", ok
    Wv2Log.LogI "        " & p.LastSettleInfo
    TestBool "  the breakdown is net:hit", (InStr(1, p.LastSettleInfo, "net:hit") > 0)

    TestBool "can arm even with a pattern that will not hit", _
             p.ArmNetworkSignal("this-never-matches-xyz")
    p.EvalSync "window.chain(100)"
    ok = p.WaitSettled(3)
    TestBool "*False if it does not hit*", (ok = False)
    Wv2Log.LogI "        " & p.LastSettleInfo
    TestBool "  the breakdown is net:miss", (InStr(1, p.LastSettleInfo, "net:miss") > 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) *signal-lost* (if the arm vanishes, do not quietly fix it) ---"

    TestBool "can arm", p.ArmContentSignal("#slot")
    TestBool "  rebuild the probe (the arm vanishes along with it)", p.SpaProbeReset()
    ok = p.WaitSettled(3)
    TestBool "*False if the arm vanished*", (ok = False)
    Wv2Log.LogI "        " & p.LastSettleInfo
    TestBool "  the breakdown is signal-lost", _
             (InStr(1, p.LastSettleInfo, "signal-lost") > 0)
    TestBool "  *returns at once (does not make you wait)*", True

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (7) DisarmSignals (withdraw by hand) ---"

    TestBool "can arm", p.ArmContentSignal("#slot")
    TestBool "  DisarmSignals succeeds", p.DisarmSignals()
    ok = p.WaitSettled(3)
    TestBool "*if withdrawn, quiescence is reached normally*", ok
    Wv2Log.LogI "        " & p.LastSettleInfo

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback depth (expected 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D4_Signal end ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_D4_Log (verification of D-4e: in-flight ledger / diagnostic log / URL signal)
' ============================================================
Public Sub Test_D4_Log()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim st As String
    Dim n As Long
    Dim il As String

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D4_Log: Browser is not started." & _
                    "Run UserForm1.Show vbModeless and StartWebView2_Full first."
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD4ProbeHtml())
    If p Is Nothing Then
        Wv2Log.LogI "Test_D4_Log: failed to create the tab."
        Exit Sub
    End If
    If Not D2WaitTitle(p, ("D-4 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&)), 10) Then
        Wv2Log.LogI "Test_D4_Log: could not confirm that the verification page loaded."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D4_Log start ================"
    Wv2Log.LogI "  --- (1) Diagnostic log ON / OFF ---"

    TestBool "disabled by default", (p.SpaProbeLogging = False)
    p.SpaProbeLogging = True
    TestBool "  can be enabled", p.SpaProbeLogging

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) *Events accumulate* ---"

    p.EvalSync "window.chain(100)"
    D3Pump 1.5
    st = p.SpaProbeState()
    TestBool "the log count has increased", (Wv2Json.JsonGetNum(st, "lgN") > 0)

    n = p.SpaProbeDrainLog()
    TestBool "*can be fetched*", (n > 0)
    TestBool "  empty once fetched", (p.SpaProbeDrainLog() = 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) *The in-flight ledger* ---"

    ' *Plant a request that has been in flight since 30 seconds ago*
    '   On our own page a blob fetch finishes in 1ms and vanishes within one EvalSync round trip
    '   (38ms). To check the ledger and the lingering judgment deterministically,
    '   we directly push in an item equivalent to a long poll.
    p.EvalSync "(function(){var q=window.__wv2p;" & _
               "q.ifl.push({u:'fake://long-poll',t:performance.now()-30000,w:'xhr'});" & _
               "return 1;})()"

    st = p.SpaProbeState()
    Wv2Json.JsonPickStr st, "il", il
    Wv2Log.LogI "        il = " & il
    TestBool "*the in-flight request appears in the list*", (InStr(1, il, "fake://long-poll") > 0)
    TestBool "  the type and elapsed time are known", (InStr(1, il, "xhr 30") > 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3b) *Lingering judgment (StaleInflightMs)* ---"

    TestBool "the default is 10000ms", (p.StaleInflightMs = 10000)
    TestBool "*a lingering request is excluded from the quiescence judgment (ifn=0)*", _
             (Wv2Json.JsonGetNum(st, "ifn") = 0)
    TestBool "  it is counted as lingering (ifo=1)", _
             (Wv2Json.JsonGetNum(st, "ifo") = 1)
    TestBool "*quiescence is reached even with a lingering request*", p.WaitSettled(5)
    Wv2Log.LogI "        " & p.LastSettleInfo

    p.StaleInflightMs = 0
    st = p.SpaProbeState()
    TestBool "*if disabled, it is counted (ifn=1)*", (Wv2Json.JsonGetNum(st, "ifn") = 1)
    TestBool "  in that case quiescence is not reached", (p.WaitSettled(2) = False)
    Wv2Log.LogI "        " & p.LastSettleInfo
    p.StaleInflightMs = 10000

    ' clean up the planted item
    p.EvalSync "(function(){window.__wv2p.ifl=[];return 1;})()"
    st = p.SpaProbeState()
    Wv2Json.JsonPickStr st, "il", il
    TestBool "  once cleaned up it vanishes from the list", (Len(il) = 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) *URL signal* ---"
    Wv2Log.LogI "        URL of this page = " & p.View_GetSource()

    TestBool "can arm (a string that should hit)", p.ArmUrlSignal("blank")
    TestBool "*quiescence is reached if it is contained in the URL*", p.WaitSettled(5)
    Wv2Log.LogI "        " & p.LastSettleInfo
    TestBool "  the breakdown is url:hit", (InStr(1, p.LastSettleInfo, "url:hit") > 0)

    TestBool "can arm (a string that will not hit)", p.ArmUrlSignal("zzz-not-here")
    TestBool "*False if it is not contained*", (p.WaitSettled(2) = False)
    Wv2Log.LogI "        " & p.LastSettleInfo
    TestBool "  the breakdown is url:miss", (InStr(1, p.LastSettleInfo, "url:miss") > 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) *Show the requests in flight at timeout* (the lesson of D-4d) ---"

    p.SpaProbeLogging = False
    TestBool "  can be returned to disabled", (p.SpaProbeLogging = False)

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback depth (expected 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D4_Log end ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_D4_Site (D-4d: *reconnaissance of Google Maps*)
'
'   *No pass/fail verdict* Real sites go down and their DOM changes without notice (design principle 75).
'   What is done here is devoted to *observing and leaving numbers*. What is judged is only
'   a few structural points such as "could the tab be opened" and "was the search box found".
'
'   What we want to know:
'     (a) *Does quiescence ever come to Maps?* On a page where tile fetching and animation do not stop,
'         how far does q (the time since the last DOM change) grow?
'     (b) *How noisy is it?* DOM changes and communications per second
'     (c) *Round-trip cost on a heavy page* (the production value of unknown 2. It was +0.9ms on a light page)
'     (d) *What becomes the marker of completion?* How document.title and the URL change after a search
'         (the /@latitude,longitude,zoom in the URL is also the very purpose of obtaining coordinates)
'
'   *Accesses an external site for real* The only Test_*. Needs a network.
' ============================================================
Public Sub Test_D4_Site(Optional ByVal searchAddr As String = (ChrW(&H6771&) & ChrW(&H4EAC&) & ChrW(&H90FD&) & ChrW(&H5343&) & ChrW(&H4EE3&) & ChrW(&H7530&) & ChrW(&H533A&) & ChrW(&H4E38&) & ChrW(&H306E&) & ChrW(&H5185&) & "1-9-1"))
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim el As Wv2Element
    Dim ms As Single
    Dim clicked As Boolean

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D4_Site: Browser is not started." & _
                    "Run UserForm1.Show vbModeless and StartWebView2_Full first."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D4_Site start ================"
    Wv2Log.LogI "  *The aim is observation. No pass/fail is given (because real sites change)*"
    Wv2Log.LogI "  address to search: " & searchAddr

    Set p = b.AddTabWithUrl("https://www.google.com/maps")
    If p Is Nothing Then
        Wv2Log.LogI "  [FAIL] failed to create the tab."
        m_ngCount = m_ngCount + 1
        TestCountPrint
        Exit Sub
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (1) Wait for loading (grab the search box with D-3's WaitFor) ---"

    ' *First wait until the Pane is in a state to accept JS*
    '   Right after opening a tab there is no View yet, and EvalSync fails with no-view.
    '   WaitFor "gives up without waiting on failure", so calling it as is fails immediately.
    TestBool "the Pane has come to accept JS", D4WaitPane(p, 30)

    ' *Selectors are tried in order of candidates* The id of a real site may change to an auto-generated one
    '   (measured on 2026-08-22: the former #searchboxinput was gone and the id was ucc-1.
    '    name='q' and the form structure remained, so that one is tried first)
    Set el = D4FindFirst(p, "input[name='q']" & Chr$(1) & _
                            "#searchboxinput" & Chr$(1) & _
                            "form input[type='text']", 30)
    TestBool "the search box was found", Not (el Is Nothing)
    If el Is Nothing Then
        Wv2Log.LogI "  *Not found* There may be a consent screen or a layout change."
        Wv2Log.LogI "        title = " & p.View_GetDocumentTitle()
        Wv2Log.LogI "        url   = " & p.View_GetSource()
        Wv2Log.LogI "        LastEvalOk=" & p.LastEvalOk & " err=" & p.LastEvalError
        TestCountPrint
        Exit Sub
    End If

    Wv2Log.LogI "        title = " & p.View_GetDocumentTitle()
    Wv2Log.LogI "        url   = " & p.View_GetSource()

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) *Round-trip cost (production value of unknown 2)* ---"

    ms = D4RoundTrip(p, 20)
    Wv2Log.LogI "        1 round trip before installing the probe = " & Format$(ms, "0.0") & " ms"
    Wv2Log.LogI "        state = " & p.SpaProbeState()
    ms = D4RoundTrip(p, 20)
    Wv2Log.LogI "        1 round trip after installing the probe = " & Format$(ms, "0.0") & " ms"
    Wv2Log.LogI "        (on a light page it was 35.4 -> 36.3 ms)"

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) *Does quiescence come?* Observe for 15 seconds doing nothing ---"
    D4Watch p, 15

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) Try the quiescence wait (the result is only recorded) ---"

    Wv2Log.LogI "        WaitSettled(10) = " & p.WaitSettled(10)
    Wv2Log.LogI "        " & p.LastSettleInfo

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) Enter the address and search ---"

    ' *Use the tools of D-4e* Turn on the diagnostic log, arm the URL change, then operate
    p.SpaProbeLogging = True
    Wv2Log.LogI "        enabled the diagnostic log"
    TestBool "the URL signal can be armed", p.ArmUrlSignal("/place/")

    Set el = D4FindFirst(p, "input[name='q']" & Chr$(1) & _
                            "#searchboxinput" & Chr$(1) & _
                            "form input[type='text']", 5)
    TestBool "the search box can be grabbed again", Not (el Is Nothing)
    If el Is Nothing Then
        TestCountPrint
        Exit Sub
    End If

    el.value = searchAddr
    TestBool "*the address could be written (D-3's .Value =)*", el.LastOk
    Wv2Log.LogI "        route = " & el.LastInfo
    Wv2Log.LogI "        read back = " & el.value

    ' *Run the search* Press the button if there is one. Otherwise synthesize the Enter key.
    Set el = D4FindFirst(p, ("button[aria-label='" & ChrW(&H691C&) & ChrW(&H7D22&) & "']") & Chr$(1) & _
                            "#searchbox-searchbutton" & Chr$(1) & _
                            "button[aria-label='Search']", 3)
    If el Is Nothing Then
        Wv2Log.LogI "        there is no search button, so synthesizing Enter"
        ' *Make CSS attribute values double quotes* JS strings are single quotes,
        '   and inside the selector double quotes (written with """"), so that
        '   they can be nested without using any backslash (project rule).
        p.EvalSync "(function(){var e=document.querySelector('input[name=""q""]')" & _
                   "||document.getElementById('searchboxinput');" & _
                   "if(!e){return 0;}" & _
                   "e.focus();e.dispatchEvent(new KeyboardEvent('keydown'," & _
                   "{key:'Enter',code:'Enter',keyCode:13,which:13,bubbles:true}));" & _
                   "return 1;})()"
        clicked = p.LastEvalOk
    Else
        clicked = el.Click()
        Wv2Log.LogI "        pressed the search button (route=" & el.LastInfo & ")"
    End If
    TestBool "the search could be run", clicked

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) *Wait with the armed URL signal* ---"

    Wv2Log.LogI "        WaitSettled(20) = " & p.WaitSettled(20)
    Wv2Log.LogI "        " & p.LastSettleInfo
    Wv2Log.LogI "        *On timeout, the requests lingering in 'in flight:' appear*"

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6b) *Add individual exclusions and wait again* (a demonstration of plan A) ---"
    Wv2Log.LogI "        Exclude the lingering request by name from the quiescence judgment"

    TestBool "AddIgnoreNetwork works", p.AddIgnoreNetwork("/search?tbm=map")
    Wv2Log.LogI "        WaitSettled(10) = " & p.WaitSettled(10)
    Wv2Log.LogI "        " & p.LastSettleInfo
    p.ClearIgnoreNetwork

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6c) 10 seconds after the search ---"
    D4Watch p, 10

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6d) *Diagnostic log* what was happening ---"
    Wv2Log.LogI "        number fetched: " & p.SpaProbeDrainLog()
    Wv2Log.LogI "        (cut off at 60 items. The rest can be fetched on the next call)"
    Wv2Log.LogI "        number fetched: " & p.SpaProbeDrainLog()
    p.SpaProbeLogging = False

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (7) Result ---"
    Wv2Log.LogI "        title = " & p.View_GetDocumentTitle()
    Wv2Log.LogI "        url   = " & p.View_GetSource()
    Wv2Log.LogI "        *If /@latitude,longitude,zoom is in the URL, the coordinates can be obtained*"
    Wv2Log.LogI "        WaitSettled(10) = " & p.WaitSettled(10)
    Wv2Log.LogI "        " & p.LastSettleInfo
    Wv2Log.LogI "        state = " & p.SpaProbeState()

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback depth (expected 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "  NOTE: the number of verdicts here is only a few structural points. Read the body of the log for the content of the observation."
    Wv2Log.LogI "================ Test_D4_Site end ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' D4FindFirst (D-4d: try candidate selectors in order and return the first that hits)
'
'   *Selectors of real sites change without notice (design principle 75)*
'   Do not bet on one; line up candidates and *leave in the log which one hit*. The next time it breaks
'   you can tell "what changed". Candidates are passed separated by Chr$(1).
'
'   Only the first one waits for timeoutSec (assuming the page is still loading).
'   The second and later are just looked at in one round trip. If all miss, Nothing.
' ============================================================
Private Function D4FindFirst(ByVal p As Wv2Pane, _
                             ByVal selectorList As String, _
                             ByVal timeoutSec As Single) As Wv2Element
    Dim cands As Variant
    Dim i As Long
    Dim el As Wv2Element

    cands = Split(selectorList, Chr$(1))

    For i = LBound(cands) To UBound(cands)
        If i = LBound(cands) Then
            Set el = p.WaitFor(CStr(cands(i)), timeoutSec)
        Else
            Set el = p.QuerySelector(CStr(cands(i)))
        End If

        If Not el Is Nothing Then
            Wv2Log.LogI "        found with selector [" & cands(i) & "]"
            Set D4FindFirst = el
            Exit Function
        End If

        Wv2Log.LogI "        selector [" & cands(i) & "] missed" & _
                    IIf(p.LastEvalOk, "", " (failed: " & p.LastEvalError & ")")
    Next i
End Function

' ============================================================
' D4WaitPane (D-4d: wait until the Pane accepts JS)
'
'   *Right after opening a tab EvalSync fails with no-view*
'   WaitFor of D-3 and WaitSettled of D-4 both follow the policy that "a failure is not fixed by waiting"
'   (design principle 104), so calling them before preparation makes them give up at once.
'   For rounds that open a real site, insert this first.
' ============================================================
Private Function D4WaitPane(ByVal p As Wv2Pane, ByVal timeoutSec As Single) As Boolean
    Dim t0 As Single

    t0 = Timer
    Do
        p.EvalSync "1", 3
        If p.LastEvalOk Then
            D4WaitPane = True
            Exit Function
        End If
        If D3Took(t0) >= timeoutSec Then
            Wv2Log.LogW "D4WaitPane: timed out err=" & p.LastEvalError
            Exit Function
        End If
        D3Pump 0.3
    Loop
End Function


' ============================================================
' D4Watch (D-4d: record the state, title and URL every second)
'
'   *A tool for seeing whether quiescence comes* If q keeps growing, it has gone quiet.
'   The title and URL are taken via COM (EvalSync is not used, so the observation is light).
' ============================================================
Private Sub D4Watch(ByVal p As Wv2Pane, ByVal seconds As Long)
    Dim i As Long
    Dim st As String
    Dim ttl As String
    Dim url As String

    For i = 1 To seconds
        st = p.SpaProbeState()
        ttl = p.View_GetDocumentTitle()
        url = p.View_GetSource()

        Wv2Log.LogI "        [" & Format$(i, "00") & "s] " & _
                    "q=" & Wv2Json.JsonGetNum(st, "q") & _
                    " m=" & Wv2Json.JsonGetNum(st, "m") & _
                    " n=" & Wv2Json.JsonGetNum(st, "n") & _
                    " f=" & Wv2Json.JsonGetNum(st, "f") & _
                    " x=" & Wv2Json.JsonGetNum(st, "x") & _
                    " rp=" & Wv2Json.JsonGetNum(st, "rp") & _
                    " | " & Left$(ttl, 40) & " | " & Left$(url, 70)

        D3Pump 1
    Next i
End Sub

' ============================================================
' Test_D4_Dom (D-4d supplement: *peek at the DOM of the active tab*)
'
'   *A tool for deciding "what to wait for" by observation rather than guesswork*
'   Selectors of real sites change without notice (design principle 75). Before writing
'   selectors by guesswork, look at what is actually there.
'
'   Usage: display the target page in the browser, then in the Immediate window
'           Test_D4_Dom           ... lists input fields, buttons and elements with a role
'           Test_D4_Dom "h1,h2"  ... lists by the specified selector
'
'   Information output: tag / id / name / type / class (first 40 chars) / placeholder /
'             aria-label / whether visible / text (first 30 chars)
'
'   *The prototype of D-4e (diagnostic log)* It does not take a time series, but for first learning
'   "what is there right now" this is enough.
' ============================================================
Public Sub Test_D4_Dom(Optional ByVal sel As String = "")
    Dim p As Wv2Pane
    Dim js As String
    Dim res As String
    Dim parts As Variant
    Dim i As Long

    Set p = UserForm1.GetActivePane
    If p Is Nothing Then
        Wv2Log.LogI "Test_D4_Dom: there is no active Pane."
        Exit Sub
    End If

    If Len(sel) = 0 Then
        sel = "input,textarea,select,[role=combobox],[role=searchbox]," & _
              "[role=search],button[aria-label],form"
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "================ Test_D4_Dom start ================"
    Wv2Log.LogI "  url   = " & p.View_GetSource()
    Wv2Log.LogI "  title = " & p.View_GetDocumentTitle()
    Wv2Log.LogI "  selector: " & sel

    js = "(function(){var out=[],q=document.querySelectorAll(" & _
         p.JsQuote(sel) & ");" & _
         "out.push('total '+q.length+' items');" & _
         "for(var i=0;i<q.length&&out.length<41;i++){var e=q[i];" & _
         "var g=function(a){var v=(e.getAttribute?e.getAttribute(a):null);" & _
         "return v?String(v).slice(0,40):'-';};" & _
         "out.push(e.tagName" & _
         "+' id='+(e.id||'-')" & _
         "+' name='+(e.name||'-')" & _
         "+' type='+(e.type||'-')" & _
         "+' cls='+String(e.className||'-').slice(0,40)" & _
         "+' ph='+g('placeholder')" & _
         "+' aria='+g('aria-label')" & _
         "+' vis='+(e.offsetParent!==null)" & _
         "+' txt='+String(e.textContent||'').trim().slice(0,30));}" & _
         "return out.join(String.fromCharCode(1));})()"

    res = p.EvalSync(js, 10)
    If Not p.LastEvalOk Then
        Wv2Log.LogI "  *failed* err=" & p.LastEvalError
        Wv2Log.LogI "================ Test_D4_Dom end ================"
        Exit Sub
    End If

    parts = Split(Wv2Json.JsonUnescape(res), ChrW$(1))
    For i = LBound(parts) To UBound(parts)
        Wv2Log.LogI "  " & parts(i)
    Next i

    Wv2Log.LogI "================ Test_D4_Dom end ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_D5_Geocode (verification of D-5: address -> coordinates)
'
'   *Checks whether the parts of the D axis have become usable for business work*
'   Processes 3 addresses *in succession in the same tab*, and outputs latitude/longitude and the normalized name.
'
'   *Accesses an external site for real* Needs a network.
'   It is a real site, so the verdict is lenient (whether the coordinates fall within a plausible range).
' ============================================================
Public Sub Test_D5_Geocode()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim addrs As Variant
    Dim i As Long
    Dim lat As Double
    Dim lng As Double
    Dim nm As String
    Dim ok As Boolean
    Dim t0 As Single

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D5_Geocode: Browser is not started." & _
                    "Run UserForm1.Show vbModeless and StartWebView2_Full first."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D5_Geocode start ================"

    t0 = Timer
    Set p = Wv2Maps.MapsOpen(b)
    TestBool "*Maps could be opened into an operable state*", Not (p Is Nothing)
    Wv2Log.LogI "        time taken = " & Format$(D3Took(t0), "0.0") & " s"
    If p Is Nothing Then
        Wv2Log.LogI "        reason = " & Wv2Maps.MapsLastError
        TestCountPrint
        Exit Sub
    End If

    addrs = Array( _
        (ChrW(&H6771&) & ChrW(&H4EAC&) & ChrW(&H90FD&) & ChrW(&H5343&) & ChrW(&H4EE3&) & ChrW(&H7530&) & ChrW(&H533A&) & ChrW(&H4E38&) & ChrW(&H306E&) & ChrW(&H5185&) & "1-9-1"), _
        (ChrW(&H5927&) & ChrW(&H962A&) & ChrW(&H5E9C&) & ChrW(&H5927&) & ChrW(&H962A&) & ChrW(&H5E02&) & ChrW(&H4E2D&) & ChrW(&H592E&) & ChrW(&H533A&) & ChrW(&H5927&) & ChrW(&H962A&) & ChrW(&H57CE&) & "1-1"), _
        (ChrW(&H5317&) & ChrW(&H6D77&) & ChrW(&H9053&) & ChrW(&H672D&) & ChrW(&H5E4C&) & ChrW(&H5E02&) & ChrW(&H4E2D&) & ChrW(&H592E&) & ChrW(&H533A&) & ChrW(&H5317&) & "1" & ChrW(&H6761&) & ChrW(&H897F&) & "2" & ChrW(&H4E01&) & ChrW(&H76EE&)))

    For i = LBound(addrs) To UBound(addrs)
        Wv2Log.LogI ""
        Wv2Log.LogI "  --- (" & (i + 1) & ") " & addrs(i) & " ---"

        t0 = Timer
        ok = Wv2Maps.MapsGeocode(p, CStr(addrs(i)), lat, lng, nm)

        Wv2Log.LogI "        time taken = " & Format$(D3Took(t0), "0.0") & " s"
        Wv2Log.LogI "        latitude/longitude = " & lat & ", " & lng
        Wv2Log.LogI "        name     = " & nm
        If Not ok Then Wv2Log.LogI "        reason   = " & Wv2Maps.MapsLastError

        TestBool "  *settled to exactly 1*", ok
        TestBool "  latitude is in the range of Japan (20-46)", (lat > 20 And lat < 46)
        TestBool "  longitude is in the range of Japan (122-154)", (lng > 122 And lng < 154)
        TestBool "  the name was obtained", (Len(nm) > 0)
    Next i

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- Handling of failure ---"

    ok = Wv2Maps.MapsGeocode(p, ("ZZZZ" & ChrW(&H5B58&) & ChrW(&H5728&) & ChrW(&H3057&) & ChrW(&H306A&) & ChrW(&H3044&) & ChrW(&H4F4F&) & ChrW(&H6240&) & "ZZZZ"), lat, lng, nm, 8)
    Wv2Log.LogI "        return value = " & ok & " reason = " & Wv2Maps.MapsLastError
    TestBool "*a nonsense address is not settled*", (ok = False)

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback depth (expected 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D5_Geocode end ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_D5_Sheet (verification of D-5b: sheet integration)
'
'   *Create a new workbook to try it and close it without saving*
'   If a sheet is added to the development workbook, Excel asks whether to save on exit, which is annoying.
'
'   *Accesses an external site for real* Needs a network.
' ============================================================
Public Sub Test_D5_Sheet()
    Dim wb As Workbook
    Dim sh As Object
    Dim n As Long

    If UserForm1.CurrentBrowser Is Nothing Then
        Wv2Log.LogI "Test_D5_Sheet: Browser is not started."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D5_Sheet start ================"

    Set wb = Workbooks.Add
    Set sh = wb.Worksheets(1)

    sh.Cells(1, 1).value = (ChrW(&H4F4F&) & ChrW(&H6240&))
    sh.Cells(1, 2).value = (ChrW(&H7DEF&) & ChrW(&H5EA6&))
    sh.Cells(1, 3).value = (ChrW(&H7D4C&) & ChrW(&H5EA6&))
    sh.Cells(1, 4).value = (ChrW(&H6B63&) & ChrW(&H898F&) & ChrW(&H5316&) & ChrW(&H4F4F&) & ChrW(&H6240&))
    sh.Cells(1, 5).value = (ChrW(&H72B6&) & ChrW(&H614B&))
    sh.Cells(2, 1).value = (ChrW(&H6771&) & ChrW(&H4EAC&) & ChrW(&H90FD&) & ChrW(&H5343&) & ChrW(&H4EE3&) & ChrW(&H7530&) & ChrW(&H533A&) & ChrW(&H4E38&) & ChrW(&H306E&) & ChrW(&H5185&) & "1-9-1")
    sh.Cells(3, 1).value = (ChrW(&H5927&) & ChrW(&H962A&) & ChrW(&H5E9C&) & ChrW(&H5927&) & ChrW(&H962A&) & ChrW(&H5E02&) & ChrW(&H4E2D&) & ChrW(&H592E&) & ChrW(&H533A&) & ChrW(&H5927&) & ChrW(&H962A&) & ChrW(&H57CE&) & "1-1")
    sh.Cells(4, 1).value = ("ZZZZ" & ChrW(&H5B58&) & ChrW(&H5728&) & ChrW(&H3057&) & ChrW(&H306A&) & ChrW(&H3044&) & ChrW(&H4F4F&) & ChrW(&H6240&) & "ZZZZ")

    Wv2Log.LogI "  Process 3 rows (1 of which is nonsense)"
    n = Wv2Maps.MapsGeocodeSheet(sh)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- Contents written ---"
    Wv2Log.LogI "  row 2: " & sh.Cells(2, 2).value & " / " & _
                sh.Cells(2, 3).value & " / " & sh.Cells(2, 4).value & " / " & sh.Cells(2, 5).value
    Wv2Log.LogI "  row 3: " & sh.Cells(3, 2).value & " / " & _
                sh.Cells(3, 3).value & " / " & sh.Cells(3, 4).value & " / " & sh.Cells(3, 5).value
    Wv2Log.LogI "  row 4: " & sh.Cells(4, 2).value & " / " & _
                sh.Cells(4, 3).value & " / " & sh.Cells(4, 4).value & " / " & sh.Cells(4, 5).value

    TestBool "*2 rows became ok*", (n = 2)
    TestBool "  the latitude of row 2 is Tokyo", (Abs(sh.Cells(2, 2).value - 35.68) < 0.1)
    TestBool "  the latitude of row 3 is Osaka", (Abs(sh.Cells(3, 2).value - 34.69) < 0.1)
    TestBool "  the status of row 2 is ok", (sh.Cells(2, 5).value = "ok")
    TestBool "*the nonsense row does not become ok*", (sh.Cells(4, 5).value <> "ok")
    TestBool "  its reason is kept", (Len(sh.Cells(4, 5).value) > 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- *Can it resume?* (rows that are ok are skipped) ---"
    n = Wv2Maps.MapsGeocodeSheet(sh)
    TestBool "the 2nd run returns the same count (counted even when skipped)", (n = 2)

    wb.Close SaveChanges:=False

    Wv2Log.LogI ""
    TestCountPrint
    Wv2Log.LogI "================ Test_D5_Sheet end ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_D6_All (verification of D-6: QuerySelectorAll and lifetime management)
' ============================================================
Public Sub Test_D6_All()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim els As Collection
    Dim el As Wv2Element
    Dim ids As String
    Dim n0 As Long

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D6_All: Browser is not started."
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD4ProbeHtml())
    If p Is Nothing Then
        Wv2Log.LogI "Test_D6_All: failed to create the tab."
        Exit Sub
    End If
    If Not D2WaitTitle(p, ("D-4 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&)), 10) Then
        Wv2Log.LogI "Test_D6_All: could not confirm that the verification page loaded."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D6_All start ================"
    Wv2Log.LogI "  --- (1) Grab them all together ---"

    Set els = p.QuerySelectorAll("div")
    TestBool "*a Collection is returned*", Not (els Is Nothing)
    TestBool "  there are 3", (els.Count = 3)
    TestBool "  not truncated", (p.LastAllTruncated = False)

    ' *The point of issue 2 is that it can be looped over naturally with For Each*
    For Each el In els
        ids = ids & el.GetAttribute("id") & ","
    Next el
    Wv2Log.LogI "        ids in document order = " & ids
    TestBool "*returned in document order*", (ids = "target,noise,slot,")

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) Distinguishing 0 results from failure (design principle 93) ---"

    Set els = p.QuerySelectorAll("blockquote")
    TestBool "a Collection is returned even for 0 results (not Nothing)", Not (els Is Nothing)
    TestBool "  Count is 0", (els.Count = 0)
    TestBool "  *LastEvalOk=True (really none)*", (p.LastEvalOk = True)

    Set els = p.QuerySelectorAll("###")
    TestBool "a Collection is returned even for an invalid selector", Not (els Is Nothing)
    TestBool "  Count is 0", (els.Count = 0)
    TestBool "  *LastEvalOk=False (failure)*", (p.LastEvalOk = False)
    Wv2Log.LogI "        LastEvalError = " & p.LastEvalError

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) *Truncation is not done silently* ---"

    Set els = p.QuerySelectorAll("div", 2)
    TestBool "2 items as per the limit", (els.Count = 2)
    TestBool "*it is known that it was truncated*", p.LastAllTruncated

    Set els = p.QuerySelectorAll("div", 200)
    TestBool "  returns to False when within the limit", (p.LastAllTruncated = False)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) The grabbed elements can be used as they are ---"

    Set els = p.QuerySelectorAll("div")
    Set el = els(1)
    TestEq "the first item can be read", el, el.GetAttribute("id"), "target"
    TestBool "  can write (SetAttribute of D-3)", _
             el.SetAttribute("data-mark", "1")
    TestEq "  can be read back", el, el.GetAttribute("data-mark"), "1"

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) *Lifetime management only shows* ---"

    n0 = p.ElementCount
    Wv2Log.LogI "        number now accumulated in the registry = " & n0
    TestBool "ElementCount can be read", (n0 > 0)

    Set els = p.QuerySelectorAll("div")
    TestBool "*it piles up with every grab (nothing is discarded on its own)*", (p.ElementCount > n0)

    TestBool "it can be rebuilt with ClearElementRegistry", p.ClearElementRegistry()
    TestBool "  the registry becomes empty", (p.ElementCount = 0)
    TestBool "  *the element in hand becomes stale*", (el.IsStale = True)

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback depth (expected 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D6_All end ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_D6_Pick (verification of D-6b: take the first from the candidate list)
'
'   *An exit to see whether D-6's QuerySelectorAll works on a real site*
'   A category search (e.g. "convenience store near Tokyo Station") produces a candidate list, so
'   we confirm that the result changes depending on whether pickFirst is used.
'
'   *Accesses an external site for real* It is a real site, so the verdict is lenient.
' ============================================================
Public Sub Test_D6_Pick()
    Dim p As Wv2Pane
    Dim lat As Double
    Dim lng As Double
    Dim nm As String
    Dim ok As Boolean

    If UserForm1.CurrentBrowser Is Nothing Then
        Wv2Log.LogI "Test_D6_Pick: Browser is not started."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D6_Pick start ================"

    Set p = Wv2Maps.MapsOpen(UserForm1.CurrentBrowser)
    TestBool "Maps could be opened", Not (p Is Nothing)
    If p Is Nothing Then
        TestCountPrint
        Exit Sub
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (1) A search that yields multiple candidates (do not pick) ---"

    ok = Wv2Maps.MapsGeocode(p, (ChrW(&H30B3&) & ChrW(&H30F3&) & ChrW(&H30D3&) & ChrW(&H30CB&) & " " & ChrW(&H6771&) & ChrW(&H4EAC&) & ChrW(&H99C5&)), lat, lng, nm, 15)
    Wv2Log.LogI "        return value=" & ok & " reason=" & Wv2Maps.MapsLastError
    Wv2Log.LogI "        coordinates=" & lat & "," & lng & "  name=" & nm
    TestBool "*it is not settled if we do not pick*", (ok = False)
    TestBool "  not picked from the candidates", (Wv2Maps.MapsLastPicked = False)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) *Take the first of the candidates* ---"

    ok = Wv2Maps.MapsGeocode(p, (ChrW(&H30B3&) & ChrW(&H30F3&) & ChrW(&H30D3&) & ChrW(&H30CB&) & " " & ChrW(&H6771&) & ChrW(&H4EAC&) & ChrW(&H99C5&)), lat, lng, nm, 15, True)
    Wv2Log.LogI "        return value=" & ok & " reason=" & Wv2Maps.MapsLastError
    Wv2Log.LogI "        coordinates=" & lat & "," & lng
    Wv2Log.LogI "        name=" & nm
    Wv2Log.LogI "        picked from the candidates=" & Wv2Maps.MapsLastPicked

    TestBool "*it is settled if we take a candidate*", ok
    If ok Then
        TestBool "  *the mark of being picked from the candidates is set*", Wv2Maps.MapsLastPicked
        TestBool "  latitude is in the range of Japan", (lat > 20 And lat < 46)
        TestBool "  longitude is in the range of Japan", (lng > 122 And lng < 154)
        TestBool "  the name was obtained", (Len(nm) > 0)
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) No mark is set for an address that settles ---"

    ok = Wv2Maps.MapsGeocode(p, (ChrW(&H6771&) & ChrW(&H4EAC&) & ChrW(&H90FD&) & ChrW(&H5343&) & ChrW(&H4EE3&) & ChrW(&H7530&) & ChrW(&H533A&) & ChrW(&H4E38&) & ChrW(&H306E&) & ChrW(&H5185&) & "1-9-1"), lat, lng, nm, 15, True)
    TestBool "it settles", ok
    TestBool "*the mark of being picked from the candidates is not set*", (Wv2Maps.MapsLastPicked = False)
    Wv2Log.LogI "        coordinates=" & lat & "," & lng & "  name=" & nm

    Wv2Log.LogI ""
    Wv2Log.LogI "  number accumulated in the registry: " & p.ElementCount
    Wv2Log.LogI "  in-callback depth (expected 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D6_Pick end ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_D7_Cancel (verification of D-7: the abort entry point and the denominator)
'
'   *Needs neither the network nor the browser* The verification of aborting while hitting a real site is
'   Test_D7_Sheet (press Esc by hand).
' ============================================================
Public Sub Test_D7_Cancel()
    Dim wb As Workbook
    Dim sh As Object
    Dim sh2 As Object
    Dim lat As Double
    Dim lng As Double
    Dim nm As String
    Dim ok As Boolean
    Dim n As Long
    Dim beforeTabs As Long

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D7_Cancel start ================"

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (1) Entry point to raise an abort from outside ---"

    Wv2Maps.MapsCancel = True
    TestBool "*if raised, it stands*", (Wv2Maps.MapsCancel = True)
    Wv2Maps.MapsCancel = False
    TestBool "  if lowered, it goes down", (Wv2Maps.MapsCancel = False)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) *The abort state is discarded at entry* (design principle 106) ---"
    Wv2Log.LogI "        Leave it raised and see that the next call is not killed"

    Wv2Maps.MapsCancel = True
    ok = Wv2Maps.MapsGeocode(Nothing, (ChrW(&H6771&) & ChrW(&H4EAC&) & ChrW(&H99C5&)), lat, lng, nm, 1)
    Wv2Log.LogI "        return value = " & ok & " reason = " & Wv2Maps.MapsLastError
    TestBool "can be called (the body fails with no-pane)", (ok = False)
    TestBool "*the reason is not canceled* (because it was discarded at entry)", _
             (Wv2Maps.MapsLastError = "no-pane")
    TestBool "  the abort request is lowered", (Wv2Maps.MapsCancel = False)
    TestBool "  the mark of having ended by abort is also lowered", (Wv2Maps.MapsCanceled = False)

    ' --- D-7b: *If you forget to restore it, Esc stops working in all of Excel* ---
    Wv2Log.LogI "        EnableCancelKey = " & Application.EnableCancelKey & _
                " (xlInterrupt = " & xlInterrupt & ")"
    TestBool "*EnableCancelKey is restored*", _
             (Application.EnableCancelKey = xlInterrupt)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) *Count the denominator* (MapsCountRows) ---"

    Set wb = Workbooks.Add
    Set sh = wb.Worksheets(1)
    sh.Cells(1, 1).value = (ChrW(&H4F4F&) & ChrW(&H6240&))
    sh.Cells(2, 1).value = (ChrW(&H6771&) & ChrW(&H4EAC&) & ChrW(&H90FD&) & ChrW(&H5343&) & ChrW(&H4EE3&) & ChrW(&H7530&) & ChrW(&H533A&) & ChrW(&H4E38&) & ChrW(&H306E&) & ChrW(&H5185&) & "1-9-1")
    sh.Cells(3, 1).value = (ChrW(&H5927&) & ChrW(&H962A&) & ChrW(&H5E9C&) & ChrW(&H5927&) & ChrW(&H962A&) & ChrW(&H5E02&) & ChrW(&H4E2D&) & ChrW(&H592E&) & ChrW(&H533A&) & ChrW(&H5927&) & ChrW(&H962A&) & ChrW(&H57CE&) & "1-1")
    sh.Cells(4, 1).value = (ChrW(&H4EAC&) & ChrW(&H90FD&) & ChrW(&H5E9C&) & ChrW(&H4EAC&) & ChrW(&H90FD&) & ChrW(&H5E02&) & ChrW(&H4E0B&) & ChrW(&H4EAC&) & ChrW(&H533A&) & ChrW(&H6771&) & ChrW(&H5869&) & ChrW(&H5C0F&) & ChrW(&H8DEF&) & ChrW(&H753A&))
    ' Row 5 is empty - *it stops here*
    sh.Cells(6, 1).value = (ChrW(&H7A7A&) & ChrW(&H767D&) & ChrW(&H306E&) & ChrW(&H5411&) & ChrW(&H3053&) & ChrW(&H3046&) & ChrW(&H306F&) & ChrW(&H6570&) & ChrW(&H3048&) & ChrW(&H306A&) & ChrW(&H3044&))

    TestBool "*stops at the empty cell and counts 3*", (Wv2Maps.MapsCountRows(sh) = 3)
    TestBool "  starting from row 3 gives 2", (Wv2Maps.MapsCountRows(sh, 3) = 2)
    TestBool "  looking at a column with no address gives 0", (Wv2Maps.MapsCountRows(sh, 2, 5) = 0)
    TestBool "  if there is no sheet, 0", (Wv2Maps.MapsCountRows(Nothing) = 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) *With 0 rows it returns without opening a tab* ---"

    Set sh2 = wb.Worksheets.Add
    beforeTabs = -1
    If Not UserForm1.CurrentBrowser Is Nothing Then
        beforeTabs = UserForm1.CurrentBrowser.TabCount
    End If

    n = Wv2Maps.MapsGeocodeSheet(sh2)
    TestBool "returns 0", (n = 0)
    TestBool "  the reason is no-rows", (Wv2Maps.MapsLastError = "no-rows")

    If beforeTabs >= 0 Then
        Wv2Log.LogI "        tab count " & beforeTabs & " -> " & _
                    UserForm1.CurrentBrowser.TabCount
        TestBool "*the tabs did not increase*", _
                 (UserForm1.CurrentBrowser.TabCount = beforeTabs)
    Else
        Wv2Log.LogI "        (the browser is not started, so the tab-count verdict is skipped)"
    End If

    Application.StatusBar = Empty   ' D-7d: clear what the previous test left before looking
    TestBool "  the status bar is not dirty", _
             (TypeName(Application.StatusBar) = "Boolean")

    wb.Close SaveChanges:=False

    Wv2Log.LogI ""
    TestCountPrint
    Wv2Log.LogI "================ Test_D7_Cancel end ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_D7_Sheet (verification of D-7b: *press Esc by hand during execution*)
'
'   *Accesses an external site for real* Needs a network.
'   *Create a new workbook to try it and close it without saving*
'
'   Procedure:
'     1) When run, it starts processing 4 addresses in order (about 8 seconds each).
'     2) *Press Esc around the 2nd one, watching the status bar*
'        There is no need to hold it down. The focus may be anywhere.
'        *From D-7b on, "code execution has been interrupted" does not appear* (it is
'        received as error 18 and turned into an abort). If it appears, D-7b is not working.
'     3) Judge that it was aborted and that *nothing was written to the row where it was aborted*.
'     4) Call it again as is and judge that *it fills in from where it left off* (design principle 110).
'
'   *It FAILs if you do not press Esc* If you forgot to press, run it again.
'   *All addresses are chosen to "settle to exactly 1"* (in D-7 an ambiguous
'   address was mixed in, each took 27 seconds, and the verdict became indirect).
' ============================================================
Public Sub Test_D7_Sheet()
    Dim wb As Workbook
    Dim sh As Object
    Dim n As Long
    Dim n2 As Long
    Dim i As Long
    Dim blanks As Long
    Dim wasCanceled As Boolean

    If UserForm1.CurrentBrowser Is Nothing Then
        Wv2Log.LogI "Test_D7_Sheet: Browser is not started."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D7_Sheet start ================"
    Wv2Log.LogI "  *** Please press Esc around the 2nd one ***"
    Debug.Print ""
    Debug.Print "  *** Please press Esc around the 2nd one ***"
    Debug.Print "  (No need to mash it. Once is enough)"

    Set wb = Workbooks.Add
    Set sh = wb.Worksheets(1)
    sh.Cells(1, 1).value = (ChrW(&H4F4F&) & ChrW(&H6240&))
    sh.Cells(2, 1).value = (ChrW(&H6771&) & ChrW(&H4EAC&) & ChrW(&H90FD&) & ChrW(&H5343&) & ChrW(&H4EE3&) & ChrW(&H7530&) & ChrW(&H533A&) & ChrW(&H4E38&) & ChrW(&H306E&) & ChrW(&H5185&) & "1-9-1")
    sh.Cells(3, 1).value = (ChrW(&H5927&) & ChrW(&H962A&) & ChrW(&H5E9C&) & ChrW(&H5927&) & ChrW(&H962A&) & ChrW(&H5E02&) & ChrW(&H4E2D&) & ChrW(&H592E&) & ChrW(&H533A&) & ChrW(&H5927&) & ChrW(&H962A&) & ChrW(&H57CE&) & "1-1")
    sh.Cells(4, 1).value = (ChrW(&H5317&) & ChrW(&H6D77&) & ChrW(&H9053&) & ChrW(&H672D&) & ChrW(&H5E4C&) & ChrW(&H5E02&) & ChrW(&H5317&) & ChrW(&H533A&) & ChrW(&H5317&) & "6" & ChrW(&H6761&) & ChrW(&H897F&) & "4" & ChrW(&H4E01&) & ChrW(&H76EE&))
    sh.Cells(5, 1).value = (ChrW(&H5BAE&) & ChrW(&H57CE&) & ChrW(&H770C&) & ChrW(&H4ED9&) & ChrW(&H53F0&) & ChrW(&H5E02&) & ChrW(&H9752&) & ChrW(&H8449&) & ChrW(&H533A&) & ChrW(&H4E2D&) & ChrW(&H592E&) & "1-1-1")

    n = Wv2Maps.MapsGeocodeSheet(sh)
    wasCanceled = Wv2Maps.MapsCanceled

    Wv2Log.LogI ""
    Wv2Log.LogI "        ok = " & n & " / aborted = " & wasCanceled & _
                " / reason = " & Wv2Maps.MapsLastError
    TestDumpStatusBar "StatusBar"
    Wv2Log.LogI "        *calls of abort detection = " & Wv2Maps.MapsCheckCount & " times*"

    TestBool "*it was aborted* (FAIL if you did not press. Please try once more)", wasCanceled
    TestBool "  the reason is canceled", (Wv2Maps.MapsLastError = "canceled")
    ' D-7d: restoring with Empty is now used, so the verdict was brought back
    TestBool "*the status bar is restored*", _
             (TypeName(Application.StatusBar) = "Boolean")
    TestBool "*EnableCancelKey is restored*", _
             (Application.EnableCancelKey = xlInterrupt)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- *Did it leave a half-written row?* (design principle 111) ---"

    blanks = 0
    For i = 2 To 5
        If Len(CStr(sh.Cells(i, 5).value)) = 0 Then
            blanks = blanks + 1
            TestBool "  row " & i & " is untouched (coordinates also empty)", _
                     (Len(CStr(sh.Cells(i, 2).value)) = 0 And _
                      Len(CStr(sh.Cells(i, 3).value)) = 0)
        End If
    Next i
    Wv2Log.LogI "        untouched rows = " & blanks & " items"
    ' *This is the direct evidence of "it stopped midway"* (n < 4 is an indirect indicator so it is not used)
    TestBool "*there are untouched rows*", (blanks >= 1)

    ' --- D-7e: *show a human that it stopped, and wait for the finger to come off* ---
    Wv2Log.LogI ""
    Debug.Print ""
    If wasCanceled Then
        Debug.Print "  *** It stopped (aborted after processing " & n & " items) ***"
        Wv2Log.LogI "  *** It stopped (aborted after processing " & n & " items) ***"
    Else
        Debug.Print "  *It did not stop* Did you press Esc?"
        Wv2Log.LogI "  *It did not stop*"
    End If
    Debug.Print "  *Please take your finger off Esc* We will try to resume once it is quiet"
    TestWaitEscReleased 15

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- *Call again and continue from where it left off* (this time please do not press Esc) ---"
    Debug.Print "  --- Resuming (please do not press Esc) ---"

    n2 = Wv2Maps.MapsGeocodeSheet(sh)
    Wv2Log.LogI "        ok of the 2nd run = " & n2 & " / aborted = " & Wv2Maps.MapsCanceled

    blanks = 0
    For i = 2 To 5
        Wv2Log.LogI "        row " & i & ": " & sh.Cells(i, 2).value & " / " & _
                    sh.Cells(i, 3).value & " / " & sh.Cells(i, 5).value
        If Len(CStr(sh.Cells(i, 5).value)) = 0 Then blanks = blanks + 1
    Next i

    TestBool "*the 2nd run was not aborted* (discarded at entry)", _
             (Wv2Maps.MapsCanceled = False)
    TestBool "*the untouched rows are gone*", (blanks = 0)
    TestBool "  ok has not decreased", (n2 >= n)
    TestBool "  EnableCancelKey is restored", _
             (Application.EnableCancelKey = xlInterrupt)
    TestBool "  the status bar is restored", _
             (TypeName(Application.StatusBar) = "Boolean")
    TestDumpStatusBar "StatusBar"

    wb.Close SaveChanges:=False

    Wv2Log.LogI ""
    TestCountPrint
    Wv2Log.LogI "================ Test_D7_Sheet end ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_D7_Resume (verification of D-7b: *can it resume from the post-abort state?*)
'
'   *No need to press Esc* Artificially builds a sheet as it is after an abort -- the first half
'   filled in and the second half untouched -- and calls again from there.
'   This lets the regression check of "resume" run automatically every time.
'
'   *Accesses an external site for real* Only 2 items are processed (about 15 seconds).
' ============================================================
Public Sub Test_D7_Resume()
    Dim wb As Workbook
    Dim sh As Object
    Dim n As Long
    Dim i As Long
    Dim blanks As Long

    If UserForm1.CurrentBrowser Is Nothing Then
        Wv2Log.LogI "Test_D7_Resume: Browser is not started."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D7_Resume start ================"

    Set wb = Workbooks.Add
    Set sh = wb.Worksheets(1)
    sh.Cells(1, 1).value = (ChrW(&H4F4F&) & ChrW(&H6240&))
    sh.Cells(2, 1).value = (ChrW(&H6771&) & ChrW(&H4EAC&) & ChrW(&H90FD&) & ChrW(&H5343&) & ChrW(&H4EE3&) & ChrW(&H7530&) & ChrW(&H533A&) & ChrW(&H4E38&) & ChrW(&H306E&) & ChrW(&H5185&) & "1-9-1")
    sh.Cells(3, 1).value = (ChrW(&H5927&) & ChrW(&H962A&) & ChrW(&H5E9C&) & ChrW(&H5927&) & ChrW(&H962A&) & ChrW(&H5E02&) & ChrW(&H4E2D&) & ChrW(&H592E&) & ChrW(&H533A&) & ChrW(&H5927&) & ChrW(&H962A&) & ChrW(&H57CE&) & "1-1")
    sh.Cells(4, 1).value = (ChrW(&H5317&) & ChrW(&H6D77&) & ChrW(&H9053&) & ChrW(&H672D&) & ChrW(&H5E4C&) & ChrW(&H5E02&) & ChrW(&H5317&) & ChrW(&H533A&) & ChrW(&H5317&) & "6" & ChrW(&H6761&) & ChrW(&H897F&) & "4" & ChrW(&H4E01&) & ChrW(&H76EE&))
    sh.Cells(5, 1).value = (ChrW(&H5BAE&) & ChrW(&H57CE&) & ChrW(&H770C&) & ChrW(&H4ED9&) & ChrW(&H53F0&) & ChrW(&H5E02&) & ChrW(&H9752&) & ChrW(&H8449&) & ChrW(&H533A&) & ChrW(&H4E2D&) & ChrW(&H592E&) & "1-1-1")

    ' --- *Artificially build the post-abort state* the first 2 rows are already processed ---
    sh.Cells(2, 2).value = 11.111
    sh.Cells(2, 3).value = 22.222
    sh.Cells(2, 4).value = (ChrW(&H898B&) & ChrW(&H5F35&) & ChrW(&H308A&) & ChrW(&H756A&) & " 1")
    sh.Cells(2, 5).value = "ok"
    sh.Cells(3, 2).value = 33.333
    sh.Cells(3, 3).value = 44.444
    sh.Cells(3, 4).value = (ChrW(&H898B&) & ChrW(&H5F35&) & ChrW(&H308A&) & ChrW(&H756A&) & " 2")
    sh.Cells(3, 5).value = "ok(candidate 1)"

    Wv2Log.LogI "  Mark the first 2 of the 4 rows as 'done' and then call"
    n = Wv2Maps.MapsGeocodeSheet(sh)

    Wv2Log.LogI "        return value = " & n & " / aborted = " & Wv2Maps.MapsCanceled
    For i = 2 To 5
        Wv2Log.LogI "        row " & i & ": " & sh.Cells(i, 2).value & " / " & _
                    sh.Cells(i, 3).value & " / " & sh.Cells(i, 5).value
    Next i

    ' *The done rows must not be touched* (do the sentinel values remain as they are?)
    TestBool "*done rows are not rewritten* (row 2)", (sh.Cells(2, 2).value = 11.111)
    TestBool "  ok(candidate 1) also counts as done (row 3)", (sh.Cells(3, 2).value = 33.333)

    blanks = 0
    For i = 2 To 5
        If Len(CStr(sh.Cells(i, 5).value)) = 0 Then blanks = blanks + 1
    Next i
    TestBool "*the untouched rows were filled in*", (blanks = 0)
    TestBool "  coordinates went into row 4", (Len(CStr(sh.Cells(4, 2).value)) > 0)
    TestBool "  coordinates went into row 5", (Len(CStr(sh.Cells(5, 2).value)) > 0)
    TestBool "*the return value also counts the 'done' ones* (4 = 2 + 2)", (n = 4)
    TestBool "  not aborted", (Wv2Maps.MapsCanceled = False)
    TestBool "  EnableCancelKey is restored", _
             (Application.EnableCancelKey = xlInterrupt)

    wb.Close SaveChanges:=False

    Wv2Log.LogI ""
    TestCountPrint
    Wv2Log.LogI "================ Test_D7_Resume end ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_K4_Quiet (experiment of K-4a) - *quiet all tabs before closing*
'
'   Isolating the cause of slow release. The hypothesis is
'   *"the pages keep running so WebView2 is busy and COM responses are slow"*.
'   If all tabs are sent to about:blank and then closed, it should become faster.
'
'   *Not one line of product code is changed* What is done here is the same as what a user can do by hand
'   (navigate to another page). If it works, how to put it into the product side will be decided separately.
'
'   Usage:
'     1) Wv2Log.LogStart
'     2) UserForm1.Show vbModeless -> StartWebView2_Full
'     3) Wv2Maps.MapsOpen UserForm1.CurrentBrowser   (make one heavy page)
'     4) *Test_K4_Quiet 3* / *Test_K4_Quiet 1* / *Test_K4_Quiet 0*
'     5) Close the form with x
'
'   *K-4b: the wait seconds became an argument* With 3 seconds there is no gain in total time
'   (it takes 1.6 seconds to quiet + 3 seconds of waiting = 4.6 seconds, and release goes 6.7 -> 0.37 seconds).
'   *If it works with 0 seconds (just navigating away) we can win in total too*, so measure that.
'
'   Measured: do nothing = 6.718 s / wait 3 s = 0.367 s
' ============================================================
Public Sub Test_K4_Quiet(Optional ByVal waitSec As Single = 3)
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim i As Long
    Dim t0 As Single

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_K4_Quiet: Browser is not started."
        Exit Sub
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "================ Test_K4_Quiet start ================"
    Wv2Log.LogI "  *Send all " & b.TabCount & " tabs to about:blank*" & _
                " (then wait " & waitSec & " s)"

    t0 = Timer

    For i = 1 To b.TabCount
        Set p = b.PaneAt(i)
        If Not p Is Nothing Then
            p.View_Navigate "about:blank"
            Wv2Log.LogI "        tab " & i & " to about:blank"
        End If
    Next i

    Wv2Log.LogI "  *Finished sending (" & Format$(Timer - t0, "0.00") & " s)*"

    ' *Right after sending, things are "about to stop"*, so wait until they settle
    '   if waitSec = 0 do not wait (to measure whether just sending is enough)
    If waitSec > 0 Then
        t0 = Timer
        Do
            DoEvents
        Loop Until Timer - t0 >= waitSec Or Timer < t0
    End If

    Wv2Log.LogI "  *Quieted. In this state, please close the form with x*"
    Debug.Print ""
    Debug.Print "  *Quieted. Please close the form with x*"
    Wv2Log.LogI "================ Test_K4_Quiet end ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_K4_Step (K-4c) - *bisect what "slows down the closing"*
'
'   What is known from measurements:
'     AddTabWithUrl only ... closing is *fast*
'     MapsOpen equivalent     ... closing is *slow* (5 seconds)
'   The differences are firing EvalSync many times and building the element registry
'   with QuerySelector. Reproduce them stage by stage to check which one is responsible.
'
'   Usage: *No need to restart Excel at each stage. Just reopen the form*:
'     Wv2Log.LogStart
'     UserForm1.Show vbModeless -> StartWebView2_Full
'     *UserForm1.CurrentBrowser.QuietOnShutdown = False*  <- do not forget this
'     Test_K4_Step 1   (or 2 / 3 / 4)
'     close the form with x -> look at the release time in the log
'
'   *QuietOnShutdown = False is needed* K-4c made it quiet things before release, so
'   as it is everything becomes fast and nothing is isolated (design principle 103: check the probe).
'
'   Contents of each stage:
'     1 ... AddTabWithUrl only (known: fast)
'     2 ... + EvalSync 20 times
'     3 ... + QuerySelector 5 times
'     4 ... MapsOpen equivalent (everything)                        (known: slow)
'     5 ... *just run DoEvents for 15 seconds* (no JS fired)   (measured: fast)
'     6 ... *EvalSync 300 times*     (the count of MapsOpen)
'     7 ... *View_GetSource 300 times* (just reading via COM)   (measured: fast)
'     8 ... + wait until JS goes through           (stage 1 of MapsOpen)
'     9 ... + *wait for the search box with WaitFor* (the only element not yet tried)
'    10 ... + wait until /@ appears           (= MapsOpen equivalent)   (measured: fast)
'    11 ... + *EnableCancelKey operation + AddIgnoreNetwork* (the last difference)
'
'   *Proceed 8 -> 9 -> 10; the stage where it became slow is the answer*
'
'   *What was learned from measurements (isolation of K-4c)*
'     Step 1 + 30 seconds wait ... fast  <- but *VBA was stopped*
'     Step 2 (EvalSync x 20) ... fast
'     Step 3 (+QuerySelector x 5) ... fast / fast even with +30 seconds wait
'     MapsOpen ... *slow*
'     Step 5 (DoEvents 15 seconds / 690,000 times) ... fast
'   -> *"firing JS", "time" and "DoEvents" are all innocent*.
'     What remains is *the count* (MapsOpen fires EvalSync around 200 times). That is stages 6/7.
' ============================================================
Public Sub Test_K4_Step(Optional ByVal stepNo As Long = 1)
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim el As Wv2Element
    Dim i As Long
    Dim t0 As Single
    Dim loops As Long
    Dim savedKey As Long

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_K4_Step: Browser is not started."
        Exit Sub
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "================ Test_K4_Step " & stepNo & " start ================"
    If b.QuietOnShutdown Then
        Wv2Log.LogW "  *QuietOnShutdown is still True* It does not isolate anything."
        Wv2Log.LogW "  Type UserForm1.CurrentBrowser.QuietOnShutdown = False first."
        Debug.Print "  *Please type QuietOnShutdown = False first*"
    End If

    ' --- stage 1: open a tab (common to all stages) ---
    t0 = Timer
    Set p = b.AddTabWithUrl("https://www.google.com/maps")
    Wv2Log.LogI "  [1] AddTabWithUrl ... " & Format$(Timer - t0, "0.00") & " s"
    If p Is Nothing Then
        Wv2Log.LogI "      could not open a tab"
        Exit Sub
    End If

    If stepNo >= 2 And stepNo <= 4 Then
        ' --- stage 2: EvalSync 20 times ---
        t0 = Timer
        For i = 1 To 20
            p.EvalSync "1", 3
        Next i
        Wv2Log.LogI "  [2] EvalSync x 20 ... " & Format$(Timer - t0, "0.00") & " s"
    End If

    If stepNo >= 3 And stepNo <= 4 Then
        ' --- stage 3: QuerySelector 5 times (build the element registry) ---
        t0 = Timer
        For i = 1 To 5
            Set el = p.QuerySelector("input")
        Next i
        Wv2Log.LogI "  [3] QuerySelector x 5 ... " & Format$(Timer - t0, "0.00") & _
                    " s (registry " & p.ElementCount & " items)"
    End If

    If stepNo >= 11 Then
        ' [11] *wrap in the same order as MapsOpen*
        '   The real one is "arm before entering -> do everything -> AddIgnoreNetwork at the end -> disarm".
        '   *An arm that lasts only an instant is not a reproduction*, so apply it here and remove it at the end.
        savedKey = Application.EnableCancelKey
        Application.EnableCancelKey = xlDisabled
        Wv2Log.LogI "  [11] set EnableCancelKey = xlDisabled"
    End If

    If stepNo >= 8 Then
        ' --- stages 8/9/10 (K-4c): *reproduce the contents of MapsOpen stage by stage* ---
        '   After 6 straight losses trying to guess the candidates, we stopped guessing and
        '   adopted the *take MapsOpen apart, stop partway and close* approach.
        '   MapsOpen as a whole is definitely slow, so there must be a boundary somewhere.

        ' [8] wait until JS goes through (stage 1 of MapsOpenCore)
        t0 = Timer
        Do
            p.EvalSync "1", 3
            If p.LastEvalOk Then Exit Do
            If Timer - t0 >= 30 Or Timer < t0 Then Exit Do
            DoEvents
        Loop
        Wv2Log.LogI "  [8] until JS goes through ... " & Format$(Timer - t0, "0.00") & " s"

        If stepNo >= 9 Then
            ' [9] *WaitFor until the search box appears* (the only element not yet tried)
            t0 = Timer
            Set el = p.WaitFor("input[name='q']", 30)
            Wv2Log.LogI "  [9] WaitFor ... " & Format$(Timer - t0, "0.00") & " s " & _
                        IIf(el Is Nothing, "(not found)", "(found)") & _
                        " registry " & p.ElementCount & " items"
        End If

        If stepNo >= 10 Then
            ' [10] wait until the map has a position (= MapsOpen equivalent)
            t0 = Timer
            Do
                If InStr(1, p.View_GetSource(), "/@") > 0 Then Exit Do
                If Timer - t0 >= 8 Or Timer < t0 Then Exit Do
                DoEvents
            Loop
            Wv2Log.LogI "  [10] until /@ appears ... " & Format$(Timer - t0, "0.00") & " s"
        End If

        If stepNo >= 11 Then
            p.AddIgnoreNetwork "/search?tbm=map"
            Application.EnableCancelKey = savedKey
            Wv2Log.LogI "  [11] did AddIgnoreNetwork and restored EnableCancelKey"
        End If
    End If

    If stepNo = 6 Or stepNo = 7 Then
        ' --- stages 6/7 (K-4c): *increase the count by an order of magnitude* ---
        '   EvalSync of stage 2 was fast at 20 times, but MapsOpen polls for 17 seconds, so
        '   by calculation it fires *around 200 times*. The order of magnitude is different.
        '     6 ... EvalSync (fires JS) 300 times
        '     7 ... View_GetSource (just reads the URL via COM) 300 times
        '   *Which one works tells whether "the JS side piles up" or "the COM side piles up"*
        t0 = Timer
        For i = 1 To 300
            If stepNo = 6 Then
                p.EvalSync "1", 3
            Else
                p.View_GetSource
            End If
        Next i
        Wv2Log.LogI "  [" & stepNo & "] " & _
                    IIf(stepNo = 6, "EvalSync", "View_GetSource") & _
                    " x 300 ... " & Format$(Timer - t0, "0.00") & " s"
        Wv2Log.LogI "      registry " & p.ElementCount & " items / " & _
                    "the last EvalSync was " & p.LastEvalOk
    End If

    If stepNo = 5 Then
        ' --- stage 5 (K-4c): *just run DoEvents* no JS is fired even once ---
        '   Step 1 + 30 seconds wait was fast, but at that time VBA was *stopped*.
        '   MapsOpen keeps running DoEvents for 15 seconds = *WebView2 events keep being
        '   delivered to VBA callbacks*. This is the only difference.
        t0 = Timer
        loops = 0
        Do
            DoEvents
            loops = loops + 1
        Loop Until Timer - t0 >= 15 Or Timer < t0
        Wv2Log.LogI "  [5] ran DoEvents for 15 seconds ... " & loops & " loops"
    End If

    If stepNo = 4 Then
        ' --- stage 4: MapsOpen equivalent (the whole thing including waits) ---
        '   *It is = 4, not >= 4* Rolling it into stage 5 would not isolate anything
        '   (we actually blundered on this in K-4c. Design principle 103: check the probe itself)
        t0 = Timer
        Wv2Maps.MapsOpen b
        Wv2Log.LogI "  [4] MapsOpen ... " & Format$(Timer - t0, "0.00") & " s"
    End If

    Wv2Log.LogI "  *In this state, please close the form with x*"
    Debug.Print ""
    Debug.Print "  *Ran up to stage " & stepNo & ". Please close with x*"
    Wv2Log.LogI "================ Test_K4_Step " & stepNo & " end ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_K4_Help (procedure for K-4a)
' ============================================================
Public Sub Test_K4_Help()
    Debug.Print ""
    Debug.Print "=========================================================="
    Debug.Print " K-4a measurement procedure (slow release when the form is closed)"
    Debug.Print "=========================================================="
    Debug.Print ""
    Debug.Print "  *Start with Wv2Log.LogStart each time* One run is contained in one file."
    Debug.Print ""
    Debug.Print "  --- Case 1: bare state (measured 0.336 s) ---"
    Debug.Print "  1) Wv2Log.LogStart"
    Debug.Print "  2) UserForm1.Show vbModeless -> StartWebView2_Full"
    Debug.Print "  3) Close at once with x"
    Debug.Print ""
    Debug.Print "  --- Case 2: after opening a heavy page (measured *6.718 s*) ---"
    Debug.Print "  3) Wv2Maps.MapsOpen UserForm1.CurrentBrowser"
    Debug.Print "  4) Close with x"
    Debug.Print ""
    Debug.Print "  --- Case 3: quiet before closing (measured *0.367 s*) ---"
    Debug.Print "  3) Wv2Maps.MapsOpen UserForm1.CurrentBrowser"
    Debug.Print "  4) *Test_K4_Quiet 3*  (send to about:blank and wait 3 seconds)"
    Debug.Print "  5) Close with x"
    Debug.Print ""
    Debug.Print "  --- Case 3b: *cut the wait time* (to be measured) ---"
    Debug.Print "  4) Test_K4_Quiet 1   ... wait 1 second"
    Debug.Print "  4) Test_K4_Quiet 0   ... *do not wait (just send)*"
    Debug.Print "  *This is the one that wins on total time* It takes 1.6 seconds to quiet things down, so"
    Debug.Print "  waiting 3 seconds makes a total of 5 seconds, not much different from 'do nothing 6.7 seconds'."
    Debug.Print ""
    Debug.Print "  --- Case 4: after closing the heavy tab (to be measured) ---"
    Debug.Print "  3) Wv2Maps.MapsOpen UserForm1.CurrentBrowser"
    Debug.Print "  4) UserForm1.CurrentBrowser.CloseTab 4   ' the tab number of Maps"
    Debug.Print "  5) Close with x"
    Debug.Print ""
    Debug.Print "  *What to look at is the release part of the log* The elapsed time from Wv2NavBar.Shutdown start to"
    Debug.Print "  Wv2Browser.Terminate completion. ?Wv2Log.LogPath"
    Debug.Print ""
    Debug.Print "  --- K-4c: isolation after the countermeasure went in ---"
    Debug.Print "  *By default it navigates to about:blank before release, so it is fast*"
    Debug.Print "  To reproduce the slow state:"
    Debug.Print "    UserForm1.CurrentBrowser.QuietOnShutdown = False"
    Debug.Print "  Bisecting what makes it slow:"
    Debug.Print "    Test_K4_Step 1  ... AddTabWithUrl only (known: fast)"
    Debug.Print "    Test_K4_Step 2  ... + EvalSync x 20"
    Debug.Print "    Test_K4_Step 3  ... + QuerySelector x 5"
    Debug.Print "    Test_K4_Step 4  ... MapsOpen equivalent (known: slow)"
    Debug.Print ""
End Sub

' ============================================================
' Test_D7_Key (probe of D-7d) - *can GetAsyncKeyState see Esc?*
'
'   Esc could not be picked up in D-7c. The cause is one of two:
'     A. MapsPump is not running, so it is not called in the first place
'     B. Even when called, the API cannot see Esc
'   *This probe measures only B* It does not go through the product code's wait,
'   and runs on its own here to record the raw value (design principle 103).
'
'   *EnableCancelKey is set to xlDisabled during execution* Otherwise VBA breaks on Esc
'   and nothing is measured. It always ends in 10 seconds and is always restored.
'
'   Needs neither the network nor the browser.
' ============================================================
Public Sub Test_D7_Key()
    Dim savedKey As Long
    Dim t0 As Single
    Dim polls As Long
    Dim hits As Long
    Dim firstHit As Single
    Dim lastSec As Long
    Dim k As Long
    Dim seen As Long

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D7_Key start ================"

    ' D-7e: *"10 seconds from now" finishes before it can be read* so it was changed to wait until pressed.
    Debug.Print ""
    Debug.Print "  *** Please press Esc now ***"
    Debug.Print "  (It ends immediately once pressed. It waits at most 30 seconds)"
    Wv2Log.LogI "  *wait until pressed* (at most 30 seconds)"

    savedKey = Application.EnableCancelKey
    Application.EnableCancelKey = xlDisabled   ' otherwise it breaks on Esc

    firstHit = -1
    lastSec = -1
    t0 = Timer
    Do
        DoEvents
        polls = polls + 1

        k = Wv2Maps.MapsRawEscState()
        If k <> 0 Then
            hits = hits + 1
            If firstHit < 0 Then
                firstHit = Timer - t0
                seen = k
                Debug.Print "  *detected* raw value = " & k
            End If
    ' it is enough to know it was pressed (it runs 40,000 times per second so it accumulates quickly)
            If hits >= 500 Then Exit Do
        End If

        If Int(Timer - t0) <> lastSec Then
            lastSec = Int(Timer - t0)
    Debug.Print "  ... " & (30 - lastSec) & " seconds left"
        End If

        If Timer - t0 >= 30 Or Timer < t0 Then Exit Do
    Loop

    Application.EnableCancelKey = savedKey

    Wv2Log.LogI ""
    Wv2Log.LogI "        polling " & polls & " times / " & _
                Format$(Timer - t0, "0.0") & " s"
    Wv2Log.LogI "        detections " & hits & " times"
    If hits > 0 Then
        Wv2Log.LogI "        first detection " & Format$(firstHit, "0.00") & " s later"
        Wv2Log.LogI "        raw value  " & seen & " (&H" & Hex$(seen And &HFFFF&) & ")"
    End If

    TestBool "*GetAsyncKeyState can see Esc*", (hits > 0)

    Wv2Log.LogI ""
    If hits > 0 Then
        Wv2Log.LogI "  *Conclusion: the API can see it*"
    Else
        Wv2Log.LogI "  *Conclusion: it could not be detected even after waiting 30 seconds*"
        Wv2Log.LogI "         *Unless you forgot to press*, the way keys are picked up must be changed."
    End If

    TestCountPrint
    Wv2Log.LogI "================ Test_D7_Key end ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' TestWaitEscReleased (D-7e, Private) - *wait until the finger comes off Esc*
'
'   In D-7d, when the 2nd run was started right after the abort, the Esc that was being mashed
'   took effect on it too and 3 verdicts FAILed. The product side was correct and
'   *the sequencing of the verification was bad*. Wait until it is quiet, then go on.
'
'   *Set xlDisabled while waiting* If left restored, an Esc pressed
'   during the wait makes VBA break.
' ============================================================
Private Sub TestWaitEscReleased(ByVal maxSec As Single)
    Dim savedKey As Long
    Dim t0 As Single
    Dim quiet As Single

    savedKey = Application.EnableCancelKey
    Application.EnableCancelKey = xlDisabled

    quiet = -1
    t0 = Timer
    Do
        DoEvents
        If Wv2Maps.MapsRawEscState() = 0 Then
            If quiet < 0 Then quiet = Timer
            If Timer - quiet >= 2 Then Exit Do   ' *if quiet for 2 seconds, consider the finger lifted*
        Else
            quiet = -1
        End If
        If Timer - t0 >= maxSec Or Timer < t0 Then Exit Do
    Loop

    Application.EnableCancelKey = savedKey
    Wv2Log.LogI "  (Esc has gone quiet. Resuming)"
End Sub

' ============================================================
' Test_D7_StatusBar (probe of D-7b)
'
'   *Measures only Excel's raw behavior without going through product code* (design principle 103)
'   In the D-7 real-machine run, "*the status bar is restored*" FAILed in all 3 places.
'   Separates whether the way the verdict is written (looking with TypeName) is bad, or Excel does not return control
'   during execution. *No verdict is made. Only facts are output.*
'
'   Needs neither the network nor the browser.
' ============================================================
Public Sub Test_D7_StatusBar()
    Wv2Log.LogI ""
    Wv2Log.LogI "================ Test_D7_StatusBar start ================"
    Wv2Log.LogI "  *Measured in D-7b: putting False left the string ""FALSE""*"
    Wv2Log.LogI "  Measure the correct way to restore, candidate by candidate. *If the type returns to Boolean, it is the right one*"
    Wv2Log.LogI ""

    TestDumpStatusBar "(0) when nothing has been done"
    Wv2Log.LogI ""

    TestStatusTry "(1) False       ", False
    TestStatusTry "(2) Empty       ", Empty
    TestStatusTry "(3) empty string """"   ", ""
    TestStatusTry "(4) vbNullString", vbNullString
    TestStatusTry "(5) CVar(False) ", CVar(False)

    Wv2Log.LogI ""
    Wv2Log.LogI "  *Whichever gives type=Boolean is the correct way to restore*"
    Wv2Log.LogI "  If all are String, this Excel cannot return control -- in that case"
    Wv2Log.LogI "  *leaving it as an empty string* does less real harm (better than displaying FALSE)."

    ' *Restore with Empty when the measurement is done* (this is what turned out right in the D-7d measurement)
    Application.StatusBar = Empty
    TestDumpStatusBar "(9) Empty for cleanup"

    Wv2Log.LogI "================ Test_D7_StatusBar end ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' TestStatusTry (D-7c, Private) - try one candidate way of restoring
'
'   *Each time, "put a string in and then restore"* Restoring without putting one in leaves Excel
'   in the state where it already holds control, so nothing has been measured (design principle 103).
'   Also look at the value after DoEvents -- because Excel may take it back later.
' ============================================================
Private Sub TestStatusTry(ByVal tag As String, ByVal v As Variant)
    Dim before As String
    Dim after As String

    Application.StatusBar = "Testing 12/100"
    Application.StatusBar = v
    before = "[" & CStr(Application.StatusBar) & "] type=" & _
             TypeName(Application.StatusBar)

    DoEvents
    after = "[" & CStr(Application.StatusBar) & "] type=" & _
            TypeName(Application.StatusBar)

    Wv2Log.LogI "        " & tag & " -> immediately after " & before & " / after DoEvents " & after
End Sub

' ============================================================
' TestDumpStatusBar (D-7b, Private) - output the measured value of the status bar
'   *Keep both the value itself and the type* Neither one alone can isolate the cause.
' ============================================================
Private Sub TestDumpStatusBar(ByVal tag As String)
    Wv2Log.LogI "        " & tag & " = [" & CStr(Application.StatusBar) & _
                "] type=" & TypeName(Application.StatusBar)
End Sub

' ============================================================
' Test_D7_Help (procedure for D-7)
' ============================================================
Public Sub Test_D7_Help()
    Debug.Print ""
    Debug.Print "=========================================================="
    Debug.Print " D-7 verification procedure (progress display and abort)"
    Debug.Print "=========================================================="
    Debug.Print ""
    Debug.Print "  --- Preparation ---"
    Debug.Print "  1) Wv2Log.LogStart"
    Debug.Print "  2) Run UserForm1.Show vbModeless, then run StartWebView2_Full"
    Debug.Print "  3) *Wait until the event burst settles* (spec fact 20)"
    Debug.Print ""
    Debug.Print "  --- Run (from the top in order) ---"
    Debug.Print "  4) Test_D7_Key ... *probe* whether GetAsyncKeyState can see"
    Debug.Print "     Esc. *Waits until pressed* (at most 30 seconds)."
    Debug.Print "     No network needed. The remaining seconds appear in the Immediate window."
    Debug.Print "  4b) Test_D7_StatusBar ... *probe* measures how to restore the status bar"
    Debug.Print "     candidate by candidate. No network needed, a few seconds."
    Debug.Print "     *No verdict. Only facts are output* (design principle 103)."
    Debug.Print "  5) Test_D7_Cancel ... *no network needed* the abort entry point and the denominator (finishes quickly)"
    Debug.Print "  6) Test_D7_Resume ... *automatic* resuming from the post-abort state (2 items, 15 seconds)"
    Debug.Print "  7) Test_D7_Sheet  ... *real site* processes 4 items."
    Debug.Print "     *Press Esc around the 2nd one* while watching the status bar."
    Debug.Print "     *No need to mash it. Once works* (mashing also takes effect on the resume)."
    Debug.Print "     There is no need to hold it down, and the focus may be anywhere."
    Debug.Print "     When it stops, *** It stopped *** appears in the Immediate window."
    Debug.Print "     *From D-7c on, neither the abort dialog nor runtime error 18 appears*"
    Debug.Print "     If they appear, it is not working. The log shows *an abort was requested with Esc*."
    Debug.Print "     After that it calls again automatically and we see whether it fills in from where it left off."
    Debug.Print ""
    Debug.Print "  --- Read the verdicts in the log file (design principle 105) ---"
    Debug.Print "  ?Wv2Log.LogPath"
    Debug.Print ""
    Debug.Print "  --- How to use it in practice ---"
    Debug.Print "  Wv2Maps.MapsGeocodeSheet Sheets(""AddressBook"")"
    Debug.Print "    During processing the status bar shows 12/100 (ok 11 / failed 1)."
    Debug.Print "    *It can be stopped any time with Esc* Nothing is written to the row where it stopped."
    Debug.Print "    Calling it again skips the ok rows and goes on from where it left off."
    Debug.Print "  ?Wv2Maps.MapsCanceled   ' did it end by an abort?"
    Debug.Print "  ?Wv2Maps.MapsCountRows(Sheets(""AddressBook""))  ' count first how many there are"
    Debug.Print "  Wv2Maps.MapsCancel = True  ' stop from outside (from the Immediate window)"
    Debug.Print ""
End Sub

' ============================================================
' Test_D4_Help (procedure for D-4)
' ============================================================
Public Sub Test_D4_Help()
    Debug.Print ""
    Debug.Print "=========================================================="
    Debug.Print " D-4 verification procedure (probe and quiescence wait)"
    Debug.Print "=========================================================="
    Debug.Print ""
    Debug.Print "  --- Preparation ---"
    Debug.Print "  1) Wv2Log.LogStart"
    Debug.Print "  2) Run UserForm1.Show vbModeless, then run StartWebView2_Full"
    Debug.Print "  3) *Wait until the event burst settles* (spec fact 20)"
    Debug.Print ""
    Debug.Print "  --- Run ---"
    Debug.Print "  4) Test_D4_Probe  ... *done* probe and health check (28 OK)"
    Debug.Print "  5) Test_D4_Settle ... *done* quiescence wait (22 OK)"
    Debug.Print "  6) Test_D4_Signal ... *done* explicit signal (31 OK)"
    Debug.Print "  7) Test_D4_Log    ... *verification of D-4e* diagnostic log / in-flight / URL signal"
    Debug.Print "  8) Test_D4_Site   ... *reconnaissance of D-4d* Google Maps (about 70 seconds)"
    Debug.Print "  9) Test_D5_Geocode ... *D-5* address -> coordinates, 3 in succession (about 60 seconds)"
    Debug.Print "     A demonstration of Wv2Maps.MapsOpen / MapsGeocode. Exactly the form used in business work."
    Debug.Print " 10) Test_D5_Sheet   ... *D-5b* sheet integration (try it by creating a new workbook)"
    Debug.Print " 11) Test_D6_All     ... *D-6* QuerySelectorAll and lifetime management"
    Debug.Print " 12) Test_D6_Pick    ... *D-6b* take the first from the candidate list (real site)"
    Debug.Print " 13) Test_D7_Cancel  ... *D-7* the abort entry point and the denominator (no network needed)"
    Debug.Print " 14) Test_D7_Resume  ... *D-7b* resuming from the post-abort state (automatic)"
    Debug.Print " 15) Test_D7_Sheet   ... *D-7b* progress and abort (press Esc during execution)"
    Debug.Print "     *The only Test_* that accesses an external site for real* Needs a network."
    Debug.Print "     To change the address, type Test_D4_Site ""another address""."
    Debug.Print ""
    Debug.Print "  --- Tools for investigating real sites ---"
    Debug.Print "  Test_D4_Dom          ... *lists what is on the currently open tab*"
    Debug.Print "  Test_D4_Dom ""h1,h2""  ... lists by the specified selector"
    Debug.Print ""
    Debug.Print "  --- Diagnostic log (D-4e) - *decide what to wait for by observation* ---"
    Debug.Print "  p.SpaProbeLogging = True   ' start accumulating (default False)"
    Debug.Print "  ... operate ..."
    Debug.Print "  ?p.SpaProbeDrainLog        ' fetch and flush to the log (60 items at a time)"
    Debug.Print "  p.SpaProbeLogging = False"
    Debug.Print ""
    Debug.Print "  --- URL signal (D-4e) ---"
    Debug.Print "  p.ArmUrlSignal ""/place/""   ' wait for the URL to change (supports pushState)"
    Debug.Print "    *For address search on Maps this was the most reliable marker*"
    Debug.Print "    Selectors of real sites change without notice. Look before writing by guesswork."
    Debug.Print ""
    Debug.Print "  *Verdicts remain in the log file* Look at the trailing *Verdicts n*"
    Debug.Print "    ?Wv2Log.LogPath tells you the location of the file"
    Debug.Print ""
    Debug.Print "  --- What to look at ---"
    Debug.Print "  - (1) undefined before installation = the probe is not attached to a Pane that is not called"
    Debug.Print "  - (2) *round-trip time before and after installation* Look at the numbers themselves in the log"
    Debug.Print "      Spec fact 52 is 38ms. Check that setting up the observation does not make it drop drastically"
    Debug.Print "  - (5) fetch and XHR are counted and return to 0 on completion"
    Debug.Print "  - *(6) is the heart of D-4a* Even if the page replaces window.fetch,"
    Debug.Print "      it repairs itself automatically at the next call (rp>=1) and is still counted afterwards"
    Debug.Print "  - (8) the probe disappears on navigation and can be set up again"
    Debug.Print ""
    Debug.Print "  --- What to look at (Test_D4_Settle) *D-4b* ---"
    Debug.Print "  - *(2) is the heart* At the point the wait ends, #target is the 'chain result'"
    Debug.Print "      = evidence that it waits not only for the communication to finish but also for the DOM update after it"
    Debug.Print "      (it spans fetch -> rewrite 300ms later)"
    Debug.Print "  - (3) on a page that does not settle, False + LastEvalOk=True (not a failure)"
    Debug.Print "  - (4) *with an exclusion the same page becomes quiet* = the noise exclusion works"
    Debug.Print "  - (6) with auto wait OFF the pre-update value is seen, with ON the post-update value"
    Debug.Print "  - A wait whose slack in LastSettleInfo is small (under 100ms) is *borderline*"
    Debug.Print "      It may just have slipped in by chance. Raise stableMs or"
    Debug.Print "      use an explicit signal together"
    Debug.Print ""
    Debug.Print "  --- What to look at (Test_D4_Signal) *D-4c* ---"
    Debug.Print "  - *(3) is the heart* If nothing happens after arming, it becomes False even if the page is quiet"
    Debug.Print "      = *being ignored by the app can be detected*"
    Debug.Print "      With the quiescence wait alone this would succeed"
    Debug.Print "  - (1) a nonexistent element / invalid selector fails to arm *on the spot* (fail-fast)"
    Debug.Print "  - (4) arm is one-shot. It is not carried over to the next wait"
    Debug.Print "  - (6) if the arm vanishes, it fails with signal-lost (it does not quietly re-arm)"
    Debug.Print ""
    Debug.Print "  --- What to look at (Test_D4_Site) *D-4d aims at observation* ---"
    Debug.Print "  - The number of verdicts is only a few structural points. *Read the body of the log for the content*"
    Debug.Print "  - The per-second lines of (3)(6) [01s] q=.. m=.. n=.. are the stars"
    Debug.Print "      q grows      -> quiescence comes (the quiescence wait is usable)"
    Debug.Print "      q does not grow -> *an exclusion or an explicit signal is needed*"
    Debug.Print "  - The round-trip cost of (2). On a light page it was 35.4 -> 36.3 ms"
    Debug.Print "  - If rp (rebuild count) has increased, Maps is replacing fetch"
    Debug.Print "  - Whether the url of (7) contains /@latitude,longitude = whether coordinates can be obtained"
    Debug.Print ""
    Debug.Print "  --- If you want to try by hand ---"
    Debug.Print "  Set p = UserForm1.GetActivePane"
    Debug.Print "  ?p.SpaProbeState"
    Debug.Print "  ?p.SpaProbeHealthy"
    Debug.Print "  ?p.SpaProbeReset"
    Debug.Print "  ?p.WaitSettled(10)     : ?p.LastSettleInfo"
    Debug.Print "  p.IgnoreSelectors = ""#clock, .ticker"""
    Debug.Print "  p.AddIgnoreNetwork ""google-analytics"""
    Debug.Print "  p.AutoWaitAfterAction = True   ' wait automatically after an operation"
    Debug.Print ""
    Debug.Print "  --- Explicit signal (D-4c) - *arm, then act* ---"
    Debug.Print "  p.ArmContentSignal ""#result-table""   ' wait for an existing container to be rewritten"
    Debug.Print "  p.ArmNetworkSignal ""/api/search""     ' wait for that request to complete"
    Debug.Print "  el.Click"
    Debug.Print "  ?p.WaitSettled(10) : ?p.LastSettleInfo"
    Debug.Print "  ?p.DisarmSignals   ' when withdrawing"
    Debug.Print ""
    Debug.Print "  ** Even if WaitSettled is True, it is not proof that ""the app's processing has finished"" **"
    Debug.Print "    It is merely quiet. It cannot be told whether it is quiet because it was ignored or because it finished."
    Debug.Print "    For important operations, use D-4c's explicit signal (arm) together."
    Debug.Print ""
    Debug.Print "  *The quiescence window is measured from the moment WaitSettled is called* Even if it was quiet before the call,"
    Debug.Print "    it is not satisfied at once (it watches for stableMs at the very least). However,"
    Debug.Print "    processing that starts after stableMs is in principle missed."
    Debug.Print ""
End Sub



' ============================================================
' *N-1: verification of network request capture*
'
'   Test_N1_Capture ... *a regression test that is self-contained with its own HTML* (issue 8)
'                     First check that capture works reliably without relying on real sites.
'   Test_N1_Watch   ... *start capturing the traffic of the currently open tab* (practical entry point)
'   Test_N1_Drain   ... flush the accumulated entries to the log and empty the buffer
'   Test_N1_Stop    ... stop capturing
'   Test_N1_Help    ... procedure
'
'   *Verdicts go to Wv2Log* (design principle 105). The Immediate window is flushed away by the plumbing log.
' ============================================================


' ============================================================
' Test_N1_Capture (regression test of N-1b)
'
'   *What the first N-1 real-machine run revealed*
'     Requests served through a virtual host (SetVirtualHostNameToFolderMapping)
'     do not show up in WebResourceRequested. That was the cause of the first run's total failure, and
'     *the wiring was correct from the start* (on a real site all 10 were captured).
'
'   *So the targets were made 2 routes + 1 control route (issue 1, plan G)*
'
'     Plan F  http://127.0.0.1:59999/...   unreachable local address
'          *Does not depend on external services* The connection always fails, but
'          WebResourceRequested is an event for "stopping / rewriting a request", so
'          it fires before going out to the network -- supposedly. We check that here.
'          - *Avoid the ports Chromium blocks* (1/7/9/11/13/.../10080 etc.).
'            59999 is not blocked.
'          - *Not stopped as mixed content despite being http* 127.0.0.1 and localhost are
'            by spec "potentially trustworthy origins", so even when called from an https page
'            they are not treated as mixed content.
'
'     Plan D  https://httpbingo.org/...    external service
'          A route that arrives cleanly, in exchange for needing the net.
'
'     Control https://appassets.netprobe/data.json (virtual host)
'          *Count that it is NOT captured to confirm* = turn the spec fact we found into a regression test
'
'   The page itself is served on the virtual host as before. *Serving works without problems*.
'   Only the event does not show up.
'
'   *Whichever target does not arrive, the verdicts that came up empty are stated explicitly in the log*
'   (so as not to pile up OKs in a situation where "not arriving" is expected)
'
'   *Do not assume the type (N-1c)*
'     Even when fired with fetch(), WebView2 reports it with the XHR type. The verdict is made by "was it captured",
'     and *what type it actually arrived as is written to the log and counted*.
' ============================================================
Public Sub Test_N1_Capture()
    Dim b   As Wv2Browser
    Dim p   As Wv2Pane
    Dim el  As Wv2Element
    Dim folderPath   As String
    Dim hr           As Long
    Dim localOk      As Boolean
    Dim netOk        As Boolean
    Dim totalBefore  As Long

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_N1_Capture: Browser is not started." & _
                    "Run UserForm1.StartWebView2_Full first."
        Exit Sub
    End If

    folderPath = N1WriteFolder()
    If LenB(folderPath) = 0 Then
        Wv2Log.LogI "Test_N1_Capture: failed to write out the verification page. Aborting."
        Exit Sub
    End If

    Set p = b.AddTab()
    If p Is Nothing Then
        Wv2Log.LogI "Test_N1_Capture: failed to create the tab."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_N1_Capture start ================"
    Wv2Log.LogI "        plan F target = " & N1_LOCAL & "  (unreachable local)"
    Wv2Log.LogI "        plan D target = " & N1_NET & "  (external service)"
    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (0) Preparation (*start capturing before Navigate*) ---"

    TestBool "NetCaptureStart succeeds", p.NetCaptureStart()
    TestBool "  capturing is on", (p.NetCaptureOn = True)

    hr = p.View3_SetVirtualHostNameToFolderMapping(N1_HOST, folderPath, 1)   ' 1 = ALLOW
    TestBool "  the virtual host mapping can be made", (hr = 0)

    hr = p.View_Navigate("https://" & N1_HOST & "/netprobe.html")
    TestBool "  Navigate succeeds", (hr = 0)

    If Not D2WaitTitle(p, ("N-1 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&)), 10) Then
        Wv2Log.LogI "Test_N1_Capture: could not confirm that the verification page loaded."
        p.NetLogDrain
        p.NetCaptureStop
        TestCountPrint
        Exit Sub
    End If

    ' Leave enough time for the resources (CSS / images) in the page to finish going out
    D3Pump 3

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (1) *Requests to the virtual host do not show up in the event (regression of the spec fact)* ---"
    ' *Do not use indirect indicators* (design principle 112) The page can certainly be read
    ' (the title was obtained), yet we count that its DOCUMENT is not there even once.
    TestBool "*the page's own DOCUMENT is not there even once*", _
             (N1Find(p, "", "", "netprobe.html") = 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) fetch / XHR - 2 target routes + control ---"

    N1Fire p, "(function(){" & _
              "fetch('" & N1_LOCAL & "/n1?k=f-local').catch(function(){});" & _
              "fetch('" & N1_NET & "/get?k=f-net').catch(function(){});" & _
              "fetch('data.json?k=f-vhost').catch(function(){});" & _
              "var a=new XMLHttpRequest();a.open('GET','" & N1_LOCAL & _
              "/n1?k=x-local',true);a.send();" & _
              "var c=new XMLHttpRequest();c.open('GET','" & N1_NET & _
              "/get?k=x-net',true);c.send();" & _
              "return 1;})()"

    ' *Do not put the type in the condition (N-1c)* A fetch does not necessarily arrive as FETCH.
    localOk = N1Wait(p, "GET", "", "k=f-local", 6)
    TestBool "*plan F: a fetch to the unreachable local is captured*", localOk
    TestBool "  plan F: XHR is captured too", N1Wait(p, "GET", "", "k=x-local", 4)

    netOk = N1Wait(p, "GET", "", "k=f-net", 8)
    TestBool "*plan D: a fetch to the external host is captured*", netOk
    TestBool "  plan D: XHR is captured too", N1Wait(p, "GET", "", "k=x-net", 4)

    TestBool "*control: a fetch to the virtual host is not captured*", _
             (N1Find(p, "", "", "k=f-vhost") = 0)

    ' --- *Count what type they arrived as (N-1c)* ---
    '   Whether what was fired with fetch arrives as FETCH or XHR is *not decided
    '   by our convenience*. The type the real machine returned is written out as is and judged.
    Wv2Log.LogI "        arrived types: f-local=" & N1CtxOf(p, "k=f-local") & _
                "  x-local=" & N1CtxOf(p, "k=x-local") & _
                "  f-net=" & N1CtxOf(p, "k=f-net") & _
                "  x-net=" & N1CtxOf(p, "k=x-net")
    TestBool "*fetch arrives as the XHR type, not FETCH*", _
             (N1CtxOf(p, "k=f-local") = "XHR" Or N1CtxOf(p, "k=f-net") = "XHR")

    If Not localOk Then
        Wv2Log.LogW "  NOTE: plan F did not arrive even once. Read the subsequent plan F verdicts as having come up empty."
    End If
    If Not netOk Then
        Wv2Log.LogW "  NOTE: plan D did not arrive even once (no net or httpbingo.org is down)." & _
                    "Read the subsequent plan D verdicts as having come up empty."
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) *Images and CSS do not arrive (default filter)* ---"
    ' The page requests n1.css / n1.png / image/png *for both routes*.
    ' We count that the default filter (DOCUMENT / XHR / FETCH) rejects them.
    TestBool "*n1.css is not there even once*", (N1Find(p, "", "", "n1.css") = 0)
    TestBool "*n1.png is not there even once*", (N1Find(p, "", "", "n1.png") = 0)
    TestBool "*image/png is not there even once*", (N1Find(p, "", "", "image/png") = 0)
    If Not (localOk Or netOk) Then
        Wv2Log.LogW "  NOTE: *these 3 came up empty* No target arrived at all, so " & _
                    """not arriving"" is not thanks to the filter."
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3b) *Output everything captured so far* ---"
    ' *So as not to fill in by guesswork (N-1c)*
    '   In the next (4), dropping the capacity to 3 erases the contents. In N-1b that lost
    '   the record of (2), and we ended up guessing "seq 1/2 should have been f-local and f-net".
    '   Always output it before it disappears (design principle 112).
    p.NetLogDrain

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) *On overflow it is not discarded silently; the number discarded is visible* ---"

    p.NetLogCapacity = 3
    TestBool "the capacity can be set to 3", (p.NetLogCapacity = 3)
    TestBool "  both the contents and the total return to 0", (p.NetLogCount = 0 And p.NetLogTotal = 0)

    N1Fire p, "(function(){for(var i=0;i<5;i++){" & _
              "fetch('" & N1_LOCAL & "/n1?n='+i).catch(function(){});" & _
              "fetch('" & N1_NET & "/get?n='+i).catch(function(){});" & _
              "}return 10;})()"
    D3Pump 4

    Wv2Log.LogI "        total fired " & p.NetLogTotal & " / on hand " & p.NetLogCount & _
                " / overflowed " & p.NetLogDropped & " items"
    TestBool "5 or more fired", (p.NetLogTotal >= 5)
    TestBool "*on hand is only 3, the capacity*", (p.NetLogCount = 3)
    TestBool "*the overflow count is counted*", (p.NetLogDropped = p.NetLogTotal - 3)

    ' --- *N-1b: the total remains even after draining (issue 3 plan Y)* ---
    '   In the log of the first real-machine run, "total fired 0" appeared and the fact that 10 had arrived disappeared.
    '   A verdict to stop that from recurring.
    totalBefore = p.NetLogTotal
    p.NetLogDrain
    TestBool "*the total fired remains even after draining (cumulative)*", (p.NetLogTotal = totalBefore)
    TestBool "  only the on-hand part becomes empty", (p.NetLogCount = 0)
    p.NetLogClear
    TestBool "  with NetLogClear the total also returns to 0", (p.NetLogTotal = 0)

    p.NetLogCapacity = 500

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) *When stopped it really stops arriving, and can be re-cast* ---"

    TestBool "NetCaptureStop succeeds", p.NetCaptureStop()
    TestBool "  capturing is off", (p.NetCaptureOn = False)

    p.NetLogClear
    N1Fire p, "(function(){" & _
              "fetch('" & N1_LOCAL & "/n1?k=afterstop').catch(function(){});" & _
              "fetch('" & N1_NET & "/get?k=afterstop').catch(function(){});return 1;})()"
    D3Pump 3
    ' *The filter is also removed*, so it is correct that nothing arrives after stopping.
    TestBool "*nothing arrives after stopping*", (p.NetLogTotal = 0)

    TestBool "NetCaptureStart can be done again", p.NetCaptureStart()
    N1Fire p, "(function(){" & _
              "fetch('" & N1_LOCAL & "/n1?k=restart').catch(function(){});" & _
              "fetch('" & N1_NET & "/get?k=restart').catch(function(){});return 1;})()"
    TestBool "*the re-cast net captures again*", N1Wait(p, "GET", "", "k=restart", 8)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) *Page navigation by form POST is visible* ---"
    ' *This is the very reason method B (native events) was chosen*
    '   With wrapping of in-page JS, this one line could not be seen by any means.

    p.NetLogClear
    Set el = p.QuerySelector("#postbtn-local")
    TestBool "plan F: the submit button can be grabbed", Not (el Is Nothing)
    If Not (el Is Nothing) Then
        TestBool "  plan F: it can be clicked", el.Click()
        TestBool "*plan F: the POST is captured as a DOCUMENT*", _
                 N1Wait(p, "POST", "DOCUMENT", "127.0.0.1", 8)
    End If

    ' Return to the verification page (it is a virtual host, so this navigation itself does not show up in the event)
    p.NetLogClear
    If N1Nav(p, "https://" & N1_HOST & "/netprobe.html", ("N-1 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&)), 10) Then
        Set el = p.QuerySelector("#postbtn-net")
        TestBool "plan D: the submit button can be grabbed", Not (el Is Nothing)
        If Not (el Is Nothing) Then
            TestBool "  plan D: it can be clicked", el.Click()
            TestBool "*plan D: the POST is captured as a DOCUMENT*", _
                     N1Wait(p, "POST", "DOCUMENT", "httpbingo.org", 10)
        End If
    Else
        Wv2Log.LogI "        could not return to the verification page, so skipping the POST of plan D"
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (7) *GET of DOCUMENT* (issue 2 plan a) ---"
    ' The page's own DOCUMENT is on the virtual host so it cannot be captured. Navigate directly to the target to capture it.

    p.NetLogClear
    p.View_Navigate N1_LOCAL & "/n1-doc-get"
    TestBool "*plan F: DOCUMENT GET is captured*", _
             N1Wait(p, "GET", "DOCUMENT", "n1-doc-get", 8)

    p.NetLogClear
    p.View_Navigate N1_NET & "/get?k=doc"
    TestBool "*plan D: DOCUMENT GET is captured*", _
             N1Wait(p, "GET", "DOCUMENT", "k=doc", 10)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- Output everything remaining on hand ---"
    p.NetLogDrain
    p.NetCaptureStop

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback depth (expected 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_N1_Capture end ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' Test_N1_Site (the isolating role for N-1 and a feel for a real site)
'
'   *Accesses an external site for real* Because Test_N1_Capture depends on the virtual host,
'   this isolates in one move "is the wiring bad, or is it only that requests to the virtual host do not show up".
'
'   What it does is the same as (1) of Test_N1_Capture:
'     create a tab -> *start capturing before Navigate* -> open -> output what has accumulated.
'
'   It is a real site, so the count changes every time. *The verdict is only "is there at least one DOCUMENT"*
'   (the number of contents is not counted).
' ============================================================
Public Sub Test_N1_Site(Optional ByVal url As String = "https://www.google.com/")
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim hr As Long

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_N1_Site: Browser is not started."
        Exit Sub
    End If

    Set p = b.AddTab()
    If p Is Nothing Then
        Wv2Log.LogI "Test_N1_Site: failed to create the tab."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_N1_Site start (" & url & ") ================"

    TestBool "NetCaptureStart succeeds", p.NetCaptureStart()

    hr = p.View_Navigate(url)
    TestBool "  Navigate succeeds", (hr = 0)

    ' Real sites are slow, so wait a little longer (the count does not matter)
    D3Pump 6

    TestBool "*at least one DOCUMENT is captured*", _
             (N1Find(p, "", "DOCUMENT", "") > 0)
    TestBool "  something is captured", (p.NetLogCount > 0)

    Wv2Log.LogI ""
    p.NetLogDrain
    p.NetCaptureStop

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback depth (expected 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_N1_Site end ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' Test_N1_Watch (practical entry point of N-1)
'
'   *Start capturing the traffic of the currently open tab*
'   A tool for seeing "operate by hand once, and what is being hit in what order".
'
'   If contexts is omitted, the 3 types DOCUMENT / XHR / FETCH.
'   If you want to see images and CSS too, type Test_N1_Watch "ALL".
' ============================================================
Public Sub Test_N1_Watch(Optional ByVal contexts As String = "")
    Dim b As Wv2Browser
    Dim p As Wv2Pane

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Debug.Print "[N-1] Browser is not started."
        Exit Sub
    End If
    Set p = b.ActivePane
    If p Is Nothing Then
        Debug.Print "[N-1] There is no active tab."
        Exit Sub
    End If

    If p.NetCaptureStart(contexts) Then
        Debug.Print "[N-1] Capture started. Operate the page by hand and then type Test_N1_Drain."
        Debug.Print "      *Verdicts and the list go to the log file* (%APPDATA%\Wv2Browser\logs)"
    Else
        Debug.Print "[N-1] Could not start capture. Look at the log."
    End If
End Sub


' ============================================================
' Test_N1_Drain (N-1) - flush the accumulated entries to the log and empty the buffer
' ============================================================
Public Sub Test_N1_Drain()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim n As Long

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then Exit Sub
    Set p = b.ActivePane
    If p Is Nothing Then Exit Sub

    n = p.NetLogDrain
    Debug.Print "[N-1] Flushed " & n & " items to the log (total fired " & p.NetLogTotal & ")."
End Sub


' ============================================================
' Test_N1_Stop (N-1) - stop capturing (also removes the filter)
' ============================================================
Public Sub Test_N1_Stop()
    Dim b As Wv2Browser
    Dim p As Wv2Pane

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then Exit Sub
    Set p = b.ActivePane
    If p Is Nothing Then Exit Sub

    p.NetLogDrain
    If p.NetCaptureStop() Then
        Debug.Print "[N-1] Capture ended."
    Else
        Debug.Print "[N-1] Capture ended, but there were failures in the cleanup. Look at the log."
    End If
End Sub


' ============================================================
' Test_N1_Help (procedure for N-1)
' ============================================================
Public Sub Test_N1_Help()
    Debug.Print "==== N-1 real-machine procedure (seeing what the browser is hitting) ===="
    Debug.Print ""
    Debug.Print "  [Regression test (zero external dependency)]"
    Debug.Print "    1) UserForm1.Show vbModeless      ' *spec fact 54* Show first"
    Debug.Print "    2) UserForm1.StartWebView2_Full"
    Debug.Print "    3) Wv2Log.LogStart                ' contain this one test run in one file"
    Debug.Print "    4) Test_N1_Capture"
    Debug.Print "    -> *Read the verdicts in the log file* the newest one under %APPDATA%\Wv2Browser\logs."
    Debug.Print "       Look at the trailing '*Verdicts n: OK x / FAIL y*'."
    Debug.Print ""
    Debug.Print "  [Isolation when it fails]"
    Debug.Print "    Test_N1_Site                      ' do the same on a real site"
    Debug.Print "    - 0 items there too      -> *a wiring problem* (vtable / IID / filter)"
    Debug.Print "    - only captured there    -> a problem with the targets of the verification page"
    Debug.Print ""
    Debug.Print "  [*Spec confirmed in N-1*]"
    Debug.Print "    Requests served through a virtual host (SetVirtualHostNameToFolderMapping)"
    Debug.Print "    *do not show up in WebResourceRequested* (all of DOCUMENT / fetch / XHR / images / CSS)."
    Debug.Print "    Because the runtime returns them directly with an internal resource handler."
    Debug.Print "    The targets of Test_N1_Capture avoid this by being placed on 2 routes:"
    Debug.Print "      Plan F  http://127.0.0.1:59999/...  unreachable local (zero external dependency)"
    Debug.Print "      Plan D  https://httpbingo.org/...   external service (needs the net)"
    Debug.Print ""
    Debug.Print "    *fetch() arrives as the XHR type, not FETCH*"
    Debug.Print "    WebView2 reports a request fired with JS fetch as XML_HTTP_REQUEST(7)."
    Debug.Print "    Do not assume when narrowing by type. The default filter"
    Debug.Print "    keeps FETCH(8) cast too (harmless, and it also catches it if it changes in the future)."
    Debug.Print ""
    Debug.Print "  [Using it on an actual site]"
    Debug.Print "    1) Open the page you want to investigate in a tab"
    Debug.Print "    2) Test_N1_Watch                  ' start capturing (DOCUMENT / XHR / FETCH)"
    Debug.Print "       Test_N1_Watch ""ALL""            ' when you want to see images and CSS too"
    Debug.Print "    3) *Operate the page by hand once* (log in, search, make it spit out a CSV...)"
    Debug.Print "    4) Test_N1_Drain                  ' flush the accumulated entries to the log and empty the buffer"
    Debug.Print "    5) Test_N1_Stop                   ' done (the filter is removed too)"
    Debug.Print ""
    Debug.Print "  --- How to read the log ---"
    Debug.Print "    #    elapsed ms  method type       URL"
    Debug.Print "    The elapsed ms is the time since capture started. *The point is that you can see the order and the intervals*."
    Debug.Print ""
    Debug.Print "  --- What N-1 cannot do (later stages) ---"
    Debug.Print "    - request headers and body          -> N-2"
    Debug.Print "    - response status and headers      -> N-3"
    Debug.Print "    - response body                      -> N-4"
    Debug.Print "    - turning it into PowerShell's Invoke-WebRequest  -> N-5"
    Debug.Print ""
    Debug.Print "  --- Cautions ---"
    Debug.Print "  - *Start capturing before Navigate* Otherwise you miss the first DOCUMENT."
    Debug.Print "  - The default ring is 500 entries. Overflow shows up in NetLogDropped (it is not cut silently)."
    Debug.Print "    If that is not enough it can be increased, e.g. p.NetLogCapacity = 5000."
    Debug.Print "  - *Spec fact 20* Do not break/step through until the event burst settles."
End Sub


' ============================================================
' N1Fire (N-1, Private) - fire one JS statement on the page side
'   *JS strings are always single quotes* (project rule)
' ============================================================
Private Sub N1Fire(ByVal p As Wv2Pane, ByVal js As String)
    p.EvalSync js, 5
    If Not p.LastEvalOk Then
        Wv2Log.LogI "        (N1Fire failed err=" & p.LastEvalError & ")"
    End If
End Sub


' ============================================================
' N1Find (N-1, Private) - find the first entry that matches the condition
'
'   If methodWant / ctxWant / uriPart is an empty string it means "do not care".
'   The return value is a 1-based position that can be passed to NetLogLine. 0 if not found.
'
'   *A tool so as not to settle for "it probably arrived"* (design principle 112)
' ============================================================
Private Function N1Find(ByVal p As Wv2Pane, _
                        ByVal methodWant As String, _
                        ByVal ctxWant As String, _
                        ByVal uriPart As String) As Long
    Dim i As Long
    Dim f As Variant

    For i = 1 To p.NetLogCount
        f = Split(p.NetLogLine(i), vbTab)
        If UBound(f) >= 4 Then
            If (LenB(methodWant) = 0 Or UCase$(CStr(f(2))) = UCase$(methodWant)) Then
                If (LenB(ctxWant) = 0 Or UCase$(CStr(f(3))) = UCase$(ctxWant)) Then
                    If (LenB(uriPart) = 0 Or _
                        InStr(1, CStr(f(4)), uriPart, vbTextCompare) > 0) Then
                        N1Find = i
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function


' ============================================================
' N1Wait (N-1, Private) - wait until one entry matching the condition arrives
'   True when it arrives. The row found is output to the log.
' ============================================================
Private Function N1Wait(ByVal p As Wv2Pane, _
                        ByVal methodWant As String, _
                        ByVal ctxWant As String, _
                        ByVal uriPart As String, _
                        ByVal timeoutSec As Single) As Boolean
    Dim t0 As Single
    Dim k  As Long

    t0 = Timer
    Do
        DoEvents
        k = N1Find(p, methodWant, ctxWant, uriPart)
        If k > 0 Then
            Wv2Log.LogI "        " & Replace$(p.NetLogLine(k), vbTab, "  ")
            N1Wait = True
            Exit Function
        End If
        If (Timer - t0) > timeoutSec Then Exit Do
    Loop
End Function


' ============================================================
' N1CtxOf (N-1c, Private) - returns *what type that request arrived as*
'   If not found, "(absent)". *A tool for observing without assuming the type*
' ============================================================
Private Function N1CtxOf(ByVal p As Wv2Pane, ByVal uriPart As String) As String
    Dim k As Long
    Dim f As Variant

    k = N1Find(p, "", "", uriPart)
    If k = 0 Then
        N1CtxOf = "(absent)"
        Exit Function
    End If

    f = Split(p.NetLogLine(k), vbTab)
    If UBound(f) >= 3 Then N1CtxOf = CStr(f(3))
End Function


' ============================================================
' N1Nav (N-1b, Private) - navigate and confirm the landing by title
'   Return value: True if it reached the target page.
' ============================================================
Private Function N1Nav(ByVal p As Wv2Pane, _
                       ByVal url As String, _
                       ByVal wantTitle As String, _
                       ByVal timeoutSec As Single) As Boolean
    If p.View_Navigate(url) <> 0 Then Exit Function
    N1Nav = D2WaitTitle(p, wantTitle, timeoutSec)
End Function


' ============================================================
' N1WriteFolder (N-1, Private) - write out the set of verification pages to %TEMP%
'   Return value: the absolute path of the folder written. Empty string on failure.
'
'   What goes there:
'     netprobe.html ... the entry. Has a form / fetch / CSS / image for both routes of targets
'     data.json     ... *the target on the virtual host (control)* Needed to count that it is not captured
'
'   *echo.html was dropped in N-1b* The form's destination was moved outside the virtual host (plan F / plan D),
'   so there is no need to place a destination in the same folder.
' ============================================================
Private Function N1WriteFolder() As String
    Dim folderPath As String

    folderPath = Environ$("TEMP")
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    folderPath = folderPath & N1_FOLDER

    If Not WriteUtf8NoBom(folderPath, "netprobe.html", BuildN1ProbeHtml()) Then Exit Function
    If Not WriteUtf8NoBom(folderPath, "data.json", "{""ok"":1}") Then Exit Function

    N1WriteFolder = folderPath
End Function


' ============================================================
' BuildN1ProbeHtml (the verification page of N-1b)
'
'   *It is static HTML, so quotation marks may be written with VBA's ""* (it is not JS)
'   *No JS at all is placed on this page itself* fetch / XHR are fired with EvalSync, so
'     by not putting any mechanism on the page side, the cause of what was sent narrows to one.
'
'   *All targets are outside the virtual host* (N-1b)
'     It turned out on the real machine that requests to the virtual host do not show up in WebResourceRequested, so
'     the destinations of form / link / img were placed on the 2 routes of plan F (127.0.0.1) and plan D (httpbingo.org).
'     Both point to nonexistent resources and so give 404 / connection failure, but
'     *all we want to see is the single line "a request went out"* (design principle 112).
' ============================================================
Private Function BuildN1ProbeHtml() As String
    Dim s As String

    s = "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    s = s & ("<title>N-1 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&) & "</title>") & vbLf

    ' --- resources that should not be captured (CSS). One for each route ---
    s = s & "<link rel=""stylesheet"" href=""" & N1_LOCAL & "/n1.css"">" & vbLf
    s = s & "<link rel=""stylesheet"" href=""" & N1_NET & "/n1.css"">" & vbLf

    s = s & "<style>" & vbLf
    s = s & "  body{font-family:'Segoe UI','Meiryo',sans-serif;padding:36px 28px;" & _
            "background:#0e121b;color:#e8eaed;}" & vbLf
    s = s & "  h1{font-size:22px;margin:0 0 6px;}" & vbLf
    s = s & "  .lead{font-size:13px;color:#9aa7bd;line-height:1.7;margin-bottom:22px;}" & vbLf
    s = s & "  .row{border:1px solid rgba(255,255,255,.12);border-radius:12px;" & _
            "padding:16px 18px;margin-bottom:14px;background:rgba(255,255,255,.04);}" & vbLf
    s = s & "  .tag{font-size:10.5px;letter-spacing:.08em;color:#6ea8fe;" & _
            "border:1px solid rgba(110,168,254,.4);border-radius:999px;padding:2px 9px;}" & vbLf
    s = s & "  button{font-family:inherit;font-size:14px;font-weight:600;color:#0e121b;" & _
            "background:#6ea8fe;border:0;border-radius:8px;padding:9px 18px;" & _
            "cursor:pointer;margin-top:10px;}" & vbLf
    s = s & "  .note{font-size:12px;color:#8ea2c8;margin-top:10px;line-height:1.6;}" & vbLf
    s = s & "  .ep{font-family:'Consolas','Courier New',monospace;font-size:11.5px;" & _
            "color:#7d8aa0;}" & vbLf
    s = s & "</style></head><body>" & vbLf
    s = s & "<h1>N-1 probe</h1>" & vbLf
    s = s & "<div class=""lead"">A page to check whether WebResourceRequested can really pick things up." & _
            "This page itself is served on the virtual host, so" & _
            "<b>loading the page itself does not show up in the event</b>." & _
            "All the targets are placed outside the virtual host.</div>" & vbLf

    ' --- plan F: POST to the unreachable local ---
    s = s & "<div class=""row"">" & vbLf
    s = s & "  <span class=""tag"">Plan F</span>" & vbLf
    s = s & "  <form id=""f-local"" action=""" & N1_LOCAL & "/echo"" method=""post"">" & vbLf
    s = s & "    <input type=""hidden"" name=""probe"" value=""wv2-n1-local"">" & vbLf
    s = s & ("    <input type=""hidden"" name=""nihongo"" value=""" & ChrW(&H65E5&) & ChrW(&H672C&) & ChrW(&H8A9E&) & ChrW(&H3000&) & ChrW(&H5168&) & ChrW(&H89D2&) & ChrW(&H30B9&) & ChrW(&H30DA&) & ChrW(&H30FC&) & ChrW(&H30B9&) & ChrW(&H5165&) & ChrW(&H308A&) & """>") & vbLf
    s = s & "    <button type=""submit"" id=""postbtn-local"">POST (local)</button>" & vbLf
    s = s & "  </form>" & vbLf
    s = s & "  <div class=""note"">The connection always fails, but" & _
            "it passes if <b>the fact that a request went out</b> is captured." & _
            "<br><span class=""ep"">" & N1_LOCAL & "/echo</span></div>" & vbLf
    s = s & "</div>" & vbLf

    ' --- plan D: POST to an external service ---
    s = s & "<div class=""row"">" & vbLf
    s = s & "  <span class=""tag"">Plan D</span>" & vbLf
    s = s & "  <form id=""f-net"" action=""" & N1_NET & "/post"" method=""post"">" & vbLf
    s = s & "    <input type=""hidden"" name=""probe"" value=""wv2-n1-net"">" & vbLf
    s = s & ("    <input type=""hidden"" name=""nihongo"" value=""" & ChrW(&H65E5&) & ChrW(&H672C&) & ChrW(&H8A9E&) & ChrW(&H3000&) & ChrW(&H5168&) & ChrW(&H89D2&) & ChrW(&H30B9&) & ChrW(&H30DA&) & ChrW(&H30FC&) & ChrW(&H30B9&) & ChrW(&H5165&) & ChrW(&H308A&) & """>") & vbLf
    s = s & "    <button type=""submit"" id=""postbtn-net"">POST (external)</button>" & vbLf
    s = s & "  </form>" & vbLf
    s = s & "  <div class=""note"">Needs the net. If it arrives, a 200 JSON is returned." & _
            "<br><span class=""ep"">" & N1_NET & "/post</span></div>" & vbLf
    s = s & "</div>" & vbLf

    ' --- resources that should not be captured (images). One for each route ---
    s = s & "<div class=""row"">" & vbLf
    s = s & "  <img src=""" & N1_LOCAL & "/n1.png"" alt="""" width=""1"" height=""1"">" & vbLf
    s = s & "  <img src=""" & N1_NET & "/image/png"" alt="""" width=""1"" height=""1"">" & vbLf
    s = s & "  <div class=""note"">CSS and images do have requests go out, but" & _
            "they are not captured by the default filter (DOCUMENT / XHR / FETCH)." & _
            "*We count that they are not captured to confirm*</div>" & vbLf
    s = s & "</div>" & vbLf

    s = s & "</body></html>"

    BuildN1ProbeHtml = s
End Function







' ============================================================
' Test_N2_Detail (regression test of N-2)
'
'   Uses the same targets as N-1 (plan F = unreachable local / plan D = httpbingo.org).
'   The page may stay served on the virtual host (only the event does not show up = spec fact 69).
'
'   What it looks at:
'     (0) *Detail is OFF by default* (issue 8 plan kappa)
'     (1) headers can be read / *GET does not read the body* (issue 7 plan alpha)
'     (2) the body can be read. *UTF-8 Japanese round-trips*
'     (3) *it is known that it was cut at the limit* (not cut silently)
'     (4) *Cookie-equivalent headers are masked* (issue 3 plan M)
'     (5) ** reading does not break the traffic ** (issue 2 plan X / issue 9 plan D')
'         -- this is what N-2 most wants to confirm
'     (6) the detail ring is separate from the list (issue 5 plan S)
'     (7) turning it back OFF stops only the detail from increasing
' ============================================================
Public Sub Test_N2_Detail()
    Dim b   As Wv2Browser
    Dim p   As Wv2Pane
    Dim el  As Wv2Element
    Dim folderPath As String
    Dim hr  As Long
    Dim k   As Long
    Dim n0  As Long
    Dim localOk As Boolean
    Dim netOk   As Boolean
    Dim hdr As String
    Dim txt As String

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_N2_Detail: Browser is not started."
        Exit Sub
    End If

    folderPath = N1WriteFolder()
    If LenB(folderPath) = 0 Then
        Wv2Log.LogI "Test_N2_Detail: failed to write out the verification page. Aborting."
        Exit Sub
    End If

    Set p = b.AddTab()
    If p Is Nothing Then
        Wv2Log.LogI "Test_N2_Detail: failed to create the tab."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_N2_Detail start ================"
    Wv2Log.LogI "        plan F target = " & N1_LOCAL & "  (unreachable local)"
    Wv2Log.LogI "        plan D target = " & N1_NET & "  (external service)"
    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (0) *Detail is OFF by default* (issue 8 plan kappa) ---"

    TestBool "*the default of NetDetailOn is False*", (p.NetDetailOn = False)
    TestBool "  the default capacity of the detail ring is 50", (p.NetDetailCapacity = 50)
    TestBool "  the default body limit is 64KB", (p.NetBodyMaxBytes = 65536)
    TestBool "  *the default of masking is ON*", (p.NetRedact = True)

    TestBool "NetCaptureStart succeeds", p.NetCaptureStart()
    hr = p.View3_SetVirtualHostNameToFolderMapping(N1_HOST, folderPath, 1)   ' 1 = ALLOW
    TestBool "  the virtual host mapping can be made", (hr = 0)
    hr = p.View_Navigate("https://" & N1_HOST & "/netprobe.html")
    TestBool "  Navigate succeeds", (hr = 0)

    If Not D2WaitTitle(p, ("N-1 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&)), 10) Then
        Wv2Log.LogI "Test_N2_Detail: could not confirm that the verification page loaded."
        p.NetCaptureStop
        TestCountPrint
        Exit Sub
    End If
    D3Pump 2

    ' *From here we take the detail*
    p.NetDetailOn = True
    p.NetLogClear
    p.NetDetailClear

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (1) Headers can be read / *GET does not read the body* ---"

    N1Fire p, "(function(){" & _
              "fetch('" & N1_LOCAL & "/n2h?k=hdr-local').catch(function(){});" & _
              "fetch('" & N1_NET & "/get?k=hdr-net').catch(function(){});" & _
              "return 1;})()"

    localOk = N2Wait(p, "k=hdr-local", 6)
    TestBool "*plan F: the detail can be obtained*", localOk
    netOk = N2Wait(p, "k=hdr-net", 8)
    TestBool "*plan D: the detail can be obtained*", netOk

    If Not (localOk Or netOk) Then
        Wv2Log.LogW "  NOTE: *no target arrived at all* Read the subsequent verdicts as having come up empty."
    End If

    k = N2Find(p, "k=hdr-local")
    If k = 0 Then k = N2Find(p, "k=hdr-net")
    If k > 0 Then
        hdr = p.NetDetailHeaders(k)
        Wv2Log.LogI "        headers (first 3 lines):"
        Wv2Log.LogI "        " & Replace$(Left$(hdr, 240), vbLf, " / ")
        TestBool "*at least one header can be read*", (InStr(1, hdr, ":") > 0)
        TestBool "*User-Agent is there*", (InStr(1, hdr, "user-agent", vbTextCompare) > 0)
        TestBool "*it is a GET, so the body is 0 bytes* (issue 7 plan alpha)", _
                 (Split(p.NetDetailLine(k), vbTab)(3) = "0")
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) The body can be read / *UTF-8 Japanese round-trips* ---"

    p.NetLogClear
    p.NetDetailClear
    ' *Passing a string body to fetch gives Content-Type: text/plain*
    '   This is a CORS "simple request" so no preflight is inserted.
    '   Even with an unreachable target the request itself goes out.
    N1Fire p, "(function(){var o={method:'POST'," & _
              ("body:'probe=wv2-n2-local&nihongo=" & ChrW(&H65E5&) & ChrW(&H672C&) & ChrW(&H8A9E&) & "'};") & _
              "fetch('" & N1_LOCAL & "/n2b',o).catch(function(){});" & _
              "return 1;})()"

    TestBool "*the POST detail can be obtained*", N2Wait(p, "/n2b", 6)
    k = N2Find(p, "/n2b")
    If k > 0 Then
        Wv2Log.LogI "        " & Replace$(p.NetDetailLine(k), vbTab, "  ")
        Wv2Log.LogI "        body = [" & p.NetDetailBody(k) & "]"
        TestBool "*the contents of the body can be read*", _
                 (InStr(1, p.NetDetailBody(k), "probe=wv2-n2-local") > 0)
        TestBool "*UTF-8 Japanese comes back without garbling*", _
                 (InStr(1, p.NetDetailBody(k), (ChrW(&H65E5&) & ChrW(&H672C&) & ChrW(&H8A9E&))) > 0)
        TestBool "  the byte count is filled in", (CLng(Split(p.NetDetailLine(k), vbTab)(3)) > 0)
        TestBool "  not truncated", (Split(p.NetDetailLine(k), vbTab)(4) = "False")
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) *It is known that it was cut at the limit* ---"

    p.NetLogClear
    p.NetDetailClear
    p.NetBodyMaxBytes = 100
    TestBool "the limit can be set to 100 bytes", (p.NetBodyMaxBytes = 100)

    N1Fire p, "(function(){var s='';for(var i=0;i<200;i++){s=s+'0123456789';}" & _
              "fetch('" & N1_LOCAL & "/n2big',{method:'POST',body:s}).catch(function(){});" & _
              "return 1;})()"

    TestBool "*the detail of the large body can be obtained*", N2Wait(p, "/n2big", 6)
    k = N2Find(p, "/n2big")
    If k > 0 Then
        Wv2Log.LogI "        " & Replace$(p.NetDetailLine(k), vbTab, "  ")
        TestBool "*it stops at 100 bytes*", (Split(p.NetDetailLine(k), vbTab)(3) = "100")
        TestBool "*it is known that it was cut*", (Split(p.NetDetailLine(k), vbTab)(4) = "True")
    End If
    p.NetBodyMaxBytes = 65536

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) *Secret headers are masked* (issue 3 plan M) ---"
    ' *A Cookie cannot be produced when the target is unreachable*, so the mechanism itself is
    '   checked with "an always present header" = User-Agent. Cookie / Authorization are
    '   in the default list (NetRedactDefaults).

    p.NetLogClear
    p.NetDetailClear
    p.AddRedactHeader "user-agent"

    N1Fire p, "(function(){fetch('" & N1_LOCAL & "/n2r?k=redact').catch(function(){});" & _
              "return 1;})()"
    TestBool "the detail can be obtained", N2Wait(p, "k=redact", 6)
    k = N2Find(p, "k=redact")
    If k > 0 Then
        hdr = p.NetDetailHeaders(k)
        TestBool "*it is masked*", (InStr(1, hdr, "<redacted") > 0)
        TestBool "*the content is not leaked*", (InStr(1, hdr, "Mozilla") = 0)

        p.NetRedact = False
        hdr = p.NetDetailHeaders(k)
        TestBool "*with NetRedact = False it can be seen raw*", (InStr(1, hdr, "Mozilla") > 0)
        p.NetRedact = True
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) ** Reading does not break the traffic ** (issue 2 plan X / issue 9 plan D') ---"
    ' *This is the crux of N-2*
    '   Reading an IStream advances its position. If Seek(0) fails to put it back,
    '   *WebView2 sends an empty body*. Rather than settling it with "the POST seemed to succeed",
    '   *count by matching against the body the destination returns* (design principle 112).
    '   httpbingo.org/post returns the contents of the received form as is in JSON.

    If Not netOk Then
        Wv2Log.LogW "  NOTE: the external host has not arrived, so this section is skipped. *The crux of N-2 stays unconfirmed*"
    Else
        p.NetLogClear
        p.NetDetailClear
        p.NetDetailUriFilter = "httpbingo.org/post"

        Set el = p.QuerySelector("#postbtn-net")
        TestBool "the submit button can be grabbed", Not (el Is Nothing)
        If Not (el Is Nothing) Then
            TestBool "  it can be clicked", el.Click()

            TestBool "*the POST detail can be obtained*", N2Wait(p, "httpbingo.org/post", 10)
            k = N2Find(p, "httpbingo.org/post")
            If k > 0 Then
                Wv2Log.LogI "        " & Replace$(p.NetDetailLine(k), vbTab, "  ")
                Wv2Log.LogI "        body = [" & Left$(p.NetDetailBody(k), 200) & "]"
                TestBool "*the body we read contains probe*", _
                         (InStr(1, p.NetDetailBody(k), "wv2-n1-net") > 0)
            End If

            ' *Did the destination really receive it?* Check with the body text
            txt = N2PageText(p, 15)
            Wv2Log.LogI "        response of the destination (first 200 chars):"
            Wv2Log.LogI "        " & Left$(Replace$(txt, vbLf, " "), 200)
            TestBool "** the body reached the destination unbroken **", _
                     (InStr(1, txt, "wv2-n1-net") > 0)
            TestBool "  *the Japanese also arrived unbroken*", _
                     (InStr(1, txt, (ChrW(&H65E5&) & ChrW(&H672C&) & ChrW(&H8A9E&))) > 0)
        End If

        p.NetDetailUriFilter = ""
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) The detail ring is separate from the list (issue 5 plan S) ---"

    p.NetLogClear
    p.NetDetailClear
    p.NetDetailCapacity = 2
    TestBool "only the detail capacity can be set to 2", (p.NetDetailCapacity = 2)
    TestBool "*the list capacity stays at 500*", (p.NetLogCapacity = 500)

    N1Fire p, "(function(){for(var i=0;i<5;i++){" & _
              "fetch('" & N1_LOCAL & "/n2ring?n='+i).catch(function(){});}return 5;})()"
    D3Pump 3

    Wv2Log.LogI "        list " & p.NetLogTotal & " items / detail " & p.NetDetailCount & " items"
    TestBool "*the list accumulates 5 or more*", (p.NetLogTotal >= 5)
    TestBool "*the detail is only 2, the capacity*", (p.NetDetailCount = 2)
    TestBool "*the detail also counts the overflow* (not cut silently)", _
             (p.NetDetailDropped > 0)
    p.NetDetailCapacity = 50

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (7) *Turning it back OFF stops only the detail from increasing* ---"

    p.NetDetailOn = False
    p.NetLogClear
    p.NetDetailClear
    N1Fire p, "(function(){fetch('" & N1_LOCAL & "/n2off?k=off').catch(function(){});" & _
              "return 1;})()"
    D3Pump 3

    n0 = p.NetLogTotal
    Wv2Log.LogI "        list " & n0 & " items / detail " & p.NetDetailCount & " items"
    TestBool "*the list increases (N-1 is alive)*", (n0 > 0)
    TestBool "*the detail does not increase at all*", (p.NetDetailCount = 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- Output all the details ---"
    p.NetDetailOn = True
    p.NetDetailClear
    N1Fire p, "(function(){fetch('" & N1_LOCAL & "/n2last?k=last').catch(function(){});" & _
              "return 1;})()"
    N2Wait p, "k=last", 6
    p.NetDetailAll

    p.NetCaptureStop

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback depth (expected 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_N2_Detail end ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' Test_N2_Help (procedure for N-2)
' ============================================================
Public Sub Test_N2_Help()
    Debug.Print "==== N-2 real-machine procedure (request headers and body) ===="
    Debug.Print ""
    Debug.Print "  [Regression test]"
    Debug.Print "    1) UserForm1.Show vbModeless      ' *spec fact 54*"
    Debug.Print "    2) UserForm1.StartWebView2_Full"
    Debug.Print "    3) Wv2Log.LogStart"
    Debug.Print "    4) Test_N2_Detail"
    Debug.Print "    -> Read the verdicts in the log file. *The crux is (5) 'the body reached the destination"
    Debug.Print "       unbroken'* (needs an external host)."
    Debug.Print ""
    Debug.Print "  [Using it on an actual site]"
    Debug.Print "    1) Open the page you want to investigate in a tab"
    Debug.Print "    2) Set p = UserForm1.CurrentBrowser.ActivePane"
    Debug.Print "       p.NetDetailOn = True          ' *detail is OFF by default*"
    Debug.Print "       Test_N1_Watch"
    Debug.Print "    3) *Operate the page by hand once*"
    Debug.Print "    4) Test_N1_Drain                 ' first find a hit in the list"
    Debug.Print "       p.NetDetailAll                ' output all the details"
    Debug.Print "       p.NetDetail 3                 ' output only the 3rd"
    Debug.Print "    5) Test_N1_Stop"
    Debug.Print ""
    Debug.Print "  --- Narrowing down ---"
    Debug.Print "    p.NetDetailUriFilter = ""/api/"" ' take details only for this URL"
    Debug.Print "    p.NetBodyMaxBytes = 200000       ' body limit (default 64KB)"
    Debug.Print "    p.NetDetailCapacity = 200        ' capacity of the detail ring (default 50)"
    Debug.Print ""
    Debug.Print "  --- *Handling of secrets* ---"
    Debug.Print "    By default Cookie / Authorization / X-Api-Key etc. are masked on output."
    Debug.Print "    p.AddRedactHeader ""x-my-token""   ' add a target to mask"
    Debug.Print "    p.NetRedact = False              ' *comes out in plain text. Be careful before handing the log to someone*"
    Debug.Print "    NOTE: on the ring it is held raw. The masking is only at output time."
    Debug.Print ""
    Debug.Print "  --- *What we were careful about in N-2* ---"
    Debug.Print "  - An IStream advances its position when read. *If it is not put back, an empty body is sent*, so"
    Debug.Print "    Seek(0) is done after reading. That nothing is broken is counted in (5)."
    Debug.Print "  - The body can only be read synchronously inside the handler (args are valid"
    Debug.Print "    only during Invoke). So *always cut at the limit*. That it was cut shows in the detail row."
    Debug.Print "  - The body of a GET is not read (there is almost none)."
    Debug.Print ""
    Debug.Print "  --- What N-2 cannot do (later stages) ---"
    Debug.Print "    - response status and headers      -> N-3"
    Debug.Print "    - response body                      -> N-4"
    Debug.Print "    - turning it into PowerShell's Invoke-WebRequest  -> N-5"
End Sub


' ============================================================
' N2Find (N-2, Private) - find a detail by partial match of the URL
'   The return value is a 1-based position that can be passed to NetDetailLine etc. 0 if absent.
' ============================================================
Private Function N2Find(ByVal p As Wv2Pane, ByVal uriPart As String) As Long
    Dim i As Long
    Dim f As Variant

    For i = 1 To p.NetDetailCount
        f = Split(p.NetDetailLine(i), vbTab)
        If UBound(f) >= 2 Then
            If InStr(1, CStr(f(2)), uriPart, vbTextCompare) > 0 Then
                N2Find = i
                Exit Function
            End If
        End If
    Next i
End Function


' ============================================================
' N2Wait (N-2, Private) - wait until that detail arrives
' ============================================================
Private Function N2Wait(ByVal p As Wv2Pane, _
                        ByVal uriPart As String, _
                        ByVal timeoutSec As Single) As Boolean
    Dim t0 As Single

    t0 = Timer
    Do
        DoEvents
        If N2Find(p, uriPart) > 0 Then
            N2Wait = True
            Exit Function
        End If
        If (Timer - t0) > timeoutSec Then Exit Do
    Loop
End Function


' ============================================================
' N2PageText (N-2, Private) - read the body text of the current page
'   Used to confirm *what the destination actually received*.
'   Waits a little until it becomes readable (it may be empty right after navigation).
' ============================================================
Private Function N2PageText(ByVal p As Wv2Pane, ByVal timeoutSec As Single) As String
    Dim t0  As Single
    Dim res As String
    Dim cur As String

    t0 = Timer
    Do
        DoEvents
        res = p.EvalSync("document.body.innerText", 5)
        If p.LastEvalOk Then
            cur = Wv2Json.JsonUnescape(res)
            If Len(cur) > 20 Then
                N2PageText = cur
                Exit Function
            End If
        End If
        If (Timer - t0) > timeoutSec Then Exit Do
    Loop
    N2PageText = cur
End Function


' ============================================================
' Test_N3_Status (regression test of N-3)
'
'   Because *this event has no filter*, only those corresponding to requests that are in the list are
'   recorded (issue 1 plan A). The targets are the same 2 routes as N-1 / N-2.
'
'   What it looks at:
'     (0) the response event is cast (NetRespOn)
'     (1) *an ordinary 200 is attached to the list row*
'     (2) *404 and 500 appear as they are* (do not look only at success)
'     (3) *a redirect is lined up with 1 request and 1 response* (issue 6 plan Y)
'     (4) response headers can be read / *masking also works on response headers*
'         *the section that looks at response contents waits with N3WaitAny* (with N2Wait it is satisfied the instant
'         the request goes out, so it always reads empty)
'     (5) *an unreachable target gets no response = blank* (issue 7 plan alpha)
'     (6) *responses not in the list are only counted* (issue 5 plan W)
'     (7) if detail is OFF, response headers are not taken (the status is attached)
'
'   *(1) to (4) need an external host* An unreachable local gets no response, so
'   verifying the status needs a real destination. When it does not arrive, that is stated and skipped.
' ============================================================
Public Sub Test_N3_Status()
    Dim b   As Wv2Browser
    Dim p   As Wv2Pane
    Dim folderPath As String
    Dim hr  As Long
    Dim k   As Long
    Dim n0  As Long
    Dim netOk As Boolean
    Dim hdr As String

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_N3_Status: Browser is not started."
        Exit Sub
    End If

    folderPath = N1WriteFolder()
    If LenB(folderPath) = 0 Then
        Wv2Log.LogI "Test_N3_Status: failed to write out the verification page. Aborting."
        Exit Sub
    End If

    Set p = b.AddTab()
    If p Is Nothing Then
        Wv2Log.LogI "Test_N3_Status: failed to create the tab."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_N3_Status start ================"
    Wv2Log.LogI "        plan F target = " & N1_LOCAL & "  (unreachable local)"
    Wv2Log.LogI "        plan D target = " & N1_NET & "  (external service)"
    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (0) The response event is cast ---"

    TestBool "NetCaptureStart succeeds", p.NetCaptureStart()
    TestBool "*the response event is also cast (NetRespOn)*", (p.NetRespOn = True)

    hr = p.View3_SetVirtualHostNameToFolderMapping(N1_HOST, folderPath, 1)   ' 1 = ALLOW
    TestBool "  the virtual host mapping can be made", (hr = 0)
    hr = p.View_Navigate("https://" & N1_HOST & "/netprobe.html")
    TestBool "  Navigate succeeds", (hr = 0)

    If Not D2WaitTitle(p, ("N-1 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&)), 10) Then
        Wv2Log.LogI "Test_N3_Status: could not confirm that the verification page loaded."
        p.NetCaptureStop
        TestCountPrint
        Exit Sub
    End If
    D3Pump 2

    p.NetDetailOn = True
    p.NetLogClear
    p.NetDetailClear

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (1) *An ordinary 200 is attached to the list row* ---"

    N1Fire p, "(function(){fetch('" & N1_NET & "/get?k=n3-200').catch(function(){});" & _
              "return 1;})()"
    netOk = N3Wait(p, "k=n3-200", 200, 12)
    TestBool "*200 is attached*", netOk
    If netOk Then
        k = N1Find(p, "", "", "k=n3-200")
        Wv2Log.LogI "        " & Replace$(p.NetLogLine(k), vbTab, "  ")
        TestBool "  can also be read with NetLogStatus", (p.NetLogStatus(k) = 200)
    Else
        Wv2Log.LogW "  NOTE: *the external host has not arrived* Read (1) to (4) as having come up empty."
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) *404 and 500 appear as they are* ---"
    ' *Do not look only at success* The value of an investigation tool is that failures are visible as failures.

    If netOk Then
        p.NetLogClear
        N1Fire p, "(function(){" & _
                  "fetch('" & N1_NET & "/status/404?k=n3-404').catch(function(){});" & _
                  "fetch('" & N1_NET & "/status/500?k=n3-500').catch(function(){});" & _
                  "return 1;})()"
        TestBool "*404 is attached*", N3Wait(p, "k=n3-404", 404, 12)
        TestBool "*500 is attached*", N3Wait(p, "k=n3-500", 500, 12)
        k = N1Find(p, "", "", "k=n3-404")
        If k > 0 Then Wv2Log.LogI "        " & Replace$(p.NetLogLine(k), vbTab, "  ")
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) *A redirect is lined up with 1 request and 1 response* (issue 6 plan Y) ---"
    ' With a redirect the request itself fires several times, so responses are also attached one to one.
    ' *Check that "two responses get attached to the same URL and overwrite" does not happen*

    If netOk Then
        p.NetLogClear
        N1Fire p, "(function(){fetch('" & N1_NET & "/redirect/2?k=n3-redir')" & _
                  ".catch(function(){});return 1;})()"
        D3Pump 6
        n0 = N3CountWithStatus(p)
        Wv2Log.LogI "        list " & p.NetLogCount & " items / of which those with a response attached " & n0 & " items"
        N3DumpAll p
        TestBool "*a redirect lines up multiple rows*", (p.NetLogCount >= 2)
        TestBool "*every row has exactly one response attached*", (n0 = p.NetLogCount)
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) Response headers can be read / *masking also works on response headers* ---"

    If netOk Then
        p.NetLogClear
        p.NetDetailClear
        N1Fire p, "(function(){fetch('" & N1_NET & _
                  "/response-headers?X-N3-Test=hello&k=n3-hdr').catch(function(){});" & _
                  "return 1;})()"
        ' *Wait until the response is attached (N-3c)*
        '   N2Wait is "until the detail appears" = it is satisfied the instant the request goes out.
        '   If you want to read the response headers, it is meaningless unless you wait *until the response arrives*.
        TestBool "*can wait until the response is attached*", N3WaitAny(p, "k=n3-hdr", 15)
        k = N2Find(p, "k=n3-hdr")
        If k > 0 Then
            hdr = p.NetDetailRespHeaders(k)
            Wv2Log.LogI "        response headers: " & Replace$(Left$(hdr, 300), vbLf, " / ")
            ' *Output how far it got, as a whole (N-3b)*
            '   If the [response] line appears, it has reached at least NetAttachRespHeaders.
            '   If it does not, it is failing at the matching of detail entries.
            p.NetDetail k
            TestBool "*the response headers can be read*", (InStr(1, hdr, ":") > 0)
            TestBool "*the specified header has come back*", _
                     (InStr(1, hdr, "X-N3-Test", vbTextCompare) > 0)
            TestBool "  held separately from the request headers", _
                     (p.NetDetailHeaders(k) <> hdr)
        End If

        ' --- *Does masking work on the response headers side too?* ---
        '   *Do not rely on an external Set-Cookie* Add the content-type, which can already be read,
        '   to the masked set and *compare the same single entry before and after masking*.
        '   This is also direct proof of issue 3 plan M (mask at output time, not at storage time).
        If k > 0 Then
            p.AddRedactHeader "content-type"
            hdr = p.NetDetailRespHeaders(k)
            Wv2Log.LogI "        after masking: " & Replace$(Left$(hdr, 240), vbLf, " / ")
            TestBool "*masking also works on the response headers*", (InStr(1, hdr, "<redacted") > 0)
            TestBool "*the content (application/json) is gone*", _
                     (InStr(1, hdr, "application/json") = 0)

            ' *It is held raw on the ring*, so setting False brings it back
            p.NetRedact = False
            hdr = p.NetDetailRespHeaders(k)
            TestBool "*held raw on the ring (comes back with NetRedact = False)*", _
                     (InStr(1, hdr, "application/json") > 0)
            p.NetRedact = True
        End If

        ' --- Also try with a real Set-Cookie (an add-on. If it does not arrive, record it as unverified) ---
        '   *Do not let it pass silently when it did not run* Looking like FAIL 0 while in fact
        '   nothing was confirmed is the worst.
        p.NetLogClear
        p.NetDetailClear
        N1Fire p, "(function(){fetch('" & N1_NET & _
                  "/cookies/set?n3secret=topsecret123&k=n3-cookie'" & _
                  ",{redirect:'manual'}).catch(function(){});return 1;})()"
        If N3WaitAny(p, "k=n3-cookie", 12) Then
            k = N2Find(p, "k=n3-cookie")
            hdr = p.NetDetailRespHeaders(k)
            Wv2Log.LogI "        response headers: " & Replace$(Left$(hdr, 300), vbLf, " / ")
            If InStr(1, hdr, "et-cookie", vbTextCompare) > 0 Then
                TestBool "*the contents of the real Set-Cookie are also masked*", _
                         (InStr(1, hdr, "topsecret123") = 0)
            Else
                Wv2Log.LogW "  NOTE: *there was no Set-Cookie in the response headers* " & _
                            "The verdict of this add-on did not run (unverified)."
            End If
        Else
            Wv2Log.LogW "  NOTE: *the Set-Cookie add-on verdict did not run (no response arrived)* " & _
                        "perhaps a request with redirect manual does not show up in the response event." & _
                        "We do not assert it and put it on the open items. The masking mechanism itself was verified above."
        End If
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) *An unreachable target gets no response = blank* (issue 7 plan alpha) ---"
    ' *Confirming that 0 is not written* Do not mix "not arrived yet" with a real 0.

    p.NetLogClear
    p.NetDetailClear
    N1Fire p, "(function(){fetch('" & N1_LOCAL & "/n3none?k=n3-none').catch(function(){});" & _
              "return 1;})()"
    TestBool "the request appears in the list", N1Wait(p, "", "", "k=n3-none", 8)
    D3Pump 3
    k = N1Find(p, "", "", "k=n3-none")
    If k > 0 Then
        Wv2Log.LogI "        " & Replace$(p.NetLogLine(k), vbTab, "  ")
        TestBool "*the status stays 0 (not arrived)*", (p.NetLogStatus(k) = 0)
        TestBool "*the display is blank (0 is not written)*", _
                 (Split(p.NetLogLine(k), vbTab)(5) = "")
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) *Responses not in the list are only counted* (issue 5 plan W) ---"
    ' *Deliberately fire a request outside the capture targets and count it*
    '   An image is of the IMAGE type so it does not go on the list (DOCUMENT / XHR / FETCH).
    '   But this event has no filter, so *the response does arrive*.
    '   We check that it is counted as "no counterpart" and does not appear in the list.
    '   *Do not make it an always-true verdict like ">= 0"* (that confirms nothing)

    If netOk Then
        p.NetLogClear
        n0 = p.NetRespUnmatched
        N1Fire p, "(function(){var im=new Image();" & _
                  "im.src='" & N1_NET & "/image/png?k=n3-orphan';" & _
                  "return 1;})()"
        D3Pump 6
        Wv2Log.LogI "        responses with no counterpart in the list = " & p.NetRespUnmatched & " items " & _
                    "(right after clearing it was " & n0 & " items)"
        TestBool "*responses outside the capture targets are counted*", (p.NetRespUnmatched > n0)
        TestBool "*that image is not on the list*", (N1Find(p, "", "", "k=n3-orphan") = 0)
        TestBool "  list misses are accounted separately (NetLogDropped)", (p.NetLogDropped = 0)
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (7) The status is attached even when detail is OFF ---"
    ' *Only the response headers count as detail (issue 4 plan U)* The status is part of the list.

    If netOk Then
        p.NetDetailOn = False
        p.NetLogClear
        p.NetDetailClear
        N1Fire p, "(function(){fetch('" & N1_NET & "/get?k=n3-off').catch(function(){});" & _
                  "return 1;})()"
        TestBool "*the status is attached even with detail OFF*", N3Wait(p, "k=n3-off", 200, 12)
        TestBool "  the detail does not increase at all", (p.NetDetailCount = 0)
        p.NetDetailOn = True
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- Output everything remaining on hand ---"
    p.NetLogDrain
    p.NetCaptureStop

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback depth (expected 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_N3_Status end ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' Test_N3_Help (procedure for N-3)
' ============================================================
Public Sub Test_N3_Help()
    Debug.Print "==== N-3 real-machine procedure (response status and headers) ===="
    Debug.Print ""
    Debug.Print "  [Regression test]"
    Debug.Print "    1) UserForm1.Show vbModeless      ' *spec fact 54*"
    Debug.Print "    2) UserForm1.StartWebView2_Full"
    Debug.Print "    3) Wv2Log.LogStart"
    Debug.Print "    4) Test_N3_Status"
    Debug.Print "    -> *(1) to (4) need an external host (httpbingo.org)*"
    Debug.Print "       An unreachable local gets no response, so verifying the status"
    Debug.Print "       needs a real destination. When it does not arrive, it is shown as having come up empty."
    Debug.Print ""
    Debug.Print "  [How to read the list (one column added in N-3)]"
    Debug.Print "    #   elapsed ms  method type       status  URL"
    Debug.Print "    1       94  GET      DOCUMENT    200  https://..."
    Debug.Print "    2      141  GET      XHR         404  https://..."
    Debug.Print "    3      609  GET      XHR              https://...   <- *blank = not arrived yet*"
    Debug.Print ""
    Debug.Print "  [Response headers]"
    Debug.Print "    p.NetDetailOn = True             ' *response headers count as detail*"
    Debug.Print "    p.NetDetail 3                    ' the 3rd detail (request + response)"
    Debug.Print "    p.NetDetailRespHeaders(3)        ' take only the response headers"
    Debug.Print "    NOTE: Set-Cookie is masked by default (NetRedact)."
    Debug.Print ""
    Debug.Print "  --- *What we were careful about in N-3* ---"
    Debug.Print "  - *This event has no filter* Images and CSS all arrive. So"
    Debug.Print "    only those corresponding to requests already in the list are recorded. If you want to see everything,"
    Debug.Print "    widen the list side with Test_N1_Watch ""ALL"" and the responses follow."
    Debug.Print "  - Matching is by URI + method. *Concurrent requests to the same URL"
    Debug.Print "    get mixed up, and the mix-up cannot be detected* (args have no clue"
    Debug.Print "    for linking). A redirect fires the request several times so it is one to one."
    Debug.Print "  - *NetLogDrain also erases rows whose response has not yet arrived*, so"
    Debug.Print "    those responses go to NetRespUnmatched. Do not flush in a hurry."
    Debug.Print "  - A large NetRespUnmatched is *normal* (responses outside the capture targets)."
    Debug.Print "    When suspecting misses, look at NetLogDropped."
    Debug.Print ""
    Debug.Print "  --- What N-3 cannot do (later stages) ---"
    Debug.Print "    - response body                      -> N-4 (GetContent is asynchronous)"
    Debug.Print "    - turning it into PowerShell's Invoke-WebRequest  -> N-5"
End Sub


' ============================================================
' N3Wait (N-3, Private) - wait until the specified status is attached to that URL
' ============================================================
Private Function N3Wait(ByVal p As Wv2Pane, _
                        ByVal uriPart As String, _
                        ByVal wantStatus As Long, _
                        ByVal timeoutSec As Single) As Boolean
    Dim t0 As Single
    Dim k  As Long

    t0 = Timer
    Do
        DoEvents
        k = N1Find(p, "", "", uriPart)
        If k > 0 Then
            If p.NetLogStatus(k) = wantStatus Then
                N3Wait = True
                Exit Function
            End If
        End If
        If (Timer - t0) > timeoutSec Then Exit Do
    Loop
End Function


' ============================================================
' N3WaitAny (N-3c, Private) - *wait until a response is attached, whatever the status*
'
'   N3Wait is "until this status is attached", N2Wait is "until the detail appears".
'   *The detail is created the instant the request goes out*, so waiting with N2Wait and reading
'   the response headers always gives empty (we stepped on this in the first N-3 real-machine run).
'   When you want to see the response contents, always wait with this one.
' ============================================================
Private Function N3WaitAny(ByVal p As Wv2Pane, _
                           ByVal uriPart As String, _
                           ByVal timeoutSec As Single) As Boolean
    Dim t0 As Single
    Dim k  As Long

    t0 = Timer
    Do
        DoEvents
        k = N1Find(p, "", "", uriPart)
        If k > 0 Then
            If p.NetLogStatus(k) <> 0 Then
                N3WaitAny = True
                Exit Function
            End If
        End If
        If (Timer - t0) > timeoutSec Then Exit Do
    Loop
End Function


' ============================================================
' N3CountWithStatus (N-3, Private) - the number of rows that have a response attached
' ============================================================
Private Function N3CountWithStatus(ByVal p As Wv2Pane) As Long
    Dim i As Long
    For i = 1 To p.NetLogCount
        If p.NetLogStatus(i) <> 0 Then N3CountWithStatus = N3CountWithStatus + 1
    Next i
End Function


' ============================================================
' N3DumpAll (N-3, Private) - output the list to the log as is (does not erase)
'   *Unlike NetLogDrain it does not empty it* because it would be a problem if it vanished midway while used for verdicts.
' ============================================================
Private Sub N3DumpAll(ByVal p As Wv2Pane)
    Dim i As Long
    For i = 1 To p.NetLogCount
        Wv2Log.LogI "        " & Replace$(p.NetLogLine(i), vbTab, "  ")
    Next i
End Sub



' ============================================================
' Test_N4_Body (regression test of N-4)
'
'   *N-4 is the first stage that is asynchronous* GetContent takes a handler and returns immediately,
'   and the body arrives later. So *always wait "until the body arrives"* (design principle 120).
'   Do not repeat the failure of N-3 where we waited for "until the detail appears" and read empty.
'
'   What it looks at:
'     (0) *the body is OFF by default* (issue 1 plan A)
'     (1) the JSON body can be read / *everything that arrived is kept* (issue 2 plan P')
'     (2) *the handling of compression settles to one or the other* -- the biggest unknown of N-4
'     (3) HTML can be read too
'     (4) binary is not output as contents but kept as hex
'     (5) *it is known that it was cut at the limit*
'     (6) *distinguish a response with no body (204) from "not arrived yet"*
'     (7) turning it back OFF stops only the body from being taken (the status and headers are attached)
'
'   *(1) to (6) need an external host* An unreachable local gets no response, so no body either.
' ============================================================
Public Sub Test_N4_Body()
    Dim b   As Wv2Browser
    Dim p   As Wv2Pane
    Dim folderPath As String
    Dim hr  As Long
    Dim k   As Long
    Dim netOk As Boolean
    Dim bodyTxt As String
    Dim hexHead As String
    Dim f   As Variant
    Dim isPlain As Boolean
    Dim isGz    As Boolean

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_N4_Body: Browser is not started."
        Exit Sub
    End If

    folderPath = N1WriteFolder()
    If LenB(folderPath) = 0 Then
        Wv2Log.LogI "Test_N4_Body: failed to write out the verification page. Aborting."
        Exit Sub
    End If

    Set p = b.AddTab()
    If p Is Nothing Then
        Wv2Log.LogI "Test_N4_Body: failed to create the tab."
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_N4_Body start ================"
    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (0) *The body is OFF by default* (issue 1 plan A) ---"

    TestBool "*the default of NetRespBodyOn is False*", (p.NetRespBodyOn = False)
    TestBool "  the default body limit is 64KB", (p.NetRespBodyMaxBytes = 65536)

    TestBool "NetCaptureStart succeeds", p.NetCaptureStart()
    hr = p.View3_SetVirtualHostNameToFolderMapping(N1_HOST, folderPath, 1)   ' 1 = ALLOW
    TestBool "  the virtual host mapping can be made", (hr = 0)
    hr = p.View_Navigate("https://" & N1_HOST & "/netprobe.html")
    TestBool "  Navigate succeeds", (hr = 0)

    If Not D2WaitTitle(p, ("N-1 " & ChrW(&H30D7&) & ChrW(&H30ED&) & ChrW(&H30FC&) & ChrW(&H30D6&)), 10) Then
        Wv2Log.LogI "Test_N4_Body: could not confirm that the verification page loaded."
        p.NetCaptureStop
        TestCountPrint
        Exit Sub
    End If
    D3Pump 2

    ' *Taking the body needs the detail as well* (because the storage place is the detail entry)
    p.NetDetailOn = True
    p.NetRespBodyOn = True

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (1) The JSON body can be read / *everything that arrived is kept* ---"

    p.NetLogClear
    p.NetDetailClear
    N1Fire p, "(function(){fetch('" & N1_NET & "/json?k=n4-json').catch(function(){});" & _
              "return 1;})()"

    netOk = N4Wait(p, "k=n4-json", 15)
    TestBool "*can wait until the body arrives*", netOk
    If Not netOk Then
        Wv2Log.LogW "  NOTE: *the external host has not arrived* Read (1) to (6) as having come up empty."
    Else
        k = N2Find(p, "k=n4-json")
        Wv2Log.LogI "        identity = " & Replace$(p.NetDetailRespLine(k), vbTab, "  ")
        Wv2Log.LogI "        body (first 160 chars) = " & Left$(p.NetDetailRespBody(k), 160)
        f = Split(p.NetDetailRespLine(k), vbTab)
        TestBool "*the byte count is filled in*", (CLng(f(0)) > 0)
        TestBool "*it is judged to be text*", (f(2) = "True")
        TestBool "*the hex of the first 32 bytes is kept*", (LenB(f(4)) > 0)
        TestBool "*the contents of the JSON can be read*", _
                 (InStr(1, p.NetDetailRespBody(k), "slideshow") > 0 Or _
                  InStr(1, p.NetDetailRespBody(k), "{") > 0)
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) *Settling the handling of compression* (the biggest unknown of N-4) ---"
    ' *Do not assume (design principle 117)*
    '   httpbingo's /gzip returns gzip-compressed JSON. Its contents
    '   contain "gzipped": true. Therefore
    '     - if gzipped can be read in the body -> GetContent returns *decoded*
    '     - if the leading hex is 1F 8B       -> GetContent returns *still compressed*
    '   exactly one of the two must hold. That is made the verdict.

    If netOk Then
        p.NetLogClear
        p.NetDetailClear
        N1Fire p, "(function(){fetch('" & N1_NET & "/gzip?k=n4-gzip').catch(function(){});" & _
                  "return 1;})()"
        If N4Wait(p, "k=n4-gzip", 15) Then
            k = N2Find(p, "k=n4-gzip")
            Wv2Log.LogI "        identity = " & Replace$(p.NetDetailRespLine(k), vbTab, "  ")
            bodyTxt = p.NetDetailRespBody(k)
            f = Split(p.NetDetailRespLine(k), vbTab)
            hexHead = CStr(f(4))
            Wv2Log.LogI "        body (first 120 chars) = " & Left$(bodyTxt, 120)

            isPlain = (InStr(1, bodyTxt, "gzipped") > 0)
            isGz = (Left$(hexHead, 5) = "1F 8B")

            TestBool "*the handling of compression settled to one or the other*", (isPlain Xor isGz)
            If isPlain Then
                Wv2Log.LogI "        -> *GetContent returns the decoded body*"
            ElseIf isGz Then
                Wv2Log.LogI "        -> *GetContent returns it still compressed* " & _
                            "(VBA cannot decompress it, so design it to keep only the encoding and the byte count)"
            Else
                Wv2Log.LogW "        -> *Cannot say either way* Read the identity and the body as they are."
            End If
        End If
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) HTML can be read too ---"

    If netOk Then
        p.NetLogClear
        p.NetDetailClear
        N1Fire p, "(function(){fetch('" & N1_NET & "/html?k=n4-html').catch(function(){});" & _
                  "return 1;})()"
        If N4Wait(p, "k=n4-html", 15) Then
            k = N2Find(p, "k=n4-html")
            Wv2Log.LogI "        identity = " & Replace$(p.NetDetailRespLine(k), vbTab, "  ")
            TestBool "*the contents of the HTML can be read*", _
                     (InStr(1, p.NetDetailRespBody(k), "<html", vbTextCompare) > 0 Or _
                      InStr(1, p.NetDetailRespBody(k), "Melville") > 0)
        End If
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) Binary is not output as contents but kept as hex ---"

    If netOk Then
        p.NetLogClear
        p.NetDetailClear
        N1Fire p, "(function(){fetch('" & N1_NET & "/image/png?k=n4-png')" & _
                  ".catch(function(){});return 1;})()"
        If N4Wait(p, "k=n4-png", 15) Then
            k = N2Find(p, "k=n4-png")
            Wv2Log.LogI "        identity = " & Replace$(p.NetDetailRespLine(k), vbTab, "  ")
            f = Split(p.NetDetailRespLine(k), vbTab)
            TestBool "*it is judged to be binary*", (f(2) = "False")
            TestBool "  the body is empty (contents are not output)", (LenB(p.NetDetailRespBody(k)) = 0)
            TestBool "*the hex is kept*", (LenB(f(4)) > 0)
        End If
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) *It is known that it was cut at the limit* ---"

    If netOk Then
        p.NetLogClear
        p.NetDetailClear
        p.NetRespBodyMaxBytes = 100
        TestBool "the limit can be set to 100 bytes", (p.NetRespBodyMaxBytes = 100)
        N1Fire p, "(function(){fetch('" & N1_NET & "/json?k=n4-cut').catch(function(){});" & _
                  "return 1;})()"
        If N4Wait(p, "k=n4-cut", 15) Then
            k = N2Find(p, "k=n4-cut")
            Wv2Log.LogI "        identity = " & Replace$(p.NetDetailRespLine(k), vbTab, "  ")
            f = Split(p.NetDetailRespLine(k), vbTab)
            TestBool "*it stops at 100 bytes*", (f(0) = "100")
            TestBool "*it is known that it was cut*", (f(1) = "True")
        End If
        p.NetRespBodyMaxBytes = 65536
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) *Distinguish a response with no body from 'not arrived yet'* ---"
    ' 204 has no body. It is 0 bytes but *it did arrive*.
    ' An unreachable target *does not arrive in the first place*. Do not mix these two (design principle 111).

    If netOk Then
        p.NetLogClear
        p.NetDetailClear
        N1Fire p, "(function(){fetch('" & N1_NET & "/status/204?k=n4-204')" & _
                  ".catch(function(){});return 1;})()"
        If N4Wait(p, "k=n4-204", 15) Then
            k = N2Find(p, "k=n4-204")
            Wv2Log.LogI "        identity = " & Replace$(p.NetDetailRespLine(k), vbTab, "  ")
            f = Split(p.NetDetailRespLine(k), vbTab)
            TestBool "*0 bytes but it has arrived*", (f(0) = "0" And f(5) = "True")
            ' *204 completes with ERROR_NO_DATA. This is not a failure*
            '   Writing "fetch failed" would mislead N-5 about 'whether there is a body' (design principle 111).
            TestBool "*do not write 'fetch failed' (there is simply no body)*", _
                     (InStr(1, CStr(f(3)), "fetch failed") = 0)
        End If
    End If

    p.NetLogClear
    p.NetDetailClear
    N1Fire p, "(function(){fetch('" & N1_LOCAL & "/n4none?k=n4-none')" & _
              ".catch(function(){});return 1;})()"
    D3Pump 4
    k = N2Find(p, "k=n4-none")
    If k > 0 Then
        Wv2Log.LogI "        identity = " & Replace$(p.NetDetailRespLine(k), vbTab, "  ")
        f = Split(p.NetDetailRespLine(k), vbTab)
        TestBool "*if unreachable it stays 'not arrived'*", (f(5) = "False")
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (7) *Turning it back OFF stops only the body from being taken* ---"
    ' *The status and headers are attached* = confirms that the body alone is a separate switch.

    If netOk Then
        p.NetRespBodyOn = False
        p.NetLogClear
        p.NetDetailClear
        N1Fire p, "(function(){fetch('" & N1_NET & "/json?k=n4-off').catch(function(){});" & _
                  "return 1;})()"
        If N3WaitAny(p, "k=n4-off", 15) Then
            D3Pump 3
            k = N2Find(p, "k=n4-off")
            Wv2Log.LogI "        identity = " & Replace$(p.NetDetailRespLine(k), vbTab, "  ")
            f = Split(p.NetDetailRespLine(k), vbTab)
            TestBool "*the body does not arrive*", (f(5) = "False")
            TestBool "*the status is attached*", (p.NetLogStatus(N1Find(p, "", "", "k=n4-off")) > 0)
            TestBool "*the response headers are attached too*", (InStr(1, p.NetDetailRespHeaders(k), ":") > 0)
        End If
        p.NetRespBodyOn = True
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "        bodies with no place to store = " & p.NetRespBodyOrphan & " items"

    p.NetCaptureStop

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback depth (expected 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_N4_Body end ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' Test_N4_Help (procedure for N-4)
' ============================================================
Public Sub Test_N4_Help()
    Debug.Print "==== N-4 real-machine procedure (response body) ===="
    Debug.Print ""
    Debug.Print "  [Regression test]"
    Debug.Print "    1) UserForm1.Show vbModeless      ' *spec fact 54*"
    Debug.Print "    2) UserForm1.StartWebView2_Full"
    Debug.Print "    3) Wv2Log.LogStart"
    Debug.Print "    4) Test_N4_Body"
    Debug.Print "    -> *(1) to (6) need an external host (httpbingo.org)*"
    Debug.Print ""
    Debug.Print "  [Usage]"
    Debug.Print "    p.NetDetailOn = True             ' *the place to store the body is the detail entry*"
    Debug.Print "    p.NetRespBodyOn = True           ' *the body is a separate switch (default OFF)*"
    Debug.Print "    Test_N1_Watch"
    Debug.Print "    ' ... operate by hand ..."
    Debug.Print "    Test_N1_Drain                    ' find a hit in the list"
    Debug.Print "    p.NetDetail 3                    ' request + response + body"
    Debug.Print "    p.NetDetailRespBody(3)           ' take only the body"
    Debug.Print "    p.NetDetailRespLine(3)           ' byte count/cut/kind/encoding/hex/arrived"
    Debug.Print "    p.NetRespBodyMaxBytes = 500000   ' body limit (default 64KB)"
    Debug.Print ""
    Debug.Print "  --- *Where the construction changed in N-4* ---"
    Debug.Print "  - *First asynchronous* GetContent takes a handler and returns immediately, and the body"
    Debug.Print "    arrives later. When judging, *wait until the body arrives* (design principle 120)."
    Debug.Print "  - *Unlike N-2, Seek(0) is not needed* It only reads a copy of the response and"
    Debug.Print "    does not touch the traffic itself."
    Debug.Print "  - *ResponseView is held until completion* (because the SDK does not say whether"
    Debug.Print "    it may be released). Always untie it with NetCaptureStop."
    Debug.Print ""
    Debug.Print "  --- *About compression* ---"
    Debug.Print "  Responses arrive as gzip / br / zstd. Whether GetContent returns the decoded body was"
    Debug.Print "  confirmed on the real machine ((2) of Test_N4_Body). Either way,"
    Debug.Print "  *content-encoding and the hex of the first 32 bytes are always kept*, so"
    Debug.Print "  when the body cannot be read, look there."
    Debug.Print ""
    Debug.Print "  --- What N-4 cannot do (later stages) ---"
    Debug.Print "    - turning it into PowerShell's Invoke-WebRequest  -> N-5"
End Sub


' ============================================================
' N4Wait (N-4, Private) - wait *until the body arrives*
'
'   *Neither "until the detail appears" nor "until the status is attached"*
'   The body arrives last in GetContent's completion handler, so wait for that.
'   In N-3 we mixed this up here and burned three real-machine runs (design principle 120).
' ============================================================
Private Function N4Wait(ByVal p As Wv2Pane, _
                        ByVal uriPart As String, _
                        ByVal timeoutSec As Single) As Boolean
    Dim t0 As Single
    Dim k  As Long
    Dim f  As Variant

    t0 = Timer
    Do
        DoEvents
        k = N2Find(p, uriPart)
        If k > 0 Then
            f = Split(p.NetDetailRespLine(k), vbTab)
            If UBound(f) >= 5 Then
                If f(5) = "True" Then          ' *arrived*
                    N4Wait = True
                    Exit Function
                End If
            End If
        End If
        If (Timer - t0) > timeoutSec Then Exit Do
    Loop
End Function
