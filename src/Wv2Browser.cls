VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Wv2Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''
' --- Wv2Browser.cls  Stage K-4c (quiet all tabs before release) ---
'
'   *Measured (K-4a/b)* If a page is running, COM responses do not come back.
'
'     Do nothing                     ... release 5.05 s (one COM round trip is 300-500ms)
'     send to about:blank (no wait)  ... release 0.38 s (round trip 8-16ms)
'     send to about:blank (wait 3 s) ... release 0.37 s  <- *waiting makes no difference*
'
'   *No need to wait* It takes effect the moment we send it. In total 5.05 s -> 1.98 s.
'   Sending takes 1.6 s because COM is still slow at that point.
'
'   *The TabBar / NavBar Panes are excluded* They are our own light HTML, and measured at 23-31ms.
'
'   *QuietOnShutdown = False turns it off* When isolating causes (Test_K4_Step),
'   quieting would ruin the measurement, so an escape hatch is left.
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Browser.cls  Stage K-4a (move the release-path logging to Wv2Log) ---
'
'   *The goal is measurement* We received a report that release is slow when the form is closed (about 0.5 s pause per item),
'   but this path was still Debug.Print and left nothing in the log, so
'   we could not tell where the time was going.
'
'   Wv2Log has millisecond timestamps, so *simply replacing lets us read the durations*.
'   Moreover, measuring twice with Wv2Log.LogEcho = False / True separates
'   *"COM release is slow" from "output to the Immediate window is slow"*.
'
'   *We do not fix anything at this stage. We only measure.* (because we guessed wrong 3 times in a row in D-7)
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Browser.cls  Stage K-2 (fix for the bug where a search crushes the settings tab) ---
'
'   *Symptom* If you enter a search term in the address bar while the settings tab is active,
'     the settings tab itself navigates to the search results and the settings screen disappears. Moreover, pressing the gear again
'     is judged as "a settings tab already exists" and no new settings tab opens.
'
'   *Root cause* IsSettings was set once at creation and never went back to False after navigation.
'     The flag did not reflect reality, so OpenSettingsTab's duplicate check
'     was grabbing the "former settings tab".
'     *It is not only an address-bar problem* The same happens with links, back, reload and JS-initiated navigation.
'     So a symptomatic fix that blocks NavigateActive does not cure it.
'
'   *Fix (A-2)* Change the check from a flag to a method of *asking the page itself*.
'     Embed window.__wv2Settings = true in the settings screen's HTML, and OpenSettingsTab fires
'     EvalSync (D-1) at each tab to ask directly "is this tab showing the settings screen right now".
'     It depends on neither the address representation nor the navigation path, and a tab
'     that returned to the settings screen via back is also recognized correctly.
'     *Put the authority of the check on the side that actually holds the state (design principle 91)*
'
'   *Fix (B-1)* Even with an accurate check, back navigation does not go through OpenSettingsTab,
'     so two settings views can end up side by side. So that the values do not disagree in that case,
'     *when the search engine changes, push the current value to all tabs showing the settings screen*
'     (just the global sync of design principle 60).
'     *To avoid the spec fact 27 deadlock, wait a beat with SafeTimer*
'     SetEngine is called via hostObjects (JS is blocked),
'     so EvalSync must not be fired on the spot.
'
'   *The IsSettings flag is kept, but is no longer used for the duplicate check*
'     It is only a record of the origin "created as a settings tab".
''''''''''''''''''''''''''''''''''
' --- Wv2Browser.cls  Stage 9.29 (custom search URL input in the settings screen + cleanup of a spare reservation) ---
'
'   Scope of 9.29 (the main part is on the Wv2SettingsBridge side. This is the UI and minor cleanup):
'
'   *(A) Add a custom input field to BuildSettingsHtml (issue 4, option a)*
'     Place it as an *independent section below* the card grid. The reason it is not a 5th card
'     is to keep EngineList() pure as the "preset definition"
'     (so the 9.18 design of "not duplicating engine names on the HTML side" is not broken).
'       - Input field + [Apply] button. While custom is in use, the section frame glows and an "in use" badge appears.
'       - While typing, Settings.PreviewUrlForTemplate gives a live preview *before applying*.
'         Encoding is not hand-rolled in JS but goes through VBA, so it is exactly the same as production.
'       - Invalid input (not starting with http:// / https://) is shown in red on the spot,
'         and applying it touches neither the Browser nor the saved file at all.
'     *Do not use <form>* Enter would submit and navigate the settings screen itself.
'       Handle it with the onclick / onkeydown of <input> + <button>.
'     *Use single quotes for JS strings and attribute values* A backslash + double quote cannot be written inside a VBA string
'       (VBA treats " as the end of the string). Write the static HTML part with VBA's "".
'
'   *(B) Stopped the useless reservation in EnsurePrewarmScheduled (open item 1 of 9.28b)*
'     *Cause established during implementation (at agreement time the plan was "do not set it if m_prewarmClaimed")*
'       m_adoptTimer_TimerCall runs in the order AdoptPrewarmedNow -> CreatePrewarmPaneNow.
'       CreatePaneCore in the latter spins DoEvents in WaitPaneReady, so
'       in the middle of that the AddTab of the spare-exhausted fallback interrupts *in a nested way*. From the point of view of
'       EnsurePrewarmScheduled at the end of this AddTab, m_prewarmPane is still
'       Nothing (the outer call has not assigned yet), so a reservation was set, and 1.5 seconds later the outer call
'       had finished warming the spare and "a spare tab already exists (skipping creation)" was logged.
'       -> m_prewarmClaimed is already False at this point, so it is not a usable condition.
'         The correct condition was *whether it is nested* = m_creatingPane > 0.
'     *Trade-off* If the outer spare creation fails, this AddTab does not set a reservation,
'       so the spare stays empty until the next AddTab / switch. An automatic retry
'       (firing RequestPrewarmAsync on failure) is not included, because in an environment where creation fails persistently
'       it would become an endless retry every 1.5 seconds.
'
' --- Wv2Browser.cls  Stage 9.28 (make prewarm delegation the default + 2 small spare-related fixes) ---
'
'   Scope of 9.28 (a round that collects 3 items from 9.27's "unresolved / next-stage open items". All small):
'
'   *(A) Change the default of NewWindowMode from 0 -> 2*
'     In 9.27 the main part (prewarm delegation via put_NewWindow) passed on the real machine, yet the default stayed at
'     0 (current behavior). Round 3 of 9.27 demonstrated that forgetting to switch the mode
'     takes POST links back to 405, so *make the state where POST opens correctly the plain default*.
'       - During PREWARM_DELAY_MS (1500ms) after startup there is no spare, and we fall back to the current behavior.
'         This is by design. During that time the first page is still loading and a link cannot be clicked, so
'         there is no real harm, and even if clicked it does not break; it only emits a *-marked log.
'       - We do not shorten the delay. 1500ms was placed deliberately as a safety margin because in 9.27 issue 1 "right after Init,
'         Controller creation is continuous, the busiest period, and the window of spec fact 20 opens"; cutting it
'         for the sake of making this the default would get the order backwards.
'       - Cost: one extra WebView2 stays resident at all times (one more msedgewebview2.exe).
'         In an environment where that is unacceptable, firing NewWindowMode = 0 after startup restores the old behavior.
'
'   *(B) Change m_creatingPane from Boolean -> Long (counter)*
'     In 9.27 round 3, nesting of CreatePaneCore actually occurred (while WaitPaneReady of the prewarm creation
'     was spinning DoEvents, m_newTabTimer_TimerCall called
'     AddTab -> CreatePaneCore). The result was fine, but with a Boolean
'     *the flag drops when the inner one finishes*, so the next one can run even though the outer one is still creating.
'     With increment/decrement, "> 0 if even one is being created" is preserved.
'
'   *(C) Aligned the early-return log of RequestPrewarmAsync with reality*
'     Whether the spare is "usable" or "already Claimed (awaiting adoption)" means completely different things, yet
'     9.27 reported both as "a spare tab already exists, so do nothing". Branch
'     and use separate wording. *Behavior is unchanged* (see the after-the-fact report below).
'
'   *Not touched* the put_NewWindow path itself / the prewarm creation, adoption and replenishment
'     logic / Deferral (policy of not using it stays) / other classes.
'     Wv2Pane.cls is unchanged from 9.27 (it is built to read the mode from Browser each time, so
'     changing only the default makes the new default take effect).
'
' --- Wv2Browser.cls  Stage 9.27 (main part: delegation to our own tab via put_NewWindow / prewarm approach) ---
'
'   Scope of 9.27:
'     Spec fact 23 (9.26b) established that "if the navigation itself is handed to the runtime, POST is fully preserved".
'     The remaining issue is UX -- putting that navigation on *our own tab* rather than a runtime-managed popup.
'     If we pass "the CoreWebView2 we prepared" to args.put_NewWindow (vtable index 4), the runtime
'     loads the original request onto that WebView.
'
'   *Adopted approach = prewarm (plan B / issues 1-7 agreed)*
'     Keep one empty Pane warmed at all times, and the moment NewWindowRequested arrives,
'     fire put_NewWindow(spare ViewPtr) and Exit immediately.
'     -> The greatest advantage is that *no Deferral (GetDeferral / Complete) is used at all*.
'       The Deferral approach "carries args / deferral COM references out of the event", so
'       it needs our own AddRef/Release management and walks next to spec fact 10 (Sentinel MEM_FREE).
'       With prewarm, all we do inside the event is 1-2 vtable calls, and there is zero waiting.
'       It touches neither spec fact 14 (SafeTimer deadlock) nor 20 (EBMode window).
'
'   *Added*
'     [Declarations]
'       - m_newWindowMode           : 0=current (GET interception) / 1=popup delegation / 2=prewarm delegation
'       - m_setHandledWithNewWindow : whether to use put_NewWindow together with put_Handled (default True)
'       - m_prewarmPane / m_prewarmClaimed : the spare tab itself and a "reserved" flag
'       - m_prewarmTimer / m_prewarmPending : one-shot SafeTimer for spare-tab creation
'       - m_adoptTimer / m_adoptPending / m_adoptUrl : one-shot SafeTimer for the adoption processing
'       - m_creatingPane            : Pane-creation-in-progress flag (re-entrancy guard for the prewarm TimerCall)
'     [Split-out = issue 2 option a]
'       - CreatePaneCore / CreatePaneCoreBody : the first half of the old AddTab (Pane creation to handler registration).
'         Called from both AddTab and the prewarm creation. AddTab became a thin function of just "CreatePaneCore + the 3 second-half steps
'         (m_tabs.Add / TabAdded / ActivateTab)".
'     [New methods]
'       - NewWindowMode (Get/Let)         : mode switch (3 values unified = issue 4 option b)
'       - PopupDelegateMode (Get/Let)     : *compatibility wrapper* delegates to NewWindowMode internally.
'                                           Test_9_26_Popup_On/_Off work as is.
'       - SetHandledWithNewWindow (Get/Let): the insurance switch of issue 5
'       - ClaimPrewarmViewPtr             : takes out the spare's ViewPtr "with a reservation"
'       - ReleasePrewarmClaim             : returns the reservation when put_NewWindow fails
'       - RequestPrewarmAsync             : reservation of spare-tab creation (one-shot Timer)
'       - RequestAdoptPrewarmedAsync      : reservation of adoption processing (one-shot Timer)
'       - Debug_PrewarmStatus             : state dump for verification
'
'   *Flow*
'     [Startup] Init succeeds -> RequestPrewarmAsync (after 1500ms) -> CreatePrewarmPaneNow
'            -> create the spare Pane and hide it in the back with Ctrl_PutIsVisible False (issue 3)
'     [Fire] Wv2Pane.View_OnNewWindowRequested (mode 2)
'            -> reserve the spare with ClaimPrewarmViewPtr -> put_NewWindow -> put_Handled
'            -> RequestAdoptPrewarmedAsync -> immediate Exit Sub (*wait for nothing inside the event*)
'     [Adopt] m_adoptTimer_TimerCall (after the handler completes)
'            -> AdoptPrewarmedNow : m_tabs.Add -> RaiseEvent TabAdded -> ActivateTab
'            -> CreatePrewarmPaneNow : replenish the next spare (issue 6: adoption first, replenishment after)
'
'   *When the spare is unusable* In any of: not created / not Ready / reserved / put_NewWindow failed,
'     Wv2Pane falls back to the current behavior (put_Handled + RequestNewTabAsync = GET interception).
'     A POST link returns to 405, but *that is better than breaking*. The fact of falling back is
'     always *logged with a star mark*, so it can be told apart on the real machine.
'
'   *Log wording changes* Because the first half of the old AddTab moved to CreatePaneCore,
'     lines such as "Wv2Browser.AddTab: Controller created OK" become
'     "Wv2Browser.CreatePaneCore: ...". Be careful when matching against past logs.
'
'   *Not touched* RequestNewTabAsync / m_newTabTimer / the settings-tab family / the 5 events /
'     the search-engine family / ActivateTab / CloseTab / MoveTab.
'
' --- Wv2Browser.cls  Stage 9.26 (toggle for the NewWindowRequested delegation probe) ---
'
'   Scope of 9.26 (a round that only does the "isolation" of the POST-link new-tab expansion returning 405):
'     *Background* In the workplace core system (Apache Tomcat/7.0.76), clicking "Appointment Management" after login
'       gives HTTP 405 (Request method 'GET' not supported). The cause is not an implementation bug but
'       a structural hole in the approach: the current Wv2Pane.View_OnNewWindowRequested
'         (1) takes out "only" the URL string with args.get_Uri,
'         (2) discards the runtime's own navigation with args.put_Handled(TRUE), and
'         (3) re-navigates in a new tab (= always GET),
'       so the original POST's method/body/Referer are lost at (2).
'
'     *What this stage does = only confirm the fact* Add a mode that does "not set" put_Handled,
'       and let the WebView2 runtime itself open it in a popup window (the 9.6a behavior).
'       If the business screen opens normally there, it is established that "delegating to the runtime preserves POST",
'       and we know the investment in the main part (the approach of passing another tab's CoreWebView2 via args.put_NewWindow)
'       will pay off. Conversely, if it is still 405, the cause is something other than the POST body (Referer/Cookie/session),
'       and we can change direction before starting on the main part.
'       *The costliest thing is to plunge into nested Deferral (a deadlock hot spot) without confirming this*
'
'     *This mode is not a permanent measure* The popup is a separate window outside our own tab UI, and
'       we do not hold its Controller, so we cannot control closing, arranging and so on. It is strictly
'       for verification (and as a stopgap when it is truly needed for work). Default is OFF.
'
'   *Added (managed collectively on the Browser side = agreed with tarboh)*
'     - Private m_popupDelegateMode As Boolean        (declaration section at the top of the module)
'     - Public Property Get/Let PopupDelegateMode     (default False = current behavior unchanged)
'     Since the flag is on the Browser rather than per Pane, it is enough for every tab's Wv2Pane to
'     just look at m_browser.PopupDelegateMode and they all stay in step. Switching is also one time regardless of the tab count.
'
'   *How to call it from the Immediate window* It uses the existing UserForm1.CurrentBrowser (the 9.17 scaffold),
'     so UserForm1.frm remains unchanged. The actual operations are wrapped in
'     Test_9_26_Popup_On / _Off / _Help of Wv2Tests.bas.
'
'   *Not touched* AddTab / RequestNewTabAsync / SafeTimer delayed execution (design principle 63) /
'     the 5 events / the search-engine family. No design principles or spec facts added either (this stage is observation only).
'
' --- Wv2Browser.cls  Stage 9.20 (persistence of search-engine choice, restore side) ---
'
'   Additions in 9.20 (start with the previously chosen search engine):
'     *Restore the saved engine in Class_Initialize (option i)*
'       The moment the Browser is born, a temporary Wv2SettingsBridge is created with New
'       and LoadEngineTemplate() is called; if there is a template string of the previously saved
'       search engine, m_searchTemplate is overwritten with it.
'       If nothing is saved, the default (Google) stays. Whether launched via UserForm or directly,
'       it is settled when the Browser is born, so every path consistently starts from the previous value.
'
'     *Name <-> template resolution stays confined to the Bridge*
'       Browser does not know the vocabulary of engine names ("bing" etc.). The Bridge's
'       LoadEngineTemplate() returns "a template string or empty", so
'       Browser just puts it into m_searchTemplate. The save side (SaveEngineName
'       right after SetEngine) is also confined to the Bridge (9.20 consolidates into the bridge).
'
'     *Direction of dependency* Class_Initialize does New Wv2SettingsBridge, but this is the same practice
'       already used inside OpenSettingsTab (only a type reference, so
'       no circular reference). LoadEngineTemplate is a side-effect-free read, so
'       BindBrowser is also unnecessary (it just reads the file without a Browser reference).
'
' --- Wv2Browser.cls  Stage 9.19 ---
'
'   Additions in 9.19 (prevent duplicate settings tabs + default display of the settings-screen preview):
'     *Make the default of the settings-screen hover preview the "currently selected engine"*
'       Added showCurrentPreview() to the JS of BuildSettingsHtml, and in (a) right after opening the settings screen
'       (not yet hovered) and (b) when the cursor leaves every card,
'       it shows the preview for the currently selected engine. Previously the preview of the last
'       hovered card stayed on.
'       *The return check is done with per-card mouseenter/mouseleave + a hoverId flag*
'       With mouseleave on the whole grid container, the gaps between cards and moves to another card
'       are still inside the container and it does not fire, which was counterintuitive (it did not return until
'       fully leaving the grid area). Set hoverId on each card, and on mouseleave, via
'       setTimeout(0), "if no card is hovered any more, return to the selected one", which
'       avoids flicker when moving between cards while returning the moment the cursor leaves a card (the next card's mouseenter sets
'       hoverId again first in the same tick, so
'       the return processing is skipped).
'
'     *Look for an existing settings tab at the start of OpenSettingsTab*
'       In 9.18 a new settings tab was created every time [gear] was pressed, and mashing it produced
'       many settings tabs. In 9.19, at the start of OpenSettingsTab m_tabs is scanned
'       from the front, and if a tab with Pane.IsSettings = True already exists, it just
'       does ActivateTab on it and Exit Sub (no new creation). If none, it creates as before
'       and sets IsSettings = True on the created Pane.
'
'     *Identification uses the Wv2Pane.IsSettings flag (plan A; see the 9.19 development notes for the design)*
'       Not found in the scan = there is no settings tab / it was closed, so the liveness check is
'       merged with the scan. There is no need to hold a separate settings Pane reference on the Browser side
'       and watch for dangling. m_settingsBridge still holds one bridge, that of the last
'       opened settings tab (existing tabs each keep the reference from attach time).
'
' --- Wv2Browser.cls  Stage 9.18 ---
'
'   Additions in 9.18 (settings tab + search-engine UI via AddHostObjectToScript):
'     *New route for opening the settings tab* (part 3)
'       TabBar [gear] button -> OnPaneWebMessage "settings" -> RequestSettingsTabAsync.
'       Creating the settings tab also goes through AddTab -> WaitPaneReady internally, so
'       calling it synchronously from inside the WebMessage handler deadlocks, as with newtab (spec fact 14).
'       It is deferred with a dedicated one-shot SafeTimer (m_settingsTimer).
'       - RequestSettingsTabAsync(): set the pending flag and start a New SafeTimer as a one-shot
'       - m_settingsTimer_TimerCall: call OpenSettingsTab after the handler completes
'       - OpenSettingsTab(): AddTab -> bridge attach -> NavigateToString -> ActivateTab
'       - BuildSettingsHtml(): builds the HTML of the search-engine selection card UI
'
'     *Expose Wv2SettingsBridge to JS with AddHostObjectToScript* (part 2-b)
'       Attach Wv2SettingsBridge to the settings tab's Pane so the settings screen's JS can call VBA directly,
'       like hostObjects.sync.Settings.SetEngine("bing").
'       The reason for going through the bridge instead of handing over the whole Wv2Browser is to narrow the "surface exposed to JS" to the search
'       engine only (so CloseTab etc. cannot be hit from the settings screen).
'       Resolution of preset name <-> template was consolidated into the bridge (the SE_* constants and name resolution
'       that UserForm1 held up to 9.17 were moved here. DRY).
'
'     *Real-machine verification takes attach before NavigateToString (order A) as the first candidate*
'       The hypothesis is that if registration comes first, the loaded JS can use hostObjects from the start.
'       If that fails, fall back to the ready handshake (order B).
'
'     *Lifetime* The attached bridge is held in m_settingsBridge as one instance and is
'       referenced by WebView2 while the settings tab is alive. It is released in Shutdown.
'
'   Additions in 9.16 (search-engine switching):
'     *Make the search URL template a setting (option 1-A + option 2-A)*
'       The Google-fixed search URL hard-coded in BuildSearchUrl in 9.14 is held in the
'       module variable m_searchTemplate so it can be swapped. Default stays
'       Google (with no setting, behavior is exactly the same as 9.14).
'
'     *The template is a "query-prefix type" string* (option 1-A)
'       A prefix template like "https://www.google.com/search?q=" to which the (encoded) search term is
'       simply concatenated at the end. Major engines such as Google / Bing / DuckDuckGo / Yahoo can
'       all be expressed in this form. An Enum + lookup table (option 1-B) adds state,
'       so we do not take it and hold the template string directly (YAGNI). If a rare engine whose "{q}" is not at the end is needed
'       in future, only then do we extend to a placeholder approach.
'
'     *The only setting entry is the main Property, no UI (option 2-A)*
'       It is read and written via Public Property Get/Let SearchEngine. If an empty string is passed it
'       falls back to the default Google template; only a light guard (of the URL
'       validity checking. The template is a trusted value entered by the developer). Settings from the UI
'       (combo box etc.) are split out to another stage (to keep this stage small
'       in granularity so it can be verified on its own on the real machine).
'
'     *Added Debug_SearchEngine (logic verification entry)*
'       A thin public wrapper that returns the current m_searchTemplate. Wv2Thunks.Test_9_16_SearchEngine
'       uses it to verify "when the engine is switched, does the BuildSearchUrl result follow" without
'       launching WebView2 (same positioning as Debug_NormalizeUrl).
'
'   Cleanup items of 9.16 (done in the same stage):
'     *Split the pure URL functions out to Wv2Url.bas (option 2-A / option 3-A)*
'       IsAlphaOnly / UrlEncodeComponent / Utf8Bytes (pure data transformations whose output
'       is determined by input alone) were moved to Wv2Url.bas. Together with Wv2Json.bas it forms a pair of
'       "stateless transformation utilities". NormalizeUrl / BuildSearchUrl depend on
'       the instance state of the search-engine setting (m_searchTemplate), i.e.
'       "browser behavior (policy)", so they stay in Wv2Browser and internally call
'       Wv2Url.IsAlphaOnly / Wv2Url.UrlEncodeComponent (separation of policy and
'       mechanism). Logic is unchanged; only the location and the call qualifier changed.
'
'     *Consolidated the feature-verification Subs into Wv2Tests.bas (option 1-A)*
'       Test_9_14/15/16 etc. (pure-logic verification that only hits Wv2Browser's Debug_*) were
'       moved from Wv2Thunks.bas to Wv2Tests.bas. Wv2Thunks is purified into the thunk foundation.
'       On the Wv2Browser side the verification entries (Debug_*) remain public as they are (no change).
'
' --- Wv2Browser.cls  Stage 9.15 ---
'
'   Additions in 9.15 (drag reordering of tabs):
'     *Added MoveTab(fromIndex, toIndex)* (option A-1)
'       Moves an element of m_tabs (1-based Collection) from fromIndex to toIndex.
'       A Collection has no element-move API, so it is rearranged by Remove -> Add at the target position (before/
'       after specified). The tab body (Pane) stays as is; only the order changes.
'
'     *Active tracking is done by object reference (option A-1)*
'       Reordering moves the active tab's index, so the active Pane is grabbed
'       before reordering, and after reordering m_tabs is scanned linearly to reverse-look-up the new index
'       (the same reference-match pattern as NotifyTitleChanged). The visible state does not change, but
'       ActivateTab is called at the end to set it right (same style as CloseTab).
'
'     *Sync reuses the existing path; no new event is added (option A-2b)*
'       Calling ActivateTab(new active) after reordering raises RaiseEvent ActiveChanged,
'       which runs PushSyncToJs (global sync, design principle 60) of TabBar/NavBar.
'       Global sync has the philosophy of "when state changes, resend the whole tab state", so
'       even if the active tab itself does not change, the reordering result is reflected in JS as is.
'       -> Public Event can stay at 5 kinds (no TabsReordered is created).
'
'     *Added Debug_TabOrderTitles (logic verification entry)*
'       A thin public wrapper that returns the titles in the current tab order (or (t<index>) if none)
'       joined with "|". Wv2Thunks.Test_9_15_Reorder uses it to verify the reordering result of MoveTab
'       without launching WebView2 (same positioning as Debug_NormalizeUrl).
'
' --- Wv2Browser.cls  Stage 9.14 ---
'
'   Additions in 9.14 (search-term fallback, default Google):
'     *Typing "words to search" in the address bar sends it to Google search*
'       Collects the open item that was put on hold in 9.12a as "out of scope / next stage" (the comment in NormalizeUrl).
'       Input with no dot / with whitespace is regarded as a search term rather than a URL, and is
'       converted to https://www.google.com/search?q=... and navigated to.
'
'     *Extended NormalizeUrl to three branches (option A)*
'       (1) Has a scheme (contains "://" / "about:" etc.)      -> as is
'       (2) Contains a dot and contains no whitespace          -> prefix https:// (URL)
'       (3) Otherwise (contains whitespace / no dot)           -> search term -> BuildSearchUrl
'       Nearly the same feel as the Chrome/Edge address bar. "localhost:8080" falls into (1) through the existing colon
'       branch (IsAlphaOnly), so it is not searched by mistake.
'
'     *URL encoding is implemented without an external COM library (option E-1)*
'       Percent-encoding requires conversion to a UTF-8 byte sequence (Japanese search term ->
'       %E6%97%A5%E6%9C%AC etc.). This project's WebView2 control uses no external COM at all by policy, but
'       UTF-8 conversion is a mere data-transformation utility unrelated to the layer of
'       "machine-code thunks, vtables, SAFEARRAY", and ADODB.Stream is already adopted for HTML output
'       (so it is not philosophically gray). Write the string to ADODB.Stream with
'       Charset="utf-8", read the bytes back, and escape everything other than RFC3986 unreserved characters
'       to %XX.
'
'     *Location = Private helpers of Wv2Browser (option P-1)*
'       The search fallback is the Browser's responsibility (same place as NormalizeUrl), so
'       UrlEncodeComponent / Utf8Bytes / BuildSearchUrl are held Private inside Wv2Browser.
'       If URL handling grows in future, extract them to Wv2Url.bas at a cleanup stage (the same rhythm as the later
'       extraction of JSON in 9.13). Wv2Json.bas is declared "JSON only", so do not mix them in.
'
' --- Wv2Browser.cls  Stage 9.12a ---
'
'   Additions in 9.12a (foundation for the address bar / nav buttons):
'     *New event TabUrlChanged(index, url, canGoBack, canGoForward)*
'       Exactly the same loose-coupling pattern as TabTitleChanged of 9.11a. The Browser does not know
'       the nav bar exists and only fires state changes. The nav bar subscribes with WithEvents.
'     - NotifyUrlChanged(senderPane, url, canBack, canFwd): called from a Pane, it
'       reverse-looks-up the index by object-reference match and does RaiseEvent TabUrlChanged.
'       (same shape as NotifyTitleChanged)
'
'     *Promote HistoryChanged to a standard handler (option 2-A-1)*
'       Previously HistoryChanged was "for SPA" and registered only in EnableSpaHandlers
'       (9.10d-2b, 9.10c-2 issue 10). However, the enabled state of the back/forward buttons needs to be tracked on all tabs, not
'       only SPA, so it is moved to RegisterStandardHandlers and
'       registered automatically on all tabs. -> Standard handlers go from 4 kinds to 5.
'       EnableSpaHandlers was changed to handle only DOMContentLoaded (SPA-specific
'       events remain there). To avoid double registration, the registration of
'       HistoryChanged was removed from EnableSpaHandlers.
'
'     *Delegating operations to the active tab (public API for the nav bar)*
'       The nav bar wants to send operations to the "active tab", but letting it touch the Pane directly
'       breaks loose coupling. We provide thin methods that the Browser proxies.
'         NavigateActive(url) / GoBackActive / GoForwardActive /
'         ReloadActive / StopActive
'         GetActiveUrl / GetActiveCanGoBack / GetActiveCanGoForward
'       *Design principle 63* None of them involves "waiting for completion", so it is
'         safe to call them synchronously from inside an event handler (no deferred execution like newtab in 9.11d is needed).
'
'     *URL scheme completion (option 3-A)* If there is no scheme, NavigateActive prefixes
'       https:// ("example.com" -> "https://example.com"). The search-term fallback is
'       out of scope (next stage).
'
' --- Wv2Browser.cls  Stage 9.11c ---
'
'   Additions in 9.11c (match the initial Bounds of a new tab to the content rectangle, option A):
'     *Problem* AddTab made the initial Bounds of a new tab GetClientRect(m_hwndHost) = the whole
'       frame. It did not subtract the tab-bar band (top 36px), so a tab
'       added after startup overlapped the tab bar (not fixed until UserForm_Resize ran).
'     *Structure of the cause* The Browser does not know the tab bar exists (loose coupling), so the actual
'       content rectangle (the area minus the tab bar) is known only to the upper level (UserForm).
'     *Remedy (option A)* While keeping the upper level as the source of truth for layout, the Browser caches "the content rectangle the
'       upper level applied last" and applies it consistently to new tabs.
'       - New fields: m_lastContentRect As RECT / m_hasContentRect As Boolean
'       - ResizeAll: saves the passed rc to m_lastContentRect and sets m_hasContentRect=True.
'       - AddTab: if m_hasContentRect is True, uses that rectangle as the initial Bounds.
'         When unset (ResizeAll not yet run), falls back to GetClientRect (whole area) as before.
'       *Loose coupling is unchanged* The Browser still does not know the tab bar; it just
'       remembers "the area the upper level specified". UserForm is unchanged (UserForm_Resize at startup already
'       calls ResizeAll, so subsequent AddTabs are automatically born with the correct rectangle).
'
' --- Wv2Browser.cls  Stage 9.11b ---
'
'   Additions in 9.11b (an entry that lends the Environment to Wv2TabBar):
'     - Added Property Get Environment() As Wv2Environment. So that when Wv2TabBar
'       creates the pane for drawing the tab bar, it shares the same Environment as the Browser
'       (one Environment is shared by all panes).
'       *Loose coupling* The Browser only "lends the Environment" and does not know
'       what the borrower uses it for (tab bar or something else).
'
'   Additions in 9.11a (event entries for a loosely coupled tab bar + CloseTab, option A):
'     *Loose coupling* The Browser knows nothing about the tab bar implementation (WebView2-based / MsForms-based / none).
'       It only notifies the outside of state changes with Public Event (RaiseEvent passes harmlessly through
'       even with zero listeners). Receivers receive with WithEvents m_browser.
'     - New events: TabAdded(index) / TabClosed(index) / ActiveChanged(index) /
'       TabTitleChanged(index, title)
'     - AddTab: RaiseEvent TabAdded right after adding to the Collection (before ActivateTab).
'       ActivateTab is called right after, so when a tab is added two fire, in the
'       order TabAdded -> ActiveChanged.
'     - ActivateTab: RaiseEvent ActiveChanged on success.
'     - CloseTab(index): new. ClearBrowserRef (design principle 58) -> Ctrl_Close (vtable 24,
'       double Close is harmless) -> m_tabs.Remove -> active correction -> RaiseEvent TabClosed.
'       Active correction has 4 branches (0 left = zero tabs allowed / the closed one was before the active /
'       the active itself = prefer the right neighbor, else the left / after the active). Visibility repair is
'       consolidated in ActivateTab (design principle 57). The minimum-tab-count policy is left to the upper level.
'     - NotifyTitleChanged(senderPane, title): new. Receives a title change from a
'       Pane, reverse-looks-up the index by object-reference match and does RaiseEvent TabTitleChanged.
'
'   Additions in 9.10e (support for making NewWindowRequested open a new tab, option Y + option beta Timer):
'     *Deadlock avoidance (option beta)* Calling AddTabWithUrl directly from Wv2Pane inside the NewWindowRequested handler
'       deadlocks (spec fact 14), so it is executed after the handler completes by
'       RequestNewTabAsync + SafeTimer (one-shot).
'     - New fields: m_newTabTimer (WithEvents SafeTimer) / m_pendingUrls (pending queue)
'     - RequestNewTabAsync(url): pushes onto the pending queue and starts a New SafeTimer as a one-shot
'       (SafeTimer.StartTimer ignores a restart with the same interval and m_uElapse remains after stopping,
'        so we do New for every request to make sure it starts)
'     - m_newTabTimer_TimerCall: outside the handler, processes the queue in order with AddTabWithUrl
'     - AddTab runs pane.SetBrowser Me and gives each Pane a back-reference to itself (the Browser)
'       (so that Wv2Pane can call AddTabWithUrl
'       at NewWindowRequested time).
'     - Shutdown: breaks the circular reference (Browser.m_tabs <-> Pane.m_browser). Calls each Pane's
'       ClearBrowserRef and then releases. Called explicitly from UserForm_Terminate
'       (design principle 58). The implementation cost on the UserForm side is only one Shutdown call line.
'
'   Additions in 9.10d-2b (SPA helper, option 2-c + option F-1):
'     - EnableSpaHandlers(pane): a public helper that registers the SPA events (HistoryChanged + DOMContentLoaded)
'       on a pane. Internally it just calls pane.View_AddHistoryChanged +
'       pane.View2_AddDOMContentLoaded. SPA-only events not included in the standard 4 handlers
'       (RegisterStandardHandlers) are gathered here
'       (9.10c-2 issue 10: SPA ones are registered explicitly). It can also be retrofitted onto existing tabs.
'     - AddTabWithUrlForSpa(hostName, folderPath, fileName): creates an SPA tab in one go.
'       In the order AddTab -> View3_SetVirtualHostNameToFolderMapping (before Navigate, after getting View) ->
'       Navigate(https://<host>/<file>) -> EnableSpaHandlers.
'       *Spec fact 13* State the Navigate target down to the file name (the root "/" resolves to a
'       directory and gives ERR_ACCESS_DENIED).
'
'   Additions in 9.10c-2:
'     - RegisterStandardHandlers(pane): automatically registers the standard 4 handlers (NavigationStarting/
'       Completed/DocumentTitleChanged/NewWindowRequested) inside AddTab
'       (option 10-A).
'
'   Additions in 9.10c-1:
'     - ActivateTab(index): realizes tab visibility control (only the active tab is visible)
'       with pane.Ctrl_PutIsVisible. Visibility control is consolidated in this one place (design principle 57).
'       AddTab delegates to it.
'
'   A umbrella class that bundles multiple tabs. It holds exactly one Wv2Environment and
'   creates and manages N Wv2Panes (= tabs) from it.
'   The core of the concept in which the UserForm fields will in future be consolidated from m_env + m_pane into a single m_browser
'   (the UserForm connection itself is 9.10c).
'
'   Design policy (decisions of stage 9.10):
'     - Class name Wv2Browser (unified Wv2 prefix, design principle 38)
'     - Wv2Pane is preserved (option D-1). This class only bundles them with a Collection, and
'       each tab's View_/Ctrl_/Settings_ are reached via ActivePane.
'     - One Environment is shared by all tabs (option E-1).
'     - Waiting for asynchronous completion is confined inside AddTab / Init (issue 1, plan a).
'       -> The Do While ... DoEvents wait logic can be removed from the UserForm side.
'
'   Scope of 9.10b (skeleton only):
'     Build  : Init / AddTab / AddTabWithUrl / AddTabWithHtml / TabCount /
'             ActivePane / Class_Terminate
'     Not built: ActivateTab (display switching) / RemoveTab / UserForm connection /
'               mouse-resize linkage (all 9.10c or later)
'     -> All tabs are created overlapping on the same m_hwndHost (visually the last
'       one is in front). The purpose of 9.10b is only to confirm that "each Pane reaches Controller/View"
'       (issue 2).
'
'   Release order (design principle 43: from child to parent):
'     Class_Terminate releases all Panes and then releases the Environment.
'     Each Wv2Pane's Class_Terminate releases in the order Settings->View2->View->Ctrl,
'     so this class keeps the order at the granularity of "Panes -> Environment".
''''''''''''''''''''''''''''''''''

Option Explicit

' ============================================================
' Browser state
' ============================================================
Public Enum BrowserState
    Bs_Created = 0     ' right after New, before Init
    Bs_EnvReady = 1    ' Environment creation complete, tabs can be added
    Bs_Failed = 2      ' Environment creation failed
End Enum

' --- Timeout seconds for waiting on asynchronous completion (same value as the existing StartWebView2) ---
Private Const READY_TIMEOUT As Single = 10#

' --- Added in stage 9.16: default search URL template (prefix type, the search term is concatenated at the end) ---
Private Const DEFAULT_SEARCH_TEMPLATE As String = "https://www.google.com/search?q="

' ============================================================
' Fields
' ============================================================
Private m_env      As Wv2Environment   ' one, shared by all tabs
Private m_tabs     As Collection       ' holds Wv2Panes (1-based)
Private m_dbgTagMap As Collection            ' 9.15 for verification: ObjPtr(key) -> tag ("A","B",..). Nothing when unused in production
Private m_hwndHost As LongPtr          ' parent frame HWND (common to all tabs)
Private m_state    As BrowserState

' --- Added in stage 9.16: search URL template of the search engine (option 1-A) ---
'   A prefix template to which BuildSearchUrl concatenates the "encoded search term" at the end.
'   Default is Google. Set in Class_Initialize and swapped via Property SearchEngine.
'   Examples: "https://www.google.com/search?q="  /  "https://www.bing.com/search?q="
Private m_searchTemplate As String

' Added in stage 9.10c-1: index of the active tab (1-based, 0 = no active).
'   The only state ActivateTab updates. ActivePane refers to it.
Private m_activeIndex As Long

' --- Added in stage 9.11c: content rectangle cache (option A) ---
'   Remembers the rectangle the upper level (UserForm) last applied in ResizeAll and reuses it as
'   the initial Bounds of AddTab. This way tabs added after startup are also placed in
'   the correct area that avoids the tab-bar band. While m_hasContentRect is False (ResizeAll not yet run)
'   AddTab falls back to GetClientRect (whole area).
'   *The Browser does not know the tab bar* What goes in here is "the area the upper level specified",
'   and the Browser itself does not compute the tab-bar height (loose coupling is maintained).
Private m_lastContentRect As RECT
Private m_hasContentRect  As Boolean

' Added in stage 9.10e (Timer deferred execution):
'   A one-shot SafeTimer and a pending URL queue for deferring the new-tab creation of
'   NewWindowRequested until after the handler completes. TimerCall is received with WithEvents.
Private WithEvents m_newTabTimer As SafeTimer
Attribute m_newTabTimer.VB_VarHelpID = -1
Private m_pendingUrls As Collection

' Added in stage 9.18 (deferred creation of the settings tab):
'   RequestSettingsTabAsync is called from the "settings" case of OnPaneWebMessage, triggered by the
'   TabBar [gear] button. The process of opening the settings tab also goes internally through
'   AddTab -> WaitPaneReady (waiting for Controller creation to complete), so calling it synchronously from inside the WebMessage
'   handler falls into the deadlock of spec fact 14 (same structure as newtab).
'   So it is deferred with a one-shot SafeTimer of the same type as newtab.
'   *Do not piggyback on the newtab queue (m_pendingUrls)* That one is for URL strings
'     only. The settings tab needs just a "an open request arrived flag", so it has its own
'     (more readable than mixing kinds, and straightforward for the one-shot of design principle 63).
'   *m_settingsBridge* Holds the Wv2SettingsBridge attached to the settings tab.
'     It is referenced by WebView2 while the settings tab is alive, so the Browser also holds one.
' K-4c: whether to send all tabs to about:blank before release (default True)
Private m_quietOnShutdown As Boolean

Private WithEvents m_settingsTimer As SafeTimer
Attribute m_settingsTimer.VB_VarHelpID = -1

' --- K-2: global sync of the settings view (design principle 60) ---
'   When the search engine changes, push the current value to all tabs showing the settings screen.
'   *The reason for inserting a SafeTimer is to avoid the deadlock of spec fact 27*
Private WithEvents m_syncTimer As SafeTimer
Attribute m_syncTimer.VB_VarHelpID = -1
Private m_syncPending As Boolean
Private m_settingsPending As Boolean
Private m_settingsBridge As Wv2SettingsBridge

' --- Added in stage 9.26 -> removed in 9.27: m_popupDelegateMode (Boolean) ---
'   The delegation toggle of 9.26 was a single Boolean, but in 9.27 there became three modes, so it was
'   merged into m_newWindowMode (Long) (issue 4 option b). Because lining up two Booleans would create
'   4 combinations, some meaningless.
'   *The PopupDelegateMode property itself remains as a compatibility wrapper* so
'     Test_9_26_Popup_On / _Off still work as is in 9.27.

' --- Added in stage 9.27: the 3-value mode of NewWindowRequested and the prewarm mechanism ---
'   *Mode values (correspond to the Private Const on the Wv2Pane side)*
'     0 = NWM_LEGACY  : current behavior. put_Handled(TRUE) + RequestNewTabAsync makes our own new tab.
'                       Only the URL is carried over, so POST turns into GET (= the cause of 405).
'     1 = NWM_POPUP   : the probe of 9.26. Does not set put_Handled and leaves it all to the runtime.
'                       POST is preserved but it becomes a popup and does not appear on the tab bar.
'     2 = NWM_PREWARM : the favorite of 9.27. Passes the spare tab's CoreWebView2 to put_NewWindow.
'                       POST is preserved and it also fits in our own tab.
'   *m_setHandledWithNewWindow (insurance of issue 5)*
'     Whether to also set put_Handled(TRUE) when put_NewWindow is set.
'     The official sample sets both, so the default is True. It can be set to False for isolation if
'     "stays white" or "opens twice" should ever happen.
Private m_newWindowMode As Long
Private m_setHandledWithNewWindow As Boolean

'   *Prewarm (spare tab)*
'     m_prewarmPane    : the warmed empty Pane (not put in m_tabs = does not appear on the tab bar)
'     m_prewarmClaimed : indicates already handed to put_NewWindow (= awaiting adoption). Handing it over twice
'                        would put two navigations on the same WebView, so Claim restricts it to once.
Private m_prewarmPane As Wv2Pane
Private m_prewarmClaimed As Boolean

'   *One-shot Timer for spare-tab creation (issue 1 option b)*
'     Reserved right after Init succeeds, and fired when the startup sequence (Env -> first tab -> TabBar/NavBar)
'     has settled. A delay so as not to insert the creation of the 4th Controller into the busiest stretch of
'     startup (avoids the window of spec fact 20).
Private WithEvents m_prewarmTimer As SafeTimer
Attribute m_prewarmTimer.VB_VarHelpID = -1
Private m_prewarmPending As Boolean

'   *One-shot Timer for adoption processing*
'     Doing everything up to m_tabs.Add / ActivateTab inside the NewWindowRequested handler
'     would run heavy processing inside the event. Same type as RequestNewTabAsync, it is processed
'     in the order adoption -> replenishment in the TimerCall after the handler completes.
Private WithEvents m_adoptTimer As SafeTimer
Attribute m_adoptTimer.VB_VarHelpID = -1
Private m_adoptPending As Boolean
Private m_adoptUrl As String

'   *Pane-creation-in-progress counter (re-entrancy guard)*
'     CreatePaneCore spins DoEvents in its internal WaitPaneReady, so in the middle of that
'     m_prewarmTimer_TimerCall can interrupt. Watches so as not to start another creation during creation.
'     *Changed from Boolean to Long in 9.28* Nesting actually occurred in 9.27 round 3.
'     With a Boolean the flag drops when the inner creation finishes, so even if the outer one is still creating
'     the next creation can run. With increment/decrement, "> 0 if even one is being created" is preserved.
Private m_creatingPane As Long

'   Wait time (ms) from reserving the spare-tab creation until it is actually created.
'   The purpose is to wait for the startup sequence to finish, so if too short it gets skipped by the re-entrancy guard,
'   and if too long it misses the first POST link. 1.5 seconds is the initial value.
Private Const PREWARM_DELAY_MS As Long = 1500

' --- Added in stage 9.11a: event entries for a loosely coupled tab bar (option A) ---
'   The Browser knows nothing about "who is listening". Whether the tab bar is WebView2-based,
'   MsForms-based or unnecessary, this entry is the same. RaiseEvent passes harmlessly through
'   even with zero listeners. Receivers receive with Private WithEvents m_browser As Wv2Browser.
'   - TabAdded(index)              : a tab was added (right after adding to the Collection)
'   - TabClosed(index)             : a tab was closed (right after removal from the Collection)
'   - ActiveChanged(index)         : the active tab changed (on ActivateTab success)
'   - TabTitleChanged(index,title) : a tab's title changed (from a Pane via NotifyTitleChanged)
Public Event TabAdded(ByVal index As Long)
Public Event TabClosed(ByVal index As Long)
Public Event ActiveChanged(ByVal index As Long)
Public Event TabTitleChanged(ByVal index As Long, ByVal title As String)

' --- 9.12a: notify URL state changes (for the address bar / back-forward buttons) ---
'   Fired from a Pane's NavigationCompleted / HistoryChanged via NotifyUrlChanged.
'   *Loose coupling* The Browser does not know the nav bar exists. Even with zero listeners RaiseEvent
'   passes harmlessly through.
Public Event TabUrlChanged(ByVal index As Long, ByVal url As String, _
                           ByVal canGoBack As Boolean, ByVal canGoForward As Boolean)


' ============================================================
' Class_Initialize
' ============================================================
Private Sub Class_Initialize()
    m_quietOnShutdown = True             ' K-4c (send to about:blank before release)
    Set m_tabs = New Collection
    Set m_env = Nothing
    m_hwndHost = 0
    m_state = Bs_Created
    m_activeIndex = 0
    Set m_pendingUrls = New Collection   ' 9.10e (NewWindowRequested pending URL queue)
    Set m_newTabTimer = Nothing          ' 9.10e (New per request)
    m_hasContentRect = False             ' 9.11c (content rectangle not yet cached)
    m_searchTemplate = DEFAULT_SEARCH_TEMPLATE  ' 9.16 (default Google. Can be swapped via Property)
    Set m_settingsTimer = Nothing        ' 9.18 (New per request)
    m_settingsPending = False            ' 9.18 (pending flag for settings-tab creation)
    Set m_settingsBridge = Nothing       ' 9.18 (the bridge to attach)

    ' 9.27: initialization of the NewWindowRequested mode and the prewarm mechanism
    '   *Default changed from 0 to 2 in 9.28* Since the main part passed on the real machine, the state where a POST link
    '     opens correctly in our own tab is made the "plain state". Only during PREWARM_DELAY_MS after startup
    '     is the spare not yet created and we fall back to the current behavior, but during that time the first
    '     page is still loading and a link cannot be clicked, so there is no real harm.
    '     To get the old behavior, fire NewWindowMode = 0 after startup.
    m_newWindowMode = 2                  ' default is prewarm delegation (NWM_PREWARM)
    m_setHandledWithNewWindow = True     ' use put_NewWindow together with put_Handled (issue 5)
    Set m_prewarmPane = Nothing
    m_prewarmClaimed = False
    Set m_prewarmTimer = Nothing         ' New per request
    m_prewarmPending = False
    Set m_adoptTimer = Nothing           ' New per request
    m_adoptPending = False
    m_adoptUrl = ""
    m_creatingPane = 0                   ' 9.28: Boolean -> Long (nesting counter)

    ' 9.20: restore the saved search engine (option i).
    '   New a temporary bridge and get the saved template with LoadEngineTemplate().
    '   If non-empty, overwrite m_searchTemplate; if empty, the default (Google) stays.
    '   The bridge holds no Browser reference and only reads the file, so BindBrowser is unnecessary.
    Dim initBridge As Wv2SettingsBridge
    Set initBridge = New Wv2SettingsBridge
    Dim savedTemplate As String
    savedTemplate = initBridge.LoadEngineTemplate()
    If Len(savedTemplate) > 0 Then
        m_searchTemplate = savedTemplate
        Debug.Print "Wv2Browser.Class_Initialize: restored saved engine = " & savedTemplate
    End If
    Set initBridge = Nothing
End Sub


' ============================================================
' Init ? create one Environment and wait until Ready
'   hwndHost   : HWND of the parent frame on which all tabs sit
'   userDataPath: WebView2's userDataFolder
'   Return value: True if the Environment becomes Ready
' ============================================================
Public Function Init( _
    ByVal hwndHost As LongPtr, _
    ByVal userDataPath As String) As Boolean

    If m_state <> Bs_Created Then
        Debug.Print "Wv2Browser.Init: called twice (state=" & m_state & ")"
        Init = (m_state = Bs_EnvReady)
        Exit Function
    End If

    m_hwndHost = hwndHost

    ' --- Create the Environment ---
    Set m_env = New Wv2Environment
    m_env.Init userDataPath

    ' --- Wait for Ready (the wait loop is confined inside this class = issue 1 a) ---
    If Not WaitEnvReady() Then
        Debug.Print "Wv2Browser.Init: Environment creation failed (state=" & _
                    m_env.state & ", lastError=&H" & Hex(m_env.LastError) & ")"
        m_state = Bs_Failed
        Exit Function
    End If

    Debug.Print "Wv2Browser.Init: Environment created OK (pEnv=" & _
                m_env.EnvironmentPtr & ")"
    m_state = Bs_EnvReady

    ' --- 9.27: reserve the creation of the spare tab (prewarm) (issue 1 option b) ---
    '   The reason for not creating it synchronously here is to avoid the busiest stretch of the startup sequence (right after this,
    '   the Controllers of the first tab / TabBar / NavBar are created one after another).
    '   Quietly warm it after PREWARM_DELAY_MS with a one-shot Timer.
    RequestPrewarmAsync

    Init = True
End Function


' ============================================================
' AddTab ? add one empty-shell tab (does not Navigate)
'   Creates a new Wv2Pane, creates a Controller from the shared Environment,
'   waits for Ready, sets Bounds, and adds it to the Collection.
'   Return value: the created Wv2Pane (Nothing on failure)
' ============================================================
Public Function AddTab() As Wv2Pane
    ' --- The first half (Pane creation to standard handler registration) was split out to CreatePaneCore (9.27) ---
    Dim pane As Wv2Pane
    Set pane = CreatePaneCore()
    If pane Is Nothing Then Exit Function

    ' --- 6. Add to the Collection ---
    m_tabs.Add pane
    Debug.Print "Wv2Browser.AddTab: tab added OK (current tab count=" & m_tabs.Count & ")"

    ' --- 6.5 Notify the tab addition (9.11a, loosely coupled event) ---
    '   Fire it before ActivateTab. This way the listener receives in the order "first added, then
    '   active changed" (AddTab calls ActivateTab right after).
    RaiseEvent TabAdded(m_tabs.Count)

    ' --- 7. Automatically activate the added tab (issue 7 a) ---
    '   Visibility control is not done here but delegated to ActivateTab (separation of responsibility).
    '   This keeps "the visibility-control logic in the single place ActivateTab".
    ActivateTab m_tabs.Count

    ' --- 8. If there is no spare tab, reserve its creation again (9.27) ---
    '   During the startup sequence prewarm may be skipped by the re-entrancy guard.
    '   If we re-arm here every time AddTab finishes, one is eventually always warmed.
    EnsurePrewarmScheduled

    Set AddTab = pane
End Function


' ============================================================
' CreatePaneCore / CreatePaneCoreBody (9.27, split out of the first half of the old AddTab = issue 2 option a)
'
'   Responsible up to "making one Pane and getting it into a usable state". Registration in m_tabs,
'   firing TabAdded and ActivateTab are [not done]. There are two callers:
'     - AddTab               -> registers the made Pane as is as a tab
'     - CreatePrewarmPaneNow -> hides the made Pane in the back as a spare
'
'   *Reason for the two-tier structure (Core / CoreBody)*
'     CoreBody has multiple Exit Functions on failure. If the re-entrancy guard (m_creatingPane) were raised and lowered
'     inside CoreBody, we would need to guarantee lowering it at every Exit, and the code gets dirty.
'     Sandwiching it in a thin wrapper (Core) lowers the guard in one place wherever Body exits.
'
'   *Why this guard is needed* WaitPaneReady inside CoreBody spins DoEvents. If
'     m_prewarmTimer_TimerCall interrupts in the middle, Pane creation becomes nested. We look at the guard and
'     make it back off (prevention so as not to approach spec facts 14 / 20).
' ============================================================
Private Function CreatePaneCore() As Wv2Pane
    ' *9.28: flag -> counter* Even if called nested, it does not return to 0 until the outer one finishes.
    m_creatingPane = m_creatingPane + 1
    Set CreatePaneCore = CreatePaneCoreBody()
    m_creatingPane = m_creatingPane - 1
    If m_creatingPane < 0 Then m_creatingPane = 0   ' lower-bound clip just in case
End Function

Private Function CreatePaneCoreBody() As Wv2Pane
    If m_state <> Bs_EnvReady Then
        Debug.Print "Wv2Browser.CreatePaneCore: Environment not Ready (state=" & m_state & ")"
        Exit Function
    End If

    ' --- 1. Create the Pane and create a Controller from the shared Environment ---
    Dim pane As Wv2Pane
    Set pane = New Wv2Pane
    pane.Init m_env, m_hwndHost

    ' --- 2. Wait for the Controller to be Ready (confined inside) ---
    If Not WaitPaneReady(pane) Then
        If pane.IsFailed Then
            Debug.Print "Wv2Browser.CreatePaneCore: Controller creation failed (state=" & _
                        pane.state & ", lastError=&H" & Hex(pane.LastError) & ")"
        Else
            Debug.Print "Wv2Browser.CreatePaneCore: Controller creation timed out (state=" & _
                        pane.state & ")"
        End If
        Exit Function
    End If
    Debug.Print "Wv2Browser.CreatePaneCore: Controller created OK (pCtrl=" & pane.CtrlPtr & ")"

    ' --- 3. Set the initial Bounds (9.11c: content rectangle cache first, option A) ---
    '   If the upper level (UserForm) has told us the rectangle via ResizeAll, use it
    '   (the correct area that avoids the tab-bar band). If not yet, the whole frame as before.
    '   In any case UserForm_Resize reissues and sets the final layout.
    Dim rc As RECT
    If m_hasContentRect Then
        rc = m_lastContentRect
        Debug.Print "Wv2Browser.CreatePaneCore: using the cached content rectangle for the initial Bounds"
    Else
        GetClientRect m_hwndHost, rc
        Debug.Print "Wv2Browser.CreatePaneCore: using the whole frame for the initial Bounds (ResizeAll not yet run)"
    End If
    Dim hr As Long
    hr = pane.Ctrl_PutBounds(rc)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.CreatePaneCore: Ctrl_PutBounds failed hr=&H" & Hex(hr)
        ' A Bounds failure is not fatal, so continue (the tab itself was created)
    End If

    ' --- 4. Get ICoreWebView2 (View) ---
    '   *Important* Controller creation alone leaves m_pView = 0. View_Navigate /
    '   View_NavigateToString need m_pView, so always get it here.
    '   *Even more important in 9.27* What is passed to put_NewWindow is this ViewPtr itself.
    hr = pane.Ctrl_GetCoreWebView2()
    If hr <> 0 Then
        Debug.Print "Wv2Browser.CreatePaneCore: Ctrl_GetCoreWebView2 failed hr=&H" & Hex(hr)
        Exit Function
    End If
    If pane.ViewPtr = 0 Then
        Debug.Print "Wv2Browser.CreatePaneCore: ViewPtr is 0 (retrieval succeeded but the value is empty)"
        Exit Function
    End If
    Debug.Print "Wv2Browser.CreatePaneCore: View obtained OK (pView=" & pane.ViewPtr & ")"

    ' --- 5. Automatically register the standard navigation handlers (option 10-A, 9.10c-2) ---
    '   Foundational events (NavigationStarting/Completed/DocumentTitleChanged/
    '   NewWindowRequested/HistoryChanged) are registered automatically on all tabs.
    '   *Also register on the spare tab* so the title and URL can be tracked from the moment it is adopted.
    RegisterStandardHandlers pane

    ' --- 5.5 Set a back-reference on the Wv2Pane (9.10e, option Y) ---
    '   So that the Pane can call the Browser's functions at NewWindowRequested.
    '   *This becomes a circular reference, so at release it is broken via Shutdown -> ClearBrowserRef (design principle 58).
    pane.SetBrowser Me

    Set CreatePaneCoreBody = pane
End Function


' ============================================================
' RegisterStandardHandlers ? register the standard navigation handlers on one tab
'   Added in stage 9.10c-2 (option 10-A). Follows steps 10?13 of the existing StartWebView2
'   verbatim. The responsibility of handler registration is consolidated in this one place (in the spirit of design principle 57).
'
'   *Subject to design principles 45/52* These View_AddXxx are handler-type methods that internally hand a COM object
'   to WebView2. The Wv2Pane-side implementation is correctly written on the premise that it goes through VTABLE_OBJ_OFFSET,
'   so here we only need to call Wv2Pane's public methods
'   (we do not touch the slots ourselves).
' ============================================================
Private Sub RegisterStandardHandlers(ByVal pane As Wv2Pane)
    Dim hr As Long

    hr = pane.View_AddNavigationStarting()
    If hr <> 0 Then Debug.Print "Wv2Browser.RegisterStandardHandlers: NavigationStarting failed hr=&H" & Hex(hr)

    hr = pane.View_AddNavigationCompleted()
    If hr <> 0 Then Debug.Print "Wv2Browser.RegisterStandardHandlers: NavigationCompleted failed hr=&H" & Hex(hr)

    hr = pane.View_AddDocumentTitleChanged()
    If hr <> 0 Then Debug.Print "Wv2Browser.RegisterStandardHandlers: DocumentTitleChanged failed hr=&H" & Hex(hr)

    hr = pane.View_AddNewWindowRequested()
    If hr <> 0 Then Debug.Print "Wv2Browser.RegisterStandardHandlers: NewWindowRequested failed hr=&H" & Hex(hr)

    ' --- 9.12a: promote HistoryChanged to a standard handler (option 2-A-1) ---
    '   The enabled state of the back/forward buttons needs to be tracked on all tabs, not only SPA, so
    '   we stop treating it as SPA-only (EnableSpaHandlers) and register it automatically on all tabs.
    '   Besides pushState/replaceState/popstate, it also fires on ordinary back/forward.
    hr = pane.View_AddHistoryChanged()
    If hr <> 0 Then Debug.Print "Wv2Browser.RegisterStandardHandlers: HistoryChanged failed hr=&H" & Hex(hr)

    Debug.Print "Wv2Browser.RegisterStandardHandlers: 5 standard handlers registered"
End Sub


' ============================================================
' EnableSpaHandlers ? register event handlers for SPA on a pane (9.10d-2b, option 2-c)
'
'   Registers 2 kinds of events for observing an SPA (history operations + DOM construction):
'     - HistoryChanged   (base ICoreWebView2, vtable 13) : fires on all of pushState/replaceState/
'                          popstate (spec fact 12)
'     - DOMContentLoaded (ICoreWebView2_2, vtable 64)     : fires only on the first load
'                          (does not fire on pushState = division of roles with HistoryChanged)
'
'   Not included in the standard 4 handlers (RegisterStandardHandlers); registered
'   explicitly only on SPA tabs (9.10c-2 issue 10). AddTabWithUrlForSpa calls it internally,
'   but it is made Public so it can also be retrofitted onto existing tabs (option 2-c).
'
'   *Design principle 45* View_AddHistoryChanged / View2_AddDOMContentLoaded are already correctly implemented
'   on the Wv2Pane side on the premise that they go through VTABLE_OBJ_OFFSET. Here we just
'   call those public methods (we do not touch the slots ourselves).
' ============================================================
Public Sub EnableSpaHandlers(ByVal pane As Wv2Pane)
    If pane Is Nothing Then
        Debug.Print "Wv2Browser.EnableSpaHandlers: pane is Nothing"
        Exit Sub
    End If

    ' --- 9.12a: HistoryChanged was removed from here ---
    '   it was promoted to RegisterStandardHandlers (automatic registration on all tabs), so registering again here would
    '   double-register it. Leave only the SPA-specific event (DOMContentLoaded).
    Dim hr As Long
    hr = pane.View2_AddDOMContentLoaded()
    If hr <> 0 Then Debug.Print "Wv2Browser.EnableSpaHandlers: DOMContentLoaded failed hr=&H" & Hex(hr)

    Debug.Print "Wv2Browser.EnableSpaHandlers: 1 SPA handler " & _
                "(DOMContentLoaded) registered; note: HistoryChanged was promoted to a standard handler (9.12a)"
End Sub


' ============================================================
' AddTabWithUrl ? add a tab and Navigate to a URL
' ============================================================
Public Function AddTabWithUrl(ByVal url As String) As Wv2Pane
    Dim pane As Wv2Pane
    Set pane = AddTab()
    If pane Is Nothing Then Exit Function

    Dim hr As Long
    hr = pane.View_Navigate(url)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.AddTabWithUrl: View_Navigate failed hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Browser.AddTabWithUrl: Navigate OK (" & url & ")"
    End If

    Set AddTabWithUrl = pane
End Function


' ============================================================
' AddTabWithHtml ? add a tab and display an HTML string on the spot
'
'   *Spec fact 6* The origin of a page displayed with NavigateToString is about:blank
'   (= the "null" origin). With this origin, history.pushState /
'   replaceState are rejected with SecurityError (web standard).
'   A full-fledged SPA test involving pushState/replaceState is done in 9.10d using a tab
'   with an http(s) origin (AddTabWithUrl).
' ============================================================
Public Function AddTabWithHtml(ByVal htmlContent As String) As Wv2Pane
    Dim pane As Wv2Pane
    Set pane = AddTab()
    If pane Is Nothing Then Exit Function

    Dim hr As Long
    hr = pane.View_NavigateToString(htmlContent)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.AddTabWithHtml: View_NavigateToString failed hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Browser.AddTabWithHtml: NavigateToString OK (" & _
                    Len(htmlContent) & " chars)"
    End If

    Set AddTabWithHtml = pane
End Function


' ============================================================
' AddTabWithUrlForSpa ? create an SPA tab in one go (9.10d-2b, option F-1)
'
'   The entry method for SPA verification and operation: maps a local folder to a virtual host, opens the HTML there
'   with an https origin, and even registers the SPA handlers.
'
'   Steps (order is important):
'     1. AddTab            : get Controller/View, register the standard 4 handlers, activate
'     2. SetVirtualHostNameToFolderMapping : *required before Navigate and after getting View*
'                            (getting View is completed inside AddTab, so the order is OK)
'     3. View_Navigate     : https://<hostName>/<fileName>
'                            *Spec fact 13* The root "/" gives ERR_ACCESS_DENIED, so
'                            state it down to the file name
'     4. EnableSpaHandlers : register HistoryChanged + DOMContentLoaded
'
'   Arguments:
'     hostName   : virtual host name (e.g. "appassets.example")
'     folderPath : absolute path of the local folder to map to
'     fileName   : start file name in the folder (e.g. "index.html")
'
'   accessKind is fixed to ALLOW (=1). Make it an argument if it becomes necessary.
'
'   Return value: the created Wv2Pane (Nothing on failure; however, if mapping/Navigate fails after
'           the tab was created, the tab itself is returned = partial success can be grasped)
' ============================================================
Public Function AddTabWithUrlForSpa( _
    ByVal hostName As String, _
    ByVal folderPath As String, _
    ByVal fileName As String) As Wv2Pane

    ' --- 1. Add the shell tab (up to getting View, registering standard handlers, activation) ---
    Dim pane As Wv2Pane
    Set pane = AddTab()
    If pane Is Nothing Then Exit Function

    ' --- 2. Set the virtual host mapping (*required before Navigate) ---
    Dim hr As Long
    hr = pane.View3_SetVirtualHostNameToFolderMapping(hostName, folderPath, 2)   ' 2 = DENY_CORS (English fork; was 1 = ALLOW, which let any site read the mapped folder)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.AddTabWithUrlForSpa: mapping failed hr=&H" & Hex(hr)
        Set AddTabWithUrlForSpa = pane   ' the tab itself is already created, so return it
        Exit Function
    End If

    ' --- 3. Navigate to https://<host>/<file> (spec fact 13: state it down to the file name) ---
    Dim url As String
    url = "https://" & hostName & "/" & fileName
    hr = pane.View_Navigate(url)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.AddTabWithUrlForSpa: Navigate failed hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Browser.AddTabWithUrlForSpa: Navigate OK (" & url & ")"
    End If

    ' --- 4. Register the SPA handlers (HistoryChanged + DOMContentLoaded) ---
    EnableSpaHandlers pane

    Set AddTabWithUrlForSpa = pane
End Function


' ============================================================
' TabCount ? current number of tabs
' ============================================================
Public Property Get TabCount() As Long
    TabCount = m_tabs.Count
End Property


' ============================================================
' ActivePane ? return the active tab's Wv2Pane
'   Formalized on an activeIndex basis in 9.10c-1 (formalization of issue 4).
'   Returns the tab pointed at by m_activeIndex (1-based, 0 = none).
' ============================================================
Public Property Get ActivePane() As Wv2Pane
    If m_activeIndex < 1 Or m_activeIndex > m_tabs.Count Then
        Debug.Print "Wv2Browser.ActivePane: no active tab (activeIndex=" & _
                    m_activeIndex & ", tabCount=" & m_tabs.Count & ")"
        Set ActivePane = Nothing
        Exit Property
    End If
    Set ActivePane = m_tabs.Item(m_activeIndex)
End Property

' ============================================================
' Environment ? lend out the shared Environment (9.11b, for Wv2TabBar)
'   Wv2TabBar uses the Environment (shared by all panes) held by this
'   Browser to create the single pane for drawing the tab bar. Use it after Init completes
'   (Bs_EnvReady). While not Ready it returns Nothing.
' ============================================================
Public Property Get Environment() As Wv2Environment
    Set Environment = m_env
End Property


' ============================================================
' ActiveIndex ? index of the active tab (1-based, 0 = none)
' ============================================================
Public Property Get ActiveIndex() As Long
    ActiveIndex = m_activeIndex
End Property


' ============================================================
' ActivateTab ? make only the tab with the given index visible and the others invisible
'   Added in stage 9.10c-1. The only entry for tab switching (single source of truth).
'
'   Argument: index (1-based)
'   Return value: True on success
'
'   Implementation policy (issue 6 plan A): loop over all tabs and do Ctrl_PutIsVisible(True)
'   only for the tab at index and False for the others. Blunt but reliable, and
'   there are no holes such as "all invisible". The tab count is realistically small, so no performance problem.
'
'   *Responsibility* This method is in charge of only "visibility control + updating activeIndex".
'   It does not add or delete tabs. AddTab and the like call (delegate to) this method, so that
'   the visibility-control logic is consolidated in one place.
' ============================================================
Public Function ActivateTab(ByVal index As Long) As Boolean
    If index < 1 Or index > m_tabs.Count Then
        Debug.Print "Wv2Browser.ActivateTab: index out of range (index=" & index & _
                    ", tabCount=" & m_tabs.Count & ")"
        Exit Function
    End If

    Dim i As Long
    Dim pane As Wv2Pane
    Dim hr As Long
    For i = 1 To m_tabs.Count
        Set pane = m_tabs.Item(i)
        hr = pane.Ctrl_PutIsVisible(i = index)
        If hr <> 0 Then
            Debug.Print "Wv2Browser.ActivateTab: PutIsVisible failed for tab " & i & " hr=&H" & Hex(hr)
        End If
    Next i

    m_activeIndex = index
    Debug.Print "Wv2Browser.ActivateTab: tab " & index & " activated (visible)"

    ' --- Notify the active change (9.11a, loosely coupled event) ---
    RaiseEvent ActiveChanged(index)
    ActivateTab = True
End Function


' ============================================================
' CloseTab ? close the tab at the given index (added in stage 9.11a)
'
'   Argument: index (1-based)
'   Return value: True on success
'
'   Steps:
'     1. index range check
'     2. Break the back-reference of the Pane being closed (ClearBrowserRef, design principle 58) ->
'        explicitly close the Controller with Ctrl_Close (vtable 24, double Close is harmless)
'     3. m_tabs.Remove index (Pane.Class_Terminate when the reference count reaches 0)
'     4. Correct m_activeIndex + repair the visible state (consolidated in ActivateTab, design principle 57):
'        - closed the active : prefer the right neighbor, else activate the left (the end)
'        - closed one before the active : decrement m_activeIndex by 1 to keep the same tab
'        - tab count became 0 : m_activeIndex = 0 (no active, zero tabs allowed)
'     5. RaiseEvent TabClosed(index) (if the active changes, ActiveChanged is also emitted via ActivateTab)
'
'   *Zero tabs allowed* The Browser has no minimum-tab-count policy. By leaving it to the upper level (the tab-bar
'   implementation side), we do not bring in extra constraints even for uses that need no tab bar.
' ============================================================
Public Function CloseTab(ByVal index As Long) As Boolean
    If index < 1 Or index > m_tabs.Count Then
        Debug.Print "Wv2Browser.CloseTab: index out of range (index=" & index & _
                    ", tabCount=" & m_tabs.Count & ")"
        Exit Function
    End If

    Dim pane As Wv2Pane
    Set pane = m_tabs.Item(index)

    ' --- 1. Break the back-reference first (design principle 58) ---
    '   Cut the circular reference Browser<->Pane before release. Same procedure as Shutdown.
    pane.ClearBrowserRef

    ' --- 2. Explicitly close the Controller ---
    '   Ctrl_Close Releases View3/View2/View/Ctrl in order and also
    '   settles the registered handlers (a double Close is S_OK and harmless).
    Dim hr As Long
    hr = pane.Ctrl_Close()
    If hr <> 0 Then
        Debug.Print "Wv2Browser.CloseTab: Ctrl_Close failed hr=&H" & Hex(hr) & " (continuing)"
    End If

    ' --- 3. Remove from the Collection (Pane.Class_Terminate when the reference count reaches 0) ---
    Set pane = Nothing
    m_tabs.Remove index
    Debug.Print "Wv2Browser.CloseTab: closed tab " & index & " (remaining tab count=" & m_tabs.Count & ")"

    ' --- 4. Active correction + repair of the visible state ---
    Dim newActive As Long
    If m_tabs.Count = 0 Then
        ' 0 tabs: no active (zero tabs allowed)
        m_activeIndex = 0
        Debug.Print "Wv2Browser.CloseTab: all tabs closed (activeIndex=0)"
    ElseIf index < m_activeIndex Then
        ' Closed one before the active: the index shifts back by one, so correct it and
        ' keep the "same tab". The visible state does not change, but set it right with ActivateTab.
        newActive = m_activeIndex - 1
        ActivateTab newActive
    ElseIf index = m_activeIndex Then
        ' Closed the active: prefer the right neighbor, else the left (the end).
        '   After Remove, the original right neighbor is now at the same index. If the index exceeds
        '   the remaining tab count (when the end was closed), activate the end.
        If index <= m_tabs.Count Then
            newActive = index
        Else
            newActive = m_tabs.Count
        End If
        ActivateTab newActive
    Else
        ' Closed one after the active: activeIndex does not change.
        '   The visible state does not change either, so ActivateTab is unnecessary, but we set it right just in case.
        ActivateTab m_activeIndex
    End If

    ' --- 5. Notify the tab close (9.11a, loosely coupled event) ---
    RaiseEvent TabClosed(index)
    CloseTab = True
End Function


' ============================================================
' MoveTab ? reorder a tab from fromIndex to toIndex (added in stage 9.15)
'
'   Caller: {"cmd":"reorder","from":N,"to":M} in Wv2TabBar.OnPaneWebMessage
'   Both arguments are 1-based. If from == to or out of range, does nothing and returns False.
'
'   *Steps*
'     1. Grab the active Pane by reference (because the index moves with reordering).
'     2. Remove the Pane to move and re-insert it at the target position with Add (before/after).
'        - to == 1            : to the front (Add pane, , 1)
'        - otherwise          : after the (to-1)th (Add pane, , , to-1)
'          Note: compute carefully with the base index on the Collection after Remove.
'     3. Reverse-look-up the grabbed active Pane by linear scan and update m_activeIndex.
'     4. Set the visible state right with ActivateTab -> RaiseEvent ActiveChanged -> global sync
'        (design principle 60, option A-2b). No new event is fired.
'
'   *No waiting for completion* It only has array operations (Remove/Add) and PutIsVisible and contains no wait like Controller
'     creation, so it is safe to call synchronously from inside a WebMessage handler
'     (same as activate/close. The SafeTimer detour like newtab is unnecessary).
' ============================================================
Public Function MoveTab(ByVal fromIndex As Long, ByVal toIndex As Long) As Boolean
    Dim n As Long
    n = m_tabs.Count

    ' --- Range check / no-change check ---
    If fromIndex < 1 Or fromIndex > n Then
        Debug.Print "Wv2Browser.MoveTab: from out of range (from=" & fromIndex & ", tabCount=" & n & ")"
        Exit Function
    End If
    If toIndex < 1 Or toIndex > n Then
        Debug.Print "Wv2Browser.MoveTab: to out of range (to=" & toIndex & ", tabCount=" & n & ")"
        Exit Function
    End If
    If fromIndex = toIndex Then
        ' The position does not change, so do nothing (treated as success).
        MoveTab = True
        Exit Function
    End If

    ' --- 1. Grab the active Pane by reference (because the index moves with reordering) ---
    Dim activePaneRef As Wv2Pane
    If m_activeIndex >= 1 And m_activeIndex <= n Then
        Set activePaneRef = m_tabs.Item(m_activeIndex)
    End If

    ' --- 2. Take out the target Pane and Remove ---
    Dim moving As Wv2Pane
    Set moving = m_tabs.Item(fromIndex)
    m_tabs.Remove fromIndex

    ' After Remove the Collection has n-1 elements. Insert moving here at the toIndex position.
    '   The before/after of Collection.Add refers to the "index on the Collection after Remove".
    '   The goal is that "after reordering, moving is at the toIndex-th position".
    If toIndex = 1 Then
        ' To the front: insert before the current first (before:=1).
        m_tabs.Add moving, , 1
    Else
        ' To be the toIndex-th = insert after the (toIndex-1)th element.
        '   Even in the Collection after Remove, elements 1..(toIndex-1) remain in the original order
        '   (the effect of removing fromIndex reaches only toIndex onward, or
        '   when fromIndex < toIndex the (toIndex-1)th matches the original toIndex-th).
        '   In either case after:=(toIndex-1) makes it the "toIndex-th after reordering".
        m_tabs.Add moving, , , toIndex - 1
    End If

    Debug.Print "Wv2Browser.MoveTab: moved tab from " & fromIndex & " -> " & toIndex & " (tabCount=" & m_tabs.Count & ")"

    ' --- 3. Reverse-look-up the new index of the active Pane ---
    If Not activePaneRef Is Nothing Then
        Dim i As Long
        For i = 1 To m_tabs.Count
            If m_tabs.Item(i) Is activePaneRef Then
                m_activeIndex = i
                Exit For
            End If
        Next i
    End If

    ' --- 4. Set the visible state right -> ActiveChanged -> global sync (option A-2b) ---
    '   The visible state itself does not change, but going through ActivateTab prompts
    '   PushSyncToJs and reflects the reordering result in TabBar/NavBar.
    If m_activeIndex >= 1 And m_activeIndex <= m_tabs.Count Then
        ActivateTab m_activeIndex
    End If

    MoveTab = True
End Function


' ============================================================
' Debug_TabOrderTitles ? return the current tab order as identification tags joined with "|" (9.15, verification entry)
'
'   A thin public wrapper for verifying the reordering result of MoveTab without launching WebView2.
'   A normal tab returns its title ("(t<index>)" if unset), but during verification
'   Debug_SeedDummyTabs records the correspondence of each dummy Pane's ObjPtr -> serial tag ("A","B",...)
'   in m_dbgTagMap, so even if the index moves through reordering,
'   "which number that tab originally was" can be tracked stably.
'   Same positioning as Debug_NormalizeUrl, a "living verification asset".
' ============================================================
Public Function Debug_TabOrderTitles() As String
    Dim i As Long
    Dim parts As String
    Dim tag As String
    For i = 1 To m_tabs.Count
        tag = ""
        ' If there is a verification tag map, look up by ObjPtr (stable identification robust to reordering)
        If Not m_dbgTagMap Is Nothing Then
            Dim key As String
            key = CStr(ObjPtr(m_tabs.Item(i)))
            On Error Resume Next
            tag = m_dbgTagMap.Item(key)
            On Error GoTo 0
        End If
        ' If there is no tag, regard it as a production tab and use the title (placeholder if unset)
        If Len(tag) = 0 Then
            On Error Resume Next
            tag = m_tabs.Item(i).DocumentTitle
            On Error GoTo 0
            If Len(tag) = 0 Then tag = "(t" & i & ")"
        End If
        If i > 1 Then parts = parts & "|"
        parts = parts & tag
    Next i
    Debug_TabOrderTitles = parts
End Function


' ============================================================
' Debug_SeedDummyTabs ? plant n tagged dummy tabs for verification (9.15, verification entry)
'
'   An alternative entry so that Test_9_15_Reorder can verify only the reordering logic of MoveTab without launching WebView2.
'   Pushes n empty Wv2Panes onto m_tabs and builds m_dbgTagMap, which maps each Pane's ObjPtr to a
'   serial tag ("A","B","C",...). activeIndex is 1.
'   It verifies by passing through the production MoveTab as is, so there is no double management.
'   Visibility control (PutIsVisible in ActivateTab) emits a failure log internally because the dummy Pane's View is Nothing,
'   but it is harmless (it does not affect order verification).
' ============================================================
Public Sub Debug_SeedDummyTabs(ByVal n As Long)
    Set m_tabs = New Collection
    Set m_dbgTagMap = New Collection
    Dim i As Long
    Dim p As Wv2Pane
    Dim tag As String
    For i = 1 To n
        Set p = New Wv2Pane
        m_tabs.Add p
        tag = Chr$(64 + i)   ' 1->"A", 2->"B", ...
        m_dbgTagMap.Add tag, CStr(ObjPtr(p))
    Next i
    If n >= 1 Then
        m_activeIndex = 1
    Else
        m_activeIndex = 0
    End If
End Sub


' ============================================================
' Debug_ActiveIndex ? return the current m_activeIndex (9.15, verification entry)
'   Used to verify active tracking after MoveTab.
' ============================================================
Public Function Debug_ActiveIndex() As Long
    Debug_ActiveIndex = m_activeIndex
End Function


' ============================================================
' NotifyTitleChanged ? receive a title change from a Pane, reverse-look-up the index and
'   fire RaiseEvent TabTitleChanged (added in stage 9.11a)
'
'   Caller: Wv2Pane.View_OnDocumentTitleChanged (via the back-reference m_browser)
'   Arguments: senderPane = the Pane whose title changed (itself, Me), title = the new title
'
'   *Reverse lookup* The Collection finds the index by object-reference match. The tab count is
'   realistically small, so linear search is enough. If not found (a delayed event of an already closed
'   tab, etc.), do nothing.
' ============================================================
Public Sub NotifyTitleChanged(ByVal senderPane As Wv2Pane, ByVal title As String)
    Dim i As Long
    For i = 1 To m_tabs.Count
        If m_tabs.Item(i) Is senderPane Then
            RaiseEvent TabTitleChanged(i, title)
            Exit Sub
        End If
    Next i
    Debug.Print "Wv2Browser.NotifyTitleChanged: senderPane not found in m_tabs (ignored)"
End Sub


' ============================================================
' NotifyUrlChanged (9.12a, receives a URL state notification from a Pane and reverse-looks-up the index)
'
'   Called from a Pane's NavigationCompleted / HistoryChanged. Reverse-looks-up the index by object-reference
'   match and does RaiseEvent TabUrlChanged.
'   (exactly the same structure as NotifyTitleChanged of 9.11a)
'
'   *Not called from the tab-bar pane* The tab-bar/nav-bar panes are not in m_tabs, and moreover
'   they have no m_browser back-reference (SetBrowser is not done),
'   so it does not reach here.
' ============================================================
Public Sub NotifyUrlChanged(ByVal senderPane As Wv2Pane, ByVal url As String, _
                            ByVal canGoBack As Boolean, ByVal canGoForward As Boolean)
    Dim i As Long
    For i = 1 To m_tabs.Count
        If m_tabs.Item(i) Is senderPane Then
            RaiseEvent TabUrlChanged(i, url, canGoBack, canGoForward)
            Exit Sub
        End If
    Next i
    Debug.Print "Wv2Browser.NotifyUrlChanged: senderPane not found in m_tabs (ignored)"
End Sub


' ============================================================
' Delegating operations to the active tab (9.12a, public API for the nav bar)
'
'   The nav bar wants to send operations to the "active tab", but letting it touch the Pane directly breaks loose
'   coupling (the nav bar would come to know the Browser's internal structure). We provide thin methods that the Browser
'   proxies so that the nav bar talks only to the Browser.
'
'   *Design principle 63* None of them involves "waiting for completion" (Navigate/GoBack/GoForward/
'     Reload/Stop are all fire-and-forget), so it is safe to call synchronously from inside a WebMessage event handler.
'     No deferred execution like the newtab of 9.11d (waiting for Controller creation) is needed.
'
'   When there are 0 tabs (no active), does nothing and returns False / an empty string.
' ============================================================
Public Function NavigateActive(ByVal url As String) As Boolean
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then
        Debug.Print "Wv2Browser.NavigateActive: no active tab (ignored)"
        Exit Function
    End If

    Dim target As String
    target = NormalizeUrl(url)
    If Len(target) = 0 Then
        Debug.Print "Wv2Browser.NavigateActive: URL is empty (ignored)"
        Exit Function
    End If

    Dim hr As Long
    hr = pane.View_Navigate(target)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.NavigateActive: Navigate failed hr=&H" & Hex(hr) & " (" & target & ")"
        Exit Function
    End If

    Debug.Print "Wv2Browser.NavigateActive: Navigate OK (" & target & ")"
    NavigateActive = True
End Function

Public Function GoBackActive() As Boolean
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then Exit Function

    If Not pane.View_GetCanGoBack() Then
        Debug.Print "Wv2Browser.GoBackActive: cannot go back (no history)"
        Exit Function
    End If

    Dim hr As Long
    hr = pane.View_GoBack()
    If hr <> 0 Then
        Debug.Print "Wv2Browser.GoBackActive: GoBack failed hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2Browser.GoBackActive: GoBack OK"
    GoBackActive = True
End Function

Public Function GoForwardActive() As Boolean
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then Exit Function

    If Not pane.View_GetCanGoForward() Then
        Debug.Print "Wv2Browser.GoForwardActive: cannot go forward (no history)"
        Exit Function
    End If

    Dim hr As Long
    hr = pane.View_GoForward()
    If hr <> 0 Then
        Debug.Print "Wv2Browser.GoForwardActive: GoForward failed hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2Browser.GoForwardActive: GoForward OK"
    GoForwardActive = True
End Function

Public Function ReloadActive() As Boolean
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then Exit Function

    Dim hr As Long
    hr = pane.View_Reload()
    If hr <> 0 Then
        Debug.Print "Wv2Browser.ReloadActive: Reload failed hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2Browser.ReloadActive: Reload OK"
    ReloadActive = True
End Function

Public Function StopActive() As Boolean
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then Exit Function

    Dim hr As Long
    hr = pane.View_Stop()
    If hr <> 0 Then
        Debug.Print "Wv2Browser.StopActive: Stop failed hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2Browser.StopActive: Stop OK"
    StopActive = True
End Function


' --- Current state of the active tab (for the nav bar's initial display) ---
'   A read entry for the nav bar to query "the current state" before the event (TabUrlChanged)
'   arrives or right after a tab switch.
Public Function GetActiveUrl() As String
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then Exit Function
    If Not pane.IsReady Then Exit Function
    GetActiveUrl = pane.View_GetSource()
End Function

Public Function GetActiveCanGoBack() As Boolean
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then Exit Function
    If Not pane.IsReady Then Exit Function
    GetActiveCanGoBack = pane.View_GetCanGoBack()
End Function

Public Function GetActiveCanGoForward() As Boolean
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then Exit Function
    If Not pane.IsReady Then Exit Function
    GetActiveCanGoForward = pane.View_GetCanGoForward()
End Function


' ============================================================
' NormalizeUrl (9.12a, scheme completion of address-bar input, option 3-A)
'
'   Shapes the string the user typed in the address bar into a navigable URL.
'     "example.com"        -> "https://example.com"
'     "https://example.com" -> as is
'     "about:blank"         -> as is (regarded as having a scheme)
'
'   *The search-term fallback is out of scope (option 3-A)* The feature of throwing input such as "words to search"
'     at a search engine needs URL encoding, so it is deferred to the next stage.
'     For now it only supplies a scheme and Navigates, and if it cannot be resolved the browser's
'     error page appears.
' ============================================================
Private Function NormalizeUrl(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    If Len(t) = 0 Then Exit Function

    ' Whether a scheme already exists (contains "://", or a scheme form such as "about:")
    If InStr(t, "://") > 0 Then
        NormalizeUrl = t
        Exit Function
    End If

    ' Has a scheme such as about: / data: / file: (the colon comes before any whitespace)
    Dim colonPos As Long
    colonPos = InStr(t, ":")
    If colonPos > 1 Then
        Dim head As String
        head = Left$(t, colonPos - 1)
        ' The scheme name is regarded as letters only (so as not to misjudge "example.com:8080",
        ' anything containing digits is excluded)
        If Wv2Url.IsAlphaOnly(head) Then
            NormalizeUrl = t
            Exit Function
        End If
    End If

    ' By here "has a scheme" has been excluded. Sort the rest into URL / search term (option A, 9.14).
    '   - Contains a dot and no whitespace -> regard as a host name and prefix https:// (URL)
    '   - Otherwise (contains whitespace / no dot) -> regard as a search term and go to Google search
    ' Nearly the same judgment as the Chrome/Edge address bar. A bare "localhost" falls to search, but this is rare in practice.
    If InStr(t, ".") > 0 And InStr(t, " ") = 0 Then
        NormalizeUrl = "https://" & t
    Else
        NormalizeUrl = BuildSearchUrl(t)
    End If
End Function


' ============================================================
' Debug_NormalizeUrl (9.14, logic verification entry)
'
'   NormalizeUrl is Private, so this is a thin public wrapper for checking only the result from outside.
'   It passes through the very same NormalizeUrl that NavigateActive uses internally, so
'   the verification logic is not double-managed against the production logic.
'   (Called from Test_9_14_SearchFallback in Wv2Thunks.bas. Same positioning as the low-level tests left in 9.13,
'    a "living verification asset".)
' ============================================================
Public Function Debug_NormalizeUrl(ByVal s As String) As String
    Debug_NormalizeUrl = NormalizeUrl(s)
End Function


' ============================================================
' BuildSearchUrl (9.14 -> 9.16, search-term fallback, option A/option E-1/option P-1)
'
'   Converts a search term into a search engine URL.
'     "<nihon tenki in Japanese>"  -> m_searchTemplate & "%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97"
'
'   *9.16: the engine can be swapped via m_searchTemplate* (option 1-A)
'     In 9.14 "https://www.google.com/search?q=" was hard-coded, but
'     now the module variable m_searchTemplate is used as the prefix template.
'     Default is Google (set in Class_Initialize). Change it with Property SearchEngine.
' ============================================================
Private Function BuildSearchUrl(ByVal terms As String) As String
    BuildSearchUrl = m_searchTemplate & Wv2Url.UrlEncodeComponent(terms)
End Function


' ============================================================
' SearchEngine (9.16, search URL template of the search engine, option 1-A + option 2-A)
'
'   Reads and writes the prefix template of the search URL used by the search-term fallback
'   (the third branch of NormalizeUrl). The encoded search term is concatenated at the end.
'     Example: Get -> "https://www.google.com/search?q="
'         Let "https://www.bing.com/search?q=" switches to Bing
'
'   *Empty guard* If an empty string (or whitespace only) is passed, it falls back to the default Google template.
'     No validity check as a URL is done (the template is a trusted value entered by the developer,
'     and the policy is not to do excessive validation).
' ============================================================
' ============================================================
' ============================================================
' NewWindowMode (9.27, a 3-value switch that chooses how NewWindowRequested is handled)
'
'   0 = current behavior (LEGACY)   : takes only the URL with args.get_Uri, does put_Handled(TRUE) and
'                             re-Navigates in our own new tab. Navigate is *always GET*,
'                             so a POST link gives 405 (this is how it was before 9.26).
'   1 = popup delegation     : passes through without setting put_Handled. The runtime opens its own window
'                             and continues the original request (demonstrated by spec fact 23).
'                             POST is preserved but it does not appear on the tab bar, and closing it is manual.
'   2 = prewarm delegation   : *the favorite of 9.27* Passes the warmed spare tab's CoreWebView2 to
'                             args.put_NewWindow (vtable 4). The runtime loads
'                             the original request onto that WebView, so POST is preserved
'                             and it also fits in our own tab.
'
'   *Out-of-range values are rounded to 0* This is a switch hit from the Immediate window on a real machine, so
'     we err on the safe side so a typo does not fall into undefined behavior.
'   *Common to all tabs* The Wv2Pane side reads m_browser.NewWindowMode every time, so switching
'     takes effect immediately even while running (no need to reopen tabs).
'   *Default value = 2 (changed from 0 in 9.28)* See Class_Initialize. Only during
'     PREWARM_DELAY_MS right after startup is the spare not yet created, and new-window requests during that time
'     fall back to the current behavior (a *-marked log is emitted).
' ============================================================
Public Property Get NewWindowMode() As Long
    NewWindowMode = m_newWindowMode
End Property

Public Property Let NewWindowMode(ByVal value As Long)
    Dim v As Long
    v = value
    If v < 0 Or v > 2 Then
        Debug.Print "Wv2Browser.NewWindowMode: out-of-range value " & value & " rounded to 0"
        v = 0
    End If
    m_newWindowMode = v

    Select Case v
        Case 0
            Debug.Print "Wv2Browser.NewWindowMode = 0 (current behavior: put_Handled + own new tab. POST turns into GET)"
        Case 1
            Debug.Print "Wv2Browser.NewWindowMode = 1 (popup delegation: left to the runtime. POST is preserved but does not appear in a tab)"
        Case 2
            Debug.Print "Wv2Browser.NewWindowMode = 2 (prewarm delegation: load onto the spare tab with put_NewWindow)"
            Debug.Print "    Spare tab state: " & Debug_PrewarmStatus()
            If Not PrewarmAvailable Then
                Debug.Print "    *No spare yet, so creation is being reserved (falls back to the current behavior until warmed)*"
                RequestPrewarmAsync
            End If
    End Select
End Property


' ============================================================
' SetHandledWithNewWindow (9.27, insurance switch of issue 5)
'
'   Whether to also set put_Handled(TRUE) right after setting the spare tab's CoreWebView2
'   on put_NewWindow. The official sample sets both, so the default is True.
'
'   *What the insurance is for* In theory the possibility remains that "once NewWindow is set, Handled is unnecessary / setting it
'     actually discards the navigation". If the spare tab should stay white or
'     two windows should open, set this to False and retry, and
'     *without fixing the code* you can isolate which is right.
' ============================================================
Public Property Get SetHandledWithNewWindow() As Boolean
    SetHandledWithNewWindow = m_setHandledWithNewWindow
End Property

Public Property Let SetHandledWithNewWindow(ByVal value As Boolean)
    m_setHandledWithNewWindow = value
    If value Then
        Debug.Print "Wv2Browser.SetHandledWithNewWindow = True (also set put_Handled after put_NewWindow)"
    Else
        Debug.Print "Wv2Browser.SetHandledWithNewWindow = False (fire only put_NewWindow)"
    End If
End Property


' ============================================================
' PopupDelegateMode (9.26 -> turned into a compatibility wrapper in 9.27)
'
'   A Boolean toggle introduced in 9.26. In 9.27 it was merged into NewWindowMode (3 values), so it
'   became *a thin wrapper that only delegates to NewWindowMode internally*.
'   Because of this, Test_9_26_Popup_On / _Off still work as is in 9.27.
'
'   Get : True when NewWindowMode = 1 (popup delegation)
'   Let : True -> NewWindowMode = 1 / False -> NewWindowMode = 0
'         Note: setting False *also clears prewarm delegation (2)*.
'            To return to the favorite of 9.27, explicitly set NewWindowMode = 2.
' ============================================================
Public Property Get PopupDelegateMode() As Boolean
    PopupDelegateMode = (m_newWindowMode = 1)
End Property

Public Property Let PopupDelegateMode(ByVal value As Boolean)
    If value Then
        NewWindowMode = 1
    Else
        NewWindowMode = 0
    End If
End Property


Public Property Get SearchEngine() As String
    SearchEngine = m_searchTemplate
End Property

Public Property Let SearchEngine(ByVal template As String)
    Dim t As String
    t = Trim$(template)
    If Len(t) = 0 Then
        m_searchTemplate = DEFAULT_SEARCH_TEMPLATE
        Debug.Print "Wv2Browser.SearchEngine: empty template, falling back to the default (Google)"
    Else
        m_searchTemplate = t
        Wv2Log.LogD "Wv2Browser.SearchEngine: search template set = " & t
    End If

    ' --- K-2: reserve the reflection onto all tabs showing the settings screen ---
    RequestSettingsSyncAsync
End Property


' ============================================================
' Debug_SearchEngine (9.16, logic verification entry)
'
'   A thin public wrapper that returns the current m_searchTemplate. BuildSearchUrl is Private, so
'   the verification side (Test_9_16_SearchEngine) uses it to confirm that the engine setting is reflected.
'   It returns the very same m_searchTemplate that the production BuildSearchUrl uses, so
'   the verification logic is not double-managed (same positioning as Debug_NormalizeUrl).
' ============================================================
Public Function Debug_SearchEngine() As String
    Debug_SearchEngine = m_searchTemplate
End Function



' ============================================================
' ResizeAll ? reissue the Bounds of all tabs to the given rectangle
'   Added in stage 9.10c-1. Called from UserForm_Resize in a single call (issue 8 a).
'
'   Why Bounds are reissued to invisible tabs too: so that the size does not
'   go off when they are activated later. All tabs are updated regardless of visible/invisible.
' ============================================================
Public Sub ResizeAll(ByRef rc As RECT)
    ' --- 9.11c: cache the content rectangle applied by the upper level (option A) ---
    '   Subsequent AddTabs use this rectangle as the initial Bounds and do not overlap the tab bar.
    m_lastContentRect = rc
    m_hasContentRect = True

    Dim i As Long
    Dim pane As Wv2Pane
    Dim hr As Long
    For i = 1 To m_tabs.Count
        Set pane = m_tabs.Item(i)
        If pane.IsReady Then
            hr = pane.Ctrl_PutBounds(rc)
            If hr <> 0 Then
                Debug.Print "Wv2Browser.ResizeAll: PutBounds failed for tab " & i & " hr=&H" & Hex(hr)
            End If
        End If
    Next i
End Sub


' ============================================================
' PaneAt ? get a tab by index (1-based) (for verification)
' ============================================================
Public Property Get PaneAt(ByVal index As Long) As Wv2Pane
    If index < 1 Or index > m_tabs.Count Then
        Set PaneAt = Nothing
        Exit Property
    End If
    Set PaneAt = m_tabs.Item(index)
End Property


' ============================================================
' state ? current Browser state
' ============================================================
Public Property Get state() As BrowserState
    state = m_state
End Property


' ============================================================
' Helper for waiting on asynchronous completion (private, confines the wait logic to this class)
' ============================================================
Private Function WaitEnvReady() As Boolean
    Dim t As Single
    t = Timer
    Do While (Not m_env.IsReady) And (Not m_env.IsFailed) And (Wv2Log.SecondsSince(t) < READY_TIMEOUT)
        DoEvents
    Loop
    WaitEnvReady = m_env.IsReady
End Function

Private Function WaitPaneReady(ByVal pane As Wv2Pane) As Boolean
    Dim t As Single
    t = Timer
    Do While (Not pane.IsReady) And (Not pane.IsFailed) And (Wv2Log.SecondsSince(t) < READY_TIMEOUT)
        DoEvents
    Loop
    WaitPaneReady = pane.IsReady
End Function


' ============================================================
' RequestNewTabAsync (9.10e, option beta: deferred new-tab creation for NewWindowRequested)
'
'   Called from Wv2Pane.View_OnNewWindowRequested. Calling AddTabWithUrl directly here would
'   deadlock (spec fact 14), so the URL is pushed onto the pending queue, a one-shot
'   SafeTimer is started, and the real processing is done in m_newTabTimer_TimerCall after the handler completes.
'
'   *Why a New SafeTimer per request* SafeTimer.StartTimer
'   "ignores if uElapse = m_uElapse", and m_uElapse remains even after stopping (StartTimer 0).
'   If the same instance is reused, a restart with the same interval is ignored, so for every request
'   we create a new instance (m_uElapse=0) to make sure it starts. With consecutive requests, even if the previous
'   timer is discarded all URLs remain in the pending queue, so none are missed.
' ============================================================
Public Sub RequestNewTabAsync(ByVal url As String)
    If m_pendingUrls Is Nothing Then Set m_pendingUrls = New Collection
    m_pendingUrls.Add url

    Set m_newTabTimer = New SafeTimer      ' new instance per request (starts reliably)
    m_newTabTimer.StartTimer 1             ' TimerCall after 1ms (after the handler completes)

    Debug.Print "Wv2Browser.RequestNewTabAsync: pending URL added (" & url & _
                "). Queue count=" & m_pendingUrls.Count & ", one-shot Timer started"
End Sub


' ============================================================
' RequestSettingsTabAsync (9.18, entry for deferred creation of the settings tab)
'
'   Called from the TabBar [gear] button -> OnPaneWebMessage "settings". The process of opening the settings tab
'   goes internally through AddTab -> WaitPaneReady, so calling it synchronously from inside a WebMessage handler
'   falls into the deadlock of spec fact 14. So, of the same type as newtab,
'   it sets the pending flag, starts a one-shot SafeTimer and returns immediately. After the handler completes,
'   the actual settings-tab creation runs in m_settingsTimer_TimerCall.
'
'   *The reason for a New SafeTimer per request is the same as RequestNewTabAsync*
'     StartTimer "ignores a restart with the same interval" and m_uElapse remains after stopping,
'     so reuse would get the restart ignored. A new instance per request starts reliably.
'   *Multiple requests are harmless* Even when already pending, it just re-arms a new timer. The TimerCall side
'     looks at the flag and opens only once (a minimal guard so mashing does not make settings tabs proliferate).
' ============================================================
Public Sub RequestSettingsTabAsync()
    m_settingsPending = True

    Set m_settingsTimer = New SafeTimer    ' new instance per request (starts reliably)
    m_settingsTimer.StartTimer 1           ' TimerCall after 1ms (after the handler completes)

    Debug.Print "Wv2Browser.RequestSettingsTabAsync: settings-tab creation reserved (one-shot Timer started)"
End Sub


' ============================================================
' m_settingsTimer_TimerCall (9.18, receiver for the one-shot firing of settings-tab creation)
'
'   Called at a time when the WebMessage handler has already returned. If
'   OpenSettingsTab is called here, AddTab's WaitPaneReady completes normally
'   (same structure as the TimerCall of newtab).
' ============================================================
Private Sub m_settingsTimer_TimerCall(ByVal wTime As Long, ByRef varData As Variant)
    ' Make it one-shot: stop repeated firing
    If Not m_settingsTimer Is Nothing Then m_settingsTimer.StartTimer 0

    If Not m_settingsPending Then Exit Sub
    m_settingsPending = False

    Debug.Print "Wv2Browser.TimerCall(settings): opening the settings tab"
    OpenSettingsTab
End Sub


' ============================================================
' IsSettingsView (K-2)
'
'   Ask the page itself whether that Pane is *now* showing the settings screen.
'   window.__wv2Settings = true is embedded in the settings screen's HTML.
'
'   *The way to read the return value is the same style as Elem_CheckStale* Have it return an object and
'     read it with InStr. What EvalSync returns is the JSON text of the value, so
'     this is less prone to mix-ups than comparing the boolean directly.
'
'   *When it cannot be determined, False (the safe side)* Loading, already Closed, and in-callback are
'     all treated as "not the settings screen". Even if missed, a new settings tab merely opens, and that does
'     less harm than wrongly grabbing a tab of another page (distinguished by LastEvalOk of design principle 93).
' ============================================================
Private Function IsSettingsView(ByVal pane As Wv2Pane) As Boolean
    If pane Is Nothing Then Exit Function
    ' English fork: only a pane we opened as the settings tab qualifies. The JS flag
    ' alone is page-controlled, so any site could claim to be the settings view.
    If Not pane.IsSettings Then Exit Function

    Dim js As String
    Dim res As String
    js = "(function(){try{return {s:(window.__wv2Settings===true)?'yes':'no'};}" & _
         "catch(e){return {s:'no'};}})()"

    res = pane.EvalSync(js, 2)

    If Not pane.LastEvalOk Then
        Wv2Log.LogD "Wv2Browser.IsSettingsView: cannot determine, so treating as not the settings screen (" & _
                    pane.LastEvalError & ")"
        Exit Function
    End If

    IsSettingsView = (InStr(1, res, """s"":""yes""") > 0)
End Function


' ============================================================
' RequestSettingsSyncAsync / m_syncTimer_TimerCall (K-2)
'
'   *Why wait a beat* SetEngine / SetCustomEngine are called via hostObjects
'   (JS->VBA). Per spec fact 27, during that time *the calling pane's JS is
'   completely blocked*. Firing EvalSync from here gets no completion notification back and
'   hangs until timeout. Fire it after getting out of the handler with a SafeTimer
'   (same type as RequestSettingsTabAsync / RequestNewTabAsync).
' ============================================================
Public Sub RequestSettingsSyncAsync()
    m_syncPending = True

    Set m_syncTimer = New SafeTimer
    m_syncTimer.StartTimer 1
End Sub


Private Sub m_syncTimer_TimerCall(ByVal wTime As Long, ByRef varData As Variant)
    ' Make it one-shot: stop repeated firing
    If Not m_syncTimer Is Nothing Then m_syncTimer.StartTimer 0

    If Not m_syncPending Then Exit Sub
    m_syncPending = False

    SyncSettingsViewsNow
End Sub


' ============================================================
' SyncSettingsViewsNow (K-2, global sync of the settings views)
'
'   Push the current engine name to all tabs showing the settings screen.
'   *Design principle 60 (VBA->JS is a global sync)* Do not send diffs; pass the current value as is.
'
'   This is needed because even with an accurate check, two settings views can end up side by side.
'   Back navigation does not go through OpenSettingsTab, so it cannot be blocked by the check at open time.
'   The role of this sync is to keep them from disagreeing even when side by side.
'
'   *Do not overwrite the contents of the custom input field* Because the user may be in the middle of typing.
'     Only the selection state and the preview are aligned.
' ============================================================
Private Sub SyncSettingsViewsNow()
    If m_tabs Is Nothing Then Exit Sub
    If m_tabs.Count = 0 Then Exit Sub

    ' If there is no bridge, the settings screen has never been opened
    If m_settingsBridge Is Nothing Then Exit Sub

    Dim syncName As String
    syncName = m_settingsBridge.GetEngine()
    If Len(syncName) = 0 Then Exit Sub

    Dim syncIdx As Long
    Dim syncPane As Wv2Pane
    Dim syncJs As String
    Dim res As String
    Dim hit As Long

    For syncIdx = 1 To m_tabs.Count
        Set syncPane = m_tabs.Item(syncIdx)
        If IsSettingsView(syncPane) Then
            syncJs = "(function(){try{return {r:applyFromHost(" & _
                     syncPane.JsQuote(syncName) & ")};}catch(e){return {r:'err'};}})()"
            res = syncPane.EvalSync(syncJs, 2)
            If syncPane.LastEvalOk Then
                If InStr(1, res, """r"":""ok""") > 0 Then
                    hit = hit + 1
                Else
                    Wv2Log.LogW "Wv2Browser.SyncSettingsViewsNow: index " & syncIdx & _
                                " failed to apply (" & res & ")"
                End If
            Else
                Wv2Log.LogW "Wv2Browser.SyncSettingsViewsNow: index " & syncIdx & _
                            " was not reached (" & syncPane.LastEvalError & ")"
            End If
        End If
    Next syncIdx

    If hit > 0 Then
        Wv2Log.LogD "Wv2Browser.SyncSettingsViewsNow: settings view count " & hit & _
                    " -> reflected " & syncName & " to them"
    End If
End Sub

' ============================================================
' OpenSettingsTab (9.18, open the settings tab and attach the bridge)
'
'   Steps (order A: attach, then Navigate):
'     1. Create an empty tab (Pane + Controller) with AddTab.
'     2. Prepare Wv2SettingsBridge and tie ourselves (the Browser) to it with BindBrowser.
'     3. Attach it to that Pane with View_AddHostObjectToScript("Settings", bridge).
'        *Attach before NavigateToString* The hypothesis (order A) is that the loaded JS
'        can use hostObjects.Settings from the start.
'     4. Build the settings-screen HTML with BuildSettingsHtml and display it with View_NavigateToString.
'     5. Activate the created tab (ActivateTab).
'
'   *Holding the bridge* The attached bridge is referenced by WebView2 while the settings tab
'     is alive. The Browser also holds one in m_settingsBridge to keep it surely alive
'     (with only a local variable it could be released the instant this Sub exits). If several settings tabs
'     are opened it is overwritten by the last bridge, but each tab's WebView2 keeps the reference
'     from its own attach time, so the behavior of existing tabs is not affected.
' ============================================================
Public Sub OpenSettingsTab()
    ' --- 9.19: if a settings tab already exists, just activate it ---
    '   Scan m_tabs from the front for a Pane with IsSettings = True. If found,
    '   activate that index and Exit Sub (no new creation). A closed
    '   settings tab has left m_tabs, so it is naturally not found by this scan.
    ' --- K-2: not the flag but *ask the page itself* ---
    '   IsSettings did not go back to False after navigation and grabbed the former settings tab.
    '   If we ask for window.__wv2Settings with EvalSync, regardless of the navigation path we can tell
    '   "whether that tab is the settings screen right now" (a tab that returned via back is also picked up).
    Dim j As Long
    For j = 1 To m_tabs.Count
        If IsSettingsView(m_tabs.Item(j)) Then
            Wv2Log.LogD "Wv2Browser.OpenSettingsTab: found a tab (index " & j & _
                        ") showing the settings screen, so activating it"
            ActivateTab j
            Exit Sub
        End If
    Next j

    Dim pane As Wv2Pane
    Set pane = AddTab()
    If pane Is Nothing Then
        Debug.Print "Wv2Browser.OpenSettingsTab: AddTab failed"
        Exit Sub
    End If

    ' --- 9.19: mark the created Pane as a settings tab ---
    '   Subsequent [gear] clicks look at this flag to prevent duplicates.
    pane.IsSettings = True

    ' Prepare the bridge and tie the Browser to it
    Set m_settingsBridge = New Wv2SettingsBridge
    m_settingsBridge.BindBrowser Me

    ' *Order A: attach before Navigate*
    Dim hr As Long
    hr = pane.View_AddHostObjectToScript("Settings", m_settingsBridge)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.OpenSettingsTab: AddHostObjectToScript failed hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Browser.OpenSettingsTab: Settings bridge attached OK"
    End If

    ' Display the settings-screen HTML
    hr = pane.View_NavigateToString(BuildSettingsHtml())
    If hr <> 0 Then
        Debug.Print "Wv2Browser.OpenSettingsTab: NavigateToString failed hr=&H" & Hex(hr)
    End If

    ' Activate the created tab (reverse-look-up the new index)
    Dim i As Long
    For i = 1 To m_tabs.Count
        If m_tabs.Item(i) Is pane Then
            ActivateTab i
            Exit For
        End If
    Next i
End Sub


' ============================================================
' BuildSettingsHtml (9.18, build the settings-screen HTML)
'
'   A card UI for choosing the search engine. Calls VBA directly via hostObjects.sync.Settings
'   (postMessage commands are not used):
'     - At startup      : get the card definitions with Settings.EngineList() and draw.
'                         Get the current engine with Settings.GetEngine() and highlight the selection.
'     - Card click      : switch with Settings.SetEngine(id). The return value confirms the selection.
'     - Card hover      : Settings.PreviewUrlFor(id, sample word) shows the generated URL
'                         live (the charm of synchronous calls).
'
'   *Text is written as actual characters (CRLF/UTF-8/no BOM is guaranteed by Python at output time)*
'   *hostObjects.sync is the synchronous host-object call of WebView2*
'     The return value becomes the JS return value as is (confirmed in the old implementation).
' ============================================================
Private Function BuildSettingsHtml() As String
    Dim s As String
    s = "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    s = s & "<title>Settings</title>" & vbLf
    s = s & "<style>" & vbLf
    s = s & "  *{box-sizing:border-box;margin:0;padding:0;}" & vbLf
    s = s & "  html,body{height:100%;font-family:'Segoe UI','Meiryo',sans-serif;" & _
            "color:#e8eaed;background:radial-gradient(1200px 600px at 20% -10%,#2b3550 0%,#151a26 55%,#0e121b 100%);}" & vbLf
    s = s & "  .wrap{max-width:840px;margin:0 auto;padding:48px 28px 64px;}" & vbLf
    s = s & "  .eyebrow{letter-spacing:.22em;font-size:11px;color:#8ea2c8;text-transform:uppercase;}" & vbLf
    s = s & "  h1{font-size:26px;font-weight:650;margin:6px 0 4px;}" & vbLf
    s = s & "  .lead{color:#9aa7bd;font-size:13.5px;margin-bottom:28px;}" & vbLf
    s = s & "  .grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:16px;}" & vbLf
    s = s & "  .card{position:relative;border:1px solid rgba(255,255,255,.08);border-radius:14px;" & _
            "padding:18px 18px 16px;cursor:pointer;background:rgba(255,255,255,.03);" & _
            "transition:transform .12s ease,border-color .12s ease,background .12s ease;overflow:hidden;}" & vbLf
    s = s & "  .card:hover{transform:translateY(-2px);border-color:rgba(120,160,255,.5);background:rgba(255,255,255,.06);}" & vbLf
    s = s & "  .card.sel{border-color:#6ea8fe;background:linear-gradient(180deg,rgba(110,168,254,.16),rgba(110,168,254,.05));" & _
            "box-shadow:0 0 0 1px rgba(110,168,254,.35),0 10px 30px -12px rgba(110,168,254,.5);}" & vbLf
    s = s & "  .card .dot{position:absolute;top:16px;right:16px;width:16px;height:16px;border-radius:50%;" & _
            "border:2px solid rgba(255,255,255,.3);}" & vbLf
    s = s & "  .card.sel .dot{border-color:#6ea8fe;background:#6ea8fe;box-shadow:0 0 10px #6ea8fe;}" & vbLf
    s = s & "  .card .name{font-size:17px;font-weight:620;margin-bottom:2px;}" & vbLf
    s = s & "  .card .host{font-size:11.5px;color:#8ea2c8;word-break:break-all;}" & vbLf
    s = s & "  .badge{display:inline-block;margin-top:10px;font-size:10.5px;letter-spacing:.08em;" & _
            "color:#6ea8fe;border:1px solid rgba(110,168,254,.4);border-radius:999px;padding:2px 9px;opacity:0;" & _
            "transition:opacity .12s ease;}" & vbLf
    s = s & "  .card.sel .badge{opacity:1;}" & vbLf
    s = s & "  .preview{margin-top:26px;border-top:1px solid rgba(255,255,255,.08);padding-top:16px;}" & vbLf
    s = s & "  .preview .plabel{font-size:11px;letter-spacing:.14em;color:#8ea2c8;text-transform:uppercase;margin-bottom:6px;}" & vbLf
    s = s & "  .preview .purl{font-family:'Consolas','Courier New',monospace;font-size:12.5px;color:#cfe0ff;" & _
            "background:rgba(0,0,0,.25);border-radius:8px;padding:10px 12px;word-break:break-all;min-height:38px;}" & vbLf
    s = s & "  .sample{color:#7d8aa0;}" & vbLf
    ' --- 9.29: custom input field (an arbitrary template outside the presets) ---
    s = s & "  .custom{margin-top:22px;border:1px solid rgba(255,255,255,.08);border-radius:14px;" & _
            "padding:16px 18px;background:rgba(255,255,255,.03);}" & vbLf
    s = s & "  .custom.sel{border-color:#6ea8fe;" & _
            "background:linear-gradient(180deg,rgba(110,168,254,.16),rgba(110,168,254,.05));" & _
            "box-shadow:0 0 0 1px rgba(110,168,254,.35);}" & vbLf
    s = s & "  .chead{display:flex;align-items:center;gap:10px;}" & vbLf
    s = s & "  .cname{font-size:15px;font-weight:620;}" & vbLf
    s = s & "  .cbadge{font-size:10.5px;letter-spacing:.08em;color:#6ea8fe;" & _
            "border:1px solid rgba(110,168,254,.4);border-radius:999px;padding:2px 9px;" & _
            "opacity:0;transition:opacity .12s ease;}" & vbLf
    s = s & "  .custom.sel .cbadge{opacity:1;}" & vbLf
    s = s & "  .chint{font-size:11.5px;color:#8ea2c8;margin:6px 0 12px;}" & vbLf
    s = s & "  .crow{display:flex;gap:10px;}" & vbLf
    s = s & "  .crow input{flex:1;min-width:0;font-family:'Consolas','Courier New',monospace;" & _
            "font-size:12.5px;color:#e8eaed;background:rgba(0,0,0,.25);" & _
            "border:1px solid rgba(255,255,255,.12);border-radius:8px;padding:9px 11px;outline:none;}" & vbLf
    s = s & "  .crow input:focus{border-color:#6ea8fe;}" & vbLf
    s = s & "  .crow button{font-size:13px;color:#0e121b;background:#6ea8fe;border:0;" & _
            "border-radius:8px;padding:9px 18px;cursor:pointer;font-weight:620;}" & vbLf
    s = s & "  .crow button:hover{background:#8fbcff;}" & vbLf
    s = s & "  .cmsg{margin-top:10px;font-family:'Consolas','Courier New',monospace;font-size:12px;" & _
            "color:#a9c6ff;word-break:break-all;min-height:16px;}" & vbLf
    s = s & "  .cmsg.err{color:#ff9d9d;}" & vbLf
    s = s & "</style></head><body>" & vbLf
    s = s & "<div class=""wrap"">" & vbLf
    s = s & "  <div class=""eyebrow"">Search Engine</div>" & vbLf
    s = s & "  <h1>Choose a search engine</h1>" & vbLf
    s = s & "  <div class=""lead"">Choose which search engine to send to when you enter something that is not a URL in the address bar. " & _
            "Hover the mouse over a card to check the generated search URL on the spot.</div>" & vbLf
    s = s & "  <div class=""grid"" id=""grid""></div>" & vbLf
    ' --- 9.29: custom input field. Do not use a form (because Enter would submit) ---
    s = s & "  <div class=""custom"" id=""custom"">" & vbLf
    s = s & "    <div class=""chead""><span class=""cname"">Custom (other than presets)</span>" & _
            "<span class=""cbadge"">In use</span></div>" & vbLf
    s = s & "    <div class=""chint"">Enter a URL to which the search term is appended at the end. " & _
            "Example: https://example.com/search?q=</div>" & vbLf
    s = s & "    <div class=""crow"">" & vbLf
    s = s & "      <input id=""ctpl"" type=""text"" spellcheck=""false"" " & _
            "placeholder=""https://example.com/search?q="">" & vbLf
    s = s & "      <button id=""capply"">Apply</button>" & vbLf
    s = s & "    </div>" & vbLf
    s = s & "    <div class=""cmsg"" id=""cmsg""></div>" & vbLf
    s = s & "  </div>" & vbLf
    s = s & "  <div class=""preview"">" & vbLf
    s = s & "    <div class=""plabel"">Preview</div>" & vbLf
    s = s & "    <div class=""purl"" id=""purl""><span class=""sample"">Shows a preview for the current engine</span></div>" & vbLf
    s = s & "  </div>" & vbLf
    s = s & "</div>" & vbLf
    s = s & "<script>" & vbLf
    ' --- Reference to hostObjects.sync.Settings. Guarded so it does not fail in environments without it ---
    s = s & "  var S = (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects" & _
            " && window.chrome.webview.hostObjects.sync) ? window.chrome.webview.hostObjects.sync.Settings : null;" & vbLf
    s = s & "  var SAMPLE = 'excel vba';" & vbLf
    s = s & "  var current = '';" & vbLf
    ' --- 9.19c: id of the card the cursor is currently over (null when not hovered) ---
    s = s & "  var hoverId = null;" & vbLf
    ' --- K-2: a marker that this page is the settings screen ---
    '   OpenSettingsTab asks for it with EvalSync. *Put it early*
    '   so that even if later JS fails, the page can still declare itself the settings screen
    '   (if it cannot, the gear opens a second settings tab).
    s = s & "  window.__wv2Settings = true;" & vbLf
    ' --- Synchronously get the current engine and card definitions from VBA ---
    s = s & "  function loadState(){" & vbLf
    s = s & "    if(!S){ document.getElementById('grid').textContent='Host object not connected'; return null; }" & vbLf
    s = s & "    try{ current = S.GetEngine(); }catch(e){ current=''; }" & vbLf
    s = s & "    try{ return JSON.parse(S.EngineList()); }catch(e){ return []; }" & vbLf
    s = s & "  }" & vbLf
    ' --- Extract the host name from the URL template (for appearance) ---
    s = s & "  function hostOf(tpl){ try{ return (tpl.split('/')[2]||tpl); }catch(e){ return tpl; } }" & vbLf
    s = s & "  function markSel(){" & vbLf
    s = s & "    var cs=document.querySelectorAll('.card');" & vbLf
    s = s & "    for(var i=0;i<cs.length;i++){ cs[i].classList.toggle('sel', cs[i].getAttribute('data-id')===current); }" & vbLf
    ' --- 9.29: update the selection highlight of the custom section from the same single place ---
    s = s & "    markCustom();" & vbLf
    s = s & "  }" & vbLf
    s = s & "  function showPreview(id){" & vbLf
    s = s & "    var el=document.getElementById('purl'); if(!S){ return; }" & vbLf
    s = s & "    var u=''; try{ u=S.PreviewUrlFor(id, SAMPLE); }catch(e){ u=''; }" & vbLf
    s = s & "    el.innerHTML = '<span class=""sample"">' + SAMPLE + '</span> &nbsp;&rarr;&nbsp; ' + (u||'(failed to get)');" & vbLf
    s = s & "  }" & vbLf
    ' --- 9.19b: show the preview of the currently selected engine (default when not hovered / after leaving) ---
    s = s & "  function showCurrentPreview(){" & vbLf
    s = s & "    if(!S){ return; }" & vbLf
    s = s & "    if(current){ showPreview(current); }" & vbLf
    s = s & "  }" & vbLf
    s = s & "  function build(list){" & vbLf
    s = s & "    var g=document.getElementById('grid'); g.innerHTML='';" & vbLf
    s = s & "    for(var i=0;i<list.length;i++){ (function(item){" & vbLf
    s = s & "      var c=document.createElement('div'); c.className='card'; c.setAttribute('data-id',item.id);" & vbLf
    s = s & "      c.innerHTML='<span class=""dot""></span>' +" & vbLf
    s = s & "        '<div class=""name"">'+item.label+'</div>' +" & vbLf
    s = s & "        '<div class=""host"">'+hostOf(item.template)+'</div>' +" & vbLf
    s = s & "        '<div class=""badge"">Selected</div>';" & vbLf
    ' --- 9.19c: while hovering, set hoverId and show that card's preview ---
    s = s & "      c.onmouseenter=function(){ hoverId=item.id; showPreview(item.id); };" & vbLf
    ' --- 9.19c: on leaving, return to the selected one only if the cursor has not moved onto another card ---
    s = s & "      c.onmouseleave=function(){ hoverId=null; setTimeout(function(){ if(hoverId===null){ showCurrentPreview(); } },0); };" & vbLf
    s = s & "      c.onclick=function(){" & vbLf
    s = s & "        if(!S) return;" & vbLf
    s = s & "        var res=''; try{ res=S.SetEngine(item.id); }catch(e){ res=''; }" & vbLf
    s = s & "        if(res){ current=res; } else { current=item.id; }" & vbLf
    s = s & "        markSel(); showPreview(item.id);" & vbLf
    s = s & "      };" & vbLf
    s = s & "      g.appendChild(c);" & vbLf
    s = s & "    })(list[i]); }" & vbLf
    s = s & "    markSel();" & vbLf
    ' --- 9.19c: right after opening (not yet hovered), also show the preview of the selected engine ---
    '   (The return when leaving a card is handled by each card's mouseleave. The mouseleave of the whole #grid
    '    is not used, because it did not fire in the gaps between cards and was counterintuitive.)
    s = s & "    showCurrentPreview();" & vbLf
    s = s & "  }" & vbLf
    ' --- 9.29: JS of the custom input field (attribute values and strings use single quotes) ---
    s = s & "  function markCustom(){" & vbLf
    s = s & "    var box=document.getElementById('custom'); if(!box){ return; }" & vbLf
    s = s & "    box.classList.toggle('sel', current==='custom');" & vbLf
    s = s & "  }" & vbLf
    s = s & "  function setMsg(text, isErr){" & vbLf
    s = s & "    var m=document.getElementById('cmsg'); if(!m){ return; }" & vbLf
    s = s & "    m.textContent = text || '';" & vbLf
    s = s & "    m.className = isErr ? 'cmsg err' : 'cmsg';" & vbLf
    s = s & "  }" & vbLf
    ' --- Live preview before applying. Encoding is left to VBA (same as production) ---
    s = s & "  function customPreview(){" & vbLf
    s = s & "    var ct=document.getElementById('ctpl'); if(!ct||!S){ return; }" & vbLf
    s = s & "    var v=ct.value;" & vbLf
    s = s & "    if(!v){ setMsg('', false); return; }" & vbLf
    s = s & "    var u=''; try{ u=S.PreviewUrlForTemplate(v, SAMPLE); }catch(e){ u=''; }" & vbLf
    s = s & "    if(u){ setMsg(SAMPLE + '  ->  ' + u, false); }" & vbLf
    s = s & "    else { setMsg('Enter a URL that starts with http:// or https://', true); }" & vbLf
    s = s & "  }" & vbLf
    s = s & "  function applyCustom(){" & vbLf
    s = s & "    var ct=document.getElementById('ctpl'); if(!ct||!S){ return; }" & vbLf
    s = s & "    var res=''; try{ res=S.SetCustomEngine(ct.value); }catch(e){ res=''; }" & vbLf
    s = s & "    if(!res){ setMsg('Could not apply. Enter a URL that starts with http:// or https://" & _
            "', true); return; }" & vbLf
    s = s & "    current=res; markSel(); showCurrentPreview();" & vbLf
    s = s & "    if(res==='custom'){ setMsg('Applied (custom)', false); }" & vbLf
    s = s & "    else { setMsg('The input was the same as the preset (' + res + '), so " & _
            "that one was selected', false); }" & vbLf
    s = s & "  }" & vbLf
    s = s & "  function initCustom(){" & vbLf
    s = s & "    var ct=document.getElementById('ctpl'); if(!ct){ return; }" & vbLf
    s = s & "    if(S){ try{ ct.value = S.GetCustomTemplate(); }catch(e){ } }" & vbLf
    s = s & "    ct.oninput = function(){ customPreview(); };" & vbLf
    s = s & "    ct.onkeydown = function(e){ if(e.key==='Enter'){ applyCustom(); } };" & vbLf
    s = s & "    var btn=document.getElementById('capply');" & vbLf
    s = s & "    if(btn){ btn.onclick=function(){ applyCustom(); }; }" & vbLf
    s = s & "    markCustom();" & vbLf
    s = s & "  }" & vbLf
    ' --- K-2: receiver for the global sync from the host ---
    '   When the engine changes in another settings tab, VBA calls this and passes the current value.
    '   The custom input field (ctpl) may be mid-typing, so it is not touched.
    s = s & "  function applyFromHost(name){" & vbLf
    s = s & "    if(!name){ return 'ng'; }" & vbLf
    s = s & "    current = name;" & vbLf
    s = s & "    markSel();" & vbLf
    s = s & "    showCurrentPreview();" & vbLf
    s = s & "    return 'ok';" & vbLf
    s = s & "  }" & vbLf
    s = s & "  var list=loadState(); if(list){ build(list); }" & vbLf
    s = s & "  initCustom();" & vbLf
    s = s & "</script>" & vbLf
    s = s & "</body></html>"
    BuildSettingsHtml = s
End Function


' ============================================================
' ============================================================
' Prewarm mechanism (9.27) -- keep one spare tab warmed
'
'   *Why warm it* If a new Controller is created inside the NewWindowRequested handler,
'     its completion (ControllerCompleted) is not delivered until the handler returns, so it deadlocks (spec fact 14).
'     Using a Deferral avoids this, but it needs our own management of carrying the
'     args / deferral COM references out of the event.
'     *If it is made in advance, all that is left during the event is to fire put_NewWindow* -- that is the whole of
'     the prewarm approach. There is not a single "wait" left inside the event.
'
'   *The spare tab is not put in m_tabs* Therefore
'     - it does not appear on the tab bar (naturally, since TabAdded is not fired)
'     - it is outside the scope of TabCount / ActivateTab / CloseTab / MoveTab
'     - NotifyTitleChanged / NotifyUrlChanged pass through with "not found (ignored)"
'   Only at the moment of adoption (AdoptPrewarmedNow) does it enter m_tabs and become an ordinary tab.
'
'   *Responsibility for release* It is not in m_tabs, so the tab-release loop of Class_Terminate
'     does not pick it up. Shutdown explicitly does ClearBrowserRef -> Set Nothing (design principle 58).
' ============================================================

' --- Whether the spare tab is "usable right now" ---
'   4 conditions: created / Controller Ready / View obtained / not yet Claimed.
Public Property Get PrewarmAvailable() As Boolean
    If m_prewarmPane Is Nothing Then Exit Property
    If m_prewarmClaimed Then Exit Property
    If Not m_prewarmPane.IsReady Then Exit Property
    PrewarmAvailable = (m_prewarmPane.ViewPtr <> 0)
End Property


' ============================================================
' ClaimPrewarmViewPtr (9.27, take out the spare tab's View "with a reservation")
'
'   Called from Wv2Pane.View_OnNewWindowRequested. If there is a usable spare, returns its ViewPtr and
'   at the same time sets m_prewarmClaimed to *seal off double acquisition*.
'
'   *Why Claim is needed* If another NewWindowRequested arrives between firing put_NewWindow and the adoption Timer
'     running, two navigations would be loaded onto the same WebView.
'     If already Claimed, it returns 0 and makes the second one fall back to the current behavior.
'
'   Return value 0 = unusable (not created / not Ready / already Claimed).
' ============================================================
Public Function ClaimPrewarmViewPtr() As LongPtr
    If Not PrewarmAvailable Then
        Debug.Print "Wv2Browser.ClaimPrewarmViewPtr: spare tab unusable -> " & Debug_PrewarmStatus()
        Exit Function
    End If
    ' *Do not read the function name as a value* In VBA, writing the function name on the right-hand side inside the Function
    '   becomes a recursive call. Take it into a temporary variable and then log it.
    Dim pv As LongPtr
    pv = m_prewarmPane.ViewPtr
    m_prewarmClaimed = True
    ClaimPrewarmViewPtr = pv
    Debug.Print "Wv2Browser.ClaimPrewarmViewPtr: spare tab reserved (ViewPtr=" & pv & ")"
End Function


' --- Cancel the Claim (rollback when put_NewWindow fails) ---
Public Sub ReleasePrewarmClaim()
    If Not m_prewarmClaimed Then Exit Sub
    m_prewarmClaimed = False
    Debug.Print "Wv2Browser.ReleasePrewarmClaim: reservation cancelled (spare tab is reusable)"
End Sub


' ============================================================
' RequestPrewarmAsync / m_prewarmTimer_TimerCall (9.27, reservation and firing of spare-tab creation)
'
'   *Why create it with a delay (issue 1 option b)* Right after Init is the busiest stretch, where Controller creation
'     is continuous: Environment -> first tab -> TabBar -> NavBar, and the window of spec fact 20
'     is open. We do not insert a 4th here; we warm it after things quiet down.
'
'   *Re-entrancy guard* WaitPaneReady of CreatePaneCore spins DoEvents, so during startup
'     this TimerCall may interrupt. If creation is in progress, *skip this time* (we do not
'     re-arm the timer -- to avoid the danger of replacing a running SafeTimer).
'     Even if skipped, EnsurePrewarmScheduled at the end of AddTab reserves again at the next opportunity, so
'     it is always warmed once the startup sequence finishes.
' ============================================================
Public Sub RequestPrewarmAsync()
    If Not m_prewarmPane Is Nothing Then
        ' *9.28: clearing up a misleading log* "A usable spare exists" and "already Claimed (awaiting adoption)"
        '   mean completely different things, yet in 9.27 both had the same wording.
        '   *In neither case do we reserve (behavior is identical to 9.27)* If already Claimed, the adoption Timer
        '   after 1ms always replenishes in the order AdoptPrewarmedNow -> CreatePrewarmPaneNow.
        '   Setting another Timer here would only add one useless log line 1.5 seconds later saying
        '   "a spare tab already exists (skipping creation)", and would not make replenishment any earlier.
        If m_prewarmClaimed Then
            Debug.Print "Wv2Browser.RequestPrewarmAsync: spare tab already Claimed (awaiting adoption). " & _
                        "It is replenished automatically right after adoption, so no reservation here"
        Else
            Debug.Print "Wv2Browser.RequestPrewarmAsync: a usable spare tab already exists, so doing nothing"
        End If
        Exit Sub
    End If

    m_prewarmPending = True
    Set m_prewarmTimer = New SafeTimer     ' new instance per request (starts reliably)
    m_prewarmTimer.StartTimer PREWARM_DELAY_MS

    Debug.Print "Wv2Browser.RequestPrewarmAsync: spare-tab creation reserved (after " & _
                PREWARM_DELAY_MS & "ms)"
End Sub


' --- If there is no spare and nothing is reserved, reserve (called from the end of AddTab) ---
'   *9.29: do not set it when nested* While a spare tab is being created (CreatePrewarmPaneNow ->
'   while WaitPaneReady of CreatePaneCore is spinning DoEvents), the AddTab of the spare-exhausted
'   fallback may interrupt and come here. At that time
'   m_prewarmPane is still Nothing (the outer call has not assigned yet), so a reservation was set, and
'   1.5 seconds later a useless log "a spare tab already exists (skipping creation)" appeared.
'   Once the outer one completes there will be one spare, so there is no need to set it here.
Private Sub EnsurePrewarmScheduled()
    If Not m_prewarmPane Is Nothing Then Exit Sub
    If m_prewarmPending Then Exit Sub

    If m_creatingPane > 0 Then
        Debug.Print "Wv2Browser.EnsurePrewarmScheduled: a spare tab is being created (nested), so " & _
                    "not reserving (the outer completion will make one)"
        Exit Sub
    End If

    RequestPrewarmAsync
End Sub


Private Sub m_prewarmTimer_TimerCall(ByVal wTime As Long, ByRef varData As Variant)
    ' Make it one-shot: stop repeated firing
    If Not m_prewarmTimer Is Nothing Then m_prewarmTimer.StartTimer 0

    If Not m_prewarmPending Then Exit Sub

    ' *Re-entrancy guard* Skip if this interrupted during Pane creation (during DoEvents of WaitPaneReady)
    If m_creatingPane > 0 Then
        m_prewarmPending = False
        Debug.Print "Wv2Browser.TimerCall(prewarm): a Pane was being created, so skipping this time " & _
                    "(it will be reserved again when AddTab completes)"
        Exit Sub
    End If

    m_prewarmPending = False
    CreatePrewarmPaneNow
End Sub


' ============================================================
' CreatePrewarmPaneNow (9.27, actual creation of the spare tab)
'
'   Makes a Pane with CreatePaneCore and *hides it in the back with Ctrl_PutIsVisible False* (issue 3).
'   Bounds were already set by CreatePaneCore to the same rectangle as existing tabs, so at the moment of adoption
'   ActivateTab just makes it visible and it appears in the right place (does not break the unification of visibility control).
'
'   *Do not Navigate* The spare stays about:blank. After it is passed to put_NewWindow,
'     the runtime loads the original request (with the body if it is a POST) onto this WebView.
' ============================================================
Private Sub CreatePrewarmPaneNow()
    If Not m_prewarmPane Is Nothing Then
        Debug.Print "Wv2Browser.Prewarm: a spare tab already exists (skipping creation)"
        Exit Sub
    End If
    If m_state <> Bs_EnvReady Then
        Debug.Print "Wv2Browser.Prewarm: Environment not Ready, so not creating (state=" & m_state & ")"
        Exit Sub
    End If

    Debug.Print "Wv2Browser.Prewarm: starting spare-tab creation"
    Dim pane As Wv2Pane
    Set pane = CreatePaneCore()
    If pane Is Nothing Then
        Debug.Print "Wv2Browser.Prewarm: *spare-tab creation failed* (will be retried at the next AddTab / switch)"
        Exit Sub
    End If

    Dim hr As Long
    hr = pane.Ctrl_PutIsVisible(False)     ' hide in the back (so it is not seen as a tab)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.Prewarm: Ctrl_PutIsVisible(False) failed hr=&H" & Hex(hr)
    End If

    Set m_prewarmPane = pane
    m_prewarmClaimed = False
    Debug.Print "Wv2Browser.Prewarm: spare tab warmed (ViewPtr=" & pane.ViewPtr & _
                ", displayed tab count=" & m_tabs.Count & " <- excludes the spare)"
End Sub


' ============================================================
' RequestAdoptPrewarmedAsync / m_adoptTimer_TimerCall (9.27, adoption and replenishment)
'
'   Called from inside the event right after Wv2Pane fires put_NewWindow. The handler returns immediately, and
'   in the TimerCall afterwards
'     (1) adoption    : put the spare into m_tabs and make it an ordinary tab via TabAdded -> ActivateTab
'     (2) replenish   : make the next spare
'   are processed in this order (issue 6: adoption first. Replenishment is heavy because of waiting for Controller creation, so it is put off
'   and the user is shown the screen first).
'
'   *Waiting is allowed inside this TimerCall* The NewWindowRequested handler has already returned,
'     so WaitPaneReady of CreatePaneCore completes normally (the deadlock condition of
'     spec fact 14 is avoided). Same structure as m_newTabTimer_TimerCall.
' ============================================================
Public Sub RequestAdoptPrewarmedAsync(ByVal url As String)
    m_adoptPending = True
    m_adoptUrl = url

    Set m_adoptTimer = New SafeTimer      ' new instance per request (starts reliably)
    m_adoptTimer.StartTimer 1             ' TimerCall after 1ms (after the handler completes)

    Debug.Print "Wv2Browser.RequestAdoptPrewarmedAsync: spare-tab adoption reserved (" & url & ")"
End Sub


Private Sub m_adoptTimer_TimerCall(ByVal wTime As Long, ByRef varData As Variant)
    ' Make it one-shot: stop repeated firing
    If Not m_adoptTimer Is Nothing Then m_adoptTimer.StartTimer 0

    If Not m_adoptPending Then Exit Sub
    m_adoptPending = False

    AdoptPrewarmedNow          ' (1) adopt -- show the screen first
    CreatePrewarmPaneNow       ' (2) replenish -- warm the next spare (heavy, so later)
End Sub


' ============================================================
' AdoptPrewarmedNow (9.27, promote the spare tab to an ordinary tab)
'
'   *Let go of m_prewarmPane first* By making it Nothing before m_tabs.Add,
'     the replenishment right after (CreatePrewarmPaneNow) does not misjudge "a spare already exists".
'   *Leave making it visible to ActivateTab* The visible state hidden in CreatePrewarmPaneNow is
'     restored by ActivateTab. The visibility-control logic stays gathered in one place.
' ============================================================
Private Sub AdoptPrewarmedNow()
    If m_prewarmPane Is Nothing Then
        Debug.Print "Wv2Browser.Adopt: *no spare tab* (skipping adoption. put_NewWindow has" & _
                    " been fired, so the navigation is left hanging)"
        m_prewarmClaimed = False
        Exit Sub
    End If

    Dim pane As Wv2Pane
    Set pane = m_prewarmPane
    Set m_prewarmPane = Nothing
    m_prewarmClaimed = False

    m_tabs.Add pane
    Debug.Print "Wv2Browser.Adopt: spare tab adopted as a tab (current tab count=" & m_tabs.Count & _
                ", requested URL=" & m_adoptUrl & ")"

    RaiseEvent TabAdded(m_tabs.Count)
    ActivateTab m_tabs.Count
End Sub


' --- For verification: return the spare tab's state as a one-line string ---
Public Function Debug_PrewarmStatus() As String
    If m_prewarmPane Is Nothing Then
        Debug_PrewarmStatus = "no spare (pending=" & CStr(m_prewarmPending) & ")"
        Exit Function
    End If
    Debug_PrewarmStatus = "spare present ViewPtr=" & m_prewarmPane.ViewPtr & _
                          " / IsReady=" & CStr(m_prewarmPane.IsReady) & _
                          " / Claimed=" & CStr(m_prewarmClaimed) & _
                          " / usable=" & CStr(PrewarmAvailable)
End Function


' ============================================================
' m_newTabTimer_TimerCall (9.10e, receiver for the one-shot firing of SafeTimer)
'
'   Called at a time when the NewWindowRequested handler has already returned.
'   If AddTabWithUrl is done here, the new Controller's ControllerCompleted is delivered
'   normally, so it does not deadlock.
'
'   First stop repeated firing with StartTimer 0 (make it one-shot), then process the pending queue in order.
' ============================================================
Private Sub m_newTabTimer_TimerCall(ByVal wTime As Long, ByRef varData As Variant)
    ' Make it one-shot: stop repeated firing
    If Not m_newTabTimer Is Nothing Then m_newTabTimer.StartTimer 0

    If m_pendingUrls Is Nothing Then Exit Sub
    Debug.Print "Wv2Browser.TimerCall: processing pending URLs (queue count=" & m_pendingUrls.Count & ")"

    Do While m_pendingUrls.Count > 0
        Dim url As String
        url = m_pendingUrls(1)
        m_pendingUrls.Remove 1

        Dim pane As Wv2Pane
        Set pane = AddTabWithUrl(url)
        If pane Is Nothing Then
            Debug.Print "  RequestNewTab: AddTabWithUrl failed (" & url & ")"
        Else
            Debug.Print "  RequestNewTab: new tab created OK (" & url & ", ViewPtr=" & _
                        pane.ViewPtr & ")"
        End If
    Loop
End Sub


' ============================================================
' Shutdown (9.10e, break the circular reference and then release)
'
'   Since 9.10e Wv2Pane holds a back-reference (m_browser) to Wv2Browser,
'   so a circular reference Browser.m_tabs <-> Pane.m_browser arises.
'   VBA is reference-counting GC, so even if the UserForm simply does Set m_browser = Nothing,
'   neither is released because of the circular reference (Class_Terminate is not called either).
'
'   So this Shutdown is called explicitly from UserForm_Terminate, cutting each Pane's back-reference
'   (ClearBrowserRef) before release. This leaves only the one-way reference Browser->Pane,
'   and the following Set m_browser = Nothing lets Class_Terminate run normally
'   (release in the order all Panes -> Environment).
'
'   *The implementation cost on the UserForm side is only this one Shutdown call line* (the aim of option Y).
'   Idempotent (safe to call multiple times).
' ============================================================
' ============================================================
' QuietOnShutdown (K-4c) - whether to quiet all tabs before release (default True)
'
'   *Setting False reproduces the slow release* Used for isolating causes.
'   There is no reason to set False in actual operation.
' ============================================================
Public Property Get QuietOnShutdown() As Boolean
    QuietOnShutdown = m_quietOnShutdown
End Property

Public Property Let QuietOnShutdown(ByVal newValue As Boolean)
    m_quietOnShutdown = newValue
    Wv2Log.LogD "Wv2Browser.QuietOnShutdown = " & newValue
End Property


' ============================================================
' QuietAllPanes (K-4c, Private) - send all tabs to about:blank before release
'
'   *No waiting* It takes effect the moment we send it (0 seconds and 3 seconds gave the same result).
'   *Continue even on failure* Release cannot be stopped midway. View_Navigate rejects
'   m_pView = 0 by itself, so it does not become an error.
' ============================================================
Private Sub QuietAllPanes()
    Dim pane As Wv2Pane
    Dim n As Long
    Dim t0 As Single

    t0 = Timer

    ' *The spare tab is not in m_tabs* (9.27) so fire it separately
    If Not m_prewarmPane Is Nothing Then
        m_prewarmPane.View_Navigate "about:blank"
        n = n + 1
    End If

    If Not m_tabs Is Nothing Then
        For Each pane In m_tabs
            If Not pane Is Nothing Then
                pane.View_Navigate "about:blank"
                n = n + 1
            End If
        Next pane
    End If

    Wv2Log.LogD "Wv2Browser.QuietAllPanes: " & n & " sent to about:blank (" & _
                Format$(Timer - t0, "0.00") & " s)"
End Sub


Public Sub Shutdown()
    Wv2Log.LogD "Wv2Browser.Shutdown: *start*"   ' K-4a

    ' --- K-4c: *quiet all tabs before release* ---
    If m_quietOnShutdown Then QuietAllPanes

    ' 9.10e (Timer): stop and release the pending timer
    If Not m_newTabTimer Is Nothing Then
        m_newTabTimer.StartTimer 0
        Set m_newTabTimer = Nothing
    End If

    ' 9.18: release the settings-tab timer and bridge
    If Not m_settingsTimer Is Nothing Then
        m_settingsTimer.StartTimer 0
        Set m_settingsTimer = Nothing
    End If
    m_settingsPending = False
    Set m_settingsBridge = Nothing

    ' 9.27: stop the prewarm mechanism and release the spare tab
    '   *The spare is not in m_tabs* so it is not picked up by the all-tabs loop below.
    '   Unless the back-reference is cut here before setting it to Nothing, the circular reference prevents release (design principle 58).
    If Not m_prewarmTimer Is Nothing Then
        m_prewarmTimer.StartTimer 0
        Set m_prewarmTimer = Nothing
    End If
    If Not m_adoptTimer Is Nothing Then
        m_adoptTimer.StartTimer 0
        Set m_adoptTimer = Nothing
    End If
    m_prewarmPending = False
    m_adoptPending = False
    If Not m_prewarmPane Is Nothing Then
        m_prewarmPane.ClearBrowserRef
        Set m_prewarmPane = Nothing
        m_prewarmClaimed = False
        Wv2Log.LogD "Wv2Browser.Shutdown: releasing the spare tab (prewarm)"
    End If

    If m_tabs Is Nothing Then Exit Sub
    Dim pane As Wv2Pane
    For Each pane In m_tabs
        If Not pane Is Nothing Then pane.ClearBrowserRef
    Next pane
    Wv2Log.LogD "Wv2Browser.Shutdown: cut the back-references of all " & m_tabs.Count & " tabs"
End Sub


' ============================================================
' CloseAll (English fork)
'   Explicitly closes every tab's controller (and the prewarmed spare), then
'   runs Shutdown. Use it when you tear the browser down but keep the host
'   window alive (e.g. to start a new browser on the same form or hwnd).
'   Shutdown alone relies on the host window being destroyed to make WebView2
'   close the controllers; without that, the Browser <-> Pane <-> handler
'   cycles keep every WebView alive. Ctrl_Close costs ~1.4 s per pane (K-4),
'   which is why the normal form-close path does not call it.
' ============================================================
Public Sub CloseAll()
    Dim pane As Wv2Pane
    If Not m_tabs Is Nothing Then
        For Each pane In m_tabs
            If Not pane Is Nothing Then pane.Ctrl_Close
        Next pane
    End If
    If Not m_prewarmPane Is Nothing Then m_prewarmPane.Ctrl_Close
    Shutdown
End Sub


' ============================================================
' Class_Terminate ? release all tabs -> release the Environment (design principle 43: from child to parent)
' ============================================================
Private Sub Class_Terminate()
    Wv2Log.LogD "Wv2Browser.Terminate start (tab count=" & m_tabs.Count & ")"

    ' --- 0. 9.27: insurance release of the spare tab ---
    '   Normally Shutdown has already cleaned up. This is insurance for when we arrive here without going through Shutdown.
    If Not m_prewarmPane Is Nothing Then
        m_prewarmPane.ClearBrowserRef
        Set m_prewarmPane = Nothing
        Wv2Log.LogD "Wv2Browser.Terminate: releasing the spare tab (prewarm)"
    End If

    ' --- 1. Release all Panes (each Pane's Class_Terminate releases the View family in order) ---
    '   Removing from the Collection cuts the reference, and when the reference count reaches 0
    '   Wv2Pane.Class_Terminate runs. Remove from the end in order.
    Dim i As Long
    For i = m_tabs.Count To 1 Step -1
        m_tabs.Remove i
    Next i
    Set m_tabs = Nothing

    ' --- 2. Release the Environment (after all Panes = parent last) ---
    Set m_env = Nothing

    m_hwndHost = 0
    m_state = Bs_Created
    Wv2Log.LogD "Wv2Browser.Terminate done"
End Sub

