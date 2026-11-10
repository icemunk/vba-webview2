VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Wv2TabBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''
' --- Wv2TabBar.cls  stage K-4a (move the release-path logging to Wv2Log) ---
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
' --- Wv2TabBar.cls  stage 9.32b (wheel horizontal scroll + keeping the scroll position) ---
'
'   Scope of 9.32b (*JS only. Not one line of CSS or VBA is touched*):
'     (1) Attach a wheel handler to #tabs to convert vertical wheel into horizontal scroll
'     (2) Save and restore scrollLeft around render()
'     (3) If the active tab is off screen, bring it into the visible range (nearest)
'
'   *Background 1: no automatic conversion happened (spec fact 28)*
'     In 9.32 we kept overflow-x:auto and hid ::-webkit-scrollbar, expecting that
'     "Chromium converts the wheel to horizontal on a horizontal container that cannot
'     scroll vertically", but *on the real machine it did not react*. It depends on the ancestors' overflow
'     configuration and the hit target, so it cannot be relied on. An explicit wheel handler is the reliable way.
'
'   *Background 2: render() resets scrollLeft to 0*
'     render() runs tabsEl.innerHTML = '' at the start, so *every full sync
'     returns the horizontal scroll position to the left edge*. While there are few tabs no scrolling
'     happens at all so nobody noticed, but it becomes apparent once scrolling is possible.
'     Moreover PushSyncToJs is also sent *every time a title changes*, so while you are looking at the right side
'     the view jumps back to the left edge just because some tab finished loading.
'     On top of that, "a new tab added with + is off screen and cannot be seen" also happens.
'
'   Agreed points (all 3 as Claude recommended):
'     1 Wheel = *a wheel handler on #tabs* (point 4 of 9.32 is not withdrawn)
'                  because going back to a thin 3px bar brings back the problem of squashed tabs.
'                  The handler is attached to #tabs only once. *Rewriting innerHTML does not replace
'                  #tabs itself, so the listener survives*.
'                  Specify {passive:false} and call preventDefault.
'     2 Scroll amount = *add deltaY as is*. One notch is about 100px = a little over 1 tab.
'                  If it is not enough it can be tuned later with a single number.
'     3 Keeping position = *save/restore + bring the active tab in with nearest*
'                  Restore alone leaves a new tab off screen and invisible.
'                  scrollIntoView({block:'nearest',inline:'nearest'}) does
'                  *not move when the element is already visible*, so the side effects are minimal.
'                  The order is "restore -> nearest".
'
'   *Post-hoc report 1: added the condition "only when the active tab changed" to the nearest step*
'     With point 3 as it was, merely receiving a title-update sync while looking at the right side
'     would pull the view back to the active tab's position (contrary to the aim of point 3,
'     "minimize side effects of moving by itself"). So we keep lastActive (a single number) at top level
'     and scrollIntoView *only when active differs from last time*. Since it is a memory of one number,
'     it does not undermine the intent of the full-sync method (design principle 60).
'
'   *Post-hoc report 2: correcting the description in the 9.32 header*
'     The 9.32 block says "design principles 82 and spec facts 26, no change in either",
'     but as a result of real-machine testing *design principles 83 and 84 and spec facts 27 and 28 were added*.
'     The correct count at the completion of 9.32b is 84 design principles and 28 spec facts. The dev notes are authoritative.
'       Design principle 83: when something feels slow, suspect the "order" before "making it faster"
'       Design principle 84: if the next full sync always overwrites it, an advance display is not state keeping
'       Spec fact 27: while a hostObjects.sync call is in progress, the JS of the calling pane is completely
'                    blocked and cannot even receive messages sent with PostWebMessageAsString
'                    (the cause of the 9.32 latency)
'       Spec fact 28: in WebView2 (Chromium 151) a vertical wheel is not automatically converted to
'                    horizontal scroll on an overflow-x:auto horizontal container
'
'   *Not touched* All CSS / all VBA logic / the 5 Browser events /
'     the ready handshake / full sync / the SafeTimer path / the postMessage
'     fallback / drag handling / optimistic update (9.32) / Wv2Browser /
'     Wv2Pane / Wv2NavBar / UserForm1.
'

''''''''''''''''''''''''''''''''''
' --- Wv2TabBar.cls  stage 9.32 (feel: optimistic update on tab switch + tidying the scrollbar) ---
'
'   Scope of 9.32 (*JS and CSS only. Not one line of VBA logic is touched*):
'     (1) Remove the *perceived latency* of tab switching (optimistic update)
'     (2) Hide the horizontal scrollbar that appears inside the tab bar when tabs increase
'     (3) Activating on drag start is *not added* (keeping the status quo was decided explicitly)
'
'   *Background: the cause of the latency was structural (found in the 9.31 real-machine test)*
'     The previous call chain:
'       JS: tab.onclick
'         `-> ho.TabBar.HostActivate(n)   <- JS is completely blocked here
'               `-> VBA: HostActivate -> m_browser.ActivateTab   <- heavy (display switch)
'                     `-> RaiseEvent ActiveChanged
'                           `-> TabBar.m_browser_ActiveChanged -> PushSyncToJs
'                                 `-> get every tab's DocumentTitle via COM
'                                 `-> View_PostWebMessageAsString(json) <- send only
'               <-- HostActivate finally returns
'         `-> JS is unblocked -> only now is the message event processed
'               `-> render() wipes and regenerates the DOM -> *only here does the color first change*
'     hostObjects.sync is a synchronous RPC, so until VBA finishes everything, the tab bar's
'     JS does not advance a single instruction. The message PushSyncToJs sent cannot be received
'     either because JS is blocked. *The order "main body switch -> log -> tab bar" was a necessity of the design*.
'     The perceived delay was no accident.
'
'   *The order must not be swapped on the VBA side*
'     HostActivate runs in the middle of a synchronous call from JS, so even if PushSyncToJs
'     is called first inside it, JS is blocked and cannot receive it. Forcing it in enters the
'     deadlock territory of spec fact 14. *Swapping the order is correctly done on the JS side*
'
'   Agreed points (all 7 as Claude recommended):
'     1 Optimistic update = *adopt*. Right after the click, JS switches the active class
'                    itself and then calls the Host. The full sync that comes later from VBA
'                    always overwrites it, so *JS is not made to hold state*,
'                    and it does not conflict with design principle 60 (full sync). On failure it reverts at the next sync.
'     2 *Waiting for paint = double requestAnimationFrame* (the most important decision of this round)
'                    If we block on a synchronous COM call right after changing the DOM,
'                    there is a risk that *it freezes before being painted* and point 1 becomes meaningless.
'                    setTimeout(...,0) only guarantees "the next task" and does not
'                    guarantee painting. Double rAF is the standard technique that guarantees "after one frame has been drawn".
'                    The cost is only that switching is about 16ms later.
'     3 close     = *not adopted*. Closing removes an element from the DOM and rolling back on failure
'                    is complicated. Add only activate first and see its effect
'                    (the intent of design principle 82: keep it in a form where we can tell what worked).
'     4 Scrollbar = *hide completely* #tabs::-webkit-scrollbar{display:none}
'                    overflow-x:auto stays, so the mouse wheel can scroll horizontally
'                    (Chromium converts the wheel to horizontal on a horizontal container that cannot scroll vertically).
'                    The idea of leaving a thin 3px bar is rejected, because a line permanently sitting
'                    right under the tab bar is distracting in a different way.
'     5 Tab width = *keep min-width:80px*. Lowering it to 60px only increases
'                    "unreadably narrow tabs" and does not change the overflow situation itself.
'     6 Activation  = *do not activate (status quo)*
'       on D&D        - Reordering is "tidying the appearance", not "changing what is being viewed"
'                      - ActivateTab is heavy, so running it on every grab would worsen the feel
'                      - The original activates because it has a feature to pull a tab out into
'                        a separate window. This implementation lacks that feature, so there is no reason to follow
'     7 Verification = Test_9_32_Help (visual check).
'
'   *Post-hoc report: moved HO() to top level*
'     Previously HO() was declared inside the for loop in render() and happened to work
'     through hoisting (it existed only in render's scope). In 9.32 sendActivate() had to be
'     taken out of render, so HO() was also
'     *moved next to postCmd (top level)*. Behavior is the same but the scope becomes
'     straightforward. The declaration inside render has been deleted.
'
'   *Not touched* All VBA logic (the 3 Host methods / OnPaneWebMessage /
'     PushSyncToJs / Init / Resize / Shutdown) / the 5 Browser events /
'     the ready handshake / full sync / the SafeTimer path / the postMessage
'     fallback / drag handling / Wv2Browser / Wv2Pane / Wv2NavBar /
'     UserForm1. Design principles 82 and spec facts 26, no change in either.
'

''''''''''''''''''''''''''''''''''
' --- Wv2TabBar.cls  stage 9.31 (appearance: tab hover + show/hide control of the close button) ---
'
'   Scope of 9.31 (*CSS only. Not one line of JS or VBA logic is touched*):
'     (A-1) Lighten the background when the mouse is over a tab
'     (A-2) Show the close button "only when active or hovered" (equivalent to Chrome)
'
'   *The only change is inside the <style> of BuildTabBarHtml.* render() / drag handling /
'     Host methods / OnPaneWebMessage / Init / full sync are all unchanged.
'
'   Agreed points (all 8 as Claude recommended):
'     1 Method     = *CSS :hover only. JS is not touched*
'                    render() discards the DOM every time with innerHTML='', so if JS held the
'                    hover state it would always be lost at each sync. With CSS it needs
'                    zero state and does not conflict with the full-sync method (design principle 60).
'     2 Hover background = *inactive only* #e7e7e7 -> #efefef.
'                    The active one is already #fff, the brightest, so it is unchanged (same as Chrome).
'                    Writing .tab:not(.active):hover gives specificity 0,3,0, which beats
'                    .tab.active (0,2,0). *It does not depend on the order of writing*
'     3 Display condition = *active or hover* (equivalent to Chrome).
'                    The active tab is the one you want to close "right now", so it should be
'                    clickable before the mouse is even over it.
'     4 *How to hide = visibility:hidden* (the most important decision of this round)
'                    - With display:none the 18px + 6px margin disappears and the title width
'                      expands and shrinks, so the text jitters at every hover.
'                    - With opacity:0 the invisible button can still be clicked (misfires).
'                    - visibility:hidden disappears while keeping its place, and pointer
'                      events are disabled too.
'                    The cost is that "the width usable for the title is always 24px less",
'                    but with min-width:80px this was judged acceptable.
'     5 Transition = *background color only .12s*
'                    visibility is a discrete value so transition has no effect, and the close
'                    button switches instantly (this is the intended behavior).
'     6 Color of the x = *keep as is* (#888 / #ccc background on hover).
'                    It is more reliable to judge after seeing on the real machine how it looks with display control in.
'     7 While dragging = *add nothing*. Suppressing hover during .dragging is not done in advance
'                    (lesson from 9.27 point 7: with a preemptive correction we could never
'                    know whether it was needed). If it is unpleasant on the real machine, add it in the next stage.
'     8 Verification = Test_9_31_Help (visual procedure). At the same time we work off *open item 1 of 9.30 =
'                    the full run-through check of v0_5_2* (new-tab expansion of POST links /
'                    pre-warm / drag reordering) with the same procedure.
'
'   *Not touched* the 5 Browser events / ready handshake / full sync /
'     the 3 hostObjects methods (HostActivate/HostClose/HostReorder) / postMessage
'     fallback / SafeTimer path / 3-tier layout / Wv2Browser / Wv2Pane /
'     Wv2NavBar / UserForm1. Design principles 82 and spec facts 26, no change in either.
'

''''''''''''''''''''''''''''''''''
' --- Wv2TabBar.cls  stage 9.23 (tidying: consolidate activate/close/reorder handling into Host methods) ---
'
'   Scope of 9.23 (remove the duplicated handling of the 3 commands completed in 9.21/9.22. Plan C-1):
'     *Background* When 9.21/9.22 moved activate/close/reorder to the hostObjects path,
'       the same processing (calls to ActivateTab/CloseTab/MoveTab and the failure log) was
'       duplicated in 2 places: "inside the Cases of OnPaneWebMessage" and "inside HostActivate/HostClose/HostReorder".
'       A breeding ground for the accident of fixing one and forgetting to fix the other.
'
'     *Decision = Plan C (keep the insurance and consolidate the processing)* The JS-side hybrid of the 3 commands
'       (hostObjects first, postMessage fallback) is left as is. The fallback
'       is not discarded as the safety net for the browser's core operations (switch/close/reorder).
'       On top of that, the bodies of Case activate/close/reorder, which are passed through at fallback time, are
'       delegated to calls of the corresponding Host methods (HostActivate/HostClose/HostReorder). This
'       gathers the logic that calls ActivateTab/CloseTab/MoveTab and the failure log into
'       one place in the Host methods (DRY).
'
'     *Plan C-1 (fire-and-forget, not taking the return value)* The Case side calls HostXxx and discards the result. The failure
'       log is already emitted inside the Host method, so the double log on the Case side is removed. However, the Debug.Print
'       at the Case entrance (-> activate etc.) is kept as evidence that the postMessage path was taken.
'       The JSON parse (JsonGetNum) is needed to restore the arguments on the fallback path,
'       so it stays on the Case side (this is the correct form. On the hostObjects path JS passes
'       the numbers directly so no parsing is needed; they are restored from JSON only at fallback time).
'
'   *Not touched* the 3 cases ready/newtab/settings (postMessage is the right way) /
'     BuildTabBarHtml on the JS side (the hybrid structure is unchanged) / the bodies of the 3 Host methods /
'     the attach in Init / the 5 Browser events / full sync / ready hand-
'     shake / SafeTimer path / 3-tier layout. Design principles 71 and spec facts 20,
'     no change in either (existing principles = Plan 4-A / practice of DRY. No new spec fact).
'

''''''''''''''''''''''''''''''''''
' --- Wv2TabBar.cls  stage 9.22 (trial: hostObjects conversion of close / reorder) ---
'
'   Scope of 9.22 (a natural development of 9.21. Following activate, close/reorder are also
'     trial-migrated to direct hostObjects calls and compared side by side with the postMessage version):
'     *Aim* In 9.21 hostObjects on a real URL was established for activate alone (spec fact 19) and
'       separability was demonstrated. As the next step, close / reorder (both
'       synchronous operations with no completion wait = Plan 4-A) are extended horizontally to the hostObjects path.
'       With this the "3 commands that suit hostObjects naturally (activate/close/reorder)"
'       are complete, and in the next stage (9.23 planned) we can judge with data whether it is worth actually
'       shrinking the Wv2Json parser + the Select Case of OnPaneWebMessage.
'
'     *Not touched (to guarantee stability)* ready / newtab / settings stay
'       as postMessage. ready is a notification (not an RPC). newtab/settings go internally
'       through WaitPaneReady (waiting for Controller creation to complete), so a direct sync call
'       raises the deadlock danger of spec fact 14. -> These 3 are unchanged, as in 9.21.
'
'     *Entry points added* (all copied in the same form as HostActivate of 9.21)
'       - Public Function HostClose(index) As Boolean
'         Called directly from JS as hostObjects.sync.TabBar.HostClose(n).
'         The body is m_browser.CloseTab(index), same as the close case. CloseTab is only
'         Ctrl_Close with no completion wait (Plan 4-A), so a direct sync call is safe.
'       - Public Function HostReorder(fromIdx, toIdx) As Boolean
'         Called directly from JS as hostObjects.sync.TabBar.HostReorder(from, to).
'         The body is m_browser.MoveTab(fromIdx, toIdx), same as the reorder case.
'         MoveTab is only array manipulation + PutIsVisible with no completion wait (Plan 4-A).
'         *The VBA argument names are fromIdx/toIdx* (From has been reported to clash as a dialect.
'         The JS side calls with positional arguments so the argument names are invisible to JS and can be freely chosen).
'
'     *The JS side is a hybrid (the same three-tier structure as activate in 9.21)*
'       On the x click of close / the drop of reorder, if hostObjects.sync.TabBar can be
'       obtained, it calls HostClose/HostReorder as a direct sync call and leaves [host] in the console;
'       if the host call throws, it catches and falls to postCmd with [host-fail->postmsg];
'       if there is no hostObjects, it falls to the conventional postCmd with [postmsg]. The existing HO()
'       helper (9.21) is reused as is. The sending of newtab/settings/ready is unchanged.
'
'   *Not touched* the 5 Browser events / full sync (VBA->JS push stays
'     postMessage) / ready handshake / SafeTimer path (newtab/settings) / 3-tier
'     layout. Design principles 71 and spec facts 19, no change in either (within the reach of spec fact 19.
'     close/reorder also hold with the same real URL + attach before Navigate as activate).
'
' --- Wv2TabBar.cls  stage 9.21 (trial: hostObjects conversion of activate) ---
'
'   Scope of 9.21 (only activate of JS->VBA goes to a direct hostObjects call):
'     *Aim* Of the 6 TabBar commands (ready/activate/close/newtab/settings/
'       reorder), only activate is trial-migrated to the hostObjects path and
'       compared side by side with the existing postMessage version. close/reorder/newtab/settings/
'       ready are not touched at all and stay postMessage (to guarantee stability).
'
'     *Core of the verification = does spec fact 18 hold even with a Navigate to a real URL?*
'       Spec fact 18 was established in the 9.18 settings tab (NavigateToString / data: URL).
'       TabBar loads the real URL appassets.tabbar with View_Navigate.
'       We confirm on the real machine whether "if AddHostObjectToScript is done before Navigate, the JS
'       loaded afterwards can grab hostObjects from its first run" also works on a real URL.
'       -> If it works, it is backed up that the reach of spec fact 18 extends to real URLs too, and it becomes the foundation for
'         the horizontal extension to close/reorder and the spread to NavBar.
'
'     *Entry points added*
'       - Public Function HostActivate(index) As Boolean
'         Called directly from JS as hostObjects.sync.TabBar.HostActivate(n).
'         The body is m_browser.ActivateTab(index), same as the existing activate case.
'         ActivateTab is only Ctrl_PutIsVisible with no completion wait (Plan 4-A), so
'         even a JS->VBA sync direct call does not deadlock.
'         On success ActiveChanged->PushSyncToJs runs and the highlight follows (full sync unchanged).
'       - m_barPane.View_AddHostObjectToScript("TabBar", Me) before Navigate in Init.
'         Placed right after the mapping and right before Navigate (order A of spec fact 18).
'
'     *The JS side is a hybrid (with fallback)*
'       On a click on the tab body, if hostObjects.sync.TabBar can be obtained, it calls
'       HostActivate and leaves [host] in the console. If not, it falls to the conventional postCmd(activate)
'       and leaves [postmsg] (insurance for when order A fails. Same form as the null guard of the settings tab).
'       Sending of anything other than activate is unchanged.
'
'   *Not touched* the 5 Browser events / full sync (VBA->JS push stays
'     postMessage) / ready handshake / SafeTimer path / 3-tier layout /
'     the existing paths of close, reorder, newtab, settings. No increase or decrease in design principles and spec facts
'     (within the reach of spec fact 18).
'
' --- Wv2TabBar.cls  stage 9.18 ---
'
'   Additions in 9.18 (settings button -> settings tab):
'     *Added a [gear] (settings) button at the right end of the tab bar* (part 3)
'       Added a #settings element to BuildTabBarHtml (pushed to the right end with margin-left:auto),
'       a hover style, and onclick. On click it posts {"cmd":"settings"} via postMessage.
'     *Added a "settings" case to OnPaneWebMessage*
'       It calls Browser.RequestSettingsTabAsync. *It must not be opened synchronously*
'       Creating the settings tab also passes through WaitPaneReady internally, so, to avoid the same deadlock as newtab
'       (spec fact 14), it is run delayed via SafeTimer.
'     *TabBar itself knows nothing about the settings content (loose coupling kept)*
'       TabBar only conveys the intent "I want to open settings" to Browser. Actually
'       building the settings tab and attaching the bridge is Browser's responsibility. TabBar's
'       5-event subscription structure / full sync / ready handshake are unchanged.
'
' --- Wv2TabBar.cls  stage 9.15 ---
'
'   Additions in 9.15 (drag reordering of tabs):
'     *Grab a tab and reorder it with HTML5 Drag & Drop (Plan B-1)*
'       Give each tab div of BuildTabBarHtml draggable=true and wire up dragstart /
'       dragover / drop / dragend. A minimal implementation that only gives the dragged tab .dragging and the drop
'       target candidate the visual class .dragover (polish goes to candidate 2).
'
'     *On drop, postMessage {"cmd":"reorder","from":N,"to":M}*
'       from/to are 1-based. A reorder case is added to OnPaneWebMessage on the VBA side, and
'       Browser.MoveTab from, to is called synchronously (array manipulation only, no completion wait ->
'       safe like activate/close. No SafeTimer detour like newtab is needed).
'       When MoveTab succeeds, ActiveChanged -> PushSyncToJs runs (full sync,
'       design principle 60, Plan A-2b), and the tab bar is redrawn in the reordered order. The JS side
'       does not move the DOM optimistically; it waits for the sync from VBA and redraws (VBA is the source of truth).
'
' --- Wv2TabBar.cls  stage 9.13 ---
'
'   Changes in 9.13 (tidying: sharing the JSON helpers):
'     Removed JsonEscape / GetJsonStr / GetJsonNum that were held as Private, and
'     switched to using the Public functions of the newly created standard module Wv2Json.bas.
'     The same functions were duplicated in Wv2NavBar too; they are consolidated in one place.
'       GetJsonStr -> Wv2Json.JsonGetStr
'       GetJsonNum -> Wv2Json.JsonGetNum
'       JsonEscape -> Wv2Json.JsonEscape (same name)
'     *The \u path of JsonEscape was also fixed.* High characters (Japanese etc.) no longer
'     get garbled into \uXXXX (spec fact 16. Behavior is unchanged but the log is easier to read).
'
' --- Wv2TabBar.cls  stage 9.11d ---
'
'   Scope of 9.11d (JS->VBA direction, enabling clicks):
'     Connect clicks on the tab bar to VBA tab operations. With this, Layer 2
'     (the WebView2 tab bar) is completed in both directions.
'       - Click on the tab body -> {"cmd":"activate","index":N} -> Browser.ActivateTab N
'       - Click on the x button -> {"cmd":"close","index":N}    -> Browser.CloseTab N
'       - Click on the + button -> {"cmd":"newtab"}            -> Browser.RequestNewTabAsync
'
'   *Command interpretation = Plan 1-A (minimal home-made parser)* VBA has no standard JSON parser, and
'     external COM (ScriptControl etc.) is also not used in this project. The sender is only the JS of our own
'     HTML, so the input is fully trustworthy and a general parser is overkill. We prepare our own minimal functions
'     (GetJsonStr / GetJsonNum) that pull out cmd and index. Since VBA->JS is
'     JSON, JS->VBA is also aligned to JSON, so it extends naturally even if commands increase in the future.
'
'   *Re-entrancy safety = Plan 4-A* OnPaneWebMessage runs inside a WebView2 event
'     handler. When calling another WebView2 operation from here, we must beware of spec fact 14 (waiting inside a handler
'     for an asynchronous WebView2 operation to complete deadlocks).
'       - activate (ActivateTab)  : only Ctrl_PutIsVisible. No completion wait -> safe synchronously
'       - close    (CloseTab)     : only Ctrl_Close. No completion wait     -> safe synchronously
'       - newtab   (AddTabWithUrl): *waits for Controller creation to complete with WaitPaneReady*
'                                   -> falls under spec fact 14. It deadlocks.
'     Therefore only newtab is sent to Browser.RequestNewTabAsync (the SafeTimer one-shot
'     mechanism of 9.10e), and the handler returns immediately. The path with a track record in NewWindowRequested is
'     reused as is (design principle 63).
'
'   *Close policy = Plan 2-A (0 tabs allowed)* Even the last tab can be closed.
'     Consistent with the 9.11a CloseTab decision "the minimum-tab-count policy is left to the upper layer".
'     When there are 0 tabs, the JS side shows "(no tabs)" (already handled in the render of 9.11c).
'     You can recover from the + button.
'
'   *Initial URL of newtab = Plan 3-B (constant HOME_URL)* Held as a constant so that it can be made
'     configurable from the UserForm in the future.
'
' --- Wv2TabBar.cls  stage 9.11c ---
'
'   Scope of 9.11c (VBA->JS direction, drawing the tab list. Display only):
'     Receiving Browser's events (TabAdded/TabClosed/ActiveChanged/TabTitleChanged),
'     send the whole tab list to the tab bar JS and have it redraw. Clicks do not
'     work yet (only the vessel is prepared. activate/close/newtab are 9.11d).
'
'   *Full-sync method (not diffs)* VBA->JS sends the state of all tabs every time:
'       {"type":"sync","tabs":[{"title":"..."},...],"active":N}
'     The JS side redraws the whole DOM from the received array (declarative rendering).
'     JS needs no index correction logic, which is robust (VBA is the source of truth). The initial sync is also a single sync.
'
'   *ready handshake (Plan X)* When loading completes, JS does a postMessage of
'     {"cmd":"ready"}. Only on receiving it does VBA return the initial sync.
'     A PostWebMessageAsString sent before the JS-side message listener is registered gets dropped
'     (in 9.11b the title events flowed unprocessed in exactly this time window), so
'     the timing of the initial sync is taken with the ready signal, which is more reliable than DOMContentLoaded.
'     It also becomes the foundation for receiving clicks in 9.11d.
'
'   *Source of truth for titles = Pane (option 3)* The title is cached in
'     Wv2Pane.m_title (added on the Pane side in 9.11c). PushSyncToJs reads it with
'     m_browser.PaneAt(i).DocumentTitle. It is not affected by index drift.
'
'   *Pane->TabBar back reference = Plan P (WebMessage sink)* To receive postMessage coming from the
'     JS of the tab bar pane (m_barPane), TabBar registers itself as the WebMessage sink of
'     m_barPane (m_barPane.SetWebMsgSink Me).
'     The receiving method is OnPaneWebMessage(pane, source, msg). At release time Shutdown
'     disconnects it via ClearWebMsgSink (design principle 58).
'
'   Implementation breakdown of 9.11c:
'     - Init: register itself as sink on m_barPane + enable reception with View_AddWebMessageReceived.
'     - 4 event handlers: replaced by a common call to PushSyncToJs.
'     - OnPaneWebMessage: on receiving {"cmd":"ready"}, do the initial sync with PushSyncToJs.
'     - PushSyncToJs: turns TabCount/PaneAt(i).DocumentTitle/ActiveIndex into JSON and
'       sends it to JS with m_barPane.View_PostWebMessageAsString. Minimal escaping
'       with JsonEscape.
'     - BuildTabBarHtml: added ready sending + sync receive drawing (clicks are only a vessel).
'
'   WebView2-based tab bar UI. Subscribes to the events of Wv2Browser (WithEvents) to
'   update the look of the tab bar, and converts clicks on the tab bar into operations on Wv2Browser
'   (ActivateTab / CloseTab / new tab).
'
'   *Loose coupling (core policy of stage 9.11)*
'     Wv2Browser knows nothing about the existence of Wv2TabBar. This class is the "listening side"
'     of Browser's Public Events, and from Browser's point of view it is the same as having zero listeners.
'     A user who does not need a tab bar simply does not instantiate this class.
'     A user who wants an MsForms tab bar can connect their own class to Browser
'     with WithEvents instead of this class (this class is one WebView2-based implementation).
'
'   *Ownership structure (Plan b, settled in stage 9.11)*
'     UserForm1
'       | m_browser As Wv2Browser   (manages the content tab group)
'       ` m_tabBar  As Wv2TabBar    (this class. The tab bar UI)
'     This class internally holds m_barPane As Wv2Pane (a single pane for drawing the tab bar).
'     This m_barPane is not put into Wv2Browser.m_tabs (it is not placed under tab management,
'     so that it does not mix with normal tabs). Therefore AddTab is not used and
'     Wv2Pane is created directly.
'
'   *Environment sharing (9.11b)*
'     m_barPane shares the same Environment as Wv2Browser
'     (one Environment is shared by all panes). Init receives browser.Environment and
'     passes it to m_barPane.Init.
'
'   Scope of 9.11b (skeleton only):
'     Build : Init (pane creation -> get View -> virtual host mapping ->
'             show static HTML) / Resize / Shutdown / empty implementations of the
'             WithEvents handlers (Debug.Print only)
'     Not built: drawing the tab list (9.11c) / two-way communication click -> operation (9.11d)
'     -> In 9.11b only "TabBar" is displayed in the top band. The purpose is to confirm that the content area
'       shifts below it and coexists.
'
'   *Pane creation procedure (modeled on Wv2Browser.AddTab, but without the tab management part)*
'     1. New Wv2Pane -> Init(env, hwnd)
'     2. Wait for Controller Ready (poll IsReady / IsFailed, DoEvents)
'     3. Ctrl_GetCoreWebView2 (get View. If omitted, Navigate gives E_POINTER, design principle 56)
'     4. View3_SetVirtualHostNameToFolderMapping (required before Navigate)
'     5. View_Navigate(https://<host>/tabbar.html) (spec fact 13: specify the file name explicitly)
'
'   *HTML delivery (reusing SetVirtualHostNameToFolderMapping from 9.10d)*
'     Write the tab bar HTML to a local folder in UTF-8 and display it as
'     https://appassets.tabbar/tabbar.html. Because Japanese titles will be poured in
'     in the future (9.11c), the writing is done in UTF-8 (no BOM) with ADODB.Stream.
''''''''''''''''''''''''''''''''''

Option Explicit

' ============================================================
' Fields
' ============================================================
Private WithEvents m_browser As Wv2Browser   ' Subscribes to Browser's events
Attribute m_browser.VB_VarHelpID = -1
Private m_barPane   As Wv2Pane                ' Single pane for drawing the tab bar
Private m_hwndHost  As LongPtr               ' hwnd of the host (parent frame)
Private m_folderPath As String               ' Delivery folder of the tab bar HTML
Private m_ready     As Boolean               ' Init-complete flag

' Delivery host name of the tab bar (kept separate from appassets.example on the content side)
Private Const TABBAR_HOST As String = "appassets.tabbar"
Private Const TABBAR_FILE As String = "tabbar.html"

' --- 9.11d: initial URL opened by the + button (newtab) (Plan 3-B) ---
'   Holding it as a constant leaves room to make it configurable from the UserForm in the future.
'   With about:blank the title is empty and the tab stays "(loading)",
'   so a real site is made the default, prioritizing how the verification looks.
Private Const HOME_URL As String = "https://www.bing.com"

' Timeout for waiting for Controller Ready (seconds)
Private Const READY_TIMEOUT As Single = 10#


' ============================================================
' Class_Initialize
' ============================================================
Private Sub Class_Initialize()
    Set m_browser = Nothing
    Set m_barPane = Nothing
    m_hwndHost = 0
    m_folderPath = ""
    m_ready = False
End Sub


' ============================================================
' Init ? Starts the tab bar
'
'   Arguments:
'     browser    : the Wv2Browser to subscribe to (the one managing the content tab group)
'     hwnd       : hwnd of the host (parent frame) (parent of m_barPane)
'     folderPath : local folder to write the tab bar HTML to
'   Return value: True on success
'
'   *Environment is obtained shared from browser.Environment (9.11b)*
' ============================================================
Public Function Init( _
    ByVal browser As Wv2Browser, _
    ByVal hwnd As LongPtr, _
    ByVal folderPath As String) As Boolean

    If browser Is Nothing Then
        Debug.Print "Wv2TabBar.Init: browser is Nothing"
        Exit Function
    End If

    Dim env As Wv2Environment
    Set env = browser.Environment
    If env Is Nothing Then
        Debug.Print "Wv2TabBar.Init: browser.Environment is Nothing (Browser.Init not finished?)"
        Exit Function
    End If

    Set m_browser = browser        ' *WithEvents subscription starts with this
    m_hwndHost = hwnd
    m_folderPath = folderPath

    ' --- 1. Write out the tab bar HTML ---
    If Not WriteTabBarHtml() Then
        Debug.Print "Wv2TabBar.Init: HTML write failed"
        Exit Function
    End If

    ' --- 2. Create the pane for the tab bar (AddTab is not used = not placed under tab management) ---
    Set m_barPane = New Wv2Pane
    m_barPane.Init env, m_hwndHost

    ' --- 3. Wait for Controller Ready ---
    If Not WaitPaneReady(m_barPane) Then
        If m_barPane.IsFailed Then
            Debug.Print "Wv2TabBar.Init: Controller creation failed (lastError=&H" & _
                        Hex(m_barPane.LastError) & ")"
        Else
            Debug.Print "Wv2TabBar.Init: Controller creation timed out"
        End If
        Exit Function
    End If
    Debug.Print "Wv2TabBar.Init: Controller creation OK (pCtrl=" & m_barPane.CtrlPtr & ")"

    ' --- 4. Get View (if omitted, Navigate gives E_POINTER, design principle 56) ---
    Dim hr As Long
    hr = m_barPane.Ctrl_GetCoreWebView2()
    If hr <> 0 Or m_barPane.ViewPtr = 0 Then
        Debug.Print "Wv2TabBar.Init: View acquisition failed hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2TabBar.Init: View acquisition OK (pView=" & m_barPane.ViewPtr & ")"

    ' --- 5. Virtual host mapping (required before Navigate) ---
    hr = m_barPane.View3_SetVirtualHostNameToFolderMapping(TABBAR_HOST, m_folderPath, 2)  ' 2 = DENY_CORS (English fork; was 1 = ALLOW, which let any site read the mapped folder)
    If hr <> 0 Then
        Debug.Print "Wv2TabBar.Init: mapping failed hr=&H" & Hex(hr)
        Exit Function
    End If

    ' --- 5b. Expose hostObjects (9.21, required before Navigate: order A of spec fact 18) ---
    '   *Verification point: does it also hold with a real-URL Navigate?* The "if you attach before Navigate,
    '   JS can grab hostObjects from its first run" established with the settings tab (NavigateToString)
    '   -- we test whether it also works with the real URL to appassets.tabbar.
    '   What is attached is TabBar itself (Me). From JS it is called as
    '   window.chrome.webview.hostObjects.sync.TabBar.HostActivate(n).
    '   Failure is not fatal (the JS side falls back to postMessage), so
    '   here we only log hr and Init continues.
    Dim hrHost As Long
    hrHost = m_barPane.View_AddHostObjectToScript("TabBar", Me)
    If hrHost <> 0 Then
        Debug.Print "Wv2TabBar.Init: AddHostObjectToScript(TabBar) failed hr=&H" & Hex(hrHost) & " (JS falls back to postMessage)"
    Else
        Debug.Print "Wv2TabBar.Init: TabBar bridge attach OK (activate prefers hostObjects)"
    End If

    ' --- 6. Navigate (spec fact 13: specify down to the file name) ---
    Dim url As String
    url = "https://" & TABBAR_HOST & "/" & TABBAR_FILE
    hr = m_barPane.View_Navigate(url)
    If hr <> 0 Then
        Debug.Print "Wv2TabBar.Init: Navigate failed hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2TabBar.Init: Navigate OK (" & url & ")"

    ' --- 7. Register WebMessage sink + enable reception (9.11c, Plan P) ---
    '   Make it possible to receive postMessage from the tab bar JS (ready / future clicks)
    '   in OnPaneWebMessage.
    '   *The initial sync is "not fired here"* A PostMessage before the JS message
    '   listener is registered gets dropped (lesson of 9.11b). We do PushSyncToJs inside OnPaneWebMessage
    '   after JS sends {"cmd":"ready"} (Plan X).
    m_barPane.SetWebMsgSink Me
    hr = m_barPane.View_AddWebMessageReceived()
    If hr <> 0 Then
        Debug.Print "Wv2TabBar.Init: View_AddWebMessageReceived failed hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2TabBar.Init: WebMessage reception enabled (waiting for ready)"

    m_ready = True
    Init = True
    Debug.Print "Wv2TabBar.Init: done (tab bar started, initial sync after ready is received)"
End Function


' ============================================================
' Resize ? Sets the Bounds of the tab bar pane
'   The UserForm_Resize passes in the rectangle of the top band.
'   (The UserForm side calculates the layout split, and Browser does not know the tab bar)
' ============================================================
Public Sub Resize(ByRef barRect As RECT)
    If m_barPane Is Nothing Then Exit Sub
    Dim hr As Long
    hr = m_barPane.Ctrl_PutBounds(barRect)
    If hr <> 0 Then
        Debug.Print "Wv2TabBar.Resize: Ctrl_PutBounds failed hr=&H" & Hex(hr)
    End If
    m_barPane.Ctrl_PutIsVisible True   ' The tab bar is always visible
End Sub


' ============================================================
' Shutdown ? Releases the tab bar pane
'   Called from UserForm_Terminate before m_browser.Shutdown.
'   This class references m_browser with WithEvents, but this is not a cycle
'   (Browser does not know this class), so it is enough to just cancel the subscription with
'   Set m_browser = Nothing.
' ============================================================
Public Sub Shutdown()
    Wv2Log.LogD "Wv2TabBar.Shutdown start"
    If Not m_barPane Is Nothing Then
        m_barPane.ClearWebMsgSink   ' 9.11c: cut the sink registration first (design principle 58)
        m_barPane.Ctrl_Close
        Set m_barPane = Nothing
    End If
    Set m_browser = Nothing   ' Cancel the WithEvents subscription
    m_ready = False
    Wv2Log.LogD "Wv2TabBar.Shutdown done"
End Sub


' ============================================================
' IsReady
' ============================================================
Public Property Get IsReady() As Boolean
    IsReady = m_ready
End Property


' ============================================================
' *Wv2Browser event handlers (9.11c: full-sync method)*
'   For any event they only call PushSyncToJs, which "sends the whole state of all tabs to JS".
'   Diffs are not handled (VBA is the source of truth, JS redraws declaratively).
'   *Note* These fire on Browser events, but the initial sync (the first one right after
'   startup) is triggered by JS's ready signal (OnPaneWebMessage). The sync of events that came before
'   ready may be dropped because the JS listener is not registered yet,
'   but since the latest state is sent in full on receiving ready, it ends up consistent
'   (an advantage of full sync).
' ============================================================
Private Sub m_browser_TabAdded(ByVal index As Long)
    Debug.Print "Wv2TabBar[EVT] TabAdded(index=" & index & ") -> PushSyncToJs"
    PushSyncToJs
End Sub

Private Sub m_browser_TabClosed(ByVal index As Long)
    Debug.Print "Wv2TabBar[EVT] TabClosed(index=" & index & ") -> PushSyncToJs"
    PushSyncToJs
End Sub

Private Sub m_browser_ActiveChanged(ByVal index As Long)
    Debug.Print "Wv2TabBar[EVT] ActiveChanged(index=" & index & ") -> PushSyncToJs"
    PushSyncToJs
End Sub

Private Sub m_browser_TabTitleChanged(ByVal index As Long, ByVal title As String)
    Debug.Print "Wv2TabBar[EVT] TabTitleChanged(index=" & index & ", title=" & title & ") -> PushSyncToJs"
    PushSyncToJs
End Sub


' ============================================================
' OnPaneWebMessage (9.11c, Plan P: receiving end of the m_barPane WebMessage sink)
'   The tab bar JS's window.chrome.webview.postMessage(...) arrives here via Wv2Pane.
'   As a result of registering with Wv2Pane.SetWebMsgSink Me, Wv2Pane's
'   View_OnWebMessageReceived calls OnPaneWebMessage(pane, source, msg)
'   from there.
'
'   Arguments:
'     pane   : the sending pane (always m_barPane here)
'     source : the sender URL
'     msg    : the string JS sent (in this implementation JSON such as {"cmd":"ready"})
'
'   In 9.11c only "ready" is interpreted. The command interpretation for the click types (activate/close/newtab)
'   is added in 9.11d.
'   *No JSON parser is used* VBA has no standard JSON parser, so in 9.11c
'   a simple check of whether "ready" is contained is enough (the received commands are only those
'   sent by the JS of our own HTML, so they are trustworthy). When commands increase in 9.11d it is extended to a minimal parse.
' ============================================================
Public Sub OnPaneWebMessage(ByVal pane As Wv2Pane, ByVal source As String, ByVal msg As String)
    Debug.Print "Wv2TabBar.OnPaneWebMessage: msg=" & msg

    If m_browser Is Nothing Then
        Debug.Print "  -> ignored because m_browser is Nothing"
        Exit Sub
    End If

    Dim cmd As String
    cmd = JsonGetStr(msg, "cmd")

    ' idx is used on the activate/close fallback path to restore the index from JSON.
    ' ok became unnecessary in 9.23 (processing and the failure log are consolidated into the Host methods, so the Case side does not take the result).
    Dim idx As Long

    Select Case cmd

        Case "ready"
            ' --- 9.11c: JS listener registration is complete. Return the initial sync (Plan X) ---
            Debug.Print "  -> ready received. Sending the initial sync."
            PushSyncToJs

        Case "activate"
            ' --- 9.11d/9.23: click on the tab body -> switch active (postMessage fallback path) ---
            '   *9.23: processing consolidated into HostActivate (Plan C-1)* Only when JS in the hybrid
            '   could not grab hostObjects does it fall here. The ActivateTab call and the
            '   failure log are already consolidated inside HostActivate, so here we only restore the index from JSON
            '   and delegate to HostActivate fire-and-forget (the return value is not taken).
            idx = JsonGetNum(msg, "index")
            Debug.Print "  -> activate(index=" & idx & ") [postmsg fallback -> HostActivate]"
            HostActivate idx

        Case "close"
            ' --- 9.11d/9.23: click on the x button -> close the tab (postMessage fallback path) ---
            '   *9.23: processing consolidated into HostClose (Plan C-1)* Only when hostObjects could not
            '   be grabbed does it fall here. The CloseTab call, 0-tab allowance (Plan 2-A) and the failure log are
            '   all consolidated inside HostClose. Here we restore the index and delegate fire-and-forget.
            idx = JsonGetNum(msg, "index")
            Debug.Print "  -> close(index=" & idx & ") [postmsg fallback -> HostClose]"
            HostClose idx

        Case "newtab"
            ' --- 9.11d: click on the + button -> new tab ---
            '   ** Do not call AddTabWithUrl synchronously **
            '   AddTabWithUrl internally does WaitPaneReady (waiting for Controller creation to complete),
            '   so calling it from inside a WebMessage handler falls into the deadlock of
            '   spec fact 14 (the same structure as NewWindowRequested in 9.10e).
            '   RequestNewTabAsync queues the URL, starts the SafeTimer as a one-shot
            '   and returns immediately, so the actual tab creation runs after the handler completes (Plan 4-A).
            Debug.Print "  -> newtab (via RequestNewTabAsync, " & HOME_URL & ")"
            m_browser.RequestNewTabAsync HOME_URL

        Case "settings"
            ' --- 9.18: click on the [gear] button -> open the settings tab ---
            '   ** Do not call OpenSettingsTab synchronously **
            '   Creating the settings tab also passes internally through AddTab -> WaitPaneReady (waiting for
            '   Controller creation to complete), so calling it directly from inside a WebMessage handler falls into the same
            '   spec fact 14 deadlock as newtab. RequestSettingsTabAsync sets
            '   a pending flag, starts the SafeTimer as a one-shot and returns immediately, so
            '   the actual settings tab creation runs after the handler completes (same judgment as Plan 4-A).
            Debug.Print "  -> settings (via RequestSettingsTabAsync)"
            m_browser.RequestSettingsTabAsync

        Case "reorder"
            ' --- 9.15/9.23: drag reordering -> move the tab (postMessage fallback path) ---
            '   *9.23: processing consolidated into HostReorder (Plan C-1)* Only when hostObjects could not
            '   be grabbed does it fall here. from/to are both 1-based. The MoveTab call and the failure log are
            '   consolidated inside HostReorder. Here we restore from/to and delegate fire-and-forget.
            Dim fromIdx As Long, toIdx As Long
            fromIdx = JsonGetNum(msg, "from")
            toIdx = JsonGetNum(msg, "to")
            Debug.Print "  -> reorder(from=" & fromIdx & ", to=" & toIdx & ") [postmsg fallback -> HostReorder]"
            HostReorder fromIdx, toIdx

        Case ""
            Debug.Print "  -> could not extract cmd (msg=" & msg & ")"

        Case Else
            Debug.Print "  -> unknown command: " & cmd

    End Select
End Sub


' ============================================================
' HostActivate (9.21, the activate entry of the hostObjects path)
'
'   Called directly from JS as window.chrome.webview.hostObjects.sync.TabBar.HostActivate(n)
'   (a synchronous RPC that does not go through postMessage). The body is the same as the
'   activate case of OnPaneWebMessage: it just calls m_browser.ActivateTab synchronously.
'
'   *Safety of the synchronous call (Plan 4-A)* ActivateTab is only Ctrl_PutIsVisible and
'     contains no completion wait, so even a JS->VBA sync direct call does not fall into the
'     deadlock of spec fact 14. On success RaiseEvent ActiveChanged ->
'     PushSyncToJs runs, and the active highlight of the tab bar follows via full sync.
'
'   The argument index is 1-based (same as data-index on the JS side). The return value is success or failure (the JS side
'   does not need to look at it, but it is returned as Boolean so that the round trip of the return value can also be checked in the comparison test).
' ============================================================
Public Function HostActivate(ByVal index As Long) As Boolean
    Debug.Print "Wv2TabBar.HostActivate: index=" & index & " (via host)"
    If m_browser Is Nothing Then
        Debug.Print "  -> ignored because m_browser is Nothing"
        HostActivate = False
        Exit Function
    End If
    Dim ok As Boolean
    ok = m_browser.ActivateTab(index)
    If Not ok Then
        Debug.Print "    [warning] ActivateTab failed (index=" & index & ")"
    End If
    HostActivate = ok
End Function


' ============================================================
' HostClose (9.22, the close entry of the hostObjects path)
'
'   Called directly from JS as window.chrome.webview.hostObjects.sync.TabBar.HostClose(n)
'   (a synchronous RPC that does not go through postMessage). The body is the same as the
'   close case of OnPaneWebMessage: it just calls m_browser.CloseTab synchronously.
'
'   *Safety of the synchronous call (Plan 4-A)* CloseTab is only Ctrl_Close and does not
'     contain a completion wait, so even a JS->VBA sync direct call does not fall into the spec fact 14 deadlock.
'     *0 tabs allowed (Plan 2-A)*: even the last tab can be closed, the active
'     correction inside CloseTab (the 4 branches of 9.11a) runs, and via RaiseEvent ActiveChanged /
'     TabClosed -> PushSyncToJs the tab bar follows through full sync.
'
'   The argument index is 1-based (same as data-index on the JS side). The return value is success or failure.
' ============================================================
Public Function HostClose(ByVal index As Long) As Boolean
    Debug.Print "Wv2TabBar.HostClose: index=" & index & " (via host)"
    If m_browser Is Nothing Then
        Debug.Print "  -> ignored because m_browser is Nothing"
        HostClose = False
        Exit Function
    End If
    Dim ok As Boolean
    ok = m_browser.CloseTab(index)
    If Not ok Then
        Debug.Print "    [warning] CloseTab failed (index=" & index & ")"
    End If
    HostClose = ok
End Function


' ============================================================
' HostReorder (9.22, the reorder entry of the hostObjects path)
'
'   Called directly from JS as window.chrome.webview.hostObjects.sync.TabBar.HostReorder(from, to)
'   (a synchronous RPC that does not go through postMessage). The body is the same as the
'   reorder case of OnPaneWebMessage: it just calls m_browser.MoveTab synchronously.
'
'   *Safety of the synchronous call (Plan 4-A)* MoveTab contains only array manipulation (Remove/Add) and
'     PutIsVisible and no completion wait, so even a sync direct call
'     does not deadlock. On success RaiseEvent ActiveChanged -> PushSyncToJs
'     runs, and the full-sync redraw runs in the reordered order.
'
'   *Argument names fromIdx/toIdx* Avoided because From has been reported to clash in some dialects. The JS side
'     calls with positional arguments (HostReorder(from,to)) so the VBA argument names are invisible to JS.
'     Both are 1-based (same meaning as JS's from/to). The return value is success or failure.
' ============================================================
Public Function HostReorder(ByVal fromIdx As Long, ByVal toIdx As Long) As Boolean
    Debug.Print "Wv2TabBar.HostReorder: from=" & fromIdx & " to=" & toIdx & " (via host)"
    If m_browser Is Nothing Then
        Debug.Print "  -> ignored because m_browser is Nothing"
        HostReorder = False
        Exit Function
    End If
    Dim ok As Boolean
    ok = m_browser.MoveTab(fromIdx, toIdx)
    If Not ok Then
        Debug.Print "    [warning] MoveTab failed (from=" & fromIdx & ", to=" & toIdx & ")"
    End If
    HostReorder = ok
End Function


' ============================================================
' PushSyncToJs (9.11c, full sync: turn the state of all tabs into JSON and send it to JS)
'
'   Shape of the JSON sent:
'     {"type":"sync","tabs":[{"title":"..."},{"title":"..."}],"active":N}
'   active is 1-based (0 = no active tab). The JS side redraws the whole DOM from this
'   array.
'
'   The source of truth for titles is Wv2Pane.DocumentTitle (option 3). Take each pane with m_browser.PaneAt(i)
'   and read DocumentTitle. A pane whose title has not been obtained yet gives an empty string,
'   so the JS side may fall back to (loading) etc. (this implementation
'   sends the empty string as is and leaves the display to JS).
' ============================================================
Private Sub PushSyncToJs()
    If m_barPane Is Nothing Then Exit Sub
    If m_browser Is Nothing Then Exit Sub
    If Not m_ready Then Exit Sub

    Dim n As Long
    n = m_browser.TabCount

    Dim json As String
    json = "{""type"":""sync"",""tabs"":["

    Dim i As Long
    Dim pane As Wv2Pane
    Dim title As String
    For i = 1 To n
        Set pane = m_browser.PaneAt(i)
        If pane Is Nothing Then
            title = ""
        Else
            title = pane.DocumentTitle
        End If
        If i > 1 Then json = json & ","
        json = json & "{""title"":""" & JsonEscape(title) & """}"
    Next i

    json = json & "],""active"":" & m_browser.ActiveIndex & "}"

    Dim hr As Long
    hr = m_barPane.View_PostWebMessageAsString(json)
    If hr <> 0 Then
        Debug.Print "Wv2TabBar.PushSyncToJs: PostWebMessageAsString failed hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2TabBar.PushSyncToJs: send OK (" & n & " tabs, active=" & _
                    m_browser.ActiveIndex & ")"
        Debug.Print "  json=" & json
    End If
End Sub


' ============================================================
' WaitPaneReady ? Wait for Controller Ready (same form as Wv2Browser.WaitPaneReady)
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
' WriteTabBarHtml ? Writes the tab bar HTML to the folder in UTF-8 (no BOM)
'   9.11b is static (only shows "TabBar" in the center). JS is added from 9.11c on.
'   *UTF-8 writing* ADODB.Stream is used so that Japanese titles can be poured in in the future
'   (VBA's Print # depends on the default code page and garbles Japanese).
' ============================================================
Private Function WriteTabBarHtml() As Boolean
    On Error GoTo eh

    ' Create the folder if it does not exist
    If Dir(m_folderPath, vbDirectory) = "" Then
        MkDir m_folderPath
        Debug.Print "Wv2TabBar.WriteTabBarHtml: creating folder " & m_folderPath
    End If

    Dim fullPath As String
    fullPath = m_folderPath
    If Right$(fullPath, 1) <> "\" Then fullPath = fullPath & "\"
    fullPath = fullPath & TABBAR_FILE

    Dim html As String
    html = BuildTabBarHtml()

    ' Write UTF-8 (no BOM) with ADODB.Stream
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                 ' adTypeText
    st.Charset = "UTF-8"
    st.Open
    st.WriteText html

    ' Remove the UTF-8 BOM and save
    Dim bin As Object
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 2
    bin.Charset = "UTF-8"
    bin.Open
    bin.WriteText html
    ' Write once as text, then re-read as binary and discard the first 3 bytes (BOM)
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

    st.Close

    Debug.Print "Wv2TabBar.WriteTabBarHtml: write OK " & fullPath & " (" & Len(html) & " chars, UTF-8)"
    WriteTabBarHtml = True
    Exit Function
eh:
    Debug.Print "Wv2TabBar.WriteTabBarHtml: error " & Err.Number & " " & Err.Description
    WriteTabBarHtml = False
End Function


' ============================================================
' BuildTabBarHtml ? Builds the tab bar HTML (9.11c: supports sync drawing)
'
'   JS behavior in 9.11c:
'     1. On load completion it sends window.chrome.webview.postMessage('{"cmd":"ready"}')
'        (ready handshake, Plan X). With this VBA returns the initial sync.
'     2. It receives sync with window.chrome.webview.addEventListener('message', ...) and
'        redraws the whole tab DOM from {type:"sync", tabs:[{title}], active:N}
'        (full sync, declarative rendering).
'     3. 9.11d: send clicks to VBA. Tab body -> activate, x -> close,
'        + -> newtab. A click on x calls stopPropagation so that it does not propagate to the
'        tab body's onclick (so the parent's activate does not misfire).
'
'   *The JSON string embeds literals using VBA's double-quote escaping ("")*
' ============================================================
Private Function BuildTabBarHtml() As String
    Dim s As String
    s = "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    s = s & "<style>" & vbLf
    s = s & "  html,body{margin:0;padding:0;height:100%;box-sizing:border-box;" & _
            "font-family:'Segoe UI','Meiryo',sans-serif;}" & vbLf
    s = s & "  *{box-sizing:border-box;}" & vbLf
    s = s & "  body{display:flex;align-items:stretch;overflow:hidden;" & _
            "background:#f3f3f3;border-bottom:1px solid #ccc;}" & vbLf
    s = s & "  #tabs{display:flex;align-items:stretch;height:100%;overflow-x:auto;}" & vbLf
    ' --- 9.32 (point 4): hide the horizontal scrollbar ---
    '     Because overflow-x:auto remains, *the wheel can scroll horizontally*
    '     (Chromium converts the wheel to horizontal on a horizontal container that cannot scroll vertically).
    '     When the bar appears it eats the height of the tab bar from the inside and squashes the tabs, so it is hidden.
    s = s & "  #tabs::-webkit-scrollbar{display:none;}" & vbLf
    ' --- 9.31 (A-1): switch only the background color smoothly over 0.12 s (point 5) ---
    s = s & "  .tab{display:flex;align-items:center;max-width:200px;min-width:80px;" & _
            "height:100%;padding:0 6px 0 12px;border-right:1px solid #ccc;" & _
            "background:#e7e7e7;cursor:pointer;font-size:13px;color:#333;" & _
            "white-space:nowrap;transition:background-color .12s;}" & vbLf
    s = s & "  .tab.active{background:#fff;color:#000;font-weight:600;}" & vbLf
    ' --- 9.31 (A-1): hover background. *Inactive only* (point 2) ---
    '     :not(.active) gives specificity 0,3,0. It is stronger than .tab.active (0,2,0), so
    '     the #fff of the active tab is kept regardless of the order of writing.
    s = s & "  .tab:not(.active):hover{background:#efefef;}" & vbLf
    s = s & "  .tab .title{overflow:hidden;text-overflow:ellipsis;flex:1;}" & vbLf
    ' --- 9.31 (A-2): hide the close button by default (points 3, 4) ---
    '     *Why visibility:hidden is used*
    '       With display:none the 18px + 6px margin disappears and the title width expands and shrinks,
    '       so the text jitters at every hover. With opacity:0 the invisible button
    '       can be clicked. With visibility:hidden it disappears while keeping its place,
    '       and pointer events are also disabled.
    s = s & "  .tab .close{margin-left:6px;width:18px;height:18px;line-height:18px;" & _
            "text-align:center;border-radius:3px;color:#888;visibility:hidden;}" & vbLf
    ' --- 9.31 (A-2): show only when active or hovered (equivalent to Chrome) ---
    '     Both have specificity 0,3,0, stronger than .tab .close (0,2,0).
    s = s & "  .tab.active .close,.tab:hover .close{visibility:visible;}" & vbLf
    s = s & "  .tab .close:hover{background:#ccc;color:#000;}" & vbLf
    s = s & "  .tab.dragging{opacity:0.45;}" & vbLf
    s = s & "  .tab.dragover{box-shadow:inset 3px 0 0 #4a90e2;}" & vbLf
    s = s & "  .tab.dragover-after{box-shadow:inset -3px 0 0 #4a90e2;}" & vbLf
    s = s & "  #newtab{display:flex;align-items:center;justify-content:center;" & _
            "width:32px;height:100%;cursor:pointer;font-size:18px;color:#555;}" & vbLf
    s = s & "  #newtab:hover{background:#e0e0e0;}" & vbLf
    ' --- 9.18: settings ([gear]) button. To the right of newtab. Pushed to the right end ---
    s = s & "  #settings{display:flex;align-items:center;justify-content:center;" & _
            "width:34px;height:100%;cursor:pointer;font-size:16px;color:#666;margin-left:auto;}" & vbLf
    s = s & "  #settings:hover{background:#e0e0e0;color:#333;}" & vbLf
    s = s & "  #empty{padding:0 12px;display:flex;align-items:center;color:#999;font-size:13px;}" & vbLf
    s = s & "</style></head><body>" & vbLf
    s = s & "  <div id=""tabs""></div>" & vbLf
    s = s & "  <div id=""newtab"" title=""New tab"">+</div>" & vbLf
    ' --- 9.18: settings button ([gear]). To the right end with margin-left:auto ---
    s = s & "  <div id=""settings"" title=""Settings"">&#9881;</div>" & vbLf
    s = s & "<script>" & vbLf
    s = s & "  var W = window.chrome && window.chrome.webview;" & vbLf
    s = s & "  function postCmd(obj){ if(W){ W.postMessage(JSON.stringify(obj)); } }" & vbLf
    ' --- 9.15: drag state and insertion-marker cleanup helper ---
    ' --- 9.32: HO() moved from inside render to top level (post-hoc report) ---
    s = s & "  function HO(){ return (W && W.hostObjects && W.hostObjects.sync) ? W.hostObjects.sync : null; }" & vbLf
    ' --- 9.32 (point 1): optimistic update. Change only the look first without waiting for VBA ---
    '     The next full sync always overwrites it, so no state remains on the JS side (compatible with design principle 60).
    s = s & "  function markActive(el){" & vbLf
    s = s & "    var xs = document.querySelectorAll('.tab.active');" & vbLf
    s = s & "    for(var k=0;k<xs.length;k++){ xs[k].classList.remove('active'); }" & vbLf
    s = s & "    if(el){ el.classList.add('active'); }" & vbLf
    s = s & "  }" & vbLf
    ' --- 9.32 (point 2): run fn "after one frame has been drawn" with a double rAF ---
    '     Because blocking on a synchronous COM call right after changing the DOM freezes it before painting.
    s = s & "  function afterPaint(fn){" & vbLf
    s = s & "    if(window.requestAnimationFrame){" & vbLf
    s = s & "      requestAnimationFrame(function(){ requestAnimationFrame(fn); });" & vbLf
    s = s & "    } else {" & vbLf
    s = s & "      setTimeout(fn, 0);" & vbLf
    s = s & "    }" & vbLf
    s = s & "  }" & vbLf
    ' --- 9.32: split out the sending part of activate (the body is identical to up to 9.31) ---
    s = s & "  function sendActivate(n){" & vbLf
    s = s & "    var ho = HO();" & vbLf
    s = s & "    if(ho && ho.TabBar){" & vbLf
    s = s & "      try{ ho.TabBar.HostActivate(n); console.log('[host] activate '+n); }" & vbLf
    s = s & "      catch(e){ console.log('[host-fail->postmsg] '+e); postCmd({cmd:'activate', index:n}); }" & vbLf
    s = s & "    } else {" & vbLf
    s = s & "      console.log('[postmsg] activate '+n); postCmd({cmd:'activate', index:n});" & vbLf
    s = s & "    }" & vbLf
    s = s & "  }" & vbLf
    ' --- 9.32b: remember the previous active (firing condition for the nearest step. Post-hoc report 1) ---
    s = s & "  var lastActive = -1;" & vbLf
    s = s & "  var dragSrc = null;" & vbLf
    s = s & "  function clearMarks(){" & vbLf
    s = s & "    var xs = document.querySelectorAll('.dragover, .dragover-after');" & vbLf
    s = s & "    for(var k=0;k<xs.length;k++){ xs[k].classList.remove('dragover'); xs[k].classList.remove('dragover-after'); }" & vbLf
    s = s & "  }" & vbLf
    ' --- Draw the whole tab list (full sync) ---
    s = s & "  function render(state){" & vbLf
    s = s & "    var tabsEl = document.getElementById('tabs');" & vbLf
    ' --- 9.32b (point 3): innerHTML='' resets scrollLeft to 0, so save it ---
    s = s & "    var keepLeft = tabsEl.scrollLeft;" & vbLf
    s = s & "    tabsEl.innerHTML = '';" & vbLf
    s = s & "    var list = (state && state.tabs) ? state.tabs : [];" & vbLf
    s = s & "    var active = (state && state.active) ? state.active : 0;" & vbLf
    s = s & "    if(list.length === 0){" & vbLf
    s = s & "      var e = document.createElement('div');" & vbLf
    s = s & "      e.id = 'empty'; e.textContent = '(no tabs)';" & vbLf
    s = s & "      tabsEl.appendChild(e); return;" & vbLf
    s = s & "    }" & vbLf
    s = s & "    for(var i=0; i<list.length; i++){" & vbLf
    s = s & "      var idx = i + 1;" & vbLf                     ' 1-based
    s = s & "      var t = list[i] || {};" & vbLf
    s = s & "      var title = (t.title && t.title.length) ? t.title : '(loading)';" & vbLf
    s = s & "      var tab = document.createElement('div');" & vbLf
    s = s & "      tab.className = 'tab' + (idx === active ? ' active' : '');" & vbLf
    s = s & "      tab.setAttribute('data-index', idx);" & vbLf
    s = s & "      tab.draggable = true;" & vbLf
    s = s & "      var ti = document.createElement('span');" & vbLf
    s = s & "      ti.className = 'title'; ti.textContent = title; ti.title = title;" & vbLf
    s = s & "      var cl = document.createElement('span');" & vbLf
    s = s & "      cl.className = 'close'; cl.textContent = '\u00d7';" & vbLf
    ' --- Clicks are only a vessel (content comes in 9.11d). For now only a debug log, nothing is sent ---
    ' --- 9.21: activate prefers hostObjects + postMessage fallback ---
    '   HO() returns hostObjects.sync if it can be obtained, otherwise null. Same form as the settings tab's
    '   null guard. If obtained, TabBar.HostActivate(n) is called directly and synchronously.
    ' --- 9.32 (points 1, 2): change the look first, let one frame be drawn, then call VBA ---
    s = s & "      tab.onclick = function(ev){" & vbLf
    s = s & "        if(ev.target && ev.target.className === 'close') return;" & vbLf
    s = s & "        var n = parseInt(this.getAttribute('data-index'), 10);" & vbLf
    s = s & "        markActive(this);" & vbLf
    s = s & "        afterPaint(function(){ sendActivate(n); });" & vbLf
    s = s & "      };" & vbLf
    ' --- 9.22: close also prefers hostObjects + postMessage fallback ---
    '   The same three-tier structure as activate (HO() reuses the one defined in 9.21).
    s = s & "      cl.onclick = function(ev){" & vbLf
    s = s & "        ev.stopPropagation();" & vbLf
    s = s & "        var n = parseInt(this.parentNode.getAttribute('data-index'), 10);" & vbLf
    s = s & "        var ho = HO();" & vbLf
    s = s & "        if(ho && ho.TabBar){" & vbLf
    s = s & "          try{ ho.TabBar.HostClose(n); console.log('[host] close '+n); }" & vbLf
    s = s & "          catch(e){ console.log('[host-fail->postmsg] '+e); postCmd({cmd:'close', index:n}); }" & vbLf
    s = s & "        } else {" & vbLf
    s = s & "          console.log('[postmsg] close '+n); postCmd({cmd:'close', index:n});" & vbLf
    s = s & "        }" & vbLf
    s = s & "      };" & vbLf
    ' --- 9.15: reorder with HTML5 Drag & Drop (Plan B-1) ---
    '   dragstart keeps the grabbed source index in dragSrc. dragover looks at "whether it is the
    '   left half or the right half of the target tab" and attaches the insertion position marker (.dragover / .dragover-after).
    '   drop computes from/to and postMessages reorder, leaving it to VBA.
    '   The DOM is not moved by itself; it is left to the sync redraw from VBA (VBA is the source of truth).
    s = s & "      tab.ondragstart = function(ev){" & vbLf
    s = s & "        dragSrc = parseInt(this.getAttribute('data-index'),10);" & vbLf
    s = s & "        this.classList.add('dragging');" & vbLf
    s = s & "        if(ev.dataTransfer){ ev.dataTransfer.effectAllowed='move'; try{ev.dataTransfer.setData('text/plain', String(dragSrc));}catch(e){} }" & vbLf
    s = s & "      };" & vbLf
    s = s & "      tab.ondragover = function(ev){" & vbLf
    s = s & "        ev.preventDefault();" & vbLf
    s = s & "        if(ev.dataTransfer){ ev.dataTransfer.dropEffect='move'; }" & vbLf
    s = s & "        clearMarks();" & vbLf
    s = s & "        var r = this.getBoundingClientRect();" & vbLf
    s = s & "        var after = (ev.clientX - r.left) > (r.width/2);" & vbLf
    s = s & "        this.classList.add(after ? 'dragover-after' : 'dragover');" & vbLf
    s = s & "      };" & vbLf
    s = s & "      tab.ondragleave = function(ev){ this.classList.remove('dragover'); this.classList.remove('dragover-after'); };" & vbLf
    s = s & "      tab.ondrop = function(ev){" & vbLf
    s = s & "        ev.preventDefault();" & vbLf
    s = s & "        var target = parseInt(this.getAttribute('data-index'),10);" & vbLf
    s = s & "        var r = this.getBoundingClientRect();" & vbLf
    s = s & "        var after = (ev.clientX - r.left) > (r.width/2);" & vbLf
    s = s & "        clearMarks();" & vbLf
    s = s & "        if(dragSrc == null) return;" & vbLf
    ' --- compute from/to as 1-based ---
    '   after=false -> put it at the target's position (make it the target-th).
    '   after=true  -> put it right after the target (the (target+1)-th).
    '   However if the source position dragSrc is before the insertion position, the
    '   target index shifts one earlier by the amount removed, so correct for it (VBA MoveTab
    '   expects "to-th after the reorder", so the final to is fixed here).
    s = s & "        var dest = after ? target + 1 : target;" & vbLf
    s = s & "        if(dragSrc < dest){ dest = dest - 1; }" & vbLf
    s = s & "        if(dest < 1) dest = 1;" & vbLf
    ' --- 9.22: reorder also prefers hostObjects + postMessage fallback ---
    '   The same three-tier structure as activate/close. HostReorder(from, to) is called with positional arguments.
    s = s & "        if(dragSrc !== dest){" & vbLf
    s = s & "          var ho = HO();" & vbLf
    s = s & "          if(ho && ho.TabBar){" & vbLf
    s = s & "            try{ ho.TabBar.HostReorder(dragSrc, dest); console.log('[host] reorder '+dragSrc+'->'+dest); }" & vbLf
    s = s & "            catch(e){ console.log('[host-fail->postmsg] '+e); postCmd({cmd:'reorder', from:dragSrc, to:dest}); }" & vbLf
    s = s & "          } else {" & vbLf
    s = s & "            console.log('[postmsg] reorder '+dragSrc+'->'+dest); postCmd({cmd:'reorder', from:dragSrc, to:dest});" & vbLf
    s = s & "          }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        dragSrc = null;" & vbLf
    s = s & "      };" & vbLf
    s = s & "      tab.ondragend = function(ev){ this.classList.remove('dragging'); clearMarks(); dragSrc = null; };" & vbLf
    s = s & "      tab.appendChild(ti); tab.appendChild(cl);" & vbLf
    s = s & "      tabsEl.appendChild(tab);" & vbLf
    s = s & "    }" & vbLf
    ' --- 9.32b (point 3): restore -> bring in with nearest only when the active tab changed ---
    s = s & "    tabsEl.scrollLeft = keepLeft;" & vbLf
    s = s & "    if(active !== lastActive){" & vbLf
    s = s & "      var act = tabsEl.querySelector('.tab.active');" & vbLf
    s = s & "      if(act && act.scrollIntoView){" & vbLf
    s = s & "        try{ act.scrollIntoView({block:'nearest', inline:'nearest'}); }catch(e){}" & vbLf
    s = s & "      }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    lastActive = active;" & vbLf
    s = s & "  }" & vbLf
    ' --- The new-tab button is also only a vessel ---
    ' --- 9.32b (points 1, 2): convert vertical wheel into horizontal scroll ---
    '     Even if innerHTML is rewritten the #tabs element itself is not replaced, so
    '     a listener attached once survives. passive:false is for preventDefault.
    s = s & "  var tabsRoot = document.getElementById('tabs');" & vbLf
    s = s & "  if(tabsRoot){" & vbLf
    s = s & "    tabsRoot.addEventListener('wheel', function(ev){" & vbLf
    s = s & "      if(tabsRoot.scrollWidth <= tabsRoot.clientWidth) return;" & vbLf
    s = s & "      var d = ev.deltaY || 0;" & vbLf
    s = s & "      if(d === 0) return;" & vbLf
    s = s & "      tabsRoot.scrollLeft += d;" & vbLf
    s = s & "      ev.preventDefault();" & vbLf
    s = s & "    }, {passive:false});" & vbLf
    s = s & "  }" & vbLf
    s = s & "  document.getElementById('newtab').onclick = function(){" & vbLf
    s = s & "    postCmd({cmd:'newtab'});" & vbLf
    s = s & "  };" & vbLf
    ' --- 9.18: settings button -> settings command ---
    s = s & "  document.getElementById('settings').onclick = function(){" & vbLf
    s = s & "    postCmd({cmd:'settings'});" & vbLf
    s = s & "  };" & vbLf
    ' --- Receive sync from VBA ---
    s = s & "  if(W){" & vbLf
    s = s & "    W.addEventListener('message', function(ev){" & vbLf
    s = s & "      var data = ev.data;" & vbLf
    s = s & "      if(typeof data === 'string'){ try{ data = JSON.parse(data); }catch(e){ return; } }" & vbLf
    s = s & "      if(data && data.type === 'sync'){ render(data); }" & vbLf
    s = s & "    });" & vbLf
    s = s & "  }" & vbLf
    ' --- ready handshake: send ready after the listener is registered (Plan X) ---
    s = s & "  render({tabs:[], active:0});" & vbLf
    s = s & "  postCmd({cmd:'ready'});" & vbLf
    s = s & "</script>" & vbLf
    s = s & "</body></html>"
    BuildTabBarHtml = s
End Function

