# VBA-WebView2 Development Notes Part 3 — Axis D (DOM / Page Automation)

The implementation record of roadmap R-4 (MSHTML equivalent). It is a direct investment in concept ①, "ease of use that excels at
automatically controlling web pages". **The current place to append.**

---

## Stage D-1 — A synchronous-retrieval primitive for ExecuteScript (EvalSync)

### Where it fits

The first step of roadmap R-4 (MSHTML equivalent). **We have not yet built the element registry or Wv2Element, and
we focused only on confirming "whether a path that uses an asynchronous API with a synchronous face can safely hold up on the current architecture"**.
If this passes, the rest reduces to a lateral rollout of plan C-1.

### What the preliminary investigation revealed (it was lighter than expected)

- `View_ExecuteScript` assigns a callback ID on each call and takes an independent slot
  (9.8c, plan a) = **concurrent calls are already safe**
- `View_ExecuteScriptSafe` returns `{ok:true,value:…}` / `{ok:false,errorName:…}` (9.9c)
- The DoEvents loop of `WaitPaneReady` is a precedent for waiting (`READY_TIMEOUT = 10#`)
- The old version (`旧WebView2.cls`, the standard-thunk method) has a precedent in `ExecuteScriptSync`,
  and the skeleton of flag + DoEvents + timeout already worked

-> D-1 was not about "building an asynchronous mechanism" but **"laying one layer over it that matches callbackId against results and waits"**.

### Issues and decisions (8 items agreed in one go)

| # | Issue | Decision |
|---|---|---|
| 1 | Method name | `EvalSync` with no prefix. `View_` / `Ctrl_` mean a one-to-one correspondence with a COM method, so they are not attached to a high-level API with a built-in wait loop |
| 2 | Receiver of results | A pending table keyed by callbackId (design principle 87). A direct call of `UserForm1.OnExecuteScriptResult` is just Debug.Print, so it is kept |
| 3 | Shape of return value | `String` (= the JSON representation of value). Success is told by `LastEvalOk` / `LastEvalError`. Type conversion is D-2 onward |
| 4 | Via Safe | Always go through `View_ExecuteScriptSafe`. Raw, a JS exception returns as S_OK + "null" and cannot be told from a normal null |
| 5 | Timeout | Default 5 seconds. When exceeded, **leave it as discarded**, and a late arrival looks at it and is thrown away |
| 6 | Concurrency and nesting | Allowed. They are separated by callbackId so they are not mixed up. Only log the depth |
| 7 | in-callback guard | ★The highlight of this stage★ If `m_inCallback > 0`, fail immediately without firing JS (design principle 88) |
| 8 | Verification | `Test_D1_Eval` / `Test_D1_Guard` / `Test_D1_Help` |

### Added Public members (Wv2Pane)

- `EvalSync(jsExpr, [timeoutSec=5]) As String`
- `LastEvalOk As Boolean` / `LastEvalError As String`
- `ResetCallbackGuard()` / `InCallbackDepth As Long`
- `Debug_SetInCallback(n)` (injection port for verification only)

Usage examples:

```
?UserForm1.GetActivePane.EvalSync("document.title")                        → "Bぷよ"  (quoted)
?UserForm1.GetActivePane.EvalSync("document.querySelectorAll('a').length")  → 22
```

### Post-implementation report (design principle 79)

1. **Separated the handlers into `_Core`.** At agreement time we explained "+1 at the entry and -1 at the exit of each View_On*", but in the actual code
   `View_OnNewWindowRequested` had 5 mid-procedure Exit Subs and
   the others had one each. We renamed the body to `Private View_OnXxx_Core`, and a thin Public wrapper of the same name
   does `+1 -> Core -> -1` (7 of them). -> Generalized as design principle 89.
2. **The counter is also restored on error but the error is not swallowed.** After restoring with `On Error GoTo GuardExit`,
   we re-throw the same error with `Err.Raise`. `ComCallbackHandler` has no error handling and previously
   it propagated as is, so this is to keep the behavior unchanged.
3. **No guard was attached to `OnExecuteScriptCompleted`.** It is the receiver of results and no
   user code runs inside it. Attaching one would raise the counter during EvalSync's DoEvents and be confusing.
4. **JSON extraction was made a Private function `JsonPickStr` inside Wv2Pane.**
   `Wv2Json.JsonGetStr` is, by design principle 64, an implementation that **does not unescape**, so it cuts off at the first
   escaped quote inside a doubly JSON-encoded value. If it becomes needed elsewhere,
   promote it to Wv2Json.
5. **Wrapped with `JSON.stringify`.**
   When we fire `View_ExecuteScriptSafe("JSON.stringify(" & jsExpr & ")")`, the value Safe returns is
   **always a string**, and one limited parser can extract it regardless of type.

### Checks done before implementation

- Formatting: UTF-8 without BOM / CRLF only / no double CR / zero bare LF
- Syntax: parity of quotes, comments to the right of the line-continuation `_`, matching open/close of 27 Subs, 46 Functions, 13 Properties
- **Reproduced `JsonPickStr`'s algorithm in Python and verified with real data**:
  12 types x with/without non-ASCII outer characters as `\uXXXX` = all 24 cases passed.
  Also confirmed detection of `"ok":false` in the exception path and extraction of errorName / errorMessage.
  Note: at first the simulator was all NG, but the cause was `json.dumps` inserting `": "` and spaces,
    whereas JS's `JSON.stringify` inserts no spaces. **It was insufficient construction on the checking side.**
- **Reserved-word check (design principle 90)**: cross-checked 49 identifiers. Made permanent after the `eNum` accident.

### Real-machine test results (all passed)

- **Types**: integer / decimal / string / Japanese / with quotes / with backslash / boolean /
  null / undefined / object / array / 300 characters, all decoded correctly
  (`{"ok":true,"value":"{\"a\":1,\"b\":[1,2,3]}"}` -> `{"a":1,"b":[1,2,3]}`)
- **Real pages**: `document.title` -> `"Bぷよ"`, `location.href`, `document.body.innerText.length` -> 305,
  `document.querySelectorAll('a').length` -> 22
- **Failure paths**: ReferenceError / Error: boom / syntax-error-or-null each classified correctly
- **Timeout**: made cb19 discarded and returned in 1 second -> log of discarding the late arrival ->
  the immediately following cb20 was normal. **The design of issue 5 was supported by measurement**
- **Guard**: in 2) of `Test_D1_Guard`, **no log at all** of `View_ExecuteScriptSafe` appears =
  it is rejected before JS is issued. It returns immediately without hanging
- **Regression (the effect of re-wrapping 7 View_On* handlers)**: adding / switching / closing tabs, D&D reordering,
  URL entry, back, forward, reload, link new-tab expansion, search engine change in the settings tab,
  all normal. **The POST link of `Test_9_26_PostProbe` too, down to pre-warm delegation,
  `put_NewWindow`, the POST body, `Sec-Fetch-User: ?1`, Referer, and Origin, was completely as before**
  (spec fact 26 reproduced). We could confirm that wrapping is harmless.

### Accident we stepped on

**★Collision with a reserved word★** `Dim eNum As Long` was a compile error. VBA does not
distinguish identifier case, so `eNum` = `Enum`. It came to light through tarboh's remark. -> Made the check
permanent as design principle 90.

### Changed files

Only 2: `Wv2Pane.cls` (D-1 stage) / `Wv2Tests.bas` (D-1 stage).

### Design principles established (86 to 90)

- **Design principle 86 (D-1) Stage numbers are assigned as "axis symbol + sequence number".**
  If a single sequence keeps being extended, work of different natures lines up in the same column and cannot be read later.
  When the axis changes, split the symbol rather than increasing the number. Past axes are frozen and not rewritten.

- **Design principle 87 (D-1) When using an asynchronous API with a synchronous face, do not receive the result in a "single variable".**
  Receive it in a pending table keyed by callbackId. With a single variable, concurrent calls mix up
  the results. Even though the caller side is already built to be concurrency-safe (View_ExecuteScript has had
  callbackId numbering since 9.8c), throwing that property away in an upper layer is a design regression.
  ★Do not delete a timed-out entry but leave it as "discarded"★. If deleted,
  a completion notification arriving late becomes an "unknown ID", clouding the log, and we cannot even tell whether it actually came.

- **Design principle 88 (D-1) In territory where the debugger cannot be used, turn "hanging" into an "immediate error".**
  Spec fact 14 (waiting synchronously from inside an event handler deadlocks) occurs in territory where, by spec fact 20,
  neither breaking nor stepping is possible, so stepping on it makes finding the cause extremely painful.
  Place a guard so it cannot be stepped on, write the reason for rejection to the log, and prepare a means of recovery (ResetCallbackGuard)
  at the same time. The existence of the guard itself serves as documentation.

- **Design principle 89 (D-1) When adding pre/post processing to an existing procedure that has Exits in the middle,
  separate the body into `_Core` and wrap it with a thin wrapper.**
  The method of adding 2 lines at the entry and exit creates paths where the restore is missed if there are mid-procedure Exits
  (View_OnNewWindowRequested had 5 Exit Subs). With a wrapper, all paths can be gathered to one point with
  `On Error GoTo`, and the public name does not change, so callers need no modification.
  ★Do not swallow errors; re-throw with Err.Raise★ (do not change the previous propagation behavior).

- **Design principle 90 (D-1) Check new identifiers against the VBA reserved-word list "ignoring case".**
  VBA does not distinguish identifier case, so `eNum` becomes the reserved word `Enum` itself and
  causes a compile error (actually stepped on in D-1). They collide even if they look different.
  We made permanent in the Python check process: "pick up `Dim x As` / `ByVal x As` / `Private x As` with a regex and
  cross-check, lower-cased, against a set of about 170 reserved words".

### Spec facts established (29 to 32)

- **Spec Fact 29 (D-1) The keys of ExecuteScript's result JSON are reordered alphabetically.**
  Although View_ExecuteScriptSafe's JS builds the object in the order `{ok, errorName, errorMessage, stack}`,
  by the time it reaches VBA it is
  `{"errorMessage":...,"errorName":...,"ok":false,"stack":...}`.
  The runtime is rebuilding the result. **Writing position-dependent parsing will break.**
  Search by key name.

- **Spec Fact 30 (D-1) In the result JSON, `<` arrives escaped as `<`.**
  It actually appeared in the exception's stack as `at <anonymous>:1:31`.
  Therefore, if you decode JSON yourself, **decoding `\uXXXX` is not insurance but mandatory**.
  When D-2 takes innerHTML, every tag will come in this form.

- **Spec Fact 31 (D-1) JSON.stringify(undefined) returns undefined, so the value key
  disappears altogether from the result object.**
  Only `{"ok":true}` arrives. This can be used as a signal that "the return value was undefined"
  (not an error). EvalSync detects this and returns the string undefined.

- **Spec Fact 32 (D-1) While JS is blocking the thread, ExecuteScript's completion notification
  does not arrive, and arrives all at once the moment the block is released.**
  In a verification that timed out a 3-second busy loop at 1 second, the completion notification
  arrived late after the cutoff (cb19). The same structure as spec fact 27 (while hostObjects.sync is running, the calling pane's JS is
  completely blocked), and **the waiting side must always have
  a way to handle the timeout and late arrival.**
  Note: the fact that a JS syntax error returns as errorCode=0 + resultJson="null" was confirmed in 9.9c.
    We reconfirmed the same behavior in D-1 too (EvalSync treats it as syntax-error-or-null).

### Notes for D-2 onward

- Strings **come back quoted**. In D-2, decide the type per property of `Wv2Element` and strip.
- `jsExpr` is an "expression", not a "statement". Wrap multiple statements in an immediately invoked function (inheriting the Safe wrapper's constraint).
- The element registry (`__wv2els[]`) **disappears on page navigation**, so have `Wv2Element`
  hold three things, "Pane reference + handle + generation number", so that stale references can be detected.
- `QuerySelectorAll` takes **just one round trip** if the JS side registers all hits at once and returns the handle array.
  The return value is VBA's standard `Collection`. We do not create a `Wv2Elements` class.
- We do not create `Wv2Document` (Wv2Pane already has title / URL / readyState / DOMContentLoaded,
  and interposing one would only duplicate). The search entry point is attached directly to `Wv2Pane`.
- ★Do not break/step-execute even during EvalSync's DoEvents loop★ (a supplement to spec fact 20).
- ★When a JS result is needed inside a handler, call after one beat with SafeTimer★
  (the same pattern as `RequestNewTabAsync`). A direct call fails immediately via the in-callback guard.

---

## Stage D-2 — Element registry and reading with Wv2Element

### Where it fits

Since D-1 demonstrated the path of "using an asynchronous API with a synchronous face", this is the production layer that sits on top of it.
It is responsible up to **grabbing an element and reading its value**; writing (`.Value =` / `.Click`) is sent to D-3,
and collections (`QuerySelectorAll`) and lifetime management to D-4.

### Issues and decisions (9 items agreed in one go)

| # | Issue | Decision |
|---|---|---|
| 1 | Timing of injecting the JS-side registry | **Plan c: do not inject.** Make the search expression itself take the form "create it myself if absent" |
| 2 | Issuer of the generation number | **Plan b: JS issues it** (`'g' + Date.now() + '-' + Math.random()`). The authority for judgment is always the JS side |
| 3 | Release of handles | **Plan a: do not release.** Just one `ClearElementRegistry` for manual cleanup |
| 4 | How to return when not found | **Plan a: `Nothing`.** Success is told by `LastEvalOk` |
| 5 | Public members of `Wv2Element` | 8 read kinds (below). Writing is D-3 |
| 6 | Where to put quote stripping | **Plan b: consolidate in `Wv2Json.bas`** |
| 7 | Whether to include `QuerySelectorAll` | **Plan a: not included (singular only).** Only the JS side is made multi-capable |
| 8 | Verification page | **Plan b: our own HTML via `NavigateToString`** |
| 9 | Verification Subs | 3: `Test_D2_Find` / `Test_D2_Stale` / `Test_D2_Help` |

### ★Issue 1: do not inject (self-bootstrapping)★

We adopt neither `AddScriptToExecuteOnDocumentCreated` (it newly requires a HandlerKind and vtable for the completion handler)
nor injecting every time at `DOMContentLoaded` (since it is inside a handler, `EvalSync` cannot be used, and it would also fire at the Panes of the tab bar /
nav bar / settings tab). Instead, **we put a common prologue at the head of every JS expression of D-2**.

```js
if(typeof window.__wv2==='undefined'||!window.__wv2){
  window.__wv2={gen:'g'+Date.now()+'-'+Math.random(),els:[]};}
var R=window.__wv2;
```

- No new COM mechanism is needed
- Round trips do not increase (judgment, creation, and search fit in one expression)
- Whether it is erased by page navigation / SPA rewriting / `about:blank`, it is
  automatically recreated at the next access (the same self-healing as design principle 74)

**Drawback**: It cannot be used for the purpose of planting a hook **before** page load. When that is needed,
implement `AddScriptToExecuteOnDocumentCreated` as an independent stage.

### ★Issue 2: the generation is issued by JS★

The idea of counting navigations on the VBA side is **out of line with reality**: with SPA `pushState` the DOM remains but the event fires, and
conversely even when the in-page JS swaps elements no event fires.
If the JS side issues it, "the registry disappears = the generation disappears = recreation always gives a different value" is guaranteed.

`Wv2Pane.m_domGen` is a cache of "the generation last heard from JS", and **is used only for the fast path**.
If the local generation and the cache differ, stale is settled without a round trip; even when they match, JS may have recreated it,
so we always ask. **When the cache is old it always falls on the "ask" side**,
so misjudgment does not occur.

### Added Public members

**`Wv2Pane`**

| Member | Contents |
|---|---|
| `GetElementById(id)` / `QuerySelector(sel)` | Returns a `Wv2Element`. `Nothing` if absent |
| `ClearElementRegistry()` | Recreates the registry with a new generation |
| `CurrentDomGen` | Generation cache (not the authority) |
| `JsQuote(s)` *(Friend)* | Turns a VBA string into a JS string literal |
| `Elem_Read` / `Elem_CheckStale` *(Friend)* | Paths dedicated to `Wv2Element` |

**`Wv2Element` (new class)**

`TagName` / `InnerText` / `InnerHTML` / `Value` / `GetAttribute(name)` / `IsStale` /
`Handle` / `Generation` / `LastOk` / `LastError`

**`Wv2Json`**

`JsonUnescapeAt(json, startPos, outVal)` / `JsonUnescape(raw)` / `JsonPickStr(json, key, outVal)`

### ★How to read Nothing and the empty string (the convention of issue 4)★

| Situation | Judgment |
|---|---|
| `Nothing` + `Pane.LastEvalOk = True` | **Really does not exist** (normal) |
| `Nothing` + `Pane.LastEvalOk = False` | **Failure**. The reason is `Pane.LastEvalError` |
| Empty string + `Element.LastOk = True` | **Really empty** (no text / no attribute) |
| Empty string + `Element.LastOk = False` | **Failure**. The reason is `Element.LastError` (`stale` / `no-pane` / `in-callback` / `timeout(n s)`) |

### Post-implementation report (design principle 79)

- **(a) Moved `JsonPickStr` from `Wv2Pane` to `Wv2Json` (Private -> Public).**
  The agreement on issue 6 was "newly create `JsonUnescape` and rewrite `JsonPickStr` to call it",
  but `Wv2Element` also needs the same decoding, so leaving it Private in `Wv2Pane` would make 2 implementations.
  The scanning part was ported identically to `Wv2Json.JsonUnescapeAt`, and the logic was not changed by even one character.
  The existing `JsonGetStr` (does not unescape, design principle 64) was **not changed at all**.
- **(b) Added `LastOk` / `LastError` to `Wv2Element` (issue 5's 8 members + 2).**
  The read family returns an empty string on failure too, so we need a way to distinguish "really empty" from "stale".
  `Pane.LastEvalOk` is the success of the JS execution itself, so it is True when stale and cannot substitute.
- **(c) Made the return values of `Elem_Read` / `Elem_CheckStale` `Boolean` + a status string rather than an Enum.**
  There is no precedent in this project of whether a class module's `Public Enum` is visible unqualified from other modules
  (`ControllerState` and `BrowserState` are used only where they are defined),
  and a compile error would waste a round trip. This also matches D-1's style.

### Checks done before implementation

- **Python advance verification (34 cases, NG 0)** — reproduced `JsonUnescape` in Python and verified with
  `<` / quotes / backslashes / newlines / control characters / **surrogate pairs** /
  the `\uXXXX` expanded form before dropping it into VBA.
  ★At first 2 emoji cases were NG, but the cause was **insufficient construction on the checking side**★
  (in Python `str` concatenation the halves of a surrogate do not recombine). The same pattern as D-1.
- **Reconstructed the JS and ran it with node (11 cases, NG 0)** — we rebuilt the JS strings VBA assembles in Python and
  **actually ran them with node** together with a minimal DOM stub to confirm even the behavior.
  Not just a syntax check: registry self-creation, the `found` judgment, `has:false`, generation mismatch,
  missing subscript, and stale after cleanup were all crushed before execution. **This technique is effective so we will continue it.**
- Formatting, reserved words (341 declared identifiers), duplicate Public names (15 names), JS-in-VBA rules
- **Full-count cross-check with the original** — the 64 lines that vanished from `Wv2Pane` were all exactly the same as the relocated `JsonPickStr` and
  the 3 rewritten call lines, guaranteeing that there are zero unintended omissions.

### Real-machine test results (all passed)

`Test_D2_Find` / `Test_D2_Stale` / the regression `Test_D1_Eval` / `Test_D1_Guard` all
showed `[OK  ]` in the visible range, and the `in-callback` depth was also 0. Tab operations and the settings tab were also normal.

★The log scrolled away and we could not read the first half, but **the consecutive handle numbers settled its health**★
The start of `(5)` was `handle=9`, and `(1)` to `(4)` take exactly 9 elements. If even one acquisition
had failed, `D2El` would have returned a dummy and the handle would have shifted by one, so this match does not occur by chance.

**Issue 2 was proven on the real machine.** Before navigation `g1787142010489-0.8716307937600973` ->
after navigation `g1787142011015-0.07404330783805646`. Although VBA did nothing, the JS side's
registry vanished and was recreated, so the generation changed and `s:"stale"` was returned.

### Design principles established (91 to 93)

- **Design principle 91 (D-2) Judge the validity of a state on the side that actually holds that state.**
  Use a local cache only as a "fast path that **errs only toward the safe side**".
  This is the form in which the generation check is done on the JS side and `m_domGen` is used only to save round trips.
  Always confirm that when the cache is old it falls toward the "ask" side.

- **Design principle 92 (D-2) For a mechanism placed where the execution environment disappears, do not inject it;
  make the expression that uses it "prepare it itself" and let it self-heal.**
  Once you start managing injection timing, you end up having to track every path of navigation, SPA rewriting, and `about:blank`.
  If the prologue is embedded in every expression, the management itself disappears.
  The same idea as design principle 74 (global sync has self-healing against misses).

- **Design principle 93 (D-2) For an API that returns the same type on failure as on success, always prepare
  a paired entry point for asking about success.**
  A read API that returns an empty string cannot distinguish "really empty" from "failure". We placed the same pair as D-1's
  `LastEvalOk` / `LastEvalError` on `Wv2Element` too.
  **Do not be satisfied with counting members at design time; always ask "can the caller decide from this return value alone?"**

### Spec facts established (33 to 34)

- **Spec Fact 33 (D-2) In the result JSON, only `<` is escaped to `<`, and `>` arrives
  as is.**
  The asymmetric form `at <anonymous>` was observed in D-2's real-machine log.
  This is the exact form of spec fact 30. Decoding works either way so there is no actual harm, but
  when writing a simulator, match the real thing.

- **Spec Fact 34 (D-2) A VBA String is a sequence of UTF-16 code units, so surrogate pairs are restored correctly
  by adding each half with `ChrW$`.**
  `CLng("&HD83D")` becomes -10179 as a 16-bit signed value, but `ChrW$` interprets a negative value as
  65536 + n, so the D83D character is obtained, and concatenating it with the next DE00 forms the pair.
  ★Be careful when simulating in Python★ Python's `str` concatenation does not recombine them, so
  unless you hold them as a code-unit sequence and decode all at once at the end, a **false NG** appears.

### Changed files

4: `Wv2Element.cls` (**new**) / `Wv2Pane.cls` / `Wv2Json.bas` / `Wv2Tests.bas`.
`Wv2Browser.cls` was unchanged (`ActivePane` / `AddTabWithHtml` already existed).

### Notes for D-3 onward

- **Reading takes one round trip per property.** Reading a large amount in a loop is slow, so
  when you want to fetch in bulk, writing one JS expression with `EvalSync` is faster. This is D-2's trade-off.
- `QuerySelectorAll` is set up so it can be added merely by changing the JS-side registration from `push` to a loop.
- **A detached element does not become stale.** If the generation is the same and it remains in the array, it can be read.
  To see whether it has left the DOM, `document.contains(e)` would have to be added separately.
- D-4 (lifetime management) will handle the matter of `els` growing as searches repeat within the same page.
  ★We do not adopt release from `Class_Terminate`★ (we cannot control the context of Terminate, and
  it easily steps on the in-callback guard or a Pane that has already had `Ctrl_Close`).
---

## Stage D-3 ― Writing and operating + waiting on DOM conditions

### Where it fits

Since D-2 completed the means of "reading the current DOM", this takes charge of its flip side, **writing and operating**, together with
**waiting for change**, which is always needed the moment `.Click` is introduced.
A direct investment in concept ① (ease of use that excels at automatically controlling web pages).
Quiescence waiting (MutationObserver / network idle) and `QuerySelectorAll` / lifetime management are sent to D-4.

### Issues and decisions (8 items agreed in one go)

| # | Issue | Decision |
|---|---|---|
| 1 | Scope of this round | **Plan A: writing and operating + waiting on DOM conditions.** If waiting is sent to the next stage, automation is unusable |
| 2 | Shape of waiting | **Plan C: explicit waiting is first-class, implicit waiting is opt-in** (`ImplicitWaitSec` defaults to 0) |
| 3 | Location of the wait loop | **Polling on the VBA side** (settled by measurement of unknown 1. Below) |
| 4 | Polling interval | **Plan A: fixed 100 ms** |
| 5 | Condition to wait for | **Plan A: look only at existence.** Visibility is split into a separate API (`WaitVisible`) |
| 6 | How `.Value =` is conveyed | **Native setter -> `input` -> `change`. Fire both** |
| 7 | Shape of `.Click` | **Plan A: prefer `e.click()`, otherwise dispatch a `MouseEvent`** |
| 8 | Return value convention | Follow D-2's `LastOk` / `LastError`. `Value` stays a Property Let |

### ★We found an error in the premise at the very first step★

The `Test_D3_Probe_Promise` in the draft of the issues fired a Promise with `EvalSync`, but
**this cannot measure unknown 1 (does ExecuteScript wait for a Promise?).**

`EvalSync` fires `View_ExecuteScriptSafe("JSON.stringify(" & expression & ")")`. That is,
**even if the expression returns a Promise, it is `JSON.stringify`'d on the spot and collapses to `{}`.**
By the time it reaches ExecuteScript the return value is no longer a Promise, so even if it had the property of waiting we could not observe it.

-> We reorganized the measurement into two stages (not a single line of product code was added):

| | What we look at | Judgment |
|---|---|---|
| (A) Via `EvalSync` | That the Promise collapses to `{}` | Automatic judgment |
| (B) Direct call of `View_ExecuteScript` | The `resultJson=` line that `OnExecuteScriptCompleted` outputs | ★Look by eye★ |

(B)'s result does not pass through `EvalSync`'s pending table so VBA cannot receive it.
So we bracket it with the markers `★★★ from here / to here ★★★` and read it by **whether** a `resultJson=` appeared
between them.

### ★Measurement result for unknown 1: ExecuteScript does not wait for a Promise★

```
(B-1) Promise.resolve(42)         → resultJson={}  resultLen=2
(B-2) new Promise(function(r){})  → resultJson={}  resultLen=2   ← ★decisive★
```

Even when handed a **Promise that never resolves**, the completion notification came immediately. If it waited, this would have
stayed silent, so this settles it (spec fact 49). Therefore for issue 3 we implemented, as in the draft's skeleton,
**polling on the VBA side**.

★Even if it had "waited", to use it we would have had to rebuild D-1's wrapping (the form that bites `JSON.stringify` first)★,
so the price of the "wait on the JS side" plan was higher than issue 3's estimate.

### Added Public members

**Wv2Element (writing and operating)**

| Member | Form | Meaning |
|---|---|---|
| `Value = v` | Property Let | Writes a value. ★There is no return value, so success is via `LastOk`★ |
| `Click()` | `Boolean` | Clicks |
| `SetAttribute(n, v)` | `Boolean` | Writes an attribute |
| `LastInfo` | `String` | The path chosen by the previous operation (`setter` / `direct` / `click` / `dispatch` / `ok`) |

**Wv2Pane (waiting on DOM conditions)**

| Member | Form | Meaning |
|---|---|---|
| `WaitFor(selector, timeoutSec)` | `Wv2Element` | Waits until it appears (default 10 seconds) |
| `WaitGone(selector, timeoutSec)` | `Boolean` | Waits until it disappears (default 10 seconds) |
| `ImplicitWaitSec` | `Single` | Implicit waiting with ★default 0 = disabled★ (opt-in) |

Internally there are 5: `Elem_Exec` (Friend, the "firing side" paired with `Elem_Read`) / `FindWithWait` /
`ExistsNow` / `PollPause` / `ElapsedSec`.

### ★The meaning of success (the easiest place to misunderstand in D-3)★

**`True` means "the JS ran without an exception", not "the intended effect occurred".**
`Click`ing a disabled button raises no exception so `True` is returned. The same for writing to a readonly input.
`.Value =` on a `div` merely grows one custom property,
yet `LastOk = True` as well (the path is `direct`). **When the effect is required, read it back and confirm.**

### ★Issue 6: what `.Value =` does (the core of SPA)★

If you plainly write `e.value = x`, React / Vue **do not notice** the change of value.
The framework layers its own value setter (value tracker) over the element instance,
and an assignment that goes through it gets recorded as "a value it already knows".

```js
var d = Object.getOwnPropertyDescriptor(e.constructor.prototype, 'value');
if (d && d.set) { d.set.call(e, v); } else { e.value = v; }
e.dispatchEvent(new Event('input',  { bubbles: true }));
e.dispatchEvent(new Event('change', { bubbles: true }));
```

What is obtained from `e.constructor.prototype` is the **native** descriptor, so
it can bypass the instance-side tracker (spec fact 50). `input` is watched by the React family and
`change` by the jQuery family, so **we fire both**. In plain HTML `change` fires one extra time,
but we judged that the accident of "some pages do not react" costs more.

### ★Design of waiting★

**The meaning of the return values follows D-2's convention (design principle 93) as is.**

| | `LastEvalOk = True` | `LastEvalOk = False` |
|---|---|---|
| `WaitFor` is `Nothing` | Did not appear in time (normal) | Failure (invalid selector / in-callback) |
| `WaitGone` is `False` | Did not disappear in time (normal) | Failure |

★**The moment it is known to be a failure, it exits without waiting**★ (design principle 104). Selector syntax errors and
in-callback (spec fact 14) do not get fixed by waiting. If it persisted, calling from inside an event handler would
freeze Excel for the specified seconds. On the real machine an invalid selector with a 5-second timeout returned in **0.05 seconds**.

**Why implicit waiting is not the default** (issue 2): If it waited by default, confirming "it really isn't there" would
always mean waiting out the timeout, muddying the meaning of D-2's decision "`Nothing` + `LastEvalOk=True` = really does not exist".
Only those who want the Selenium style raise `ImplicitWaitSec`.

### Post-implementation report (design principle 79)

1. ★**Added `Wv2Element.LastInfo`**★ (issue 8's 3 members + 1)
   Whether `.Value =` went **via the native setter (`setter`) or a plain assignment (`direct`)** is
   the only clue to whether it worked on an SPA, and without it unknown 3 could not be verified.
   It is a record of "the previous operation" that returns to an empty string when a read-family member is called, and is not used to decide success.

2. ★**`WaitGone` uses a dedicated existence check (`ExistsNow`) that does not push onto the registry**★
   Running `FindOne` would push the not-yet-vanished element onto `R.els` at every poll
   (70 of them after waiting 10 seconds). Since there is no need to grab it, we prepared a separate expression that returns only `{g:R.gen,found:!!e}`.

3. ★**Corrected for `Timer` crossing midnight (`ElapsedSec`)**★
   `Timer` returns to 0 at midnight, so a plain subtraction goes negative and **the wait loop never
   exits** (spec fact 53). We correct by adding the crossed amount (86400).
   Since the real-machine test ran at 1:53, this is not someone else's problem.

4. **Silenced the "not found" log during polling** (added `quiet` to `FindOne`).
   If it appeared on every round trip it would fill the Immediate window. An ordinary
   `QuerySelector` without setting `ImplicitWaitSec` still prints, as in D-2.

5. **Changed `QuerySelector` / `GetElementById` to go via `FindWithWait`.**
   However, `ImplicitWaitSec` defaults to 0, so unless it is set the behavior is the same as D-2 (answers in one round trip).

### Checks done before implementation ― ★we reassembled the JS from the VBA source and ran it against a real DOM★

This PC had no node (cscript's JScript is ES3-equivalent with no `Object.defineProperty` or
`JSON`, and no DOM, so it is no substitute). Instead we **ran Edge headless**
(`msedge --headless=new --dump-dom`. The same Chromium as WebView2).

How the harness was built: we wrote a mini evaluator that **reads the concatenation expressions in `src/*.cls` themselves and expands them to JS**
(supporting string literals / variables / calls of `JsQuote` `CStr` `JsRegistryPrologue`), and reproduced even the wrapping of
`View_ExecuteScriptSafe` + `EvalSync`. The string from `JSON.stringify(eval(wrapped))` on the page side is
**exactly the `resultJson` that VBA receives**, so we could reproduce the equivalent of `JsonPickStr` in Python and cross-check end to end.
The verification page (`BuildD3ProbeHtml`) was also extracted from VBA, so **the implementation and the verification cannot drift apart**.

All 30-plus were OK. The breakdown: the setter path / round trips of symbols, quotes, and backslashes /
`input`+`change` 3 times each / ★a React-style tracker not noticing a plain assignment but noticing D-3's way of writing★ / both `click` and `dispatch` paths /
`isTrusted=false` / round trip of attributes / an invalid attribute name giving False via `InvalidCharacterError` / writing to a div being `direct` /
select and textarea / stale after the registry is rebuilt.

★**This harness could not be used for D-3b (waiting)**★
When tarboh's Edge is running, `msedge.exe` hands the command line to the existing instance and exits immediately
(spec fact 55). For `ExistsNow`'s JS we limited ourselves to assembling it for visual inspection + a bracket/quote
matching check, and left the verification of meaning to the real machine's `Test_D3_Wait`.

### Real-machine test results (all passed)

| Test | Result |
|---|---|
| `Test_D3_Probe_Promise` | Unknown 1 settled (above) |
| `Test_D3_Write` | ★33 OK / FAIL 0★ |
| `Test_D3_Framework` | ★7 OK★ (notified 0 -> 1, trackedSet stays at 1) |
| `Test_D3_Wait` | ★29 OK / FAIL 0★ |
| Regression `Test_D2_Find` / `_Stale` / `Test_D1_Eval` / `_Guard` | ★31 / 8 / 20 / 3 all OK★ |

**The measured wait times lined up cleanly** (the trick is a 800 ms `setTimeout`):

| Scene | Limit | Measured | Reading |
|---|---|---|---|
| Element already present | ― | **0.04 s** | Returns without waiting |
| Appears late | 5 s | **0.89 s** | Delay of 90 ms = about 1 polling cycle |
| Never appears | 1 s | **1.02 s** | Stops as specified |
| ★Invalid selector★ | 5 s | **0.05 s** | ★Gave up without waiting★ |
| `WaitGone` | 5 s | **0.88 s** | Same |
| `ImplicitWaitSec=5` | 5 s | **0.88 s** | The opt-in works |

In `Test_D3_Wait` (2), **we could write straight into the element we waited for and grabbed**.
It is evidence that waiting (D-3b) and writing (D-3a) are connected, and
usage such as `p.WaitFor("#result", 5).Value = "..."` really works.

### ★Changed where verification is output (to K-1's log)★

The 40-plus checks of `Test_D3_Write` could not be read in the Immediate window. Each `EvalSync` round trip outputs
15 lines of plumbing log, so 600-plus lines flowed past and overflowed the buffer.

So we made **the checks of all D-1 / D-2 / D-3 tests output to `Wv2Log`**.
`Wv2Log` also fires `Debug.Print` internally, so what shows in the Immediate window does not change.
At the end of each test a one-line tally is output:

```
★33 checks: OK 33 / FAIL 0★
```

We also merged the check helpers into one pair (`D2Bool` / `D2Eq` and D-3's version were the same thing, so they became
`TestBool` / `TestEq`). Now **pass/fail can be told just by searching the log file**
(design principle 105). `Test_D2_Find` has the tally in its mid-procedure `Exit Sub` path too.

### Gap in the procedure sheet (★discovered on the real machine★)

The preparation steps of `Test_D*_Help` were **missing `UserForm1.Show vbModeless`**.
`hWnd_Frame` is obtained with `Frame1.[_GethWnd]`, so unless the form's window actually exists,
it stays 0 and `Browser.Init` fails (spec fact 54). tarboh had been filling it in out of habit, so
it did not become an accident. We fixed the 3 Helps of D-1 / D-2 / D-3.

### Design principles established (103 to 105)

- **Design principle 103 (D-3) Before measuring, confirm that the property you want to measure is preserved by the measurement path.**
  We tried to measure "does ExecuteScript wait for a Promise" with `EvalSync`, but `EvalSync` bites
  `JSON.stringify` first, so the Promise collapsed to `{}` midway along the path.
  **Cross-check the probe itself** to see whether it has become a form that always returns the same answer.

- **Design principle 104 (D-3) A waiting API distinguishes "not yet" from "failure", and stops
  waiting the moment it is known to be a failure.**
  A retry is a mechanism that bets on "it will be fixed as time passes", so it must not be applied to things that time does not fix
  (syntax errors / permissions / in-callback). Without this separation, we get the worst behavior in which
  the more wrong the call, the longer it freezes.

- **Design principle 105 (D-3) Output the verification's checks not to a place that scrolls away but to a place that can be searched later.**
  The Immediate window easily overflows with plumbing logs. If the checks cannot be read, it is the same as not having verified.
  **Merely changing the output destination changes the cost of re-running verification from "run everything again" to
  "open a file".**

### Spec facts established (49 to 55)

- **Spec Fact 49 (D-3) `ExecuteScript` does not wait for a Promise.**
  Even if the return value is a Promise, the completion notification returns **immediately** as `{}`. Confirmed with a Promise that never resolves.
  Therefore the form "wait on the JS side and return the result" cannot be written, and waiting becomes
  polling on the VBA side.

- **Spec Fact 50 (D-3) A value tracker such as React's layers `value` over the element instance, so
  unless you `call` the native prototype setter, the change goes unnoticed.**
  `e.value = x` goes through the tracker, and the tracker records it as "a value I wrote myself", so
  even if `input` is fired right after, it judges there is no difference.
  With `Object.getOwnPropertyDescriptor(e.constructor.prototype,'value').set.call(e, x)`
  the tracker can be bypassed and the difference is seen at `input`. **Confirmed on the real machine: notified 0 -> 1.**

- **Spec Fact 51 (D-3) The `isTrusted` of a synthetic event dispatched from JS is false.**
  The same for `e.click()` and `new MouseEvent('click')`. It cannot work in principle on a page that looks at `isTrusted` and rejects. The values observed on the real machine were `click/true/false` (type/bubbles/isTrusted).

- **Spec Fact 52 (D-3) One `EvalSync` round trip measures about 38 ms.**
  33 checks, about 40 round trips, took 1.41 seconds. An order of magnitude slower than issue 4's estimate (a few ms).
  If you poll at 100 ms intervals, the **effective judgment interval becomes about 140 ms**.

- **Spec Fact 53 (D-3) `Timer` returns to 0 at midnight.**
  A wait loop that measures elapsed time as `Timer - t0` goes negative across a date change and **never exits**.
  If negative, correct by adding 86400.

- **Spec Fact 54 (D-3) The HWND of a control on a UserForm cannot be obtained until after `Show`.**
  `Frame1.[_GethWnd]` returns 0, and with `hWnd_Frame` left at 0, `Browser.Init` fails.
  `StartWebView2_Full` presupposes `UserForm1.Show vbModeless`.

- **Spec Fact 55 (D-3) When Edge is running, `msedge.exe` hands the command line to the existing
  instance and exits immediately.**
  Separating `--user-data-dir` does not help. The symptom is "exit code 0, empty standard output", and
  verification by `--headless --dump-dom` becomes impossible. **The way to tell is `msedge --version`**
  (if nothing is printed, it was handed off). `msedgewebview2.exe` is the same.

### Changed files

3: `Wv2Pane.cls` / `Wv2Element.cls` / `Wv2Tests.bas`. There are no new modules.
`Wv2Browser.cls` / `Wv2Json.bas` were unchanged.

### Notes for D-4 onward

- **Quiescence waiting (MutationObserver + fetch/XHR counters) is D-4's main event.**
  What D-3 can wait for is only up to "DOM conditions". It cannot wait for "until all communication ends and things go quiet".
  ★For wrapping `fetch` / `XHR`, first decide the consistency with design principle 92 (do not inject)★
  The key is whether the counters can be recreated every time by the self-generating prologue.
- **`WaitVisible` (wait until visible) is unimplemented.** Since issue 5 settled on "look only at existence",
  add it as a separate API when needed (`offsetParent !== null`, etc.).
- **`ImplicitWaitSec` does not apply to `WaitFor` / `WaitGone`.** This is to prevent explicit waiting from
  persisting twice. This asymmetry is intentional, so if you change it, change it together with the reasoning.
- **Polling is a long `DoEvents` stretch** (as in unknown 2). If you wait 10 seconds, you spend 10 seconds
  in the danger zone of spec fact 20 the whole time. We wrote this in the help of `Test_D3_Wait` too.
- `Wv2Element` is still **singular** only. `QuerySelectorAll` and lifetime management remain in D-4.

---
## Stage D-4 ― Quiescence waiting (judging when an SPA has "calmed down")

### Where it fits

What D-3 introduced was waiting **for cases where you can write what to wait for** (`WaitFor` / `WaitGone`).
The real subject pointed out by an outside recommendation — what to do **when the target to wait for cannot be specified** — was
untouched. D-4 fills that in.

**It was cut into 6 steps, a to f, each verified independently on the real machine.**

| Step | Contents | Real machine |
|---|---|---|
| D-4a | In-page probe + health check | 28 OK |
| D-4b | `WaitSettled` + noise exclusion + auto-waiting | 22 OK |
| D-4c | Explicit signals (`Arm*`) | 31 OK |
| D-4d | ★Reconnaissance of Google Maps★ (observation only) | 5 structure checks OK |
| D-4e | Diagnostic log / in-flight ledger / URL signal | 21 OK |
| D-4f | Excluding lingering requests (`StaleInflightMs`) | Included in the above |

### Issues and decisions (10 items agreed in one go)

| # | Issue | Decision |
|---|---|---|
| 1 | Scope | Quiescence waiting alone (`QuerySelectorAll` / lifetime management are D-5) |
| 2 | How to place the probe | **Extension of the self-generating prologue** (maintaining design principle 92) |
| 3 | Health check | **Include** |
| 4 | Judgment expression and defaults | `inflight=0` and quiet 500 ms / limit 10 seconds / polling 100 ms |
| 5 | Noise exclusion | Include (URL patterns + selectors). Blocking is not done |
| 6 | Explicit signals | `ArmContentSignal` / `ArmNetworkSignal`. One-shot |
| 7 | Coupling of waiting and operation | **`WaitSettled` is first-class + `AutoWaitAfterAction` defaults to False** |
| 8 | Return value and naming | Follow D-3's convention. **`WaitSettled`** rather than `WaitIdle` |
| 9 | Diagnostic log | Include in a minimal form |
| 10 | Verification | Our own HTML + ★the real site is Google Maps★ |

### Added Public members

| Member | Meaning |
|---|---|
| `WaitSettled(timeoutSec, stableMs)` | Wait until communication and the DOM calm down |
| `LastSettleInfo` | Breakdown of the wait (`stable q=611ms(eff=611) slack=111ms polls=8 ...`) |
| `AutoWaitAfterAction` | ★Default False★ Wait automatically after an operation |
| `AddIgnoreNetwork` / `ClearIgnoreNetwork` / `IgnoreSelectors` | Exclusion from the quiescence judgment |
| `ArmContentSignal` / `ArmNetworkSignal` / `ArmUrlSignal` / `DisarmSignals` | Explicit signals |
| `StaleInflightMs` | ★Default 10 seconds★ The threshold beyond which lingering requests are not counted |
| `SpaProbeState` / `SpaProbeHealthy` / `SpaProbeReset` | State of the probe |
| `SpaProbeLogging` / `SpaProbeDrainLog` | Diagnostic log |

### ★How the probe (window.__wv2p) is built★

It wraps `fetch` and `XMLHttpRequest.prototype.open/send` to count in-flight requests, and
holds the time of the last DOM change with a `MutationObserver`. **It is not injected** (design principle 92).
The expression that is called prepares it itself, and ★performs a health check at every call and rebuilds it if broken★.

What the health check looks at: version number / `window.fetch === p.fw` / `XHR.open===p.xo` and
`send===p.xs` / liveness of `MutationObserver`. If even one is missing, it **re-wraps the current `window.fetch`**
and rebuilds (it wraps along with the wrapper the framework layered on, so it does not break the other side's features).

★Communication that started before the probe was placed cannot be counted★ This is by specification. Combined with D-4c's
operation of "arm just before the operation", there is no actual harm.

### ★The quiescence judgment expression★

```
Quiet = (in-flight excluding lingering ones is 0) and (stableMs or more since the last DOM change)
Defaults: stableMs=500 / limit 10 seconds / polling 100 ms / lingering threshold 10 seconds
```

★The quiet window is measured from "the moment `WaitSettled` is called"★ (D-4b's post-implementation report).
Even if it was quiet before the call, it does not suddenly hold. Otherwise **"has not started yet" and
"has already finished" cannot be distinguished** (it is normal for the effect of an operation to start tens of ms late).
Therefore ★it watches for at least stableMs★. However, processing that starts after stableMs is
missed in principle — that is the job of explicit signals.

### ★★ Even if WaitSettled is True, it is not proof of "finished" ★★

It is **merely quiet**. It cannot distinguish whether the app is quiet because it ignored the request or because it finished.
That is why we named it `WaitSettled` rather than `WaitIdle` (issue 8).
For important operations, **arm an explicit signal just before the operation**.

### Explicit signals (arm, then act)

| API | What it waits for |
|---|---|
| `ArmContentSignal(selector)` | A ★pre-existing★ subtree is rewritten. **If absent at arm time, immediately False** |
| `ArmNetworkSignal(pattern)` | A request whose URL partially matches completes |
| `ArmUrlSignal(fragment)` | ★The URL comes to contain that string★ (supports pushState) |

- **One-shot.** Consumed by the next single `WaitSettled` (whether it succeeds or times out)
- **When it hits, the quiet window is measured again** (take the whole thing once more after the evidence)
- ★**If the arm vanishes, do not silently repair it**★ If the probe is gone along with the arm due to navigation, fail immediately with `signal-lost`.
  If it just kept waiting, it would turn into a "wait with no condition" and return success the moment things went quiet
- ★**Even for excluded requests, the explicit-signal judgment is still made**★ Exclusion means only
  "remove from the quiescence judgment", not "pretend we did not see it"

### ★D-4d: What we learned from reconnaissance of Google Maps★

**The real site is Google Maps, and the aim is address search** (motivated by tarboh's practical use in his work).

#### Address search went through ― even coordinates could be obtained

```
.Value = "東京都千代田区丸の内1-9-1" (path=setter) → click button[aria-label='検索']
→ title = 〒100-0005 東京都千代田区丸の内１丁目９−１ - Google マップ
→ url   = .../place/.../@35.6824246,139.7682844,17z/...
```

★35.6824246, 139.7682844 = Tokyo Station★ D-3's writing and clicking and D-4's waiting
meshed on a real site. **The coordinates are in the URL, so the marker of completion and the value we want are the same thing**.

#### The prediction was wrong ― Maps was quiet

In 15 seconds of doing nothing there was **1 DOM change and 1 communication**. The map is drawn with canvas / WebGL, so
neither tile fetches nor animation show up in the probe. The assessment that "it never quiets down so exclusion is essential" was wrong.
Right after a search it moves as `m=36 -> 198 -> 202` and **converges in about 2 seconds**.

#### ★The real problem: lingering streaming requests★

```
timeout inflight=1 q=11399ms polls=142 signal=url:hit
| running: xhr 20060ms /search?tbm=map&authuser=0&hl=ja&gl=jp&pb=!4m12!...
```

The URL signal hit and the DOM was quiet for 11 seconds, but **one XHR that keeps streaming search results stays
open**, so `inflight=0` never holds. A real example of the "page that never becomes idle" that the reference product warned about.

★We could pin this down because we introduced the in-flight ledger (D-4e)★
Until then all we could say was "for some reason quiescence never comes".

#### Countermeasure in two tiers (D-4f)

| | Means | Measured on Maps |
|---|---|---|
| General | `StaleInflightMs` (default 10 seconds) | `WaitSettled` succeeds in **10.05 seconds** |
| Precise | `AddIgnoreNetwork "/search?tbm=map"` | Succeeds in **0.61 seconds** |

Since we **tilted the default toward the general side**, it works for the time being even on unknown sites. When speed and precision
are needed, exclude by name. ★What is being discarded★: a request that is merely truly slow also stops being
counted once it exceeds the threshold. However, at that point the DOM has also been still for the quiet window, so
it does not contradict `WaitSettled`'s meaning of "it looks calmed down".

#### ★The diagnostic log showed cause and effect★

```
19583 DOM  x2                              ← wrote the address
19585 NET+ /s?tbm=map&suggest=p&q=東京...   ← suggest request
19690 NET- 105ms                           ← response
19700 DOM  x357                            ← rendering of suggestions
19723 NET+ /search?tbm=map&pb=...           ← ★the search proper = lingers★
20250 NET+ /maps/_/MapsWizUi/data/batchexecute?rpcids=hspqX
20346 NET- 95ms
```

A **structure identical** to the reference product arming `rpcids=tDoGIe` on Google Flights appeared in our log too.
If you arm `batchexecute?rpcids=...`, you can catch "the moment the search returned" precisely. ★We can now decide
what to wait for by observation rather than guesswork★

### ★Do not write selectors by guesswork★ (D-4d's accident)

At first we waited 30 seconds for `#searchboxinput` and did not find it. We built a DOM dump tool
(`Test_D4_Dom`) on the spot and looked in:

```
INPUT id=ucc-1 name=q type=text cls=UGojuc ...   ← ★the id is auto-generated★
BUTTON aria=検索 type=submit cls=mL3xi
```

**The former id was gone and is now `ucc-1`**. `name='q'` and `aria-label='検索'` remained.
We changed it to try candidates in order and ★leave a log of which one hit★ (`D4FindFirst`).
`Test_D4_Dom` remains as a tool usable whenever automating any site in the future.

### Post-implementation report (design principle 79)

1. ★**Made the quiet window be measured from "the time of the call"**★ (above)
2. ★**Always send the exclusion list even when empty**★ If we make it "don't send because it's empty", **the previous setting
   remains on the page side**, and `IgnoreSelectors = ""` or `ClearIgnoreNetwork` would have no effect.
   One FAIL in the real-machine test was this
3. **Silenced the "not found" log during polling** (`quiet` in `FindOne`)
4. ★**Added `D4WaitPane` on the verification side**★ Right after opening a tab, `EvalSync` fails with `no-view`, and
   `WaitFor` / `WaitSettled`, under the policy of "do not wait on failure" (design principle 104),
   give up immediately. In rounds that open a real site, wait for the Pane to be ready first (we also put a note in the
   comment of `WaitFor`)
5. **Newly created `Test_D4_Dom` (DOM dump).** It was not among the issues, but once a selector misses,
   the only path left is "fix it by guesswork"
6. ★**Added a read-only guard to `import.ps1`**★ (an improvement to axis E. Below)

### ★Closed one hole in the tooling (axis E)★

If Excel has the workbook open, `Workbooks.Open` **succeeds read-only**.
Then both the replacement and the post-verification succeed against "the in-memory copy", and
★nothing reaches the disk yet it reports "written back (cross-check OK)"★.
We actually stepped on this in D-4b's write-back. `import.ps1` now looks at `wb.ReadOnly` at `-Apply` time and
aborts.

### Checks done before implementation

**62 cases** in a real Edge headless DOM. The method of reassembling the JS from the VBA source is the same as D-3.

★This harness caught 2 real bugs★

1. `repairs:(prev&&prev.repairs?prev.repairs+1:0)` is **falsy when `repairs` is 0 and
   collapses to 0**. The first auto-repair would have never been recorded
2. `SpaProbeReset` set `window.__wv2p=null`, so **the generation went back to 1** and
   we could no longer trace when it was rebuilt. We changed it to "break the version number and make it rebuild"

★Conversely, there was also one bug the harness missed★ (post-implementation report 2 above).
Because `SpaProbeApplyJs` was stubbed with an empty string, **the VBA-side assembly logic was never
passed through even once**. We were breaking design principle 98 (verify by the same path as production) ourselves.
Immediately afterward we fixed the harness to use the production path too.

### Design principles established (106 to 109)

- **Design principle 106 (D-4) Settings held by the other side are sent every time, even when empty.**
  "Don't send because it hasn't changed" or "don't send because it's empty" means **the previous state
  remains on the other side**. If the sender is to be the sole authority, overwrite every time.

- **Design principle 107 (D-4) For identifiers that grab an external target, do not write them by guesswork; decide by observing.
  Build the observing tool first.**
  A real site's ids change to auto-generated. If you fix it by guesswork, the next time it breaks you
  cannot tell what changed. At the very least, line up candidates and ★leave a record of which one hit★.

- **Design principle 108 (D-4) Prepare both a general compromise and a precise specification, and tilt the default toward the general side.**
  Excluding lingering requests is this. The default (a cutoff by time) makes it work even on unknown sites, and
  exclusion by name or explicit signals win back precision and speed.
  The condition is ★to write in the documentation what the compromise discards★.

- **Design principle 109 (D-4) "Calmed down" can be judged only within what is observable.
  State plainly that rendering that cannot be observed cannot be waited for.**
  Canvas / WebGL appear neither in the DOM nor in communication. Do not return something
  that cannot be judged "pretending it can be judged".

### Spec facts established (56 to 61)

- **Spec Fact 56 (D-4) The Google Maps map is drawn with canvas / WebGL, and neither tile fetches nor
  animation show up in the MutationObserver or in fetch/XHR.**
  In 15 seconds of doing nothing there was 1 DOM change and 1 communication. Conversely,
  ★"has the map finished drawing" cannot be judged from the DOM side★

- **Spec Fact 57 (D-4) `MutationObserver` record delivery is a microtask, so
  even reading within the same JS turn does not reflect it.**
  `EvalSync` ends the turn so it is no problem on the real machine, but
  if you write a "change, then read" verification within a single expression, a **false NG** appears.

- **Spec Fact 58 (D-4) Google Maps search keeps `/search?tbm=map` open while streaming results.**
  It stays in-flight for 20 seconds or more. A form you are sure to step on
  if you try to measure an SPA's "completion" by the completion of communication.

- **Spec Fact 59 (D-4) The ids of real sites can change to auto-generated.**
  Maps' search box had gone from `#searchboxinput` to `id=ucc-1`.
  The `name` attribute, `aria-label`, and structure live longer.

- **Spec Fact 60 (D-4) The standard output of `msedge --headless` is empty when launched directly from Git Bash or Python's
  `subprocess`.**
  It can be obtained via PowerShell's `Start-Process -RedirectStandardOutput`.
  The symptom, "exit code 0, no output", is the same as spec fact 55 (handoff while Edge is running), so
  it is confusing.

- **Spec Fact 61 (D-4) When Excel has the workbook open, `Workbooks.Open` succeeds
  read-only.** Writing and post-verification both succeed against the in-memory copy, so
  ★"written back" is reported even though it has not reached the disk★.

### Changed files

`Wv2Pane.cls` / `Wv2Element.cls` (one line of auto-wait hook) / `Wv2Tests.bas` /
`tools/import.ps1` (read-only guard). There are no new modules.

### Notes for D-5 onward

- ★**Completion of canvas rendering cannot be observed**★ (spec fact 56). It matters for the
  use of saving a map as an image. **Neither printing nor screen capture is implemented**, so if we do it,
  `CapturePreviewAsync` (one completion handler + IStream) is the direct approach.
  ★Rendering completion can be judged from outside by "take two shots with a short gap, and if identical it has stopped"★
- `QuerySelectorAll` and element lifetime management remain in D-5 (carried over from D-2).
- `WaitVisible` (wait until visible) is also unimplemented.
- ★Shadow DOM and iframes are not reached★ (as in D-3's notes).
- **The next step for business use is to make "address -> coordinates" a single function.**
  All the parts are in place (`.Value =` / `.Click` / `ArmUrlSignal` / `WaitSettled` /
  extracting coordinates from the URL).

---
## Stage D-5 ― Practical use: address -> coordinates (Wv2Maps)

### Where it fits

★The first example of folding the parts built on axis D into a form usable as is in business.★
Since D-4d demonstrated that Google Maps address search goes through on a real site, we turned that procedure into
a single function. **We added no new mechanism at all** ― only a combination of D-2 to D-4.

### Newly added module `Wv2Maps.bas`

| Member | Meaning |
|---|---|
| `MapsOpen(browser)` | Opens a Maps tab and waits until it is operable. Returns a `Wv2Pane` |
| `MapsGeocode(pane, address, outLat, outLng, outName)` | Searches an address and obtains the coordinates and the normalized name |
| `MapsLastError` | The reason for a failure |

```vba
Dim p As Wv2Pane, lat As Double, lng As Double, nm As String
Set p = Wv2Maps.MapsOpen(UserForm1.CurrentBrowser)
If Wv2Maps.MapsGeocode(p, "東京都千代田区丸の内1-9-1", lat, lng, nm) Then
    Debug.Print lat & "," & lng & "  " & nm
End If
```

★The same Pane can be reused★ so addresses can be processed one after another (batch).
`outName` receives the address Maps normalized (`〒100-0005 東京都千代田区丸の内１丁目９−１`),
so **it can also be used to normalize notation variants** as a bonus.

### ★Why it was isolated into a separate module★

Selectors and URL shapes **change at Google's convenience without notice** (design principle 75). If mixed into the product core
(`Wv2Pane` / `Wv2Element`), then when it breaks we would not know where to fix.
**Confine what changes to one place.**

This judgment actually paid off. The handling of Maps search leaving a streaming request open
(spec fact 58) — `AddIgnoreNetwork "/search?tbm=map"` — is also
**Maps-specific knowledge**, so we put it inside `MapsOpen`. As a result, the time per item went from
★10.4 seconds to 2.0 seconds★ (because it no longer waits out `StaleInflightMs`' 10 seconds every time).

### ★A bug stepped on on the real machine: the coordinates lag by one item★

The first implementation waited for "completion when the URL differs from the URL before the search". When we looked up 3 in a row:

| | Name | Coordinates |
|---|---|---|
| (1) 東京駅 (Tokyo Station) | Correct | Correct |
| (2) 大阪城 (Osaka Castle) | Correct | ★Still Tokyo Station★ |
| (3) 札幌 (Sapporo) | Correct | ★Still Osaka Castle★ |

Cause: when you press search, Maps **first changes to `/maps/search/<search term>/@<previous center>`**,
and then settles on `/maps/place/<name>/@<new coordinates>/data=!3d<lat>!4d<lng>`
(spec fact 62). We grabbed at the first change, so the title was new but the coordinates were the previous ones,
an off-by-one.

★The correct wait condition is "contains both `/place/` and `!3d`"★
`!3d<lat>!4d<lng>` is **attached only once the place is settled** (spec fact 63).

★**Only the first item had passed**★ because the first search is a transition from `/maps` and the timing
happens to mesh. **Without a test lining up 3 items, we would have missed it.**

### Where to get the coordinates from

- First choice: `!3d<lat>!4d<lng>` … ★the coordinates of the place itself★
- Second choice: `/@<lat>,<lng>,<zoom>` … the center of the map
- ★Use `Val` (not `CDbl`)★ `CDbl` is locale-dependent and breaks in environments where the decimal separator is `,`.
  `Val` always reads `.` as the decimal point and stops at non-digit characters (`!` or `,`).

### Return value convention (following design principle 93)

| | Meaning |
|---|---|
| `True` | Arrived at `/place/` = settled on a single item |
| `False` + `ambiguous` | Multiple candidates (stayed on `/search/`). ★The coordinates receive the map's center★ |
| `False` + `no-searchbox` / `no-search-trigger` | The structure changed (investigate again with `Test_D4_Dom`) |
| `False` + `timeout-url` / `no-coords` | Timed out / the coordinates could not be read from the URL |

### Real-machine test results

| Address | Coordinates | Time |
|---|---|---|
| 東京都千代田区丸の内1-9-1 | 35.6824246, 139.7682844 | 3.7 s |
| 大阪府大阪市中央区大阪城1-1 | 34.6852142, 135.5257536 | 2.0 s |
| 北海道札幌市中央区北1条西2丁目 | 43.0626377, 141.3539212 | 2.1 s |
| ZZZZ存在しない住所ZZZZ (nonexistent address) | ― | `False` + `ambiguous` |

**All 14 OK.** Reproducible with `Test_D5_Geocode`.

### Design principle established (110)

- **Design principle 110 (D-5) Test a stateful counterpart on the "second time".**
  A verification of just one item **structurally misses the kind of bug where the previous state lingers**.
  D-5's coordinate lag (only the first is correct) and D-4's exclusion list that could not be emptied
  (design principle 106) both appeared first on the second time.
  ★For real-site verification, line up at least 3★ ― with 2 it is hard to tell from "by chance".

### Spec facts established (62 to 63)

- **Spec Fact 62 (D-5) Google Maps search first changes to `/maps/search/<search term>/@<previous center>`
  and only later settles on `/maps/place/<name>/@<new coordinates>/data=...`.**
  "Completion when the URL changes" is too early.

- **Spec Fact 63 (D-5) In a Maps URL, `!3d<lat>!4d<lng>` is the coordinates of the place itself, and
  `/@<lat>,<lng>` is the center of the map.** `!3d` is attached only once the place is settled, so
  it can also be used as a marker for the completion judgment.

### Changed files

`Wv2Maps.bas` (**new**) / `Wv2Tests.bas`. The product core is unchanged.

### Notes

- **Sheet integration was added in D-5b** (below).
- **Routes (travel time and distance between 2 points)** are untouched. More complex than address search
  (two input fields, origin and destination, and a choice of travel mode).
- ★When `ambiguous`, there is no means of choosing from the candidate list★ In practice there are scenes where
  "take the first one if there are multiple candidates" is wanted. It can be written once `QuerySelectorAll` (D-6) is in.

---
### D-5b: Sheet integration (column of addresses -> columns of coordinates)

```vba
Wv2Maps.MapsGeocodeSheet ActiveSheet
```

| Column | Contents |
|---|---|
| A | Address (input) |
| B / C | Latitude / longitude |
| D | Normalized address |
| E | Status (`ok` / `ambiguous` / `not-found` / `no-searchbox` …) |

The column positions can be changed by arguments. It writes row by row and runs `DoEvents`, so
**the cells fill in on the screen during processing**.

★**It can be resumed from where it left off even if stopped partway**★ Rows whose status is `ok` are skipped. Even if stopped partway through 100 items,
calling it again processes only the rest. To redo, clear the status column.

### ★2 bugs stepped on on the real machine ― both appeared on the "second time"★

#### 1. A nonsense address got the previous row's coordinates

```
Row 4: 34.6852142 / 135.5257536 / ZZZZ存在しない住所ZZZZ / ambiguous
```

Because the search did not resolve, **the map did not move**, and the previous row's (Osaka Castle) position simply remained in
`/@`. The original policy "if there are multiple candidates, leave the center coordinates" backfired.
Although it can be told from the status column, ★if a `VLOOKUP` picks it up, it mixes in as wrong data★.

**The fix**: ★look at whether the coordinates changed before and after the search★

| | Meaning | Coordinates |
|---|---|---|
| The map moved | Multiple candidates (a wide area was specified, etc.). The center coordinates have meaning | Write (`ambiguous`) |
| Did not move | Nothing was found. ★Leftovers from the previous search★ | **Do not write** (`not-found`) |

#### 2. On the second run, it fell into `ambiguous` again

`MapsGeocodeSheet` opens a new tab with `MapsOpen` every time. On the second run, finished rows are skipped, so
the nonsense address was looked up **without the map ever having moved**, and there are no coordinates to serve as the comparison reference
(the URL right after opening is `https://www.google.com/maps`, which does not contain `/@`), so
"cannot judge whether it moved" -> it fell into `ambiguous`.

**The fix**: in `MapsOpen`, ★wait until `/@` appears in the URL★ (it attaches in about 1 second) +
if there is still no reference, **conservatively make it `not-found`**.
★Saying "not found" does less harm than leaving behind dubious coordinates★

### Design principle established (111)

- **Design principle 111 (D-5b) Do not leave behind a value that cannot be judged; leave the reason.**
  "Put in a plausible value just in case" means the later stage (whether a person or a formula) **believes it**.
  A blank + a reason conveys that "there is none". ★"It's fine because there is a status column" does not hold★
  ― the referencing side sometimes looks only at the value column.

### Real-machine test results (D-5b)

**All 7 OK.** Processed 3 rows (one of them nonsense), 2 rows were `ok`, and
the nonsense row had an **empty coordinate field and `not-found`**. We also confirmed ★the judgment does not waver between the 1st and 2nd runs★
(the resume path). 0.9 to 1.8 seconds per item.

### Changed files (D-5b)

`Wv2Maps.bas` / `Wv2Tests.bas`. The product core is unchanged.

---
## Stage D-6 ― QuerySelectorAll and lifetime management

### Where it fits

Clears the "returning multiple" and "growth of the registry" carried over from D-2.
It is also the premise for the open item stacked up in D-5 (choosing from Maps' candidate list).

### Issues and decisions (6 items agreed in one go)

| # | Issue | Decision |
|---|---|---|
| 1 | Scope | `QuerySelectorAll` + lifetime management (`WaitVisible` is a separate stage) |
| 2 | How to return multiple | ★**`Collection`**★ |
| 3 | Upper limit of count | Default 200. ★Notify if truncated★ |
| 4 | Lifetime management | ★**Only show it. Do not discard on our own**★ |
| 5 | How the JS side returns | Return handles as a comma-separated **string** |
| 6 | Verification | Our own HTML + the real site is Maps candidate selection |

### ★Issue 2: why Collection★

You can write `For Each el In els`, most natural in VBA. Zero items is a plain `Count = 0`.

- ★Why not an array★ **`UBound` of an empty array is a runtime error**, forcing the caller
  to defend every time.
- ★Why not a custom class (`Wv2Elements`)★ To make `For Each` work, a member attribute line
  (`Attribute Item.VB_UserMemId = 0`) is needed, but
  ★`import.ps1` drops attribute lines by design★ (spec fact 47), so it is **a poor match for the round trip**.
  ★An example where the choice of language feature was decided by a constraint of the round-trip tool★

### ★Issue 4: lifetime management is "only show it"★

`QuerySelectorAll` can stack 200 at once, so the registry grows faster than D-2 assumed.
But **if old ones are discarded on our own, the `Wv2Element` in hand silently becomes stale**.
D-2 issue 3 (do not release) and design principle 91 (the authority is on the JS side) apply as they are.

| Means | Contents |
|---|---|
| Monitoring | Show the current number with `ElementCount` |
| Warning | Notify with `LogW` when it exceeds 5000 (no discarding) |
| Manual | `ClearElementRegistry` (existing since D-2) |

### Added Public members

| Member | Meaning |
|---|---|
| `QuerySelectorAll(selector, maxCount)` | A `Collection` of `Wv2Element` (default 200 items) |
| `LastAllTruncated` | Whether the previous call was cut off at the limit |
| `ElementCount` | The number stacked in the registry (-1 on failure) |

### Implementation

We merely changed D-2's `FindOne` from `push` to a loop. The JSON returned is
`{g:generation, n:actual number on the page, k:number returned, cut:whether truncated, h:'0,1,2', tot:total stacked}`.
★Handles are a string★ (`Wv2Json` cannot handle arrays, and `Split` is smaller than
extending the parser ― issue 5).

### Real-machine test results

**All 21 OK.** The preceding Edge headless was 73 cases.

- Document order is preserved (`target,noise,slot`)
- Zero items is `Count=0` + `LastEvalOk=True` / an invalid selector is `Count=0` + `LastEvalOk=False`
- When truncated, `LastAllTruncated=True` + a warning log ("the page has 3 but stopped at 2")
- `ElementCount` 11 -> increases each time we grab -> 0 with `ClearElementRegistry` ->
  ★the elements in hand become stale★

### Changed files

`Wv2Pane.cls` / `Wv2Tests.bas`.

---
### D-6b: Take the first from the candidate list

The open item stacked up in D-5. It can be written only once `QuerySelectorAll` is in.

```vba
ok = Wv2Maps.MapsGeocode(p, "コンビニ 東京駅", lat, lng, nm, 15, True)
If ok And Wv2Maps.MapsLastPicked Then Debug.Print "候補の 1 件目です"
```

★**Write "settled" and "chosen from candidates" differently**★ The sheet's status column is also `ok` and `ok(候補1)` ("ok (candidate 1)").
If both were `ok`, things of different precision would line up with the same face and could not be told apart later
(the same line as design principle 111). The default is "do not choose".

### ★A bug stepped on on the real machine: an indirect indicator was used for the judgment★

D-5b distinguished `not-found` from `ambiguous` by "after searching, **did the map move**", but
**when the same search term is looked up consecutively the map does not move**, so even though the candidate list was
showing it became `not-found`, and ★the path that takes a candidate was never entered even once★.

The fix is to **count the candidate links directly with `QuerySelectorAll`**:

| Candidate links | Judgment |
|---|---|
| 0 | `not-found` (return no coordinates) |
| 1 or more | `ambiguous` (+ to the 1st one if `pickFirst`) |

★A form where D-6's tool replaced D-5's weakness★. The movement judgment (`MapsSameSpot`) was deleted.

### Design principle established (112)

- **Design principle 112 (D-6b) Do not use an indirect indicator for judging state. If it can be counted directly, count it.**
  "The map moved = something was found" seems plausible at first glance, but **there are other reasons for it not moving**
  (the same search was repeated / the result was at the same place).
  ★An indirect indicator "looks like it is right", so the route by which it goes wrong is hard to notice★
  If there is no means of counting, it can't be helped, but replace it once one becomes available.

### Real-machine test results

**`Test_D6_Pick` all 10 OK.**

```
Candidate links: 5 [a[href*='/maps/place/']]
★Took the 1st candidate★ (コンビニ 東京駅)
Coordinates=35.6808858,139.7676894  Name=NewDays 八重洲中央改札内  Chosen from candidates=True
```

For an address that settles (Tokyo Station) no mark is attached, and with the "do not choose" setting it stays `ambiguous`.
The regression `Test_D5_Sheet` was also all 7 OK ― no place link was mixed into a page with zero search results
(0 candidates), and it stayed `not-found` + an empty coordinate field.

### Changed files (D-6b)

`Wv2Maps.bas` / `Wv2Tests.bas`. The product core is unchanged.

---
## Stage D-7 ― Progress display and interruption

### Where it fits

D-5b got as far as sheet integration, but ★there was no way to stop it★. `MapsGeocodeSheet` keeps running until
the addresses run out, and to quit partway you had to force-quit Excel.
This is the very premise of putting it on real data, so we close it.

★**This stage was rebuilt 5 times (D-7 -> 7b -> 7c -> 7d -> 7e).**★
**Three of them were because "my premise was wrong on the real machine".** The history itself is an asset so
we keep all of it ― as long as we deal with Esc, we will certainly step on the same trap again.

### Issues and decisions (10 items agreed in one go)

| # | Issue | Decision |
|---|---|---|
| 1 | Scope | Confined to `Wv2Maps`. Do not touch the product core |
| 2 | Where to put the interrupt flag | A module variable of `Wv2Maps` + a Public Property |
| 3 | ★How to convey the interrupt★ | Esc. Details below (rebuilt 3 times) |
| 4 | Granularity of stopping | At row boundaries. ★Do not write a half-finished row★ (design principle 111) |
| 5 | How to show progress | `Application.StatusBar` + log |
| 6 | Denominator | Count in one pass beforehand. If 0 items, do not open a tab |
| 7 | How to return the interruption | Do not change the meaning of the return value. `MapsCanceled` + `MapsLastError = "canceled"` |
| 8 | Lifetime of the flag | ★Always clear at the entrance★ (design principle 106). But only "the outermost" |
| 9 | A single `MapsGeocode` | Make it effective |
| 10 | Verification | The part that needs no network, and the real site + pressing Esc by hand |

### Added Public members

| Member | Meaning |
|---|---|
| `MapsCancel` (Get / Let) | Request an interrupt / read the current request |
| `MapsCanceled` | Whether the previous call ended by interruption |
| `MapsCountRows` | Count the number of rows to process first (the entry point for producing the denominator) |
| `MapsCheckCount` | How many times the interrupt check could be called in the previous batch (for D-7d's measurement) |
| `MapsRawEscState` | The raw value of `GetAsyncKeyState` (observation port for probes) |

### ★The story of issue 3 ― we failed 3 times to catch Esc★

| Try | What we did | What happened on the real machine |
|---|---|---|
| D-7 | Look only with `GetAsyncKeyState` | ★Never fired even once★ |
| D-7b | Receive error 18 with `EnableCancelKey = xlErrorHandler` | ★Did not reach the handler★ |
| D-7c | Do not raise the error itself with `EnableCancelKey = xlDisabled` | ★The interrupt held for the first time★ |

**Why D-7 did not fire**: ★The VBA runtime's Esc arrives regardless of focus★.
The design-time premise that "`EnableCancelKey` doesn't arrive unless Excel has focus so use `GetAsyncKeyState`" was
**backwards**. It broke before we got to read it.

**Why D-7b did not arrive**: Error 18 occurs mainly
★inside WebView2's COM callbacks (`View_On*`)★, where the caller is not VBA
(via a thunk) so it does not trace back to `On Error GoTo`. `Wv2Pane`'s guard re-raised it,
and it became an unhandled dialog at the COM boundary. ★During a batch, what DoEvents processes
is exactly those callbacks, so this was not a rare case but the main path.★

**Why D-7c works**: With `xlDisabled` the error itself does not occur and execution does not stop.
If it does not stop, the next time `MapsPump` runs, `GetAsyncKeyState` can pick it up.

★**The reason to avoid a break is not only usability**★ It is exactly spec fact 20's crash window.
A break while WebView2 events are running can become an AV.

★**What we lose**★ During a batch run, Excel's Esc / Ctrl+Break do not work. However, all waits
have timeouts (`EvalSync` / `WaitFor` / `WaitSettled`), and `MapsPump` runs
every 0.2 to 0.3 seconds, so it does not freeze forever.

### ★D-7d: the status bar does not come back with `False`★

"Restore with `Application.StatusBar = False` when done" was **counterproductive in this environment**.
Measured by the probe (`Test_D7_StatusBar`):

| Value put in | Reading it back | |
|---|---|---|
| `False` | `[FALSE]` type=String | ★Stringified as a value★ |
| `Empty` | `[False]` type=Boolean | ★Correct★ |
| Empty string `""` | `[False]` type=Boolean | Correct |
| `vbNullString` | `[False]` type=Boolean | Correct |
| `CVar(False)` | `[FALSE]` type=String | ★Stringified as a value★ |

`FALSE` also stayed on the screen. ★It was neither the product side nor the judging side, but Excel's behavior★
―― at D-7 we guessed "the way the judgment is written is bad", and **that was wrong too**.
**We misdiagnosed twice in a row until we wrote the probe.**

### ★D-7d/e: counting "the number of times" does not reveal "coverage over time"★

At D-7c the interrupt held, but the state was ★tapping Esc "pon" did not work, whereas holding it did★.
To isolate the cause we measured two things:

| | Measured |
|---|---|
| `Test_D7_Key` (DoEvents spinning idle) | 56923 times / 1.4 s = **40,659 per second** |
| During a batch (`MapsCheckCount`) | 29866 times / 35.5 s = **841 per second** |

At 841 per second, even a 100 ms keypress should be caught 84 times, but it is not.
★Because polling is not uniform in time.★ Assuming it ran at the same speed,

```
29866 ÷ 40659 ≒ 0.73 seconds  →  we were watching for about 2% of the 35.5 seconds
```

★**The remaining 98% is waiting on the `Wv2Pane` side (`EvalSync` / `WaitFor` / `WaitSettled`), and
there is not a single detection point there**★. So a "pon" misses 98% of the time, while a long press hits.
`GetAsyncKeyState`'s bit0 (pressed since last time) is **not working** ―― if it were, a keypress
in the blank time would be picked up at the next poll.

★**D-7d's conclusion that "it's 565 per second so the calls are not insufficient" was wrong**★
What should have been counted was **time coverage, not count**. We thought we had kept design principle 112 (count directly), but
★we counted the wrong thing★.

### ★D-7e: make a verification observed by a person visible to the person★

In D-7d's real-machine test **the interrupt had succeeded**, but to tarboh it looked as though "nothing stopped at all".
★Because the verification Sub automatically ran a 2nd round right after the interrupt★.
Moreover the Esc he mashed also took effect on the 2nd round, and 3 checks FAILed
(the product side was correct, ★only the arrangement of the verification was bad★).

What we fixed:

- When interrupted, print `★★★ Stopped (interrupted after processing n items) ★★★` **to the Immediate window**
- With ★`TestWaitEscReleased`★, **wait until the finger leaves Esc (2 seconds of quiet)** before resuming
- Also set `xlDisabled` during that wait (otherwise a keypress during the wait would break)
- `Test_D7_Key` was changed from "10 seconds from now" to a form that ★waits until pressed★
  (the previous version **ended before reading**, so it misjudged forgetting to press as "the API can't see it")

### Real-machine test results

**`Test_D7_Sheet` all 13 OK** (interrupt -> display -> release finger -> automatic resume -> all filled).
`Test_D7_Cancel` 15, `Test_D7_Resume` 8, and `Test_D7_Key` 1 also passed.

```
★★★ Stopped (interrupted after processing 1 item) ★★★
★13 checks: OK 13 / FAIL 0★
```

### Design principles established (113 / 114)

- **Design principle 113 (D-7d) Evaluate intermittent observation not by "count" but by "coverage over time".**
  841 times per second looks sufficient, but if it is concentrated in 2% of the time, 98% is missed.
  ★A count gives false reassurance that "it's enough"★. To get the coverage rate,
  take the ratio to the "frequency when run continuously" (this stage's 40,659 per second is the baseline).

- **Design principle 114 (D-7e) Make a verification observed by a person visible to the person.**
  If the next thing is run right after an interrupt, the moment of "it stopped" cannot be seen. **It is reported as a failure
  although it holds**, so isolating the cause wastes a whole round trip.
  ★For a verification that makes a person press a key, wait before the press, and stop and show the result of the press.★

### Changed files (the whole of D-7)

`Wv2Maps.bas` / `Wv2Tests.bas`. **The product core (`Wv2Pane` / `Wv2Element`) is unchanged.**
`Wv2Thunks.bas` changed only in the spelling of a label (below).

### ★Accident: a case difference in a label rewrote the whole project★

The `Cleanup:` added in D-7b collided with the existing label `CleanUp:` in `Wv2Thunks.Test_Environment_Real`, and
★by spec fact 38 the spelling across the whole project was unified★
(behavior is identical). **It had passed through `check_idents.py` beforehand with "no collisions"** ――
★that script does not pick up labels as identifiers★. We put it on axis E's open-items list.
For now we cross-check manually as follows:

```python
# Collect ^([A-Za-z]\w*):\s*$ from src/*.bas, *.cls and look for collisions with lower()
```

There are currently 7 labels (`Canceled` / `Cleanup` / `GuardExit` / `Resolved` / `done` /
`eh` / `fail`), with no collisions.

### Remaining holes (open items for the next stage)

1. ★**Esc does not work with a "pon" (a long press is needed)**★ The cause is the 2% coverage above.
   To close it, we wire a detection point into `Wv2Pane`'s wait loops = ★touching the product core★.
   The issue is "where to put the interrupt flag" (create a cross-cutting module, or a public property).
2. ★**Cannot interrupt during the 10 to 15 seconds of `MapsOpen`**★ Measured: `★interrupt checks 2 times / 10.5 seconds★`.
   Nearly all of the waiting is `AddTabWithUrl` and `WaitFor` (both on the `Wv2Pane` side).
   **It can be solved at the same time by the same wiring as 1.**

---
## Outside recommendation — The method of judging "waiting for completion" on SPAs (received 2026-08-20)

★Not a stage. Recorded as input for when designing D-3 / D-4.★

### Source

A recommendation sent to this project by **a participant in the SeleniumVBA project**.
Received via tarboh. The original text (translated from Japanese):

> I was a little curious about SPA support. A mechanism for catching things like HistoryChanged and DOMContentLoaded
> is already implemented, isn't it?
>
> On the other hand, in cases where multiple asynchronous processes chain together, such as a button click followed by
> fetch/XHR and then a DOM update by React/Vue, I wonder what method the automation wrapper you are about to build
> will adopt for the completion wait that judges "it's OK to proceed to the next operation".
>
> There seem to be various signals available — DOM condition waits, MutationObserver, Network Idle, WebView2's various events —
> but for SPAs there is no decisive one where watching any single one is certain, so it's a tricky matter.

### Why it hits home

**It strikes at the core of concept ①, "ease of use that excels at automatically controlling web pages".**

What was assembled up to D-2 is a means of "**reading the current DOM**" (`EvalSync` / element registry /
generation judgment); there is not yet any means of "**waiting for the change to finish**".
The moment D-3 (`.Click` / `.Value =`) is introduced, a chain of **operation -> wait -> next operation** is
certain to be needed. If the way of waiting is wrong there, we get automation that
"passes locally but fails in production".

What we currently have and how far each reaches:

| What we have | Implemented | Does it reach SPA chains? |
|---|---|---|
| `NavigationCompleted` | 9.4b | **Does not reach** (does not fire on SPA screen transitions) |
| `DOMContentLoaded` | 9.9 | **Does not reach** (only the initial document) |
| `HistoryChanged` | 9.9 / 9.10 | **Partially** (picks up transitions accompanied by pushState, but not re-renders without it) |
| `EvalSync` | D-1 | Can synchronously evaluate any JS, so it **can serve as a foundation for waiting** |
| Element registry + generation | D-2 | Tells "whether the DOM was rebuilt" but not "whether it has settled" |

### The signals the recommendation lists, and the current assessment

**Not decided yet.** Placed here as material to put to a one-go agreement as issues of D-3 / D-4.

| Signal | Strengths | Weaknesses |
|---|---|---|
| **DOM condition wait** (until a given selector appears/disappears/the text changes) | **The intent is explicit**. Few false detections. Light to implement (an `EvalSync` loop suffices) | The caller must write "what to wait for" every time. A burden on ease of use |
| **MutationObserver** | Captures the DOM change itself. General | Judging "the change has stopped" needs the concept of a **quiet period**. Some pages never quiet down because of animations or ads |
| **Network Idle** | Framework-independent. Strong against fetch/XHR chains | Cannot be obtained directly from WebView2. `fetch` / `XMLHttpRequest` must be wrapped and counted on the JS side (**a good match for design principle 92's self-generating prologue**). Some pages never become idle because of long polling or WebSocket |
| **WebView2's various events** | Already wired | As in the table above, they do not reach SPA chains |

**Current personal view (not agreed)**: The remark that there is no decisive one is correct, and
**rather than choosing a single signal, it appears sound to lean toward "composition + an explicit cutoff"**.

- The default is "**DOM condition wait**" — make first-class the form in which the user can write what they want to wait for.
  Concept ① is not "universal automatic judgment" but "**ease of writing**", so
  a form like `WaitFor("#result", 5000)` could be the main event.
- As a supplement, provide "**quiescence waiting**" (a composition of MutationObserver + a Network counter) as insurance for when the target to wait for
  cannot be specified.
- **Always have a timeout, and return failure when exceeded** (put it on design principle 93's
  `LastOk` / `LastError` pattern). ★Because of spec facts 20 / 14,
  always place the wait loop under the same in-callback guard as `EvalSync`.★

### Cautions when starting

- ★**Waiting is needed at the same time as D-3**★ If we introduce `.Click` and then say "waiting comes in the next stage",
  automation in the meantime is effectively unusable. Always ask in D-3's issues **whether to include a minimal form of waiting**.
- Wrapping `fetch` / `XHR` **leaves traces on the page side**, so consistency with design principle 92 (do not inject /
  the expression that uses it prepares it itself) must be decided first.
  The key is whether the counters can be recreated every time by the self-generating prologue.
- SeleniumVBA very likely has an existing solution to the same problem.
  **It is worth investigating SeleniumVBA's wait API before starting** (ask the recommender if possible).

### The recommender, and the reference product introduced (added 2026-08-22)

**The recommender is sele_chan** (a member of the Excel-Fun community).

★**The same person is the author of the product below**★ (it came out on 2026-08-22; it had at first been
mistakenly recorded as "the introduced product"). The handle is used differently depending on the venue:

| Venue | Handle |
|---|---|
| GitHub / Medium | `hanamichi77777` |
| note / Qiita / Discord | `sele_chan` |

Grounds: on Discord they write from the position of **deciding the design**, with statements such as "I updated to v4.3 and added...", "I don't want to design
it by doubling the number of Shadow-only methods", and "implementation complete".
The Qiita article too announces their own release, saying "the public version is at the following GitHub link".

sele_chan also publishes articles. Places to consult if work stalls around SeleniumVBA:

- `https://note.com/sele_chan`
- `https://qiita.com/sele_chan`

**WebDriver BiDi for SeleniumVBA** (author sele_chan = hanamichi77777 on GitHub, MIT, v4.3)
`https://github.com/hanamichi77777/WebDriver-BiDi-for-SeleniumVBA`

★**What is written here is secondhand information from just reading the README**★ (`src/` is 3 files: `BiDiCommandWrapper.cls`
467 KB / `BiDiSocketCommunicator.cls` 203 KB / `BiDi_Sample.bas` 26 KB. The contents are unread).
When actually adopting anything, consult the original.

#### What it does

An **extension** of SeleniumVBA (not a replacement). From VBA it speaks **W3C WebDriver BiDi** directly
over a **native WinSock WebSocket**, subscribing to browser-side events.
The aim is **exactly the same** as this project's:

> In SPAs where "request completion = page ready" does not hold, how do we obtain the grounds for
> judging that it is OK to proceed to the next operation?

It is striking that the verification benchmarks are real sites of the hardest kind: the **ServiceNow login form** (Shadow DOM + asynchronous UI) and
**Google Flights flight search** (dynamic suggestions / fare calendar / asynchronous results).

#### Design of waiting (★this is the main subject★)

| Mechanism | Contents |
|---|---|
| **idle consensus** | Infers "stable enough" from a **composition** of network activity + Fetch/XHR counters + DOM mutations + a **quiet window (slackMs)** |
| **`ArmNetworkSignal(pattern)`** | For the next one time only, wait for a specific network request/response |
| **`ArmContentSignal(xpath)`** | Wait for an **existing** DOM subtree to be rewritten. ★If the target is absent at arm time, an immediate error★ |
| **`ArmVisibilitySignal(xpath)`** | Wait for an element to become visible |
| **`AddIdleIgnoreNetworkPattern`** | Lets the request through but **excludes it from the quiescence judgment** (for always-chattering telemetry) |
| **`ExecuteEnableResourceBlocking`** | **Does not let the request be sent** (stronger. Can change the page's behavior) |
| **`waitForCompletion`** | Choose per operation whether to wait. The stable window (e.g. 1000 ms) is also specified per operation |
| **Discovery Log** | A diagnostic tool that records network / DOM mutations / suppressed noise / slackMs / stability judgments **on the same timeline** |

★**The most important pattern is "arm, then act"**★ A signal is **one-shot** and
is consumed by the next wait. To do the same operation twice, arm twice.

#### Limits the author states explicitly (honest, and they apply to us as they are)

- ★**SPA completion is an "inference", not a guarantee**★ Internal logical completion cannot be proven, and
  future delayed processing cannot be ruled out either
- **Use explicit signals for important operations.** Rather than relying on quiescence alone, it is safer to arm `ArmContentSignal` /
  `ArmNetworkSignal` just before the operation
- **If the ignore patterns are too broad, even meaningful requests are excluded and an too-early STABLE is produced**

#### Differences from ours ― what can be borrowed / what cannot

| | Theirs | Ours (WebView2 + `EvalSync`) |
|---|---|---|
| Path | WebDriver BiDi (WebSocket) | Only `ExecuteScript` round trips |
| Network observation | Can be taken from the protocol **without touching the page** | ★The only way is to wrap `fetch` / `XHR` inside the page★ |
| DOM changes | BiDi + in-page probe | Place a MutationObserver in the page |
| Dependencies | Needs a WebDriver / the browser's BiDi endpoint | Zero additional dependencies |

**The communication path itself cannot be borrowed.** It needs a BiDi endpoint, which clashes with this project's
premise of zero external libraries. **What should be borrowed is the shape of the API and how limits are handled.**

#### Concrete implications for D-4

1. ★**"Arm, then act" relieves the tension with design principle 92**★
   Counting network activity requires "planting the wrap before requests start", which by nature does not mesh well with the self-generating prologue
   (lazy generation). **But with arm-then-act, if the wrap is created at arm time, it is
   always before the operation.** If we cut the spec as "count only from arm time onward",
   there is a good chance we can avoid bringing out `AddScriptToExecuteOnDocumentCreated`.
2. **Do not make quiescence waiting a standalone API; make it composition + cutoff.**
   Put it on our `LastEvalOk` / `LastError` pattern, and ★write the fact that "it is an inference" into the meaning of the return value★
   (e.g. distinguish `STABLE` / `TIMEOUT` / `UNKNOWN`).
3. **Design noise exclusion (removal from the idle judgment) and blocking (not letting it be sent) as separate things.**
   We have no means of blocking for now, so first only exclusion.
4. **Make the waiting policy selectable per operation.** D-3's `.Click` is designed not to wait, so
   D-4 will add the form "wait until quiescence after this operation". Whether to change the default is an issue.
5. ★**We can build the equivalent of the Discovery Log too, and the effect is large**★
   We already have K-1's `Wv2Log`. If we output what the in-page probe collected — "requests / DOM changes / judgments" — as a
   timeline, **we can decide "what should be waited for" by observation rather than guesswork**.
   That they even envision an operation of feeding the log to an AI for analysis
   fits as is with this project's development cycle (Claude directly reads the log).

#### Holes of ours that came into view along the way (open items for D-4 onward)

- ★**Cannot penetrate Shadow DOM**★ `document.querySelector` does not look inside a shadow root.
  It is essential on real sites like ServiceNow (they have `ExecuteShadowClick`).
- ★**Cannot cross iframes**★ They specify a browsing context ID and hit directly inside the frame.
  We look only directly under `document`.
- They also handle file-selection dialogs / downloads / new-tab capture, but
  in light of concepts ① and ② these are **low priority** for us (new tabs are already handled separately in 9.27).

### What happened to the reference product afterwards (shared on 2026-08-22)

The author sele_chan shared the status of `WebDriver BiDi for SeleniumVBA` v4.3.
★There are 3 points that directly help us.★

#### 1. ★Recording DOM changes during manual operation★ (an extension of the Discovery Log)

The operation the author describes:

```
Manually observe an unknown SPA
        ↓
View Network + DOM on the same time axis
        ↓
Investigate what the data source is
        ↓
Decide what to use as the completion condition
        ↓
Finally design the automation method
```

"If you feed a log recorded by manual operation to an AI for the first site you tackle,
you can set a guideline on what wait strategy to proceed with before implementation."

★**We can follow this as is with our toolset**★ D-4e's probe
works even without automatic operation, so if you set `SpaProbeLogging = True`, then **operate by hand**, and
extract with `SpaProbeDrainLog`, you get the same record.
**The order itself of "observe before implementing" is the same philosophy as design principle 107.**

#### 2. VirusTotal 0 detections (including Bkav)

0 detections out of 64 vendors in v4.3. ★Even Bkav, the false-positive king, judged it clean★.
This is a useful reference for our open item since 9.30 (Bkav false positives).
It shows that an .xlsm with VBA + WinSock can also get a zero verdict.

#### 3. ★For Shadow DOM, they made existing methods support it instead of adding dedicated methods★

The path the author took and the conclusion:

> Methods inside ShadowDOM only support clicking so far. However, I don't want to design it by
> doubling the number of Shadow-only methods like `ExecuteShadowInput` / `ExecuteShadowSelect` / `ExecuteShadowGetText` …
> alongside those for the normal DOM, so I'd like to experiment toward unifying them.

> Implementation complete. All methods that use XPath on the normal DOM
> can now be used on ShadowDOM too.

★**This can be borrowed as a starting point when we tackle Shadow DOM**★
Rather than creating Shadow-only twins (`ShadowClick`, etc.) on the `Wv2Element` side,
the sound approach is to **make the grabbing side (`QuerySelector`) support piercing shadow roots**.

Note that "XPath can't be used in Shadow DOM" is the same type of problem for us:
`document.querySelector` does not pierce shadow roots. To pierce, we would write
a recursive search that walks `shadowRoot` on the JS side (in a form that rides on D-2's prologue).


### Comparison with Playwright ― ★third-party SPAs have no completion signal★ (received 2026-08-23)

sele_chan's reaction to the D-4 / D-7 development report. Via tarboh. The original text (translated from Japanese):

> Regarding waiting for SPA sites, in Playwright the position is "SPA creator = test executor",
> so I hear the stance is to catch the completion signal created on the SPA application side.
>
> However, what we are seeking is the completion condition of SPA sites created by third parties. With third-party
> SPAs there are cases where no completion signal is made explicit, so the point seems to be how accurately
> we can observe and discover it even in that situation. As a supplement, the BiDi extension takes the approach of
> having an AI read an observation log (with only the necessary parts extracted) and infer the completion condition.

#### Why this hits home ― ★our standing was put into words★

The reason for "there is no decisive one" is explained as a **difference in premise**.

| | Playwright | This project / BiDi extension |
|---|---|---|
| Target SPA | ★One they built themselves★ | ★One a third party built★ |
| Judging completion | Catch the completion signal the app emits | ★There is no signal so observe and discover it★ |
| Therefore | Waiting is "waiting for a promise" | Waiting is "**inferring**" |

★**Our decision (D-4) to write "it is an inference" into the return value was the correct response to this
difference in premise**★ If we designed the API thinking we were on the same ground as Playwright,
we would be making the false promise to callers that "if you wait, it's surely finished".

#### Correspondence with our toolset

| Their method | Our real thing | Status |
|---|---|---|
| Have an AI read an observation log and infer the completion condition | `SpaProbeLogging` / `SpaProbeDrainLog` (D-4e) | ★Already have★ |
| ★A log with only the necessary parts extracted★ | Reduce with `AddIgnoreNetwork` / `IgnoreSelectors` | Partial |
| Can record even with manual operation | The probe works even without automatic operation | ★Already have★ |

★**The difference is the single point of "extraction"**★ They **extract only what is necessary** on the premise of having an AI read it.
Our `SpaProbeDrainLog` spits out 60 at a time raw.
It fits well with a development cycle in which Claude reads the log directly, but
**what decides "what to drop and what to keep" is, for now, only the exclusion settings**.

-> Open item: ★extraction of the observation log★ (decide on real sites what to keep so that the completion condition can be hit).
  It can wait until real demand appears, but **it may pay off before Shadow DOM / iframe**.
