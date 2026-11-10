# VBA-WebView2 Development Notes Part 5 — Axis K (Foundations and Paying Down Debt)

The axis that deals with the foundation underneath every feature, and with the debts we have put off.
The thunk layer (`Wv2Thunks.bas`), the Sentinel, and the logging infrastructure that supports development all live here.

Unlike axis D (moving features forward) and axis E (the development environment), **axis K is by nature
an "investment in not moving forward."** It adds no visible features, but without it the verification that comes next cannot stand.

**The current place to append.**

---

## Stage K-1 — Writing debug logs to a file (`Wv2Log`)

### Where it fits

During the real-machine test of D-2, **the Immediate window did not retain enough lines, so the first half of the log scrolled away and
could not be checked.** The consecutive numbering of the handles let us confirm that things were healthy, but
the essential danger was ★not noticing that it had scrolled away★.

Axis D does dozens of round trips per test, so this will certainly recur. **It is the foundation for all later verification.**
tarboh himself proposed, "Wouldn't it be better to build something that writes to a text file?", and at the completion of D-2
we agreed "this is next". It was an open item.

We started once `import.ps1 -AllowNew` worked in E-3, which made it possible to bring `Wv2Log.bas` in
**without hand-building it in the VBE**.

### Issues and decisions (9 items agreed in one go)

| # | Issue | Decision |
|---|---|---|
| 1 | Output location | `%APPDATA%\Wv2Browser\logs\`. Same root as `settings.txt`. Avoided putting it next to the workbook because that **could collide with OneDrive sync** |
| 2 | How to split files | **One file per launch** (`wv2_yyyymmdd_hhnnss.log`). Keep 20 and delete older ones automatically. Axis D is "1 test = dozens of round trips", so having **one test run closed within one file** helps verification |
| 3 | Log levels | Four levels (`LOG_ERROR` / `LOG_WARN` / `LOG_INFO` / `LOG_DEBUG`). The threshold can be changed at run time. **The default during the transition is `LOG_DEBUG`** (so nothing equivalent to the existing `Debug.Print` is dropped) |
| 4 | Echoing to `Debug.Print` | ★`LogPrint` also fires `Debug.Print` internally★ The caller writes just one line. This allows a **mechanical 1-to-1 replacement** for migration, and what shows in the Immediate window does not change |
| 5 | Scope of migration | Only the axis-D parts of `Wv2Pane`. The rest next time or later (**do not rewrite every existing `Debug.Print` at once**) |
| 6 | Write method | **Keep the file open and write every line.** Force a flush right after `ERROR` / `WARN`. The idea of buffering in memory and writing at exit was rejected because, per spec fact 20, crashes are a realistic possibility |
| 7 | Re-entrancy handling | **Counter + pending queue** (same pattern as design principle 78). Discarding lines that arrive during re-entry would betray the very purpose of K-1, "not noticing that it scrolled away", so they are queued and written when we leave |
| 8 | Character encoding | Log is UTF-8 / no BOM / CRLF. **Do not reuse** `Wv2Tests.WriteUtf8NoBom` (`Wv2Tests` is positioned as "verification only, removable when porting", so product code must not depend on it) |
| 9 | Identifier naming | ★Avoid one-character identifiers★ They are a target of spec fact 43. Adopted `LogE` / `LogW` / `LogI` / `LogD` (`E` / `W` / `I` / `D` rejected) |

### The `Wv2Log` API

```vba
Wv2Log.LogE "..."      ' ERROR (forced flush right after writing)
Wv2Log.LogW "..."      ' WARN  (same)
Wv2Log.LogI "..."      ' INFO
Wv2Log.LogD "..."      ' DEBUG

?Wv2Log.LogPath        ' Full path of the current log file
Wv2Log.LogLevel = LOG_WARN   ' Threshold
Wv2Log.LogEcho = False       ' Stop echoing to the Immediate window
Wv2Log.LogStart        ' Switch to a new log (restart a verification)
Wv2Log.LogStop / LogFlush
```

The output looks like this.

```
=== Wv2Log start 2026-08-20 08:36:20 ===
=== path = C:\Users\gugug\AppData\Roaming\Wv2Browser\logs\wv2_20260820_083620.log ===
=== The 6 digits at the start of each line are a sequence number. A gap means the log dropped a line ===
000001 2026-08-20 08:36:20.033 [ERR ] K-1: ERROR line
000002 2026-08-20 08:36:20.033 [WARN] K-1: WARN line
```

### ★We numbered the lines★ (an addition that was not among the issues)

**The motivation for K-1 is "not noticing that it scrolled away", so unless we build in a way to notice,
we have solved only half the problem.** If you look at the 6 digits at the start of the line and see a gap, a line was dropped.

The verification Sub also looks at this, and "no gaps in the sequence numbers" is included in the pass criteria.

### Checks done before implementation

- **Added an identifier collision check** (`tools/check_idents.py`). It compares, lower-cased, against the existing 748.
  We started writing only after confirming that all 50 identifiers used in `Wv2Log` had no collisions.
- **Reproduced the UTF-8 encoder in Python first**, and confirmed by brute force over **all of U+0020 to U+FFFF** plus
  a thinned-out sample of the supplementary planes that it matches Python's `encode('utf-8')`.
- **Added a formatting and syntax check** (`tools/check_vba.py`). It makes standard check no. 1 permanent.

### ★A trap we stepped on: design principle 29 (the surrogate check nearly died)★

`&HD800` / `&HDBFF` / `&HDC00` / `&HDFFF`, written for the surrogate-pair check,
all exceed `&H7FFF`. **Without the `&` suffix they are evaluated as negative Integers,**

```vba
If logCode >= &HD800 And logCode <= &HDBFF Then
```

**so this condition never holds.** A defect where 4-byte characters get silently corrupted and which cannot be noticed by eye.
The newly added `check_vba.py` detected it, and we crushed it before handing over the code.

★Further confirmation came after the import★ The VBE **dropped the redundant `&`** from `&H10000&` / `&H40000&`
but **kept** it on `&HD800&` / `&HDC00&`.
**The VBE itself judges that "the danger zone is only 0x8000 to 0xFFFF".** This matches the checker's threshold.

### Real-machine test results — it took 3 tries

★1st try 4/13, 2nd try 13/14, 3rd try 14/14 and it passed.
The cause of both failures was not `Wv2Log` but **the verification Sub's side.**★

#### 1st try (4/13) — ADODB cannot read the log

`K1ReadUtf8: failed (3002) The file could not be opened.` appeared, and
the 9 items that depend on reading back all failed together.

**Two discoveries were made here.**

1. **The log file itself had been perfect from the start.** Because it was buried under the failure display,
   we read the file directly to check. UTF-8 / no BOM / CRLF, sequence numbers 1 to 10 with no gaps,
   **`[😀]` intact**, and the pending queue behaved as designed (`000008` `000009` after `000010`).
   ★The `[??]` in the Immediate window is merely a limit of CP932 display; the data is not corrupted★
2. **Checks (8) and (9) were "passing vacuously".** Even though the read-back had failed,
   `InStr("", x) = 0` let the "confirm it is absent" checks pass.
   **A check that cannot fire is worse than no check** (design principle 95), yet we were doing
   exactly that in our own verification Sub.

#### 2nd try (13/14) — ADODB cannot open it even with `Shared`

We changed `Lock Write` to `Shared`, but only (3) kept failing.
So we used PowerShell to separate **whether it was a sharing-mode problem or specific to ADODB**.

| Writer's sharing mode | .NET ReadAllBytes | Reader that specifies sharing | **ADODB.Stream** |
|---|---|---|---|
| Equivalent to `Shared` | NG | **OK** | **NG** |
| Equivalent to `Lock Write` | NG | **OK** | **NG** |
| Exclusive | NG | NG | NG |

**`ADODB.Stream` cannot open a file that another handle has open, regardless of sharing mode.**
So (3) was a check that **cannot pass with ADODB in principle**, and
★I was measuring the wrong thing with the wrong tool★. It is not a defect on the `Wv2Log` side.

At the same time we learned that "**a reader that specifies sharing can open it**", so
there is no obstacle to opening the log in an editor.

#### 3rd try (14/14) — pass

Replacing (3) with VBA's own `Open ... For Binary Access Read Shared` made every item pass.

| # | Check | Result |
|---|---|---|
| 1-2 | `LogPath` can be obtained / the file actually exists | OK |
| 3 | **The log can be read while running (shared read)** | OK |
| 4 | The log could be read back | OK |
| 5-8 | A line exists for each of the 4 levels | OK |
| 9-10 | The INFO / DEBUG lines cut by the threshold are absent | OK |
| 11 | **Surrogate pairs intact** | OK |
| 12 | Japanese intact | OK |
| 13 | **The 2 lines written during re-entry are flushed afterwards** | OK |
| 14 | **No gaps in the sequence numbers (10 lines / max 10)** | OK |

### Post-implementation report (design principle 79)

1. **Added a 6-digit sequence number at the start of each line.** It was not among the issues, but it ties directly to the motivation of K-1.
2. **Made the log file open with `Shared`.** So that it can be read in an editor while running.
   We had originally used `Lock Write`, but that turned away readers too.
3. **Added `Wv2Log.Debug_SetLogDepth`.** A hook for verifying the re-entrancy guard
   (same pattern as `Wv2Pane.Debug_SetInCallback`). Product code does not call it.
4. **Changed check (14) from "sequence order" to "absence of gaps".**
   The re-entrancy test flushes the pending queue later, so the order is deliberately swapped.
   K-1's motivation is loss, not order, so this is the right one.
5. **Made two checking tools permanent** (`check_idents.py` / `check_vba.py`).
   In line with design principle 95, we turned the constraints we stepped on into checkers rather than just fixing them on the spot.

### Changed files

| File | Contents |
|---|---|
| `src/Wv2Log.bas` | **New, 465 lines** |
| `src/Wv2Pane.cls` | Migrated the `Debug.Print` in `EvalSync` (9) and `ClearElementRegistry` (2). ★Not one line of logic was changed★ |
| `src/Wv2Tests.bas` | `Test_K1_Log` / `Test_K1_Help` / `K1Bool` / `K1ReadUtf8` / `K1CanReadWhileOpen` |
| `tools/check_idents.py` | New |
| `tools/check_vba.py` | New |

**Only 11 were migrated. 1,623 `Debug.Print` calls still remain in `src/`.**
We will move them gradually.

### Design principles established (99 to 100)

## ★ Design Principle 99 — Confirm an absence only after verifying that the target could be read ★

A search over an empty string answers "not found" for anything. If you confirm "there is no ~" in a state where
the read has failed, **it passes even though it is broken.**

We actually stepped on this in the 1st try of K-1. "There is no INFO line cut by the threshold" passed
even though the log could not be read back.

The remedy is simple: **verify first that the read succeeded, and if it did not, abort the checks that follow.**
If the output states plainly that we aborted, it also does not inflate the pass count.

## ★ Design Principle 100 — When a verification fails, first suspect whether the "way of measuring" can measure the target ★

K-1 failed twice in a row, but **both times the cause was on the verification Sub's side**, and
`Wv2Log` had been working correctly from the start.

- 1st: the reading tool (ADODB) simply could not open the file
- 2nd: we were trying to measure **a property that tool cannot measure in principle**

If you mistake a tool's constraint for a defect in the target, you go off to break something that is correct.
**Checking directly, by a path separate from the target, around the failed item** is the shortest way to isolate it
(in K-1, reading the log file directly with Python let us declare at once that "the contents are perfect").

### Spec facts established (44 to 46)

## ★ Spec Fact 44 — `ADODB.Stream.LoadFromFile` cannot open a file that another handle has open ★

★It cannot open it regardless of sharing mode★ (measured). Even if the writer has it open
with `Shared` (no lock), you get `The file could not be opened (3002)`.

On the other hand, **a reader that specifies a sharing mode can open it.**
In VBA that is `Open path For Binary Access Read Shared As #n`.
When you want to read a log that is in use, you must not use ADODB.

## ★ Spec Fact 45 — `Debug.Print` displays in CP932, but the String itself is intact ★

If you `Debug.Print` a character not in CP932, it becomes `?`.
**But that is only about display; the VBA String remains intact as UTF-16.**
If you write it to a file as UTF-8, it is preserved correctly.

In K-1's measurements, the same line that showed `Surrogate pair [??]` in the Immediate window
read `Surrogate pair [😀]` in the log file.

★Do not jump to the conclusion that "the data was corrupted" from a `?` in the Immediate window.★
This is a separate matter from spec fact 35 (the source being bound to CP932).

## ★ Spec Fact 46 — The VBE drops the redundant `&` suffix on `&H` literals ★

`&H10000&` / `&H40000&` become `&H10000` / `&H40000` after import
(0x10000 and above are Long from the start, so no `&` is needed).

On the other hand, **the `&` on `&HD800&` / `&HDC00&` stays.**
0x8000 to 0xFFFF turn into Integer without the `&`, so the VBE also judges that it is meaningful.

★This is the VBE's own confirmation that the danger zone of design principle 29 is "0x8000 to 0xFFFF".★
If the `&` has vanished after import, it is not corruption.

### Notes for K-2 onward

- **Migrating the remaining 1,623 `Debug.Print` calls.** Do not do it all at once.
  The next one that pays off is `View_On*_Core` in `Wv2Pane` (this is what floods the Immediate window with
  event bursts), and `Wv2Browser`.
- **A permanent fix for the Sentinel MEM_FREE bug (spec fact 10)** is the heart of axis K.
  `Wv2Thunks.bas` is the foundation of every feature, and by spec fact 20 it is territory where the debugger cannot be used.
  Do it when we can settle in for it.
- Log rotation is fixed at 20 files. If it becomes a problem, change the constant `LOG_KEEP_FILES`.
- **Performance is unmeasured.** We do not flush every line, so it should not be heavy,
  but we will judge after seeing how it feels in axis D's real tests (dozens of round trips).

---

## Stage K-2 — Fixing the bug where the settings tab is destroyed by a search

### Where it fits

An existing bug found in the D-2 regression check and put on the open-items list. It lay in the gap between 9.19 (introduction of `IsSettings`) and
9.24 (consolidating the NavBar's Host), and is unrelated to D-2.
Since 9.xx is frozen, we handled it on axis K as debt repayment.

**Symptom**: If you enter a search term in the address bar while the settings tab is active, the settings tab itself navigates to the
search results and the settings screen disappears. Furthermore, pressing the gear does not open a new settings tab because it
judges that "a settings tab already exists".

### True cause, and why a symptomatic fix does not work

`IsSettings` was set once at creation and **never went back to `False` even after navigation.**
Because the flag did not reflect reality, the duplicate check in `OpenSettingsTab` was
grabbing the "former settings tab".

★**It is not only an address-bar problem**★
The flag goes out of sync because "navigation happened", so the same thing occurs with **link clicks, Back, Forward, reload, and
navigation originating from JS**. A symptomatic fix that blocks `NavigateActive` will not cure it.

This view decided the choice among the later options.

### Issues and decisions

We first agreed on 6 items, but **at tarboh's remark, issue 4 was withdrawn and issue 7 was added**.

| # | Issue | Decision |
|---|---|---|
| 1 | Stage number | **K-2** (debt repayment) |
| 2' | Method of judging | **A-2: ask each tab directly with `EvalSync`** (below) |
| 3 | Judgment expression | Embed `window.__wv2Settings = true` in the settings HTML and ask `(window.__wv2Settings===true)?'yes':'no'` |
| 4 | ~~Whether to restore the flag when Back returns to it~~ | **★Withdrawn★** See below |
| 5 | Logging | Output the judgment and sync results to `Wv2Log` (the first job for K-1) |
| 6 | Verification | Show the procedure in `Test_K2_Help`. The settings tab needs a real launch, so it cannot be fully automated |
| 7 | **Inconsistency when settings views overlap** | **B-1: when a setting changes, push the current value to every visible tab** (new) |

### ★Redesign prompted by a remark★

Originally, in issue 4 I wrote, "Even if Back returns to the settings screen, we do not restore the flag. There can end up being
two settings tabs, but the actual harm is small." To this, tarboh remarked:

> I'm bothered by the contradiction that arises with the other tab's settings
> when you change a setting in one of multiple settings tabs

**This remark was correct, and my assessment was wrong.**

The settings screen's JS is implemented to **fetch the current value with `S.GetEngine()` at load time and hold it in `current`**,
so if you change it later in another tab, the side opened first keeps showing the old value.
It was not a "visual duplicate" but **a state inconsistency**.

On further investigation, we found that **the problem splits in two**:

- **(a) The gear grabs the "former settings tab"** ... a problem of judgment accuracy
- **(b) Two settings views sit side by side** ... a problem of state synchronization

★**(b) cannot be closed off by "the judgment at open time" alone**★ Because `GoBack` does not
go through `OpenSettingsTab`. However accurate the judgment, navigating the settings tab away, opening settings elsewhere, then
going Back in the original tab gives two tabs. **There was a region the judgment improvement could not reach.**

So we made it a two-stage defense: close (a) with A-2 and (b) with B-1 separately.

### Fix A-2 — Switch the judgment to "ask the page itself"

```vba
Private Function IsSettingsView(ByVal pane As Wv2Pane) As Boolean
    js = "(function(){try{return {s:(window.__wv2Settings===true)?'yes':'no'};}" & _
         "catch(e){return {s:'no'};}})()"
    res = pane.EvalSync(js, 2)
    If Not pane.LastEvalOk Then Exit Function    ' Cannot judge = not the settings screen
    IsSettingsView = (InStr(1, res, """s"":""yes""") > 0)
End Function
```

- **It depends on neither the address representation nor the navigation path.** A tab restored by Back is also picked up correctly
- ★**Put the authority for the judgment on the side that actually holds the state** (design principle 91)★ A flag can lie,
  but the page itself cannot
- **When it cannot judge, it returns False (the safe side).** Loading, Closed, and in-callback are all
  treated as "not the settings screen". If we miss one, a new settings tab merely opens, which does less harm
  than wrongly grabbing another page's tab (distinguished by `LastEvalOk` of design principle 93)
- The way the return value is read follows the same style as `Elem_CheckStale` (have it return an object and read it with `InStr`)

★**This is the first place where D-1's `EvalSync` works in product code**★. The investment in axis D paid off here.

**We did not adopt the idea of judging by address.** The settings screen is shown with `NavigateToString`, and
by design principle 18 its **source URL is `about:blank`**. `AddTabWithHtml` also uses
`NavigateToString`, so `about:blank` is not unique to the settings tab.

### Fix B-1 — Synchronizing all settings views

We placed a hook in **one place, the Property Let of `Wv2Browser.SearchEngine`**.
It is a chokepoint that is always passed through, whether via the bridge or called directly from VBA, so this alone catches every path.

★**To avoid the deadlock of spec fact 27, we wait one beat with `SafeTimer`**★
`SetEngine` is called via hostObjects (JS to VBA), and **during that time the calling pane's JS is
completely blocked.** Firing `EvalSync` from here means the completion notification never comes back and it hangs.
We get out of the handler in the same pattern as `RequestSettingsTabAsync` / `RequestNewTabAsync`.

We prepared the receiving end of the sync on the settings screen side.

```js
function applyFromHost(name){
    if(!name){ return 'ng'; }
    current = name;
    markSel();              // markCustom is also called inside markSel
    showCurrentPreview();
    return 'ok';
}
```

★**Do not overwrite the custom input field (`ctpl`)**★ because the user may be in the middle of typing.
Only the selection state and the preview are brought into line.

**We kept the `IsSettings` flag, but it is no longer used for duplicate checks.**
It was demoted to a record of its origin, "created as a settings tab" (it is a public property, so we left it rather than deleting it).

### Preliminary investigation — the pre-warm spare tab

Because it scans `m_tabs`, we checked before implementation whether we would trip over the spare tab of 9.27.

★**The spare tab does not enter `m_tabs`**★ This is stated in the code's own comments, and
the log also shows `← spare not included`. It enters for the first time at the moment it is adopted by `AdoptPrewarmedNow`. **The scan
will never step on it.**

### Checks done before implementation

- **Identifier collision check** (`check_idents.py`) — all 10 new ones had no collisions
- **Actually run the JS** — this environment has no `node`, so we used Windows Script Host's
  JScript as a substitute. But trying hand-written JS means little, so
  ★**we reconstruct the JS mechanically from the VBA source and `eval` it**★.

  ```
  [Judgment expression] (function(){try{return {s:(window.__wv2Settings===true)?'yes':'no'};}catch(e){return {s:'no'};}})()
  [Apply expression] (function(){try{return {r:applyFromHost("bing")};}catch(e){return {r:'err'};}})()
  ```

  **13 / 13 passed.** We confirmed that the strict `===` rejects `false` and the string `"true"`,
  that exceptions are swallowed and become `err`, and even that no exception escapes when `window` is broken.
  We were able to guarantee that ★what we tried and what we ship are identical★.
- **The CP932 check fired.** We had written a wave dash `〜` (U+301C) in `Test_K2_Help`,
  which was exactly the trap of spec fact 35. **We caught it before the import** and changed it to `～` (U+FF5E).
  Making the checker permanent paid off right away.

### Real-machine test results (all 15 steps normal)

| Step | What was checked | Result |
|---|---|---|
| 1-4 | After searching in the settings tab, the gear **opens a new settings tab** | **Fixed** |
| 5 | Mashing the gear does not multiply tabs (regression of 9.19) | OK |
| 6-7 | Same with link navigation and reload | OK |
| 8-13 | **With 2 settings views, a change in one is reflected in the other** | **OK** |
| 14-15 | Search engine reflection and custom URLs work as before | OK |

Steps 6-7 are ★**paths that plan A (blocking only the address bar) would not have fixed**★,
which supports the claim that we crushed the true cause.

### Post-implementation report (design principle 79)

1. **Excluded the custom input field from synchronization.** We judged that the harm of breaking what the user is typing is greater.
2. **Set the timeout of `IsSettingsView` to 2 seconds.** With the default 5 seconds,
   a single unresponsive tab makes the gear feel sluggish.

### A workbook accident caused in this round

★**During the K-2 work, we caused an accident that put the workbook into a compile error**★
The cause was how `import.ps1` handled attribute lines, and it was a **tool-side problem**, so we record it in E-4.
It is independent of the bug fix itself.

### Changed files

| File | Contents |
|---|---|
| `src/Wv2Browser.cls` | Added `IsSettingsView` / `RequestSettingsSyncAsync` / `m_syncTimer_TimerCall` / `SyncSettingsViewsNow`. Replaced the scan in `OpenSettingsTab`. Hooked the `SearchEngine` Let. Added a marker and `applyFromHost` to the settings HTML |
| `src/Wv2Tests.bas` | `Test_K2_Help` |

### Design principle established (101)

## ★ Design Principle 101 — Before choosing a symptomatic fix, count the paths by which the symptom appears ★

K-2's original plan A was "block navigation from the address bar". But
**the cause of the flag going out of sync is "navigation happened"**, and the address bar is only one path to it.
Links, Back, Forward, reload, and JS-originated navigation produce the same symptom.

**Counting the paths shows that a symptomatic fix can block only part of the whole.**
Conversely, if you crush the true cause (the flag does not reflect reality), everything is fixed without enumerating a single path.
Steps 6-7 of the real-machine test (link navigation and reload) were placed to confirm this difference.

### Notes for K-3 onward

- **The `IsSettings` flag is close to dead code.** Nobody reads it.
  Removing it means cutting a public property of `Wv2Pane`, so do it in a separate round.
- **Synchronization of settings views covers only the "engine name".** If settings items increase in the future,
  either add arguments to `applyFromHost` or change to passing a state object.
- Because it fires `EvalSync` once per tab, there is a possibility that **the gear's response becomes sluggish with an extremely large number of tabs.**
  We have not measured it. If it becomes a problem, add a fast path that "looks at the active tab first".

---

## Stage K-3 — Replacing the 76 characters lost to CP932

### Where it fits

Working off an open item discovered in E-1. **Not one line of product-code logic was changed.**
What we changed was **only symbols inside comments and in the guidance text of `Debug.Print`**.

Because VBA source is held in the project's code page (CP932),
characters not in CP932 become `?` at import time and are lost forever (spec fact 35).
The current workbook already contained this damage to the extent of 76 characters.

★**Even if we restore the original character, it becomes `?` again at import**★, so
**replacement with a substitute character that exists in CP932** was necessary.

### How we identified "what the original character was"

It cannot be told from the workbook. A `?` is nothing but a `?`.

**We identified it by cross-checking against the distributed files in `Downloads` (the UTF-8 originals Claude output during the chat era).**
If a line from the original, converted by "characters that break → `?`", matches a line in
`src/`, then which character that line's `?` was is settled.

★**All 76 characters could be identified (zero unknown)**★
Since the lines rewritten in K-1 / K-2 did not include any of the affected places,
every line matched the originals.

### Issues and decisions (6 items agreed in one go)

| # | Issue | Decision |
|---|---|---|
| 1 | Stage number | **K-3** (the "K-3: `Debug.Print` migration" on the open-items list moves down to K-4) |
| 2 | Replacement table | Below |
| 3 | Duplicate `[歯車] (歯車)` | **Collapse to just `[歯車]`** (2 places) |
| 4 | Cross-checking against past logs | The wording changes, so record it in the development notes (following the precedent of 9.28 / 9.29) |
| 5 | Scope | **All at once.** It is mechanical, and the checker guarantees it |
| 6 | Verification | The 5-point set of checks + convergence of the round trip + full-count cross-check + eyeballing the help Subs |

### Replacement table

| Original | Count | → | Reason |
|---|---|---|---|
| `〜` U+301C | 33 | **`～`** U+FF5E | Looks identical. Can be used as-is for range notation (`9.21〜9.23`, `3〜4 枚`) |
| `⚙` U+2699 | 28 | **`[歯車]`** ("gear") | CP932 has no gear character |
| `⇄` U+21C4 | 6 | **`⇔`** | All are the bidirectional correspondence `name⇄template` |
| `↔` U+2194 | 5 | **`⇔`** | Same (`Browser↔Pane`, `VBA↔JS`) |
| `⟳` U+27F3 | 3 | **`[更新]`** ("refresh") | The sequence `← → ⟳` becomes `← → [更新]` |
| `—` U+2014 | 1 | **`―`** U+2015 | Looks almost identical |

**Why `⚙` became `[歯車]`**: We decided after reading the context of all 28 places. It reads naturally in every form.

```
9.18 では [歯車] を押すたびに設定タブが新規生成され     (In 9.18, every press of [gear] creates a new settings tab)
TabBar の [歯車] ボタン → OnPaneWebMessage "settings"   (TabBar's [gear] button -> OnPaneWebMessage "settings")
★タブバー右端に [歯車] (設定) ボタンを追加★             (Added a [gear] (Settings) button at the right end of the tab bar)
[歯車] は「無ければ開く/あれば移動」。                    ([gear] means "open if absent / move if present".)
```

The rival proposal `[設定]` ("Settings") was not adopted, because `⚙ を押して設定タブを開く`
("press ⚙ to open the settings tab") would become
`[設定] を押して設定タブを開く` ("press [Settings] to open the settings tab"), a repetition.

### ★No effect on the UI at all★

We confirmed in advance that the 76 replaced places are **all VBA comments or `Debug.Print` strings**, and that
**not one of them is inside an HTML / JS string.**
The tab bar's gear button is drawn with the HTML character reference `&#9881;`, so it is unrelated.

### Verification

| Check | Result |
|---|---|
| Formatting and syntax (`check_vba.py`) | No problems (16 files) |
| CP932 round trip (`check_cp932.py`) | No problems |
| Workbook internals (`check_book.py`) | No problems (no duplicate attribute lines) |
| **Full-count cross-check** | **73 lines changed. All explainable by the replacement table. 0 missed, 0 remaining** |
| Convergence of the round trip | `import` -> `export` -> preview **0 / 16** |

★**We redid the full-count cross-check in the "forward direction"**★
The first cross-check script tried to apply the replacement table to the old lines (already `?`) from `git show HEAD`,
and falsely reported all 73 lines as "unexplainable differences".
When we fixed it to **treat the originals as the authority and trace "original -> replacement -> current line" forward**,
it came to 0. ★Once again the error was on the checking side (design principle 100)★

### Changed files

`Wv2Browser` 14 / `Wv2Tests` 46 / `Wv2SettingsBridge` 4 / `Wv2TabBar` 4 /
`Wv2NavBar` 3 / `Wv2Pane` 3 / `Wv2Json` 2 = **76 characters**, plus 2 places where `[歯車] (歯車)` was collapsed.

### ★Changes to log wording (note when cross-checking against past logs)★

The guidance text of `Debug.Print` changed, so it **no longer matches past verification logs as strings.**

| Old (corrupted display) | New |
|---|---|
| `? を押して設定タブを開く` | `[歯車] を押して設定タブを開く` |
| `2?3 回連打` | `2～3 回連打` |
| `← → ? を 1 回ずつ押して` | `← → [更新] を 1 回ずつ押して` |
| `名前?テンプレート解決` | `名前⇔テンプレート解決` |

Affected help Subs: `Test_9_18_BridgeLogic_Help` / `Test_9_19_SettingsTabDedup_Help` /
`Test_9_20_Persistence_Help` / `Test_9_21_HostActivate_Help` /
`Test_9_22_HostCloseReorder_Help` / `Test_9_23_Consolidation_Help` /
`Test_9_29_Help` / `Test_9_32_Help` / `Test_9_32b_Help`.

### Notes for K-4 onward

- **The originals of `Wv2Log` and `Module1` are not in `Downloads`**, but
  the former was newly created in K-1 (no damage) and the latter is ASCII only, so there is no problem.
- ★**From now on `check_cp932.py` stops this at a pre-gate, so the same damage will not grow**★
  In fact, in K-2 it caught the `〜` written in `Test_K2_Help` before the import.
- **K-4: migrating the remaining `Debug.Print` calls** is the next axis-K job.

---
## Stage K-4 ― Slow release when the form is closed

### Origin

A report from tarboh: "The release processing from `UserForm1.Terminate` onward
has a gap of about 0.25 seconds per item, and it is tedious."

★This stage was devoted to "measuring"★ ―― It came right after D-7, in which **we had missed 3 guesses in a row**,
so we went for measurement from the start. As a result we find that ★guessing would have been a complete waste of effort★
(below).

### K-4a: Moved the release path's logging to `Wv2Log`

This path stayed on `Debug.Print`, leaving nothing in the log, and we could not tell
where the time was going. ★`Wv2Log` has millisecond timestamps, so just replacing it lets us
read the durations★. The targets were 17 places in `Wv2Pane.Class_Terminate` / `CleanupAllPersistentHandlers` /
`Wv2Browser.Class_Terminate` + `Shutdown` / `Wv2TabBar.Shutdown` /
`Wv2NavBar.Shutdown`.

★If we measure twice, with `LogEcho = False` and `True`, we can tell apart whether "COM release is slow" or
"output to the Immediate window is slow"★ ―― that was the plan for the measurement.

### ★Measurement (1): `Debug.Print` was not the culprit★

| | Total |
|---|---|
| 1st (`LogEcho = True`) | **0.336 s** (55 lines) |
| 2nd (`LogEcho = False`) | **0.336 s** (55 lines) |

★Identical down to the millisecond★. **Output to the Immediate window was irrelevant.**
★Had we gone ahead with K-4 (migrating the remaining 1,600 `Debug.Print` calls) on a guess,
it would have been a complete waste of effort.★

And ★in the plain state (launch -> close immediately) it is 0.336 s★, so it is not slow to begin with.
**The reproduction conditions were different.**

### ★Measurement (2): opening a heavy page makes it 20 times slower★

| Condition | Total release |
|---|---|
| Plain state | 0.336 s |
| ★After opening Google Maps with `Wv2Maps.MapsOpen`★ | ★**6.718 s**★ |

The breakdown: ★one COM call goes from 8-16 ms to 330-480 ms (40 times)★.
`Ctrl_Close` is even heavier, 1.4 s each.

### ★K-4b: The log timestamps caused real harm, so we fixed them★

While measuring, we found that exactly one log line ran backwards:

```
000030  22:45:23.992
000031  22:45:23.000   ← ★the seconds did not carry over★ (should be 24.000)
000032  22:45:24.008
```

The cause was ★mixing separate time sources, `Now` (to the second) and `Timer` (milliseconds)★.
When `Timer` ticks over to the next second first, you get "seconds as before, milliseconds 000".
The K-1 comment ★stated this as known★: "it may be off by one digit at a second boundary",
but real harm occurred when we nearly misread it as "there is a step that took 1008 ms".

The fix was ★just one call to `GetLocalTime`★. It gets everything from year/month/day down to milliseconds at once, so
in principle it cannot drift. We confirmed 0 backward steps in 3 logs after applying it.

★**If the foundation of measurement is broken, the measurements themselves can no longer be trusted.**★

### ★K-4c: Countermeasure ― send all tabs to `about:blank` before release★

We tested the hypothesis that "COM responses do not come back while a page is active"
with **the same thing a user can do by hand (navigate to another page)**.

| Condition | Time to navigate away | Release | Total |
|---|---|---|---|
| Do nothing | ― | 5.05 s | 5.05 s |
| Wait 3 s | 1.6 s | 0.367 s | about 5.0 s |
| Wait 1 s | 1.47 s | 0.368 s | 1.84 s |
| ★0 s (just navigate away)★ | 1.60 s | **0.376 s** | ★**1.98 s**★ |

★**There was no need to wait at all**★ ―― it takes effect **the moment** we navigate to `about:blank`.
So we made `Wv2Browser.Shutdown` navigate all tabs + the spare tab at its start
(`QuietAllPanes`). `TabBar` / `NavBar` are lightweight self-made HTML, so they are excluded.

★It can be turned off with `QuietOnShutdown = False`★ ―― when isolating a cause, quieting would ruin
the measurement, so we left an escape hatch (design principle 103: a probe must not destroy what it wants to measure).

★Limitation★ `Wv2Browser.Shutdown` is called from `UserForm_Terminate` = **the form has
already closed**. It goes from 5 seconds to 2 seconds, but does not become "instant".
To make it "instant" we would need to navigate away first in `UserForm_QueryClose`, and that
means touching `forms/` (manual operation). **Open item**.

### ★We failed to identify the cause (9 misses in a row)★

We tried to pin down "what makes closing slow" and **missed 9 times in a row**.
Candidates we eliminated: number of `EvalSync` calls (fast even at 300) / element registry / number of `DoEvents`
calls (fast even at 690,000) / elapsed time (fast even after leaving it 30 seconds) / `AddTabWithUrl` /
`WaitFor` / `EnableCancelKey` / `AddIgnoreNetwork` / adoption of the spare tab.

★Correlating all the logs, the discriminating criterion was, without exception, "whether `Wv2Maps.MapsOpen` was called"★
(all 4 slow cases had MapsOpen, all 14 fast cases did not).
But even when we reproduced everything inside `MapsOpen` with `Test_K4_Step`, it **never once became slow**.

### ★But a big gain: what is slow was not "the Maps tab"★

Only by reading the real `MapsOpen` release, Immediate window and all, could we first see the breakdown:

| Target | Time |
|---|---|
| `Ctrl_Close` of `Wv2NavBar` | 1.73 s |
| `Ctrl_Close` of `Wv2TabBar` | 1.34 s |
| ★Spare tab (waiting hidden)★ | ★2.24 s★ |
| ★**The Maps tab**★ | ★**72 ms**★ |
| 3 normal tabs | 65-72 ms each |

★**What is slow is only the "non-active Panes"**★. We had long explained it as
"all COM round trips are slow", but that was **wrong**. WebView2 shares browser processes, so
when Maps is running and busy, **COM calls to lower-priority Panes get made to wait**
―― and thinking of it this way is also consistent with why `about:blank` works. ★This is the starting point for the next dig.★

### Changed files (K-4)

`Wv2Pane.cls` / `Wv2Browser.cls` / `Wv2TabBar.cls` / `Wv2NavBar.cls` (logging only) /
`Wv2Log.bas` (time) / `Wv2Maps.bas` (5 measurement log lines) / `Wv2Tests.bas` / `Module1.bas`.

---
## Stage K-5 ― ★Starting WebView2 crashes the VBE's IME input★

### Origin

A report from tarboh (2026-08-24):

> After running "実行用", typing even one character in the Immediate window with Japanese input on
> crashes Excel. With Japanese input off I can type, and
> pasting Japanese and running it is fine.

★It was reproducible, and a crash (work is lost)★, so we investigated it as a priority.

### Isolation (`CrashTest1` to `12` permanently kept in `Module1`)

| Condition | Symptom |
|---|---|
| `Wv2Log` only (`LogStart` / `LogI` / `LogStop` / `LogFlush` / `LogEcho`) | ★Never crashed★ |
| With a file left open, enumerating via `Show` / `Dir`, then `Show` | ★No crash★ |
| `StartWebView2_Spa` (1 tab, no UI parts) | No crash |
| ★`StartWebView2_Full`★ (with `TabBar` + `NavBar`) | ★Crashes from while the form is displayed★ |
| `Spa` + adding tabs | Unstable. ★Crashes for sure after closing the form★ |

★**The only reliable boundary is "whether WebView2 has been started even once"**★.
And ★**it does not recover even if we close the form and release WebView2**★.

### Assessment

★The IME input context (HIMC) is managed per thread★. WebView2 has its own IME handling
derived from Chromium, and its window is created as a child of `hWnd_Frame` on **the same UI thread as Excel**.
That is, ★the VBE's Immediate window and
WebView2 are fighting over the IME of the same thread★.

The reason `Full` shows the symptom sooner is thought to be that a
★WebView2 with an `input`★, `NavBar`'s address bar, is there from the start.
The observation that "moving focus to the code window can sidestep it" is circumstantial evidence that
★moving focus re-establishes the IME context★.
However, ★the conditions fluctuate (a race)★, so it cannot be relied on as a workaround.

★The same family as spec fact 20 (breaking during a WebView2 event crashes)★ ――
a **side effect of WebView2 and the VBE cohabiting in the same process and on the same thread**, and
**there is no place in our code to fix it.**

### ★Handling: remove the need to type★

The real workaround is ★not typing Japanese directly into the Immediate window★. So
we took steps in the direction of **removing the need to type**:

- We added ★alphanumeric aliases `Boot` / `BootNoQuiet` / `BootMaps`★ to the
  Japanese Sub names in `Module1` (`実行用` / `実行用2` / `実行用3`)
- The verification Subs were alphanumeric from the start (`Test_D7_Sheet`, etc.), so they are safe as they are
- When Japanese is unavoidable, ★copy and paste★ (confirmed safe)

### Spec fact established (67)

For the main text, see spec fact 67 in the index. For the means of reproduction, `CrashTest1` to `12` in `Module1`
(★we left them in place rather than deleting them★ so that we can re-measure when the environment changes).
