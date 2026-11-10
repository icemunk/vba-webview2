> Historical record (translated). These were the working rules at the time of D-2; current rules are in CLAUDE.md.

# Handover Prompt (as of D-2 completion)

> **Why this is saved**: On 2026-08-19, when development moved from ordinary chat to Claude Code,
> the handover prompt that had been written on the chat side was saved as a record.
> ★This file is a historical record, not the current rules.★
> For the current rules, see `CLAUDE.md` at the repository root.
> In particular, "Known bugs in tool usage (`present_files`, etc.)" does not apply to Claude Code.

---

## What this project is
A project that implements a WebView2 browser control using **only** Excel VBA (no external COM libraries; only
machine-code thunks, SAFEARRAY-style memory primitives, and hand-built vtables). The development language is Japanese.
**The project is complete as a form usable as a tabbed browser. It is now working on the D axis, "automatic page control (MSHTML equivalent)", and has completed D-2.**

### ★The concept of VBA-WebView2 (stated explicitly by tarboh. The criterion for what to keep and what to drop)★
**The goal is not to reproduce an ordinary browser completely.** A high degree of polish is a good thing, but
the purpose is not to surf the web leisurely and comfortably from VBA. The aim is these two points:

1. **To be excellent and easy to use for automatically controlling web pages**
2. **To provide a modern GUI by laying out the controls on a UserForm in HTML**

When in doubt, use these two points to ask "is it worth matching the original (a real browser)?"
**The D axis is a direct investment in concept 1.**

---

## ★What to do first★

### 1. Read `dev-notes-0-index.md`
**The development notes are split into 4 volumes. The index file is the entry point, and the index alone gives almost every "where we are now".**

| File | Contents | Nature |
|---|---|---|
| **`dev-notes-0-index.md`** (509 lines) | Stage numbering scheme / concept / table of contents of all stages / **full list of design principles (1 to 93)** / **full list of spec facts (1 to 34)** / roadmap / **current state** / **open items** | **Read this first** |
| `dev-notes-1-foundation.md` (6511 lines) | Preface to stage 9.2 (establishing the thunk foundation) | Frozen |
| `dev-notes-2-browser.md` (6628 lines) | Stage 9.3a to stage 9.32b | Frozen |
| `dev-notes-3-D-axis.md` (365 lines) | Stage D-1 to stage D-2 | **Active, where to append** |

The full lists have three columns: number, heading, and source stage. **When you need the full text, pull the section of the "source stage" from
the volumes. The tables in the index are summaries, not the original text.**
Sections especially worth reading: **D-1 (EvalSync) / D-2 (element registry) / stage 9.23 (consolidating into Host methods)
/ stage 9.27 (prewarm delegation via put_NewWindow) / stage 9.32 (optimistic update)**.

### 2. Consult the project files
**As of 2026-08-19, all files have been updated to the latest.** Even so, when editing a standard file,
check for version drift with grep before using it as the base.

The header stage of each file (confirmed against the actual files):

```
D-2   : Wv2Element.cls (260 lines) / Wv2Pane.cls (3981 lines) / Wv2Json.bas (370 lines) / Wv2Tests.bas (2846 lines)
9.32b : Wv2TabBar.cls        9.29 : Wv2Browser.cls / Wv2SettingsBridge.cls
9.25b : Wv2NavBar.cls        9.18 : UserForm1.frm
9.16  : Wv2Thunks.bas / Wv2Url.bas      9.9b : ComCallbackHandler.cls      9.8a : Wv2Environment.cls
```

Points for checking version drift:
- `Wv2Pane.cls` → `GetElementById` / `QuerySelector` / `Elem_Read` / `Elem_CheckStale` /
  `JsQuote` / `JsRegistryPrologue` / `CurrentDomGen` / `ClearElementRegistry` / `m_domGen`
  actually exist. ★`Private Function JsonPickStr(` is **gone now**★ (moved to Wv2Json).
  There are 7 Private Subs with `View_On*_Core` plus Public wrappers of the same names (D-1).
- `Wv2Json.bas` → `JsonUnescapeAt` / `JsonUnescape` / `JsonPickStr` actually exist as Public.
  The existing `JsonEscape` / `JsonGetStr` / `JsonGetNum` / `JsonBool` are **unchanged**.
- `Wv2Tests.bas` → `Test_D2_Find` / `Test_D2_Stale` / `Test_D2_Help` +
  `D2Eq` / `D2Bool` / `D2El` / `D2WaitTitle` / `BuildD2ProbeHtml` / `BuildD2SecondHtml`.

### 3. ★The current workbook is `WebView2_v0_5_2.xlsm`★ (clean-built at stage 9.30)
What D-2 replaced were 4 files: `Wv2Element.cls` (new) / `Wv2Pane.cls` / `Wv2Json.bas` /
`Wv2Tests.bas`.

---

## ★Where we are now: there is no work in progress★

- **D-1 (EvalSync) is complete and verified on the real machine.**
- **D-2 (element registry + reading with `Wv2Element`) is complete, verified on the real machine, and recorded in the development notes.**

### What to do next: ★propose K-1 (debug log file output) as the top priority★

Item 1 of "Open items / next-stage homework (as of D-2 completion)" at the end of the index. **tarboh himself
proposed, "Wouldn't it be better to build a mechanism that writes to a text file?", and
it was agreed in the previous session that "this is next".**

- **Why it is the top priority**: In the D-2 verification, the Immediate window did not retain enough lines, and
  **the first half of the log scrolled away and could not be checked**. We could establish health from the handle serial numbers,
  but the essential danger is that ★you cannot even notice that it scrolled away★. The D axis makes
  dozens of round trips per test, so this will certainly recur. **It is the foundation for all later verification.**
- **Design direction (not yet agreed; start by raising the points for discussion)**:
  Create `Wv2Log.bas` and consolidate into `LogPrint`, and for the time being output alongside `Debug.Print`,
  migrating step by step (**do not rewrite all existing `Debug.Print` at once**).
  Points likely to be raised: log levels / output destination and rotation /
  reuse of `Wv2Tests.WriteUtf8NoBom` / performance (flush on every line or not) /
  re-entrancy during event bursts / how to phase out the parallel output with `Debug.Print`.
- **What to do after K-1 is done**: Re-run `Test_D2_Find` and confirm the **expected-value matches of (1) to (4)**
  that scrolled away and remain unchecked. In particular `InnerHTML` (the core of spec facts 30 / 33).

#### Counter-proposals

- **The bug where the settings tab gets crushed by a search** (open item 2) — cause identified. See "Unresolved bug" below.
- **D-3 (writing and operating)** — `.Value =` / `.Click` / `.SetAttribute`. A continuation of the D axis.
- **C axis (R-3: TreeView / ListView)** — can be finished reliably with zero new mechanisms. Directly tied to concept 2.

---

## ★Unresolved bug (found in the D-2 regression check; cause identified, not fixed)★

**Symptom**: When you type a search term into the address bar while the settings tab is active, the settings tab
itself navigates to the search results and the settings screen disappears. Moreover, pressing the gear does not open a new settings tab, because
it is judged that "a settings tab already exists".

**Cause**: `Wv2Browser.NavigateActive` (around line 1427) fires `View_Navigate` without looking at whether
`ActivePane` is the settings tab. Also, `IsSettings` is set at creation and is never returned to
`False` even after navigation, so the duplicate check in `OpenSettingsTab` (`If m_tabs.Item(j).IsSettings Then`, around
line 1941) grabs the "former settings tab".

**The fix policy is undecided.** There are at least 2 options (protect the settings tab from navigation / lower the
flag when it leaves the settings screen), and the policy needs to be agreed on. It is an existing bug that lay in the gap between stages 9.19 and 9.24, and is
unrelated to D-2.

---

## ★Mandatory rules for VBA code output★

- When modifying a standard file, update the version in the header (stack the latest block at the top; keep the past ones).
  Output the **entire file after modification**, not a diff, and hand it over with `present_files`
- **The development notes are the exception: output only the diff.** The append destinations are fixed at 2 files:
  - `dev-notes-3-D-axis.md` … append the body at the end (if you make a section for the K axis, either make it a separate file as part 5,
    or confirm with tarboh. At present only the D-axis file is an active append destination)
  - `dev-notes-0-index.md` … 1 line in the table of contents + 1 line each in the principle/fact tables + update the upper-limit number in the headings
    + the heading and file layout of the "current state" + replace the open items at the end
  - The "Open items / next-stage homework" at the end **always keeps only the latest single sheet**
- For verification, Claude prepares a verification Sub / help Sub, and tarboh only has to
  **type the Sub name in the Immediate window**
- Keep CRLF / UTF-8 / no BOM with a Python script (LF internally, convert everything to CRLF at the end)
- ★When writing JS inside a VBA string, backslash + double quote is forbidden★
  (use single quotes for JS attribute values and strings. **Do not write it even in comments**)
  - Established in D-2: **just use `Wv2Pane.JsQuote`.** It is `Chr$(34) & Wv2Json.JsonEscape(s) & Chr$(34)`,
    so the escape sequence is assembled at run time and never appears in the VBA source
  - Note: the 4 lines of JSON escape implementation in `Wv2Json.bas` are fine. A simple grep gives a false positive
- ★VBA does not distinguish case in identifiers★ `eNum` collides with the reserved word `Enum` (stepped on in D-1).
  **Always check new identifiers against the reserved word list, lowercased (design principle 90)**
- ★Whether a class module's `Public Enum` is visible unqualified from other modules is unverified★
  (`ControllerState` and `BrowserState` are used only where they are defined).
  D-2 avoided it with `Boolean` + a status string. If you want to use one, verify first
- The first argument of `WritePtrNatively` is `LONG_PTR` (writing `LongPtr` gives a run-time error)
- `AddressOf` can be written only in the actual-argument position of a function (`addr = VBA.Int(AddressOf X)`)
- Writing a comment to the right of a line-continuation character (underscore) (on the same line) is a syntax error
- Place module-level variables and constants at the top of the module (before all procedures)
- ★In VBA, writing the function name on the right-hand side inside a Function is a recursive call★ (avoid it with a temporary variable)

---

## ★Known bugs in tool usage (specific to the chat environment. Does not apply to Claude Code)★

A bug occurs frequently where a tool call such as `present_files` / `view` is **not executed but displayed as literal text
in the chat, and Claude's turn ends**. It is especially easy to step on when calling `view` repeatedly in small pieces. Countermeasures:

- **Move checking file contents over to `bash_tool` + Python.** When you want to look at several places, combine them in
  a single bash call. **It has not occurred since 9.24. Consolidating into bash works.**
- ★When handling Japanese files with `bash_tool`, the shell's `grep -o` + `head -c` or `cut -c`
  give an "invalid UTF-8" error★ **Cut with Python.**
- ★When reading a file with Python, add `newline=""`★ The universal newlines of text mode
  normalize CRLF to LF, so splitting on CRLF
  silently stops working. **In D-2 the inspection script came up empty because of this.**
- ★Generate large files in `/home/claude/out/` and hand them over with `present_files`★
- ★Edit by "replace with Python + assert the match count"★ Zero accidents. Continue this.
  Writing the replacement text **out to a separate file (`frag_*.txt`) with a heredoc and loading it from Python**
  keeps the quotes and backslashes of VBA from interfering with
  Python's string literals

### ★The standard set of inspections (one more added in D-2: now 4 points)★

1. **Formatting and syntax inspection**: BOM / CRLF / double CR / bare LF / parity of quotes / comments after a line continuation /
   matching open and close of Sub, Function, Property, Enum / balance of If, Do, Select, For, With / target of Exit
2. **For rounds heavy in JS, "reassemble the JS inside the VBA and verify it"**
   ★Strengthened in D-2: not just a syntax check but **actually running it in node**★
   We wrote a minimal DOM stub and could confirm the behavior too (11 cases).
   **Assert that the same fragment exists in the VBA source, to prevent mix-ups.**
3. **Reserved word check**: pick up `Dim x As` / `ByVal x As` / `Private x As` with regular expressions and
   match them lowercased against a set of about 170 VBA reserved words
4. **★Full-count comparison against the original (made standard in D-2)★** Compare the converted file against the original with `Counter`,
   and confirm that every vanished line is an intended deletion. In D-2 we could guarantee that the 64 vanished lines of
   `Wv2Pane` were all the moved `JsonPickStr` and the 3 rewritten call lines

### ★Algorithms can be verified in Python first★
In D-1 we reproduced `JsonPickStr`, and in D-2 `JsonUnescape`, in Python, verified them with real data, and then brought them down to
VBA. It is effective, so continue.
Note: **when everything fails or part fails, first suspect the inspection side.**
  D-1: `json.dumps` puts a space after the colon but JS's `JSON.stringify` does not.
  D-2: with Python `str` concatenation, the halves of a surrogate do not recombine (a VBA String is a
  sequence of UTF-16 code units, so concatenating gives the correct pair). **Both times the cause was the inspection side.**

---

## ★Spec facts that are especially easy to forget★

### Spec fact 20 (stage 9.22) — the structural window of the EBMode gate
**Phenomenon**: If you break or step through in the debugger during or right after the run of `StartWebView2_Full`,
Excel crashes (AV). **Cause**: the 91-byte EBMode check stub that every COM callback passes through can branch only on
"the value of EBMode at the time of entering the stub", and a moment remains, structurally, in which the break transition races after the check passes.
**The countermeasure is avoidance (an operating rule)**: do not break or step through until the WebView2 event burst
has settled.
★From D-1 on, in addition to this, "do not stop during the DoEvents loop of `EvalSync`" is needed★

### Spec fact 14 and the in-callback guard (D-1)
If you call `EvalSync` from inside an event handler (`View_On*`), the runtime waits for the handler to return and
we wait for the JS to complete, so it freezes until timeout. We put a guard in D-1, so
**now it returns immediately with `LastEvalError = "in-callback"`**. When you need the result of JS inside a handler,
call it after a beat with SafeTimer (the same pattern as `RequestNewTabAsync`).
★All methods of `Wv2Element` in D-2 also go through this path★ (`LastError = "in-callback"`).

### Spec facts 30 / 33 (D-1 / D-2) — escaping of the result JSON
The opening angle bracket arrives escaped as `<`, but **the closing one arrives as is**
(D-2 measurement: `at <anonymous>`). `Wv2Json.JsonUnescapeAt` decodes it, so
the VBA side gets a plain string, but when you write a simulator, match it to the real thing.

### Key points of the D-2 design (for re-confirmation)
- **The generation is issued by JS** (`'g' + Date.now() + '-' + Math.random()`).
  **The authority of the verdict is always on the JS side** (design principle 91). `Wv2Pane.CurrentDomGen` is
  only a "fast path that errs only toward the safe side"
- **Do not inject the JS-side registry.** Embed a self-generating prologue at the head of every search expression
  (design principle 92). It self-heals across page navigation, SPA rewrites, and `about:blank`
- **The meaning of `Nothing` versus an empty string is distinguished with `LastEvalOk` / `LastOk`** (design principle 93)

---

## Development cycle (strictly observed)

(1) Align on the policy (present several options + state Claude's recommendation) → tarboh chooses →
(2) implement → (3) share the real-machine verification log → (4) record in the development notes → next stage.
**Never start implementing before the policy is agreed.** Keep each stage to a small scope that can be verified independently on the real machine.
If you find a new technical risk right before implementing, do not handle it on your own; **raise it on the spot as an issue and get agreement**.

- The form of **agreeing on the issues in a batch** has become established (stage 9.27 had 7, stage 9.32 had 7, D-1 had 8, **D-2 had 9**).
- **Always report defensive implementations and small policy adjustments after the fact** (design principle 79). In D-2 we reported 3.
- ★Lesson of stage 9.30★ If the agreed first move turns out to be wrong in prior investigation, replace the first move on the spot and re-agree.
- ★Lesson of stage 9.32b★ If, while preparing the implementation, you find "another problem we will certainly step on", add it to the same round's issues and close it together.
- ★Lesson of D-1★ **When tarboh points out an error in a premise, immediately rebuild the evaluation.**
  Also, **it is worth increasing the inspections before handing over the code**.
- ★Lesson of D-2★ **Do not be satisfied with counting the members at design time; always ask "can the caller
  decide from this return value alone?"** (the origin of design principle 93. `LastOk` / `LastError` were added afterward).
  Also, **a regression check can dig up an existing bug** (the settings tab bug). When you find one,
  do not fix it on your own; identify only the cause and pile it onto the open items.

---

## ★Context management (preventing recurrence)★

In the past there was an accident in which **the context hit its limit and we could not even output the handover prompt**.
When Claude feels the context is getting tight (after a long run of pasted real-machine logs,
after completing one stage's worth of work, etc.), **while there is still room**, it should proactively suggest "it may be about time
to hand over to a new chat" and propose outputting the handover prompt.
Do not wait until the very limit.
