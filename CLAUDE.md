# VBA-WebView2 (English fork)

Microsoft Edge **WebView2** hosted directly on an Excel VBA UserForm, in **pure VBA** — no
external COM library, no ActiveX reference, no registry changes. It uses only x64 machine-code
thunks, SAFEARRAY-based memory primitives and hand-built COM vtables called through
`DispCallFunc`. **64-bit Office only.**

This is an English fork of [tarboh/WebView2-For-Excel-VBA](https://github.com/tarboh/WebView2-For-Excel-VBA)
(MIT). Upstream history is kept; the upstream remote is named `upstream`. Working language of this
fork: **English** (code, comments, docs, commit messages).

Status inherited from upstream: usable tabbed browser; page automation ("D axis", an MSHTML-like
API) complete through D-7; network inspection ("N axis") complete through N-4.

## Concept (what to optimise for)

Not a full browser clone. Two goals:

1. **Be excellent at automating web pages from VBA.**
2. **Give UserForms a modern GUI by laying out controls in HTML.**

When in doubt, ask whether a change serves one of these two.

---

## Where the truth lives

| Path | Contents | Notes |
|---|---|---|
| `src/` | 8 standard modules + 9 classes | **Source of truth. Edit here.** |
| `src/document/` | `ThisWorkbook` / `Sheet1` | Reference only (import cannot replace them) |
| `forms/` | `UserForm1` / `TimerForm` | Managed by hand; not auto-imported |
| `book/` | `WebView2_v0.5.2.xlsm` | Current workbook; commit only at milestones |
| `docs/` | Dev notes (see below) | |
| `tools/` | `export.ps1` / `import.ps1` / `check_*.py` / `vtable.py` | Round-trip with the workbook, checks |

### Round-trip with the workbook

**Workbook → `src/`** (opens the book read-only with macros disabled; never modifies it):

```bash
powershell -ExecutionPolicy Bypass -File tools/export.ps1 -Force
```

Forms are exported only with `-Forms` (`.frx` changes a few bytes on every export).

**`src/` → workbook** (default is a dry run; `-Apply` writes, `-AllowNew` allows new modules,
`-Modules A,B` limits the set):

```bash
powershell -ExecutionPolicy Bypass -File tools/import.ps1
```

- **Close Excel first.** If the book is open, `Workbooks.Open` succeeds read-only and nothing
  reaches the disk; `import.ps1` guards against this.
- A timestamped backup goes to `book/backup/`; only changed modules are written; the book is
  re-read and compared afterwards.
- Existing modules are replaced with `AddFromString`, new ones are added with `Import`
  (`Remove` + `Import` renames the module to `Wv2Pane1` because removal is deferred).
- Member attribute lines (`Attribute m_foo.VB_VarHelpID`) are dropped on write; otherwise they
  duplicate and one becomes visible code → compile error. After `-Apply` run
  `python tools/check_book.py`; if a module got corrupted, `tools/repair_book.ps1 -Modules <name> -Apply`.
- Declaring an identifier that already exists with different casing makes VBA re-case it in the
  whole project; `import.ps1` reports this — then run `export.ps1 -Force`.
- **Keep the workbook outside OneDrive.** OneDrive opens it via a cloud URL with AutoSave on, which
  silently overwrites what `import.ps1` wrote.

### Opening the browser

```vba
UserForm1.Show vbModeless      ' must be shown first, or hWnd_Frame is 0 and startup fails
```

---

## Dev notes (`docs/`)

| File | Contents |
|---|---|
| `dev-notes-0-index.md` | **Start here.** Stage numbering, all design principles (1-122), all spec facts (1-80), roadmap, current state, open items |
| `standard-thunk-method.md` | How the x64 thunk technique works |
| `dev-notes-3-D-axis.md` | D-1 … D-7: page automation |
| `dev-notes-4-E-axis.md` | E-1 … E-4: dev environment / tooling |
| `dev-notes-5-K-axis.md` | K-1 … K-5: foundation / tech-debt |
| `dev-notes-6-N-axis.md` | N-1 … N-4: network inspection |
| `dev-notes-1-foundation.ja.md`, `dev-notes-2-browser.ja.md`, `handover-at-D2.ja.md` | Frozen history, still in Japanese |

Append to the note for the relevant axis; keep the index's open-items section current.
Notes are UTF-8 without BOM, CRLF.

---

## Rules for writing VBA here

- **Update the version header block when you change a module** (newest block on top). It is the
  only way to match a module in the book with `src/`.
- **VBA identifiers are case-insensitive.** `eNum` collides with `Enum`. Check new identifiers with
  `python tools/check_idents.py <names...>` (it compares against reserved words and every existing
  identifier, lower-cased).
- Writing the function name on the right-hand side inside a Function is a recursive call — use a temp.
- The first argument of `WritePtrNatively` is `LONG_PTR` (writing `LongPtr` gives a runtime error).
- `AddressOf` is only allowed as a call argument (`addr = VBA.Int(AddressOf X)`).
- No comment after a line-continuation underscore on the same line (syntax error).
- Module-level variables and constants go at the top of the module.
- **Never guess a COM vtable index.** An index off by one calls a different method and kills the
  process. Get it from the SDK header: `python tools/vtable.py ICoreWebView2` (set `WEBVIEW2_H` to
  the path of `WebView2.h` from the Microsoft.Web.WebView2 NuGet package).
- **Source text must stay ASCII.** VBA stores source in the project's ANSI code page; anything that
  code page cannot represent becomes `?` on import and is lost forever. Non-ASCII data (e.g. the
  Japanese test inputs) must be built at runtime with `ChrW(&Hxxxx&)`. `import.ps1` runs
  `tools/check_cp932.py` and `tools/check_ascii.py` before every import and aborts on a hit.

### JavaScript inside VBA strings

Never put backslash + double quote in VBA-built JS. Use single quotes in JS, and quote VBA values
with `Wv2Pane.JsQuote(s)` (= `Chr$(34) & Wv2Json.JsonEscape(s) & Chr$(34)`).

### Debug log

```vba
Wv2Log.LogE "..."   ' ERROR (flushed immediately)
Wv2Log.LogW "..."   ' WARN  (flushed immediately)
Wv2Log.LogI "..."   ' INFO
Wv2Log.LogD "..."   ' DEBUG
```

Also echoes to the Immediate window (`Wv2Log.LogEcho = False` turns that off). Files go to
`%APPDATA%\Wv2Browser\logs\`, one per session, 20 kept. The 6-digit line prefix is a sequence
number — a gap means lost log lines. Tests write their verdict (`*Verdicts n: OK x / FAIL y*`)
to the log; read results from the log file, not the Immediate window.

---

## Facts that are easy to forget

- **Spec fact 20 – never break into the debugger while WebView2 is starting or busy.** The
  EBMode gate stub can only check the mode on entry; a break during an event burst crashes Excel
  (access violation). This also applies while `EvalSync` is pumping `DoEvents`.
- **Spec fact 14 – no `EvalSync` from inside an event handler.** It would deadlock; it now
  returns immediately with `LastEvalError = "in-callback"`. Defer with `SafeTimer` instead.
  All `Wv2Element` methods go through the same path.
- **D-2 design:** the DOM generation id is issued by JS and JS is always the authority; no
  injected JS registry (each query carries its own prologue, so navigation/SPA rewrites self-heal);
  `Nothing` vs empty string is distinguished by `LastEvalOk` / `LastOk`.
- **`WaitSettled` = True only means "quiet", not "the app finished".** For important actions, arm an
  explicit signal (`ArmContentSignal` / `ArmNetworkSignal` / `ArmUrlSignal`) right before acting.
- Network (N axis): requests served via `SetVirtualHostNameToFolderMapping` do not raise
  `WebResourceRequested`; JS `fetch()` arrives as `XML_HTTP_REQUEST` (7), not `FETCH`; reading the
  request `IStream` advances it (must `Seek(0)` back); `WebResourceResponseReceived` has no filter;
  `GetContent` returns decoded bodies.

## API overview

Page automation (D axis): `QuerySelector` / `GetElementById` → `Wv2Element`; `QuerySelectorAll` →
`Collection`; `.Value =` / `.Click` / `.SetAttribute` (native setters, so React/Vue see them);
`WaitFor` / `WaitGone` / `ImplicitWaitSec`; `WaitSettled` / `AutoWaitAfterAction`;
`ArmContentSignal` / `ArmNetworkSignal` / `ArmUrlSignal`; `Wv2Maps.MapsGeocode` /
`MapsGeocodeSheet` (address → coordinates via Google Maps).

Network inspection (N axis): `NetCaptureStart` / `NetCaptureStop`, `NetLogDrain` / `NetLogLine(i)`,
`NetDetailOn` + `NetDetail(i)` / `NetDetailHeaders(i)` / `NetDetailBody(i)`, `NetRedact`
(cookies etc. redacted by default), `NetLogStatus(i)`, `NetRespOn`, `NetRespBodyOn` +
`NetDetailRespBody(i)`.

## Verification checklist

1. `python tools/check_vba.py` (format/syntax) and `python tools/check_ascii.py` (encoding).
2. `python tools/check_idents.py <new identifiers>`.
3. After a pure comment/message change: `python tools/check_translation.py <old_dir> src` proves
   the code is unchanged.
4. After `import.ps1 -Apply`: `python tools/check_book.py`.
5. Real-machine tests: run `Test_*` subs from the Immediate window (with `UserForm1` shown) and read
   the verdict lines in the log.
