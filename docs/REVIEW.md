# Review of tarboh/WebView2-For-Excel-VBA (English fork)

Reviewed 2026-11-10 at upstream commit `7da8e24` ("N-4"). Branch `english` of this fork.

**Verdict:** the project is what it says it is. It is a genuine, carefully engineered pure-VBA
WebView2 host. I found no malicious code, no telemetry, no downloads and no hidden behaviour. Its
low-level COM layer is correct against Microsoft's official SDK header. The bugs I did find are
robustness and security-hardening issues, not deception. They are fixed in this branch as far as
possible without running Excel; see [Not verified](#not-verified).

---

## 1. Is it safe / what it claims?

| Check | Result |
|---|---|
| Win32 APIs declared | Only memory (`VirtualAlloc/Free/Query`, `RtlMoveMemory`), windows (`SetWindowPos`, `GetClientRect`, …), timers, env vars, `DispCallFunc`, `CoTaskMemFree`, `MultiByteToWideChar`, plus the WebView2 loader. No process creation, registry, network or shell APIs. |
| Executable memory | One RWX region of thunks. Each thunk is 74 bytes of x64 code, disassembled by hand: an alive-flag check, then a call that injects `Me` into a VBA class method. It only calls back into this project's own VBA. |
| Network access | Only what the browser itself loads. The test URLs are `example.com`, `httpbin.org`, `httpbingo.org`, `postman-echo.com`, Bing, Google and DuckDuckGo. No phone-home. |
| File writes | Log files in `%APPDATA%\Wv2Browser\logs` (rotation deletes only its own `wv2_*.log`), the settings file, and HTML for the tab and navigation bars. |
| Workbook vs source | The VBA in `book/WebView2_v0.5.2.xlsm` was extracted with `olevba` and diffed against `src/` and `forms/`. They are identical apart from one identifier's casing (`cbId`/`cbid`). olevba's "VBA stomping" flag is a false positive caused by that difference. |
| Auto-run macros | None in the current workbook: no `Workbook_Open` or `Auto_Open`. `UserForm_Resize` only fires when you open the form. |
| `Old/` workbooks (removed in this branch) | Historical versions. Some of them use `Shell` and `CallByName` as experiments. They are not needed and remain only in git history. |

## 2. COM correctness: checked against WebView2.h

Every raw vtable call was mapped to its interface and method and compared with
`WebView2.h` from **Microsoft.Web.WebView2 1.0.4314.35**. There are 75 call sites; examples are
`ICoreWebView2` slots 3-71, `Controller`, `Settings`, the event-args interfaces, the
header iterators, `WebResourceResponseView` and `IStream`.

- **Indexes:** all are correct.
- **IIDs:** all 15 (the handler IIDs plus `ICoreWebView2_2` and `_3`) are correct.
- **x64 ABI:** correct for `BOOL`, `LPCWSTR`, `EventRegistrationToken`, `RECT` passed by value, `UINT64` and `IStream::Seek`.
- **Memory ownership:** strings returned as `LPWSTR` are freed with `CoTaskMemFree`.

The upstream rule "never guess a vtable index; extract it with `tools/vtable.py`" has clearly
been followed.

## 3. Issues found and what this branch did

Severity: H = can crash Excel or compromise data; M = leak or security weakness;
L = robustness.

| # | Sev | Issue | Fix in this branch |
|---|---|---|---|
| H1 | H | **Two workbooks freed each other's thunk memory.** The crash-recovery "sentinel" stored the region address in a process-wide environment variable with a fixed name (`WV2_VBA_LastRegion`). A second workbook containing this code in the same Excel process freed the first workbook's live thunk region, and the first workbook's next callback crashed Excel. | The variable name is now suffixed with a hash of `ThisWorkbook.FullName` (`Wv2Thunks.SentinelEnvName`). |
| H2 | H | After a VBA reset, the thunks still point at destroyed objects. Safety relies on the host window being destroyed (UserForm). If you host the browser in a window that outlives the reset, the next event can crash. | **Documented, not re-engineered:** the browser must be hosted in a UserForm. Making IUnknown go through checked native thunks is too risky without a test machine. |
| M1 | M | When a WebView2 call failed synchronously (`add_*`, `ExecuteScript`, `CreateController`, `GetContent`, environment creation), the callback slot leaked and kept its owner object alive. | New `Wv2Thunks.Thunks_DiscardHandler`, called on all 12 failure paths. It only acts while WebView2 holds no reference to the handler, and the slot must still belong to that handler. |
| M2 | M | Running `StartWebView2_*` twice orphaned the old browser and bars, which stayed alive and visible because of reference cycles. | `UserForm1.AlreadyStarted` makes a second start a no-op with a message. New `Wv2Browser.CloseAll` closes every controller when you want to tear down but keep the window. It is not on the normal form-close path because `Ctrl_Close` costs about 1.4 s per pane (K-4). |
| M3 | M | **The settings tab exposed a privileged host object to any site.** After the Settings tab navigated to a web page, that site could call `chrome.webview.hostObjects.Settings.SetCustomEngine("https://evil/?q=")` and redirect every future search. Any page could also claim to be the settings view (`window.__wv2Settings`). | `NavigationStarting` now removes the `Settings` host object (`View_RemoveHostObjectFromScript`, vtable 50, checked against the header) when that pane goes to `http:`/`https:`/`file:`. `IsSettingsView` also requires the pane's own `IsSettings` flag. Custom templates with control characters are rejected, so they can't inject lines into the settings file. |
| M4 | M | Hard-coded `C:\Temp\VBA_WebView2*` held the profile (cookies) and the HTML served to pages that hold host objects. Folders under `C:\` are usually writable by all users. | All of these now live under `%LOCALAPPDATA%\Wv2Browser\<Profile/TabBar/NavBar/SPA>` (`Wv2Log.AppFolder`, which creates parent folders as needed). |
| M5 | M | An error raised inside a callback returned uninitialised stack data to WebView2 as the HRESULT. | `ComCallbackHandler.Handler_Invoke` now has an error boundary: it logs the error with `Wv2Log.LogE` and returns `S_OK`. |
| M6 | M | `WebView2Loader.dll` was declared without a path. It worked only if Power Query had already loaded Office's copy; otherwise you got runtime error 53. | `Wv2Thunks.EnsureWebView2Loader` loads it by full path: `%WEBVIEW2_LOADER_PATH%`, then the workbook folder, then `Application.Path`, then Office's Power Query `bin` folder. If none is found you get a clear message. |
| L1 | L | Wait loops used `Timer - t`, so they never timed out across midnight. | Replaced with `Wv2Log.SecondsSince` / `ElapsedSec`. |
| L3 | L | `Wv2Pane` called `UserForm1` directly, so the class was not reusable without the form, and could silently re-create the form after it was unloaded. | The call was removed; it only printed values that are already printed. |
| L7 | L | `JsonGetNum` raised an overflow error on large page-supplied numbers. | Values are now clamped. |
| L9 | L | Virtual-host folder mappings used `ALLOW`, which lets any site read those local files. | Changed to `DENY_CORS` (2). |
| — | M | **Source encoding.** VBA stores source in the project's ANSI code page. The book is code page 932 (Japanese); Western Windows uses 1252. Any Japanese left in `src/` would turn into `?` on import there. | The source is now **100% ASCII**. Comments were translated. Japanese test data (URL-encoding inputs, Japanese address samples, DOM text) is built with `ChrW(&Hxxxx&)`. A decode-back check proved every value is byte-identical to the original. The new `tools/check_ascii.py` gate runs in `import.ps1`. The three Japanese-named subs in `Module1` were folded into `Boot`/`BootNoQuiet`/`BootMaps`. |
| — | L | `tools/vtable.py` had the original author's personal paths. | It now uses `WEBVIEW2_H`, or `tools/WebView2.h` (git-ignored). |

**Not changed (by design or low value):**
- **L2:** `DoEvents` re-entrancy in the settings sync loops.
- **L4:** request-body capture on non-seekable uploads. This is opt-in through `NetDetailOn`.
- **L5:** request and response bodies are logged when detail capture is on. This is opt-in; header redaction is on by default.
- **L6:** refcounts can drift if the debugger breaks inside a callback. Upstream already says not to break during callbacks (spec fact 20).
- **L8:** `dcf` maps unknown argument types to `VT_I8`. All current callers pass `Long`/`LongPtr`.
- **L10:** 32-bit Office. The project fails to compile cleanly via `#Error`.

## 4. Translation

- **What was translated:**
  - All 17 modules, both forms, the 8 tools.
  - `README.md`, `CLAUDE.md`, the docs index, the thunk write-up, and the D, E, K and N dev notes.
- **How it was checked:** `tools/check_translation.py` strips comments and string literals and compares the remaining code token-for-token with the original. It reported "code identical" for every module before the fixes above were applied.
- **Still Japanese:** the frozen history volumes `docs/dev-notes-1-foundation.ja.md`, `dev-notes-2-browser.ja.md` and `handover-at-D2.ja.md` (about 13k lines).

## Not verified

**Excel was not available on the review machine, so nothing was run.** All checks were static:
- the SDK header comparison;
- the project's own `check_vba.py` and `check_idents.py`;
- `check_translation.py` and the ASCII decode-back check.

Before relying on this branch:

1. **Push the source into the workbook.** On a PC with 64-bit Excel, enable *Trust access to the VBA project
   object model*, close Excel, then run:
   ```
   powershell -ExecutionPolicy Bypass -File tools/import.ps1 -Apply -AllowNew
   python tools/check_book.py
   ```
   `UserForm1.frm` must be re-imported by hand: remove `UserForm1`, then import `forms/UserForm1.frm`.
2. **Get the loader DLL.** Put `WebView2Loader.dll` (x64, from the NuGet package) next to the workbook, unless the
   Power Query copy exists.
3. **Run the tests.** Run `Boot`, then the `Test_*` subs (with `UserForm1` shown), and read the verdict lines in the log. Pay special attention to:
   - the settings tab (open Settings, then navigate that tab to a website);
   - restart (`StartWebView2_Full` twice);
   - the tab bar and nav bar after the `DENY_CORS` change.
