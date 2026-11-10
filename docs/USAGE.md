# Using WebView2 from VBA

How to set this project up and drive a real Edge (WebView2) browser from your own Excel macros.
Everything goes through one small module, **`Wv2`**.

> **Status:** the code has passed every static check, but this fork has not yet been run in
> Excel. Do [First-time setup](#1-first-time-setup) and run `Wv2.Demo` before building on it.

## Contents

1. [First-time setup](#1-first-time-setup)
2. [Your first macro](#2-your-first-macro)
3. [The `Wv2` module](#3-the-wv2-module)
4. [Working with a page](#4-working-with-a-page)
5. [Waiting properly](#5-waiting-properly)
6. [Running JavaScript](#6-running-javascript)
7. [Watching network traffic](#7-watching-network-traffic)
8. [Using it in your own workbook](#8-using-it-in-your-own-workbook)
9. [Rules that prevent crashes](#9-rules-that-prevent-crashes)
10. [Troubleshooting](#10-troubleshooting)

---

## 1. First-time setup

**You need:** Windows 10/11, **64-bit** Excel (Microsoft 365 or 2016+), and the WebView2 Runtime
(already installed on current Windows).

1. **Copy the repo to a local folder outside OneDrive.** OneDrive's AutoSave overwrites what the
   import tool writes.
2. **Allow the import tool to write VBA.** In Excel: *File → Options → Trust Center → Trust Center
   Settings → Macro Settings →* tick **Trust access to the VBA project object model**.
3. **Close Excel**, then push the source into the workbook:

   ```
   powershell -ExecutionPolicy Bypass -File tools/import.ps1 -Apply -AllowNew
   python tools/check_book.py
   ```

4. **Replace the form by hand** (the tool does not import forms). Open
   `book/WebView2_v0.5.2.xlsm`, press `Alt+F11`, right-click `UserForm1` → *Remove* (answer *No* to
   export), then *File → Import File…* → `forms/UserForm1.frm`. Save.
5. **Provide `WebView2Loader.dll` (x64).** Most Office installs already have one and it is found
   automatically. If start-up reports that it was not found, download the
   [Microsoft.Web.WebView2](https://www.nuget.org/packages/Microsoft.Web.WebView2) NuGet package
   (a `.nupkg` is a zip), take `runtimes/win-x64/native/WebView2Loader.dll` and put it in the
   same folder as the workbook.
6. **Smoke test.** In the VBA editor open the Immediate window (`Ctrl+G`) and run:

   ```vba
   Wv2.Demo
   ```

   A browser window opens on example.com and the Immediate window ends with `Wv2.Demo: PASS`.

## 2. Your first macro

```vba
Sub ReadHeadline()
    Dim pg As Wv2Pane
    Dim el As Wv2Element

    Set pg = Wv2.OpenBrowser("https://example.com")
    If pg Is Nothing Then
        MsgBox "The browser did not start or the page did not load."
        Exit Sub
    End If

    Set el = pg.WaitFor("h1", 10)             ' CSS selector, wait up to 10 s
    If Not el Is Nothing Then Range("A1").Value = el.InnerText

    Wv2.CloseBrowser
End Sub
```

`OpenBrowser` shows the browser window, starts WebView2, loads the page and returns once it has
loaded. Every later call works on that same browser until `Wv2.CloseBrowser` (or the user closes
the window).

## 3. The `Wv2` module

| Call | What it does | Returns |
|---|---|---|
| `Wv2.OpenBrowser(url, withBars, timeoutSec)` | Opens the browser and loads `url`. `withBars:=False` hides the tab bar and address bar. If the browser is already open, opens a new tab. | The tab (`Wv2Pane`), or `Nothing` on failure |
| `Wv2.NewTab(url, timeoutSec)` | Opens another tab and waits for it to load. | The tab, or `Nothing` |
| `Wv2.Navigate(url, timeoutSec, pane)` | Loads `url` in the active tab (or `pane`) and waits. | `True` if it loaded |
| `Wv2.WaitLoaded(pane, timeoutSec)` | Waits for the load in progress to finish. | `True` if it loaded successfully |
| `Wv2.Page` | The active tab. | `Wv2Pane` or `Nothing` |
| `Wv2.ActiveBrowser` | The browser object: tabs, search engine, settings. | `Wv2Browser` or `Nothing` |
| `Wv2.IsBrowserOpen` | Whether the browser is running. | `Boolean` |
| `Wv2.Eval(js, timeoutSec, pane)` | Runs a JavaScript expression. | Result as JSON text |
| `Wv2.EvalText(js, timeoutSec, pane)` | Same, with a string result decoded. | Plain text |
| `Wv2.CloseBrowser` | Closes all tabs and the window. | — |
| `Wv2.Demo` | Smoke test. | Prints PASS / FAIL |

Defaults: `url` = `about:blank`, `withBars` = `True`, load timeouts = 30 s, JS timeout = 5 s.
A URL without a scheme gets `https://`; a path such as `C:\site\index.html` becomes a `file:` URL.

Nothing here raises an error for an ordinary failure. Check the return value; the reason is in the
log file (`?Wv2Log.LogPath`, under `%APPDATA%\Wv2Browser\logs`).

Cookies and logins are kept between runs in `%LOCALAPPDATA%\Wv2Browser\Profile`. Delete that
folder (with the browser closed) for a clean profile.

## 4. Working with a page

A tab is a `Wv2Pane`. An element is a `Wv2Element`.

```vba
Dim pg As Wv2Pane, box As Wv2Element, btn As Wv2Element, row As Variant

Set pg = Wv2.Page

' Find
Set box = pg.QuerySelector("input[name='q']")       ' Nothing if not found
Set btn = pg.GetElementById("search-button")
For Each row In pg.QuerySelectorAll("table#results tr")
    Debug.Print row.InnerText
Next row

' Read
Debug.Print box.Value, box.TagName, box.GetAttribute("placeholder"), box.InnerHTML

' Write and click (uses the browser's native setters, so React/Vue pages notice)
box.Value = "excel vba"
If Not btn.Click() Then Debug.Print btn.LastError
```

- A lookup returns `Nothing` when the element is not there. Always test before use.
- After any element call, `el.LastOk` / `el.LastError` say whether it worked. An empty string with
  `LastOk = True` is a real empty value.
- When the page reloads or re-renders, old elements go stale (`el.IsStale`). Look them up again.
- `QuerySelectorAll` returns at most a fixed number of elements; `pg.LastAllTruncated` says
  whether the list was cut. Pass a larger `maxCount` as the second argument if needed.

## 5. Waiting properly

Web pages change after they "finish loading". Never use fixed sleeps; wait for what you need.

| You want | Use |
|---|---|
| An element to appear | `pg.WaitFor("css selector", seconds)` → the element, or `Nothing` on timeout |
| An element to disappear (spinner, dialog) | `pg.WaitGone("css selector", seconds)` |
| A full page load after `Navigate` or a link click | `Wv2.WaitLoaded` |
| The page to go quiet (no network, no DOM changes) | `pg.WaitSettled(seconds)` |
| Every lookup to wait a little by default | `pg.ImplicitWaitSec = 5` |

`WaitSettled = True` only means "quiet", not "the app finished". For a step that matters (submit,
save, search) arm a signal for the thing you expect **just before** acting, then wait:

```vba
pg.ArmContentSignal "div.results"        ' or ArmNetworkSignal "/api/search", ArmUrlSignal "/done"
btn.Click
If Not pg.WaitSettled(20) Then Debug.Print "no result: " & pg.LastSettleInfo
```

## 6. Running JavaScript

```vba
Debug.Print Wv2.EvalText("document.title")                    ' Example Domain
Debug.Print Wv2.Eval("document.querySelectorAll('a').length") ' 1
Debug.Print Wv2.Eval("({a:1, b:[2,3]})")                      ' {"a":1,"b":[2,3]}
```

- Pass an **expression**. For several statements wrap them:
  `"(function(){ var x = 1; return x + 1; })()"`.
- **Promises are not awaited.** An expression that returns a Promise comes back as `{}`. Have the
  page store the result (for example in a `window` variable) and read it with a second call.
- On failure the result is `""`: check `Wv2.Page.LastEvalOk` and `Wv2.Page.LastEvalError`.
- **Quoting.** Use single quotes inside the JavaScript. To insert a VBA string, let JSON do the
  quoting instead of gluing quotes yourself:

  ```vba
  js = "document.querySelector('#name').value = " & Chr$(34) & Wv2Json.JsonEscape(userText) & Chr$(34)
  ```

- Read JSON results with `Wv2Json.JsonGetStr` / `JsonGetNum`.

## 7. Watching network traffic

```vba
Dim pg As Wv2Pane, i As Long
Set pg = Wv2.Page

pg.NetDetailOn = True                 ' keep headers and bodies, not just the list
pg.NetCaptureStart , "*api*"          ' URL filter with wildcards; "*" = everything
Wv2.Navigate "https://httpbin.org/get"

For i = 1 To pg.NetLogCount
    Debug.Print pg.NetLogLine(i)      ' number, ms, method, kind, URL, status, reason (tab-separated)
Next i
pg.NetCaptureStop
```

Cookie and authorization headers are redacted in the capture by default (`pg.NetRedact`).
Response bodies need `pg.NetRespBodyOn = True` and are read with `pg.NetDetailRespBody(i)`.
Details: `docs/dev-notes-6-N-axis.md`.

## 8. Using it in your own workbook

The simplest route is to keep working in a copy of `book/WebView2_v0.5.2.xlsm`. To add the browser
to another 64-bit workbook instead:

1. In the VBA editor use *File → Import File…* for every file in `src/` **except**
   `Wv2Tests.bas` and `Module1.bas` (tests and experiments), and for both forms in `forms/`
   (`UserForm1.frm`, `TimerForm.frm`; keep the `.frx` files beside them). Do not import
   `src/document/`.
2. *Debug → Compile VBAProject* must finish without errors.
3. Run `Wv2.Demo`.

No references need ticking and nothing is registered on the PC, so the workbook can be handed to
another user as a single file (plus `WebView2Loader.dll` if their Office has none).

**Your own HTML as the user interface.** Put your HTML in a folder and open it without the bars:

```vba
Set pg = Wv2.OpenBrowser("C:\MyApp\ui\index.html", withBars:=False)
```

The window's size and caption are those of `UserForm1` — change them in the form designer.

## 9. Rules that prevent crashes

These come from how the callbacks are built; breaking them takes Excel down with unsaved work.

1. **Never pause the code while the browser is busy.** No breakpoints, `Stop` or single-stepping
   (`F8`) while a page loads or while any wait / `Eval` call is running. Debug with
   `Debug.Print` or `Wv2Log.LogI`, or stop on a line where nothing is in flight.
2. **Save before you run** while developing.
3. **Close the browser before pressing Reset** (the square Stop button) in the VBA editor, and
   before closing the workbook.
4. **Keep the browser in `UserForm1`.** Do not host it in a window that survives a VBA reset.
5. **Do not call the wait or `Eval` functions from inside a browser event handler.** They return
   failure at once there (they would otherwise deadlock). Start a `SafeTimer` and do the work
   from its tick.
6. **Do not type with an IME (Japanese/Chinese/Korean input) in the Immediate window** while the
   browser is open. This is a known upstream crash.
7. **One Excel task at a time.** While a macro is waiting, the user can still click around Excel;
   do not let them start a second macro that drives the browser.

## 10. Troubleshooting

| Symptom | Cause and fix |
|---|---|
| "WebView2Loader.dll (x64) was not found" | Put the x64 DLL beside the workbook (setup step 5), or set the environment variable `WEBVIEW2_LOADER_PATH` to its full path. |
| Compile error on open, `#Error` line highlighted | 32-bit Office. Only 64-bit is supported. |
| `OpenBrowser` returns `Nothing`, window closes again | Start-up failed. Read the last lines of the log (`?Wv2Log.LogPath`). Most often the WebView2 Runtime is missing or the loader DLL was not found. |
| `OpenBrowser` returns `Nothing`, window stays open | The page did not load in time or the address failed. `Wv2.Page` is still usable; retry with `Wv2.Navigate`. |
| "already started on this form" | The browser is already running. Use `Wv2.NewTab` / `Wv2.Navigate`, or `Wv2.CloseBrowser` first. |
| A lookup returns `Nothing` on a page that shows the element | It was not there yet (use `WaitFor`), or it is inside an `<iframe>` — frames are not searched. |
| Excel crashes when the form opens after you edited modules | VBA lost an `AddressOf` target. Run *Debug → Compile VBAProject*, save, restart Excel. |
| Module shows `?` where text used to be | A non-ASCII character went into the source. Keep source ASCII; build other text with `ChrW`. `python tools/check_ascii.py` finds them. |
| Import tool says nothing was written | Excel was still open, or the workbook is in OneDrive. |

More depth: [`REVIEW.md`](REVIEW.md) (audit and known limits), [`../CLAUDE.md`](../CLAUDE.md)
(rules for changing the code), `dev-notes-0-index.md` (design history).
