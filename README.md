# VBA + WebView2 for Excel

Host **Microsoft Edge WebView2** directly on an Excel VBA UserForm and drive it from VBA:
no .NET wrapper, no ActiveX/COM registration, no references to set, no admin rights.

English fork of [tarboh/WebView2-For-Excel-VBA](https://github.com/tarboh/WebView2-For-Excel-VBA)
(MIT licence, original author tarboh). Translated, reviewed and patched; see
[`docs/REVIEW.md`](docs/REVIEW.md) for the audit and the list of changes.

## What it does

- Creates a WebView2 environment through `CreateCoreWebView2EnvironmentWithOptions` in
  `WebView2Loader.dll`, then calls every COM method by vtable index with `DispCallFunc`.
- COM callbacks (environment created, navigation completed, script results, web messages,
  resource requests, …) are received through hand-built IUnknown vtables plus small x64
  machine-code thunks that forward into VBA class methods.
- On top of that sits a tabbed browser UserForm and an automation API:
  - **Page automation** — `QuerySelector`, `QuerySelectorAll`, `GetElementById` → `Wv2Element`
    with `.Value`, `.Click`, `.SetAttribute`; `WaitFor`, `WaitGone`, `WaitSettled`,
    `ArmContentSignal` / `ArmNetworkSignal` / `ArmUrlSignal`; synchronous `EvalSync` for JS.
  - **Network inspection** — capture requests/responses (`NetCaptureStart`, `NetDetail*`,
    `NetDetailRespBody`), with cookie/auth headers redacted by default.
  - **HTML GUI for UserForms** — tab bar, navigation bar and settings page are HTML inside WebView2.
  - Example: `Wv2Maps.MapsGeocodeSheet` geocodes a column of addresses via Google Maps.

## Requirements

- Windows 10/11 with the **WebView2 Runtime** (preinstalled on current Windows).
- **64-bit** Microsoft 365 / Office 2016+ Excel. 32-bit Office is not supported (the thunks are
  x64 code).
- `WebView2Loader.dll` (x64). Most Office installs already have one — see
  [Locating WebView2Loader.dll](#locating-webview2loaderdll).
- Macros enabled for the workbook.

## Quick start

1. Open `book/WebView2_v0.5.2.xlsm` (keep it outside OneDrive — AutoSave fights the import tools).
2. In the VBA editor's Immediate window run the smoke test. It opens example.com and prints
   `Wv2.Demo: PASS`:

   ```vba
   Wv2.Demo
   ```

3. Automate a page from your own macro. Everything goes through the `Wv2` module:

   ```vba
   Dim pg As Wv2Pane
   Dim el As Wv2Element
   Set pg = Wv2.OpenBrowser("https://example.com")   ' shows the browser, waits for the load
   If pg Is Nothing Then Exit Sub
   Set el = pg.WaitFor("h1", 10)                     ' polls the DOM until the element exists
   If Not el Is Nothing Then Debug.Print el.InnerText
   Wv2.CloseBrowser
   ```

**Full guide: [`docs/USAGE.md`](docs/USAGE.md)** — first-time setup, the `Wv2` API, waiting,
JavaScript, network capture, using it in your own workbook, and troubleshooting.

The workbook in `book/` still holds the upstream code until you run the import step in the guide.

Never break into the debugger or single-step while WebView2 is starting or busy (including while
`EvalSync` waits) — the callback gate cannot protect that moment and Excel will crash.

## Locating WebView2Loader.dll

The loader is not part of Windows. At start-up the code looks for it in this order and loads the
first one that works:

1. the full path in the environment variable `WEBVIEW2_LOADER_PATH`;
2. the workbook's folder;
3. Excel's folder;
4. the copy Office ships with Power Query (`...\Office16\ADDINS\Microsoft Power Query for Excel Integrated\bin`).

If none is found you get a message box. Take the x64 `WebView2Loader.dll` from the
[Microsoft.Web.WebView2](https://www.nuget.org/packages/Microsoft.Web.WebView2) NuGet package
(`runtimes/win-x64/native/`) and put it next to the workbook.

## Repository layout

| Path | Contents |
|---|---|
| `src/` | VBA modules and classes — the source of truth |
| `forms/` | UserForms (imported by hand) |
| `book/` | The workbook |
| `tools/` | `export.ps1` / `import.ps1` round-trip, checks, `vtable.py` |
| `docs/` | **`USAGE.md` (how to use it)**, the review, the thunk write-up, design notes |

Developer rules and workflow: [`CLAUDE.md`](CLAUDE.md).

## Troubleshooting: Excel crashes when the form opens

VBA sometimes does not resolve `AddressOf` targets in a standard module that are never referenced
elsewhere, leaving a bad pointer in the hand-built vtable. Keeping a dummy reference forces it:

```vba
' Forces VBA to resolve the AddressOf target
Public Sub RegisterNavigationCompleted_()
    Static vTable As LongPtr
    vTable = GetAddr(AddressOf Handler_QueryInterface)
End Sub
```

## Credits

Original project by **tarboh**. The original author thanks **Mokkun**
([reference article](https://eschamali.github.io/StarterWebScrapingKit/#userform-powershell))
for the key technical idea. Licensed under MIT — see [LICENSE](LICENSE).
