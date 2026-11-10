# VBA-WebView2 Development Notes ― Axis N (Network Investigation)

★Newly established 2026-08-25.★ The index is `dev-notes-0-index.md`.

---
## What is axis N

★**See what the browser is actually hitting, using nothing but Excel VBA.**★
Build the equivalent of F12's Network tab with zero external libraries.

### Why it fits the concept

The ★first step★ of concept ① (ease of use that excels at automatically controlling web pages)
is always "investigate what to hit". The origin was a voice from real-world practice that this is the most tedious part
(2026-08-25, tarboh):

> I operated a core system, had it spit out a CSV, and automated that with PowerShell.
> To POST over HTTP and read the responses, ★investigating the log in F12's Network tab
> was extremely tedious★. Couldn't we build an investigation tool that extracts only the necessary parts?

★Including the exit, it becomes value that F12 does not have★ ―― it captures a request and
**emits it as PowerShell `Invoke-WebRequest` code**. The real-world flow of "operate once by hand in the browser ->
receive that traffic as PowerShell code -> put it on Task Scheduler"
becomes the tool as it is (there is `Copy as cURL`, but no PowerShell version).

### ★Choice of method (decided)★

| | A. Extend wrapping of in-page JS | B. WebView2 native events |
|---|---|---|
| Implementation | Light (just add to D-4's probe) | Heavy (new wiring of the COM path + `IStream`) |
| `fetch` / `XHR` | Visible | Visible |
| ★Page navigation by form POST★ | ★**Not visible**★ | ★Visible★ |
| Images / CSS / redirects | Not visible | Visible |

★**B is settled**★. Because the target core system is ★centered on form POST (classic style)★ and
A cannot see it at all.

### How the stages are cut

| Stage | Contents |
|---|---|
| **N-1** | Wire up `WebResourceRequested`, giving a ★list of URL + method + type★ |
| N-2 | The request's **headers and body** (reading from `IStream`) |
| N-3 | The response's **status and headers** (`WebResourceResponseReceived`) |
| N-4 | The response **body** |
| N-5 | ★**Emit as PowerShell `Invoke-WebRequest` code**★ |

★N-1 alone is already practical★ ― if you can see "what is hit in what order after operating once by hand",
the part that takes the most time in F12 is taken care of.

---
## Stage N-1 ― Wiring up WebResourceRequested (complete)

### Issues and decisions (8 items agreed in one go)

| # | Issue | Decision |
|---|---|---|
| 1 | Axis and volume | ★Newly establish axis N and this file★ |
| 2 | Method | Native events (above) |
| 3 | Scope of N-1 | ★Only URL / method / type★. Headers, body, and response come in later stages |
| 4 | Default filter | ★Narrow to Document / XHR / Fetch★ (exclude images, CSS, fonts; an entry point for widening is provided) |
| 5 | Handler etiquette | ★Record only and return immediately★ (below) |
| 6 | Where records are kept | ★In-memory ring buffer + retrieval API★ (same pattern as D-4's `SpaProbeDrainLog`) |
| 7 | Exit | N-1 is log only. Sheet output and code generation come in later stages |
| 8 | Verification | ★Make a form POST with our own HTML★ (don't rely on a real site; first see that it is reliably captured) |

★Reason for issue 5★ `WebResourceRequested` is an event for "stopping / rewriting a request", and
★if the handler is slow, the communication itself gets jammed★. It is also territory of spec facts 14 / 20, and
**you must not call `EvalSync` inside the handler**. Record only and return immediately.

★Reason for issue 6★ A single page emits dozens to hundreds of requests. Writing directly to the log would overflow by volume
(the same judgment as when we built `SpaProbeDrainLog` in D-4).

### ★Investigation: pinning down the vtable index★

★**Writing a COM vtable index by guesswork causes a crash.**★
We worked backward from the existing wiring (6 entries such as `remove_NavigationStarting = 8`) and
guessed `add_WebResourceRequested = 54`, but ★**the correct answer was 55**★.
We had missed counting `OpenDevToolsWindow` (51).
**Had we tried it on the real machine, we would have called a different method, off by one, and crashed.**

So we made ★a tool that extracts mechanically from the SDK header (`WebView2.h`)★:

```bash
python tools/vtable.py ICoreWebView2
python tools/vtable.py --enum COREWEBVIEW2_WEB_RESOURCE_CONTEXT
python tools/vtable.py --find WebResource
```

★Cross-check★ We confirmed that `remove_DOMContentLoaded = 65` matches the value in the existing `RemoveVtableIndexFor`.
This supports the tool's correctness.

★Location of the header★ `C:\Users\gugug\Downloads\新しいフォルダー (4)\WebView2.h`
(2.7 MB). ★It is in Downloads, so it can disappear★. It can also be specified with the environment variable `WEBVIEW2_H`, and
it is picked up if placed at `tools/WebView2.h`. **If it disappears, re-fetch it from the SDK.**

#### Numbers used in N-1 (settled)

| Target | index |
|---|---|
| `ICoreWebView2::add_WebResourceRequested` | ★**55**★ |
| `ICoreWebView2::remove_WebResourceRequested` | 56 |
| `ICoreWebView2::AddWebResourceRequestedFilter` | ★**57**★ |
| `ICoreWebView2::RemoveWebResourceRequestedFilter` | 58 |
| `...WebResourceRequestedEventArgs::get_Request` | 3 |
| `...WebResourceRequestedEventArgs::get_ResourceContext` | 7 |
| `ICoreWebView2WebResourceRequest::get_Uri` | 3 |
| `ICoreWebView2WebResourceRequest::get_Method` | 5 |

#### Numbers to be used from N-2 onward (noted in advance)

| Target | index |
|---|---|
| `ICoreWebView2WebResourceRequest::get_Content` (IStream) | 7 |
| `ICoreWebView2WebResourceRequest::get_Headers` | 9 |
| `ICoreWebView2HttpRequestHeaders::GetIterator` | 8 |
| `...HttpHeadersCollectionIterator::GetCurrentHeader` | 3 |
| `...HttpHeadersCollectionIterator::get_HasCurrentHeader` | 4 |
| `...HttpHeadersCollectionIterator::MoveNext` | 5 |
| `ICoreWebView2_2::add_WebResourceResponseReceived` | 61 |

#### `COREWEBVIEW2_WEB_RESOURCE_CONTEXT` (the filter's argument)

| Value | Meaning | N-1 default |
|---|---|---|
| 0 | ALL | ― |
| ★1★ | DOCUMENT | ★Include★ (form POST / page navigation) |
| 2 / 3 / 5 | STYLESHEET / IMAGE / FONT | Exclude |
| 6 | SCRIPT | Exclude |
| ★7★ | XML_HTTP_REQUEST | ★Include★ |
| ★8★ | FETCH | ★Include★ |
| 11 | WEBSOCKET | Exclude (not handled in N-1) |

★If not a single filter is registered, not a single event comes★, so
`AddWebResourceRequestedFilter("*", ctx)` ends up being called 3 times.

### Design principle established (115) ― at the preparation stage

- **Design principle 115 (N-1) Do not write a COM vtable index by guesswork; settle it with the SDK header.**
  This is the COM version of design principle 107 (do not write by guesswork an identifier that grabs an external target), but
  ★the danger level differs★ ―― if a selector misses it just ends in "not found", but
  ★if a vtable index is off by one, a different method is called and the whole process goes down★.
  Use working backward from the existing wiring **only as a cross-check**. In fact, in N-1's preparation
  the backward calculation was off by one. The tool is `tools/vtable.py`.

### Implementation (4 replacements, no new modules)

| File | Contents |
|---|---|
| `Wv2Thunks.bas` | `HK_WebResourceRequested = 11` / extended the upper limit of `m_iidTable` / added IID |
| `ComCallbackHandler.cls` | `Case HK_WebResourceRequested` -> `View_OnWebResourceRequested` |
| `Wv2Pane.cls` | The user-facing entry points + 4 COM wirings + receiver + ring buffer |
| `Wv2Tests.bas` | 6 `Test_N1_*` + our own HTML |

The IID was taken from the `MIDL_INTERFACE` line of `WebView2.h` (design principle 115):
`ICoreWebView2WebResourceRequestedEventHandler` = `{ab00b74c-15f1-4646-80e8-e76341d25d71}`

#### User-facing entry points

```vba
p.NetCaptureStart          ' ★If called before Navigate, capture starts from the first DOCUMENT★
' ... operate once by hand ...
p.NetLogDrain              ' Flush what has accumulated to the log and empty it
p.NetCaptureStop           ' Also remove the filters
```

| Member | Meaning |
|---|---|
| `NetCaptureStart(contexts, uriFilter)` | Start capture. Default is `DOCUMENT,XHR,FETCH` / `"*"` |
| `NetCaptureStop()` | End capture. ★Also removes the filters that were set★ |
| `NetCaptureOn` | Whether capturing |
| `NetLogDrain()` | Flush what has accumulated to `Wv2Log` and empty ★only the contents★ |
| `NetLogLine(i)` | Read one entry without flushing (oldest first, 1-based, 5 `vbTab`-separated items) |
| `NetLogCount / NetLogTotal / NetLogDropped` | On hand / cumulative fires / cumulative overflow |
| `NetLogCapacity` | The ring's capacity (default 500). Changing it rebuilds the contents |
| `NetLogClear()` | Reset everything, including the cumulative figures, to 0 |

The low-level layer is exposed as `View_AddWebResourceRequested` / `View_RemoveWebResourceRequested` /
`View_AddWebResourceRequestedFilter` / `View_RemoveWebResourceRequestedFilter`
(for when you want fine control of the types).

#### ★Thoroughly "record only and return immediately" (issue 5)★

Inside the receiver (`View_OnWebResourceRequested`):

- Fire neither `Wv2Log` nor `Debug.Print` ―― ★the 4 `Debug.Print` lines present in the other `Case`s
  were deliberately removed for this type only★
- Also leave empty the `funcName` passed to `dcf` / `GetStringProperty`, so that even on failure
  nothing is printed to the Immediate window (a failure is left in the ring as `(err)`)
- ★Do not re-raise errors★ ―― the other receivers `Err.Raise`, but if an
  error dialog appears in the middle of communication, the operation cannot continue. The fact that we swallowed it always
  appears in the output of `NetLogDrain`

`get_Request` returns already AddRef'd, so always `ComRelease` it. This is a path passed through hundreds of times
per page, so a leak here leaks visibly.

#### Used `GetTickCount` as the time source for elapsed ms

★`Timer` cannot be used★ ―― `Timer` is a Single, "seconds since midnight", so
at afternoon values (around 50,000 seconds) the Single does not have enough significant digits and the ms resolution disappears.
Moreover it rolls back when crossing a date. `GetTickCount` is ms since boot and increases monotonically.
The return value is a Long so it wraps around in about 49.7 days, but we compute the difference after raising it to Double, and
when negative add 2^32 to bring it back (`NetElapsedMs`).

---

### Cross-checks before implementation (what we crushed before taking it to the real machine)

| Cross-check | Result |
|---|---|
| Reconfirmed vtable indexes with `tools/vtable.py` | 55/56/57/58, args 3/7, request 3/5, all 17 enum values |
| Brute-forced the ring buffer in Python | Capacity 1 to 8 x count 0 to 30 compared with `deque`, **502 cases matched** |
| Reassembled the verification page from the VBA source and confirmed in a real DOM | 15 OK in Edge headless |

★An algorithm can be verified in Python first★ (the same move as D-1's `JsonPickStr` and D-2's
`JsonUnescape`). If we noticed a mistake in the ring's index computation only after writing it in VBA, we would lose
a whole real-machine round trip, so we crushed it first by brute force.

---

### Real-machine test ― ★it took 3 tries★

#### 1st try: 24 checks, OK 14 / FAIL 10 (total fires 0)

**The event never came even once.** Although `add_WebResourceRequested` and
`AddWebResourceRequestedFilter` x 3 were all S_OK, `Invoke` was called 0 times.

★4 of the 14 OKs were vacuous passes★ ("there is not a single `probe.css`",
"not a single one comes after stopping", etc.). Since nothing was captured they passed automatically, and
were not evidence.

**★There was a hole in the test design★** ―― **every single one** of the verification page's communications went
**through the virtual host**. The first DOCUMENT, the fetch / XHR to `data.json`, and the POST to `echo.html` were
all `https://appassets.netprobe/`.
So this log
could **not distinguish at all** between

- (a) the wiring not working (vtable / IID / filter)
- (b) only requests to the virtual host not riding on the event

Because we prioritized "don't rely on a real site" (issue 8), we put in no control at all, and
**we had not fully cross-checked the probe itself** (the same shape of failure as design principle 103).

★The `Test_N1_Site` we had prepared for isolation paid off★ ―― it was set up to print
that guidance to the log automatically on failure, so there was no hesitation about the next move.

#### 1st try (b): `Test_N1_Site` ― 4 checks, OK 4 / FAIL 0

On a real site (Google), **all 10 were captured**.

```
#   Elapsed ms  Method   Type       URL
1       94  GET      DOCUMENT   https://www.google.com/
2      141  GET      DOCUMENT   https://www.google.com/
3      609  GET      XHR        .../async/hpba?...
4      672  OPTIONS  XHR        https://ogads-pa.clients6.google.com/$rpc/...
5      906  GET      XHR        .../xjs/_/js/...
6      953  GET      XHR        .../complete/s?q&...
7     1016  GET      XHR        .../async/hpba?...
8     1062  POST     XHR        https://ogads-pa.clients6.google.com/$rpc/...
9     1500  OPTIONS  XHR        https://play.google.com/log?...
10    1687  POST     XHR        https://play.google.com/log?...
```

- The methods are obtained correctly (GET / POST / **OPTIONS** = even the CORS preflight)
- The types are separated correctly
- `uri="*"` is working (other hosts besides `google.com` are also picked up)
- Not a single image, CSS, or script is present (the default filter is working)

**★The wiring was correct from the start.★** The cause was settled by (b) ―― this is spec fact 69.

#### 2nd try (N-1b): 34 checks, OK 31 / FAIL 3

We reassembled the targets into 2 systems + a control (issue 1, plan G / issue 2, plan a).

| | Target | Nature |
|---|---|---|
| **Plan F** | `http://127.0.0.1:59999/…` | Unreachable. ★Zero external dependencies★ |
| **Plan D** | `https://httpbingo.org/…` | Needs the net, but arrives straightforwardly |
| **Control** | The virtual host's `data.json` | ★Count that it is not captured★ |

★We crushed plan F's 2 pitfalls beforehand★

- **Port number** ―― Avoid the ports Chromium blocks (1/7/9/11/13/…/10080).
  With a blocked port the request may stop before being issued. 59999 is not blocked
- **Mixed content** ―― By spec `127.0.0.1` and `localhost` are
  "potentially trustworthy origins", so calling http from an https page is
  not treated as mixed content. ★Using a name such as `.invalid` gets stuck here★

★**The heart of N-1 was captured**★

```
1  188  POST  DOCUMENT  http://127.0.0.1:59999/echo      ← page navigation by form POST
1  421  POST  DOCUMENT  https://httpbingo.org/post       ← same (external)
1  188  GET   DOCUMENT  http://127.0.0.1:59999/n1-doc-get
1  203  GET   DOCUMENT  https://httpbingo.org/get?k=doc
```

**A path that wrapping in-page JS can never see** (navigation by form POST) was captured in
both systems. The reason we chose method B paid off here.

The remaining 3 FAILs were all fetch checks. Looking at the drain of (4):

```
8   500  GET  XHR  https://httpbingo.org/get?n=3
9   562  GET  XHR  http://127.0.0.1:59999/n1?n=4
10  609  GET  XHR  https://httpbingo.org/get?n=4
```

**All 10 requests fired with `fetch()` arrive with the `XHR` type.**
The test failed only because it presumed "if it is fetch, the type should be FETCH", and
★they were captured after all★. This is spec fact 70.

#### 3rd try (N-1c): 35 checks, OK 35 / FAIL 0

We fixed only the test side (the implementation was untouched).

1. **Stopped presuming the type** ―― the judgment is made on "was it captured".
   Linked with this, the false reports of `localOk` / `netOk` ("not a single one arrived") also vanished
2. ★**Added a check that counts "fetch arrives with the XHR type"**★ ―― we turn the spec fact we found
   directly into a regression test. We prepared a tool `N1CtxOf` that writes out the observed type, and
   ★made sure that while crushing a presumption we did not bring in a new presumption★
3. **Inserted one drain in (3b)** ―― in the 2nd try, at the point where (4) dropped the capacity to 3,
   the record of (2) vanished, and ★we ended up filling in by guesswork that "seq 1/2 should have been f-local and f-net"★.
   Always output before it disappears

```
Types that arrived: f-local=XHR  x-local=XHR  f-net=XHR  x-net=XHR
--- 4 network requests (4 total fires) ---
  1  3437  GET  XHR  http://127.0.0.1:59999/n1?k=f-local
  2  3500  GET  XHR  https://httpbingo.org/get?k=f-net
  3  3547  GET  XHR  http://127.0.0.1:59999/n1?k=x-local
  4  3609  GET  XHR  https://httpbingo.org/get?k=x-net
```

With seq 1 / 2 actually existing, the guess from the 2nd try is settled as a record.

---

### Fixed along the way ― ★a line where the numbers look like a lie★ (issue 3, plan Y)

The end of the 2nd try's log looked like this.

```
Wv2Pane.NetCaptureStop: capture ended (total fires 0 / on hand 0 / overflow 0)
```

**In fact 10 had come.** This was because `NetLogDrain` reset the cumulative counters to
0 along with the contents, so a person looking only at this line would read it as "not a single one came".
`NetLogTotal` is described as "total fires since capture start", so the description and the implementation
disagreed.

- `NetLogDrain` … empties only the ring's contents. ★Keeps the cumulative figures★
- `NetLogClear` … resets everything, including the cumulative figures
- `NetCaptureStart` … goes through `NetLogClear`, so it counts from 0 every time

We also added a check to prevent recurrence (`★Total fires remain even after draining (cumulative)★`).

---
### Design principles established (116 / 117)

- **Design principle 116 (N-1) A self-contained verification can say nothing when the closed world itself is the cause.
  ★Always mix in one control★.**
  In N-1's 1st try, we prioritized "don't rely on a real site" and put every target on the virtual host,
  and as a result, when everything failed, we could not
  distinguish "the wiring is bad" from "the virtual host is special". ★The more closed a verification, the more it needs a control that is not closed★.
  A relative of design principle 103 (cross-check the probe itself).

- **Design principle 117 (N-1) Do not presume the classification the other side returns, in your own words.
  ★Write the observed values to the log and then judge★.**
  We produced 3 FAILs from the presumption that since we fired it with `fetch()`, the type should be `FETCH`. In reality it arrives as `XHR`.
  When using a ★vocabulary the other side decides★ — classifications, enumerations, state names — as a condition, first output the measured values and confirm.
  The vocabulary version of design principle 112 (do not use indirect indicators).

### Spec facts established (69 / 70 / 71)

- **Spec Fact 69** ★Requests served via `SetVirtualHostNameToFolderMapping` do not
  ride on `WebResourceRequested`★ (all of DOCUMENT / fetch / XHR / images / CSS).
  Because the runtime returns them directly through an internal asset handler. ★Serving itself works without problems.
  Only the event does not ride★, so the verification page can stay on the virtual host, and
  **only the target needs to be moved outside**.

- **Spec Fact 70** ★A JS `fetch()` request arrives not as `FETCH`(8) but as
  `XML_HTTP_REQUEST`(7)★. On the real machine we fired 10 fetches and all 10 were XHR.
  Do not presume when narrowing by type. The default filter
  also keeps `FETCH`(8) registered (it does no harm, and would pick it up if this changes in the future).

- **Spec Fact 71** ★`http://127.0.0.1:<port>` is not stopped by mixed content even when called from an https page★
  (by spec a "potentially trustworthy origin").
  However ★avoid ports Chromium blocks (1/7/9/11/13/…/10080, etc.)★.
  Given these two, ★an unreachable local address can be used as a "target that generates
  a network request with zero external dependencies"★ (the connection fails but the request is sent).

---

### What became possible with N-1

★**After operating once by hand, you can see what is hit in what order.**★

```
Test_N1_Watch          ' Start capturing the traffic of the currently open tab
Test_N1_Watch "ALL"    ' When you want to see even images and CSS
' ... operate the page by hand (log in, search, have it spit out a CSV...)
Test_N1_Drain          ' Flush what has accumulated to the log and empty it
Test_N1_Stop           ' Done (the filters are also removed)
```

The regression test is `Test_N1_Capture` (35 checks), the isolation is `Test_N1_Site`, and the procedure is `Test_N1_Help`.

### What to do next (-> consumed in N-2)

- ~~**N-2** The request's **headers and body**~~ -> ★Complete (the N-2 stage below)★
- Open item: ★The `FETCH`(8) filter may not actually be used★
  (spec fact 70). Whether fetch is captured without registering it is unconfirmed. It does no harm so we leave it as is for now

---
## Stage N-2 ― Request headers and body (complete)

On top of N-1's list (URL / method / type), we add ★what is needed to reproduce a single request★.

### Issues and decisions (9 items agreed in one go)

| # | Issue | Decision |
|---|---|---|
| 1 | When to read the body | ★Plan A: read up to the limit inside the handler★ |
| 2 | How to restore it after reading | ★Plan X: return to the start with `Seek(0)`★ |
| 3 | `Cookie` / `Authorization` | ★Plan M: masked by default (can be explicitly unmasked)★ |
| 4 | Representation of the body | Plan P: look at `Content-Type`; if text, UTF-8; otherwise byte count + leading hex |
| 5 | How records are held | ★Plan S: details in a separate ring (default 50 entries). The list stays at 500★ |
| 6 | Retrieving details | Plan V: the list stays one line each; `NetDetail(i)` outputs just one |
| 7 | Which requests to read the body for | Plan alpha: by default only non-GET (+ an entry point to narrow by URL is also provided) |
| 8 | Default for detail capture | ★Plan kappa: details entirely OFF by default★ |
| 9 | How to confirm "it is not broken" | ★Plan D': compare against the JSON httpbingo returns★ |

★Issues 1 and 2 are the substance of N-2 itself★ ―― the rest are etiquette of operation.

**Reason for issue 1** `args` is valid only during Invoke, so ★the body can only be read
synchronously inside the handler★. Plan B (postpone with `GetDeferral`) keeps that request blocked
while the Deferral is held, so we would return to VBA's event loop with the communication stopped, which is
more dangerous than plan A. Plan C (length only) does not reach N-5.

**Reason issue 8 is kappa** ★N-1's list is complete as a "light tool that can always be used"★.
Adding the header iterator (about 60 COM calls per request) and body reading
by default would ruin that goodness. Details are needed only after we know "I want to reproduce this one",
so opt-in is natural.

**Reason issue 3 is M** In each verification, the operation is ★to have an AI read the log file★. If `Cookie` goes in
plain text, the session is handed over the moment the log is handed over. The default is the safe side; unmasking is explicit.

---
### ★The heart of N-2 ― something meant only to look can break the communication★

The `IStream` returned by `get_Content` **has a position**. Reading advances the position, and if
WebView2 then sends from the same stream, ★an empty body goes out★. Whereas N-1 "recorded only and returned immediately",
never touching the traffic, **N-2 in principle puts its hands on the traffic**.

The countermeasure is to return with `Seek(0, STREAM_SEEK_SET)` (plan X). Plan Y (`Clone`) is not implemented by some
streams, and plan Z (re-insert with `put_Content`) would mean implementing our own `IStream` in COM, which would make N-2 a different scale of work.

★But this cannot be settled with "it's probably fine"★ ―― looking only at our own side, we cannot tell that we failed to
restore it. **The only way is to compare against the body the destination received** (issue 9, plan D').

---
### Settled numbers (all from the SDK header)

| Target | index |
|---|---|
| `ICoreWebView2WebResourceRequest::get_Content` -> `IStream**` | **7** |
| `ICoreWebView2WebResourceRequest::get_Headers` | **9** |
| `ICoreWebView2HttpRequestHeaders::GetIterator` | **8** |
| `...HttpHeadersCollectionIterator::GetCurrentHeader(LPWSTR*, LPWSTR*)` | **3** |
| `...::get_HasCurrentHeader(BOOL*)` / `::MoveNext(BOOL*)` | **4** / **5** |
| `IStream::Read` / `Seek` / `Stat` | **3** / **5** / **12** |

★`IStream` is not in WebView2.h★ but **we could take it from the Windows SDK's `objidl.h`**
(`C:\Program Files (x86)\Windows Kits\10\Include\10.0.26100.0\um\objidl.h`).
`tools/vtable.py` can swap the header it reads with the environment variable `WEBVIEW2_H`, so
it can be used as is:

```bash
WEBVIEW2_H=".../um/objidl.h" python tools/vtable.py IStream
```

★Do not guess even for standard COM interfaces★ (design principle 115). "`Read` is next after IUnknown so 3" was right,
but being right and having confirmed are different things.

---
### Implementation (3 replacements, no new modules)

| File | Contents |
|---|---|
| `Wv2Thunks.bas` | Declaration of `MultiByteToWideChar` + `Utf8BytesToString` |
| `Wv2Pane.cls` | Detail ring + header / body reading + masking (+760 lines) |
| `Wv2Tests.bas` | `Test_N2_Detail` / `Test_N2_Help` (+400 lines) |

#### User-facing entry points

```vba
p.NetDetailOn = True       ' ★Details are OFF by default★
Test_N1_Watch              ' Start capture (unchanged from N-1)
' ... operate by hand ...
Test_N1_Drain              ' First get a lead from the list
p.NetDetailAll             ' Output all details
p.NetDetail 3              ' Output only the 3rd
```

| Member | Meaning |
|---|---|
| `NetDetailOn` | ★Default False★ Whether to capture details |
| `NetDetailCapacity` | Capacity of the detail ring (default 50) |
| `NetBodyMaxBytes` | Upper limit of the body (default 65536) |
| `NetRedact` | ★Default True★ Whether to mask secrets |
| `AddRedactHeader(name)` / `ClearRedactHeaders()` | Targets to mask |
| `NetDetailUriFilter` | Substring match of URLs for which to capture details |
| `NetDetailCount` / `NetDetailDropped` | Count on hand / ★count discarded by overflow★ |
| `NetDetailLine(i)` | seq / method / URL / byte count / truncated / Content-Type |
| `NetDetailHeaders(i)` / `NetDetailBody(i)` | Headers (masking applied) / body text |
| `NetDetail(i)` / `NetDetailAll()` / `NetDetailClear()` | Output / output all / discard |

#### ★Secrets are masked "when output"★

They are held **raw** on the ring, and masked at the moment `NetDetailHeaders` / `NetDetail` return/output them.
If we masked at storage time they could not be viewed later, but in this form
they can be unmasked on the spot with `NetRedact = False`. The property that **plain text goes to disk
only when "output" is written** is preserved.

The default list is `cookie` / `set-cookie` / `authorization` / `proxy-authorization` /
`x-api-key` / `x-auth-token` / `x-csrf-token` / `api-key`.

#### Used `MultiByteToWideChar` for UTF-8 decoding

★`StrConv(b, vbUnicode)` will not do★ ―― it interprets using the project's code page
(CP932 in a Japanese environment), so passing a UTF-8 byte sequence garbles it.
`ADODB.Stream` is an external COM so it cannot be used in product code (used only in verification code).
`MultiByteToWideChar` is a Win32 API, so it fits in the same category as the already-used `lstrlenW` /
`RtlMoveMemory`.

---
### Cross-checks before implementation

| Cross-check | Result |
|---|---|
| Call `MultiByteToWideChar` for real via ctypes | **14 OK** |
| Mirror masking / text detection / header extraction in Python | **24 OK** |

★We verified the very foundation we depend on★ (design principle 103). `Utf8BytesToString` merely calls
`MultiByteToWideChar` in 2 steps (get length -> convert), so calling it in the same order from ctypes
reveals the real behavior. The gain was confirming the behavior ★when truncated at the limit in the middle of a multibyte character★:

```
First 16 bytes -> 'abc=1&x=日本[U+FFFD]'   <- the part before is readable
First 17 bytes -> 'abc=1&x=日本語'
```

We also confirmed that adding `MB_ERR_INVALID_CHARS` makes **the whole conversion fail**, resulting in "the truncated body cannot be
read at all". So `dwFlags = 0` is right.

★Only the Edge headless DOM verification was skipped★ (Edge was running and handed off, spec fact 55).
What we wanted to confirm was "the Content-Type when passing a string body to fetch", and
the fetch spec fixes it as `text/plain;charset=UTF-8`, and that this value is
judged as text had been confirmed on the Python side. On the real machine it was as expected too.

---
### Real-machine test ― ★38 checks, OK 38 / FAIL 0 (passed on the first try)★

★The heart★:

```
1  POST  https://httpbingo.org/post  133  False  application/x-www-form-urlencoded
Body = [probe=wv2-n1-net&nihongo=%E6%97%A5%E6%9C%AC%E8%AA%9E...]
[OK] ★★The body reached the destination undamaged★★
[OK]   ★Japanese also reached undamaged★
```

**`Seek(0)` works. Reading does not break the communication.**
`Seek`'s by-value `LARGE_INTEGER` also went through with `dcf`'s `vbLongLong`.

Others:

```
(2) 1  POST  http://127.0.0.1:59999/n2b    36  False  text/plain;charset=UTF-8
    Body = [probe=wv2-n2-local&nihongo=日本語]
(3) 1  POST  http://127.0.0.1:59999/n2big 100  True   text/plain;charset=UTF-8
(4) User-Agent: <125 characters masked>
```

★Two discoveries from the log★

- **Japanese in a form POST is carried percent-encoded** (`%E6%97%A5%E6%9C%AC…`).
  It is per the `x-www-form-urlencoded` spec, but ★it matters when N-5 emits PowerShell★
- `sec-ch-ua` / `sec-ch-ua-mobile` / `sec-ch-ua-platform` (Client Hints) are also visible.
  The headers needed for reproduction are all there
- The `Origin` of requests that originate from a virtual-host page becomes `https://appassets.netprobe`

---
### Fixed along the way ― ★an inconsistency with a principle we found ourselves★

Only the detail ring silently overwrote on overflow. The list has `NetLogDropped` but the detail did not, which is
an inconsistency, so we added `NetDetailDropped`, made `NetDetailAll` warn as well,
and added one check. **Because we noticed it ourselves before taking it to the real machine, it was cheap.**

---
### Design principles established (118 / 119)

- **Design principle 118 (N-2) When an observation touches its target, count "that we did not break it"
  ★by a separate path★.**
  Reading the `IStream` advances the position = the observation can change the communication. We restored it with `Seek(0)`, but
  ★evidence that it works can never be obtained from our own side★. It is known only by comparing against the
  body the destination received. One level above design principle 112 (do not use indirect indicators),
  ★this is the etiquette for scenes where "our own observation result itself cannot be trusted"★.

- **Design principle 119 (N-2) Mask secrets not "when storing" but "when outputting".**
  Hold them raw on the ring, and mask at the moment they become a log or a return value. Masking at storage time means
  they cannot be viewed later and no entry point for unmasking can be made. In this form,
  while preserving the property that ★plain text goes to disk only when "output" is written★,
  it can be unmasked with `NetRedact = False` when needed.

### Spec facts established (72 to 75)

- **Spec Fact 72** ★The `IStream` from `get_Content` advances its position when read, but
  returning with `Seek(0, STREAM_SEEK_SET)` does not break the communication★ (measured. The body reached the destination).

- **Spec Fact 73** If you pass a **string** body to JS `fetch()`, the Content-Type becomes
  `text/plain;charset=UTF-8`. ★Since it is a CORS "simple request", no preflight is
  interposed★, so even for an unreachable target the request itself is sent (usable in verification).

- **Spec Fact 74** An HTML form POST is `application/x-www-form-urlencoded`, and
  ★Japanese is carried percent-encoded★ (`日本語` -> `%E6%97%A5%E6%9C%AC%E8%AA%9E`).

- **Spec Fact 75** `MultiByteToWideChar(CP_UTF8, 0, ...)` turns a multibyte character cut in the middle
  into U+FFFD and **returns the rest**. ★Adding `MB_ERR_INVALID_CHARS` makes the whole conversion fail★,
  so it must not be added for uses that truncate at a limit.

---
### What became possible with N-2

★**Everything needed to reproduce this one request is now in place.**★ What remains is the response side and the exit.

### What to do next (-> consumed in N-3)

- ~~**N-3** The response's **status and headers**~~ -> ★Complete (the N-3 stage below)★
- Open item: ★`multipart/form-data` (file upload) is treated as binary★.
  Whether to read the contents delimited by the boundary string depends on what N-5 needs
- Open item: the `FETCH`(8) filter may not actually be used (spec fact 70. Carried over from N-1)

---
## Stage N-3 ― Response status and headers (complete)

A ★result★ was attached to each line of the list. ★Response headers★ were attached to the detail.

```
#   Elapsed ms  Method   Type       Status  URL
1       94  GET      XHR          200 https://httpbingo.org/get?k=n3-off
```

### Issues and decisions (8 items agreed in one go)

| # | Issue | Decision |
|---|---|---|
| 1 | Which responses to record | ★Plan A: only those corresponding to requests already in the list★ (+ an entry point to narrow by URL) |
| 2 | Key for matching with requests | Plan P: attach to the not-yet-matched request by URI + method |
| 3 | Where to write | ★Plan S: add the status to the list's line★ |
| 4 | Response headers | Plan U: add to the detail entry (only when `NetDetailOn`) |
| 5 | Responses arriving after overflow | Plan W: discard and ★count★ |
| 6 | Redirects | Plan Y: 1 request 1 response (if already filled, go to the next unmatched) |
| 7 | Responses that have not arrived yet | ★Plan alpha: leave blank★ (`0` cannot be told apart from a real one) |
| 8 | Verification | Both (httpbingo + the unreachable local one as a control for "does not come") |

---
### ★The central problem of N-3 ― this event has no filter★

`AddWebResourceResponseReceivedFilter` is ★absent★ from the vtable of `ICoreWebView2_2`
(after 61/62 comes `NavigateWithWebResourceRequest`). That is,
**images, CSS, and fonts all come**. Moreover args holds only `Request` and
`Response`, and ★the type (`COREWEBVIEW2_WEB_RESOURCE_CONTEXT`) is not included★,
so we cannot narrow by type either.

So we ★record only those corresponding to requests already in the list (issue 1, plan A)★.
N-1's net already defines "what we want to see", so we ride on it. If we want everything,
widening the list side with `NetCaptureStart "ALL"` brings the responses along automatically.
★Do not hold the narrowing in two places★

On the other hand, ★N-3 does not touch the traffic★ (the body `GetContent` is asynchronous = N-4).
There is none of N-2's "reading breaks it" risk.

---
### Settled numbers (all from the SDK header)

| Target | index |
|---|---|
| `ICoreWebView2_2::add_WebResourceResponseReceived` | **61** (★call it on `m_pView2`★) |
| Same `remove_` | **62** |
| `...ResponseReceivedEventArgs::get_Request` / `get_Response` | **3** / **4** |
| `ICoreWebView2WebResourceResponseView::get_Headers` | **3** |
| Same `get_StatusCode`(int*) / `get_ReasonPhrase` | **4** / **5** |
| Same `GetContent` (★asynchronous, needs a completion handler★) | **6** (N-4) |
| `ICoreWebView2HttpResponseHeaders::GetIterator` | ★**7**★ |

The IID is `{7de9898a-24f5-40c3-a2de-d4f458e69828}`, and `Invoke(sender, args)` has 2 arguments.

#### ★A trap we nearly stepped on ― GetIterator's index differs between request and response sides★

| Interface | GetIterator |
|---|---|
| `ICoreWebView2HttpRequestHeaders` | **8** |
| `ICoreWebView2HttpResponseHeaders` | ★**7**★ |

The response side has no `put_` family and `AppendHeader` comes first, so **the order is shifted**.
Had we reused N-2's code on the assumption that "it must be the same", ★we would have called `GetHeaders`
off by one★ (and since `GetHeaders` takes different arguments, it could crash).
A living example of design principle 115.

The iterator's type itself is `ICoreWebView2HttpHeadersCollectionIterator` on both
(3 = `GetCurrentHeader` / 4 = `get_HasCurrentHeader` / 5 = `MoveNext`), so
we split the reading out into `NetReadHeadersIter(pHeaders, iterIndex)` and made it
★call with the index as an argument★.

---
### Implementation (4 replacements, no new modules)

| File | Contents |
|---|---|
| `Wv2Thunks.bas` | `HK_WebResourceResponseReceived = 12` + IID |
| `ComCallbackHandler.cls` | Added `Case` -> `View2_OnWebResourceResponseReceived` |
| `Wv2Pane.cls` | Status column in the list + response headers + matching |
| `Wv2Tests.bas` | `Test_N3_Status` / `Test_N3_Help` |

#### Added Public members

| Member | Meaning |
|---|---|
| `NetLogStatus(i)` | The status of that row (★0 = not arrived yet★) |
| `NetDetailRespHeaders(i)` | Response headers (returned with masking applied) |
| `NetRespOn` | Whether the response event is being picked up |
| `NetRespUnmatched` | ★Number of responses that had no counterpart in the list★ |

`NetLogLine` went from 5 items to 7 items, but ★the status was added at the end★, so
the existing subscripts (0 to 4) did not move (the N-1 / N-2 tests pass unmodified).
In the human-readable `NetLogFormat`, it is placed between the type and the URL, as in F12.

---
### Cross-check before implementation ― ★one error in the agreed approach was found★

We ran the central matching algorithm in Python (18 OK).

In issue 2 we had proposed "attach to the **newest** unmatched request" and received approval, but
**running it showed it was wrong**:

```
The same URL is sent twice and then the responses return (Req1 -> Req2 -> Resp(200) -> Resp(404))
    From newest -> [(1, 404), (2, 200)]   <- ★swapped★
    From oldest -> [(1, 200), (2, 404)]   <- correct
```

In scenes with one round trip at a time in order, **both give the same result**, so
★the oldest (FIFO) is always equal or better★. We changed to oldest before implementing.

★Even something agreed on can be fixed if run before implementation★ ―― had we noticed only after taking it to the real machine,
it would have shown up in the hardest-to-decipher form: "for some reason the order is swapped".

In addition, we actually hit `httpbingo`'s endpoints (`/status/404` / `/redirect/2` /
`/response-headers` / `/cookies/set`) and confirmed the behavior before running.

---
### Real-machine test ― ★it took 4 tries. But the implementation was never wrong★

| Try | Result | Cause |
|---|---|---|
| 1 | 22 checks, FAIL 2 | Response headers empty |
| 2 | Same (instrumentation added) | ★The reach point was revealed: `[response] not arrived yet`★ |
| 3 | 22 checks, OK 22 | ★The test's way of waiting was wrong★ (the implementation is innocent) |
| 4 | ★25 checks, OK 25★ | Filled in the checks that had not run |

#### 1st try ― the central part passed in one go

```
1    78  GET  XHR  https://httpbingo.org/redirect/2?k=n3-redir      302
2   609  GET  XHR  https://httpbingo.org/relative-redirect/1        302
3  1156  GET  XHR  https://httpbingo.org/get                        200
[OK] ★Every row has exactly one response attached★
```

★The 3-hop redirect lined up cleanly as 1 request 1 response★ ―― the matching (FIFO) works
in the field. 200 / 404 / 500 were also attached, and the unreachable target stayed blank.

The FAILs were only the 2 response-header checks.

#### 2nd try ― ★we added instrumentation and isolated it in one round★

With only "it was empty", we could not tell where it became empty. Since we are inside the handler,
we cannot fire logs, so we added instrumentation that ★writes the reason into the return value string itself★
(the shape of design principle 111). We also had the test call `NetDetail` so the reach point would be visible.

The result was ★`[response] not arrived yet`★ ―― it had not reached `NetAttachRespHeaders`.
That is, **the header reading is innocent**. Not one of the diagnostic strings we prepared got its turn.

#### 3rd try ― ★the cause was my test's way of waiting★

The timestamps were decisive:

```
23:11:08.694  (4) start
23:11:09.035  details obtained     <- 0.34 seconds later
23:11:09.037  response headers: empty
```

The round trip to httpbingo took 0.94 seconds in (1) and 6 seconds in (3). In 0.34 seconds the
response had not returned. (4) was waiting with `N2Wait` (until **the detail appears**).
But ★the detail is created the moment the request is sent★, so this means nothing at all about the response having
"arrived". In (1) we correctly used `N3Wait` (until **the status is attached**),
but in (4) we mixed them up.

When we added `N3WaitAny` (until a response is attached regardless of status) and fixed it,
the response headers were read in one go:

```
[Response headers: 10]
  access-control-allow-origin: https://appassets.netprobe
  content-type: application/json; charset=utf-8
  x-n3-test: hello          <- ★the header we specified came back★
```

★`GetIterator = 7` was right, and the implementation was correct from the start.★

#### 4th try ― ★even when OK, a check that "did not run" is the same as a vacuous pass★

The 3rd try was "OK 22 / FAIL 0", but ★not one `Set-Cookie` check
had run★. The response never came and the whole `If` block was skipped, and
that did not appear anywhere in the numbers.

- Instead of relying on an external `Set-Cookie`, we changed it so that ★we add the already-readable `content-type` to the masked targets
  and compare reading the same one entry before and after masking★
- We treated `Set-Cookie` as an add-on and ★issue a warning if it did not run★

```
[OK] ★Masking also works on response headers★
[OK] ★The contents (application/json) are gone★
[OK] ★Raw on the ring (comes back with NetRedact = False)★
Note: ★The Set-Cookie add-on check did not run (no response came)★
```

The third one is what matters ―― ★if we had masked at storage time, it would not
come back even with `NetRedact = False`★, so this settles that "we mask at output time" (design principle 119).

---
### Design principles established (120 / 121)

- **Design principle 120 (N-3) Do not wait for "ready" in place of "done".
  ★Make the condition you wait for the very event you want to see★.**
  The detail entry is created ★the moment the request is sent★, and the response comes after.
  We were reading the response headers after waiting "until the detail appears", so every time we were reading
  something that had not arrived yet. ★If you substitute a nearby, easier-to-observe event,
  the failure appears in the hardest-to-decipher form, "the value is empty"★.
  The time version of design principle 112 (do not use indirect indicators).

- **Design principle 121 (N-3) ★A check that did not run is more dangerous than a FAIL★.
  A check skipped by a condition vanishes from the numbers as "OK 0 / FAIL 0".
  Always put the fact that it was skipped in the log.**
  N-3's 3rd try was "OK 22 / FAIL 0", but the `Set-Cookie` we wanted to confirm
  had not run even once. ★The worst kind is one that looks like a pass while confirming nothing★.
  The same root as design principle 116 (mix in a control) / N-1's 4 vacuous passes.

### Spec facts established (76 to 78)

- **Spec Fact 76** ★`ICoreWebView2HttpResponseHeaders::GetIterator` is 7, and
  `ICoreWebView2HttpRequestHeaders` is 8★. The response side has no `put_` family and
  `AppendHeader` comes first, so the order is shifted. Reusing code is off by one.

- **Spec Fact 77** ★`WebResourceResponseReceived` has no filter★
  (`AddWebResourceResponseReceivedFilter` is absent from the vtable).
  Images and CSS all come. Moreover args does not include the type, so we cannot narrow by type either.

- **Spec Fact 78** HTTP/2 responses have no reason phrase, so
  `get_ReasonPhrase` returns an empty string (only the status code comes).
  An empty part after `200` is not abnormal.

---
### What became possible with N-3

★**"When, what, how it was hit, and what happened" is complete in a single list line.**★
What remains is the response body (N-4) and the exit (N-5).

### What to do next (-> consumed in N-4)

- ~~**N-4** The response **body**~~ -> ★Complete (the N-4 stage below)★
- Open item: ★A cross-origin request of `fetch(..., {redirect:'manual'})` seems not to ride on
  `WebResourceResponseReceived`★ (in N-3 the response did not come either time).
  ★This is an observation, not an assertion★. Investigate if it becomes necessary in N-4 or later.
- Open item: the diagnostic instrumentation (the path that returns `(diagnostic: ...)`) is left in place. It does no harm, but
  fold it away if it gets in the way in N-4 or later.

---
## Stage N-4 ― Response body (complete)

★The construction differs from N-1 to N-3. It becomes asynchronous for the first time.★
`GetContent` returns immediately after being handed a handler, and the body arrives later.

### Issues and decisions (8 items agreed in one go)

| # | Issue | Decision |
|---|---|---|
| 1 | Switch for capturing the body | ★Plan A: newly add `NetRespBodyOn` (default False)★ |
| 2 | If compressed | ★Plan P': first build it in an observable form★ (always keep the encoding and leading hex) |
| 3 | Limit | Plan S: hold a separate `NetRespBodyMaxBytes` (default 64KB) |
| 4 | Waiting for completion | Plan X: put the seq in `CallbackId` (reuse the mechanism of 9.8c) |
| 5 | Lifetime of `ResponseView` | ★Plan beta: hold with AddRef until completion★ |
| 6 | If it disappears before completion | Plan W: discard and count (`NetRespBodyOrphan`) |
| 7 | How to output the body | Plan U: `NetDetailRespBody(i)` + add to `NetDetail(i)` |
| 8 | Verification | Both (httpbingo's `/json` `/gzip` `/html` `/image/png` `/status/204`) |

---
### ★N-4's 3 worries★

**① Compression** ― In N-3's log, `content-encoding: zstd` was attached to a response.
Whether what `GetContent` returns is already decoded is ★not stated in the SDK header★.
If it stays compressed, decoding zstd / br in VBA is out of the question.

**② Volume** ― Response bodies are an order of magnitude different from request bodies (HTML hundreds of KB, images MB).
Moreover, by spec fact 77 there is no filter, so capturing unconditionally would be a disaster.

**③ Lifetime of `ResponseView`** ― The header does not state "may we release it after calling `GetContent`".
If we get it wrong it appears as ★access to a released object
= crash★, and by spec fact 20 the debugger cannot be used either.

On the other hand, ★N-2's worry is gone★ ―― this only reads a copy of the response and does not
put hands on the traffic, so `Seek(0)` is not needed (design principle 118 has no role).

---
### Settled numbers

| Target | index / value |
|---|---|
| `ICoreWebView2WebResourceResponseView::GetContent(handler)` | **6** |
| Completion handler IID | `{875738e1-9fa2-40e3-8b74-2e8972dd6fe7}` |
| `Invoke(HRESULT errorCode, IStream* result)` | **3** |

★The shape of `Invoke` is exactly the same as `HK_ExecuteScriptCompleted`★, so
the `CallbackId` mechanism built in 9.8c could be used as is. We put
the request's serial number in `CallbackId` to tell which response's body it is.

---
### Implementation (4 replacements, no new modules)

```vba
p.NetDetailOn = True      ' ★The body's home is the detail entry★
p.NetRespBodyOn = True    ' ★The body has a separate switch (default OFF)★
' ... operate ...
p.NetDetail 3             ' request + response + body
p.NetDetailRespLine(3)    ' byte count/truncated/kind/encoding/hex/arrived
```

| Member | Meaning |
|---|---|
| `NetRespBodyOn` | ★Default False★ Whether to capture the body |
| `NetRespBodyMaxBytes` | Upper limit of the body (default 65536) |
| `NetRespBodyOrphan` | Number of bodies that had no home |
| `NetDetailRespBody(i)` | The body (contents only when text) |
| `NetDetailRespLine(i)` | Byte count / truncated / kind / encoding / leading hex / ★whether it arrived★ |

#### ★Why we hold "arrived" separately★

`204 No Content` is ★0 bytes but has arrived★. An unreachable target
★never arrives in the first place★. If we lump these two together as "0 bytes",
the wait (`N4Wait`) waits forever on a 204. So we hold `m_detRespGot` separately
(design principle 111 ― do not leave behind a value that cannot be judged).

#### Reading was consolidated into one

Request body (N-2) and response body (N-4) are read the same way, and
★the only difference is whether we return with `Seek(0)`★. So we split it out into `NetReadStream`
and made returning the responsibility of `NetReadBody` (request side).

---
### Cross-checks before implementation ― ★34 static checks★

We crushed by machine the failures that melted away in N-3:

| What we looked at | Why |
|---|---|
| ★`N2Wait` is not used for the body check★ | In N-3 we mixed up what to wait for and lost 3 real-machine rounds |
| The order of the 6 items of `NetDetailRespLine` and the meaning of the subscripts the test reads | If the subscripts shift, it shows up as "the value is empty" and cannot be deciphered |
| `NetReadBody` does `Seek(0)` / `NetReadStream` does not | If they get mixed, N-2's communication breakage recurs |
| `AddRef` -> hold -> `Release` at completion are paired | A leak means a crash or a leakage |
| The compression check is decided as "exactly one of the two" with `Xor` | Don't miss "neither" |

In addition, we actually hit httpbingo's `/gzip` and confirmed that ★`"gzipped": true` is in
the contents★ before running (this becomes the foundation of the check).

---
### Real-machine test ― ★24 checks, OK 24 / FAIL 0 (passed on the first try)★

#### ★The biggest unknown was settled★

```
Profile = 1395  False  True  gzip  7B 0A 20 20 22 61 72 67 73 22 ...  True
-> ★GetContent returns the decoded body★
```

Although `content-encoding: gzip` was attached, the leading hex is `7B` (= `{`),
★plain-text JSON★. The `zstd` of `/json` `/html` was the same.

That is, ★`content-encoding` only indicates the encoding on the wire, and what `GetContent`
returns is after decoding★. **You must not judge "it is compressed" by looking at the header's encoding.**
This ties directly to emitting `Invoke-WebRequest` in N-5
(it too decompresses automatically, so there is no need to reproduce the encoding).

★Because we made it a check that "is always settled as one of the two" instead of presuming, it was settled in 1 round★
―― had we only observed, it would have ended with "it was gzip yet readable, why?".

#### Others

```
(4) 8090  False False       89 50 4E 47 0D 0A 1A 0A ...  True   <- PNG magic
(5)  100  True  True  zstd  7B 0A 20 20 ...              True   <- truncated at the limit
(6)    0  False False                                    True   <- 204: arrived
       0  False False                                    False  <- unreachable: does not arrive
(7)    0  False False                                    False  <- body OFF (headers are attached)
```

---
### Fixed along the way ― ★we were mixing "failure" and "no body"★

In the 1st try, (6) came out like this:

```
Profile = 0  False  False  (fetch failed hr=&H800700E8)  True
```

The check was OK, but ★the display was misleading★. `0x800700E8` =
`HRESULT_FROM_WIN32(ERROR_NO_DATA)`, which is **not a failure but "there is no body"**.
It would mislead us when judging "is there a body" in N-5 (design principle 111).

We named `E_NO_DATA`, split the branches, and added one check to confirm again
(★do not stack up unverified work★ = applying design principle 121 to ourselves).

---
### Design principle established (122)

- **Design principle 122 (N-4) When there can be two answers,
  write a check that ★is always settled as one of them★ (make "neither" a FAIL).**
  The next move after design principle 117 (do not presume the classification the other side returns). Merely not presuming
  ends in ★just observing★, and no pass/fail is given. In N-4 we made it the form
  "the body shows `gzipped` readable **Xor** it starts with `1F 8B`", so
  ★the spec was settled in one real-machine round, and it also remained as a regression test★.
  A pattern for turning the unknown not into "investigating" but into "a check".

### Spec facts established (79 / 80)

- **Spec Fact 79** ★`GetContent` returns the decoded body★
  Even with `content-encoding: gzip` / `zstd` attached, the contents are plain text.
  ★You must not judge "it is compressed" by looking at the header's encoding★.

- **Spec Fact 80** On `204 No Content`, the completion of `GetContent` comes with
  `0x800700E8` (`HRESULT_FROM_WIN32(ERROR_NO_DATA)`).
  ★This is not a failure but "there is no body"★.

---
### What became possible with N-4

★**A single communication, from request to response body, can be obtained in full.**★
Only the exit (N-5) remains.

### What to do next

- **N-5** ★**Emit as PowerShell `Invoke-WebRequest` code**★
  Only once we get here does it become value that F12 does not have. The materials in hand are all in place:
  URL / method / request headers / request body / status / response headers / response body.
  - ★Spec fact 74★ Japanese in a form POST is carried percent-encoded.
    To make code that can be pasted as is, we must decide whether to decode it or pass it raw
  - ★Spec fact 79★ Responses can be obtained decoded, so there is no need to
    reproduce `content-encoding` (`Invoke-WebRequest` also decompresses automatically)
  - ★Handling of secrets★ We built the masking mechanism for Cookie / Authorization in N-2, but
    ★in code generation they must go into "code that works"★. The issue becomes deciding whether to emit them masked,
    to use placeholders, or to make the user explicitly unmask
