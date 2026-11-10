VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Wv2Element"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --- Wv2Element.cls  Stage D-3 (writing and operations) ---
'
'   D-3 adds 4 public members:
'
'     Value = v                  Property Let  writes the value (**has no return value**)
'     Click()             As Boolean           clicks
'     SetAttribute(n, v)  As Boolean           writes an attribute
'     LastInfo            As String            name of the route chosen by the last operation
'
'   # How to check success (follows the D-2 LastOk / LastError convention as is, issue 8)
'     Click / SetAttribute ... the return value (Boolean) is enough. Use LastError when you need the reason.
'     Value =              ... **a Property Let cannot return a value**, so
'                            **always check LastOk** to see whether the write went through.
'
'       el.Value = "abc"
'       If Not el.LastOk Then Debug.Print el.LastError
'
'     The values placed in LastError are the same as in D-2 (stale / no-pane / in-callback / timeout(n s) /
'     the JS exception message).
'
'   # **Success means "the JS ran without an exception", not "the intended effect happened"**
'     Clicking a disabled button raises no exception, so True is returned. Writing to a readonly
'     input is the same. When the effect matters, **read it back to confirm**.
'
'   # What Value = does (issue 6: **the core of SPA support**)
'     With a plain e.value = x, frameworks such as React / Vue **do not notice**
'     the value change. The framework overlays its own value setter (value tracker) on the
'     element instance, and an assignment that goes through it is recorded as
'     "a value I already know".
'
'     So D-3 fires in this order:
'       (1) take the **native setter** of value from e.constructor.prototype and,
'           if there is one, call it on e (= bypassing the tracker). LastInfo = 'setter'
'           For an element where it cannot be obtained, fall back to a plain assignment.  LastInfo = 'direct'
'       (2) fire the input event (React-style frameworks watch this)
'       (3) fire the change event (jQuery-style code watches this)
'     **Firing both** is the agreement of issue 6. On plain HTML one extra change fires, but we judged that
'     the accident of "some pages do not react" costs more than the actual harm.
'
'   # Shape of Click (issue 7, plan A)
'     If e.click() exists, use it (LastInfo = 'click'). Otherwise
'     dispatch a MouseEvent('click') with bubbles (LastInfo = 'dispatch').
'     **mousedown / mouseup are not fired**; we will think about it when drag-style UIs are needed.
'     Note that the fired events have isTrusted = false (they are synthetic events
'     made from JS, so this is inevitable). It does not work on pages that reject events by looking at isTrusted.
'
'   # Usage example
'     Dim p As Wv2Pane, el As Wv2Element
'     Set p = UserForm1.GetActivePane
'     Set el = p.QuerySelector("input[name='q']")
'     el.Value = "search term"
'     If el.LastOk Then p.QuerySelector("button[type='submit']").Click
'
'   **As of D-3 there is still no way to "wait until the change finishes"**
'     WaitFor / WaitGone, which wait for the DOM to be rewritten after a Click, will be added after
'     measuring the premise of issue 3 (Wv2Tests.Test_D3_Probe_Promise).
'''''''''''''''''''''''''''''''''''
' --- Wv2Element.cls  Stage D-2 (read operations on an element handle) ---
'
'   **The core class of the D axis (DOM / page automation). Newly added in D-2.**
'
'   # Role
'     A handle for reading, from the VBA side, one DOM element registered in the JS-side
'     registry (window.__wv2.els). It holds only these three things:
'
'       m_pane   ... which Wv2Pane (= which tab) the element belongs to
'       m_handle ... index into the JS-side window.__wv2.els[]
'       m_gen    ... registry generation at acquisition time (a string issued by JS)
'
'     **It does not hold the DOM object itself.** It asks JS each time a value is needed.
'     So reading one property costs one ExecuteScript round trip.
'     Reading a lot in a loop is slow, so to fetch many at once it is faster to
'     write one JS expression with Wv2Pane.EvalSync (a design trade-off accepted in D-2).
'
'   # How to create one
'     **Do not create with New.** It is returned by Wv2Pane.GetElementById / Wv2Pane.QuerySelector.
'     If you New it directly, m_pane is Nothing and nothing can be read.
'
'   # Public members (D-2 is **read-only**. Writing and Click are D-3)
'     TagName            As String   tag name of the element (upper case. e.g. INPUT)
'     InnerText          As String   the displayed text
'     InnerHTML          As String   inner HTML (< of spec fact 30 already decoded)
'     Value              As String   value of input / textarea / select. Empty string if none
'     GetAttribute(name) As String   attribute value. Empty string if the attribute is absent
'     IsStale            As Boolean  whether the handle has become invalid (one round trip)
'     Handle             As Long     JS-side index (for debugging)
'     Generation         As String   generation at acquisition (for debugging)
'     LastOk             As Boolean  whether the last read succeeded
'     LastError          As String   reason for the last failure
'
'   # How to read "an empty string came back" (**important**)
'     Read-type members return an empty string even on failure, so distinguish the meaning by LastOk:
'
'       LastOk = True   ... **normal** it really was empty (no text / no attribute)
'       LastOk = False  ... **failure** the reason is in LastError:
'                           stale        the registry was rebuilt (page navigation etc.)
'                           no-pane      no Pane reference (created with New)
'                           in-callback  called from inside an event handler (spec fact 14)
'                           timeout(n s) JS did not return (spec fact 32)
'                           other        EvalSync's LastEvalError is stored as is
'
'   # The idea of generation (Generation) (D-2 issue 2, plan b)
'     The generation is **issued by JS** (window.__wv2.gen). A page navigation wipes out
'     the whole JS execution environment, so on the next access the registry is rebuilt and the generation is always different.
'     The approach of counting navigation events on the VBA side (plan a) drifts from reality: with SPA pushState
'     the DOM remains yet an event fires, and conversely when in-page JS swaps elements
'     no event fires. So **the authority for the verdict is always on the JS side**.
'
'     The generation cache that Wv2Pane holds (CurrentDomGen) is used only as a fast path that
'     "skips the round trip only when it is already known to be stale". If the cache is old it
'     always falls toward asking JS, so no misjudgment occurs.
'
'   # Lifetime
'     **Class_Terminate does nothing.** It does not release the JS-side slot (D-2 issue 3).
'     We cannot control when and in what context Terminate runs, so firing EvalSync from there
'     easily steps on D-1's in-callback guard or on a Pane already closed by Ctrl_Close.
'     The registry vanishes entirely on page navigation, so the only real harm is that
'     "searching a lot within the same page makes the JS-side array grow". To discard manually,
'     call Wv2Pane.ClearElementRegistry. Proper lifetime management is handled in D-4.
''''''''''''''''''''''''''''''''''
Option Explicit


' --- Internal state ---
'   m_pane   : reference to the owning Wv2Pane (Wv2Pane does not hold Wv2Element, so
'              there is no circular reference)
'   m_handle : index into the JS-side window.__wv2.els[] (0-based)
'   m_gen    : registry generation at acquisition (a string issued by JS)
Private m_pane   As Wv2Pane
Private m_handle As Long
Private m_gen    As String

' --- Result of the last read (same style as D-1's LastEvalOk / LastEvalError) ---
Private m_lastOk    As Boolean
Private m_lastError As String

' --- Name of the route chosen by the last **operation** (D-3. 'setter' / 'direct' / 'click' etc.) ---
Private m_lastInfo  As String


Private Sub Class_Initialize()
    Set m_pane = Nothing
    m_handle = -1
    m_gen = ""
    m_lastOk = False
    m_lastError = ""
    m_lastInfo = ""
End Sub


Private Sub Class_Terminate()
    ' **Deliberately does nothing** (D-2 issue 3, plan a). See the header comment above.
    Set m_pane = Nothing
End Sub


' ============================================================
' InitElement (Friend, called only from Wv2Pane)
'   The only entry point through which Wv2Pane.GetElementById / QuerySelector
'   pour in the handle and generation received from JS.
' ============================================================
Friend Sub InitElement(ByVal ownerPane As Wv2Pane, _
                       ByVal elemHandle As Long, _
                       ByVal elemGen As String)
    Set m_pane = ownerPane
    m_handle = elemHandle
    m_gen = elemGen
    m_lastOk = True
    m_lastError = ""
    m_lastInfo = ""
End Sub


' ============================================================
' Handle / Generation (plain values for debugging)
' ============================================================
Public Property Get Handle() As Long
    Handle = m_handle
End Property

Public Property Get Generation() As String
    Generation = m_gen
End Property


' ============================================================
' LastOk / LastError (success or failure of the last read)
'
'   **Always use it to interpret an empty string result**
'     LastOk = True  -> it really was empty (normal)
'     LastOk = False -> failure. The reason is in LastError
' ============================================================
Public Property Get LastOk() As Boolean
    LastOk = m_lastOk
End Property

Public Property Get LastError() As String
    LastError = m_lastError
End Property


' ============================================================
' LastInfo (D-3) - name of the route chosen by the last **operation**
'
'   Value =      : 'setter' (via the native setter) / 'direct' (plain assignment)
'   Click        : 'click'  (e.click())             / 'dispatch' (synthetic event)
'   SetAttribute : 'ok'
'
'   **Calling a read-type member (TagName etc.) resets it to an empty string.** It is a record of "the last operation",
'   so to see the route of a write, read it **right after writing**.
'   It is for diagnostics and is not used to judge success (that is the job of LastOk / the return value).
' ============================================================
Public Property Get LastInfo() As String
    LastInfo = m_lastInfo
End Property


' ============================================================
' TagName / InnerText / InnerHTML / Value (the 4 read types of D-2)
'
'   Each costs one ExecuteScript round trip per call.
'
'   **InnerHTML always runs into spec fact 30 (< arrives as <)**
'     Wv2Json.JsonUnescape decodes it, so plain HTML is obtained on the VBA side.
'
'   **Value is an empty string for anything other than input / textarea / select**
'     Other elements have no value property, so it is undefined on the JS side.
'     In that case LastOk is True (it is not an error).
' ============================================================
Public Property Get TagName() As String
    TagName = ReadProp("e.tagName")
End Property

Public Property Get InnerText() As String
    InnerText = ReadProp("e.innerText")
End Property

Public Property Get InnerHTML() As String
    InnerHTML = ReadProp("e.innerHTML")
End Property

Public Property Get value() As String
    value = ReadProp("e.value")
End Property


' ============================================================
' GetAttribute (D-2)
'   Fetches an attribute value. If the attribute is absent: empty string + LastOk = True.
'
'   Usage example:
'     el.GetAttribute("href")
'     el.GetAttribute("data-id")
'
'   **Difference between getAttribute and .value**
'     After the user rewrites <input value="initial">,
'       .Value               -> the current value (after rewriting)
'       .GetAttribute("value") -> the attribute value (still the initial value)
'     Use them according to purpose.
' ============================================================
Public Function GetAttribute(ByVal attrName As String) As String
    If m_pane Is Nothing Then
        m_lastOk = False
        m_lastError = "no-pane"
        GetAttribute = ""
        Exit Function
    End If
    GetAttribute = ReadProp("e.getAttribute(" & m_pane.JsQuote(attrName) & ")")
End Function


' ============================================================
' Value (D-3, Property Let) - writes the value
'
'   **Being a Property Let it has no return value** Always check LastOk for success (issue 8).
'   **3 steps to make the framework notice** (see the header):
'     fire in the order native setter -> input event -> change event
'
'   Usage example:
'     el.Value = "search term"
'     If Not el.LastOk Then Debug.Print "write failed: " & el.LastError
'
'   **What happens when fired at something other than input / textarea / select**
'     A div etc. has no value property, so the native setter cannot be obtained and it falls back to
'     a plain assignment (LastInfo = 'direct'). In JS terms it just grows one custom property
'     and raises no exception, so **LastOk becomes True**.
'     Confirm "whether the effect appeared" by reading it back.
' ============================================================
Public Property Let value(ByVal newValue As String)
    Dim jsVal As String
    Dim js As String

    If m_pane Is Nothing Then
        m_lastOk = False
        m_lastError = "no-pane"
        m_lastInfo = ""
        Exit Property
    End If

    jsVal = m_pane.JsQuote(newValue)

    js = "var d=Object.getOwnPropertyDescriptor(e.constructor.prototype,'value');" & _
         "if(d&&d.set){d.set.call(e," & jsVal & ");rv='setter';}" & _
         "else{e.value=" & jsVal & ";rv='direct';}" & _
         "e.dispatchEvent(new Event('input',{bubbles:true}));" & _
         "e.dispatchEvent(new Event('change',{bubbles:true}));"

    ExecAction js, "Value"
End Property


' ============================================================
' Click (D-3) - clicks
'
'   Return value: True = ran on the JS side without an exception / False = failure (reason in LastError)
'
'   If e.click() exists it is used (LastInfo = 'click'); otherwise
'   a MouseEvent('click') is dispatched with bubbles + cancelable
'   (LastInfo = 'dispatch'). Either way **it is a synthetic event, so isTrusted is false**.
'
'   **A True return value does not guarantee "what happened as a result of the press"**
'     A disabled button also raises no exception, so True is returned. Whether page navigation or DOM
'     rewriting occurs is a separate matter, and as of D-3 there is still no way to wait for it.
'
'   Usage example:
'     If Not el.Click() Then Debug.Print el.LastError
' ============================================================
Public Function Click() As Boolean
    Dim js As String

    js = "if(typeof e.click==='function'){e.click();rv='click';}" & _
         "else{e.dispatchEvent(new MouseEvent('click'," & _
         "{bubbles:true,cancelable:true,view:window}));rv='dispatch';}"

    Click = ExecAction(js, "Click")
End Function


' ============================================================
' SetAttribute (D-3) - writes an attribute
'
'   Return value: True = written / False = failure (reason in LastError)
'
'   **For the difference between GetAttribute and .Value see the D-2 header**
'     SetAttribute("value", ...) changes only the **attribute** and does not change the **current value** of an
'     input the user has already edited. To change the value use .Value =.
'
'   When the attribute name is invalid (empty string etc.) JS throws InvalidCharacterError, so
'   False + that name and message go into LastError.
'
'   Usage example:
'     el.SetAttribute "data-id", "42"
' ============================================================
Public Function SetAttribute(ByVal attrName As String, _
                             ByVal attrValue As String) As Boolean
    Dim js As String

    SetAttribute = False

    If m_pane Is Nothing Then
        m_lastOk = False
        m_lastError = "no-pane"
        m_lastInfo = ""
        Exit Function
    End If

    js = "e.setAttribute(" & m_pane.JsQuote(attrName) & "," & _
         m_pane.JsQuote(attrValue) & ");rv='ok';"

    SetAttribute = ExecAction(js, "SetAttribute")
End Function

' ============================================================
' IsStale (D-2)
'   Returns whether this handle has become unusable.
'
'   Conditions for True:
'     - the JS-side registry was rebuilt by page navigation etc. (the generation changed)
'     - no element is at the index position (does not normally happen)
'   When it cannot be determined (communication failure etc.) it also errs toward the safe side and returns True.
'   In that case LastOk = False, so look at LastError if you need to distinguish.
'
'   **An element detached from the DOM does not become stale**
'     If the generation is the same and it remains in the array, False is returned. A detached element is
'     still a valid object in JS, so the read itself succeeds.
' ============================================================
Public Property Get IsStale() As Boolean
    Dim st As String
    Dim stale As Boolean

    m_lastOk = False
    m_lastError = ""
    m_lastInfo = ""

    If m_pane Is Nothing Then
        m_lastError = "no-pane"
        IsStale = True
        Exit Property
    End If

    stale = m_pane.Elem_CheckStale(m_handle, m_gen, st)

    If Len(st) = 0 Then
        m_lastOk = True
    Else
        m_lastError = st
    End If

    IsStale = stale
End Property


' ============================================================
' ReadProp (internal common: reads one property)
'   valueExpr is a JS expression that can use the variable e (= the target element).
' ============================================================
Private Function ReadProp(ByVal valueExpr As String) As String
    Dim v As String
    Dim st As String
    Dim ok As Boolean

    m_lastOk = False
    m_lastError = ""
    m_lastInfo = ""
    ReadProp = ""

    If m_pane Is Nothing Then
        m_lastError = "no-pane"
        Exit Function
    End If

    ok = m_pane.Elem_Read(m_handle, m_gen, valueExpr, v, st)

    If ok Then
        m_lastOk = True
        ReadProp = v
    Else
        m_lastError = st
    End If
End Function


' ============================================================
' ExecAction (D-3, internal common: fires one operation)
'   actionStmts is a **statement** that can use the variable e (the target element) on the JS side.
'   To keep the route name, assign it to the variable rv on the JS side.
'   The write-side entry paired with ReadProp; the way state is updated is aligned too.
' ============================================================
Private Function ExecAction(ByVal actionStmts As String, _
                            ByVal apiName As String) As Boolean
    Dim v As String
    Dim st As String

    m_lastOk = False
    m_lastError = ""
    m_lastInfo = ""
    ExecAction = False

    If m_pane Is Nothing Then
        m_lastError = "no-pane"
        Exit Function
    End If

    If m_pane.Elem_Exec(m_handle, m_gen, actionStmts, v, st) Then
        m_lastOk = True
        m_lastInfo = v
        ExecAction = True

        ' --- D-4b: opt-in auto wait (default False, so normally passes straight through) ---
        '   **It does not affect the return value.** Whether the operation ran and
        '   whether the page settled afterwards are separate matters (keeps the D-3 contract).
        '   The result of the wait is seen via m_pane.LastSettleInfo.
        If m_pane.AutoWaitAfterAction Then m_pane.WaitSettled
    Else
        m_lastError = st
        Wv2Log.LogW "Wv2Element." & apiName & ": failed (handle=" & m_handle & _
                    ") err=" & st
    End If
End Function

