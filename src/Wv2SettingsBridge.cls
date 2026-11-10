VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Wv2SettingsBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''
' --- Wv2SettingsBridge.cls  stage 9.29 (arbitrary templates outside the presets + persistence of custom) ---
'
'   Additions in 9.29 (allow entering an arbitrary search URL from the settings screen):
'     *The notation stays the prefix method (point 1, plan a)*
'       The same as the presets: a "template that concatenates the encoded search term at the end".
'       The %s placeholder method (Chrome/Firefox style) would require touching NormalizeUrl / PreviewUrlFor /
'       the existing preset definitions / compatibility of saved values all together,
'       so it is split out into a separate, independent stage.
'
'     *Validation is minimal (point 2, plan a)*
'       Only: non-empty + starts with http:// or https://. A search template is an
'       *incomplete URL* ending in ?q=, so putting it through ordinary URL validity checks
'       would reject correct input. The check is consolidated in NormalizeCustomTemplate.
'
'     *Storage format is 2 keys, and the template line only for custom (point 3, plan a-1)*
'         engine=custom
'         template=https://example.com/search?q=
'       When a preset is selected, only the single line engine=bing (same as 9.20).
'       *Why we do not always write 2 lines:* if the preset's URL is later corrected in code,
'         the saved value would revive the old URL by accident. We keep "for presets the name is
'         authoritative, the URL definition is authoritative in code".
'       Note: side effect: switching from custom to a preset deletes the template line, so the custom URL
'         remembered in the input field is also lost (by design).
'
'     - New Public (exposed to JS, point 5, plan a):
'         SetCustomEngine(template) As String
'           Validate -> apply to Browser -> save. The return value is the engine name that was settled.
'           If invalid, "" (Browser and the saved file are not touched at all).
'           *If the input was the same URL as a preset, return that preset's name
'             instead of "custom" and save it as a preset*
'             (follows the reverse lookup of TemplateToName as is. The settings screen also lights up that card)
'         GetCustomTemplate() As String
'           Initial value of the input field. If the current one is custom, that URL; otherwise the saved
'           custom; if none, "".
'         PreviewUrlForTemplate(template, terms) As String
'           For live preview *before applying*. Does not change Browser's state at all.
'           The encoding is passed through VBA instead of being home-made on the JS side, so it is exactly identical to production.
'
'     - New Public (VBA side):
'         LoadCustomTemplate() As String  (the saved custom template. Validated)
'         Debug_SaveSettings(engineName, template)  (wrapper for verification)
'
'     - Modifications to existing code:
'         LoadEngineName  -> generalized to a generic LoadValue(keyName), and made a thin wrapper.
'                           (to read the 2 keys engine / template with the same parser)
'         LoadEngineTemplate -> if engine=custom, returns the template line (validated.
'                           If invalid it is "", and Wv2Browser stays at the default Google)
'         SaveEngineName  -> made a thin wrapper over SaveSettingsFile(engineName, template).
'                           *The subject of the log changed to SaveSettingsFile*
'         Note: Wv2Browser side only receives "a template string or empty", so it is unchanged
'           (keeping the 9.20 design "Browser does not know the engine-name vocabulary").
'
' --- Wv2SettingsBridge.cls  stage 9.20 (persistence of the search engine choice) ---
'
'   Additions in 9.20 (carry the selected engine over to the next startup):
'     *The storage is an external file (APPDATA) - not a hidden sheet*
'       At first we also considered making a hidden Excel sheet the settings medium (Plan D), but
'       that depends on the Excel object model Sheets(...) and breaks portability to other
'       hosts such as Word (the lifeline of this project). So an external file
'       (%APPDATA%\Wv2Browser\settings.txt) is the storage. I/O uses only VBA built-ins
'       Open / Print / Line Input / Dir / MkDir, with neither Win32 nor ADODB
'       (engine names are ASCII so UTF-8 handling is unnecessary). This keeps it host-independent.
'
'     *The responsibility is integrated into this bridge (no separate module is cut)*
'       This class already holds the "command center of search engine settings" (preset constants SE_*,
'       name <-> template resolution, SetEngine/GetEngine). Persistence is the responsibility of
'       "keeping those settings across startups", so gathering it in the same class is natural
'       in terms of domain cohesion. A separate module like Wv2Store.bas is not cut
'       (the number of modules is not increased).
'
'     *The format is one-line key=value text*
'       Example: engine=bing . For now there is only the engine item, but if lines are added, future items such as a "template
'       outside the presets" go straight in as the second line. Loading splits each line
'       at '=' and picks up only the value of the engine key. Unknown keys are ignored.
'
'     *Timing of save/restore*
'       - Save: SaveEngineName is called internally right after SetEngine succeeds in resolving.
'         Persistence failures (no write permission etc.) are swallowed and production behavior is not stopped.
'       - Restore: Wv2Browser.Class_Initialize creates a temporary bridge with New and
'         calls LoadEngineTemplate() (Plan i). If there is a value, m_searchTemplate is overwritten
'         with it; otherwise the default (Google) stays. It is settled at Browser birth,
'         so whether via UserForm or direct launch, it consistently starts with the previous value.
'
'     *Entries exposed*
'       LoadEngineName()     As String  The saved engine name ("bing" etc.). "" if none.
'       LoadEngineTemplate() As String  The template string of the saved engine.
'                                        "" if invalid/none (the caller falls back to the default).
'       Debug_SaveEngineName(name)       For verification: a thin public wrapper over SaveEngineName.
'
' --- Wv2SettingsBridge.cls  stage 9.18 (new) ---
'
'   *What this class is*
'     A "thin entry" for operating the search engine from the settings screen (the HTML/JS of the settings tab).
'     This instance is exposed to JS with ICoreWebView2.AddHostObjectToScript, so that
'     JS can call VBA methods directly like window.chrome.webview.hostObjects.sync.Settings.SetEngine("bing")
'     (a direct RPC path that does not go through string parsing of postMessage
'     commands).
'
'   *Why we stand up a bridge instead of handing over the whole Wv2Browser (part 2-b)*
'     If Wv2Browser were directly AddHostObjectToScript-ed, the settings screen's JS could
'     hit even tab operations such as CloseTab / MoveTab (the surface exposed to JS is too wide).
'     This bridge exposes only the minimal API needed for search engine settings and internally
'     holds a single reference to Wv2Browser. The "surface exposed to JS" is structurally
'     minimized by the class boundary. The settings screen touches only the search engine and cannot reach tab operations.
'
'   *Consolidation of preset resolution (DRY)*
'     The mutual conversion between "google"/"bing"/"duckduckgo" (alias "ddg")/"yahoo" <-> template strings
'     is consolidated in this class. The preset constants (SE_*) and name-resolution logic that
'     UserForm1 held until 9.17 were moved here, and UserForm1's
'     UseSearchEngine was also changed to go through this bridge. Name resolution is in one place,
'     and both the Immediate window and the settings screen pass through the same path.
'
'   *About Instancing*
'     In the old WebView2 implementation (verified on the real machine), AddHostObjectToScript worked even with a class
'     of Instancing = 1 - Private (it is enough that the IDispatch VBA generates can be
'     exposed to the OS). So this class can also stay at the default Private.
'     PublicNotCreatable is unnecessary (we once assumed so, but that was wrong).
'
'   *Public methods (called from JS)*
'     GetEngine() As String          Returns the current engine name ("google" etc.).
'                                     "custom" if it matches none of the presets.
'     SetEngine(name) As String      Resolves the preset name and sets it on Browser.
'                                     Returns the engine name that was actually set.
'     PreviewUrl(terms) As String    The URL when the given terms are searched with the current engine.
'                                     For the live preview of the settings screen.
'     EngineList() As String         Returns the preset definitions as a JSON array.
'                                     Used by the settings screen to draw the cards.
'
'   *VBA->JS return values*
'     As confirmed in the old implementation, these return values can be received as is as the JS-side
'     function return values (via hostObjects.sync they return synchronously).
' ============================================================
Option Explicit

' --- Search engine presets (moved from UserForm1 in 9.17) ---
'   All are "query prefix type templates". The encoded search term is concatenated at the end.
'   DuckDuckGo's /?q= type and Yahoo's ?p= type differ only in the key name and can be expressed by
'   a single prefix template (confirmed in 9.17).
Private Const SE_GOOGLE     As String = "https://www.google.com/search?q="
Private Const SE_BING       As String = "https://www.bing.com/search?q="
Private Const SE_DUCKDUCKGO As String = "https://duckduckgo.com/?q="
Private Const SE_YAHOO      As String = "https://search.yahoo.com/search?p="

' --- Stage 9.20: persistence (where the search engine choice is saved) ---
'   Written to %APPDATA%\Wv2Browser\settings.txt as one line of key=value.
'   I/O uses VBA built-ins only (no Win32/ADODB). This keeps portability to Word etc.
Private Const STORE_SUBFOLDER As String = "Wv2Browser"
Private Const STORE_FILENAME  As String = "settings.txt"
Private Const STORE_KEY_ENGINE As String = "engine"
' 9.29: the 2nd key written only for custom (the template URL itself)
Private Const STORE_KEY_TEMPLATE As String = "template"

' --- Browser reference held internally (the target to which settings are applied) ---
Private m_browser As Wv2Browser


' ============================================================
' BindBrowser (called from the VBA side. An entry not exposed to JS)
'
'   Binds the Wv2Browser to be operated to the bridge. Browser calls it just before opening the settings tab
'   and doing AddHostObjectToScript.
'   *Lifetime* Only a one-way reference bridge->Browser. Browser does not hold the bridge
'     with WithEvents; it only hands it to WebView2 as a hostObject via the settings tab (Pane).
'     When the settings tab closes, the WebView2-side reference also disappears and the bridge is released.
' ============================================================
Public Sub BindBrowser(ByVal br As Wv2Browser)
    Set m_browser = br
End Sub


' ============================================================
' ResolveTemplate (internal helper: name -> template)
'
'   Resolves a preset name to a search template string. Case and leading/trailing whitespace are
'   ignored. An unknown name returns an empty string (the caller falls back to the default).
' ============================================================
Private Function ResolveTemplate(ByVal engineName As String) As String
    Dim key As String
    key = LCase$(Trim$(engineName))
    Select Case key
        Case "google":            ResolveTemplate = SE_GOOGLE
        Case "bing":              ResolveTemplate = SE_BING
        Case "duckduckgo", "ddg": ResolveTemplate = SE_DUCKDUCKGO
        Case "yahoo":             ResolveTemplate = SE_YAHOO
        Case Else:                ResolveTemplate = ""
    End Select
End Function


' ============================================================
' TemplateToName (internal helper: template -> name)
'
'   Reverse-looks up which preset the current template string matches.
'   If it matches none, returns "custom" (the settings screen can treat this as
'   "outside the presets").
' ============================================================
Private Function TemplateToName(ByVal template As String) As String
    Select Case template
        Case SE_GOOGLE:     TemplateToName = "google"
        Case SE_BING:       TemplateToName = "bing"
        Case SE_DUCKDUCKGO: TemplateToName = "duckduckgo"
        Case SE_YAHOO:      TemplateToName = "yahoo"
        Case Else:          TemplateToName = "custom"
    End Select
End Function


' ============================================================
' GetEngine (exposed to JS: the current engine name)
'
'   Reverse-looks up the search template currently set on Browser into a name and returns it.
'   Used for the initial display of "which one is currently selected" when the settings screen opens.
'   An empty string if Browser is not bound.
' ============================================================
Public Function GetEngine() As String
    If m_browser Is Nothing Then
        GetEngine = ""
        Exit Function
    End If
    GetEngine = TemplateToName(m_browser.Debug_SearchEngine())
End Function


' ============================================================
' SetEngine (exposed to JS: set the engine)
'
'   Resolves the preset name and sets it on Browser.SearchEngine. An unknown name falls to
'   the default (Google). The final guard against empty string/whitespace only is guaranteed by the Browser-side
'   Property Let, so here we concentrate on name resolution (no double guard).
'   The return value is "the engine name actually set" (used by the settings screen to settle the selection state).
' ============================================================
Public Function SetEngine(ByVal engineName As String) As String
    If m_browser Is Nothing Then
        Debug.Print "Wv2SettingsBridge.SetEngine: ignored because Browser is not bound (" & engineName & ")"
        SetEngine = ""
        Exit Function
    End If

    Dim template As String
    template = ResolveTemplate(engineName)

    If Len(template) = 0 Then
        Debug.Print "Wv2SettingsBridge.SetEngine: unknown engine name '" & engineName & _
                    "'. Using the default (Google)."
        template = SE_GOOGLE
    End If

    m_browser.SearchEngine = template
    SetEngine = TemplateToName(m_browser.Debug_SearchEngine())
    Debug.Print "Wv2SettingsBridge.SetEngine: '" & engineName & "' -> '" & SetEngine & "'"

    ' 9.20: persist the engine name that was resolved successfully (carry over to the next startup).
    '   Save failures are swallowed (production behavior is not stopped). custom is not saved
    '   (because outside-the-preset cannot currently be chosen from the UI. Template saving will be added in the future).
    If SetEngine <> "custom" And Len(SetEngine) > 0 Then
        SaveEngineName SetEngine
    End If
End Function


' ============================================================
' NormalizeCustomTemplate (internal helper: validation of a custom template, 9.29)
'
'   There are only 2 accepting conditions (point 2, plan a):
'     - the length is not 0 after dropping leading/trailing whitespace
'     - starts with http:// or https:// (case ignored)
'   If it passes, returns the Trim-med string; if not, "".
'   *Do not make it stricter than this.* A search template is an incomplete URL ending in ?q=, so
'     putting it through a general URL validity check would reject even correct input.
' ============================================================
Private Function NormalizeCustomTemplate(ByVal template As String) As String
    NormalizeCustomTemplate = ""

    Dim t As String
    t = Trim$(template)
    If Len(t) = 0 Then Exit Function

    ' English fork: control characters (CR/LF/TAB...) would break the key=value
    ' settings file and could inject extra lines into it.
    Dim k As Long
    For k = 1 To Len(t)
        If AscW(Mid$(t, k, 1)) < 32 Then Exit Function
    Next k

    Dim lo As String
    lo = LCase$(t)
    If Left$(lo, 7) = "http://" Or Left$(lo, 8) = "https://" Then
        NormalizeCustomTemplate = t
    End If
End Function


' ============================================================
' SetCustomEngine (exposed to JS: set an arbitrary template outside the presets, 9.29)
'
'   Called from "Apply" of the custom input field on the settings screen.
'     Return "custom"      : applied and saved as custom
'     Return "bing" etc.   : the input was the same URL as a preset, so it was applied
'                            and saved as that (the settings screen lights up that card)
'     Return ""            : invalid input. *Neither Browser nor the saved file was touched*
'
'   *Do not read the function name on the right-hand side* In VBA that can become a recursive call, so the settled
'     engine name is taken into a temporary variable (nm) and assigned just once at the end.
' ============================================================
Public Function SetCustomEngine(ByVal template As String) As String
    SetCustomEngine = ""

    If m_browser Is Nothing Then
        Debug.Print "Wv2SettingsBridge.SetCustomEngine: ignored because Browser is not bound"
        Exit Function
    End If

    Dim t As String
    t = NormalizeCustomTemplate(template)
    If Len(t) = 0 Then
        Debug.Print "Wv2SettingsBridge.SetCustomEngine: ignored because the template is invalid ('" & _
                    template & "')"
        Exit Function
    End If

    m_browser.SearchEngine = t

    Dim nm As String
    nm = TemplateToName(m_browser.Debug_SearchEngine())

    If nm = "custom" Then
        SaveSettingsFile "custom", t
    Else
        ' If the input was the same URL as a preset. Do not hold it as custom;
        ' save it as the preset (the template line is not written = point 3, plan a-1).
        SaveSettingsFile nm, ""
        Debug.Print "Wv2SettingsBridge.SetCustomEngine: input was identical to a preset, so saved as '" & _
                    nm & "'"
    End If

    SetCustomEngine = nm
    Debug.Print "Wv2SettingsBridge.SetCustomEngine: applied '" & t & "' -> '" & nm & "'"
End Function


' ============================================================
' GetCustomTemplate (exposed to JS: initial value of the custom input field, 9.29)
'
'   Priority:
'     1. the template currently in Browser, if it is custom
'     2. the custom template remaining in the saved file
'     3. "" if neither exists
'   2 works even if Browser is not bound, so the settings screen can always show something.
' ============================================================
Public Function GetCustomTemplate() As String
    Dim t As String
    t = ""

    If Not m_browser Is Nothing Then
        Dim cur As String
        cur = m_browser.Debug_SearchEngine()
        If TemplateToName(cur) = "custom" Then t = cur
    End If

    If Len(t) = 0 Then t = LoadCustomTemplate()

    GetCustomTemplate = t
End Function


' ============================================================
' PreviewUrl (exposed to JS: preview of the search URL)
'
'   Returns the URL actually jumped to when the given search term is searched with "the current engine".
'   Used for effects like seeing the generated URL live when the mouse is over an engine card
'   on the settings screen (natural only because hostObjects.sync is a synchronous call).
'   It passes through Browser's NormalizeUrl, so the result is identical to the production search fallback.
' ============================================================
Public Function PreviewUrl(ByVal terms As String) As String
    If m_browser Is Nothing Then
        PreviewUrl = ""
        Exit Function
    End If
    PreviewUrl = m_browser.Debug_NormalizeUrl(terms)
End Function


' ============================================================
' PreviewUrlFor (exposed to JS: search URL preview for a specified engine)
'
'   Returns what URL it would become if searched with the engine of the specified name,
'   not "the current engine" (only previewing without actually switching the engine). Used by the settings screen
'   to list the preview URL of each card.
'   To avoid changing Browser's state, the template concatenation is done by itself (the encoding is made
'   identical to production via Browser).
' ============================================================
Public Function PreviewUrlFor(ByVal engineName As String, ByVal terms As String) As String
    If m_browser Is Nothing Then
        PreviewUrlFor = ""
        Exit Function
    End If

    Dim template As String
    template = ResolveTemplate(engineName)
    If Len(template) = 0 Then template = SE_GOOGLE

    PreviewUrlFor = template & Wv2Url.UrlEncodeComponent(terms)
End Function


' ============================================================
' PreviewUrlForTemplate (exposed to JS: preview with a directly specified template, 9.29)
'
'   Builds and returns the URL actually jumped to from the string *currently being typed* in the custom input field.
'   It has not been applied yet, so it does not touch Browser's state at all (preview only).
'   Returns "" for a string that does not pass validation (the settings screen sees this and shows an error).
'   *The encoding is Wv2Url.UrlEncodeComponent* If encodeURIComponent were home-made on the JS side,
'     the result could drift from production (BuildSearchUrl), so it always goes through VBA.
' ============================================================
Public Function PreviewUrlForTemplate(ByVal template As String, ByVal terms As String) As String
    PreviewUrlForTemplate = ""

    Dim t As String
    t = NormalizeCustomTemplate(template)
    If Len(t) = 0 Then Exit Function

    PreviewUrlForTemplate = t & Wv2Url.UrlEncodeComponent(terms)
End Function


' ============================================================
' EngineList (exposed to JS: returns the preset definitions as JSON)
'
'   Returns, as a JSON array, the definitions the settings screen needs to draw the cards dynamically.
'   Each element is {"id":name, "label":display name, "template":template}.
'   When adding a preset, fixing this one place and the 2 helpers is automatically reflected in the settings screen
'   (engine names are not duplicated on the HTML side).
'
'   *The JSON string is built directly here* The values are only fixed alphanumeric URLs and labels and
'     contain no characters that need escaping, so they are concatenated plainly without going through Wv2Json.
' ============================================================
Public Function EngineList() As String
    Dim s As String
    s = "["
    s = s & "{""id"":""google"",""label"":""Google"",""template"":""" & SE_GOOGLE & """},"
    s = s & "{""id"":""bing"",""label"":""Bing"",""template"":""" & SE_BING & """},"
    s = s & "{""id"":""duckduckgo"",""label"":""DuckDuckGo"",""template"":""" & SE_DUCKDUCKGO & """},"
    s = s & "{""id"":""yahoo"",""label"":""Yahoo"",""template"":""" & SE_YAHOO & """}"
    s = s & "]"
    EngineList = s
End Function


' ============================================================
' SettingsPath (internal: a pure function that builds the full path of the saved file)
'
'   Returns %APPDATA%\Wv2Browser\settings.txt. In an environment where APPDATA cannot be obtained
'   it returns an empty string so that the caller can skip saving/loading.
' ============================================================
Private Function SettingsPath() As String
    Dim base As String
    base = Environ$("APPDATA")
    If Len(base) = 0 Then
        SettingsPath = ""
        Exit Function
    End If
    SettingsPath = base & "\" & STORE_SUBFOLDER & "\" & STORE_FILENAME
End Function


' ============================================================
' EnsureFolder (internal: create the save folder if it does not exist)
'
'   Prepares %APPDATA%\Wv2Browser\. If it exists, nothing is done. On the premise that the caller swallows
'   failure, it passes straight through here with On Error Resume Next.
' ============================================================
Private Sub EnsureFolder()
    Dim base As String
    base = Environ$("APPDATA")
    If Len(base) = 0 Then Exit Sub

    Dim folder As String
    folder = base & "\" & STORE_SUBFOLDER

    On Error Resume Next
    If Len(Dir$(folder, vbDirectory)) = 0 Then
        MkDir folder
    End If
    On Error GoTo 0
End Sub


' ============================================================
' LoadEngineName / LoadCustomTemplate / LoadValue (public: read the saved values)
'
'   *Generalized into LoadValue(keyName) in 9.29* To read the 2 keys engine and template
'   with the same parser. LoadEngineName became a thin wrapper.
'
'   Reads settings.txt line by line and returns the <value> of <key>=<value>.
'   *The value may contain =* Only the first = is used as the delimiter,
'   so values like template=https://example.com/search?q= can also be read without loss.
'   A missing file/folder, a read failure and a missing key all return ""
'   (the caller falls back to the default). Whether the name is valid (a preset or not) is not
'   judged here (LoadEngineTemplate does that).
' ============================================================
Public Function LoadEngineName() As String
    LoadEngineName = LoadValue(STORE_KEY_ENGINE)
End Function


' --- 9.29: the saved custom template (a value that does not pass validation falls to "") ---
Public Function LoadCustomTemplate() As String
    LoadCustomTemplate = NormalizeCustomTemplate(LoadValue(STORE_KEY_TEMPLATE))
End Function


' --- 9.29: generic parser that reads the value of 1 key from settings.txt ---
Private Function LoadValue(ByVal keyName As String) As String
    LoadValue = ""

    Dim path As String
    path = SettingsPath()
    If Len(path) = 0 Then Exit Function

    On Error GoTo done
    If Len(Dir$(path)) = 0 Then Exit Function   ' No file

    Dim fno As Integer
    fno = FreeFile
    Open path For Input As #fno

    Dim ln As String, eqPos As Long, k As String, v As String
    Do Until EOF(fno)
        Line Input #fno, ln
        eqPos = InStr(ln, "=")
        If eqPos > 0 Then
            k = LCase$(Trim$(Left$(ln, eqPos - 1)))
            v = Trim$(Mid$(ln, eqPos + 1))
            If k = LCase$(keyName) Then
                LoadValue = v
                ' There is no need to keep reading, but we do not break because the last definition takes priority.
            End If
        End If
    Loop

    Close #fno
    Exit Function

done:
    ' On a read failure: close it if open, and return "".
    On Error Resume Next
    Close #fno
    On Error GoTo 0
    LoadValue = ""
End Function


' ============================================================
' LoadEngineTemplate (public: returns the template string of the saved engine)
'
'   Reads the name with LoadEngineName, resolves it to a template with ResolveTemplate and returns it.
'   If there is no save or the name is unknown (custom or a broken value), returns "" so that the caller
'   (Wv2Browser.Class_Initialize) can leave the default (Google) as it is.
'   *Keep name <-> template resolution closed within this class, and hand Browser only "a template
'     string or empty"* Browser need not know the engine-name vocabulary.
' ============================================================
Public Function LoadEngineTemplate() As String
    Dim name As String
    name = LoadEngineName()
    If Len(name) = 0 Then
        LoadEngineTemplate = ""
        Exit Function
    End If

    ' 9.29: if custom, return the template line (already through NormalizeCustomTemplate).
    '   If the template line is broken or the line itself is missing, it is "" and the caller
    '   (Wv2Browser.Class_Initialize) stays at the default (Google).
    If LCase$(Trim$(name)) = "custom" Then
        LoadEngineTemplate = LoadCustomTemplate()
        Exit Function
    End If

    LoadEngineTemplate = ResolveTemplate(name)   ' "" is returned if unknown
End Function


' ============================================================
' SaveEngineName (internal: save the engine name)
'
'   Writes a single line engine=<name> to settings.txt (currently 1 item, so it is overwritten entirely).
'   If the folder does not exist it is created with EnsureFolder. Write failures are swallowed (a persistence
'   failure does not stop production). Called by SetEngine after it succeeds in resolving.
'   *When a 2nd and later item is added in the future, grow this into "read the existing one and replace only the relevant line".
'     For now there is one engine item, so a full rewrite is enough.
' ============================================================
Private Sub SaveEngineName(ByVal engineName As String)
    SaveSettingsFile engineName, ""
End Sub


' ============================================================
' SaveSettingsFile (internal: write out settings.txt, 9.29)
'
'   Always writes engine=<n>, and writes template=<URL> after it only when template is non-empty
'   (point 3, plan a-1). When a preset is selected it is called with template "" so there is only one line, and
'   *when switching from custom to a preset the template line disappears naturally*.
' ============================================================
Private Sub SaveSettingsFile(ByVal engineName As String, ByVal template As String)
    Dim path As String
    path = SettingsPath()
    If Len(path) = 0 Then Exit Sub

    EnsureFolder

    On Error GoTo fail
    Dim fno As Integer
    fno = FreeFile
    Open path For Output As #fno
    Print #fno, STORE_KEY_ENGINE & "=" & engineName
    If Len(template) > 0 Then
        Print #fno, STORE_KEY_TEMPLATE & "=" & template
    End If
    Close #fno
    Debug.Print "Wv2SettingsBridge.SaveSettingsFile: saved '" & engineName & "'" & _
                IIf(Len(template) > 0, " / template=" & template, "") & " -> " & path
    Exit Sub

fail:
    On Error Resume Next
    Close #fno
    On Error GoTo 0
    Debug.Print "Wv2SettingsBridge.SaveSettingsFile: save failed (swallowed) '" & engineName & "'"
End Sub


' ============================================================
' Debug_SaveEngineName (for verification: a public wrapper of SaveEngineName)
'
'   SaveEngineName is Private, so a thin public entry is prepared so that the pure-logic verification (Test_9_20_Persistence)
'   can check the save->load round trip. Not used on the production path.
' ============================================================
Public Sub Debug_SaveEngineName(ByVal engineName As String)
    SaveEngineName engineName
End Sub


' ============================================================
' Debug_SaveSettings (for verification: a public wrapper of SaveSettingsFile, 9.29)
'
'   Used by Test_9_29_Custom / Test_9_20_Persistence to "save and write back
'   engine + template as a pair". Not used on the production path.
' ============================================================
Public Sub Debug_SaveSettings(ByVal engineName As String, ByVal template As String)
    SaveSettingsFile engineName, template
End Sub

