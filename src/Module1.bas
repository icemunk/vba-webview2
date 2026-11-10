Attribute VB_Name = "Module1"
''''''''''''''''''''''''''''''''''
' --- Module1.bas  Stage K-4d (the entry point for verification lives here) ---
'
'   *Code written directly in the VBE is wiped on every import.ps1 run*
'   src\ is the source of truth, so on write-back the workbook side is overwritten with the contents of src\Module1.bas.
'   Put the things you often type during verification *here*.
''''''''''''''''''''''''''''''''''
Option Explicit

' *** Japanese names need the IME = trigger for the crash (spec fact 67) ***
'   In an Excel that has launched WebView2, typing into the Immediate window with the IME can crash it.
'   *ASCII aliases are provided, so use those* (Boot / BootNoQuiet / BootMaps).
'   English fork: the Japanese-named subs were removed; their bodies now live in these.

' Full startup set for the Test_* subs (3 fixed tabs + bars; includes LogStart so the log is
' confined to one run). *For your own macros use Wv2.OpenBrowser instead* (module Wv2).
Sub Boot()
    Wv2Log.LogStart
    UserForm1.Show vbModeless
    UserForm1.StartWebView2_Full
End Sub

' Turn off the quieting before release (K-4c). *Only when you want to reproduce a slow release*
'   Default is True (= send all tabs to about:blank before closing).
Sub BootNoQuiet()
    UserForm1.CurrentBrowser.QuietOnShutdown = False
End Sub

' Open Maps (the real MapsOpen. The K-4d measurement log is emitted)
Sub BootMaps()
    Wv2Maps.MapsOpen UserForm1.CurrentBrowser
End Sub

' ============================================================
' *Isolating the crash (K-5)*
'
'   Symptom: after Wv2Log.LogStart, once anything is logged,
'         *turning Japanese input on in the Immediate window and typing one character crashes Excel*.
'         With Japanese input off it is fine. Pasting Japanese text is also fine.
'
'   *Always save before running. These are experiments that assume a crash.*
'
'   Usage: run one -> turn Japanese input on in the Immediate window and type one character
'           -> see whether it crashes. *Neither WebView2 nor the form is used*
'
'     CrashTest1 ... LogStart only (the 3 header lines are written)
'     CrashTest2 ... LogStart + one LogI line     <- *candidate for the minimal repro*
'     CrashTest3 ... CrashTest2 + LogStop     <- if it does not crash, *the open file is the cause*
'     CrashTest4 ... LogStart + LogW (WARN automatically does Close -> Open)
'     CrashTest5 ... LogI with LogEcho = False  <- if it does not crash, *Debug.Print is the cause*
' ============================================================
Sub CrashTest1()
    Wv2Log.LogStart
    Debug.Print "--- CrashTest1 done. Turn Japanese input on and try typing one character ---"
End Sub

Sub CrashTest2()
    Wv2Log.LogStart
    Wv2Log.LogI "test"
    Debug.Print "--- CrashTest2 done. Turn Japanese input on and try typing one character ---"
End Sub

Sub CrashTest3()
    Wv2Log.LogStart
    Wv2Log.LogI "test"
    Wv2Log.LogStop            ' *close the file*
    Debug.Print "--- CrashTest3 done (file closed). Try typing one character ---"
End Sub

Sub CrashTest4()
    Wv2Log.LogStart
    Wv2Log.LogW "test"      ' WARN involves LogFlush (Close -> Open)
    Debug.Print "--- CrashTest4 done. Try typing one character ---"
End Sub

Sub CrashTest5()
    Wv2Log.LogStart
    Wv2Log.LogEcho = False    ' *do not emit Debug.Print*
    Wv2Log.LogI "test"
    Wv2Log.LogEcho = True
    Debug.Print "--- CrashTest5 done (log goes to the file only). Try typing one character ---"
End Sub

' ============================================================
' *Round 2: maybe Wv2Log is not even involved (K-5)*
'
'   With CrashTest1-5 (Wv2Log only) it *never crashed*.
'   It crashes when *UserForm1.Show is run after* LogStart.
'   In the reverse order (Show first, LogStart last) it does not crash.
'
'   -> The suspect is *"showing a modeless form while a file is left open"*.
'     LogStart does an Open on the log file and *leaves it open*.
'
'   *Always save before running.*
'   Run one -> turn Japanese input on in the Immediate window and type one character -> see whether it crashes.
' ============================================================

' [6] Sub version of pattern 2 (confirming the minimal repro)
Sub CrashTest6()
    Wv2Log.LogStart
    UserForm1.Show vbModeless
End Sub

' [7] Just reverse the order (if it does not crash, *order matters* is confirmed)
Sub CrashTest7()
    UserForm1.Show vbModeless
    Wv2Log.LogStart
End Sub

' [8] *Do not use Wv2Log; just Show with a file left open*
'     If it crashes, *Wv2Log is innocent*. It is then plain VBA behavior.
'     The open file is released when Excel is closed (a leftover .tmp is harmless).
Sub CrashTest8()
    Dim crashPath As String
    Dim n As Long

    crashPath = Environ$("APPDATA") & "\Wv2Browser\crashtest.tmp"
    n = FreeFile
    Open crashPath For Binary Access Write Shared As #n
    ' *deliberately not closed*
    Debug.Print "CrashTest8: calling Show with the file left open (" & crashPath & ")"
    UserForm1.Show vbModeless
End Sub

' [9] *Do not open a file; enumerate the log folder with Dir, then Show*
'     LogStart uses Dir / Kill in RotateOldLogs. See whether that is the cause.
Sub CrashTest9()
    Dim f As String
    Dim n As Long

    f = Dir(Environ$("APPDATA") & "\Wv2Browser\logs\*.log")
    Do While Len(f) > 0
        n = n + 1
        f = Dir
    Loop
    Debug.Print "CrashTest9: enumerated " & n & " log file(s), now calling Show"
    UserForm1.Show vbModeless
End Sub

' ============================================================
' *Round 3: where is the difference between Full and Spa (K-5)*
'
'   Measured: StartWebView2_Full -> *crashes* / StartWebView2_Spa -> *does not crash*
'   There are 3 differences:
'     (a) presence of TabBar / NavBar (*NavBar has the address bar input*)
'     (b) whether the content is a real site or local HTML (bing has a search box)
'     (c) number of WebView2 controllers (Full has 5, Spa has 1)
'
'   *Always save before running.*
'   Run -> without touching the code window, turn Japanese input on in the Immediate window and type one character.
' ============================================================

' [10] Spa + *one real-site tab* (no TabBar / NavBar)
'      If it crashes -> *(b) the content is the culprit* (a page with an input field). The UI parts are innocent
'      If it does not crash -> the case for *(a) TabBar / NavBar is the culprit* gets stronger
Sub CrashTest10()
    Wv2Log.LogStart
    UserForm1.Show vbModeless
    UserForm1.StartWebView2_Spa
    UserForm1.CurrentBrowser.AddTabWithUrl "https://www.bing.com"
    Debug.Print "CrashTest10: Spa + bing (no UI parts). Try typing one character"
End Sub

' [11] Spa + *three about:blank tabs* (only the count increases. No input field)
'      If it crashes -> *(c) the number of controllers is the culprit*
'      If it does not crash -> the count is irrelevant
Sub CrashTest11()
    Dim i As Long

    Wv2Log.LogStart
    UserForm1.Show vbModeless
    UserForm1.StartWebView2_Spa
    For i = 1 To 3
        UserForm1.CurrentBrowser.AddTabWithUrl "about:blank"
    Next i
    Debug.Print "CrashTest11: Spa + about:blank x3 (no input field). Try typing one character"
End Sub

' [12] Control: Full as is (*should crash for sure*)
'      The baseline that gives meaning to the other tests "not crashing"
Sub CrashTest12()
    Wv2Log.LogStart
    UserForm1.Show vbModeless
    UserForm1.StartWebView2_Full
    Debug.Print "CrashTest12: Full (control). Try typing one character"
End Sub
