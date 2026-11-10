Attribute VB_Name = "Wv2Log"
''''''''''''''''''''''''''''''''''
' --- Wv2Log.bas  Stage K-4b (take the time with a single GetLocalTime call) ---
'
'   *Fixed because K-4a measurements showed real harm*
'     000030  22:45:23.992
'     000031  22:45:23.000   <- *the second did not roll over* (should be 24.000)
'     000032  22:45:24.008
'
'   The cause was *mixing two separate time sources, Now and Timer*.
'   When Timer rolls over to the next second first, you get "previous second, milliseconds 000".
'   (The opposite skew can also happen. The two differ in both source and update interval.)
'
'   *Real harm* While timing the teardown I almost misread "one step took 1008 ms".
'   *If the measuring base is broken, the measurements themselves can no longer be trusted.*
'
'   The fix is just one call to GetLocalTime. Year/month/day/hour/minute/second/millisecond
'   are obtained at the same instant, so they *cannot be skewed in principle*.
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Log.bas  Stage K-1 (debug log written to a file) ---
'
'   *Why it is needed*
'     During D-2 verification, the Immediate window's line retention was too small and the first
'     half of the log scrolled away and could not be checked. Handle serial numbers let us confirm
'     health, but the real danger is *not noticing that lines scrolled away*. The D axis does dozens
'     of round trips per test, so this is certain to recur. Basis of all later verification.
'
'   *What it provides*
'     - LogE / LogW / LogI / LogD : one-line output per level
'     - LogPath                   : full path of the current log file (Property Get)
'     - LogLevel                  : threshold (Property Get/Let). Default is LOG_DEBUG
'     - LogEcho                   : also echo to Debug.Print (Property Get/Let). Default True
'     - LogStart / LogStop / LogFlush : explicit control
'
'   *Key design points (9 issues agreed in one go in K-1)*
'     - Output goes to %APPDATA%\Wv2Browser\logs\, the same root as settings.txt.
'       Putting it next to the workbook could collide with OneDrive sync, so we avoided that
'     - *One file per launch* (wv2_yyyymmdd_hhnnss.log). Old ones are auto-deleted, keeping 20.
'       The D axis is "one test = dozens of round trips", so one test per file works well
'     - *LogPrint also fires Debug.Print internally* so the caller needs only one line.
'       This lets us migrate by a mechanical 1-to-1 replacement of Debug.Print "..." -> Wv2Log.LogD "...",
'       and the Immediate window looks the same. LogEcho = False turns it off
'     - *The file stays open* and is written on every line. ERROR / WARN get a forced flush right after writing
'       The file is opened Shared. *Neither ADODB nor an external editor
'       could open it with Lock Write (measured in K-1 real-machine testing)* The only writer is this process,
'       so letting readers through without a lock suits the requirement better.
'       (close and reopen). We do not use the approach of accumulating in memory and writing at exit, because
'       a crash is a realistic possibility (spec fact 20)
'     - *The re-entrancy guard is a counter + pending queue* (same type as design principle 78). Dropping lines
'       that arrive during re-entry would betray K-1's very purpose of "not noticing that lines scrolled away",
'       so we queue them and write them when we leave
'     - *Lines get a serial number* A missing number reveals that lines were lost. Directly tied to K-1's motivation
'
'   *Character encoding*
'     The log file is UTF-8 / no BOM / CRLF. The UTF-8 conversion is hand-written in this module
'     (creating an ADODB.Stream per line would be heavy). The runtime String can hold arbitrary
'     Unicode from Web pages, so surrogate pairs are also folded correctly.
'     Note: this is a separate matter from the VBA "source" being tied to CP932 (spec fact 35).
'
'   *Do not reuse Wv2Tests.WriteUtf8NoBom*
'     Wv2Tests is positioned as "verification only, removable when porting" (it is not among the 12
'     production units). Product code must not depend on it.
''''''''''''''''''''''''''''''''''

Option Explicit

' ===== Log levels =====
Public Const LOG_ERROR As Long = 1
Public Const LOG_WARN As Long = 2
Public Const LOG_INFO As Long = 3
Public Const LOG_DEBUG As Long = 4

' --- K-4b: take the time with a single GetLocalTime call ---
'   *Do not mix Now and Timer* They are separate time sources, so they skew at second boundaries (see header).
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare PtrSafe Sub GetLocalTime Lib "kernel32" ( _
    ByRef lpSystemTime As SYSTEMTIME)

Private Const LOG_KEEP_FILES As Long = 20
Private Const LOG_PREFIX As String = "wv2_"

Private m_logNum As Long
Private m_logPath As String
Private m_logLevel As Long
Private m_logEcho As Boolean
Private m_logDepth As Long
Private m_logPending As String
Private m_logSeq As Long
Private m_logInit As Boolean


' ============================================================
' LogE / LogW / LogI / LogD
'
'   One-line output per level. The caller only needs to write this.
'   ERROR and WARN are force-flushed right after being written.
' ============================================================
Public Sub LogE(ByVal logMsg As String)
    LogWrite LOG_ERROR, logMsg
End Sub

Public Sub LogW(ByVal logMsg As String)
    LogWrite LOG_WARN, logMsg
End Sub

Public Sub LogI(ByVal logMsg As String)
    LogWrite LOG_INFO, logMsg
End Sub

Public Sub LogD(ByVal logMsg As String)
    LogWrite LOG_DEBUG, logMsg
End Sub


' ============================================================
' LogPath / LogLevel / LogEcho
' ============================================================
Public Property Get LogPath() As String
    LogPath = m_logPath
End Property

Public Property Get LogLevel() As Long
    If Not m_logInit Then InitDefaults
    LogLevel = m_logLevel
End Property

Public Property Let LogLevel(ByVal logArg As Long)
    If Not m_logInit Then InitDefaults
    If logArg < LOG_ERROR Then
        m_logLevel = LOG_ERROR
    ElseIf logArg > LOG_DEBUG Then
        m_logLevel = LOG_DEBUG
    Else
        m_logLevel = logArg
    End If
End Property

Public Property Get LogEcho() As Boolean
    If Not m_logInit Then InitDefaults
    LogEcho = m_logEcho
End Property

Public Property Let LogEcho(ByVal logArg As Boolean)
    If Not m_logInit Then InitDefaults
    m_logEcho = logArg
End Property


' ============================================================
' InitDefaults
'
'   Default values of the module variables. A standard module has no Class_Initialize,
'   so this runs once, the first time anything is accessed.
' ============================================================
Private Sub InitDefaults()
    If m_logInit Then Exit Sub
    m_logInit = True
    m_logLevel = LOG_DEBUG
    m_logEcho = True
    m_logSeq = 0
    m_logDepth = 0
    m_logPending = ""
End Sub


' ============================================================
' LogStart
'
'   Opens a new log file. If one is already open, closes it and reopens.
'   Normally there is no need to call this explicitly (the first LogD opens it automatically).
'   Use it when you want to restart verification cleanly.
' ============================================================
Public Sub LogStart()
    On Error Resume Next
    LogStop
    On Error GoTo 0

    InitDefaults
    m_logSeq = 0

    Dim logFolder As String
    logFolder = LogFolderPath()
    If Len(logFolder) = 0 Then Exit Sub

    RotateOldLogs logFolder

    m_logPath = logFolder & LOG_PREFIX & Format$(Now, "yyyymmdd_hhnnss") & ".log"

    On Error GoTo eh
    m_logNum = FreeFile
    Open m_logPath For Binary Access Write Shared As #m_logNum
    Seek #m_logNum, LOF(m_logNum) + 1

    Dim logHead As String
    logHead = "=== Wv2Log start " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ===" & vbCrLf & _
              "=== path = " & m_logPath & " ===" & vbCrLf & _
              "=== the 6 digits at the start of each line are a serial number; a gap means the log dropped lines ===" & vbCrLf
    PutBytes Utf8FromString(logHead)

    Debug.Print "Wv2Log.LogStart: " & m_logPath
    Exit Sub
eh:
    Debug.Print "Wv2Log.LogStart: failed (" & Err.Number & ") " & Err.Description
    m_logNum = 0
    m_logPath = ""
End Sub


' ============================================================
' LogStop
' ============================================================
Public Sub LogStop()
    If m_logNum = 0 Then Exit Sub
    On Error Resume Next
    Close #m_logNum
    On Error GoTo 0
    m_logNum = 0
End Sub


' ============================================================
' LogFlush
'
'   Closing and reopening pushes the OS buffer down to disk.
'   LogWrite calls this automatically for ERROR / WARN.
' ============================================================
Public Sub LogFlush()
    If m_logNum = 0 Then Exit Sub
    If Len(m_logPath) = 0 Then Exit Sub

    On Error GoTo eh
    Close #m_logNum
    m_logNum = FreeFile
    Open m_logPath For Binary Access Write Shared As #m_logNum
    Seek #m_logNum, LOF(m_logNum) + 1
    Exit Sub
eh:
    Debug.Print "Wv2Log.LogFlush: failed (" & Err.Number & ") " & Err.Description
    m_logNum = 0
End Sub


' ============================================================
' LogWrite  (core)
'
'   - early return on the threshold
'   - echo to Debug.Print (LogEcho)
'   - *re-entrancy guard* if re-entered while writing, push to the pending queue and flush it on the way out
' ============================================================
Private Sub LogWrite(ByVal logArg As Long, ByVal logMsg As String)
    If Not m_logInit Then InitDefaults
    If logArg > m_logLevel Then Exit Sub

    m_logSeq = m_logSeq + 1

    Dim logLine As String
    logLine = Right$("00000" & CStr(m_logSeq), 6) & " " & NowStamp() & _
              " [" & LevelTag(logArg) & "] " & logMsg

    If m_logEcho Then Debug.Print logLine

    ' --- if re-entered, just queue (same type as design principle 78) ---
    If m_logDepth > 0 Then
        m_logPending = m_logPending & logLine & vbCrLf
        Exit Sub
    End If

    m_logDepth = m_logDepth + 1
    On Error Resume Next

    If m_logNum = 0 Then LogStart
    If m_logNum <> 0 Then
        PutBytes Utf8FromString(logLine & vbCrLf)

        ' --- flush what was queued during re-entry ---
        If Len(m_logPending) > 0 Then
            Dim logDrain As String
            logDrain = m_logPending
            m_logPending = ""
            PutBytes Utf8FromString(logDrain)
        End If
    End If

    On Error GoTo 0
    m_logDepth = m_logDepth - 1

    If logArg <= LOG_WARN Then LogFlush
End Sub


' ============================================================
' Debug_SetLogDepth  (verification only)
'
'   A hook to drive the re-entrancy guard from outside. Same type as Wv2Pane.Debug_SetInCallback.
'   Setting it to 1 or more makes LogWrite only push to the pending queue; after returning to 0,
'   it is flushed all at once with the next line. *Do not call from product code*
' ============================================================
Public Sub Debug_SetLogDepth(ByVal logArg As Long)
    If Not m_logInit Then InitDefaults
    If logArg < 0 Then logArg = 0
    m_logDepth = logArg
End Sub

' ============================================================
' PutBytes
' ============================================================
Private Sub PutBytes(ByRef outBytes() As Byte)
    On Error Resume Next
    If m_logNum = 0 Then Exit Sub
    Put #m_logNum, , outBytes
End Sub


' ============================================================
' LevelTag
' ============================================================
Private Function LevelTag(ByVal logArg As Long) As String
    Select Case logArg
        Case LOG_ERROR: LevelTag = "ERR "
        Case LOG_WARN:  LevelTag = "WARN"
        Case LOG_INFO:  LevelTag = "INFO"
        Case Else:      LevelTag = "DBG "
    End Select
End Function


' ============================================================
' NowStamp (rebuilt in K-4b)
'
'   yyyy-mm-dd hh:nn:ss.mmm
'
'   *Call GetLocalTime once* Year/month/day through milliseconds are obtained at the same instant,
'   so the seconds and milliseconds cannot disagree in principle.
'   Previously Now (to the second) and Timer (milliseconds) were read separately,
'   so *at second boundaries a line with "previous second, milliseconds 000" appeared* (real harm in K-4a).
' ============================================================
Private Function NowStamp() As String
    Dim st As SYSTEMTIME

    GetLocalTime st

    NowStamp = Format$(st.wYear, "0000") & "-" & _
               Format$(st.wMonth, "00") & "-" & _
               Format$(st.wDay, "00") & " " & _
               Format$(st.wHour, "00") & ":" & _
               Format$(st.wMinute, "00") & ":" & _
               Format$(st.wSecond, "00") & "." & _
               Format$(st.wMilliseconds, "000")
End Function


' ============================================================
' LogFolderPath
'
'   Creates %APPDATA%\Wv2Browser\logs\ and returns it. Returns an empty string on failure.
' ============================================================
Private Function LogFolderPath() As String
    On Error GoTo eh
    Dim logFolder As String
    logFolder = Environ$("APPDATA")
    If Len(logFolder) = 0 Then Exit Function
    If Right$(logFolder, 1) <> "\" Then logFolder = logFolder & "\"

    logFolder = logFolder & "Wv2Browser\"
    If Dir(logFolder, vbDirectory) = "" Then MkDir logFolder

    logFolder = logFolder & "logs\"
    If Dir(logFolder, vbDirectory) = "" Then MkDir logFolder

    LogFolderPath = logFolder
    Exit Function
eh:
    Debug.Print "Wv2Log.LogFolderPath: failed (" & Err.Number & ") " & Err.Description
End Function


' ============================================================
' RotateOldLogs
'
'   Keeps the newest LOG_KEEP_FILES of wv2_*.log and deletes anything older.
'   File names are timestamps, so name order = time order.
'   *Dir keeps state, so collect everything first and delete afterwards*
' ============================================================
Private Sub RotateOldLogs(ByVal logFolder As String)
    On Error GoTo eh

    Dim logNames() As String
    Dim logCount As Long
    ReDim logNames(0 To 255)
    logCount = 0

    Dim logName As String
    logName = Dir(logFolder & LOG_PREFIX & "*.log")
    Do While Len(logName) > 0
        If logCount > UBound(logNames) Then ReDim Preserve logNames(0 To logCount + 255)
        logNames(logCount) = logName
        logCount = logCount + 1
        logName = Dir
    Loop

    If logCount <= LOG_KEEP_FILES Then Exit Sub

    ' --- sort by name (simple insertion sort; the count is small, so it is enough) ---
    Dim logIdx As Long
    Dim logLow As Long
    Dim logHigh As String
    For logIdx = 1 To logCount - 1
        logHigh = logNames(logIdx)
        logLow = logIdx - 1
        Do While logLow >= 0
            If logNames(logLow) <= logHigh Then Exit Do
            logNames(logLow + 1) = logNames(logLow)
            logLow = logLow - 1
        Loop
        logNames(logLow + 1) = logHigh
    Next logIdx

    Dim logKept As Long
    logKept = logCount - LOG_KEEP_FILES
    For logIdx = 0 To logKept - 1
        On Error Resume Next
        Kill logFolder & logNames(logIdx)
        On Error GoTo 0
    Next logIdx

    Debug.Print "Wv2Log.RotateOldLogs: deleted " & logKept & " (kept " & LOG_KEEP_FILES & ")"
    Exit Sub
eh:
    Debug.Print "Wv2Log.RotateOldLogs: failed (" & Err.Number & ") " & Err.Description
End Sub


' ============================================================
' Utf8FromString
'
'   Converts a String (sequence of UTF-16 code units) to a UTF-8 byte array.
'   *Folds surrogate pairs into 4 bytes* because the runtime String can hold
'   arbitrary Unicode from Web pages.
'   Hand-written because creating an ADODB.Stream per line would be heavy.
' ============================================================
Private Function Utf8FromString(ByVal logText As String) As Byte()
    Dim outBytes() As Byte
    Dim logLen As Long
    Dim logIdx As Long
    Dim logCount As Long
    Dim logCode As Long
    Dim logLow As Long

    logLen = Len(logText)
    If logLen = 0 Then
        ReDim outBytes(0 To 0)
        Utf8FromString = outBytes
        Exit Function
    End If

'   Worst case 4 bytes per character
    ReDim outBytes(0 To logLen * 4 - 1)
    logCount = 0
    logIdx = 1

    Do While logIdx <= logLen
        logCode = AscW(Mid$(logText, logIdx, 1))
        ' *AscW returns 0x8000 and above as a negative Integer (spec fact 16)*
        If logCode < 0 Then logCode = logCode + 65536

        ' --- surrogate pair (high D800-DBFF + low DC00-DFFF) ---
        If logCode >= &HD800& And logCode <= &HDBFF& And logIdx < logLen Then
            logLow = AscW(Mid$(logText, logIdx + 1, 1))
            If logLow < 0 Then logLow = logLow + 65536
            If logLow >= &HDC00& And logLow <= &HDFFF& Then
                logCode = &H10000 + (logCode - &HD800&) * &H400 + (logLow - &HDC00&)
                logIdx = logIdx + 1
            End If
        End If

        If logCode < &H80 Then
            outBytes(logCount) = CByte(logCode)
            logCount = logCount + 1
        ElseIf logCode < &H800 Then
            outBytes(logCount) = CByte(&HC0 + (logCode \ &H40))
            outBytes(logCount + 1) = CByte(&H80 + (logCode Mod &H40))
            logCount = logCount + 2
        ElseIf logCode < &H10000 Then
            outBytes(logCount) = CByte(&HE0 + (logCode \ &H1000))
            outBytes(logCount + 1) = CByte(&H80 + ((logCode \ &H40) Mod &H40))
            outBytes(logCount + 2) = CByte(&H80 + (logCode Mod &H40))
            logCount = logCount + 3
        Else
            outBytes(logCount) = CByte(&HF0 + (logCode \ &H40000))
            outBytes(logCount + 1) = CByte(&H80 + ((logCode \ &H1000) Mod &H40))
            outBytes(logCount + 2) = CByte(&H80 + ((logCode \ &H40) Mod &H40))
            outBytes(logCount + 3) = CByte(&H80 + (logCode Mod &H40))
            logCount = logCount + 4
        End If

        logIdx = logIdx + 1
    Loop

    ReDim Preserve outBytes(0 To logCount - 1)
    Utf8FromString = outBytes
End Function


' ============================================================
' SecondsSince (English fork)
'   Seconds elapsed since a value taken from VBA's Timer.
'   Timer resets to 0 at midnight; plain "Timer - t" then goes negative and
'   a "(Timer - t) < timeout" wait loop never times out. This wraps correctly.
' ============================================================
Public Function SecondsSince(ByVal sinceTimer As Single) As Single
    Dim d As Single
    d = Timer - sinceTimer
    If d < 0 Then d = d + 86400
    SecondsSince = d
End Function


' ============================================================
' AppFolder (English fork)
'   Per-user working folder: %LOCALAPPDATA%\Wv2Browser\<subName>, created
'   (with all parents) if missing. Replaces the old hard-coded C:\Temp\...
'   paths: folders under C:\ are usually writable by every user of the
'   machine, which exposed the WebView2 profile (cookies) and the HTML that is
'   served to pages holding host objects. Falls back to %TEMP% when
'   LOCALAPPDATA is not set. Returns the path without a trailing backslash.
' ============================================================
Public Function AppFolder(ByVal subName As String) As String
    Dim root As String
    root = Environ$("LOCALAPPDATA")
    If LenB(root) = 0 Then root = Environ$("TEMP")
    If Right$(root, 1) = "\" Then root = Left$(root, Len(root) - 1)
    root = root & "\Wv2Browser"
    If LenB(subName) > 0 Then root = root & "\" & subName
    EnsureFolderTree root
    AppFolder = root
End Function

' Creates a folder and any missing parents. Raises if a level cannot be made.
Public Sub EnsureFolderTree(ByVal folderPath As String)
    Dim parts() As String, cur As String, i As Long
    parts = Split(folderPath, "\")
    cur = parts(0)
    For i = 1 To UBound(parts)
        If LenB(parts(i)) > 0 Then
            cur = cur & "\" & parts(i)
            If LenB(Dir$(cur, vbDirectory)) = 0 Then MkDir cur
        End If
    Next i
End Sub
