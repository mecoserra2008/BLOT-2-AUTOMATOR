Attribute VB_Name = "modBlotBestEx"
Option Explicit
'==========================================================================================
' BLOT -> Best Execution monitor  (Banco de Portugal rule: >= MIN_Q counterparties asked/deal)
' v4 - configuration on the CONFIG sheet (modConfig), pure parsing core (ParseExport),
'      raw-export archive, per-date history, holidays, re-entrancy guard, self-tests
'------------------------------------------------------------------------------------------
' Public macros (buttons on DASHBOARD / CONFIG)
'   RecordExport           learns EXPORT_SEQ from ONE manual export (records clicks / keys on Bloomberg)
'   ListWindows            logs the visible window titles (to find BBG_TITLE_PREFIX)
'   RunBlotReport          end-to-end: open BLOT -> export -> parse -> archive -> history -> alerts
'   ValidateConfig         (modConfig) checks CONFIG, writes PASS/WARN/FAIL on CONFIG
'   TestExport             TEST_EXPORT_RUNS cycles of OPEN_SEQ + EXPORT_SEQ + detection (no history)
'   CaptureClickPoint      records a CLICK:x,y token relative to the Bloomberg window
'   ImportFromFile         backfill / manual load of one or more export files
'   RunSelfTests           (modSelfTest) fixture-driven tests of the parsing core
'   AbortRun               resets a stuck run and cancels the pending poll
' Other: Setup, TestDDE, ImportActiveWorkbook, ScheduleDaily, ScheduledRun, CancelSchedule
'==========================================================================================

'------------------------------------ WIN32 ----------------------------------------------
Private Type GUID_T
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Type RECT_T
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINT_T
    X As Long
    Y As Long
End Type

Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT_T) As Long
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINT_T) As Long
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As LongPtr)
Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As LongPtr, ByVal dwId As Long, riid As GUID_T, ppvObject As Object) As Long
Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, lpiid As GUID_T) As Long

Private Const OBJID_NATIVEOM As Long = &HFFFFFFF0
Private Const MOUSE_LDOWN    As Long = &H2
Private Const MOUSE_LUP      As Long = &H4
Private Const MOUSE_RDOWN    As Long = &H8
Private Const MOUSE_RUP      As Long = &H10
Private Const SW_RESTORE     As Long = 9
Private Const SW_MAXIMIZE    As Long = 3
Private Const GA_ROOT        As Long = 2
Private Const VK_LBUTTON     As Long = &H1
Private Const VK_RBUTTON     As Long = &H2

Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetAncestor Lib "user32" (ByVal hwnd As LongPtr, ByVal gaFlags As Long) As LongPtr
Private Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, lpdwProcessId As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'------------------------------------ CONSTANTS ------------------------------------------
Public Const NCOL As Long = 10            ' DEALS / HIST columns
Public Const LV_INFO  As String = "INFO"
Public Const LV_WARN  As String = "WARN"
Public Const LV_ALERT As String = "ALERT"
Public Const LV_ERROR As String = "ERROR"

'------------------------------------ STATE ----------------------------------------------
Public gHeadless As Boolean               ' TRUE during build / headless self-tests: no popup, no e-mail
Private gT0 As Date, gSince As Date, gDeadline As Date, gNextRun As Date, gSchedPending As Boolean
Private gKnown As String, gCandKey As String, gCandRows As Long, gBusy As Boolean, gKind As String
Private gPollAt As Date, gPollPending As Boolean, gCycleAt As Date, gCyclePending As Boolean
Private gTestN As Long, gTestI As Long, gTestOK As Long, gLastErr As String, gWarned As String
Private gRecTxtWait As Double, gRecTxtMain As Boolean, gWinChecked As Boolean

'==========================================================================================
'                                    ENTRY POINTS
'==========================================================================================
Public Sub RunBlotReport()
    On Error GoTo EH
    LoadConfig
    If Not BeginRun("RUN") Then Exit Sub
    Notify "Run started: SOURCE_MODE=" & SOURCE_MODE & ", expected trade date " & _
           Format$(PrevBusinessDay(Date), "yyyy-mm-dd") & " (DAYS_BACK=" & DaysBack(Date) & ").", LV_INFO
    If Not StartAcquisition() Then FinishRun
    Exit Sub
EH:
    FinishRun
    Notify "ERROR in RunBlotReport: " & Err.Description, LV_ERROR
End Sub

Public Sub PollExport()
    Dim found As Boolean
    gPollPending = False
    If Not gBusy Then Exit Sub
    On Error GoTo EH
    found = TryAcquire()
    If found Then Exit Sub
Reschedule:
    On Error GoTo 0
    If Not gBusy Then Exit Sub
    If Now < gDeadline Then
        SchedulePoll
    Else
        OnAcquireTimeout
    End If
    Exit Sub
EH:
    gLastErr = Err.Description
    Resume Reschedule
End Sub

Public Sub AbortRun()
    CancelPending
    gBusy = False: gKind = ""
    Application.StatusBar = False
    Notify "Run aborted by the operator (state reset, pending poll cancelled).", LV_WARN, False, True
End Sub

Public Sub TestDDE()
    Dim msg As String, ch As Long
    LoadConfig
    If DdeAvailable(msg) Then
        On Error Resume Next
        ch = Application.DDEInitiate("WINBLP", "BBK")
        Application.DDEExecute ch, BBG_PANEL & "BLOT<GO>"
        Application.DDETerminate ch
        On Error GoTo 0
        Notify "TestDDE: " & msg & " Sent " & BBG_PANEL & "BLOT<GO>.", LV_INFO, True, True
    Else
        Notify "TestDDE: " & msg, LV_WARN, True, True
    End If
End Sub

Public Function DdeAvailable(ByRef msg As String) As Boolean
    Dim ch As Long, da As Boolean
    If Application.IgnoreRemoteRequests Then
        msg = "Excel option 'Ignore other applications that use Dynamic Data Exchange (DDE)' is ON: turn it off (File > Options > Advanced > General)."
        Exit Function
    End If
    da = Application.DisplayAlerts: Application.DisplayAlerts = False
    On Error GoTo Fail
    ch = Application.DDEInitiate("WINBLP", "BBK")
    Application.DDETerminate ch
    Application.DisplayAlerts = da
    msg = "DDE link WINBLP|BBK OK."
    DdeAvailable = True
    Exit Function
Fail:
    Application.DisplayAlerts = da
    msg = "DDE link WINBLP|BBK failed (" & Err.Description & "). Terminal running and logged in?"
End Function

' TEST_EXPORT_RUNS consecutive cycles of the real acquisition path; nothing is written to history
Public Sub TestExport()
    On Error GoTo EH
    LoadConfig
    If SOURCE_MODE = "TERMINAL" And Len(EXPORT_SEQ) = 0 Then
        Notify "TestExport: EXPORT_SEQ is empty (CONFIG). Record it with Capture Click Point or numbered items.", LV_WARN, True, True
        Exit Sub
    End If
    If Not BeginRun("TEST") Then Exit Sub
    gTestN = IIf(SOURCE_MODE = "TERMINAL", Application.WorksheetFunction.Max(1, TEST_EXPORT_RUNS), 1)
    gTestI = 0: gTestOK = 0
    Notify "TestExport started: " & gTestN & " cycle(s), SOURCE_MODE=" & SOURCE_MODE & ".", LV_INFO
    StartTestCycle
    Exit Sub
EH:
    FinishRun
    Notify "ERROR in TestExport: " & Err.Description, LV_ERROR, True, True
End Sub

Public Sub StartTestCycle()
    gCyclePending = False
    If Not gBusy Or gKind <> "TEST" Then Exit Sub
    On Error GoTo EH
    gTestI = gTestI + 1
    Application.StatusBar = "BLOT Test Export: cycle " & gTestI & "/" & gTestN
    If Not StartAcquisition() Then TestCycleDone False, "steps failed (see LOG)"
    Exit Sub
EH:
    TestCycleDone False, "error: " & Err.Description
End Sub

Public Sub CaptureClickPoint()
    Dim h As LongPtr, rc As RECT_T, pt As POINT_T, tok As String, ans As VbMsgBoxResult
    On Error GoTo EH
    LoadConfig
    h = BbgHwnd()
    If h = 0 Then MsgBox "Bloomberg window '" & BBG_TITLE_PREFIX & "' not found.", vbCritical: Exit Sub
    MsgBox "After OK you have 5 seconds: place the mouse over the target button in the Bloomberg window and keep it still.", vbInformation
    WaitSec 5
    GetCursorPos pt: GetWindowRect h, rc
    tok = "CLICK:" & (pt.X - rc.Left) & "," & (pt.Y - rc.Top)
    WriteLog "CaptureClickPoint: " & tok & " (window at " & rc.Left & "," & rc.Top & " size " & (rc.Right - rc.Left) & "x" & (rc.Bottom - rc.Top) & ")", LV_INFO
    ans = MsgBox("Captured step token:" & vbLf & vbLf & tok & vbLf & vbLf & _
                 "Append it to EXPORT_SEQ on the CONFIG sheet?" & vbLf & "(a WAIT:1 step is inserted before it when EXPORT_SEQ is not empty)", _
                 vbYesNo + vbQuestion, "CaptureClickPoint")
    If ans = vbYes Then
        AppendStepToConfig "EXPORT_SEQ", tok
        MsgBox "EXPORT_SEQ updated. Run Test Export to verify (5/5).", vbInformation
    End If
    Exit Sub
EH:
    Notify "ERROR in CaptureClickPoint: " & Err.Description, LV_ERROR, True, True
End Sub

Public Sub ImportActiveWorkbook()
    Dim blk As Object, wb As Workbook
    On Error GoTo EH
    LoadConfig
    Set wb = ActiveWorkbook
    If wb Is ThisWorkbook Then MsgBox "Activate the export workbook first.", vbExclamation: Exit Sub
    If Not BeginRun("IMPORT") Then Exit Sub
    Set blk = FindBlockInWorkbook(wb)
    If blk Is Nothing Then
        Notify "ImportActiveWorkbook: no header row with a dealer column (HDR_DEALERS) in " & wb.Name & ".", LV_WARN, True, True
    Else
        ProcessData blk("v"), blk("raw"), "workbook " & blk("label"), wb, ""
    End If
    FinishRun
    Exit Sub
EH:
    FinishRun
    Notify "ERROR in ImportActiveWorkbook: " & Err.Description, LV_ERROR
End Sub

Public Sub ImportFromFile()
    Dim files As Variant, f As Variant, n As Long, tot As Long
    On Error GoTo EH
    LoadConfig
    files = Application.GetOpenFilename("BLOT export (*.xls*;*.csv),*.xls*;*.csv", , "Select BLOT export file(s)", , True)
    If Not IsArray(files) Then Exit Sub
    If Not BeginRun("IMPORT") Then Exit Sub
    For Each f In files
        tot = tot + 1
        If ImportOneFile(CStr(f)) Then n = n + 1
    Next
    FinishRun
    Notify "Import finished: " & n & " of " & tot & " file(s) loaded.", LV_INFO, True, True
    Exit Sub
EH:
    FinishRun
    Notify "ERROR in ImportFromFile: " & Err.Description, LV_ERROR
End Sub

Private Function ImportOneFile(ByVal f As String) As Boolean
    Dim blk As Object, wb As Workbook
    On Error GoTo EH
    If IsCsvPath(f) Then
        Set blk = CsvBlock(f)
    Else
        Set wb = Workbooks.Open(f, UpdateLinks:=0, ReadOnly:=True)
        Set blk = FindBlockInWorkbook(wb)
    End If
    If blk Is Nothing Then
        Notify "Import: no header row with a dealer column (HDR_DEALERS) in " & f & ".", LV_WARN
    Else
        ImportOneFile = ProcessData(blk("v"), blk("raw"), "file " & f, Nothing, f)
    End If
    If Not wb Is Nothing Then wb.Close False
    Exit Function
EH:
    Notify "ERROR importing " & f & ": " & Err.Description, LV_ERROR
    On Error GoTo -1
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
End Function

'==========================================================================================
'                                    RECORD EXPORT (learns EXPORT_SEQ from one manual export)
'==========================================================================================
Public Sub RecordExport()
    Dim bb As LongPtr, bbPid As Long, fgRaw As LongPtr, root As LongPtr, rc As RECT_T, pt As POINT_T
    Dim t0 As Double, tLast As Double, tNow As Double, tCheck As Double, tStop As Double
    Dim toks As Collection, txt As String, prevL As Boolean, prevR As Boolean, l As Boolean, r As Boolean
    Dim kPrev(0 To 255) As Boolean, vk As Long, kd As Boolean, inB As Boolean, isMain As Boolean
    Dim known As String, newWb As String, stopWhy As String, seq As String, title As String, prefix As String
    Dim lastX As Long, lastY As Long, lastT As Double, x As Long, y As Long, winInfo As String
    Dim ans As VbMsgBoxResult, identify As Boolean
    On Error GoTo EH
    LoadConfig
    If gBusy Then MsgBox "A run is in progress: click Abort first.", vbExclamation: Exit Sub
    ans = MsgBox("RECORD EXPORT - the macro learns EXPORT_SEQ from ONE export done by hand." & vbLf & vbLf & _
        "1. YES = first run OPEN_SEQ (BLOT 2 / BLOT -1) exactly like the daily run (recommended)." & vbLf & _
        "2. Then do the export BY HAND in the Bloomberg panel, the way you always do it:" & vbLf & _
        "   clicks, right-clicks, menu numbers + <GO>, Enter, Esc, until the export opens in Excel." & vbLf & _
        "3. Recording stops by itself when the export workbook appears" & vbLf & _
        "   (or Ctrl+Shift+Q, or after 3 minutes). Clicks outside Bloomberg are ignored." & vbLf & vbLf & _
        "If the Bloomberg window is not recognised yet, your FIRST click on it only identifies it." & vbLf & vbLf & _
        "Run OPEN_SEQ first?", vbYesNoCancel + vbQuestion, "Record Export")
    If ans = vbCancel Then Exit Sub
    If ans = vbYes Then
        If Not RunSteps(ExpandSeq(OPEN_SEQ, Date)) Then Exit Sub
    End If
    known = SnapshotWorkbooks()
    bb = BbgHwnd()
    If bb <> 0 Then
        GetWindowThreadProcessId bb, bbPid
        BbgFocus
    Else
        identify = True
    End If
    Set toks = New Collection
    For vk = 1 To 254: kPrev(vk) = KeyIsDown(vk): Next
    prevL = kPrev(VK_LBUTTON): prevR = kPrev(VK_RBUTTON)
    t0 = NowSec(): tLast = t0: tCheck = t0
    Application.StatusBar = "BLOT Record Export: RECORDING - do the export in Bloomberg now (Ctrl+Shift+Q stops)"
    WriteLog "Record Export started (Bloomberg window " & IIf(bb = 0, "not recognised yet: first click identifies it", "'" & WindowTitle(bb) & "'") & ").", LV_INFO
    Do
        Sleep 8
        DoEvents
        tNow = NowSec()
        If tStop > 0 And tNow >= tStop Then Exit Do
        If KeyIsDown(&H11) And KeyIsDown(&H10) And KeyIsDown(&H51) Then stopWhy = "stopped with Ctrl+Shift+Q": Exit Do
        If tNow - t0 > 180 Then stopWhy = "3-minute limit": Exit Do

        fgRaw = GetForegroundWindow(): root = GetAncestor(fgRaw, GA_ROOT)
        inB = False: isMain = False
        If bb <> 0 And root <> 0 Then
            isMain = (root = bb)
            inB = isMain Or (WindowPid(root) = bbPid And bbPid <> 0)
        End If

        l = KeyIsDown(VK_LBUTTON): r = KeyIsDown(VK_RBUTTON)
        If ((l And Not prevL) Or (r And Not prevR)) And tStop = 0 Then
            GetCursorPos pt
            Sleep 40
            fgRaw = GetForegroundWindow(): root = GetAncestor(fgRaw, GA_ROOT)
            If identify Then
                If root <> 0 And WindowPid(root) <> WindowPid(Application.Hwnd) Then
                    bb = root: GetWindowThreadProcessId bb, bbPid: identify = False: tLast = tNow
                    WriteLog "Record Export: Bloomberg window identified as '" & WindowTitle(bb) & "'.", LV_INFO
                End If
            ElseIf bb <> 0 And (root = bb Or (WindowPid(root) = bbPid And bbPid <> 0)) Then
                GetWindowRect bb, rc
                x = pt.X - rc.Left: y = pt.Y - rc.Top
                FlushRecText toks, txt, tNow, tLast
                If l And toks.Count > 0 And tNow - lastT < 0.45 And Abs(x - lastX) <= 4 And Abs(y - lastY) <= 4 And Left$(toks(toks.Count), 6) = "CLICK:" Then
                    toks.Remove toks.Count
                    toks.Add "DCLICK:" & x & "," & y
                Else
                    AddRecTok toks, IIf(l, "CLICK:", "RCLICK:") & x & "," & y, tNow, tLast
                End If
                lastX = x: lastY = y: lastT = tNow
            End If
        End If
        prevL = l: prevR = r

        For vk = 8 To 222
            kd = KeyIsDown(vk)
            If kd And Not kPrev(vk) And inB And tStop = 0 Then RecordKey toks, txt, vk, isMain, tNow, tLast
            kPrev(vk) = kd
        Next

        If tNow - tCheck >= 1 And tStop = 0 Then
            tCheck = tNow
            newWb = NewWorkbookSince(known)
            If Len(newWb) > 0 Then stopWhy = "export detected: " & newWb: tStop = tNow + 0.5
        End If
    Loop
    FlushRecText toks, txt, NowSec(), tLast
    Application.StatusBar = False

    If toks.Count = 0 Then
        Notify "Record Export: nothing was recorded on the Bloomberg window (" & stopWhy & ").", LV_WARN, True, True
        Exit Sub
    End If
    seq = JoinTokens(toks) & "|WAIT:3"
    title = WindowTitle(bb)
    GetWindowRect bb, rc
    winInfo = (rc.Right - rc.Left) & "x" & (rc.Bottom - rc.Top) & IIf(IsZoomed(bb) <> 0, ",MAX", "")
    WriteLog "Record Export (" & stopWhy & "): EXPORT_SEQ = " & seq & " | window '" & title & "' " & winInfo, LV_INFO
    ans = MsgBox("Recorded (" & stopWhy & "):" & vbLf & vbLf & seq & vbLf & vbLf & _
                 "Bloomberg window: '" & title & "'  " & winInfo & vbLf & vbLf & _
                 "Save it on CONFIG (replaces EXPORT_SEQ)?", vbYesNo + vbQuestion, "Record Export")
    If ans <> vbYes Then Exit Sub
    SetConfigValue "EXPORT_SEQ", seq
    SetConfigValue "BBG_WINDOW", winInfo
    If Len(title) > 0 And (Len(BBG_TITLE_PREFIX) = 0 Or UCase$(Left$(title, Len(BBG_TITLE_PREFIX))) <> UCase$(BBG_TITLE_PREFIX)) Then
        prefix = InputBox("The Bloomberg window title is:" & vbLf & vbLf & title & vbLf & vbLf & _
                          "Start of the title that ALWAYS identifies this panel (saved as BBG_TITLE_PREFIX):", "Record Export", SuggestPrefix(title))
        If Len(prefix) > 0 Then SetConfigValue "BBG_TITLE_PREFIX", prefix
    End If
    If Len(newWb) = 0 Then
        MsgBox "Saved. No export workbook was detected while recording - check that the last step really exported, then run Test Export.", vbExclamation, "Record Export"
    Else
        MsgBox "Saved. Close the export workbook '" & newWb & "', then click Test Export (it must succeed 5/5).", vbInformation, "Record Export"
    End If
    Exit Sub
EH:
    Application.StatusBar = False
    Notify "ERROR in RecordExport: " & Err.Description, LV_ERROR, True, True
End Sub

Private Sub RecordKey(toks As Collection, ByRef txt As String, ByVal vk As Long, ByVal isMain As Boolean, ByVal tNow As Double, ByRef tLast As Double)
    Dim ctl As Boolean, alt As Boolean, sh As Boolean, c As String, sk As String
    Select Case vk
        Case &H10, &H11, &H12, &H14, &H5B, &H5C, &H90, &H91, &HA0 To &HA5
            Exit Sub                                          ' modifiers / locks
    End Select
    ctl = KeyIsDown(&H11): alt = KeyIsDown(&H12): sh = KeyIsDown(&H10)
    If ctl And sh And vk = &H51 Then Exit Sub                 ' stop combination
    c = KeyChar(vk)
    If Len(c) > 0 And Not ctl And Not alt Then
        If Len(txt) = 0 Then gRecTxtWait = RecWait(tNow - tLast): gRecTxtMain = isMain: tLast = tNow
        txt = txt & c
        Exit Sub
    End If
    Select Case vk
        Case &HD                                              ' Enter = <GO>
            If Len(txt) > 0 Then
                If gRecTxtMain Then
                    AddRecTokW toks, "KEYS:" & txt & "<GO>", gRecTxtWait
                Else
                    AddRecTokW toks, "SEND:" & EscKeys(txt) & "~", gRecTxtWait
                End If
                txt = "": tLast = tNow
            Else
                AddRecTok toks, "SEND:~", tNow, tLast
            End If
            Exit Sub
        Case &H8
            If Len(txt) > 0 Then txt = Left$(txt, Len(txt) - 1): Exit Sub
            sk = "{BACKSPACE}"
        Case &H1B: sk = "{ESC}"
        Case &H9:  sk = IIf(sh, "+{TAB}", "{TAB}")
        Case &H25: sk = "{LEFT}"
        Case &H26: sk = "{UP}"
        Case &H27: sk = "{RIGHT}"
        Case &H28: sk = "{DOWN}"
        Case &H21: sk = "{PGUP}"
        Case &H22: sk = "{PGDN}"
        Case &H24: sk = "{HOME}"
        Case &H23: sk = "{END}"
        Case &H2E: sk = "{DEL}"
        Case &H70 To &H7B: sk = "{F" & (vk - &H6F) & "}"
        Case Else
            If Len(c) = 0 Then Exit Sub
            sk = IIf(ctl, "^", "") & IIf(alt, "%", "") & LCase$(c)
    End Select
    FlushRecText toks, txt, tNow, tLast
    AddRecTok toks, "SEND:" & sk, tNow, tLast
End Sub

Private Function KeyChar(ByVal vk As Long) As String
    Select Case vk
        Case &H30 To &H39: KeyChar = Chr$(vk)
        Case &H60 To &H69: KeyChar = Chr$(vk - &H30)
        Case &H41 To &H5A: KeyChar = Chr$(vk)
        Case &H20: KeyChar = " "
        Case &HBD, &H6D: KeyChar = "-"
        Case &HBE, &H6E: KeyChar = "."
        Case &HBC: KeyChar = ","
    End Select
End Function

Private Sub FlushRecText(toks As Collection, ByRef txt As String, ByVal tNow As Double, ByRef tLast As Double)
    If Len(txt) = 0 Then Exit Sub
    AddRecTokW toks, "SEND:" & EscKeys(txt), gRecTxtWait
    txt = "": tLast = tNow
End Sub

Private Sub AddRecTok(toks As Collection, ByVal tok As String, ByVal tNow As Double, ByRef tLast As Double)
    AddRecTokW toks, tok, RecWait(tNow - tLast)
    tLast = tNow
End Sub

Private Sub AddRecTokW(toks As Collection, ByVal tok As String, ByVal w As Double)
    toks.Add "WAIT:" & InvariantNum(w)
    toks.Add tok
End Sub

' Human pause -> WAIT seconds: rounded up to 0.5, between 0.5 and 8
Public Function RecWait(ByVal dt As Double) As Double
    Dim w As Double
    w = -Int(-dt * 2) / 2
    If w < 0.5 Then w = 0.5
    If w > 8 Then w = 8
    RecWait = w
End Function

Private Function JoinTokens(toks As Collection) As String
    Dim t As Variant, s As String
    For Each t In toks: s = s & "|" & t: Next
    JoinTokens = Mid$(s, 2)
End Function

Public Function SuggestPrefix(ByVal title As String) As String
    Dim p As Long
    p = InStr(1, title, " - ")
    If p > 1 Then SuggestPrefix = Left$(title, p - 1) Else SuggestPrefix = title
End Function

Private Function KeyIsDown(ByVal vk As Long) As Boolean
    KeyIsDown = (GetAsyncKeyState(vk) And &H8000) <> 0
End Function

Private Function NowSec() As Double
    NowSec = CDbl(Date) * 86400# + Timer
End Function

Private Function WindowPid(ByVal h As LongPtr) As Long
    Dim pid As Long
    If h <> 0 Then GetWindowThreadProcessId h, pid
    WindowPid = pid
End Function

Private Function WindowTitle(ByVal h As LongPtr) As String
    Dim s As String, l As Long
    If h = 0 Then Exit Function
    s = String$(260, vbNullChar): l = GetWindowText(h, s, 260)
    If l > 0 Then WindowTitle = Left$(s, l)
End Function

Private Function NewWorkbookSince(ByVal known As String) As String
    Dim app As Object, wb As Object
    On Error Resume Next
    For Each app In AllExcelApps()
        For Each wb In app.Workbooks
            If InStr(1, known, "|" & app.Hwnd & ":" & wb.FullName & "|", vbTextCompare) = 0 Then NewWorkbookSince = wb.Name: Exit Function
        Next
    Next
End Function

' Visible top-level window titles; onlyLikely = Bloomberg-looking titles only
Public Function WindowCandidates(ByVal onlyLikely As Boolean) As String
    Dim h As LongPtr, t As String, s As String, n As Long
    h = FindWindowEx(0, 0, vbNullString, vbNullString)
    Do While h <> 0 And n < 80
        If IsWindowVisible(h) <> 0 Then
            t = WindowTitle(h)
            If Len(t) > 0 Then
                If Not onlyLikely Or InStr(1, t, "bloomberg", vbTextCompare) > 0 Or InStr(1, t, "blp", vbTextCompare) > 0 Or t Like "#-*" Then
                    s = s & "'" & t & "'; ": n = n + 1
                End If
            End If
        End If
        h = FindWindowEx(0, h, vbNullString, vbNullString)
    Loop
    WindowCandidates = s
End Function

Public Sub ListWindows()
    Dim c As String
    LoadConfig
    c = WindowCandidates(True)
    WriteLog "Visible windows that look like Bloomberg: " & IIf(Len(c) = 0, "(none)", c), LV_INFO
    WriteLog "All visible windows: " & WindowCandidates(False), LV_INFO
    If Application.Visible Then
        MsgBox "Windows that look like Bloomberg:" & vbLf & vbLf & IIf(Len(c) = 0, "(none - see LOG for all windows)", Replace(c, "; ", vbLf)) & vbLf & vbLf & _
               "BBG_TITLE_PREFIX must be the start of the panel title (now: '" & BBG_TITLE_PREFIX & "'). Record Export sets it for you.", vbInformation, "List Windows"
    End If
End Sub

' Before clicks: restore the window state / size used when EXPORT_SEQ was recorded (BBG_WINDOW)
Private Sub CheckBbgWindow(ByVal h As LongPtr)
    Dim rc As RECT_T, cur As String, want As String
    If h = 0 Or Len(BBG_WINDOW) = 0 Then Exit Sub
    If InStr(1, BBG_WINDOW, "MAX", vbTextCompare) > 0 And IsZoomed(h) = 0 Then ShowWindow h, SW_MAXIMIZE: WaitSec 0.8
    GetWindowRect h, rc
    cur = (rc.Right - rc.Left) & "x" & (rc.Bottom - rc.Top)
    want = Split(BBG_WINDOW, ",")(0)
    If cur <> want Then WriteLog "Bloomberg window is " & cur & " but EXPORT_SEQ was recorded at " & want & ": clicks may miss (maximize the panel on the same monitor or re-record).", LV_WARN
End Sub

' Focus Bloomberg unless a Bloomberg window (panel, menu or dialog) is already in front
Private Function BbgFocusUnlessOwn() As Boolean
    Dim h As LongPtr, pid As Long
    h = BbgHwnd(): If h = 0 Then Exit Function
    GetWindowThreadProcessId h, pid
    If WindowPid(GetAncestor(GetForegroundWindow(), GA_ROOT)) = pid And pid <> 0 Then
        BbgFocusUnlessOwn = True
    Else
        BbgFocusUnlessOwn = BbgFocus()
    End If
End Function

'==========================================================================================
'                                    SETUP
'==========================================================================================
Public Sub Setup()
    Dim ws As Worksheet, nm As Variant, prev As Worksheet, i As Long, su As Boolean, da As Boolean
    On Error GoTo EH
    su = Application.ScreenUpdating: Application.ScreenUpdating = False
    For Each nm In SheetOrder(): GetSheet CStr(nm): Next

    WriteConfigSheet
    LoadConfig

    Set ws = GetSheet(SH_MAP)
    If Len(ws.Range("A1").Value) = 0 Then
        ws.Range("A1:B1").Value = Array("Alias (as exported)", "Counterparty (legal entity)")
        ws.Range("A1:B1").Font.Bold = True: ws.Range("A1:B1").Interior.Color = RGB(217, 225, 242)
        ws.Columns("A:B").ColumnWidth = 32: ws.Columns("A:B").NumberFormat = "@"
        ws.Range("D1").Value = "One row per alias. Example: JPM -> JPMORGAN, JPM LDN -> JPMORGAN, JPMS -> JPMORGAN."
        ws.Range("D2").Value = "Matching ignores case and repeated spaces. Unmapped names are listed in LOG after every run."
        ws.Range("D1:D2").Font.Italic = True
    End If

    EnsureHeader GetSheet(SH_HIST), HistHeader()
    EnsureHeader GetSheet(SH_HD), HistDailyHeader()
    EnsureLogHeader GetSheet(SH_LOG)
    WriteReadme
    If Len(GetSheet(SH_DEALS).Range("A2").Value) = 0 Then
        BuildDashboard Empty, 0, "No run yet."
    Else
        AddSheetButtons GetSheet(SH_DASH), DashButtons()
    End If
    GetSheet(SH_FIX).Visible = xlSheetHidden

    da = Application.DisplayAlerts: Application.DisplayAlerts = False
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        If Not IsOwnSheet(ws.Name) Then
            If Application.WorksheetFunction.CountA(ws.Cells) = 0 And ws.Shapes.Count = 0 Then ws.Delete
        End If
    Next
    Application.DisplayAlerts = da
    Set prev = Nothing
    For Each nm In SheetOrder()
        Set ws = GetSheet(CStr(nm))
        If prev Is Nothing Then ws.Move Before:=ThisWorkbook.Worksheets(1) Else ws.Move After:=prev
        Set prev = ws
    Next
    GetSheet(SH_DASH).Activate
    Application.ScreenUpdating = su
    WriteLog "Setup completed (sheets, CONFIG defaults, buttons, README).", LV_INFO
    If Application.Visible And Not gHeadless Then MsgBox "Setup complete. Fill the yellow cells on CONFIG, then Validate Config.", vbInformation, "BLOT Best Execution"
    Exit Sub
EH:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Notify "ERROR in Setup: " & Err.Description, LV_ERROR
End Sub

Private Function SheetOrder() As Variant
    SheetOrder = Array(SH_README, SH_DASH, SH_CONFIG, SH_MAP, SH_DEALS, SH_HIST, SH_HD, SH_RAW, SH_LOG, SH_TESTS, SH_FIX)
End Function

Private Function IsOwnSheet(ByVal nm As String) As Boolean
    Dim s As Variant
    For Each s In SheetOrder()
        If StrComp(CStr(s), nm, vbTextCompare) = 0 Then IsOwnSheet = True: Exit Function
    Next
End Function

Public Function DashButtons() As Variant
    DashButtons = Array("Run Now|RunBlotReport", "Validate Config|ValidateConfig", "Record Export|RecordExport", _
                        "Test Export|TestExport", "Capture Click Point|CaptureClickPoint", "Import File|ImportFromFile", _
                        "Run Self-Tests|RunSelfTests", "Abort|AbortRun")
End Function

Public Sub AddSheetButtons(ws As Worksheet, defs As Variant)
    Dim i As Long, b As Object, x As Double, p() As String
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).Name Like "btn_*" Then ws.Shapes(i).Delete
    Next
    ws.Rows(1).RowHeight = 34
    x = 6
    For i = LBound(defs) To UBound(defs)
        p = Split(defs(i), "|")
        Set b = ws.Buttons.Add(x, 4, 118, 26)
        b.OnAction = p(1)
        b.Caption = p(0)
        b.Name = "btn_" & p(1)
        b.Placement = xlFreeFloating
        x = x + 124
    Next
End Sub

Private Sub WriteReadme()
    Dim ws As Worksheet, t As Collection, i As Long, s As String
    Set ws = GetSheet(SH_README)
    ws.Cells.Clear
    Set t = New Collection
    t.Add "#BLOT Best-Execution Monitor - operator guide"
    t.Add "Banco de Portugal: every deal must have quotes requested from at least 3 counterparties. This workbook checks it every business day and keeps the evidence."
    t.Add ""
    t.Add "#1. One-time setup"
    t.Add "a) Excel > Options > Advanced > General: UNCHECK 'Ignore other applications that use Dynamic Data Exchange (DDE)'."
    t.Add "b) Keep this file in a Trusted Location (or have IT sign it). Bloomberg Terminal logged in; Outlook open if e-mail is used."
    t.Add "c) CONFIG sheet: fill the yellow cells - ARCHIVE_DIR (required), EXPORT_FOLDER (if used), ALERT_TO, SHOW_POPUPS=FALSE for unattended runs."
    t.Add "d) DEALER_MAP sheet: one row per dealer alias -> legal entity (e.g. JPM, JPM LDN, JPMS -> JPMORGAN). Otherwise aliases count as different counterparties."
    t.Add "e) Click 'Validate Config' and fix every FAIL line (CONFIG columns E:G)."
    t.Add ""
    t.Add "#2. Calibrate the export (TERMINAL mode)"
    t.Add "a) Maximise the Bloomberg panel (same monitor and size as on the run days). Close other Bloomberg popups."
    t.Add "b) Click 'Record Export', answer Yes (it opens BLOT 2 / BLOT -1), then do the export BY HAND exactly as usual"
    t.Add "   (clicks, right-clicks, menu numbers + <GO>, Enter). Recording stops when the export appears in Excel."
    t.Add "   It writes EXPORT_SEQ, BBG_WINDOW and (if needed) BBG_TITLE_PREFIX on CONFIG. Close the export workbook."
    t.Add "c) Click 'Test Export': it runs TEST_EXPORT_RUNS (5) cycles. Accept EXPORT_SEQ only with 5/5 successes (see LOG)."
    t.Add "d) Check the LOG line 'Header mapping': every role must point to the right export column. Put the exact export headers first in HDR_* on CONFIG."
    t.Add "e) Pin DEALER_MODE (ROWS preferred) and set DATE_FORMAT_TEXT if the export has text dates (e.g. MM/DD/YYYY)."
    t.Add "f) Compare 5-10 real deals by hand with the BLOT inquiry detail (DEALS column NDealers / Dealers)."
    t.Add "g) If 'BLOT -1' means one CALENDAR day back, use KEYS:BLOT -{DAYS_BACK}<GO> in OPEN_SEQ so Monday covers Friday (DAYS_BACK = days to previous business day)."
    t.Add ""
    t.Add "#3. Daily operation"
    t.Add "The workbook arms a daily run at RUN_AT when it is opened (keep it open, or let Windows Task Scheduler open it before RUN_AT)."
    t.Add "Weekends and the holidays listed on CONFIG (column I) are skipped. During a TERMINAL run (about 15 s) do not touch keyboard or mouse; the screen must be unlocked."
    t.Add "Buttons: Run Now (manual run), Import File (backfill old exports), Abort (reset a stuck run), Run Self-Tests (checks the counting logic)."
    t.Add ""
    t.Add "#4. Sheets"
    t.Add "DASHBOARD  KPIs, status banner, distribution of counterparties per deal, deal lookups, exceptions, recent history."
    t.Add "DEALS      latest run, one row per deal (TradeKey prefix RFQ:/TKT:/CMP:/ORD: shows how the deal was grouped)."
    t.Add "HIST       audit trail, one row per (TradeDate, TradeKey); re-running a day updates rows, never duplicates. Column Source = archive file."
    t.Add "HIST_DAILY one row per trade date (recomputed from HIST).   RAW = latest raw export.   LOG = every run, alert, timeout and error."
    t.Add ""
    t.Add "#5. What the alerts mean"
    t.Add "'< 3'      fewer than MIN_Q distinct counterparties were asked on that deal. Investigate and document."
    t.Add "'KEY?'     rows of one deal carry different ISINs or sides: the grouping is ambiguous. Check the export / HDR_* mapping."
    t.Add "TIMEOUT    the export was not detected within TIMEOUT_S. Check the Terminal, EXPORT_SEQ and screen lock."
    t.Add "Coverage gap   a business day has no data in HIST_DAILY. Backfill it with Import File (or check BLOT -1 semantics, step 2g)."
    t.Add "No executed trades   the export was empty. Confirm that the desk did not trade that day."
    t.Add ""
    t.Add "#6. Limits of the control"
    t.Add "Only trades in the BLOT are covered: voice / off-platform trades need a separate source."
    t.Add "The dealer column must contain the counterparties ASKED, not the quotes received (that would create false alerts)."
    t.Add "Open policy points: GROUP_LEVEL (RFQ or ORDER), holiday calendar, scheduled delivery of the BLOT view by Bloomberg (option A)."
    For i = 1 To t.Count
        s = CStr(t(i))
        If Left$(s, 1) = "#" Then
            ws.Cells(i + 1, 2).Value = Mid$(s, 2)
            ws.Cells(i + 1, 2).Font.Bold = True
            ws.Cells(i + 1, 2).Font.Size = IIf(i = 1, 15, 12)
        ElseIf Len(s) > 0 Then
            ws.Cells(i + 1, 2).Value = "'" & s
        End If
    Next
    ws.Columns(1).ColumnWidth = 2
    ws.Columns(2).ColumnWidth = 160
End Sub

'==========================================================================================
'                                    SCHEDULER
'==========================================================================================
Public Sub ScheduleDaily()
    Dim t As Date
    On Error GoTo EH
    LoadConfig
    CancelSchedule
    If Not ParseTimeStr(RUN_AT, t) Then Notify "Scheduler NOT armed: RUN_AT '" & RUN_AT & "' is not hh:mm[:ss].", LV_ERROR: Exit Sub
    gNextRun = Date + t
    If gNextRun <= Now Then gNextRun = gNextRun + 1
    Application.OnTime gNextRun, ProcName("ScheduledRun")
    gSchedPending = True
    WriteLog "Scheduler armed: next run " & Format$(gNextRun, "yyyy-mm-dd hh:nn:ss") & IIf(IsBusinessDay(gNextRun), "", " (non-business day: will be skipped)"), LV_INFO
    Exit Sub
EH:
    Notify "ERROR in ScheduleDaily: " & Err.Description, LV_ERROR
End Sub

Public Sub ScheduledRun()
    gSchedPending = False
    On Error Resume Next
    LoadConfig
    If IsBusinessDay(Date) Then
        RunBlotReport
    Else
        WriteLog "Scheduled run skipped: " & Format$(Date, "yyyy-mm-dd") & " is not a business day.", LV_INFO
    End If
    ScheduleDaily
End Sub

Public Sub CancelSchedule()
    On Error Resume Next
    If gSchedPending Then Application.OnTime gNextRun, ProcName("ScheduledRun"), , False
    gSchedPending = False
End Sub

'==========================================================================================
'                                    RUN STATE
'==========================================================================================
Private Function BeginRun(ByVal kind As String) As Boolean
    If gBusy Then
        If Now > gDeadline + TimeSerial(0, 10, 0) Then
            Notify "Previous " & gKind & " run looked stuck (deadline " & Format$(gDeadline, "hh:nn:ss") & "): state reset.", LV_WARN
            CancelPending
        Else
            Notify "A " & gKind & " run is already in progress; request ignored (use Abort to reset).", LV_WARN, True, True
            Exit Function
        End If
    End If
    gBusy = True: gKind = kind: gLastErr = "": gWarned = "|"
    gDeadline = Now + TimeSerial(0, 0, TIMEOUT_S)
    BeginRun = True
End Function

Private Sub FinishRun()
    CancelPending
    gBusy = False: gKind = ""
    Application.StatusBar = False
End Sub

Private Sub CancelPending()
    On Error Resume Next
    If gPollPending Then Application.OnTime gPollAt, ProcName("PollExport"), , False
    If gCyclePending Then Application.OnTime gCycleAt, ProcName("StartTestCycle"), , False
    gPollPending = False: gCyclePending = False
End Sub

Private Sub SchedulePoll()
    gPollAt = Now + TimeSerial(0, 0, POLL_S)
    Application.OnTime gPollAt, ProcName("PollExport")
    gPollPending = True
End Sub

Private Function ProcName(ByVal p As String) As String
    ProcName = "'" & ThisWorkbook.Name & "'!" & p
End Function

' Drives the source (TERMINAL steps) and arms the poll. False = could not start (already logged).
Private Function StartAcquisition() As Boolean
    gT0 = Now: gDeadline = gT0 + TimeSerial(0, 0, TIMEOUT_S)
    gCandKey = "": gCandRows = -1: gLastErr = ""
    gKnown = SnapshotWorkbooks()
    Select Case SOURCE_MODE
        Case "TERMINAL"
            gSince = gT0
            If Len(EXPORT_SEQ) = 0 Then Notify "EXPORT_SEQ is empty: configure the export steps on CONFIG (Capture Click Point / Test Export).", LV_ERROR: Exit Function
            If Not RunSteps(ExpandSeq(OPEN_SEQ, Date)) Then Notify "OPEN_SEQ failed: could not drive the Bloomberg Terminal.", LV_ERROR: Exit Function
            If Not RunSteps(ExpandSeq(EXPORT_SEQ, Date)) Then Notify "EXPORT_SEQ failed.", LV_ERROR: Exit Function
        Case "OUTLOOK"
            gSince = Date
        Case "FOLDER"
            gSince = Date
            If Len(EXPORT_FOLDER) = 0 Then Notify "EXPORT_FOLDER is empty (FOLDER mode).", LV_ERROR: Exit Function
        Case Else
            Notify "Invalid SOURCE_MODE: " & SOURCE_MODE, LV_ERROR: Exit Function
    End Select
    Application.StatusBar = "BLOT: waiting for export (" & SOURCE_MODE & ")..."
    SchedulePoll
    StartAcquisition = True
End Function

Private Sub OnAcquireTimeout()
    Dim msg As String
    msg = "TIMEOUT: BLOT export not detected after " & TIMEOUT_S & "s (" & SOURCE_MODE & ")."
    If Len(gLastErr) > 0 Then msg = msg & " Last error: " & gLastErr
    If gKind = "TEST" Then
        TestCycleDone False, msg
    Else
        FinishRun
        Notify msg, LV_ALERT
    End If
End Sub

Private Sub HandleExport(v As Variant, vRaw As Variant, ByVal label As String, ByVal srcWb As Object, ByVal srcFile As String)
    If gKind = "TEST" Then
        TestCycleDone True, label, v
    Else
        ProcessData v, vRaw, label, srcWb, srcFile
        FinishRun
    End If
End Sub

Private Sub TestCycleDone(ByVal ok As Boolean, ByVal detail As String, Optional v As Variant)
    Dim res As Object
    CancelPending
    If ok Then
        gTestOK = gTestOK + 1
        If Not IsMissing(v) Then
            Set res = ParseExport(v, LoadDealerMap(), PrevBusinessDay(Date))
            detail = detail & " | parse: " & res("m") & " deal(s), " & res("nB") & " below " & MIN_Q & ", mode=" & res("mode") & _
                     vbLf & "Header mapping: " & res("mapping")
            If Len(res("warnings")) > 0 Then detail = detail & vbLf & "Data quality: " & res("warnings")
        End If
    End If
    Notify "TestExport cycle " & gTestI & "/" & gTestN & ": " & IIf(ok, "OK - export detected in ", "FAIL - ") & detail, IIf(ok, LV_INFO, LV_WARN), False, True
    If gTestI < gTestN Then
        gCycleAt = Now + TimeSerial(0, 0, 3)
        Application.OnTime gCycleAt, ProcName("StartTestCycle")
        gCyclePending = True
    Else
        FinishRun
        Notify "TestExport finished: " & gTestOK & "/" & gTestN & " cycle(s) succeeded." & _
               IIf(gTestOK = gTestN And gTestN >= 5, " EXPORT_SEQ can be accepted.", " Do NOT accept EXPORT_SEQ until 5/5 succeed."), _
               IIf(gTestOK = gTestN, LV_INFO, LV_WARN), True, True
    End If
End Sub

'==========================================================================================
'                                    SOURCES
'==========================================================================================
Private Function TryAcquire() As Boolean
    Select Case SOURCE_MODE
        Case "TERMINAL"
            If TryNewWorkbooks() Then TryAcquire = True: Exit Function
            If Len(EXPORT_FOLDER) > 0 Then TryAcquire = TryFile(NewestFileSince(EXPORT_FOLDER, gSince), True)
        Case "FOLDER"
            TryAcquire = TryFile(NewestFileSince(EXPORT_FOLDER, gSince), True)
        Case "OUTLOOK"
            TryAcquire = TryFile(FetchOutlookAttachment(), False)
    End Select
End Function

Private Function TryNewWorkbooks() As Boolean
    Dim app As Object, wb As Object, blk As Object, k As String
    For Each app In AllExcelApps()
        For Each wb In app.Workbooks
            k = "|" & app.Hwnd & ":" & wb.FullName & "|"
            If InStr(1, gKnown, k, vbTextCompare) = 0 Then
                Set blk = FindBlockInWorkbook(wb)
                If Not blk Is Nothing Then
                    If IsStable(k, blk("rows")) Then
                        gKnown = gKnown & Mid$(k, 2)
                        HandleExport blk("v"), blk("raw"), "workbook " & blk("label"), wb, ""
                        If CLOSE_EXPORT Or gKind = "TEST" Then CloseExport wb, app
                        TryNewWorkbooks = True
                    End If
                    Exit Function                       ' found but still populating -> next poll
                ElseIf InStr(1, gWarned, k, vbTextCompare) = 0 Then
                    gWarned = gWarned & Mid$(k, 2)
                    WriteLog "New workbook " & wb.Name & " ignored: no header row with a dealer column (HDR_DEALERS) in the first 40 rows.", LV_INFO
                End If
            End If
        Next
    Next
End Function

Private Function TryFile(ByVal f As String, ByVal needStable As Boolean) As Boolean
    Dim wb As Workbook, blk As Object
    If Len(f) = 0 Then Exit Function
    If InStr(1, gKnown, "|" & f & "|", vbTextCompare) > 0 Then Exit Function
    If needStable Then
        If Not IsStable("F|" & f, FileLen(f)) Then Exit Function
    End If
    If IsCsvPath(f) Then
        Set blk = CsvBlock(f)
    Else
        Set wb = Workbooks.Open(f, UpdateLinks:=0, ReadOnly:=True)
        Set blk = FindBlockInWorkbook(wb)
    End If
    gKnown = gKnown & f & "|"
    If blk Is Nothing Then
        Notify "File " & f & " ignored: no header row with a dealer column (HDR_DEALERS).", LV_WARN
    Else
        HandleExport blk("v"), blk("raw"), "file " & f, Nothing, f
        TryFile = True
    End If
    If Not wb Is Nothing Then wb.Close False
End Function

Private Function FetchOutlookAttachment() As String
    Dim ol As Object, fld As Object, its As Object, it As Object, att As Object, p As String, fn As String
    On Error GoTo Done
    Set ol = CreateObject("Outlook.Application")
    Set fld = OutlookFolder(ol, OUTLOOK_FOLDER)
    If fld Is Nothing Then gLastErr = "Outlook folder '" & OUTLOOK_FOLDER & "' not found": Exit Function
    Set its = fld.Items
    its.Sort "[ReceivedTime]", True
    Set it = its.GetFirst
    Do While Not it Is Nothing
        If it.Class = 43 Then
            If it.ReceivedTime < gSince Then Exit Do
            If InStr(1, it.Subject, OUTLOOK_SUBJECT, vbTextCompare) > 0 Then
                If InStr(1, gKnown, "|" & it.EntryID & "|", vbTextCompare) = 0 Then
                    For Each att In it.Attachments
                        fn = LCase$(att.FileName)
                        If fn Like "*.xls*" Or fn Like "*.csv" Then
                            p = Environ$("TEMP") & "\BLOT_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & att.FileName
                            att.SaveAsFile p
                            gKnown = gKnown & it.EntryID & "|"
                            FetchOutlookAttachment = p
                            Exit Function
                        End If
                    Next
                End If
            End If
        End If
        Set it = its.GetNext
    Loop
Done:
    If Err.Number <> 0 Then gLastErr = "Outlook: " & Err.Description
End Function

Private Function NewestFileSince(ByVal folder As String, ByVal since As Date) As String
    Dim f As String, best As String, bt As Date, ft As Date, pat As Variant
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    For Each pat In Array("*.xls*", "*.csv")
        f = Dir$(folder & pat)
        Do While Len(f) > 0
            If Left$(f, 2) <> "~$" Then
                ft = FileDateTime(folder & f)
                If ft >= since And ft > bt Then bt = ft: best = folder & f
            End If
            f = Dir$()
        Loop
    Next
    NewestFileSince = best
End Function

Private Sub CloseExport(ByVal wb As Object, ByVal app As Object)
    On Error Resume Next
    wb.Close SaveChanges:=False
    If Not app Is Application Then
        If app.Workbooks.Count = 0 Then app.Quit
    End If
End Sub

' All running Excel instances (Bloomberg export often opens a separate instance)
Private Function AllExcelApps() As Collection
    Dim col As New Collection, seen As Object, hMain As LongPtr, hDesk As LongPtr, h7 As LongPtr
    Dim w As Object, iid As GUID_T, app As Object
    Set seen = NewDict()
    col.Add Application: seen(CStr(Application.Hwnd)) = 1
    IIDFromString StrPtr("{00020400-0000-0000-C000-000000000046}"), iid
    hMain = FindWindowEx(0, 0, "XLMAIN", vbNullString)
    Do While hMain <> 0
        hDesk = FindWindowEx(hMain, 0, "XLDESK", vbNullString)
        h7 = 0
        If hDesk <> 0 Then h7 = FindWindowEx(hDesk, 0, "EXCEL7", vbNullString)
        If h7 <> 0 Then
            Set w = Nothing: Set app = Nothing
            On Error Resume Next
            If AccessibleObjectFromWindow(h7, OBJID_NATIVEOM, iid, w) = 0 Then Set app = w.Application
            If Not app Is Nothing Then
                If Not seen.Exists(CStr(app.Hwnd)) Then seen(CStr(app.Hwnd)) = 1: col.Add app
            End If
            On Error GoTo 0
        End If
        hMain = FindWindowEx(0, hMain, "XLMAIN", vbNullString)
    Loop
    Set AllExcelApps = col
End Function

Private Function SnapshotWorkbooks() As String
    Dim app As Object, wb As Object, s As String
    s = "|"
    On Error Resume Next
    For Each app In AllExcelApps()
        For Each wb In app.Workbooks
            s = s & app.Hwnd & ":" & wb.FullName & "|"
        Next
    Next
    SnapshotWorkbooks = s
End Function

Private Function IsStable(ByVal key As String, ByVal n As Long) As Boolean
    If key = gCandKey And n = gCandRows Then IsStable = True Else gCandKey = key: gCandRows = n
End Function

Private Function IsCsvPath(ByVal f As String) As Boolean
    IsCsvPath = (LCase$(Right$(f, 4)) = ".csv" Or LCase$(Right$(f, 4)) = ".txt")
End Function

' CSV is read as TEXT (no Excel conversion): avoids the PT-locale date/number misparse of Workbooks.Open
Private Function CsvBlock(ByVal f As String) As Object
    Dim a As Variant, blk As Variant, d As Object
    a = ParseCsvText(ReadTextFile(f), CSV_DELIM)
    blk = FindBlockInArray(a)
    If IsEmpty(blk) Then Exit Function
    Set d = NewDict()
    d("v") = blk: d("raw") = blk: d("rows") = UBound(blk, 1): d("label") = Dir$(f)
    Set CsvBlock = d
End Function

'==========================================================================================
'                                    TERMINAL DRIVER
'==========================================================================================
Public Function ExpandSeq(ByVal seq As String, ByVal refDate As Date) As String
    ExpandSeq = Replace(seq, "{DAYS_BACK}", CStr(DaysBack(refDate)), , , vbTextCompare)
End Function

' "" if the step sequence is syntactically valid, else the problem
Public Function ValidateSteps(ByVal seq As String) As String
    Dim st As Variant, s As String, tok As String, arg As String, p As Long, xy() As String
    For Each st In Split(seq, "|")
        s = Trim$(CStr(st))
        If Len(s) > 0 Then
            p = InStr(1, s, ":")
            If p = 0 Then ValidateSteps = "Bad step '" & s & "' (expected TOKEN:arg).": Exit Function
            tok = UCase$(Left$(s, p - 1)): arg = Mid$(s, p + 1)
            Select Case tok
                Case "KEYS", "SEND"
                    If Len(arg) = 0 Then ValidateSteps = "Empty " & tok & " step.": Exit Function
                    If tok = "KEYS" And (UCase$(Trim$(arg)) Like "[A-Z]<GO>" Or InStr(1, arg, "<N>", vbTextCompare) > 0) Then
                        ValidateSteps = "'" & s & "' looks like an unfilled placeholder: a single letter + <GO> runs a Bloomberg function (N = News). Use Record Export.": Exit Function
                    End If
                Case "CLICK", "RCLICK", "DCLICK"
                    xy = Split(arg, ",")
                    If UBound(xy) <> 1 Then ValidateSteps = "Bad " & s & " (expected " & tok & ":x,y).": Exit Function
                    If Not IsPlainNumber(Replace(Trim$(xy(0)), "-", "")) Or Not IsPlainNumber(Replace(Trim$(xy(1)), "-", "")) Then ValidateSteps = "Bad coordinates in " & s & ".": Exit Function
                Case "WAIT"
                    If Not IsPlainNumber(arg) Then ValidateSteps = "Bad " & s & " (seconds, '.' decimal).": Exit Function
                Case Else
                    ValidateSteps = "Unknown step '" & tok & "' in '" & s & "'.": Exit Function
            End Select
        End If
    Next
End Function

Private Function RunSteps(ByVal seq As String) As Boolean
    Dim st As Variant, s As String, tok As String, arg As String, p As Long, e As String
    e = ValidateSteps(seq)
    If Len(e) > 0 Then Notify "Step sequence rejected: " & e, LV_ERROR: Exit Function
    gWinChecked = False
    For Each st In Split(seq, "|")
        s = Trim$(CStr(st))
        If Len(s) > 0 Then
            p = InStr(1, s, ":")
            tok = UCase$(Left$(s, p - 1)): arg = Mid$(s, p + 1)
            Select Case tok
                Case "KEYS"
                    If Not BbgType(arg) Then Notify "Step failed: " & s & " (Bloomberg window '" & BBG_TITLE_PREFIX & "' not reachable).", LV_ERROR: Exit Function
                Case "SEND"
                    If Not BbgFocusUnlessOwn() Then Notify "Step failed: " & s & " (Bloomberg window '" & BBG_TITLE_PREFIX & "' not found: use List Windows / Record Export).", LV_ERROR: Exit Function
                    SendKeys arg, True
                Case "CLICK", "RCLICK", "DCLICK"
                    If Not BbgClick(arg, tok) Then Notify "Step failed: " & s & " (Bloomberg window '" & BBG_TITLE_PREFIX & "' not found: use List Windows / Record Export).", LV_ERROR: Exit Function
                Case "WAIT"
                    WaitSec Val(arg)
            End Select
            WaitSec STEP_GAP_S
        End If
    Next
    RunSteps = True
End Function

Private Function BbgType(ByVal txt As String) As Boolean
    Dim ch As Long, da As Boolean
    da = Application.DisplayAlerts: Application.DisplayAlerts = False
    On Error GoTo DDEFail
    ch = Application.DDEInitiate("WINBLP", "BBK")
    Application.DDEExecute ch, BBG_PANEL & txt
    Application.DDETerminate ch
    Application.DisplayAlerts = da
    BbgType = True: Exit Function
DDEFail:
    Resume UseKeys
UseKeys:
    Application.DisplayAlerts = da
    On Error GoTo Fail
    If Not BbgFocus() Then GoTo Fail
    SendKeys Replace(EscKeys(txt), "<GO>", "~"), True
    BbgType = True: Exit Function
Fail:
    BbgType = False
End Function

Private Function BbgClick(ByVal arg As String, ByVal kind As String) As Boolean
    Dim h As LongPtr, rc As RECT_T, xy() As String
    xy = Split(arg, ",")
    If UBound(xy) <> 1 Then Exit Function
    If Not BbgFocusUnlessOwn() Then Exit Function
    h = BbgHwnd()
    If Not gWinChecked Then CheckBbgWindow h: gWinChecked = True
    GetWindowRect h, rc
    SetCursorPos rc.Left + CLng(Val(xy(0))), rc.Top + CLng(Val(xy(1)))
    WaitSec 0.2
    Select Case kind
        Case "CLICK":  MouseBtn MOUSE_LDOWN, MOUSE_LUP
        Case "RCLICK": MouseBtn MOUSE_RDOWN, MOUSE_RUP
        Case "DCLICK": MouseBtn MOUSE_LDOWN, MOUSE_LUP: WaitSec 0.08: MouseBtn MOUSE_LDOWN, MOUSE_LUP
    End Select
    BbgClick = True
End Function

Private Sub MouseBtn(ByVal dn As Long, ByVal up As Long)
    mouse_event dn, 0, 0, 0, 0
    WaitSec 0.05
    mouse_event up, 0, 0, 0, 0
End Sub

Public Function BbgHwnd() As LongPtr
    Dim h As LongPtr, s As String, l As Long
    If Len(BBG_TITLE_PREFIX) = 0 Then Exit Function
    h = FindWindowEx(0, 0, vbNullString, vbNullString)
    Do While h <> 0
        If IsWindowVisible(h) <> 0 Then
            s = String$(260, vbNullChar): l = GetWindowText(h, s, 260)
            If l > 0 Then
                If UCase$(Left$(Left$(s, l), Len(BBG_TITLE_PREFIX))) = UCase$(BBG_TITLE_PREFIX) Then BbgHwnd = h: Exit Function
            End If
        End If
        h = FindWindowEx(0, h, vbNullString, vbNullString)
    Loop
End Function

Private Function BbgFocus() As Boolean
    Dim h As LongPtr, s As String, l As Long
    h = BbgHwnd(): If h = 0 Then Exit Function
    If IsIconic(h) <> 0 Then ShowWindow h, SW_RESTORE
    s = String$(260, vbNullChar): l = GetWindowText(h, s, 260)
    On Error Resume Next
    AppActivate Left$(s, l)
    On Error GoTo 0
    SetForegroundWindow h
    WaitSec 0.5
    BbgFocus = True
End Function

Private Function EscKeys(ByVal s As String) As String
    Dim i As Long, c As String, o As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If InStr(1, "+^%~(){}[]", c) > 0 Then o = o & "{" & c & "}" Else o = o & c
    Next
    EscKeys = o
End Function

'==========================================================================================
'                                    READING EXPORTS
'==========================================================================================
' First sheet with a header row (first 40 rows, >= 3 cells, contains a dealer alias).
' Returns Dictionary: v (Value2 block, parsed), raw (Value block, for RAW), rows, label
Private Function FindBlockInWorkbook(ByVal wb As Object) As Object
    Dim ws As Object, ur As Object, a As Variant, r As Long, lc As Long, d As Object
    If wb Is ThisWorkbook Then Exit Function
    For Each ws In wb.Worksheets
        Set ur = ws.UsedRange
        If ur.Columns.Count >= 3 Then
            a = ur.Value2
            If IsArray(a) Then
                r = FindHeaderRow(a, lc)
                If r > 0 Then
                    Set d = NewDict()
                    d("v") = SliceBlock(a, r, lc)
                    d("raw") = ur.Cells(r, 1).Resize(UBound(a, 1) - r + 1, lc).Value
                    d("rows") = UBound(a, 1) - r + 1
                    d("label") = wb.Name & "!" & ws.Name
                    Set FindBlockInWorkbook = d
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Function FindBlockInArray(a As Variant) As Variant
    Dim r As Long, lc As Long
    If Not IsArray(a) Then Exit Function
    r = FindHeaderRow(a, lc)
    If r > 0 Then FindBlockInArray = SliceBlock(a, r, lc)
End Function

Private Function FindHeaderRow(a As Variant, ByRef lastC As Long) As Long
    Dim r As Long, c As Long, lc As Long, rMax As Long, hdr() As Variant
    rMax = UBound(a, 1): If rMax > 40 Then rMax = 40
    For r = 1 To rMax
        lc = 0
        For c = UBound(a, 2) To 1 Step -1
            If Len(CellS(a, r, c)) > 0 Then lc = c: Exit For
        Next
        If lc >= 3 Then
            ReDim hdr(1 To 1, 1 To lc)
            For c = 1 To lc: hdr(1, c) = a(r, c): Next
            If HeaderIndex(hdr, HDR_DEALERS) > 0 Then FindHeaderRow = r: lastC = lc: Exit Function
        End If
    Next
End Function

Private Function SliceBlock(a As Variant, ByVal r0 As Long, ByVal nc As Long) As Variant
    Dim o() As Variant, i As Long, j As Long, n As Long
    n = UBound(a, 1) - r0 + 1
    ReDim o(1 To n, 1 To nc)
    For i = 1 To n
        For j = 1 To nc
            o(i, j) = a(r0 + i - 1, j)
        Next
    Next
    SliceBlock = o
End Function

' Reads a text file: UTF-8 (BOM or valid), UTF-16 LE (BOM) or Windows-1252
Public Function ReadTextFile(ByVal p As String) As String
    Dim st As Object, b() As Byte, cs As String, s As String, n As Long
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1: st.Open: st.LoadFromFile p
    n = st.Size
    If n = 0 Then st.Close: Exit Function
    b = st.Read
    st.Close
    cs = "windows-1252"
    If n >= 3 Then
        If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then cs = "utf-8"
    End If
    If n >= 2 And cs = "windows-1252" Then
        If b(0) = &HFF And b(1) = &HFE Then cs = "unicode"
    End If
    If cs = "windows-1252" Then
        If LooksUtf8(b) Then cs = "utf-8"
    End If
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = cs: st.Open: st.LoadFromFile p
    s = st.ReadText(-1)
    st.Close
    If Len(s) > 0 Then
        If Left$(s, 1) = ChrW(&HFEFF) Then s = Mid$(s, 2)
    End If
    ReadTextFile = s
End Function

Private Function LooksUtf8(b() As Byte) As Boolean
    Dim i As Long, n As Long, k As Long, j As Long, hasHi As Boolean
    i = LBound(b): n = UBound(b)
    Do While i <= n
        If b(i) < &H80 Then
            k = 0
        ElseIf (b(i) And &HE0) = &HC0 Then
            k = 1
        ElseIf (b(i) And &HF0) = &HE0 Then
            k = 2
        ElseIf (b(i) And &HF8) = &HF0 Then
            k = 3
        Else
            Exit Function
        End If
        If k > 0 Then
            hasHi = True
            If i + k > n Then Exit Function
            For j = 1 To k
                If (b(i + j) And &HC0) <> &H80 Then Exit Function
            Next
        End If
        i = i + k + 1
    Loop
    LooksUtf8 = hasHi
End Function

Public Function CsvDelimiter(ByVal txt As String, ByVal cfg As String) As String
    Dim lines As Variant, i As Long, c As Long, cand As Variant, best As String, bestN As Long, mx As Long, n As Long
    Select Case UCase$(Trim$(cfg))
        Case ";": CsvDelimiter = ";": Exit Function
        Case ",": CsvDelimiter = ",": Exit Function
        Case "TAB": CsvDelimiter = vbTab: Exit Function
    End Select
    txt = Replace(Replace(txt, vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(txt, vbLf, 21)
    best = ",": bestN = 0
    For Each cand In Array(";", ",", vbTab)
        mx = 0
        For i = LBound(lines) To Application.WorksheetFunction.Min(UBound(lines), LBound(lines) + 19)
            n = Len(lines(i)) - Len(Replace(lines(i), cand, ""))
            If n > mx Then mx = n
        Next
        If mx > bestN Then bestN = mx: best = cand
    Next
    CsvDelimiter = best
End Function

' RFC-4180 style parser: quoted fields, "" escapes, line breaks inside quotes. All values are text.
Public Function ParseCsvText(ByVal txt As String, ByVal delimCfg As String) As Variant
    Dim d As String, n As Long, i As Long, c As String, inQ As Boolean, fld As String
    Dim rowF() As String, nf As Long, rows As Collection, maxC As Long, a() As Variant, r As Long, j As Long, rw As Variant
    d = CsvDelimiter(txt, delimCfg)
    txt = Replace(Replace(txt, vbCrLf, vbLf), vbCr, vbLf)
    Set rows = New Collection
    ReDim rowF(0 To 31)
    n = Len(txt): i = 1
    Do While i <= n
        c = Mid$(txt, i, 1)
        If inQ Then
            If c = """" Then
                If Mid$(txt, i + 1, 1) = """" Then
                    fld = fld & """": i = i + 1
                Else
                    inQ = False
                End If
            Else
                fld = fld & c
            End If
        ElseIf c = """" And Len(fld) = 0 Then
            inQ = True
        ElseIf c = d Then
            PushField rowF, nf, fld: fld = ""
        ElseIf c = vbLf Then
            PushField rowF, nf, fld: fld = ""
            PushRow rows, rowF, nf, maxC
        Else
            fld = fld & c
        End If
        i = i + 1
    Loop
    If nf > 0 Or Len(fld) > 0 Then PushField rowF, nf, fld: PushRow rows, rowF, nf, maxC
    If rows.Count = 0 Then Exit Function
    ReDim a(1 To rows.Count, 1 To maxC)
    r = 0
    For Each rw In rows
        r = r + 1
        For j = 1 To maxC
            If j - 1 <= UBound(rw) Then a(r, j) = rw(j - 1) Else a(r, j) = ""
        Next
    Next
    ParseCsvText = a
End Function

Private Sub PushField(rowF() As String, ByRef nf As Long, ByVal fld As String)
    If nf > UBound(rowF) Then ReDim Preserve rowF(0 To 2 * UBound(rowF) + 1)
    rowF(nf) = fld: nf = nf + 1
End Sub

Private Sub PushRow(rows As Collection, rowF() As String, ByRef nf As Long, ByRef maxC As Long)
    Dim o() As String, j As Long
    ReDim o(0 To nf - 1)
    For j = 0 To nf - 1: o(j) = rowF(j): Next
    rows.Add o
    If nf > maxC Then maxC = nf
    nf = 0
End Sub

'==========================================================================================
'                                    PARSING CORE (pure: array in, result out)
'==========================================================================================
' v       : 2-D array, row 1 = header (1-based)
' dmap    : normalized alias -> normalized legal entity
' defTD   : trade date used when a deal has no parseable trade date
' Returns a Dictionary: ok, err, m, deals (m x NCOL), mode, mapping, warnings, nB, nBad,
'   nDropped, dropped, nSkipped, nDateFallback, unmapped, keykinds, minTD, maxTD
' Deal columns: 1 TradeDate | 2 TradeKey | 3 Security | 4 Side | 5 Qty | 6 NDealers | 7 Status
'               8 LoadedAt (wrapper) | 9 Dealers (sorted, ';') | 10 Source (wrapper)
Public Function ParseExport(v As Variant, dmap As Object, ByVal defTD As Date) As Object
    Dim res As Object, hdr() As Variant, n As Long, nc As Long, i As Long, j As Long
    Dim cDl As Long, cRfq As Long, cID As Long, cOrd As Long, cSec As Long, cDt As Long
    Dim cSide As Long, cQty As Long, cSt As Long, cTime As Long, cTrd As Long
    Dim isOrd As Boolean, tktOK As Boolean, rk() As String, lastKey As String, k As String, s As String
    Dim mode As String, live As Object, secU As Object, sideU As Object, secR As Object, sideR As Object
    Dim qtyR As Object, dtR As Object, dset As Object, cnt As Object, bad As Object, perKey As Object
    Dim ents As Object, unm As Object, kinds As Object, kk As Variant, out() As Variant
    Dim nNum As Long, nTxt As Long, multi As Boolean, nSkip As Long, nFb As Long, m As Long, r As Long
    Dim nB As Long, nBad As Long, nDrop As Long, dropped As String, warn As String, x As Long, d As Date
    Dim minTD As Date, maxTD As Date, nCmp As Long

    Set res = NewDict()
    res("ok") = False: res("err") = "": res("m") = 0: res("deals") = Empty: res("mode") = ""
    res("mapping") = "": res("warnings") = "": res("nB") = 0: res("nBad") = 0: res("nDropped") = 0
    res("dropped") = "": res("nSkipped") = 0: res("nDateFallback") = 0: res("unmapped") = ""
    res("keykinds") = "": res("minTD") = 0: res("maxTD") = 0
    If Not IsArray(v) Then res("err") = "No data.": GoTo Done
    n = UBound(v, 1): nc = UBound(v, 2)
    ReDim hdr(1 To 1, 1 To nc)
    For j = 1 To nc: hdr(1, j) = v(1, j): Next

    '--- header mapping ---
    cDl = HeaderIndex(hdr, HDR_DEALERS)
    cRfq = HeaderIndex(hdr, HDR_RFQ):     cID = HeaderIndex(hdr, HDR_ID)
    cOrd = HeaderIndex(hdr, HDR_ORDER):   cSec = HeaderIndex(hdr, HDR_SEC)
    cDt = HeaderIndex(hdr, HDR_DATE):     cSide = HeaderIndex(hdr, HDR_SIDE)
    cQty = HeaderIndex(hdr, HDR_QTY):     cSt = HeaderIndex(hdr, HDR_STATUS)
    cTime = HeaderIndex(hdr, HDR_TIME):   cTrd = HeaderIndex(hdr, HDR_TRADER)
    isOrd = (GROUP_LEVEL = "ORDER")
    res("mapping") = MapDesc(hdr, Array("DEALERS", cDl, "RFQ", cRfq, "ORDER", cOrd, "TICKET", cID, "SECURITY", cSec, _
                     "DATE", cDt, "SIDE", cSide, "QTY", cQty, "STATUS", cSt, "TIME", cTime, "TRADER", cTrd))
    warn = Collisions(Array("DEALERS", cDl, "RFQ", cRfq, "ORDER", cOrd, "TICKET", cID, "SECURITY", cSec, _
                     "DATE", cDt, "SIDE", cSide, "QTY", cQty, "STATUS", cSt, "TIME", cTime, "TRADER", cTrd))
    If cDl = 0 Then res("err") = "Dealer column not found (HDR_DEALERS).": GoTo Done
    If isOrd And cOrd = 0 Then warn = warn & "GROUP_LEVEL=ORDER but no order column (HDR_ORDER): RFQ level used. "
    If Not isOrd Then cOrd = 0

    '--- 1) ticket usable only if every dealer row without an RFQ/order id carries it ---
    tktOK = (cID > 0)
    If tktOK Then
        For i = 2 To n
            If Len(CellS(v, i, cDl)) > 0 And Not IsChildRow(v, i, cOrd, cRfq, cID, cSec) Then
                If Len(CellS(v, i, cOrd)) + Len(CellS(v, i, cRfq)) + Len(CellS(v, i, cID)) = 0 Then tktOK = False: Exit For
            End If
        Next
        If Not tktOK Then warn = warn & "Ticket column '" & NormH(hdr(1, cID)) & "' is blank on some dealer rows (ticket on the winning dealer only?): not used as key. "
    End If

    '--- 2) trade key per row; child rows (no RFQ/order/ticket/security) inherit the previous key ---
    ReDim rk(1 To n)
    For i = 2 To n
        If IsChildRow(v, i, cOrd, cRfq, cID, cSec) Then
            If Len(CellS(v, i, cDl)) > 0 Then
                If Len(lastKey) = 0 Then nSkip = nSkip + 1 Else rk(i) = lastKey
            End If
        Else
            k = MakeTradeKey(v, i, cOrd, cRfq, IIf(tktOK, cID, 0), cDt, cSec, cSide, cQty, cTime, cTrd)
            rk(i) = k: lastKey = k
        End If
    Next

    '--- 3) counting mode ---
    mode = DEALER_MODE
    If mode <> "ROWS" And mode <> "LIST" And mode <> "COUNT" Then
        Set perKey = NewDict()
        For i = 2 To n
            If Len(rk(i)) > 0 Then
                s = CellS(v, i, cDl)
                If Len(s) > 0 Then
                    If IsPlainNumber(s) Then nNum = nNum + 1 Else nTxt = nTxt + 1
                    perKey(rk(i)) = perKey(rk(i)) + 1
                    If perKey(rk(i)) > 1 Then multi = True
                End If
            End If
        Next
        If nNum > 0 And nNum >= nTxt Then
            mode = "COUNT"
        ElseIf multi Or nTxt = 0 Then
            mode = "ROWS"
        Else
            mode = "LIST"
        End If
    End If
    res("mode") = mode

    '--- 4) aggregate per key ---
    Set live = NewDict(): Set secU = NewDict(): Set sideU = NewDict(): Set secR = NewDict()
    Set sideR = NewDict(): Set qtyR = NewDict(): Set dtR = NewDict(): Set dset = NewDict()
    Set cnt = NewDict(): Set bad = NewDict(): Set unm = NewDict(): Set kinds = NewDict()
    Set ents = NewDict()
    For Each kk In dmap.Keys: ents(dmap(kk)) = 1: Next
    For i = 2 To n
        k = rk(i)
        If Len(k) > 0 Then
            If Not live.Exists(k) Then live.Add k, False: cnt.Add k, 0: dset.Add k, NewDict()
            ' a deal is live if ANY row is not cancelled/rejected (dealer-level Expired/Passed = still asked)
            If Not IsExcluded(CellS(v, i, cSt)) Then live(k) = True
            ' integrity: one ISIN and one side per key
            s = CellS(v, i, cSec)
            If Len(s) > 0 Then
                If Not secU.Exists(k) Then
                    secU(k) = UCase$(s): secR(k) = s
                ElseIf secU(k) <> UCase$(s) Then
                    bad(k) = 1
                End If
            End If
            s = CellS(v, i, cSide)
            If Len(s) > 0 Then
                If Not sideU.Exists(k) Then
                    sideU(k) = UCase$(s): sideR(k) = s
                ElseIf sideU(k) <> UCase$(s) Then
                    bad(k) = 1
                End If
            End If
            If Not qtyR.Exists(k) Then
                s = CellS(v, i, cQty)
                If Len(s) > 0 Then qtyR(k) = s
            End If
            If Not dtR.Exists(k) And cDt > 0 Then
                If Len(CellS(v, i, cDt)) > 0 Then dtR.Add k, v(i, cDt)
            End If
            ' distinct counterparties, normalized to legal entity
            s = CellS(v, i, cDl)
            Select Case mode
                Case "COUNT"
                    If IsPlainNumber(s) Then
                        x = Fix(Val(s))
                        If x > cnt(k) Then cnt(k) = x
                    End If
                Case Else
                    AddDealers dset(k), s, dmap, ents, unm, (mode = "LIST")
                    cnt(k) = dset(k).Count
            End Select
        End If
    Next

    '--- 5) output ---
    For Each kk In live.Keys
        If live(kk) Then
            m = m + 1
        Else
            nDrop = nDrop + 1
            If nDrop <= 50 Then dropped = dropped & kk & "; "
        End If
    Next
    If m > 0 Then
        ReDim out(1 To m, 1 To NCOL)
        For Each kk In live.Keys
            If live(kk) Then
                r = r + 1
                If dtR.Exists(kk) Then
                    If Not ParseDateCell(dtR(kk), DATE_FORMAT_TEXT, d) Then d = defTD: nFb = nFb + 1
                Else
                    d = defTD: nFb = nFb + 1
                End If
                out(r, 1) = d
                out(r, 2) = CStr(kk)
                If secR.Exists(kk) Then out(r, 3) = secR(kk) Else out(r, 3) = ""
                If sideR.Exists(kk) Then out(r, 4) = sideR(kk) Else out(r, 4) = ""
                If qtyR.Exists(kk) Then out(r, 5) = QtyOut(qtyR(kk)) Else out(r, 5) = ""
                out(r, 6) = CLng(cnt(kk))
                If bad.Exists(kk) Then
                    out(r, 7) = "KEY?": nBad = nBad + 1
                ElseIf cnt(kk) >= MIN_Q Then
                    out(r, 7) = "OK"
                Else
                    out(r, 7) = "< " & MIN_Q
                End If
                If cnt(kk) < MIN_Q Then nB = nB + 1
                If mode <> "COUNT" Then out(r, 9) = SortedJoin(dset(kk)) Else out(r, 9) = ""
                out(r, 8) = Empty: out(r, 10) = ""
                kinds(Left$(CStr(kk), 3)) = kinds(Left$(CStr(kk), 3)) + 1
                If minTD = 0 Or d < minTD Then minTD = d
                If d > maxTD Then maxTD = d
            End If
        Next
        res("deals") = out
    End If

    '--- 6) data-quality warnings ---
    nCmp = kinds("CMP")
    If cRfq = 0 And cOrd = 0 And Not tktOK Then
        warn = warn & "No RFQ / ticket id column: deals grouped by composite key (date|ISIN|side|qty|RFQ time|trader). Map HDR_RFQ to the inquiry id column. "
    ElseIf nCmp > 0 Then
        warn = warn & nCmp & " deal(s) without RFQ/ticket id grouped by composite key. "
    End If
    If nCmp > 0 And cTime = 0 Then warn = warn & "No RFQ time column: re-trades of the same size cannot be separated. "
    If nSkip > 0 Then warn = warn & nSkip & " dealer row(s) before the first deal row could not be attributed (not counted). "
    If nFb > 0 Then warn = warn & nFb & " deal(s) without a parseable trade date: " & Format$(defTD, "yyyy-mm-dd") & " used (check HDR_DATE / DATE_FORMAT_TEXT). "
    If cSt = 0 Then warn = warn & "No status column: every deal treated as executed. "

    res("ok") = True
    res("m") = m: res("nB") = nB: res("nBad") = nBad: res("nDropped") = nDrop: res("dropped") = dropped
    res("nSkipped") = nSkip: res("nDateFallback") = nFb: res("unmapped") = SortedJoin(unm, ", ")
    res("keykinds") = "RFQ=" & kinds("RFQ") + 0 & " ORD=" & kinds("ORD") + 0 & " TKT=" & kinds("TKT") + 0 & " CMP=" & nCmp + 0
    res("minTD") = minTD: res("maxTD") = maxTD
Done:
    res("warnings") = Trim$(warn)
    Set ParseExport = res
End Function

Private Function IsChildRow(v As Variant, ByVal i As Long, ByVal cOrd As Long, ByVal cRfq As Long, ByVal cID As Long, ByVal cSec As Long) As Boolean
    IsChildRow = (Len(CellS(v, i, cOrd)) + Len(CellS(v, i, cRfq)) + Len(CellS(v, i, cID)) + Len(CellS(v, i, cSec)) = 0)
End Function

Private Function MakeTradeKey(v As Variant, ByVal i As Long, ByVal cOrd As Long, ByVal cRfq As Long, ByVal cID As Long, _
                              ByVal cDt As Long, ByVal cSec As Long, ByVal cSide As Long, ByVal cQty As Long, _
                              ByVal cTime As Long, ByVal cTrd As Long) As String
    Dim s As String, d As Date, ds As String
    s = CellS(v, i, cOrd): If Len(s) > 0 Then MakeTradeKey = "ORD:" & s: Exit Function
    s = CellS(v, i, cRfq): If Len(s) > 0 Then MakeTradeKey = "RFQ:" & s: Exit Function
    s = CellS(v, i, cID):  If Len(s) > 0 Then MakeTradeKey = "TKT:" & s: Exit Function
    If cDt > 0 Then
        If ParseDateCell(v(i, cDt), DATE_FORMAT_TEXT, d) Then ds = Format$(d, "yyyy-mm-dd") Else ds = CellS(v, i, cDt)
    End If
    MakeTradeKey = "CMP:" & Join(Array(ds, UCase$(CellS(v, i, cSec)), UCase$(CellS(v, i, cSide)), _
                                       CellS(v, i, cQty), TimeKey(v, i, cTime), UCase$(CellS(v, i, cTrd))), "|")
End Function

Private Function TimeKey(v As Variant, ByVal i As Long, ByVal c As Long) As String
    Dim x As Variant
    If c = 0 Then Exit Function
    x = v(i, c)
    Select Case VarType(x)
        Case vbDouble, vbSingle, vbDate, vbCurrency, vbDecimal
            If CDbl(x) >= 0 And CDbl(x) < 1 Then
                TimeKey = Format$(CDate(x), "hh:nn:ss")
            ElseIf CDbl(x) >= 1 And CDbl(x) < 2958466 Then
                TimeKey = Format$(CDate(x), "yyyy-mm-dd hh:nn:ss")
            Else
                TimeKey = CellS(v, i, c)
            End If
        Case Else
            TimeKey = CellS(v, i, c)
    End Select
End Function

Private Sub AddDealers(d As Object, ByVal s As String, dmap As Object, ents As Object, unm As Object, ByVal splitList As Boolean)
    Dim i As Long, t As Variant, u As String
    If Len(s) = 0 Then Exit Sub
    If splitList Then
        For i = 1 To Len(LIST_DELIMS): s = Replace(s, Mid$(LIST_DELIMS, i, 1), Chr$(1)): Next
        s = Replace(Replace(s, vbCr, Chr$(1)), vbLf, Chr$(1))
    End If
    For Each t In Split(s, Chr$(1))
        u = NormName(CStr(t))
        If Len(u) > 0 Then
            If dmap.Exists(u) Then
                u = dmap(u)
            ElseIf Not ents.Exists(u) Then
                unm(u) = 1
            End If
            d(u) = 1
        End If
    Next
End Sub

' UPPERCASE, tabs / NBSP / line breaks -> space, repeated spaces collapsed, trimmed
Public Function NormName(ByVal s As String) As String
    s = Replace(Replace(Replace(Replace(s, vbTab, " "), Chr$(160), " "), vbCr, " "), vbLf, " ")
    Do While InStr(1, s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    NormName = UCase$(Trim$(s))
End Function

Public Function LoadDealerMap() As Object
    Dim d As Object, ws As Worksheet, mv As Variant, i As Long, lr As Long, a As String, e As String
    Set d = NewDict()
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(SH_MAP): On Error GoTo 0
    If Not ws Is Nothing Then
        lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lr >= 2 Then
            mv = ws.Range("A1:B" & lr).Value
            For i = 2 To UBound(mv, 1)
                If Not IsError(mv(i, 1)) And Not IsError(mv(i, 2)) Then
                    a = NormName(CStr(mv(i, 1))): e = NormName(CStr(mv(i, 2)))
                    If Len(a) > 0 And Len(e) > 0 Then
                        If Not d.Exists(a) Then d(a) = e
                    End If
                End If
            Next
        End If
    End If
    Set LoadDealerMap = d
End Function

Private Function IsExcluded(ByVal st As String) As Boolean
    Dim a As Variant, t As String
    If Len(st) = 0 Then Exit Function
    st = LCase$(st)
    For Each a In Split(EXCL_STATUS, "|")
        t = Trim$(CStr(a))
        If Len(t) > 0 Then
            If InStr(1, st, t, vbBinaryCompare) > 0 Then IsExcluded = True: Exit Function
        End If
    Next
End Function

Public Function HeaderIndex(hdr As Variant, ByVal aliases As String) As Long
    Dim a As Variant, t As String, j As Long
    For Each a In Split(aliases, "|")                         ' 1) exact match, by priority
        t = NormHeaderText(CStr(a))
        If Len(t) > 0 Then
            For j = LBound(hdr, 2) To UBound(hdr, 2)
                If NormH(hdr(1, j)) = t Then HeaderIndex = j: Exit Function
            Next
        End If
    Next
    For Each a In Split(aliases, "|")                         ' 2) contains, only aliases >= 5 chars
        t = NormHeaderText(CStr(a))
        If Len(t) >= 5 Then
            For j = LBound(hdr, 2) To UBound(hdr, 2)
                If InStr(1, NormH(hdr(1, j)), t, vbBinaryCompare) > 0 Then HeaderIndex = j: Exit Function
            Next
        End If
    Next
End Function

Public Function NormH(x As Variant) As String
    If IsError(x) Then Exit Function
    If IsEmpty(x) Or IsNull(x) Then Exit Function
    NormH = NormHeaderText(CStr(x))
End Function

Private Function NormHeaderText(ByVal s As String) As String
    NormHeaderText = LCase$(NormName(s))
End Function

Private Function MapDesc(hdr As Variant, roles As Variant) As String
    Dim i As Long, s As String, used As Object, j As Long, un As String
    Set used = NewDict()
    For i = LBound(roles) To UBound(roles) Step 2
        If roles(i + 1) > 0 Then
            s = s & roles(i) & "=" & ColLetter(roles(i + 1)) & " '" & NormH(hdr(1, roles(i + 1))) & "'; "
            used(roles(i + 1)) = 1
        Else
            s = s & roles(i) & "=-; "
        End If
    Next
    For j = LBound(hdr, 2) To UBound(hdr, 2)
        If Not used.Exists(j) And Len(NormH(hdr(1, j))) > 0 Then un = un & ColLetter(j) & " '" & NormH(hdr(1, j)) & "', "
    Next
    If Len(un) > 0 Then s = s & "unmapped: " & Left$(un, Len(un) - 2)
    MapDesc = s
End Function

Private Function Collisions(roles As Variant) As String
    Dim i As Long, j As Long, s As String
    For i = LBound(roles) To UBound(roles) Step 2
        For j = i + 2 To UBound(roles) Step 2
            If roles(i + 1) > 0 And roles(i + 1) = roles(j + 1) Then
                If Not ((roles(i) = "ORDER" And roles(j) = "TICKET") Or (roles(i) = "TICKET" And roles(j) = "ORDER")) Then
                    s = s & "Column " & ColLetter(roles(i + 1)) & " mapped to both " & roles(i) & " and " & roles(j) & ". "
                End If
            End If
        Next
    Next
    Collisions = s
End Function

Private Function ColLetter(ByVal c As Long) As String
    Dim s As String
    Do While c > 0
        s = Chr$(65 + ((c - 1) Mod 26)) & s
        c = (c - 1) \ 26
    Loop
    ColLetter = s
End Function

Public Function SortedJoin(d As Object, Optional ByVal sep As String = ";") As String
    Dim a() As String, i As Long, j As Long, t As String, k As Variant, n As Long
    If d.Count = 0 Then Exit Function
    ReDim a(1 To d.Count)
    For Each k In d.Keys: n = n + 1: a(n) = CStr(k): Next
    For i = 2 To n                                            ' insertion sort, binary compare
        t = a(i): j = i - 1
        Do While j >= 1
            If StrComp(a(j), t, vbBinaryCompare) <= 0 Then Exit Do
            a(j + 1) = a(j): j = j - 1
        Loop
        a(j + 1) = t
    Next
    SortedJoin = Join(a, sep)
End Function

Private Function QtyOut(ByVal s As String) As Variant
    If IsPlainNumber(s) Then QtyOut = Val(s) Else QtyOut = s
End Function

' Text of a cell, locale-independent for numbers and dates
Public Function CellS(v As Variant, ByVal i As Long, ByVal c As Long) As String
    Dim x As Variant
    If c = 0 Then Exit Function
    x = v(i, c)
    Select Case VarType(x)
        Case vbString
            CellS = Trim$(x)
        Case vbEmpty, vbNull, vbError
            CellS = ""
        Case vbBoolean
            CellS = IIf(x, "TRUE", "FALSE")
        Case vbDate
            If CDbl(x) = Int(CDbl(x)) Then CellS = Format$(x, "yyyy-mm-dd") Else CellS = Format$(x, "yyyy-mm-dd hh:nn:ss")
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte
            CellS = InvariantNum(CDbl(x))
        Case Else
            CellS = Trim$(CStr(x))
    End Select
End Function

'==========================================================================================
'                                    DATES (never through a locale string)
'==========================================================================================
Public Function ParseDateCell(ByVal x As Variant, ByVal fmt As String, ByRef d As Date) As Boolean
    Select Case VarType(x)
        Case vbDate
            d = DateSerial(Year(x), Month(x), Day(x)): ParseDateCell = True
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte
            If CDbl(x) >= 1 And CDbl(x) < 2958466 Then
                d = CDate(Int(CDbl(x))): ParseDateCell = True
            Else
                ParseDateCell = ParseDateText(InvariantNum(CDbl(x)), fmt, d)
            End If
        Case vbString
            ParseDateCell = ParseDateText(CStr(x), fmt, d)
    End Select
End Function

' ISO yyyy-mm-dd is always accepted; anything else only through fmt (MM/DD/YYYY, DD-MMM-YY, YYYYMMDD...)
Public Function ParseDateText(ByVal s As String, ByVal fmt As String, ByRef d As Date) As Boolean
    Dim runs As Collection, grp As Collection, i As Long, k As Long, p As Long, sepd As Boolean
    Dim y As Long, mo As Long, dd As Long
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    If Len(s) >= 10 Then
        If Left$(s, 10) Like "####-##-##" Then
            ParseDateText = MakeDate(CLng(Left$(s, 4)), CLng(Mid$(s, 6, 2)), CLng(Mid$(s, 9, 2)), d)
            Exit Function
        End If
    End If
    fmt = UCase$(Trim$(fmt))
    If Len(fmt) = 0 Then Exit Function
    p = InStr(1, s, " "): If p > 0 Then s = Left$(s, p - 1)
    Set runs = FormatRuns(fmt, sepd)
    If runs Is Nothing Then Exit Function
    If sepd Then
        Set grp = AlnumGroups(s)
        If grp.Count <> runs.Count Then Exit Function
        For i = 1 To runs.Count
            If Not AssignDatePart(runs(i), grp(i), y, mo, dd) Then Exit Function
        Next
    Else
        If Len(s) <> Len(fmt) Then Exit Function
        k = 1
        For i = 1 To runs.Count
            If Not AssignDatePart(runs(i), Mid$(s, k, Len(runs(i))), y, mo, dd) Then Exit Function
            k = k + Len(runs(i))
        Next
    End If
    ParseDateText = MakeDate(y, mo, dd, d)
End Function

' "" if fmt is a supported date format, else the problem
Public Function ValidateDateFormat(ByVal fmt As String) As String
    Dim runs As Collection, sepd As Boolean, r As Variant, nD As Long, nM As Long, nY As Long
    Set runs = FormatRuns(UCase$(Trim$(fmt)), sepd)
    If runs Is Nothing Then ValidateDateFormat = "Unsupported format '" & fmt & "' (use D/M/Y letters, e.g. MM/DD/YYYY, DD-MMM-YY, YYYYMMDD).": Exit Function
    For Each r In runs
        Select Case Left$(r, 1)
            Case "D": nD = nD + 1
            Case "M": nM = nM + 1
            Case "Y": nY = nY + 1
        End Select
    Next
    If nD <> 1 Or nM <> 1 Or nY <> 1 Then ValidateDateFormat = "Format '" & fmt & "' needs exactly one day, one month and one year part."
End Function

' Letter runs of a format. sepd = runs separated by non-letters; Nothing if invalid / mixed
Private Function FormatRuns(ByVal fmt As String, ByRef sepd As Boolean) As Collection
    Dim c As New Collection, i As Long, ch As String, cur As String, nSep As Long, gaps As Long, prevSep As Boolean
    If Len(fmt) = 0 Then Exit Function
    For i = 1 To Len(fmt)
        ch = Mid$(fmt, i, 1)
        If ch Like "[A-Z]" Then
            If ch <> "D" And ch <> "M" And ch <> "Y" Then Exit Function
            If Len(cur) > 0 And Left$(cur, 1) <> ch Then
                c.Add cur: cur = ""
                If Not prevSep Then gaps = gaps + 1
            End If
            cur = cur & ch: prevSep = False
        Else
            If Len(cur) > 0 Then c.Add cur: cur = ""
            If i > 1 Then nSep = nSep + 1
            prevSep = True
        End If
    Next
    If Len(cur) > 0 Then c.Add cur
    If c.Count <> 3 Then Exit Function
    If gaps > 0 And nSep > 0 Then Exit Function               ' mixed separated / positional
    sepd = (nSep > 0)
    For i = 1 To c.Count
        Select Case c(i)
            Case "D", "DD", "M", "MM", "YY", "YYYY"
            Case "MMM", "MMMM"
                If Not sepd Then Exit Function
            Case Else
                Exit Function
        End Select
        If Not sepd And (c(i) = "D" Or c(i) = "M") Then Exit Function
    Next
    Set FormatRuns = c
End Function

Private Function AlnumGroups(ByVal s As String) As Collection
    Dim c As New Collection, i As Long, ch As String, cur As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[0-9A-Za-z]" Then
            cur = cur & ch
        ElseIf Len(cur) > 0 Then
            c.Add cur: cur = ""
        End If
    Next
    If Len(cur) > 0 Then c.Add cur
    Set AlnumGroups = c
End Function

Private Function AssignDatePart(ByVal run As String, ByVal t As String, ByRef y As Long, ByRef mo As Long, ByRef dd As Long) As Boolean
    Select Case Left$(run, 1)
        Case "D"
            If Not IsDigits(t) Or Len(t) > 2 Then Exit Function
            dd = CLng(t)
        Case "M"
            If Len(run) >= 3 Then
                mo = MonthFromName(t)
                If mo = 0 Then Exit Function
            Else
                If Not IsDigits(t) Or Len(t) > 2 Then Exit Function
                mo = CLng(t)
            End If
        Case "Y"
            If Not IsDigits(t) Or Len(t) <> Len(run) Then Exit Function
            y = CLng(t)
            If Len(run) = 2 Then y = 2000 + y
    End Select
    AssignDatePart = True
End Function

Private Function MonthFromName(ByVal t As String) As Long
    Dim p As Long
    t = UCase$(Left$(t, 3))
    If Len(t) < 3 Then Exit Function
    p = InStr(1, "JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC", t)
    If p = 0 Then p = InStr(1, "JANFEVMARABRMAIJUNJULAGOSETOUTNOVDEZ", t)   ' Portuguese
    If p > 0 And (p - 1) Mod 3 = 0 Then MonthFromName = (p - 1) \ 3 + 1
End Function

Private Function MakeDate(ByVal y As Long, ByVal mo As Long, ByVal dd As Long, ByRef d As Date) As Boolean
    Dim t As Date
    If y < 1900 Or y > 9999 Or mo < 1 Or mo > 12 Or dd < 1 Or dd > 31 Then Exit Function
    t = DateSerial(y, mo, dd)
    If Month(t) <> mo Or Day(t) <> dd Then Exit Function
    d = t: MakeDate = True
End Function

'==========================================================================================
'                                    PROCESS (I/O wrapper around the core)
'==========================================================================================
Private Function ProcessData(v As Variant, vRaw As Variant, ByVal srcLabel As String, ByVal srcWb As Object, ByVal srcFile As String) As Boolean
    Dim res As Object, out As Variant, m As Long, arch As String, i As Long, ts As Date
    Dim prevMax As Double, dates As Object, msg As String, lst As String, gaps As String, td As Date
    Dim nB As Long, nBad As Long

    On Error GoTo EH
    Set res = ParseExport(v, LoadDealerMap(), PrevBusinessDay(Date))
    Notify "Parsed " & srcLabel & ": " & (UBound(v, 1) - 1) & " data row(s) | mode=" & res("mode") & _
           IIf(DEALER_MODE = "AUTO", " (auto-detected: pin DEALER_MODE on CONFIG)", " (pinned)") & _
           " | GROUP_LEVEL=" & GROUP_LEVEL & " | keys " & res("keykinds") & vbLf & "Header mapping: " & res("mapping"), LV_INFO
    If Len(res("warnings")) > 0 Then Notify "Data quality (" & srcLabel & "): " & res("warnings"), LV_WARN
    If Not res("ok") Then Notify "ERROR parsing " & srcLabel & ": " & res("err"), LV_ERROR: Exit Function
    If res("nDropped") > 0 Then WriteLog res("nDropped") & " deal(s) excluded as not executed (all rows match EXCL_STATUS): " & res("dropped"), LV_INFO
    If Len(res("unmapped")) > 0 Then WriteLog "Dealer names not in DEALER_MAP (check aliases): " & res("unmapped"), LV_INFO

    m = res("m")
    If m = 0 Then
        arch = ArchiveExport(srcWb, srcFile, vRaw, PrevBusinessDay(Date))
        WriteRaw vRaw
        Notify "ALERT: export contains no executed trades (" & srcLabel & "). Expected trade date " & _
               Format$(PrevBusinessDay(Date), "yyyy-mm-dd") & ". Confirm that the desk did not trade." & _
               IIf(Len(arch) > 0, " Archived: " & arch, ""), LV_ALERT
        ProcessData = True
        Exit Function
    End If

    out = res("deals"): ts = Now: td = res("maxTD")
    arch = ArchiveExport(srcWb, srcFile, vRaw, td)
    For i = 1 To m
        out(i, 8) = ts
        out(i, 10) = IIf(Len(arch) > 0, arch, srcLabel)
    Next
    prevMax = HistDailyMaxDate()
    WriteRaw vRaw
    WriteDeals out, m
    Set dates = UpsertHist(GetSheet(SH_HIST), out, m)
    RecalcHistDaily GetSheet(SH_HIST), GetSheet(SH_HD), dates
    gaps = CoverageGaps(prevMax, res("minTD"))
    BuildDashboard out, m, "Last run " & Format$(ts, "yyyy-mm-dd hh:nn") & " | " & srcLabel & " | mode " & res("mode") & " | " & GROUP_LEVEL & " level"

    nB = res("nB"): nBad = res("nBad")
    msg = "Trades " & DateRangeText(res("minTD"), res("maxTD")) & ": " & m & " deal(s) | " & nB & " below " & MIN_Q & _
          " counterparties (" & Format$(nB / m, "0.0%") & ") | mode=" & res("mode")
    If nBad > 0 Then msg = msg & " | " & nBad & " deal(s) with inconsistent key (KEY?)"
    If nB > 0 Or nBad > 0 Then
        For i = 1 To m
            If out(i, 7) <> "OK" Then
                lst = lst & vbLf & Format$(out(i, 1), "yyyy-mm-dd") & "  " & out(i, 2) & "  " & out(i, 3) & "  " & out(i, 4) & "  " & out(i, 5) & _
                      "  dealers=" & out(i, 6) & IIf(Len(out(i, 9)) > 0, " (" & out(i, 9) & ")", "") & "  [" & out(i, 7) & "]"
            End If
        Next
        Notify "ALERT " & msg & lst, LV_ALERT
    Else
        Notify "OK " & msg, LV_INFO
        If SEND_OK_EMAIL And Len(ALERT_TO) > 0 And Not gHeadless Then SendMail "[BLOT Best Ex] OK " & DateRangeText(res("minTD"), res("maxTD")), "OK " & msg
    End If
    If Len(gaps) > 0 Then
        Notify "ALERT coverage gap: no BLOT data loaded for business day(s) " & gaps & _
               ". Backfill with Import File, or check that OPEN_SEQ shows the previous business day (README step 2g).", LV_ALERT
    End If
    ProcessData = True
    Exit Function
EH:
    Notify "ERROR in ProcessData (" & srcLabel & "): " & Err.Description, LV_ERROR
End Function

Private Function DateRangeText(ByVal d0 As Date, ByVal d1 As Date) As String
    If d0 = d1 Then DateRangeText = Format$(d0, "yyyy-mm-dd") Else DateRangeText = Format$(d0, "yyyy-mm-dd") & ".." & Format$(d1, "yyyy-mm-dd")
End Function

' Dated copy of the raw export in ARCHIVE_DIR. Returns the archive path ("" if not archived).
Private Function ArchiveExport(ByVal srcWb As Object, ByVal srcFile As String, vRaw As Variant, ByVal td As Date) As String
    Dim base As String, ext As String, p As String, wb As Workbook
    If Len(ARCHIVE_DIR) = 0 Then
        Notify "ALERT: ARCHIVE_DIR is empty on CONFIG - the raw export was NOT archived (regulatory evidence).", LV_ALERT
        Exit Function
    End If
    base = ARCHIVE_DIR & "\BLOT_" & Format$(td, "yyyymmdd") & "_" & Format$(Now, "yyyymmdd_hhnnss")
    On Error GoTo Fallback
    If Len(srcFile) > 0 Then
        ext = Mid$(srcFile, InStrRev(srcFile, "."))
        p = base & LCase$(ext)
        FileCopy srcFile, p
    ElseIf Not srcWb Is Nothing Then
        Select Case srcWb.FileFormat
            Case 52: ext = ".xlsm"
            Case 56: ext = ".xls"
            Case 50: ext = ".xlsb"
            Case 6: ext = ".csv"
            Case Else: ext = ".xlsx"
        End Select
        p = base & ext
        srcWb.SaveCopyAs p
    Else
        GoTo Fallback
    End If
    ArchiveExport = p
    Exit Function
Fallback:
    On Error GoTo -1
    On Error GoTo Failed
    p = base & ".xlsx"
    Set wb = Workbooks.Add
    wb.Worksheets(1).Range("A1").Resize(UBound(vRaw, 1), UBound(vRaw, 2)).Value = TextSafe(vRaw)
    Application.DisplayAlerts = False
    wb.SaveAs p, 51
    Application.DisplayAlerts = True
    wb.Close False
    ArchiveExport = p
    Exit Function
Failed:
    Application.DisplayAlerts = True
    Notify "ALERT: could not archive the raw export to " & ARCHIVE_DIR & " (" & Err.Description & ").", LV_ALERT
    On Error GoTo -1
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
End Function

'==========================================================================================
'                                    OUTPUT
'==========================================================================================
Public Function HistHeader() As Variant
    HistHeader = Array("TradeDate", "TradeKey", "Security", "Side", "Qty", "NDealers", "Status", "LoadedAt", "Dealers", "Source")
End Function

Public Function HistDailyHeader() As Variant
    HistDailyHeader = Array("TradeDate", "NTrades", "NBelowMin", "PctBelowMin", "AvgDealers", "MinDealers", "MaxDealers", "UpdatedAt", "NKeyIssues")
End Function

Public Sub EnsureHeader(ws As Worksheet, h As Variant)
    Dim n As Long
    n = UBound(h) - LBound(h) + 1
    If Len(ws.Range("A1").Value) = 0 Or ws.Cells(1, n).Value <> h(UBound(h)) Then
        ws.Range("A1").Resize(1, n).Value = h
        ws.Range("A1").Resize(1, n).Font.Bold = True
        ws.Range("A1").Resize(1, n).Interior.Color = RGB(217, 225, 242)
        ws.Columns("A").NumberFormat = "yyyy-mm-dd"
        ws.Columns("A").ColumnWidth = 12
        ws.Columns("H").NumberFormat = "yyyy-mm-dd hh:mm:ss"
        ws.Columns("H").ColumnWidth = 19
        If h(1) = "TradeKey" Then
            ws.Columns("B").NumberFormat = "@"
            ws.Columns("B").ColumnWidth = 30: ws.Columns("C").ColumnWidth = 16
            ws.Columns("I").ColumnWidth = 36: ws.Columns("J").ColumnWidth = 60
        Else
            ws.Columns("D").NumberFormat = "0.0%": ws.Columns("E").NumberFormat = "0.00"
            ws.Columns("B:G").ColumnWidth = 11: ws.Columns("I").ColumnWidth = 11
        End If
    End If
End Sub

Private Sub EnsureLogHeader(ws As Worksheet)
    If Len(ws.Range("A1").Value) = 0 Then
        ws.Range("A1:C1").Value = Array("Timestamp", "Level", "Message")
        ws.Range("A1:C1").Font.Bold = True: ws.Range("A1:C1").Interior.Color = RGB(217, 225, 242)
        ws.Columns("A").ColumnWidth = 20: ws.Columns("B").ColumnWidth = 8: ws.Columns("C").ColumnWidth = 160
    End If
End Sub

Private Function TextSafe(v As Variant) As Variant
    Dim o As Variant, i As Long, j As Long
    o = v
    If Not IsArray(o) Then TextSafe = o: Exit Function
    For i = LBound(o, 1) To UBound(o, 1)
        For j = LBound(o, 2) To UBound(o, 2)
            If VarType(o(i, j)) = vbString Then
                If Len(o(i, j)) > 0 Then o(i, j) = "'" & o(i, j)
            End If
        Next
    Next
    TextSafe = o
End Function

Private Sub WriteRaw(vRaw As Variant)
    With GetSheet(SH_RAW)
        .Cells.Clear
        If IsArray(vRaw) Then .Range("A1").Resize(UBound(vRaw, 1), UBound(vRaw, 2)).Value = TextSafe(vRaw)
    End With
End Sub

Private Sub WriteDeals(out As Variant, ByVal m As Long)
    Dim ws As Worksheet
    Set ws = GetSheet(SH_DEALS)
    ws.Cells.Clear
    EnsureHeader ws, HistHeader()
    ws.Range("A2").Resize(m, NCOL).Value = TextSafe(out)
    ws.Range("A1").Resize(m + 1, NCOL).Sort Key1:=ws.Range("F1"), Order1:=xlAscending, Header:=xlYes
    ws.Columns("A:I").AutoFit
End Sub

' Upsert one row per (TradeDate, TradeKey): re-running a day never duplicates. Returns the touched dates.
Public Function UpsertHist(ws As Worksheet, out As Variant, ByVal m As Long) As Object
    Dim ex As Object, hv As Variant, lastR As Long, i As Long, j As Long, rw As Long, k As String
    Dim one() As Variant, dates As Object, ov As Variant
    EnsureHeader ws, HistHeader()
    Set ex = NewDict(): Set dates = NewDict()
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR >= 2 Then
        hv = ws.Range("A1:B" & lastR).Value
        For i = 2 To UBound(hv, 1): ex(HistKey(hv(i, 1), hv(i, 2))) = i: Next
    End If
    ReDim one(1 To 1, 1 To NCOL)
    ov = TextSafe(out)
    For i = 1 To m
        For j = 1 To NCOL: one(1, j) = ov(i, j): Next
        k = HistKey(out(i, 1), out(i, 2))
        If ex.Exists(k) Then
            rw = ex(k)
        Else
            lastR = lastR + 1: rw = lastR: ex(k) = rw
        End If
        ws.Cells(rw, 1).Resize(1, NCOL).Value = one
        dates(CLng(out(i, 1))) = 1
    Next
    Set UpsertHist = dates
End Function

Private Function HistKey(d As Variant, id As Variant) As String
    Select Case VarType(d)
        Case vbDate, vbDouble, vbLong, vbInteger, vbSingle, vbCurrency
            HistKey = Format$(CDate(Int(CDbl(d))), "yyyymmdd") & "|" & CStr(id)
        Case Else
            HistKey = CStr(d) & "|" & CStr(id)
    End Select
End Function

' HIST_DAILY rows of the given dates are recomputed from HIST (the audit trail is the source of truth)
Public Sub RecalcHistDaily(wsH As Worksheet, wsD As Worksheet, dates As Object)
    Dim hv As Variant, lastR As Long, i As Long, dk As Long, agg As Object, a As Variant, k As Variant
    Dim dv As Variant, rw As Long, rowOf As Object, q As Long, lastD As Long
    EnsureHeader wsD, HistDailyHeader()
    lastR = wsH.Cells(wsH.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Sub
    hv = wsH.Range("A1:G" & lastR).Value2
    Set agg = NewDict()
    For i = 2 To UBound(hv, 1)
        If VarType(hv(i, 1)) = vbDouble Then
            dk = CLng(Int(hv(i, 1)))
            If dates.Exists(dk) Then
                If Not agg.Exists(dk) Then agg(dk) = Array(0#, 0#, 0#, 1000000#, -1#, 0#)
                a = agg(dk)
                q = CLng(Val(CStr(hv(i, 6))))
                a(0) = a(0) + 1
                If q < MIN_Q Then a(1) = a(1) + 1
                a(2) = a(2) + q
                If q < a(3) Then a(3) = q
                If q > a(4) Then a(4) = q
                If CStr(hv(i, 7)) = "KEY?" Then a(5) = a(5) + 1
                agg(dk) = a
            End If
        End If
    Next
    Set rowOf = NewDict()
    lastD = wsD.Cells(wsD.Rows.Count, 1).End(xlUp).Row
    If lastD >= 2 Then
        dv = wsD.Range("A1:A" & lastD).Value2
        For i = 2 To UBound(dv, 1)
            If VarType(dv(i, 1)) = vbDouble Then rowOf(CLng(Int(dv(i, 1)))) = i
        Next
    End If
    For Each k In agg.Keys
        a = agg(k)
        If rowOf.Exists(k) Then
            rw = rowOf(k)
        Else
            lastD = lastD + 1: rw = lastD: rowOf(k) = rw
        End If
        wsD.Cells(rw, 1).Resize(1, 9).Value = Array(CDate(k), a(0), a(1), a(1) / a(0), a(2) / a(0), a(3), a(4), Now, a(5))
    Next
    If lastD > 2 Then wsD.Range("A1").Resize(lastD, 9).Sort Key1:=wsD.Range("A1"), Order1:=xlAscending, Header:=xlYes
End Sub

Private Function HistDailyMaxDate() As Double
    On Error Resume Next
    HistDailyMaxDate = Application.WorksheetFunction.Max(GetSheet(SH_HD).Range("A:A"))
End Function

' Business days strictly between the last loaded trade date and the first new one
Private Function CoverageGaps(ByVal prevMax As Double, ByVal newMin As Date) As String
    Dim d As Date, s As String, n As Long
    If prevMax < 1 Or CDbl(newMin) <= prevMax + 1 Then Exit Function
    For d = CDate(prevMax) + 1 To newMin - 1
        If IsBusinessDay(d) Then
            n = n + 1
            If n <= 30 Then s = s & Format$(d, "yyyy-mm-dd") & " "
        End If
    Next
    If n > 30 Then s = s & "(+" & (n - 30) & " more)"
    CoverageGaps = Trim$(s)
End Function

Private Sub BuildDashboard(out As Variant, ByVal m As Long, ByVal info As String)
    Dim ws As Worksheet, co As ChartObject, nB As Long, nBad As Long, avgQ As Double, mn As Long, mx As Long
    Dim t0 As Date, t1 As Date, i As Long, q As Long, r As Long, dist() As Long, lk As String
    Dim hd As Worksheet, lr As Long, r0 As Long, hv As Variant, rr As Long

    Set ws = GetSheet(SH_DASH)
    For Each co In ws.ChartObjects: co.Delete: Next
    On Error Resume Next: ws.Cells.Validation.Delete: On Error GoTo 0
    ws.Cells.Clear
    AddSheetButtons ws, DashButtons()
    ws.Columns("A").ColumnWidth = 2

    ws.Range("B3").Value = "BLOT - Best Execution | minimum " & MIN_Q & " counterparties asked per deal (Banco de Portugal)"
    ws.Range("B3").Font.Size = 14: ws.Range("B3").Font.Bold = True
    ws.Range("B4").Value = "'" & info
    ws.Range("B4").Font.Italic = True: ws.Range("B4").Font.Color = RGB(89, 89, 89)
    If m = 0 Then
        ws.Range("B6").Value = "No data loaded yet: fill CONFIG, click Validate Config, then Run Now (or Import File)."
        ws.Columns("B").ColumnWidth = 40
        Exit Sub
    End If

    Stats out, m, nB, nBad, avgQ, mn, mx, t0, t1
    ws.Range("B6:B11").Value = Application.Transpose(Array("Trade date(s)", "No. trades (deals)", _
        "No. deals < " & MIN_Q & " counterparties", "% deals < " & MIN_Q & " counterparties", _
        "Avg counterparties / deal", "Deals with inconsistent key (KEY?)"))
    ws.Range("B6:B11").Font.Bold = True
    ws.Range("C6").Value = "'" & DateRangeText(t0, t1)
    ws.Range("C7").Value = m
    ws.Range("C8").Value = nB
    ws.Range("C9").Value = nB / m: ws.Range("C9").NumberFormat = "0.0%"
    ws.Range("C10").Value = avgQ:  ws.Range("C10").NumberFormat = "0.00"
    ws.Range("C11").Value = nBad
    ws.Range("C6:C11").HorizontalAlignment = xlRight

    With ws.Range("B13:C13")
        .Merge
        If nB > 0 Or nBad > 0 Then
            .Value = "ALERT: " & nB & " deal(s) with fewer than " & MIN_Q & " counterparties" & IIf(nBad > 0, ", " & nBad & " KEY?", "")
            .Interior.Color = RGB(192, 0, 0)
        Else
            .Value = "OK: all deals with >= " & MIN_Q & " counterparties"
            .Interior.Color = RGB(0, 128, 0)
        End If
        .Font.Color = vbWhite: .Font.Bold = True: .HorizontalAlignment = xlCenter
    End With

    '--- distribution ---
    ReDim dist(0 To mx)
    For i = 1 To m: dist(out(i, 6)) = dist(out(i, 6)) + 1: Next
    ws.Range("E6:G6").Value = Array("No. counterparties", "No. deals", "% total"): ws.Range("E6:G6").Font.Bold = True
    For q = 0 To mx
        r = 7 + q
        ws.Cells(r, 5).Value = q: ws.Cells(r, 6).Value = dist(q): ws.Cells(r, 7).Value = dist(q) / m
        If q < MIN_Q Then ws.Range(ws.Cells(r, 5), ws.Cells(r, 7)).Interior.Color = RGB(255, 199, 206)
    Next
    ws.Range(ws.Cells(7, 7), ws.Cells(7 + mx, 7)).NumberFormat = "0.0%"

    Set co = ws.ChartObjects.Add(ws.Range("I6").Left, ws.Range("I6").Top, 420, 230)
    With co.Chart
        .ChartType = xlColumnClustered
        .SetSourceData ws.Range(ws.Cells(6, 6), ws.Cells(7 + mx, 6))
        .SeriesCollection(1).XValues = ws.Range(ws.Cells(7, 5), ws.Cells(7 + mx, 5))
        .HasTitle = True: .ChartTitle.Text = "Counterparties asked per deal"
        .HasLegend = False
        For q = 0 To mx
            If q < MIN_Q Then .SeriesCollection(1).Points(q + 1).Format.Fill.ForeColor.RGB = RGB(192, 0, 0)
        Next
    End With

    '--- per-deal lookup (latest run) ---
    r = Application.WorksheetFunction.Max(15, 9 + mx)
    ws.Cells(r, 2).Value = "Deal lookup (latest run)": ws.Cells(r, 2).Font.Bold = True: ws.Cells(r, 2).Font.Size = 12
    ws.Cells(r + 1, 2).Resize(5, 1).Value = Application.Transpose(Array("Deal (TradeKey)", "Security", "No. counterparties", "Counterparties", "Status"))
    ws.Cells(r + 1, 2).Resize(5, 1).Font.Bold = True
    With ws.Cells(r + 1, 3)
        .NumberFormat = "@"
        .Validation.Add Type:=xlValidateList, Formula1:="='" & SH_DEALS & "'!$B$2:$B$" & (m + 1)
        .Interior.Color = RGB(255, 242, 204)
        .Value = GetSheet(SH_DEALS).Range("B2").Value         ' default: worst deal
    End With
    lk = "MATCH(" & ws.Cells(r + 1, 3).Address & ",'" & SH_DEALS & "'!$B:$B,0)"
    ws.Cells(r + 2, 3).Formula = "=IFERROR(INDEX('" & SH_DEALS & "'!$C:$C," & lk & "),"""")"
    ws.Cells(r + 3, 3).Formula = "=IFERROR(INDEX('" & SH_DEALS & "'!$F:$F," & lk & "),"""")"
    ws.Cells(r + 4, 3).Formula = "=IFERROR(INDEX('" & SH_DEALS & "'!$I:$I," & lk & "),"""")"
    ws.Cells(r + 5, 3).Formula = "=IFERROR(INDEX('" & SH_DEALS & "'!$G:$G," & lk & "),"""")"
    ws.Cells(r + 2, 3).Resize(4, 1).HorizontalAlignment = xlLeft

    '--- history lookup (any date) ---
    r = r + 7
    ws.Cells(r, 2).Value = "History lookup (HIST, any date)": ws.Cells(r, 2).Font.Bold = True: ws.Cells(r, 2).Font.Size = 12
    ws.Cells(r + 1, 2).Resize(5, 1).Value = Application.Transpose(Array("TradeKey (type it)", "Trade date", "No. counterparties", "Counterparties", "Status"))
    ws.Cells(r + 1, 2).Resize(5, 1).Font.Bold = True
    ws.Cells(r + 1, 3).NumberFormat = "@": ws.Cells(r + 1, 3).Interior.Color = RGB(255, 242, 204)
    lk = "MATCH(" & ws.Cells(r + 1, 3).Address & ",'" & SH_HIST & "'!$B:$B,0)"
    ws.Cells(r + 2, 3).Formula = "=IFERROR(INDEX('" & SH_HIST & "'!$A:$A," & lk & "),"""")"
    ws.Cells(r + 2, 3).NumberFormat = "yyyy-mm-dd"
    ws.Cells(r + 3, 3).Formula = "=IFERROR(INDEX('" & SH_HIST & "'!$F:$F," & lk & "),"""")"
    ws.Cells(r + 4, 3).Formula = "=IFERROR(INDEX('" & SH_HIST & "'!$I:$I," & lk & "),"""")"
    ws.Cells(r + 5, 3).Formula = "=IFERROR(INDEX('" & SH_HIST & "'!$G:$G," & lk & "),"""")"
    ws.Cells(r + 2, 3).Resize(4, 1).HorizontalAlignment = xlLeft

    '--- exceptions ---
    r = r + 7
    ws.Cells(r, 2).Value = "Exceptions (< " & MIN_Q & " counterparties or KEY?)": ws.Cells(r, 2).Font.Bold = True: ws.Cells(r, 2).Font.Size = 12
    ws.Range(ws.Cells(r + 1, 2), ws.Cells(r + 1, 8)).Value = Array("TradeKey", "Security", "Side", "Qty", "NDealers", "Status", "Counterparties")
    ws.Range(ws.Cells(r + 1, 2), ws.Cells(r + 1, 8)).Font.Bold = True
    q = r + 1
    For i = 1 To m
        If out(i, 7) <> "OK" Then
            q = q + 1
            ws.Range(ws.Cells(q, 2), ws.Cells(q, 8)).Value = Array("'" & out(i, 2), "'" & out(i, 3), "'" & out(i, 4), out(i, 5), out(i, 6), "'" & out(i, 7), "'" & out(i, 9))
            ws.Range(ws.Cells(q, 2), ws.Cells(q, 8)).Interior.Color = RGB(255, 199, 206)
        End If
    Next
    If q = r + 1 Then ws.Cells(q + 1, 2).Value = "None."

    '--- recent history (HIST_DAILY) ---
    Set hd = GetSheet(SH_HD)
    lr = hd.Cells(hd.Rows.Count, 1).End(xlUp).Row
    r0 = 23
    ws.Cells(r0, 9).Value = "Recent history (HIST_DAILY)": ws.Cells(r0, 9).Font.Bold = True: ws.Cells(r0, 9).Font.Size = 12
    ws.Range(ws.Cells(r0 + 1, 9), ws.Cells(r0 + 1, 13)).Value = Array("Trade date", "Trades", "< " & MIN_Q, "% < " & MIN_Q, "KEY?")
    ws.Range(ws.Cells(r0 + 1, 9), ws.Cells(r0 + 1, 13)).Font.Bold = True
    If lr >= 2 Then
        rr = r0 + 1
        For i = lr To Application.WorksheetFunction.Max(2, lr - 14) Step -1
            hv = hd.Range(hd.Cells(i, 1), hd.Cells(i, 9)).Value
            rr = rr + 1
            ws.Cells(rr, 9).Value = hv(1, 1): ws.Cells(rr, 9).NumberFormat = "yyyy-mm-dd"
            ws.Cells(rr, 10).Value = hv(1, 2): ws.Cells(rr, 11).Value = hv(1, 3)
            ws.Cells(rr, 12).Value = hv(1, 4): ws.Cells(rr, 12).NumberFormat = "0.0%"
            ws.Cells(rr, 13).Value = hv(1, 9)
            If Val(CStr(hv(1, 3))) > 0 Or Val(CStr(hv(1, 9))) > 0 Then ws.Range(ws.Cells(rr, 9), ws.Cells(rr, 13)).Interior.Color = RGB(255, 199, 206)
        Next
    End If

    ws.Columns("B").ColumnWidth = 36: ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 8: ws.Columns("E:H").ColumnWidth = 14
    ws.Columns("I:M").ColumnWidth = 12
End Sub

Private Sub Stats(out As Variant, ByVal m As Long, nB As Long, nBad As Long, avgQ As Double, mn As Long, mx As Long, t0 As Date, t1 As Date)
    Dim i As Long, s As Double
    mn = 1000000: mx = 0: nB = 0: nBad = 0: t0 = 0: t1 = 0
    For i = 1 To m
        s = s + out(i, 6)
        If out(i, 6) < MIN_Q Then nB = nB + 1
        If out(i, 7) = "KEY?" Then nBad = nBad + 1
        If out(i, 6) < mn Then mn = out(i, 6)
        If out(i, 6) > mx Then mx = out(i, 6)
        If t0 = 0 Or out(i, 1) < t0 Then t0 = out(i, 1)
        If out(i, 1) > t1 Then t1 = out(i, 1)
    Next
    If m > 0 Then avgQ = s / m
End Sub

'==========================================================================================
'                                    NOTIFY / LOG
'==========================================================================================
' ALERT / ERROR are e-mailed to ALERT_TO and shown as popup when SHOW_POPUPS (never when headless)
Public Sub Notify(ByVal msg As String, Optional ByVal level As String = "INFO", _
                  Optional ByVal forcePopup As Boolean = False, Optional ByVal noMail As Boolean = False)
    Dim isAlert As Boolean
    WriteLog msg, level
    isAlert = (level = LV_ALERT Or level = LV_ERROR)
    If gHeadless Then Exit Sub
    If isAlert And Not noMail And Len(ALERT_TO) > 0 Then SendMail "[BLOT Best Ex] " & Left$(Split(msg & vbLf, vbLf)(0), 150), msg
    If Application.Visible And ((isAlert And SHOW_POPUPS) Or forcePopup) Then
        MsgBox Left$(msg, 1000), IIf(isAlert, vbExclamation, vbInformation), "BLOT Best Execution"
    End If
End Sub

Public Sub WriteLog(ByVal msg As String, ByVal level As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = GetSheet(SH_LOG)
    EnsureLogHeader ws
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now: ws.Cells(r, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(r, 2).Value = level
    ws.Cells(r, 3).Value = "'" & Left$(Replace(msg, vbCr, ""), 32000)
    If level = LV_ALERT Or level = LV_ERROR Then
        ws.Cells(r, 2).Font.Color = RGB(192, 0, 0): ws.Cells(r, 2).Font.Bold = True
    ElseIf level = LV_WARN Then
        ws.Cells(r, 2).Font.Color = RGB(156, 101, 0)
    End If
End Sub

Private Sub SendMail(ByVal subj As String, ByVal body As String)
    Dim ol As Object, mi As Object
    On Error GoTo EH
    Set ol = CreateObject("Outlook.Application")
    Set mi = ol.CreateItem(0)
    mi.To = ALERT_TO: mi.Subject = subj: mi.Body = body
    mi.Send
    Exit Sub
EH:
    WriteLog "ERROR sending e-mail to " & ALERT_TO & ": " & Err.Description, LV_ERROR
End Sub

'==========================================================================================
'                                    HELPERS
'==========================================================================================
Public Function GetSheet(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If GetSheet Is Nothing Then
        Set GetSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetSheet.Name = nm
    End If
End Function

Public Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

Private Sub WaitSec(ByVal s As Double)
    Dim t As Double: t = Timer + s
    Do While Timer < t: DoEvents: Loop
End Sub
