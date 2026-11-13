Attribute VB_Name = "modBlotBestEx"
Option Explicit
'==========================================================================================
' BLOT -> Best Execution monitor v5 (Banco de Portugal: >= MIN_Q counterparties ASKED per request)
'------------------------------------------------------------------------------------------
' Export layout (Bloomberg BLOT, one row per DEALER LEG of a quote request):
'   Security | Side | Block | Status | Qty (M) | Price | Yield | BrkrName | ... | Trade Dt | ISIN | ... | ID
' Rule:
'   request  = rows whose ID is equal except the last ID_SUFFIX_LEN characters (the dealer leg)
'              FITDNA:6AF5A884011800|03, |05, |06 = one request to 3 dealers
'              checked: one ISIN, one side, one trade date per request, else KEY?
'   asked    = distinct dealers (BrkrName, DEALER_MAP applied), whatever their status
'   quoted   = distinct dealers whose row has a price
'   filled   = a row whose status matches FILL_STATUS (winner = that dealer)
'   check    = BREACH (filled, asked < MIN_Q) | LOW (not filled, asked < MIN_Q) | KEY? | OK
' Buttons: Run Now, Process Pasted Data, Import File, Test Export, Validate Config, Abort
'==========================================================================================

'------------------------------------ WIN32 ----------------------------------------------
Private Type GUID_T
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As LongPtr, ByVal dwId As Long, riid As GUID_T, ppvObject As Object) As Long
Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, lpiid As GUID_T) As Long
Private Const OBJID_NATIVEOM As Long = &HFFFFFFF0
Private Const SW_RESTORE As Long = 9

'------------------------------------ CONSTANTS ------------------------------------------
Public Const UCOLS As Long = 16      ' INQUIRIES / HIST columns
Public Const QCOLS As Long = 13      ' QUOTES columns
Public Const LV_INFO  As String = "INFO"
Public Const LV_WARN  As String = "WARN"
Public Const LV_ALERT As String = "ALERT"
Public Const LV_ERROR As String = "ERROR"

'------------------------------------ STATE ----------------------------------------------
Public gHeadless As Boolean
Private gT0 As Date, gSince As Date, gDeadline As Date, gNextRun As Date, gSchedPending As Boolean
Private gKnown As String, gCandKey As String, gCandRows As Long, gBusy As Boolean, gKind As String
Private gPollAt As Date, gPollPending As Boolean, gCycleAt As Date, gCyclePending As Boolean
Private gTestN As Long, gTestI As Long, gTestOK As Long, gLastErr As String, gWarned As String

'==========================================================================================
'                                    ENTRY POINTS
'==========================================================================================
Public Sub RunBlotReport()
    On Error GoTo EH
    LoadConfig
    If Not BeginRun("RUN") Then Exit Sub
    Notify "Run started: expected trade date " & Format$(PrevBusinessDay(Date), "yyyy-mm-dd") & ".", LV_INFO
    If Not StartAcquisition() Then FinishRun
    Exit Sub
EH:
    FinishRun
    Notify "ERROR in RunBlotReport: " & Err.Description, LV_ERROR
End Sub

Public Sub PollExport()
    gPollPending = False
    If Not gBusy Then Exit Sub
    On Error GoTo EH
    If TryAcquire() Then Exit Sub
Reschedule:
    On Error GoTo 0
    If Not gBusy Then Exit Sub
    If Now < gDeadline Then SchedulePoll Else OnAcquireTimeout
    Exit Sub
EH:
    gLastErr = Err.Description
    Resume Reschedule
End Sub

Public Sub AbortRun()
    CancelPending
    gBusy = False: gKind = ""
    Application.StatusBar = False
    Notify "Run aborted by the operator.", LV_WARN, False, True
End Sub

' Paste the BLOT grid (with its header row) on the PASTE sheet, then this button
Public Sub ProcessPasted()
    Dim ws As Worksheet, a As Variant, r As Long, lc As Long, ur As Range, v As Variant, vRaw As Variant
    On Error GoTo EH
    LoadConfig
    Set ws = GetSheet(SH_PASTE)
    Set ur = ws.UsedRange
    If ur.Cells.CountLarge < 3 Then Notify "PASTE is empty: paste the BLOT grid (with its header row) from cell A3.", LV_WARN, True, True: Exit Sub
    a = ur.Value2
    r = FindHeaderRow(a, lc)
    If r = 0 Then Notify "PASTE: no header row with a dealer column (" & Split(HDR_DEALER, "|")(0) & ") and an ID / ISIN column.", LV_WARN, True, True: Exit Sub
    If Not BeginRun("IMPORT") Then Exit Sub
    v = SliceBlock(a, r, lc)
    vRaw = ur.Cells(r, 1).Resize(UBound(a, 1) - r + 1, lc).Value
    ProcessData v, vRaw, "pasted data", Nothing, ""
    FinishRun
    GetSheet(SH_DASH).Activate
    Exit Sub
EH:
    FinishRun
    Notify "ERROR in ProcessPasted: " & Err.Description, LV_ERROR, True, True
End Sub

Public Sub ClearPaste()
    With GetSheet(SH_PASTE)
        .Range(.Cells(3, 1), .Cells(.Rows.Count, 60)).Clear
    End With
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
    GetSheet(SH_DASH).Activate
    Notify "Import finished: " & n & " of " & tot & " file(s) loaded.", LV_INFO, True, True
    Exit Sub
EH:
    FinishRun
    Notify "ERROR in ImportFromFile: " & Err.Description, LV_ERROR
End Sub

Private Function ImportOneFile(ByVal f As String) As Boolean
    Dim blk As Object, wb As Workbook
    On Error GoTo EH
    Set blk = OpenExportFile(f, wb)
    If blk Is Nothing Then
        Notify "Import: no BLOT header row in " & f & ".", LV_WARN
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

' TEST_EXPORT_RUNS cycles of OPEN_SEQ + EXPORT_SEQ + detection; nothing written to history
Public Sub TestExport()
    On Error GoTo EH
    LoadConfig
    If Not BeginRun("TEST") Then Exit Sub
    gTestN = TEST_EXPORT_RUNS: gTestI = 0: gTestOK = 0
    Notify "TestExport started: " & gTestN & " cycles.", LV_INFO
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

Public Function DdeAvailable(ByRef msg As String) As Boolean
    Dim ch As Long, da As Boolean
    If Application.IgnoreRemoteRequests Then
        msg = "Excel option 'Ignore other applications that use DDE' is ON: turn it off (Options > Advanced > General)."
        Exit Function
    End If
    da = Application.DisplayAlerts: Application.DisplayAlerts = False
    On Error GoTo Fail
    ch = Application.DDEInitiate("WINBLP", "BBK")
    Application.DDETerminate ch
    Application.DisplayAlerts = da
    msg = "DDE link to Bloomberg OK."
    DdeAvailable = True
    Exit Function
Fail:
    Application.DisplayAlerts = da
    msg = "DDE link to Bloomberg failed (" & Err.Description & "). Terminal running and logged in?"
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
        ws.Range("A1:B1").Value = Array("BrkrName as exported", "Counterparty (legal entity)")
        Hdr ws.Range("A1:B1")
        ws.Columns("A:B").ColumnWidth = 34: ws.Columns("A:B").NumberFormat = "@"
        ws.Range("D1").Value = "One row per spelling, e.g. 'BNP PARIB.' -> BNP PARIBAS. Matching ignores case and repeated spaces."
        ws.Range("D2").Value = "Names not in this table are listed in LOG after every run."
        ws.Range("D1:D2").Font.Italic = True
    End If
    Set ws = GetSheet(SH_PASTE)
    If Len(ws.Range("A2").Value) = 0 Then
        ws.Range("A2").Value = "Paste the BLOT grid (Ctrl+V, WITH its header row) from cell A3, then click 'Process Pasted Data'."
        ws.Range("A2").Font.Bold = True
        AddSheetButtons ws, Array("Process Pasted Data|ProcessPasted", "Clear|ClearPaste")
    End If
    EnsureHeader GetSheet(SH_HIST), UnitHeader()
    EnsureHeader GetSheet(SH_HD), DailyHeader()
    EnsureLogHeader GetSheet(SH_LOG)
    If Len(GetSheet(SH_INQ).Range("A2").Value) = 0 Then BuildDashboard Empty, 0, Nothing, "No run yet."

    da = Application.DisplayAlerts: Application.DisplayAlerts = False
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        If Not IsOwnSheet(ws.Name) Then
            If Application.WorksheetFunction.CountA(ws.Cells) = 0 And ws.Shapes.Count = 0 Then ws.Delete
        End If
    Next
    Application.DisplayAlerts = da
    For Each nm In SheetOrder()
        Set ws = GetSheet(CStr(nm))
        If prev Is Nothing Then ws.Move Before:=ThisWorkbook.Worksheets(1) Else ws.Move After:=prev
        Set prev = ws
    Next
    GetSheet(SH_TESTS).Visible = xlSheetHidden
    GetSheet(SH_FIX).Visible = xlSheetHidden
    GetSheet(SH_DASH).Activate
    Application.ScreenUpdating = su
    WriteLog "Setup completed.", LV_INFO
    Exit Sub
EH:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Notify "ERROR in Setup: " & Err.Description, LV_ERROR
End Sub

Private Function SheetOrder() As Variant
    SheetOrder = Array(SH_DASH, SH_INQ, SH_QUOTES, SH_PASTE, SH_HIST, SH_HD, SH_CONFIG, SH_MAP, SH_LOG, SH_TESTS, SH_FIX)
End Function

Private Function IsOwnSheet(ByVal nm As String) As Boolean
    Dim s As Variant
    For Each s In SheetOrder()
        If StrComp(CStr(s), nm, vbTextCompare) = 0 Then IsOwnSheet = True: Exit Function
    Next
End Function

Public Function DashButtons() As Variant
    DashButtons = Array("Run Now|RunBlotReport", "Process Pasted Data|ProcessPasted", "Import File|ImportFromFile", _
                        "Test Export|TestExport", "Validate Config|ValidateConfig", "Abort|AbortRun")
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
        Set b = ws.Buttons.Add(x, 4, 128, 26)
        b.OnAction = p(1)
        b.Caption = p(0)
        b.Name = "btn_" & p(1)
        b.Placement = xlFreeFloating
        x = x + 134
    Next
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
    WriteLog "Scheduler armed: next run " & Format$(gNextRun, "yyyy-mm-dd hh:nn:ss"), LV_INFO
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
            Notify "Previous " & gKind & " run looked stuck: state reset.", LV_WARN
            CancelPending
        Else
            Notify "A " & gKind & " run is already in progress (use Abort to reset).", LV_WARN, True, True
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

Private Function StartAcquisition() As Boolean
    gT0 = Now: gDeadline = gT0 + TimeSerial(0, 0, TIMEOUT_S)
    gCandKey = "": gCandRows = -1: gLastErr = ""
    gKnown = SnapshotWorkbooks()
    gSince = gT0
    If Len(EXPORT_SEQ) = 0 Then Notify "EXPORT_SEQ is empty on CONFIG.", LV_ERROR: Exit Function
    If Not RunSteps(ExpandSeq(OPEN_SEQ, Date)) Then Notify "OPEN_SEQ failed: could not drive the Bloomberg Terminal.", LV_ERROR: Exit Function
    If Not RunSteps(ExpandSeq(EXPORT_SEQ, Date)) Then Notify "EXPORT_SEQ failed.", LV_ERROR: Exit Function
    Application.StatusBar = "BLOT: waiting for the export..."
    SchedulePoll
    StartAcquisition = True
End Function

Private Sub OnAcquireTimeout()
    Dim msg As String
    msg = "TIMEOUT: BLOT export not detected after " & TIMEOUT_S & "s (new Excel workbook" & _
          IIf(Len(EXPORT_FOLDER) > 0, " or new file in " & EXPORT_FOLDER, "") & ")."
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
            Set res = AnalyzeBlot(v, LoadDealerMap(), PrevBusinessDay(Date))
            detail = detail & " | " & res("m") & " request(s), " & res("nFilled") & " filled | " & res("mapping")
        End If
    End If
    Notify "TestExport cycle " & gTestI & "/" & gTestN & ": " & IIf(ok, "OK - ", "FAIL - ") & detail, IIf(ok, LV_INFO, LV_WARN), False, True
    If gTestI < gTestN Then
        gCycleAt = Now + TimeSerial(0, 0, 3)
        Application.OnTime gCycleAt, ProcName("StartTestCycle")
        gCyclePending = True
    Else
        FinishRun
        Notify "TestExport finished: " & gTestOK & "/" & gTestN & " cycle(s) succeeded.", IIf(gTestOK = gTestN, LV_INFO, LV_WARN), True, True
    End If
End Sub

'==========================================================================================
'                                    SOURCES
'==========================================================================================
Private Function TryAcquire() As Boolean
    If TryNewWorkbooks() Then TryAcquire = True: Exit Function
    If Len(EXPORT_FOLDER) > 0 Then TryAcquire = TryFile(NewestFileSince(EXPORT_FOLDER, gSince))
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
                        CloseExport wb, app
                        TryNewWorkbooks = True
                    End If
                    Exit Function
                ElseIf InStr(1, gWarned, k, vbTextCompare) = 0 Then
                    gWarned = gWarned & Mid$(k, 2)
                    WriteLog "New workbook " & wb.Name & " ignored: no BLOT header row.", LV_INFO
                End If
            End If
        Next
    Next
End Function

Private Function TryFile(ByVal f As String) As Boolean
    Dim wb As Workbook, blk As Object
    If Len(f) = 0 Then Exit Function
    If InStr(1, gKnown, "|" & f & "|", vbTextCompare) > 0 Then Exit Function
    If Not IsStable("F|" & f, FileLen(f)) Then Exit Function
    gKnown = gKnown & f & "|"
    Set blk = OpenExportFile(f, wb)
    If blk Is Nothing Then
        Notify "File " & f & " ignored: no BLOT header row.", LV_WARN
    Else
        HandleExport blk("v"), blk("raw"), "file " & f, Nothing, f
        TryFile = True
    End If
    If Not wb Is Nothing Then wb.Close False
End Function

' Opens an export file (CSV as text, Excel / Bloomberg XLS through Excel) and finds the BLOT block
Private Function OpenExportFile(ByVal f As String, ByRef wb As Workbook) As Object
    Dim da As Boolean, a As Variant, blk As Variant, d As Object
    If LCase$(Right$(f, 4)) = ".csv" Or LCase$(Right$(f, 4)) = ".txt" Then
        a = ParseCsvText(ReadTextFile(f), "AUTO")
        blk = FindBlockInArray(a)
        If IsEmpty(blk) Then Exit Function
        Set d = NewDict()
        d("v") = blk: d("raw") = blk: d("rows") = UBound(blk, 1): d("label") = Dir$(f)
        Set OpenExportFile = d
    Else
        da = Application.DisplayAlerts: Application.DisplayAlerts = False   ' Bloomberg XLS: "format and extension" prompt
        Set wb = Workbooks.Open(f, UpdateLinks:=0, ReadOnly:=True)
        Application.DisplayAlerts = da
        Set OpenExportFile = FindBlockInWorkbook(wb)
    End If
End Function

' Newest *.xls* / *.csv modified since 'since' (this workbook, Excel temp files and archive copies excluded)
Private Function NewestFileSince(ByVal folder As String, ByVal since As Date) As String
    Dim f As String, best As String, bt As Date, ft As Date, pat As Variant
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    For Each pat In Array("*.xls*", "*.csv")
        f = Dir$(folder & pat)
        Do While Len(f) > 0
            If Left$(f, 2) <> "~$" And StrComp(folder & f, ThisWorkbook.FullName, vbTextCompare) <> 0 _
               And Not (f Like "BLOT_########_########_######*") Then
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

'==========================================================================================
'                                    TERMINAL DRIVER (KEYS via DDE, SEND, WAIT)
'==========================================================================================
Public Function ExpandSeq(ByVal seq As String, ByVal refDate As Date) As String
    ExpandSeq = Replace(seq, "{DAYS_BACK}", CStr(DaysBack(refDate)), , , vbTextCompare)
End Function

Public Function ValidateSteps(ByVal seq As String) As String
    Dim st As Variant, s As String, tok As String, arg As String, p As Long
    For Each st In Split(seq, "|")
        s = Trim$(CStr(st))
        If Len(s) > 0 Then
            p = InStr(1, s, ":")
            If p = 0 Then ValidateSteps = "Bad step '" & s & "' (expected KEYS:..., SEND:... or WAIT:s).": Exit Function
            tok = UCase$(Left$(s, p - 1)): arg = Mid$(s, p + 1)
            Select Case tok
                Case "KEYS", "SEND"
                    If Len(arg) = 0 Then ValidateSteps = "Empty " & tok & " step.": Exit Function
                    If tok = "KEYS" And (UCase$(Trim$(arg)) Like "[A-Z]<GO>" Or InStr(1, arg, "<N>", vbTextCompare) > 0) Then
                        ValidateSteps = "'" & s & "' looks like a placeholder (a letter + <GO> runs a Bloomberg function, N = News).": Exit Function
                    End If
                Case "WAIT"
                    If Not IsPlainNumber(arg) Then ValidateSteps = "Bad " & s & " (seconds, '.' decimal).": Exit Function
                Case Else
                    ValidateSteps = "Unknown step '" & tok & "' (v5 supports KEYS, SEND, WAIT).": Exit Function
            End Select
        End If
    Next
End Function

Private Function RunSteps(ByVal seq As String) As Boolean
    Dim st As Variant, s As String, tok As String, arg As String, p As Long, e As String
    e = ValidateSteps(seq)
    If Len(e) > 0 Then Notify "Step sequence rejected: " & e, LV_ERROR: Exit Function
    For Each st In Split(seq, "|")
        s = Trim$(CStr(st))
        If Len(s) > 0 Then
            p = InStr(1, s, ":")
            tok = UCase$(Left$(s, p - 1)): arg = Mid$(s, p + 1)
            Select Case tok
                Case "KEYS"
                    If Not BbgType(arg) Then Notify "Step failed: " & s & " (no DDE link and no Bloomberg window).", LV_ERROR: Exit Function
                Case "SEND"
                    If Not BbgFocus() Then Notify "Step failed: " & s & " (window '" & BBG_TITLE_PREFIX & "' not found).", LV_ERROR: Exit Function
                    SendKeys arg, True
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

Private Function BbgHwnd() As LongPtr
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
    Dim h As LongPtr
    h = BbgHwnd(): If h = 0 Then Exit Function
    If IsIconic(h) <> 0 Then ShowWindow h, SW_RESTORE
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

' Header row = first of 40 rows with a dealer column and an ID or ISIN column
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
            If HeaderIndex(hdr, HDR_DEALER) > 0 And (HeaderIndex(hdr, HDR_ID) > 0 Or HeaderIndex(hdr, HDR_ISIN) > 0) Then
                FindHeaderRow = r: lastC = lc: Exit Function
            End If
        End If
    Next
End Function

Private Function SliceBlock(a As Variant, ByVal r0 As Long, ByVal nc As Long) As Variant
    Dim o() As Variant, i As Long, j As Long, n As Long
    n = UBound(a, 1) - r0 + 1
    ReDim o(1 To n, 1 To nc)
    For i = 1 To n
        For j = 1 To nc
            If j <= UBound(a, 2) Then o(i, j) = a(r0 + i - 1, j)
        Next
    Next
    SliceBlock = o
End Function

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
    Dim lines As Variant, i As Long, cand As Variant, best As String, bestN As Long, mx As Long, n As Long
    Select Case UCase$(Trim$(cfg))
        Case ";": CsvDelimiter = ";": Exit Function
        Case ",": CsvDelimiter = ",": Exit Function
        Case "TAB": CsvDelimiter = vbTab: Exit Function
    End Select
    txt = Replace(Replace(txt, vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(txt, vbLf, 21)
    best = ",": bestN = 0
    For Each cand In Array(vbTab, ";", ",")
        mx = 0
        For i = LBound(lines) To Application.WorksheetFunction.Min(UBound(lines), LBound(lines) + 19)
            n = Len(lines(i)) - Len(Replace(lines(i), cand, ""))
            If n > mx Then mx = n
        Next
        If mx > bestN Then bestN = mx: best = cand
    Next
    CsvDelimiter = best
End Function

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
'                                    ANALYSIS CORE (pure: array in, result out)
'==========================================================================================
' v = block, row 1 = header. Returns a Dictionary:
'   ok, err, mapping, warnings, m (units), units (m x UCOLS), nLegs, legs (nLegs x QCOLS),
'   nRequests, nFilled, nBreach, nLow, nKey, nQuotedLow, statuses, unmapped, minTD, maxTD
' Unit columns: 1 TradeDate | 2 Key | 3 Security | 4 ISIN | 5 Side | 6 Qty | 7 Requests | 8 Asked
'   9 Quoted | 10 Filled | 11 Winner | 12 FillPrice | 13 Check | 14 Dealers | 15 LoadedAt | 16 Source
' Leg columns: 1 TradeDate | 2 Key | 3 Request | 4 Security | 5 ISIN | 6 Side | 7 Qty | 8 BrkrName
'   9 Dealer (normalized) | 10 Status | 11 Price | 12 Quoted | 13 Filled
Public Function AnalyzeBlot(v As Variant, dmap As Object, ByVal defTD As Date) As Object
    Dim res As Object, hdr() As Variant, n As Long, nc As Long, i As Long, j As Long, u As Long, m As Long
    Dim cDl As Long, cID As Long, cSec As Long, cIsin As Long, cSide As Long, cSt As Long, cQty As Long, cPx As Long, cDt As Long
    Dim idx As Object, inqIsin As Object, inqSide As Object, inqDate As Object, badInq As Object, reqs As Object
    Dim bondRows As Object, bondInq As Object, stSeen As Object, unm As Object, ents As Object, kk As Variant
    Dim uDate() As Date, uKey() As String, uSec() As String, uIsin() As String, uSide() As String, uQty() As Variant
    Dim uInq() As Object, uAsk() As Object, uQuo() As Object, uFill() As Boolean, uWin() As String, uPx() As Variant
    Dim legs() As Variant, nLegs As Long, d As Date, dlr As String, nm As String, st As String, px As Double
    Dim isin As String, sec As String, sideU As String, qty As String, inq As String, uk As String, bk As String
    Dim isQ As Boolean, isF As Boolean, nNoDl As Long, nFb As Long, warn As String, s As String
    Dim out() As Variant, ord() As Long, rk() As Long, t As Long, chk As String, bad As Boolean
    Dim nFilled As Long, nBreach As Long, nLow As Long, nKey As Long, nQLow As Long, nSusp As Long
    Dim minTD As Date, maxTD As Date

    Set res = NewDict()
    res("ok") = False: res("err") = "": res("mapping") = "": res("warnings") = "": res("m") = 0
    res("units") = Empty: res("nLegs") = 0: res("legs") = Empty: res("nRequests") = 0
    res("nFilled") = 0: res("nBreach") = 0: res("nLow") = 0: res("nKey") = 0: res("nQuotedLow") = 0
    res("statuses") = "": res("unmapped") = "": res("minTD") = 0: res("maxTD") = 0
    If Not IsArray(v) Then res("err") = "No data.": GoTo Done
    n = UBound(v, 1): nc = UBound(v, 2)
    ReDim hdr(1 To 1, 1 To nc)
    For j = 1 To nc: hdr(1, j) = v(1, j): Next

    cDl = HeaderIndex(hdr, HDR_DEALER):   cID = HeaderIndex(hdr, HDR_ID)
    cSec = HeaderIndex(hdr, HDR_SECURITY): cIsin = HeaderIndex(hdr, HDR_ISIN)
    cSide = HeaderIndex(hdr, HDR_SIDE):   cSt = HeaderIndex(hdr, HDR_STATUS)
    cQty = HeaderIndex(hdr, HDR_QTY):     cPx = HeaderIndex(hdr, HDR_PRICE)
    cDt = HeaderIndex(hdr, HDR_DATE)
    res("mapping") = "Columns: " & MapDesc(hdr, Array("DEALER", cDl, "ID", cID, "SECURITY", cSec, "ISIN", cIsin, "SIDE", cSide, _
                     "STATUS", cSt, "QTY", cQty, "PRICE", cPx, "DATE", cDt))
    If cDl = 0 Then res("err") = "Dealer column not found (HDR_DEALER).": GoTo Done
    If cID = 0 And cIsin = 0 Then res("err") = "Neither an ID nor an ISIN column found (HDR_ID / HDR_ISIN).": GoTo Done
    If cID = 0 Then warn = warn & "No ID column: requests grouped by trade date | ISIN | side | qty. "
    If cSt = 0 Then warn = warn & "No status column: no request can be recognised as filled. "
    If cPx = 0 Then warn = warn & "No price column: 'quoted' cannot be measured. "

    Set idx = NewDict(): Set inqIsin = NewDict(): Set inqSide = NewDict(): Set inqDate = NewDict()
    Set badInq = NewDict(): Set reqs = NewDict(): Set bondRows = NewDict(): Set bondInq = NewDict()
    Set stSeen = NewDict(): Set unm = NewDict(): Set ents = NewDict()
    For Each kk In dmap.Keys: ents(dmap(kk)) = 1: Next
    ReDim uDate(1 To n): ReDim uKey(1 To n): ReDim uSec(1 To n): ReDim uIsin(1 To n): ReDim uSide(1 To n)
    ReDim uQty(1 To n): ReDim uInq(1 To n): ReDim uAsk(1 To n): ReDim uQuo(1 To n): ReDim uFill(1 To n)
    ReDim uWin(1 To n): ReDim uPx(1 To n): ReDim legs(1 To n, 1 To QCOLS)

    For i = 2 To n
        dlr = CellS(v, i, cDl)
        If Len(dlr) = 0 Then
            If Len(CellS(v, i, cID)) + Len(CellS(v, i, cIsin)) + Len(CellS(v, i, cSec)) > 0 Then nNoDl = nNoDl + 1
            GoTo NextRow
        End If
        If cDt > 0 Then
            If Not ParseDateCell(v(i, cDt), DATE_FORMAT_TEXT, d) Then d = defTD: nFb = nFb + 1
        Else
            d = defTD: nFb = nFb + 1
        End If
        isin = UCase$(CellS(v, i, cIsin)): sec = CellS(v, i, cSec)
        sideU = UCase$(CellS(v, i, cSide)): qty = CellS(v, i, cQty)
        inq = RequestKey(CellS(v, i, cID), d, IIf(Len(isin) > 0, isin, UCase$(sec)), sideU, qty)
        reqs(inq) = 1
        SameOrBad inqIsin, badInq, inq, isin
        SameOrBad inqSide, badInq, inq, sideU
        SameOrBad inqDate, badInq, inq, Format$(d, "yyyymmdd")

        If GROUP_LEVEL = "ORDER" Then
            uk = "ORD:" & Format$(d, "yyyy-mm-dd") & "|" & IIf(Len(isin) > 0, isin, UCase$(sec)) & "|" & sideU
        Else
            uk = inq
        End If
        If Not idx.Exists(uk) Then
            m = m + 1: idx(uk) = m
            uDate(m) = d: uKey(m) = uk: uSec(m) = sec: uIsin(m) = isin: uSide(m) = CellS(v, i, cSide)
            uQty(m) = NumOrText(v, i, cQty)
            Set uInq(m) = NewDict(): Set uAsk(m) = NewDict(): Set uQuo(m) = NewDict()
            uPx(m) = Empty
        End If
        u = idx(uk)
        uInq(u)(inq) = 1
        If Len(uSec(u)) = 0 Then uSec(u) = sec
        If Len(uIsin(u)) = 0 Then uIsin(u) = isin

        nm = NormName(dlr)
        If dmap.Exists(nm) Then
            nm = dmap(nm)
        ElseIf Not ents.Exists(nm) Then
            unm(nm) = 1
        End If
        st = CellS(v, i, cSt)
        If Len(st) = 0 Then s = "(blank)" Else s = st
        stSeen(s) = stSeen(s) + 1
        px = 0
        If cPx > 0 Then px = PriceOf(v(i, cPx))
        isQ = (px > 0)
        isF = IsFilledStatus(st)
        If isQ Then s = s & " @" & InvariantNum(px)
        If uAsk(u).Exists(nm) Then uAsk(u)(nm) = uAsk(u)(nm) & " / " & s Else uAsk(u)(nm) = s
        If isQ Then uQuo(u)(nm) = 1
        If isF Then
            uFill(u) = True
            If InStr(1, ";" & uWin(u) & ";", ";" & nm & ";") = 0 Then
                If Len(uWin(u)) > 0 Then uWin(u) = uWin(u) & ";"
                uWin(u) = uWin(u) & nm
            End If
            If IsEmpty(uPx(u)) And isQ Then uPx(u) = px
        End If

        bk = Format$(d, "yyyymmdd") & "|" & isin & "|" & sideU & "|" & qty
        bondRows(bk) = bondRows(bk) + 1
        If Not bondInq.Exists(bk) Then bondInq.Add bk, NewDict()
        bondInq(bk)(inq) = 1

        nLegs = nLegs + 1
        legs(nLegs, 1) = d: legs(nLegs, 2) = uk: legs(nLegs, 3) = inq: legs(nLegs, 4) = sec: legs(nLegs, 5) = isin
        legs(nLegs, 6) = CellS(v, i, cSide): legs(nLegs, 7) = NumOrText(v, i, cQty): legs(nLegs, 8) = dlr
        legs(nLegs, 9) = nm: legs(nLegs, 10) = st
        If isQ Then legs(nLegs, 11) = px Else legs(nLegs, 11) = ""
        legs(nLegs, 12) = IIf(isQ, "YES", "NO"): legs(nLegs, 13) = IIf(isF, "YES", "NO")
NextRow:
    Next

    ' one request per dealer row for the same bond/side/day/qty = the ID rule is probably wrong
    For Each kk In bondRows.Keys
        If bondRows(kk) >= 2 And bondInq(kk).Count = bondRows(kk) Then nSusp = nSusp + 1
    Next

    ' output, worst first: BREACH, KEY?, LOW, OK filled, OK not filled
    If m > 0 Then
        ReDim out(1 To m, 1 To UCOLS): ReDim rk(1 To m): ReDim ord(1 To m)
        For u = 1 To m
            bad = False
            For Each kk In uInq(u).Keys
                If badInq.Exists(kk) Then bad = True
            Next
            If bad Then
                chk = "KEY?": nKey = nKey + 1: rk(u) = 1
            ElseIf uAsk(u).Count < MIN_Q Then
                If uFill(u) Then
                    chk = "BREACH": nBreach = nBreach + 1: rk(u) = 0
                Else
                    chk = "LOW": nLow = nLow + 1: rk(u) = 2
                End If
            Else
                chk = "OK"
                If uFill(u) Then rk(u) = 3 Else rk(u) = 4
            End If
            If uFill(u) Then nFilled = nFilled + 1
            If uQuo(u).Count < MIN_Q Then nQLow = nQLow + 1
            out(u, 1) = uDate(u): out(u, 2) = uKey(u): out(u, 3) = uSec(u): out(u, 4) = uIsin(u): out(u, 5) = uSide(u)
            out(u, 6) = uQty(u): out(u, 7) = uInq(u).Count: out(u, 8) = uAsk(u).Count: out(u, 9) = uQuo(u).Count
            out(u, 10) = IIf(uFill(u), "YES", "NO"): out(u, 11) = uWin(u)
            If IsEmpty(uPx(u)) Then out(u, 12) = "" Else out(u, 12) = uPx(u)
            out(u, 13) = chk: out(u, 14) = DealerDetail(uAsk(u)): out(u, 15) = Empty: out(u, 16) = ""
            If minTD = 0 Or uDate(u) < minTD Then minTD = uDate(u)
            If uDate(u) > maxTD Then maxTD = uDate(u)
            ord(u) = u
        Next
        For i = 2 To m                                        ' stable insertion sort by rank
            t = ord(i): j = i - 1
            Do While j >= 1
                If rk(ord(j)) <= rk(t) Then Exit Do
                ord(j + 1) = ord(j): j = j - 1
            Loop
            ord(j + 1) = t
        Next
        res("units") = Reorder(out, ord, m)
    End If
    If nLegs > 0 Then res("legs") = TrimRows(legs, nLegs)

    If nNoDl > 0 Then warn = warn & nNoDl & " row(s) without dealer ignored. "
    If nFb > 0 Then warn = warn & nFb & " row(s) without a readable trade date: " & Format$(defTD, "yyyy-mm-dd") & " used (check DATE_FORMAT_TEXT). "
    If nSusp > 0 Then warn = warn & nSusp & " bond/side/day group(s) have one request per dealer row: check ID_SUFFIX_LEN. "
    If nKey > 0 Then warn = warn & nKey & " request(s) mix ISINs, sides or dates under one ID (KEY?). "
    res("ok") = True
    res("m") = m: res("nLegs") = nLegs: res("nRequests") = reqs.Count
    res("nFilled") = nFilled: res("nBreach") = nBreach: res("nLow") = nLow: res("nKey") = nKey: res("nQuotedLow") = nQLow
    res("statuses") = CountList(stSeen): res("unmapped") = KeyList(unm, ", ")
    res("minTD") = minTD: res("maxTD") = maxTD
Done:
    res("warnings") = Trim$(warn)
    Set AnalyzeBlot = res
End Function

' Request key: ID without its dealer-leg suffix; without ID: trade date | ISIN | side | qty
Public Function RequestKey(ByVal id As String, ByVal d As Date, ByVal isin As String, ByVal sideU As String, ByVal qty As String) As String
    id = Trim$(id)
    If Len(id) > 0 Then
        If Len(id) > ID_SUFFIX_LEN + 3 Then RequestKey = "INQ:" & Left$(id, Len(id) - ID_SUFFIX_LEN) Else RequestKey = "INQ:" & id
    Else
        RequestKey = "CMP:" & Format$(d, "yyyy-mm-dd") & "|" & isin & "|" & sideU & "|" & qty
    End If
End Function

Private Sub SameOrBad(seen As Object, bad As Object, ByVal k As String, ByVal val As String)
    If Len(val) = 0 Then Exit Sub
    If Not seen.Exists(k) Then
        seen(k) = val
    ElseIf seen(k) <> val Then
        bad(k) = 1
    End If
End Sub

Public Function IsFilledStatus(ByVal st As String) As Boolean
    Dim a As Variant, t As String
    st = LCase$(Trim$(st))
    If Len(st) = 0 Then Exit Function
    For Each a In Split(FILL_STATUS, "|")
        t = Trim$(CStr(a))
        If Len(t) > 0 Then
            If InStr(1, st, t, vbBinaryCompare) > 0 Then IsFilledStatus = True: Exit Function
        End If
    Next
End Function

' Price from a cell: real number, or text with ',' or '.' decimal (96,198 / 96.198 / 1.234,5)
Public Function PriceOf(ByVal x As Variant) As Double
    Dim s As String, pc As Long, pd As Long
    Select Case VarType(x)
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte
            PriceOf = CDbl(x)
        Case vbString
            s = Replace(Replace(Trim$(x), " ", ""), Chr$(160), "")
            If Len(s) = 0 Then Exit Function
            pc = InStrRev(s, ","): pd = InStrRev(s, ".")
            If pc > 0 And pd > 0 Then
                If pc > pd Then s = Replace(Replace(s, ".", ""), ",", ".") Else s = Replace(s, ",", "")
            ElseIf pc > 0 Then
                If Len(s) - Len(Replace(s, ",", "")) = 1 Then s = Replace(s, ",", ".") Else s = Replace(s, ",", "")
            End If
            If IsPlainNumber(s) Then PriceOf = Val(s)
    End Select
End Function

Private Function NumOrText(v As Variant, ByVal i As Long, ByVal c As Long) As Variant
    Dim x As Variant
    NumOrText = ""
    If c = 0 Then Exit Function
    x = v(i, c)
    Select Case VarType(x)
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte
            NumOrText = CDbl(x)
        Case vbString
            If PriceOf(x) <> 0 Then NumOrText = PriceOf(x) Else NumOrText = Trim$(x)
    End Select
End Function

Private Function DealerDetail(d As Object) As String
    Dim k As Variant, s As String
    For Each k In d.Keys
        s = s & "; " & k & " [" & d(k) & "]"
    Next
    DealerDetail = Mid$(s, 3)
End Function

Private Function CountList(d As Object) As String
    Dim k As Variant, s As String
    For Each k In d.Keys: s = s & ", " & k & "=" & d(k): Next
    CountList = Mid$(s, 3)
End Function

Private Function KeyList(d As Object, ByVal sep As String) As String
    Dim k As Variant, s As String
    For Each k In d.Keys: s = s & sep & k: Next
    If Len(s) > 0 Then KeyList = Mid$(s, Len(sep) + 1)
End Function

Private Function Reorder(a As Variant, ord() As Long, ByVal m As Long) As Variant
    Dim o() As Variant, i As Long, j As Long
    ReDim o(1 To m, 1 To UBound(a, 2))
    For i = 1 To m
        For j = 1 To UBound(a, 2): o(i, j) = a(ord(i), j): Next
    Next
    Reorder = o
End Function

Private Function TrimRows(a As Variant, ByVal n As Long) As Variant
    Dim o() As Variant, i As Long, j As Long
    ReDim o(1 To n, 1 To UBound(a, 2))
    For i = 1 To n
        For j = 1 To UBound(a, 2): o(i, j) = a(i, j): Next
    Next
    TrimRows = o
End Function

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

Public Function HeaderIndex(hdr As Variant, ByVal aliases As String) As Long
    Dim a As Variant, t As String, j As Long
    For Each a In Split(aliases, "|")
        t = LCase$(NormName(CStr(a)))
        If Len(t) > 0 Then
            For j = LBound(hdr, 2) To UBound(hdr, 2)
                If NormH(hdr(1, j)) = t Then HeaderIndex = j: Exit Function
            Next
        End If
    Next
    For Each a In Split(aliases, "|")
        t = LCase$(NormName(CStr(a)))
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
    NormH = LCase$(NormName(CStr(x)))
End Function

Private Function MapDesc(hdr As Variant, roles As Variant) As String
    Dim i As Long, s As String
    For i = LBound(roles) To UBound(roles) Step 2
        If roles(i + 1) > 0 Then s = s & roles(i) & "='" & NormH(hdr(1, roles(i + 1))) & "' " Else s = s & roles(i) & "=- "
    Next
    MapDesc = Trim$(s)
End Function

Public Function CellS(v As Variant, ByVal i As Long, ByVal c As Long) As String
    Dim x As Variant
    If c = 0 Then Exit Function
    If c > UBound(v, 2) Then Exit Function
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

Public Function ValidateDateFormat(ByVal fmt As String) As String
    Dim runs As Collection, sepd As Boolean, r As Variant, nD As Long, nM As Long, nY As Long
    Set runs = FormatRuns(UCase$(Trim$(fmt)), sepd)
    If runs Is Nothing Then ValidateDateFormat = "Unsupported format '" & fmt & "' (e.g. DD/MM/YYYY, MM/DD/YYYY, DD-MMM-YY).": Exit Function
    For Each r In runs
        Select Case Left$(r, 1)
            Case "D": nD = nD + 1
            Case "M": nM = nM + 1
            Case "Y": nY = nY + 1
        End Select
    Next
    If nD <> 1 Or nM <> 1 Or nY <> 1 Then ValidateDateFormat = "Format '" & fmt & "' needs one day, one month and one year part."
End Function

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
    If gaps > 0 And nSep > 0 Then Exit Function
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
    If p = 0 Then p = InStr(1, "JANFEVMARABRMAIJUNJULAGOSETOUTNOVDEZ", t)
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
'                                    PROCESS (I/O around the core)
'==========================================================================================
Private Function ProcessData(v As Variant, vRaw As Variant, ByVal srcLabel As String, ByVal srcWb As Object, ByVal srcFile As String) As Boolean
    Dim res As Object, out As Variant, m As Long, arch As String, i As Long, ts As Date, prevMax As Double
    Dim dates As Object, msg As String, lst As String, gaps As String, info As String
    On Error GoTo EH
    Set res = AnalyzeBlot(v, LoadDealerMap(), PrevBusinessDay(Date))
    WriteLog "Read " & srcLabel & ": " & res("nLegs") & " dealer row(s), " & res("nRequests") & " request(s). " & res("mapping"), LV_INFO
    If Len(res("warnings")) > 0 Then Notify "Data check (" & srcLabel & "): " & res("warnings"), LV_WARN
    If Not res("ok") Then Notify "ERROR reading " & srcLabel & ": " & res("err"), LV_ERROR: Exit Function
    If Len(res("unmapped")) > 0 Then WriteLog "Dealer names not in DEALER_MAP: " & res("unmapped"), LV_INFO
    WriteLog "Statuses seen: " & res("statuses"), LV_INFO

    m = res("m"): ts = Now
    If m > 0 Then arch = ArchiveExport(srcWb, srcFile, vRaw, res("maxTD")) Else arch = ArchiveExport(srcWb, srcFile, vRaw, PrevBusinessDay(Date))
    If m = 0 Then
        WriteSheetTable GetSheet(SH_QUOTES), QuoteHeader(), Empty, 0
        Notify "ALERT: the export has no dealer rows (" & srcLabel & "). Confirm that the desk did not trade.", LV_ALERT
        ProcessData = True: Exit Function
    End If
    out = res("units")
    For i = 1 To m
        out(i, 15) = ts
        If Len(arch) > 0 Then out(i, 16) = arch Else out(i, 16) = srcLabel
    Next
    prevMax = HistDailyMaxDate()
    WriteSheetTable GetSheet(SH_QUOTES), QuoteHeader(), res("legs"), res("nLegs")
    WriteSheetTable GetSheet(SH_INQ), UnitHeader(), out, m
    Set dates = UpsertHist(GetSheet(SH_HIST), out, m)
    RecalcHistDaily GetSheet(SH_HIST), GetSheet(SH_HD), dates
    gaps = CoverageGaps(prevMax, res("minTD"))
    info = "Last run " & Format$(ts, "yyyy-mm-dd hh:nn") & " | " & srcLabel & " | " & GROUP_LEVEL & " level | " & res("nLegs") & " dealer rows"
    BuildDashboard out, m, res, info

    msg = "Trades " & DateRangeText(res("minTD"), res("maxTD")) & ": " & m & " request(s), " & res("nFilled") & " filled, " & _
          res("nBreach") & " BREACH (filled, < " & MIN_Q & " asked), " & res("nLow") & " LOW (not filled, < " & MIN_Q & " asked), " & res("nKey") & " KEY?"
    If res("nBreach") > 0 Or res("nKey") > 0 Then
        For i = 1 To m
            If out(i, 13) = "BREACH" Or out(i, 13) = "KEY?" Then
                lst = lst & vbLf & "[" & out(i, 13) & "] " & Format$(out(i, 1), "yyyy-mm-dd") & "  " & out(i, 3) & "  " & out(i, 4) & "  " & out(i, 5) & _
                      "  asked=" & out(i, 8) & " quoted=" & out(i, 9) & "  winner=" & out(i, 11) & "  " & out(i, 14)
            End If
        Next
        Notify "ALERT " & msg & lst, LV_ALERT
    Else
        Notify "OK " & msg, LV_INFO
    End If
    If Len(gaps) > 0 Then Notify "ALERT coverage gap: no BLOT data for business day(s) " & gaps & ". Load them with Import File.", LV_ALERT
    ProcessData = True
    Exit Function
EH:
    Notify "ERROR in ProcessData (" & srcLabel & "): " & Err.Description, LV_ERROR
End Function

Private Function DateRangeText(ByVal d0 As Date, ByVal d1 As Date) As String
    If d0 = d1 Then DateRangeText = Format$(d0, "yyyy-mm-dd") Else DateRangeText = Format$(d0, "yyyy-mm-dd") & ".." & Format$(d1, "yyyy-mm-dd")
End Function

Private Function ArchiveExport(ByVal srcWb As Object, ByVal srcFile As String, vRaw As Variant, ByVal td As Date) As String
    Dim base As String, ext As String, p As String, wb As Workbook
    If Len(ARCHIVE_DIR) = 0 Then
        Notify "ALERT: ARCHIVE_DIR is empty on CONFIG - the export was NOT archived.", LV_ALERT
        Exit Function
    End If
    EnsureFolder ARCHIVE_DIR
    base = ARCHIVE_DIR & "\BLOT_" & Format$(td, "yyyymmdd") & "_" & Format$(Now, "yyyymmdd_hhnnss")
    On Error GoTo Fallback
    If Len(srcFile) > 0 Then
        p = base & LCase$(Mid$(srcFile, InStrRev(srcFile, ".")))
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
    Notify "ALERT: could not archive the export to " & ARCHIVE_DIR & " (" & Err.Description & ").", LV_ALERT
    On Error GoTo -1
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
End Function

'==========================================================================================
'                                    OUTPUT SHEETS
'==========================================================================================
Public Function UnitHeader() As Variant
    UnitHeader = Array("TradeDate", "Key", "Security", "ISIN", "Side", "Qty", "Requests", "Asked", "Quoted", "Filled", _
                       "Winner", "FillPrice", "Check", "Dealers [status @price]", "LoadedAt", "Source")
End Function

Public Function QuoteHeader() As Variant
    QuoteHeader = Array("TradeDate", "Key", "Request", "Security", "ISIN", "Side", "Qty", "BrkrName", "Dealer", "Status", "Price", "Quoted", "Filled")
End Function

Public Function DailyHeader() As Variant
    DailyHeader = Array("TradeDate", "Requests", "Filled", "Breach", "Low", "KeyIssues", "PctFilledCompliant", "AvgAsked", "UpdatedAt")
End Function

Public Sub EnsureHeader(ws As Worksheet, h As Variant)
    Dim n As Long
    n = UBound(h) - LBound(h) + 1
    If Len(ws.Range("A1").Value) = 0 Or ws.Cells(1, n).Value <> h(UBound(h)) Then
        ws.Range("A1").Resize(1, n).Value = h
        Hdr ws.Range("A1").Resize(1, n)
        ws.Columns("A").NumberFormat = "yyyy-mm-dd": ws.Columns("A").ColumnWidth = 11
        If h(1) = "Key" Then
            ws.Columns("B").NumberFormat = "@": ws.Columns("B").ColumnWidth = 30
            ws.Columns("C").ColumnWidth = 26: ws.Columns("D").ColumnWidth = 14
            ws.Columns("N").ColumnWidth = 80: ws.Columns("O").NumberFormat = "yyyy-mm-dd hh:mm"
            ws.Columns("O").ColumnWidth = 16: ws.Columns("P").ColumnWidth = 50
        ElseIf h(1) = "Requests" Then
            ws.Columns("G").NumberFormat = "0.0%": ws.Columns("H").NumberFormat = "0.0"
            ws.Columns("I").NumberFormat = "yyyy-mm-dd hh:mm": ws.Columns("B:I").ColumnWidth = 12
        End If
    End If
End Sub

Private Sub EnsureLogHeader(ws As Worksheet)
    If Len(ws.Range("A1").Value) = 0 Then
        ws.Range("A1:C1").Value = Array("Timestamp", "Level", "Message")
        Hdr ws.Range("A1:C1")
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

Private Sub WriteSheetTable(ws As Worksheet, h As Variant, data As Variant, ByVal n As Long)
    Dim nc As Long
    nc = UBound(h) - LBound(h) + 1
    ws.Cells.Clear
    ws.Range("A1").Resize(1, nc).Value = h
    Hdr ws.Range("A1").Resize(1, nc)
    ws.Columns("A").NumberFormat = "yyyy-mm-dd"
    If n > 0 Then ws.Range("A2").Resize(n, nc).Value = TextSafe(data)
    ws.Range("A1").Resize(1, nc).EntireColumn.AutoFit
    If nc = UCOLS Then
        ws.Columns(14).ColumnWidth = 80: ws.Columns(15).NumberFormat = "yyyy-mm-dd hh:mm": ws.Columns(16).ColumnWidth = 50
    End If
End Sub

' One row per (TradeDate, Key): re-running a day updates rows, never duplicates. Returns the dates touched.
Public Function UpsertHist(ws As Worksheet, out As Variant, ByVal m As Long) As Object
    Dim ex As Object, hv As Variant, lastR As Long, i As Long, j As Long, rw As Long, k As String
    Dim one() As Variant, dates As Object, ov As Variant
    EnsureHeader ws, UnitHeader()
    Set ex = NewDict(): Set dates = NewDict()
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR >= 2 Then
        hv = ws.Range("A1:B" & lastR).Value
        For i = 2 To UBound(hv, 1): ex(HistKey(hv(i, 1), hv(i, 2))) = i: Next
    End If
    ReDim one(1 To 1, 1 To UCOLS)
    ov = TextSafe(out)
    For i = 1 To m
        For j = 1 To UCOLS: one(1, j) = ov(i, j): Next
        k = HistKey(out(i, 1), out(i, 2))
        If ex.Exists(k) Then
            rw = ex(k)
        Else
            lastR = lastR + 1: rw = lastR: ex(k) = rw
        End If
        ws.Cells(rw, 1).Resize(1, UCOLS).Value = one
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

' HIST_DAILY rows of the given dates recomputed from HIST (the audit trail is the source of truth)
Public Sub RecalcHistDaily(wsH As Worksheet, wsD As Worksheet, dates As Object)
    Dim hv As Variant, lastR As Long, i As Long, dk As Long, agg As Object, a As Variant, k As Variant
    Dim dv As Variant, rw As Long, rowOf As Object, lastD As Long, ck As String, pct As Variant
    EnsureHeader wsD, DailyHeader()
    lastR = wsH.Cells(wsH.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Sub
    hv = wsH.Range("A1:M" & lastR).Value2
    Set agg = NewDict()
    For i = 2 To UBound(hv, 1)
        If VarType(hv(i, 1)) = vbDouble Then
            dk = CLng(Int(hv(i, 1)))
            If dates.Exists(dk) Then
                If Not agg.Exists(dk) Then agg(dk) = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#)  ' n, filled, breach, low, key, sumAsked, filledOK
                a = agg(dk)
                ck = CStr(hv(i, 13))
                a(0) = a(0) + 1
                If CStr(hv(i, 10)) = "YES" Then
                    a(1) = a(1) + 1
                    If ck = "OK" Then a(6) = a(6) + 1
                End If
                If ck = "BREACH" Then a(2) = a(2) + 1
                If ck = "LOW" Then a(3) = a(3) + 1
                If ck = "KEY?" Then a(4) = a(4) + 1
                a(5) = a(5) + Val(CStr(hv(i, 8)))
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
        If a(1) > 0 Then pct = a(6) / a(1) Else pct = ""
        wsD.Cells(rw, 1).Resize(1, 9).Value = Array(CDate(k), a(0), a(1), a(2), a(3), a(4), pct, a(5) / a(0), Now)
    Next
    If lastD > 2 Then wsD.Range("A1").Resize(lastD, 9).Sort Key1:=wsD.Range("A1"), Order1:=xlAscending, Header:=xlYes
End Sub

Private Function HistDailyMaxDate() As Double
    On Error Resume Next
    HistDailyMaxDate = Application.WorksheetFunction.Max(GetSheet(SH_HD).Range("A:A"))
End Function

Private Function CoverageGaps(ByVal prevMax As Double, ByVal newMin As Date) As String
    Dim d As Date, s As String, n As Long
    If prevMax < 1 Or CDbl(newMin) <= prevMax + 1 Then Exit Function
    For d = CDate(prevMax) + 1 To newMin - 1
        If IsBusinessDay(d) Then
            n = n + 1
            If n <= 30 Then s = s & Format$(d, "yyyy-mm-dd") & " "
        End If
    Next
    CoverageGaps = Trim$(s)
End Function

'==========================================================================================
'                                    DASHBOARD
'==========================================================================================
Private Sub BuildDashboard(out As Variant, ByVal m As Long, ByVal res As Object, ByVal info As String)
    Dim ws As Worksheet, i As Long, r As Long, q As Long, mx As Long, nFill As Long, nOkFill As Long
    Dim dist() As Long, hd As Worksheet, lr As Long, hv As Variant, rr As Long, c As Long, tbl() As Variant
    Dim t0 As Date, t1 As Date, unitWord As String

    Set ws = GetSheet(SH_DASH)
    ws.Cells.Clear
    On Error Resume Next: ws.Cells.FormatConditions.Delete: On Error GoTo 0
    AddSheetButtons ws, DashButtons()
    ws.Columns("A").ColumnWidth = 2
    ws.Range("B3").Value = "BLOT Best Execution  |  every quote request must ask at least " & MIN_Q & " counterparties (Banco de Portugal)"
    ws.Range("B3").Font.Size = 15: ws.Range("B3").Font.Bold = True: ws.Range("B3").Font.Color = RGB(31, 56, 100)
    ws.Range("B4").Value = "'" & info
    ws.Range("B4").Font.Italic = True: ws.Range("B4").Font.Color = RGB(89, 89, 89)
    If m = 0 Or res Is Nothing Then
        ws.Range("B6").Value = "No data yet: Run Now, or paste the BLOT grid on PASTE and click Process Pasted Data."
        ws.Columns("B").ColumnWidth = 40
        Exit Sub
    End If
    If GROUP_LEVEL = "ORDER" Then unitWord = "Orders" Else unitWord = "Quote requests"

    '--- KPIs ---
    For i = 1 To m
        If out(i, 10) = "YES" Then
            nFill = nFill + 1
            If out(i, 13) = "OK" Then nOkFill = nOkFill + 1
        End If
        If out(i, 8) > mx Then mx = out(i, 8)
        If t0 = 0 Or out(i, 1) < t0 Then t0 = out(i, 1)
        If out(i, 1) > t1 Then t1 = out(i, 1)
    Next
    ws.Range("B6:B14").Value = Application.Transpose(Array("Trade date", unitWord, "Filled", "Not filled", _
        "BREACH: filled with < " & MIN_Q & " asked", "LOW: not filled, < " & MIN_Q & " asked", _
        "Requests with < " & MIN_Q & " dealers quoting", "KEY?: ambiguous grouping", "% filled trades compliant"))
    ws.Range("C6").Value = "'" & DateRangeText(t0, t1)
    ws.Range("C7").Value = m
    ws.Range("C8").Value = nFill
    ws.Range("C9").Value = m - nFill
    ws.Range("C10").Value = res("nBreach")
    ws.Range("C11").Value = res("nLow")
    ws.Range("C12").Value = res("nQuotedLow")
    ws.Range("C13").Value = res("nKey")
    If nFill > 0 Then ws.Range("C14").Value = nOkFill / nFill Else ws.Range("C14").Value = "'-"
    ws.Range("C14").NumberFormat = "0.0%"
    ws.Range("B6:C14").Borders.Color = RGB(217, 217, 217)
    ws.Range("C6:C14").HorizontalAlignment = xlRight: ws.Range("C6:C14").Font.Bold = True
    If res("nBreach") > 0 Then ws.Range("B10:C10").Interior.Color = RGB(255, 199, 206)
    If res("nKey") > 0 Then ws.Range("B13:C13").Interior.Color = RGB(255, 235, 156)

    With ws.Range("B16:C16")
        .Merge
        If res("nBreach") > 0 Or res("nKey") > 0 Then
            .Value = "ALERT: " & res("nBreach") & " filled trade(s) with fewer than " & MIN_Q & " counterparties asked" & IIf(res("nKey") > 0, ", " & res("nKey") & " KEY?", "")
            .Interior.Color = RGB(192, 0, 0)
        Else
            .Value = "OK: every filled trade asked at least " & MIN_Q & " counterparties"
            .Interior.Color = RGB(0, 128, 0)
        End If
        .Font.Color = vbWhite: .Font.Bold = True: .HorizontalAlignment = xlCenter
    End With

    '--- rule / legend ---
    ws.Range("E6").Value = "How the check works": ws.Range("E6").Font.Bold = True
    ws.Range("E7").Value = "'One quote request = the dealer rows whose ID differs only in the last " & ID_SUFFIX_LEN & " characters (FITDNA:6AF5A884011800|03, |05, |06)."
    ws.Range("E8").Value = "'Asked = distinct dealers on the request, whatever their status (Rejected, C/Cancel and Passed dealers were asked)."
    ws.Range("E9").Value = "'Quoted = dealers whose row has a price.  Filled = a row whose status matches FILL_STATUS (" & FILL_STATUS & ")."
    ws.Range("E10").Value = "'BREACH = filled with < " & MIN_Q & " asked.  LOW = not filled with < " & MIN_Q & " asked.  KEY? = one ID with two ISINs / sides / dates."
    ws.Range("E11").Value = "'Statuses seen in this export: " & res("statuses")
    ws.Range("E11").Font.Color = RGB(31, 56, 100): ws.Range("E11").Font.Bold = True
    If Len(res("warnings")) > 0 Then
        ws.Range("E12").Value = "'Data check: " & res("warnings")
        ws.Range("E12").Font.Color = RGB(192, 0, 0)
    End If
    ws.Range("E7:E12").Font.Size = 9

    '--- distribution of dealers asked ---
    ReDim dist(0 To mx)
    For i = 1 To m: dist(out(i, 8)) = dist(out(i, 8)) + 1: Next
    ws.Range("E14").Value = "Dealers asked": ws.Range("F14").Value = unitWord
    Hdr ws.Range("E14:F14")
    For q = 0 To mx
        ws.Cells(15 + q, 5).Value = q: ws.Cells(15 + q, 6).Value = dist(q)
        If q < MIN_Q Then ws.Cells(15 + q, 5).Resize(1, 2).Interior.Color = RGB(255, 199, 206)
    Next
    ws.Range(ws.Cells(15, 6), ws.Cells(15 + mx, 6)).FormatConditions.AddDatabar

    '--- recent days ---
    Set hd = GetSheet(SH_HD)
    lr = hd.Cells(hd.Rows.Count, 1).End(xlUp).Row
    ws.Range("I14:N14").Value = Array("Recent days", "Requests", "Filled", "Breach", "Low", "KEY?")
    Hdr ws.Range("I14:N14")
    rr = 14
    If lr >= 2 Then
        For i = lr To Application.WorksheetFunction.Max(2, lr - 7) Step -1
            hv = hd.Range(hd.Cells(i, 1), hd.Cells(i, 6)).Value
            rr = rr + 1
            For c = 1 To 6: ws.Cells(rr, 8 + c).Value = hv(1, c): Next
            ws.Cells(rr, 9).NumberFormat = "yyyy-mm-dd"
            If Val(CStr(hv(1, 4))) > 0 Then ws.Range(ws.Cells(rr, 9), ws.Cells(rr, 14)).Interior.Color = RGB(255, 199, 206)
        Next
    End If

    '--- requests table ---
    r = Application.WorksheetFunction.Max(18, 16 + mx, rr) + 2
    ws.Cells(r, 2).Value = unitWord & " (worst first)": ws.Cells(r, 2).Font.Bold = True: ws.Cells(r, 2).Font.Size = 12
    r = r + 1
    ws.Cells(r, 2).Resize(1, 12).Value = Array("Trade date", "Security", "ISIN", "Side", "Qty", "Asked", "Quoted", "Filled", "Winner", "Price", "Check", "Dealers [status @price]")
    Hdr ws.Cells(r, 2).Resize(1, 12)
    ReDim tbl(1 To m, 1 To 12)
    For i = 1 To m
        tbl(i, 1) = out(i, 1): tbl(i, 2) = "'" & out(i, 3): tbl(i, 3) = "'" & out(i, 4): tbl(i, 4) = "'" & out(i, 5)
        tbl(i, 5) = out(i, 6): tbl(i, 6) = out(i, 8): tbl(i, 7) = out(i, 9): tbl(i, 8) = out(i, 10)
        tbl(i, 9) = "'" & out(i, 11): tbl(i, 10) = out(i, 12): tbl(i, 11) = out(i, 13): tbl(i, 12) = "'" & out(i, 14)
    Next
    ws.Cells(r + 1, 2).Resize(m, 12).Value = tbl
    ws.Cells(r + 1, 2).Resize(m, 1).NumberFormat = "yyyy-mm-dd"
    For i = 1 To m
        rr = r + i
        Select Case out(i, 13)
            Case "BREACH"
                ws.Cells(rr, 2).Resize(1, 12).Interior.Color = RGB(255, 199, 206): ws.Cells(rr, 12).Font.Bold = True
            Case "KEY?"
                ws.Cells(rr, 2).Resize(1, 12).Interior.Color = RGB(255, 235, 156)
            Case "LOW"
                ws.Cells(rr, 12).Interior.Color = RGB(255, 235, 156)
            Case "OK"
                ws.Cells(rr, 12).Interior.Color = RGB(198, 239, 206)
        End Select
        If out(i, 9) < MIN_Q Then ws.Cells(rr, 8).Font.Color = RGB(192, 80, 0): ws.Cells(rr, 8).Font.Bold = True
        If out(i, 10) = "YES" Then ws.Cells(rr, 9).Font.Bold = True
    Next
    ws.Cells(r, 2).Resize(m + 1, 12).Borders.Color = RGB(217, 217, 217)

    ws.Columns("B").ColumnWidth = 36: ws.Columns("C").ColumnWidth = 26: ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 13: ws.Columns("F:I").ColumnWidth = 10: ws.Columns("J").ColumnWidth = 22
    ws.Columns("K").ColumnWidth = 10: ws.Columns("L").ColumnWidth = 9: ws.Columns("M").ColumnWidth = 110
    ws.Columns("N").ColumnWidth = 8
End Sub

'==========================================================================================
'                                    NOTIFY / LOG
'==========================================================================================
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
