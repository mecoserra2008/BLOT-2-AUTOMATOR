Attribute VB_Name = "modSelfTest"
Option Explicit
'==========================================================================================
' Self-tests of the parsing core (no Terminal, no RAW / HIST / DEALS writes)
'   RunSelfTests            button: runs everything, shows a summary, activates TESTS
'   RunSelfTestsHeadless    same without any popup / e-mail; returns TRUE if all pass (build)
'   EmbedFixtures folder    copies tests/fixtures/*.csv + tests/expected/*.json into FIXTURES
'   DumpSelfTestResults dir writes the VBA deals of every fixture case as CSV (cross-check
'                           with tests/reference_impl.py)
' Fixture spec (tests/expected/<fixture>.json):
'   { "fixture": "...csv", "default_trade_date": "yyyy-mm-dd", "dealer_map": {alias: entity},
'     "cases": [ { "name": "...", "config": {KEY: value}, "rerun": bool,
'                  "expect": { "mode", "n_trades", "n_below", "n_key_issue", "n_dropped",
'                              "pct_below", "n_date_fallback", "hist_rows", "hist_daily_rows",
'                              "deals": [ {"key","date","n","status","dealers","security"} ] } } ] }
'==========================================================================================

Private gT As Worksheet, gRow As Long, gPass As Long, gFail As Long
Private gDump As Collection          ' Array(fixture, case, result dictionary)
Private jS As String, jP As Long     ' JSON reader state

'==========================================================================================
'                                    ENTRY POINTS
'==========================================================================================
Public Sub RunSelfTests()
    Dim ok As Boolean
    ok = RunSelfTestsCore()
    If Application.Visible And Not gHeadless Then
        GetSheet(SH_TESTS).Activate
        MsgBox "Self-tests: " & gPass & " passed, " & gFail & " failed.", IIf(ok, vbInformation, vbCritical), "BLOT Best Execution"
    End If
End Sub

Public Function RunSelfTestsHeadless() As Boolean
    Dim was As Boolean
    was = gHeadless: gHeadless = True
    RunSelfTestsHeadless = RunSelfTestsCore()
    gHeadless = was
End Function

Private Function RunSelfTestsCore() As Boolean
    Dim fx As Collection, f As Variant, su As Boolean
    On Error GoTo EH
    su = Application.ScreenUpdating: Application.ScreenUpdating = False
    InitTestSheet
    Set gDump = New Collection
    LoadDefaults
    UnitTests
    Set fx = GetFixtures()
    If fx.Count = 0 Then Check "fixtures", "-", "fixtures found", ">0", "0"
    For Each f In fx
        RunFixture CStr(f(0)), CStr(f(1)), CStr(f(2))
    Next
    gT.Range("D2").Value = fx.Count
    gT.Range("C2").Value = "Fixtures:"
    LoadConfig
    FinishTestSheet
    Application.ScreenUpdating = su
    WriteLog "Self-tests: " & gPass & " passed, " & gFail & " failed.", IIf(gFail = 0, LV_INFO, LV_WARN)
    RunSelfTestsCore = (gFail = 0 And gPass > 0)
    Exit Function
EH:
    Check "suite", "-", "unexpected error", "", Err.Description
    On Error Resume Next
    LoadConfig
    FinishTestSheet
    Application.ScreenUpdating = True
    RunSelfTestsCore = False
End Function

'==========================================================================================
'                                    FIXTURES
'==========================================================================================
Public Sub EmbedFixtures(ByVal folder As String)
    Dim ws As Worksheet, f As String, r As Long, csv As String, js As String, jp As String, fv As Variant
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    Set ws = GetSheet(SH_FIX)
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("Fixture", "CSV text", "Expected JSON")
    ws.Columns("A:C").NumberFormat = "@"
    r = 1
    For Each fv In DirList(folder & "*.csv")
        f = CStr(fv)
        csv = ReadTextFile(folder & f)
        jp = folder & "..\expected\" & Left$(f, Len(f) - 4) & ".json"
        If Len(Dir$(jp)) = 0 Then Err.Raise vbObjectError + 10, , "Missing expected file for " & f & ": " & jp
        js = ReadTextFile(jp)
        If Len(csv) > 32000 Or Len(js) > 32000 Then Err.Raise vbObjectError + 11, , "Fixture too large for one cell: " & f
        r = r + 1
        ws.Cells(r, 1).Value = f
        ws.Cells(r, 2).Value = csv
        ws.Cells(r, 3).Value = js
    Next
    ws.Visible = xlSheetHidden
    WriteLog "EmbedFixtures: " & (r - 1) & " fixture(s) from " & folder, LV_INFO
End Sub

' Array(name, csvText, jsonText): FIXTURES sheet if filled, else <workbook>\tests or <workbook>\..\tests
Private Function GetFixtures() As Collection
    Dim c As New Collection, ws As Worksheet, lr As Long, i As Long, folder As Variant, f As String, fv As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_FIX)
    On Error GoTo 0
    If Not ws Is Nothing Then
        lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For i = 2 To lr
            If Len(ws.Cells(i, 1).Value) > 0 Then c.Add Array(CStr(ws.Cells(i, 1).Value), CStr(ws.Cells(i, 2).Value), CStr(ws.Cells(i, 3).Value))
        Next
    End If
    If c.Count = 0 And Len(ThisWorkbook.Path) > 0 Then
        For Each folder In Array(ThisWorkbook.Path & "\tests\", ThisWorkbook.Path & "\..\tests\")
            If Len(Dir$(folder & "fixtures\*.csv")) > 0 Then
                For Each fv In DirList(folder & "fixtures\*.csv")
                    f = CStr(fv)
                    If Len(Dir$(folder & "expected\" & Left$(f, Len(f) - 4) & ".json")) > 0 Then
                        c.Add Array(f, ReadTextFile(folder & "fixtures\" & f), ReadTextFile(folder & "expected\" & Left$(f, Len(f) - 4) & ".json"))
                    End If
                Next
                Exit For
            End If
        Next
    End If
    Set GetFixtures = c
End Function

' Dir$ has one global enumeration: collect the names before any other Dir$ call
Private Function DirList(ByVal pattern As String) As Collection
    Dim c As New Collection, f As String
    f = Dir$(pattern)
    Do While Len(f) > 0
        c.Add f
        f = Dir$()
    Loop
    Set DirList = c
End Function

Private Sub RunFixture(ByVal name As String, ByVal csv As String, ByVal js As String)
    Dim spec As Object, cs As Variant, cname As String, defTD As Date, a As Variant, blk As Variant
    Dim res As Object, dm As Object, ex As Object
    On Error GoTo EH
    Set spec = JsonParse(js)
    defTD = PrevBusinessDay(Date)
    If spec.Exists("default_trade_date") Then
        If Not ParseDateText(CStr(spec("default_trade_date")), "", defTD) Then Err.Raise vbObjectError + 12, , "bad default_trade_date"
    End If
    For Each cs In spec("cases")
        cname = CStr(cs("name"))
        ApplyConfig CfgDict(cs)
        a = ParseCsvText(csv, CSV_DELIM)
        blk = FindBlockInArray(a)
        If IsEmpty(blk) Then
            Check name, cname, "header row found", "yes", "no"
        Else
            If cs.Exists("dealer_map") Then Set dm = MapDict(cs("dealer_map")) Else Set dm = MapDict(spec("dealer_map"))
            Set res = ParseExport(blk, dm, defTD)
            Set ex = cs("expect")
            CompareResult name, cname, res, ex
            If ex.Exists("hist_rows") Or ex.Exists("hist_daily_rows") Then
                HistTest name, cname, res, ex, JBool(cs, "rerun")
            End If
            gDump.Add Array(name, cname, res)
        End If
        LoadDefaults
    Next
    Exit Sub
EH:
    Check name, cname, "fixture ran without error", "", Err.Description
    LoadDefaults
End Sub

Private Function CfgDict(cs As Variant) As Object
    Dim d As Object, k As Variant, c As Object
    Set d = NewDict()
    If cs.Exists("config") Then
        Set c = cs("config")
        For Each k In c.Keys
            If IsNull(c(k)) Then d(UCase$(k)) = "" Else d(UCase$(k)) = CStr(JText(c(k)))
        Next
    End If
    Set CfgDict = d
End Function

Private Function MapDict(m As Variant) As Object
    Dim d As Object, k As Variant
    Set d = NewDict()
    If IsObject(m) Then
        For Each k In m.Keys
            d(NormName(CStr(k))) = NormName(CStr(m(k)))
        Next
    End If
    Set MapDict = d
End Function

Private Sub CompareResult(ByVal fx As String, ByVal cname As String, res As Object, ex As Object)
    Dim deals As Variant, got As Object, i As Long, e As Variant, k As String, row As Variant, m As Long

    Check fx, cname, "parse ok", "True", B2S(res("ok")) & IIf(Len(res("err")) > 0, " (" & res("err") & ")", "")
    If ex.Exists("mode") Then Check fx, cname, "mode", CStr(ex("mode")), res("mode")
    If ex.Exists("n_trades") Then Check fx, cname, "n_trades", JInt(ex("n_trades")), CStr(res("m"))
    If ex.Exists("n_below") Then Check fx, cname, "n_below", JInt(ex("n_below")), CStr(res("nB"))
    If ex.Exists("n_key_issue") Then Check fx, cname, "n_key_issue", JInt(ex("n_key_issue")), CStr(res("nBad"))
    If ex.Exists("n_dropped") Then Check fx, cname, "n_dropped", JInt(ex("n_dropped")), CStr(res("nDropped"))
    If ex.Exists("n_date_fallback") Then Check fx, cname, "n_date_fallback", JInt(ex("n_date_fallback")), CStr(res("nDateFallback"))
    If ex.Exists("pct_below") And res("m") > 0 Then
        Check fx, cname, "pct_below", Format$(CDbl(ex("pct_below")), "0.000"), Format$(res("nB") / res("m"), "0.000")
    End If
    If Not ex.Exists("deals") Then Exit Sub

    Set got = NewDict()
    m = res("m")
    If m > 0 Then
        deals = res("deals")
        For i = 1 To m: got(CStr(deals(i, 2))) = i: Next
    End If
    Check fx, cname, "deal count", CStr(ex("deals").Count), CStr(m)
    For Each e In ex("deals")
        k = CStr(e("key"))
        If Not got.Exists(k) Then
            Check fx, cname, "deal " & k, "present", "missing (got: " & JoinKeys(got) & ")"
        Else
            i = got(k)
            If e.Exists("n") Then Check fx, cname, k & " n", JInt(e("n")), CStr(deals(i, 6))
            If e.Exists("status") Then Check fx, cname, k & " status", CStr(e("status")), CStr(deals(i, 7))
            If e.Exists("date") Then Check fx, cname, k & " date", CStr(e("date")), Format$(deals(i, 1), "yyyy-mm-dd")
            If e.Exists("dealers") Then Check fx, cname, k & " dealers", CStr(e("dealers")), CStr(deals(i, 9))
            If e.Exists("security") Then Check fx, cname, k & " security", CStr(e("security")), CStr(deals(i, 3))
        End If
    Next
End Sub

' Loads the deals into scratch HIST / HIST_DAILY sheets (once, or twice when rerun) and counts rows
Private Sub HistTest(ByVal fx As String, ByVal cname As String, res As Object, ex As Object, ByVal rerun As Boolean)
    Dim wsH As Worksheet, wsD As Worksheet, out As Variant, m As Long, i As Long, dates As Object
    Dim nH1 As Long, nH2 As Long, nD As Long, da As Boolean
    m = res("m")
    If m = 0 Then Check fx, cname, "hist test has deals", ">0", "0": Exit Sub
    out = res("deals")
    For i = 1 To m: out(i, 8) = Now: out(i, 10) = "selftest": Next
    Set wsH = ScratchSheet("T_HIST"): Set wsD = ScratchSheet("T_HIST_DAILY")
    Set dates = UpsertHist(wsH, out, m)
    RecalcHistDaily wsH, wsD, dates
    nH1 = wsH.Cells(wsH.Rows.Count, 1).End(xlUp).Row - 1
    If rerun Then
        For i = 1 To m: out(i, 8) = Now + TimeSerial(0, 1, 0): Next
        Set dates = UpsertHist(wsH, out, m)
        RecalcHistDaily wsH, wsD, dates
    End If
    nH2 = wsH.Cells(wsH.Rows.Count, 1).End(xlUp).Row - 1
    nD = wsD.Cells(wsD.Rows.Count, 1).End(xlUp).Row - 1
    If rerun Then Check fx, cname, "HIST rows unchanged on re-run", CStr(nH1), CStr(nH2)
    If ex.Exists("hist_rows") Then Check fx, cname, "HIST rows", JInt(ex("hist_rows")), CStr(nH2)
    If ex.Exists("hist_daily_rows") Then Check fx, cname, "HIST_DAILY rows", JInt(ex("hist_daily_rows")), CStr(nD)
    If nD >= 1 Then
        Check fx, cname, "HIST_DAILY NTrades total", CStr(m), CStr(Application.WorksheetFunction.Sum(wsD.Range("B:B")))
        Check fx, cname, "HIST_DAILY NBelowMin total", CStr(res("nB")), CStr(Application.WorksheetFunction.Sum(wsD.Range("C:C")))
    End If
    da = Application.DisplayAlerts: Application.DisplayAlerts = False
    wsH.Delete: wsD.Delete
    Application.DisplayAlerts = da
End Sub

Private Function ScratchSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet, da As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If Not ws Is Nothing Then
        da = Application.DisplayAlerts: Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = da
    End If
    Set ScratchSheet = GetSheet(nm)
End Function

Public Sub DumpSelfTestResults(ByVal outDir As String)
    Dim it As Variant, res As Object, deals As Variant, i As Long, h As Integer, p As String, fso As Object
    If gDump Is Nothing Then RunSelfTestsHeadless
    If gDump.Count = 0 Then RunSelfTestsHeadless
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(outDir) Then fso.CreateFolder outDir
    If Right$(outDir, 1) <> "\" Then outDir = outDir & "\"
    For Each it In gDump
        Set res = it(2)
        p = outDir & Left$(it(0), Len(it(0)) - 4) & "__" & it(1) & ".csv"
        h = FreeFile
        Open p For Output As #h
        Print #h, "key,date,n,status,dealers,security,mode"
        If res("m") > 0 Then
            deals = res("deals")
            For i = 1 To res("m")
                Print #h, Q(deals(i, 2)) & "," & Q(Format$(deals(i, 1), "yyyy-mm-dd")) & "," & Q(deals(i, 6)) & "," & _
                          Q(deals(i, 7)) & "," & Q(deals(i, 9)) & "," & Q(deals(i, 3)) & "," & Q(res("mode"))
            Next
        End If
        Close #h
    Next
End Sub

Private Function Q(ByVal x As Variant) As String
    Q = """" & Replace(CStr(x), """", """""") & """"
End Function

'==========================================================================================
'                                    UNIT TESTS
'==========================================================================================
Private Sub UnitTests()
    Dim d As Date, a As Variant, savedH As Object, s As String

    '--- calendar ---
    Check "unit", "calendar", "Easter 2026", "2026-04-05", Format$(EasterSunday(2026), "yyyy-mm-dd")
    Check "unit", "calendar", "Easter 2027", "2027-03-28", Format$(EasterSunday(2027), "yyyy-mm-dd")
    Set savedH = HolidaySet
    Set HolidaySet = NewDict()
    HolidaySet(CLng(DateSerial(2026, 6, 10))) = 1
    Check "unit", "calendar", "prev bday of Thu 2026-06-11 (Wed holiday)", "2026-06-09", Format$(PrevBusinessDay(DateSerial(2026, 6, 11)), "yyyy-mm-dd")
    Check "unit", "calendar", "prev bday of Mon 2026-11-16", "2026-11-13", Format$(PrevBusinessDay(DateSerial(2026, 11, 16)), "yyyy-mm-dd")
    Check "unit", "calendar", "DAYS_BACK Mon 2026-11-16", "3", CStr(DaysBack(DateSerial(2026, 11, 16)))
    Check "unit", "calendar", "holiday is not a business day", "False", B2S(IsBusinessDay(DateSerial(2026, 6, 10)))
    Check "unit", "steps", "{DAYS_BACK} expansion", "KEYS:BLOT -2<GO>", ExpandSeq("KEYS:BLOT -{DAYS_BACK}<GO>", DateSerial(2026, 6, 11))
    Set HolidaySet = savedH

    '--- dates ---
    Check "unit", "dates", "MM/DD/YYYY 11/10/2026", "2026-11-10", DT("11/10/2026", "MM/DD/YYYY")
    Check "unit", "dates", "DD/MM/YYYY 10/11/2026 14:32", "2026-11-10", DT("10/11/2026 14:32", "DD/MM/YYYY")
    Check "unit", "dates", "DD-MMM-YY 10-Nov-26", "2026-11-10", DT("10-Nov-26", "DD-MMM-YY")
    Check "unit", "dates", "DD-MMM-YYYY 10-Dez-2026 (PT)", "2026-12-10", DT("10-Dez-2026", "DD-MMM-YYYY")
    Check "unit", "dates", "YYYYMMDD 20261110", "2026-11-10", DT("20261110", "YYYYMMDD")
    Check "unit", "dates", "M/D/YYYY 1/5/2026", "2026-01-05", DT("1/5/2026", "M/D/YYYY")
    Check "unit", "dates", "ISO without format", "2026-11-10", DT("2026-11-10", "")
    Check "unit", "dates", "ISO with time", "2026-11-10", DT("2026-11-10T09:15:00", "MM/DD/YYYY")
    Check "unit", "dates", "ambiguous text without format rejected", "FAIL", DT("11/10/2026", "")
    Check "unit", "dates", "invalid month rejected", "FAIL", DT("13/10/2026", "MM/DD/YYYY")
    Check "unit", "dates", "invalid day rejected", "FAIL", DT("02/30/2026", "MM/DD/YYYY")
    If ParseDateCell(CDbl(DateSerial(2026, 11, 10)) + 0.4, "", d) Then s = Format$(d, "yyyy-mm-dd") Else s = "FAIL"
    Check "unit", "dates", "numeric serial (Value2)", "2026-11-10", s
    Check "unit", "dates", "format validation", "", ValidateDateFormat("MM/DD/YYYY")
    Check "unit", "dates", "bad format detected", "True", B2S(Len(ValidateDateFormat("MM/YYYY")) > 0)

    '--- numbers / text ---
    Check "unit", "text", "InvariantNum 1234567890123", "1234567890123", InvariantNum(1234567890123#)
    Check "unit", "text", "InvariantNum 2.5", "2.5", InvariantNum(2.5)
    Check "unit", "text", "InvariantNum 0.25", "0.25", InvariantNum(0.25)
    Check "unit", "text", "NormName", "JPM LDN", NormName("  jpm" & vbTab & "  ldn" & Chr$(160))
    Check "unit", "text", "IsPlainNumber 1,000", "False", B2S(IsPlainNumber("1,000"))

    '--- csv ---
    a = ParseCsvText("a;b;c" & vbCrLf & """x;y"";2;""l1" & vbLf & "l2""" & vbCrLf & "q;""he said """"hi"""""";", "AUTO")
    Check "unit", "csv", "rows x cols", "3x3", (UBound(a, 1)) & "x" & (UBound(a, 2))
    Check "unit", "csv", "quoted delimiter", "x;y", CStr(a(2, 1))
    Check "unit", "csv", "line break in quotes", "l1|l2", Replace(CStr(a(2, 3)), vbLf, "|")
    Check "unit", "csv", "escaped quote", "he said ""hi""", CStr(a(3, 2))
    Check "unit", "csv", "empty last field", "", CStr(a(3, 3))
    Check "unit", "csv", "comma detected", ",", CsvDelimiter("a,b,c" & vbLf & "1,2,3", "AUTO")
    Check "unit", "csv", "tab detected", "TAB", Replace(CsvDelimiter("a" & vbTab & "b" & vbTab & "c", "AUTO"), vbTab, "TAB")

    '--- steps ---
    Check "unit", "steps", "valid sequence", "", ValidateSteps("KEYS:BLOT 2<GO>|WAIT:1.5|CLICK:10,20|RCLICK:5,5|SEND:{TAB}~")
    Check "unit", "steps", "unknown token rejected", "True", B2S(Len(ValidateSteps("FOO:1")) > 0)
    Check "unit", "steps", "bad click rejected", "True", B2S(Len(ValidateSteps("CLICK:10")) > 0)
    Check "unit", "steps", "bad wait rejected", "True", B2S(Len(ValidateSteps("WAIT:x")) > 0)
    Check "unit", "steps", "placeholder KEYS:N<GO> rejected", "True", B2S(Len(ValidateSteps("KEYS:96<GO>|WAIT:1|KEYS:N<GO>")) > 0)
    Check "unit", "steps", "numbered menu accepted", "", ValidateSteps("KEYS:96<GO>|WAIT:1|KEYS:3<GO>|WAIT:3")
    Check "unit", "record", "RecWait 0.2s", "0.5", InvariantNum(RecWait(0.2))
    Check "unit", "record", "RecWait 1.2s", "1.5", InvariantNum(RecWait(1.2))
    Check "unit", "record", "RecWait capped", "8", InvariantNum(RecWait(30))
    Check "unit", "record", "SuggestPrefix", "1-BLOOMBERG", SuggestPrefix("1-BLOOMBERG - BLOT")
    Check "unit", "calib", "default candidates", "3", CStr(SplitCandidates(DefaultOf("EXPORT_CANDIDATES")).Count)
    Check "unit", "calib", "candidate 1", "KEYS:2<GO>|WAIT:2|KEYS:2<GO>|WAIT:5", SplitCandidates(DefaultOf("EXPORT_CANDIDATES"))(1)
    Check "unit", "calib", "candidates are valid steps", "", ValidateSteps(Replace(DefaultOf("EXPORT_CANDIDATES"), "||", "|"))
    Check "unit", "calib", "Bloomberg exe wintrv.exe", "True", B2S(IsBloombergExe("wintrv.exe"))
    Check "unit", "calib", "excel.exe is not Bloomberg", "False", B2S(IsBloombergExe("excel.exe"))

    '--- rules guards ---
    s = EXCL_STATUS
    EXCL_STATUS = "cancel||"
    Check "unit", "rules", "empty EXCL_STATUS token ignored", "1", CStr(LiveCount("Done"))
    EXCL_STATUS = s
    Check "unit", "rules", "empty alias never matches empty header", "0", CStr(HeaderIndex(HdrRow(Array("", "x", "y")), "|dealer"))
    Check "unit", "rules", "header contains-match (>= 5 chars)", "2", CStr(HeaderIndex(HdrRow(Array("id", "No. of Dealers asked", "y")), "no. of dealers"))
End Sub

Private Function B2S(ByVal b As Boolean) As String
    If b Then B2S = "True" Else B2S = "False"
End Function

Private Function DT(ByVal s As String, ByVal fmt As String) As String
    Dim d As Date
    If ParseDateText(s, fmt, d) Then DT = Format$(d, "yyyy-mm-dd") Else DT = "FAIL"
End Function

Private Function HdrRow(a As Variant) As Variant
    Dim h() As Variant, j As Long
    ReDim h(1 To 1, 1 To UBound(a) - LBound(a) + 1)
    For j = LBound(a) To UBound(a): h(1, j - LBound(a) + 1) = a(j): Next
    HdrRow = h
End Function

' Number of live deals of a one-deal, 3-dealer export with the given status
Private Function LiveCount(ByVal status As String) As Long
    Dim v(1 To 4, 1 To 3) As Variant, res As Object, i As Long
    v(1, 1) = "inquiry id": v(1, 2) = "dealer": v(1, 3) = "status"
    For i = 2 To 4: v(i, 1) = "R1": v(i, 2) = "D" & i: v(i, 3) = status: Next
    Set res = ParseExport(v, NewDict(), Date)
    LiveCount = res("m")
End Function

'==========================================================================================
'                                    TESTS SHEET
'==========================================================================================
Private Sub InitTestSheet()
    Set gT = GetSheet(SH_TESTS)
    gT.Cells.Clear
    gT.Range("A1").Value = "Self-tests": gT.Range("A1").Font.Bold = True: gT.Range("A1").Font.Size = 13
    gT.Range("A2").Value = "Result"
    gT.Range("A4:F4").Value = Array("Suite", "Case", "Check", "Expected", "Actual", "Result")
    gT.Range("A4:F4").Font.Bold = True: gT.Range("A4:F4").Interior.Color = RGB(217, 225, 242)
    gT.Columns("A:E").NumberFormat = "@"
    gRow = 4: gPass = 0: gFail = 0
End Sub

Private Sub Check(ByVal suite As String, ByVal cname As String, ByVal what As String, ByVal expected As String, ByVal actual As String)
    Dim ok As Boolean
    If gT Is Nothing Then InitTestSheet
    ok = (expected = actual)
    gRow = gRow + 1
    gT.Cells(gRow, 1).Resize(1, 5).Value = Array(suite, cname, what, expected, actual)
    gT.Cells(gRow, 6).Value = IIf(ok, "PASS", "FAIL")
    gT.Cells(gRow, 6).Interior.Color = IIf(ok, RGB(198, 239, 206), RGB(255, 199, 206))
    If ok Then gPass = gPass + 1 Else gFail = gFail + 1
End Sub

Private Sub FinishTestSheet()
    If gT Is Nothing Then Exit Sub
    gT.Range("B2").Value = IIf(gFail = 0 And gPass > 0, "PASS", "FAIL") & " - " & gPass & " passed, " & gFail & " failed - " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    gT.Range("B2").Font.Bold = True
    gT.Range("B2").Interior.Color = IIf(gFail = 0 And gPass > 0, RGB(198, 239, 206), RGB(255, 199, 206))
    gT.Columns("A:B").ColumnWidth = 22: gT.Columns("C").ColumnWidth = 44
    gT.Columns("D:E").ColumnWidth = 34: gT.Columns("F").ColumnWidth = 8
End Sub

Private Function JoinKeys(d As Object) As String
    Dim k As Variant
    For Each k In d.Keys: JoinKeys = JoinKeys & k & " ": Next
End Function

'==========================================================================================
'                                    MINIMAL JSON READER
'   objects -> Scripting.Dictionary, arrays -> Collection, numbers -> Double,
'   strings -> String, true/false -> Boolean, null -> Null
'==========================================================================================
Public Function JsonParse(ByVal s As String) As Object
    jS = s: jP = 1
    SkipWs
    If Mid$(jS, jP, 1) <> "{" Then Err.Raise vbObjectError + 20, , "JSON: object expected at top level"
    Set JsonParse = JObject()
End Function

Private Function JObject() As Object
    Dim d As Object, k As String, c As String
    Set d = NewDict()
    jP = jP + 1: SkipWs
    If Mid$(jS, jP, 1) = "}" Then jP = jP + 1: Set JObject = d: Exit Function
    Do
        SkipWs
        If Mid$(jS, jP, 1) <> """" Then Err.Raise vbObjectError + 21, , "JSON: key expected at " & jP
        k = JString()
        SkipWs
        If Mid$(jS, jP, 1) <> ":" Then Err.Raise vbObjectError + 22, , "JSON: ':' expected at " & jP
        jP = jP + 1: SkipWs
        If NextIsContainer() Then Set d(k) = JContainer() Else d(k) = JScalar()
        SkipWs
        c = Mid$(jS, jP, 1): jP = jP + 1
        If c = "}" Then Exit Do
        If c <> "," Then Err.Raise vbObjectError + 23, , "JSON: ',' or '}' expected at " & (jP - 1)
    Loop
    Set JObject = d
End Function

Private Function JArray() As Collection
    Dim col As New Collection, c As String
    jP = jP + 1: SkipWs
    If Mid$(jS, jP, 1) = "]" Then jP = jP + 1: Set JArray = col: Exit Function
    Do
        SkipWs
        If NextIsContainer() Then col.Add JContainer() Else col.Add JScalar()
        SkipWs
        c = Mid$(jS, jP, 1): jP = jP + 1
        If c = "]" Then Exit Do
        If c <> "," Then Err.Raise vbObjectError + 24, , "JSON: ',' or ']' expected at " & (jP - 1)
    Loop
    Set JArray = col
End Function

Private Function NextIsContainer() As Boolean
    Dim c As String
    c = Mid$(jS, jP, 1)
    NextIsContainer = (c = "{" Or c = "[")
End Function

Private Function JContainer() As Object
    If Mid$(jS, jP, 1) = "{" Then Set JContainer = JObject() Else Set JContainer = JArray()
End Function

Private Function JScalar() As Variant
    Dim c As String, st As Long
    c = Mid$(jS, jP, 1)
    Select Case c
        Case """"
            JScalar = JString()
        Case "t"
            jP = jP + 4: JScalar = True
        Case "f"
            jP = jP + 5: JScalar = False
        Case "n"
            jP = jP + 4: JScalar = Null
        Case Else
            st = jP
            Do While jP <= Len(jS) And InStr(1, "+-.eE0123456789", Mid$(jS, jP, 1)) > 0
                jP = jP + 1
            Loop
            If jP = st Then Err.Raise vbObjectError + 25, , "JSON: unexpected '" & c & "' at " & jP
            JScalar = Val(Mid$(jS, st, jP - st))
    End Select
End Function

Private Function JString() As String
    Dim o As String, c As String
    jP = jP + 1
    Do While jP <= Len(jS)
        c = Mid$(jS, jP, 1)
        If c = """" Then
            jP = jP + 1: JString = o: Exit Function
        ElseIf c = "\" Then
            jP = jP + 1
            c = Mid$(jS, jP, 1)
            Select Case c
                Case "n": o = o & vbLf
                Case "r": o = o & vbCr
                Case "t": o = o & vbTab
                Case "b": o = o & Chr$(8)
                Case "f": o = o & Chr$(12)
                Case "u": o = o & ChrW(CLng("&H" & Mid$(jS, jP + 1, 4))): jP = jP + 4
                Case Else: o = o & c
            End Select
        Else
            o = o & c
        End If
        jP = jP + 1
    Loop
    Err.Raise vbObjectError + 26, , "JSON: unterminated string"
End Function

Private Sub SkipWs()
    Do While jP <= Len(jS)
        Select Case Mid$(jS, jP, 1)
            Case " ", vbTab, vbCr, vbLf
                jP = jP + 1
            Case Else
                Exit Sub
        End Select
    Loop
End Sub

Private Function JInt(ByVal x As Variant) As String
    JInt = CStr(CLng(x))
End Function

Private Function JText(ByVal x As Variant) As String
    Select Case VarType(x)
        Case vbBoolean: JText = IIf(x, "TRUE", "FALSE")
        Case vbDouble: JText = InvariantNum(x)
        Case Else: JText = CStr(x)
    End Select
End Function

Private Function JBool(d As Variant, ByVal k As String) As Boolean
    If d.Exists(k) Then
        If VarType(d(k)) = vbBoolean Then JBool = d(k)
    End If
End Function
