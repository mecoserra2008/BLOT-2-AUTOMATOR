Attribute VB_Name = "modConfig"
Option Explicit
'==========================================================================================
' BLOT Best-Execution monitor v5 - CONFIG sheet, validation, business-day calendar
'   LoadConfig / LoadDefaults / ApplyConfig d   fill the public variables below
'   ValidateConfig                              PASS / WARN / FAIL list on CONFIG (E:G)
'==========================================================================================

'------------------------------------ SHEETS ---------------------------------------------
Public Const SH_DASH   As String = "DASHBOARD"
Public Const SH_INQ    As String = "INQUIRIES"
Public Const SH_QUOTES As String = "QUOTES"
Public Const SH_HIST   As String = "HIST"
Public Const SH_HD     As String = "HIST_DAILY"
Public Const SH_PASTE  As String = "PASTE"
Public Const SH_CONFIG As String = "CONFIG"
Public Const SH_MAP    As String = "DEALER_MAP"
Public Const SH_LOG    As String = "LOG"
Public Const SH_TESTS  As String = "TESTS"
Public Const SH_FIX    As String = "FIXTURES"

' CONFIG layout: row 1 buttons, row 2 title, row 3 hint, row 4 header, data from row 5
Private Const CFG_HDR_ROW  As Long = 4
Private Const CFG_COL_KEY  As Long = 1
Private Const CFG_COL_VAL  As Long = 2
Private Const CFG_COL_DESC As Long = 3
Private Const CFG_COL_CHK  As Long = 5      ' E:G validation report
Private Const CFG_COL_HOL  As Long = 9      ' I:J holiday table
Private Const HOLIDAY_REF  As String = "I5:I400"

'------------------------------------ CONFIG VALUES --------------------------------------
Public BBG_PANEL As String, BBG_TITLE_PREFIX As String, OPEN_SEQ As String, EXPORT_SEQ As String
Public EXPORT_FOLDER As String, TIMEOUT_S As Long, POLL_S As Long
Public RUN_AT As String, ALERT_TO As String, SHOW_POPUPS As Boolean, SCHEDULE_MODE As String, RETRY_MIN As Long
Public MIN_Q As Long, GROUP_LEVEL As String, FILL_STATUS As String, ID_SUFFIX_LEN As Long, DATE_FORMAT_TEXT As String
Public HDR_DEALER As String, HDR_ID As String, HDR_SECURITY As String, HDR_ISIN As String, HDR_SIDE As String
Public HDR_STATUS As String, HDR_QTY As String, HDR_PRICE As String, HDR_DATE As String

' fixed (no longer on CONFIG)
Public Const STEP_GAP_S As Double = 0.7
Public Const TASK_NAME As String = "BLOT Best Execution"
Public Const TEST_EXPORT_RUNS As Long = 5

Public HolidaySet As Object          ' Long(date serial) -> 1
Public gCfgIssues As String

Private gDefs As Collection          ' Array(section, key, default, description, flag)
Private gDefIdx As Object

'==========================================================================================
'                                    DEFAULTS
'==========================================================================================
Private Sub BuildDefs()
    If Not gDefs Is Nothing Then Exit Sub
    Set gDefs = New Collection
    Set gDefIdx = NewDict()
    AddDef "BLOOMBERG", "BBG_PANEL", "<blp-1>", "Bloomberg panel used for the DDE keystrokes.", ""
    AddDef "BLOOMBERG", "BBG_TITLE_PREFIX", "1-BLOOMBERG", "Start of the panel window title (only needed for SEND: steps).", ""
    AddDef "BLOOMBERG", "OPEN_SEQ", "KEYS:BLOT 2<GO>|WAIT:3|KEYS:BLOT -1<GO>|WAIT:5", "Opens the BLOT of the previous day. KEYS:text<GO> | SEND:keys | WAIT:s. {DAYS_BACK} = days to the previous business day.", ""
    AddDef "BLOOMBERG", "EXPORT_SEQ", "KEYS:2<GO>|WAIT:2|KEYS:2<GO>|WAIT:3|KEYS:1<GO>|WAIT:5", _
           "Export through DDE: 2<GO> Export, 2<GO> 'XLS - Excel format', 1<GO> '1) Export' in the Export Blotter dialog.", ""
    AddDef "BLOOMBERG", "EXPORT_FOLDER", "", "Folder where every BLOT export is SAVED by the workbook (BLOT_<tradedate>_<timestamp>.xls*, regulatory evidence). Also watched for Bloomberg's own XLS.", "U"
    AddDef "BLOOMBERG", "TIMEOUT_S", "300", "Seconds to wait for the export before a TIMEOUT alert.", ""
    AddDef "BLOOMBERG", "POLL_S", "3", "Seconds between checks for the export.", ""

    AddDef "RUN", "RUN_AT", "08:30:00", "Daily run time (weekdays; the holidays in column I are skipped).", ""
    AddDef "RUN", "SCHEDULE_MODE", "WORKBOOK", "TASK = Windows Task Scheduler opens Excel and runs at RUN_AT (button 'Install Daily Task') | WORKBOOK = runs at RUN_AT while this workbook is open | OFF.", "U"
    AddDef "RUN", "RETRY_MIN", "10", "Scheduled run only: minutes before ONE retry when Bloomberg did not export (0 = no retry).", ""
    AddDef "RUN", "ALERT_TO", "", "E-mail recipients (';'-separated) for alerts, timeouts and errors. Empty = no e-mail.", "U"
    AddDef "RUN", "SHOW_POPUPS", "FALSE", "TRUE = popup on alerts (blocks Excel until dismissed: keep FALSE for unattended runs).", ""

    AddDef "RULE", "MIN_Q", "3", "Minimum counterparties ASKED per quote request (Banco de Portugal: 3).", ""
    AddDef "RULE", "GROUP_LEVEL", "INQUIRY", "INQUIRY = each quote request (ID without its last characters) | ORDER = all requests for the same ISIN, side and trade date counted together.", "U"
    AddDef "RULE", "FILL_STATUS", "done|filled|accepted|traded|executed|complete", "Status values ('|'-separated, substring) meaning the dealer row was EXECUTED. Check the 'Statuses seen' line on the DASHBOARD.", "U"
    AddDef "RULE", "ID_SUFFIX_LEN", "2", "ID = request + dealer leg: the last N characters are the leg (FITDNA:6AF5A884011800|03).", ""
    AddDef "RULE", "DATE_FORMAT_TEXT", "DD/MM/YYYY", "Format of TEXT dates (real date cells are always read correctly).", ""

    AddDef "HEADERS", "HDR_DEALER", "brkrname|broker name|broker|dealer|counterparty", "Dealer column (exact export header first, lowercase, '|'-separated).", ""
    AddDef "HEADERS", "HDR_ID", "id|ticket id|trade id", "Row ID (request + leg).", ""
    AddDef "HEADERS", "HDR_SECURITY", "security|description", "Security description.", ""
    AddDef "HEADERS", "HDR_ISIN", "isin", "ISIN.", ""
    AddDef "HEADERS", "HDR_SIDE", "side|b/s", "Side.", ""
    AddDef "HEADERS", "HDR_STATUS", "status", "Status of the dealer row.", ""
    AddDef "HEADERS", "HDR_QTY", "qty (m)|qty|quantity", "Quantity.", ""
    AddDef "HEADERS", "HDR_PRICE", "price", "Price quoted by the dealer (non-empty = the dealer quoted).", ""
    AddDef "HEADERS", "HDR_DATE", "trade dt|trade date", "Trade date.", ""
End Sub

Private Sub AddDef(ByVal sec As String, ByVal k As String, ByVal def As String, ByVal desc As String, ByVal flag As String)
    gDefs.Add Array(sec, k, def, desc, flag)
    gDefIdx(k) = gDefs.Count
End Sub

Public Function DefaultOf(ByVal k As String) As String
    BuildDefs
    If gDefIdx.Exists(k) Then DefaultOf = gDefs(gDefIdx(k))(2)
End Function

Public Function ConfigKeys() As Collection
    Dim c As New Collection, d As Variant
    BuildDefs
    For Each d In gDefs: c.Add d(1): Next
    Set ConfigKeys = c
End Function

'==========================================================================================
'                                    LOAD
'==========================================================================================
Public Sub LoadConfig()
    ApplyConfig ReadConfigSheet()
End Sub

Public Sub LoadDefaults()
    ApplyConfig NewDict()
End Sub

Public Sub ApplyConfig(ByVal d As Object)
    gCfgIssues = ""
    BBG_PANEL = CfgStr(d, "BBG_PANEL")
    BBG_TITLE_PREFIX = CfgStr(d, "BBG_TITLE_PREFIX")
    OPEN_SEQ = CfgStr(d, "OPEN_SEQ")
    EXPORT_SEQ = CfgStr(d, "EXPORT_SEQ")
    EXPORT_FOLDER = StripSlash(CfgStr(d, "EXPORT_FOLDER"))
    TIMEOUT_S = CLng(CfgNum(d, "TIMEOUT_S"))
    POLL_S = CLng(CfgNum(d, "POLL_S"))
    RUN_AT = CfgStr(d, "RUN_AT")
    SCHEDULE_MODE = UCase$(CfgStr(d, "SCHEDULE_MODE"))
    RETRY_MIN = CLng(CfgNum(d, "RETRY_MIN"))
    ALERT_TO = CfgStr(d, "ALERT_TO")
    SHOW_POPUPS = CfgBool(d, "SHOW_POPUPS")
    MIN_Q = CLng(CfgNum(d, "MIN_Q"))
    GROUP_LEVEL = UCase$(CfgStr(d, "GROUP_LEVEL"))
    FILL_STATUS = LCase$(CfgStr(d, "FILL_STATUS"))
    ID_SUFFIX_LEN = CLng(CfgNum(d, "ID_SUFFIX_LEN"))
    DATE_FORMAT_TEXT = UCase$(CfgStr(d, "DATE_FORMAT_TEXT"))
    HDR_DEALER = LCase$(CfgStr(d, "HDR_DEALER"))
    HDR_ID = LCase$(CfgStr(d, "HDR_ID"))
    HDR_SECURITY = LCase$(CfgStr(d, "HDR_SECURITY"))
    HDR_ISIN = LCase$(CfgStr(d, "HDR_ISIN"))
    HDR_SIDE = LCase$(CfgStr(d, "HDR_SIDE"))
    HDR_STATUS = LCase$(CfgStr(d, "HDR_STATUS"))
    HDR_QTY = LCase$(CfgStr(d, "HDR_QTY"))
    HDR_PRICE = LCase$(CfgStr(d, "HDR_PRICE"))
    HDR_DATE = LCase$(CfgStr(d, "HDR_DATE"))
    Set HolidaySet = LoadHolidays()
End Sub

Public Function ReadConfigSheet() As Object
    Dim d As Object, ws As Worksheet, v As Variant, lr As Long, i As Long, k As String
    Set d = NewDict()
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_CONFIG)
    On Error GoTo 0
    If ws Is Nothing Then Set ReadConfigSheet = d: Exit Function
    lr = ws.Cells(ws.Rows.Count, CFG_COL_KEY).End(xlUp).Row
    If lr > CFG_HDR_ROW Then
        v = ws.Range(ws.Cells(CFG_HDR_ROW + 1, CFG_COL_KEY), ws.Cells(lr, CFG_COL_VAL)).Value
        For i = 1 To UBound(v, 1)
            If Not IsError(v(i, 1)) Then
                k = UCase$(Trim$(CStr(v(i, 1))))
                If IsConfigKeyName(k) Then d(k) = v(i, 2)
            End If
        Next
    End If
    Set ReadConfigSheet = d
End Function

Private Function IsConfigKeyName(ByVal k As String) As Boolean
    Dim i As Long
    If Len(k) = 0 Then Exit Function
    For i = 1 To Len(k)
        If Not (Mid$(k, i, 1) Like "[A-Z0-9_]") Then Exit Function
    Next
    IsConfigKeyName = True
End Function

Public Function Cfg(ByVal key As String, Optional ByVal def As Variant = "") As Variant
    Dim d As Object
    Set d = ReadConfigSheet()
    If d.Exists(UCase$(key)) Then Cfg = d(UCase$(key)) Else Cfg = def
End Function

Private Function CfgStr(d As Object, ByVal k As String) As String
    If d.Exists(k) Then CfgStr = ValToStr(d(k), k) Else CfgStr = DefaultOf(k)
End Function

Private Function ValToStr(ByVal x As Variant, ByVal k As String) As String
    Select Case VarType(x)
        Case vbError, vbEmpty, vbNull
            ValToStr = ""
        Case vbBoolean
            ValToStr = IIf(x, "TRUE", "FALSE")
        Case vbDate
            If CDbl(x) < 1 Then ValToStr = Format$(x, "hh:nn:ss") Else ValToStr = Format$(x, "yyyy-mm-dd hh:nn:ss")
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte
            If k = "RUN_AT" And CDbl(x) >= 0 And CDbl(x) < 1 Then ValToStr = Format$(CDate(x), "hh:nn:ss") Else ValToStr = InvariantNum(CDbl(x))
        Case Else
            ValToStr = Trim$(CStr(x))
    End Select
End Function

Private Function CfgNum(d As Object, ByVal k As String) As Double
    Dim s As String
    s = Replace(CfgStr(d, k), ",", ".")
    If IsPlainNumber(s) Then
        CfgNum = Val(s)
    Else
        gCfgIssues = gCfgIssues & k & "='" & s & "' (not a number); "
        CfgNum = Val(DefaultOf(k))
    End If
End Function

Private Function CfgBool(d As Object, ByVal k As String) As Boolean
    Select Case UCase$(CfgStr(d, k))
        Case "TRUE", "VERDADEIRO", "1", "YES", "SIM"
            CfgBool = True
        Case "FALSE", "FALSO", "0", "NO", "NAO"
            CfgBool = False
        Case Else
            gCfgIssues = gCfgIssues & k & "='" & CfgStr(d, k) & "' (not TRUE/FALSE); "
            CfgBool = (UCase$(DefaultOf(k)) = "TRUE")
    End Select
End Function

Private Function StripSlash(ByVal p As String) As String
    Do While Len(p) > 3 And (Right$(p, 1) = "\" Or Right$(p, 1) = "/")
        p = Left$(p, Len(p) - 1)
    Loop
    StripSlash = p
End Function

'==========================================================================================
'                                    CONFIG SHEET
'==========================================================================================
' Rebuilds the CONFIG layout; existing values of known keys are kept, unknown (old) keys dropped
Public Sub WriteConfigSheet()
    Dim ws As Worksheet, old As Object, d As Variant, r As Long, lastSec As String, hol As Variant, nHol As Long
    BuildDefs
    Set ws = GetSheet(SH_CONFIG)
    Set old = ReadConfigSheet()
    hol = ws.Range(ws.Cells(CFG_HDR_ROW + 1, CFG_COL_HOL), ws.Cells(400, CFG_COL_HOL + 1)).Value
    nHol = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(CFG_HDR_ROW + 1, CFG_COL_HOL), ws.Cells(400, CFG_COL_HOL)))
    ws.Cells.Clear
    ws.Rows(1).RowHeight = 34
    ws.Range("A2").Value = "CONFIG - edit column B only (yellow = to fill), then Validate Config."
    ws.Range("A2").Font.Bold = True: ws.Range("A2").Font.Size = 13
    ws.Range("A3").Value = "Values are text: TRUE/FALSE, numbers with '.' or ','."
    ws.Range("A3").Font.Italic = True
    ws.Cells(CFG_HDR_ROW, 1).Resize(1, 3).Value = Array("Key", "Value", "Description")
    Hdr ws.Cells(CFG_HDR_ROW, 1).Resize(1, 3)
    r = CFG_HDR_ROW + 1
    For Each d In gDefs
        If d(0) <> lastSec Then
            If lastSec <> "" Then r = r + 1
            ws.Cells(r, 1).Value = "[" & d(0) & "]": ws.Cells(r, 1).Font.Bold = True
            r = r + 1: lastSec = d(0)
        End If
        ws.Cells(r, CFG_COL_KEY).Value = d(1)
        ws.Cells(r, CFG_COL_VAL).NumberFormat = "@"
        If old.Exists(d(1)) Then ws.Cells(r, CFG_COL_VAL).Value = ValToStr(old(d(1)), CStr(d(1))) Else ws.Cells(r, CFG_COL_VAL).Value = d(2)
        ws.Cells(r, CFG_COL_DESC).Value = d(3)
        If d(4) = "U" Then ws.Cells(r, CFG_COL_VAL).Interior.Color = RGB(255, 242, 204)
        r = r + 1
    Next
    ws.Columns(1).ColumnWidth = 18: ws.Columns(2).ColumnWidth = 60: ws.Columns(3).ColumnWidth = 95
    ws.Columns(4).ColumnWidth = 2: ws.Columns(5).ColumnWidth = 22: ws.Columns(6).ColumnWidth = 8
    ws.Columns(7).ColumnWidth = 70: ws.Columns(8).ColumnWidth = 2
    ws.Columns(9).ColumnWidth = 12: ws.Columns(10).ColumnWidth = 36
    WriteHolidayTable ws, hol, nHol
    AddSheetButtons ws, Array("Validate Config|ValidateConfig", "Test Export|TestExport", "Install Daily Task|InstallDailyTask", _
                              "Remove Daily Task|RemoveDailyTask", "Silence Format Prompt|SilenceFormatPrompt")
End Sub

Public Sub Hdr(rg As Range)
    rg.Font.Bold = True
    rg.Interior.Color = RGB(31, 56, 100)
    rg.Font.Color = vbWhite
End Sub

' Excel's "The file format and extension of ... don't match" warning (Microsoft setting ExtensionHardening:
' 0 = no warning, 1 = warn (default), 2 = block). Only needed when Bloomberg itself opens the export in Excel.
Private Function ExtHardeningKey() As String
    ExtHardeningKey = "HKCU\Software\Microsoft\Office\" & Application.Version & "\Excel\Security\ExtensionHardening"
End Function

Public Function ExtensionHardening() As Long
    On Error GoTo NotSet
    ExtensionHardening = CLng(CreateObject("WScript.Shell").RegRead(ExtHardeningKey()))
    Exit Function
NotSet:
    ExtensionHardening = 1
End Function

Public Sub SilenceFormatPrompt()
    If MsgBox("Switch off Excel's warning 'The file format and extension of ... don't match' for your Windows user?" & vbLf & vbLf & _
              "Microsoft setting ExtensionHardening = 0. Files still open normally; Protected View and macro security are not changed." & vbLf & _
              "Needed only when Bloomberg itself opens the export in Excel (the workbook reads the export as data without any prompt)." & vbLf & vbLf & _
              "Close and reopen Excel afterwards.", vbYesNo + vbQuestion, "Silence Format Prompt") <> vbYes Then Exit Sub
    On Error GoTo EH
    CreateObject("WScript.Shell").RegWrite ExtHardeningKey(), 0, "REG_DWORD"
    WriteLog "Excel format/extension warning switched off for this user (" & ExtHardeningKey() & " = 0).", "INFO"
    MsgBox "Done. Close and reopen Excel for it to take effect. (To undo: run RestoreFormatPrompt.)", vbInformation, "Silence Format Prompt"
    Exit Sub
EH:
    MsgBox "Could not change the setting (blocked by policy?): " & Err.Description, vbExclamation, "Silence Format Prompt"
End Sub

Public Sub RestoreFormatPrompt()
    On Error Resume Next
    CreateObject("WScript.Shell").RegDelete ExtHardeningKey()
    WriteLog "Excel format/extension warning restored (default).", "INFO"
    If Application.Visible Then MsgBox "Excel's default warning is restored. Close and reopen Excel.", vbInformation
End Sub

Public Sub SetConfigValue(ByVal key As String, ByVal v As String)
    Dim ws As Worksheet, f As Range
    Set ws = GetSheet(SH_CONFIG)
    Set f = ws.Columns(CFG_COL_KEY).Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If f Is Nothing Then Err.Raise vbObjectError + 2, , "Key " & key & " not found on CONFIG (run Setup)."
    ws.Cells(f.Row, CFG_COL_VAL).NumberFormat = "@"
    ws.Cells(f.Row, CFG_COL_VAL).Value = v
End Sub

' Site file (KEY=value per line, '#' = comment). Used by the build. Unknown keys are ignored.
Public Function ApplySiteConfig(ByVal path As String) As Long
    Dim txt As String, ln As Variant, p As Long, k As String, mine As Object, kk As Variant, n As Long
    Set mine = NewDict()
    For Each kk In ConfigKeys(): mine(kk) = 1: Next
    txt = Replace(Replace(ReadTextFile(path), vbCrLf, vbLf), vbCr, vbLf)
    For Each ln In Split(txt, vbLf)
        ln = Trim$(CStr(ln))
        If Len(ln) > 0 And Left$(ln, 1) <> "#" Then
            p = InStr(1, ln, "=")
            If p > 1 Then
                k = UCase$(Trim$(Left$(ln, p - 1)))
                If mine.Exists(k) Then SetConfigValue k, Trim$(Mid$(ln, p + 1)): n = n + 1
            End If
        End If
    Next
    WriteLog "Site configuration applied: " & n & " value(s).", "INFO"
    ApplySiteConfig = n
End Function

'==========================================================================================
'                                    HOLIDAYS / BUSINESS DAYS
'==========================================================================================
Public Function EasterSunday(ByVal y As Long) As Date
    Dim a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long
    Dim h As Long, i As Long, k As Long, l As Long, m As Long
    a = y Mod 19: b = y \ 100: c = y Mod 100: d = b \ 4: e = b Mod 4
    f = (b + 8) \ 25: g = (b - f + 1) \ 3: h = (19 * a + b - d - g + 15) Mod 30
    i = c \ 4: k = c Mod 4: l = (32 + 2 * e + 2 * i - h - k) Mod 7
    m = (a + 11 * h + 22 * l) \ 451
    EasterSunday = DateSerial(y, (h + l - 7 * m + 114) \ 31, ((h + l - 7 * m + 114) Mod 31) + 1)
End Function

' PT national + Lisbon municipal + TARGET2 closing days (union), one year
Public Function HolidaysForYear(ByVal y As Long) As Collection
    Dim c As New Collection, e As Date
    e = EasterSunday(y)
    c.Add Array(DateSerial(y, 1, 1), "New Year (PT, TARGET2)")
    c.Add Array(e - 2, "Good Friday (PT, TARGET2)")
    c.Add Array(e + 1, "Easter Monday (TARGET2)")
    c.Add Array(DateSerial(y, 4, 25), "Freedom Day (PT)")
    c.Add Array(DateSerial(y, 5, 1), "Labour Day (PT, TARGET2)")
    c.Add Array(e + 60, "Corpus Christi (PT)")
    c.Add Array(DateSerial(y, 6, 10), "Portugal Day (PT)")
    c.Add Array(DateSerial(y, 6, 13), "St Anthony (Lisbon municipal)")
    c.Add Array(DateSerial(y, 8, 15), "Assumption (PT)")
    c.Add Array(DateSerial(y, 10, 5), "Republic Day (PT)")
    c.Add Array(DateSerial(y, 11, 1), "All Saints (PT)")
    c.Add Array(DateSerial(y, 12, 1), "Restoration of Independence (PT)")
    c.Add Array(DateSerial(y, 12, 8), "Immaculate Conception (PT)")
    c.Add Array(DateSerial(y, 12, 25), "Christmas (PT, TARGET2)")
    c.Add Array(DateSerial(y, 12, 26), "Christmas holiday (TARGET2)")
    Set HolidaysForYear = c
End Function

Private Sub WriteHolidayTable(ws As Worksheet, hol As Variant, ByVal nHol As Long)
    Dim y As Long, h As Variant, arr() As Variant, n As Long, i As Long, j As Long, t As Variant
    ws.Cells(CFG_HDR_ROW - 1, CFG_COL_HOL).Value = "Holidays (runs skipped; edit freely)"
    ws.Cells(CFG_HDR_ROW - 1, CFG_COL_HOL).Font.Italic = True
    ws.Cells(CFG_HDR_ROW, CFG_COL_HOL).Resize(1, 2).Value = Array("Holiday", "Name / calendar")
    Hdr ws.Cells(CFG_HDR_ROW, CFG_COL_HOL).Resize(1, 2)
    ws.Range(HOLIDAY_REF).NumberFormat = "yyyy-mm-dd"
    If nHol > 0 Then                                           ' keep the operator's table
        ws.Cells(CFG_HDR_ROW + 1, CFG_COL_HOL).Resize(UBound(hol, 1), 2).Value = hol
        Exit Sub
    End If
    ReDim arr(1 To 60, 1 To 2)
    For y = Year(Date) To Year(Date) + 3
        For Each h In HolidaysForYear(y)
            If Weekday(h(0), vbMonday) <= 5 Then n = n + 1: arr(n, 1) = h(0): arr(n, 2) = h(1)
        Next
    Next
    For i = 1 To n - 1
        For j = i + 1 To n
            If arr(j, 1) < arr(i, 1) Then
                t = arr(i, 1): arr(i, 1) = arr(j, 1): arr(j, 1) = t
                t = arr(i, 2): arr(i, 2) = arr(j, 2): arr(j, 2) = t
            End If
        Next
    Next
    For i = 1 To n
        ws.Cells(CFG_HDR_ROW + i, CFG_COL_HOL).Value = arr(i, 1)
        ws.Cells(CFG_HDR_ROW + i, CFG_COL_HOL + 1).Value = arr(i, 2)
    Next
End Sub

Public Function LoadHolidays() As Object
    Dim d As Object, ws As Worksheet, c As Range, dt As Date
    Set d = NewDict()
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_CONFIG)
    On Error GoTo 0
    If Not ws Is Nothing Then
        For Each c In ws.Range(HOLIDAY_REF).Cells
            If Not IsEmpty(c.Value2) Then
                If ParseDateCell(c.Value2, "", dt) Then d(CLng(dt)) = 1
            End If
        Next
    End If
    Set LoadHolidays = d
End Function

Public Function IsBusinessDay(ByVal d As Date) As Boolean
    If HolidaySet Is Nothing Then Set HolidaySet = NewDict()
    d = DateSerial(Year(d), Month(d), Day(d))
    IsBusinessDay = (Weekday(d, vbMonday) <= 5) And Not HolidaySet.Exists(CLng(d))
End Function

Public Function PrevBusinessDay(ByVal d As Date) As Date
    Dim i As Long
    d = DateSerial(Year(d), Month(d), Day(d)) - 1
    Do While Not IsBusinessDay(d) And i < 30
        d = d - 1: i = i + 1
    Loop
    PrevBusinessDay = d
End Function

Public Function DaysBack(ByVal d As Date) As Long
    d = DateSerial(Year(d), Month(d), Day(d))
    DaysBack = CLng(d - PrevBusinessDay(d))
End Function

'==========================================================================================
'                                    VALIDATION
'==========================================================================================
Public Sub ValidateConfig()
    Dim nF As Long, nW As Long, ok As Boolean
    ok = ValidateConfigCore(nF, nW)
    If Application.Visible Then
        GetSheet(SH_CONFIG).Activate
        MsgBox "Validate Config: " & nF & " FAIL, " & nW & " WARN (details in CONFIG E:G).", _
               IIf(ok, vbInformation, vbExclamation), "BLOT Best Execution"
    End If
End Sub

Public Function ValidateConfigCore(Optional ByRef nFail As Long, Optional ByRef nWarn As Long) As Boolean
    Dim res As Collection, sheetCfg As Object, k As Variant, s As String, e As String, msg As String, t As Date
    On Error GoTo EH
    Set res = New Collection
    Set sheetCfg = ReadConfigSheet()
    LoadConfig

    For Each k In ConfigKeys()
        If Not sheetCfg.Exists(k) Then s = s & k & " "
    Next
    If Len(s) > 0 Then VAdd res, "Keys present", "FAIL", "Missing: " & s & "(run Setup)" Else VAdd res, "Keys present", "PASS", "All keys found."
    If Len(gCfgIssues) > 0 Then VAdd res, "Values", "FAIL", gCfgIssues Else VAdd res, "Values", "PASS", "Readable."

    e = ValidateSteps(OPEN_SEQ)
    If Len(e) > 0 Then VAdd res, "OPEN_SEQ", "FAIL", e Else VAdd res, "OPEN_SEQ", "PASS", OPEN_SEQ
    e = ValidateSteps(EXPORT_SEQ)
    If Len(EXPORT_SEQ) = 0 Then e = "Empty."
    If Len(e) > 0 Then VAdd res, "EXPORT_SEQ", "FAIL", e Else VAdd res, "EXPORT_SEQ", "PASS", EXPORT_SEQ & " (confirm with Test Export 5/5)"
    If DdeAvailable(msg) Then VAdd res, "DDE link", "PASS", msg Else VAdd res, "DDE link", "FAIL", msg

    If Len(EXPORT_FOLDER) = 0 Then
        VAdd res, "EXPORT_FOLDER", "FAIL", "Required: folder where every export is saved."
    Else
        EnsureFolder EXPORT_FOLDER
        e = FolderProblem(EXPORT_FOLDER)
        If Len(e) > 0 Then VAdd res, "EXPORT_FOLDER", "FAIL", e Else VAdd res, "EXPORT_FOLDER", "PASS", EXPORT_FOLDER
    End If

    If ParseTimeStr(RUN_AT, t) Then VAdd res, "RUN_AT", "PASS", Format$(t, "hh:nn:ss") Else VAdd res, "RUN_AT", "FAIL", "'" & RUN_AT & "' is not hh:mm[:ss]."
    Select Case SCHEDULE_MODE
        Case "TASK"
            If TaskExists() Then VAdd res, "SCHEDULE_MODE", "PASS", "TASK: Windows task '" & TASK_NAME & "' installed." Else VAdd res, "SCHEDULE_MODE", "FAIL", "TASK but the Windows task is not installed: click Install Daily Task."
        Case "WORKBOOK"
            VAdd res, "SCHEDULE_MODE", "WARN", "WORKBOOK: runs only while this workbook stays open. Install Daily Task for a fully unattended run."
        Case "OFF"
            VAdd res, "SCHEDULE_MODE", "WARN", "OFF: no scheduled run."
        Case Else
            VAdd res, "SCHEDULE_MODE", "FAIL", "'" & SCHEDULE_MODE & "' is not TASK | WORKBOOK | OFF."
    End Select
    If ThisWorkbook.ReadOnly Then VAdd res, "Workbook", "FAIL", "Opened read-only: history cannot be saved." Else VAdd res, "Workbook", "PASS", ThisWorkbook.FullName
    If ExtensionHardening() = 0 Then
        VAdd res, "Format prompt", "PASS", "Excel's 'format and extension don't match' warning is off for this user."
    Else
        VAdd res, "Format prompt", "WARN", "The workbook reads the export as data (no prompt). If BLOOMBERG opens the file in Excel after exporting, the prompt still appears: click Silence Format Prompt."
    End If
    If POLL_S < 1 Or TIMEOUT_S <= POLL_S Then VAdd res, "TIMEOUT_S / POLL_S", "FAIL", "Need POLL_S >= 1 and TIMEOUT_S > POLL_S." Else VAdd res, "TIMEOUT_S / POLL_S", "PASS", TIMEOUT_S & "s / " & POLL_S & "s"
    If Len(ALERT_TO) = 0 Then VAdd res, "ALERT_TO", "WARN", "No e-mail alerts." Else VAdd res, "ALERT_TO", "PASS", ALERT_TO
    If SHOW_POPUPS Then VAdd res, "SHOW_POPUPS", "WARN", "TRUE blocks unattended runs until dismissed." Else VAdd res, "SHOW_POPUPS", "PASS", "FALSE"

    If MIN_Q <> 3 Then VAdd res, "MIN_Q", "WARN", "Regulatory threshold is 3; now " & MIN_Q & "." Else VAdd res, "MIN_Q", "PASS", "3"
    If GROUP_LEVEL <> "INQUIRY" And GROUP_LEVEL <> "ORDER" Then
        VAdd res, "GROUP_LEVEL", "FAIL", "'" & GROUP_LEVEL & "' is not INQUIRY | ORDER."
    ElseIf GROUP_LEVEL = "ORDER" Then
        VAdd res, "GROUP_LEVEL", "WARN", "ORDER: requests for the same ISIN/side/day are counted together (the export has no order id). Compliance decision."
    Else
        VAdd res, "GROUP_LEVEL", "PASS", "INQUIRY"
    End If
    If Len(Trim$(Replace(FILL_STATUS, "|", ""))) = 0 Then VAdd res, "FILL_STATUS", "FAIL", "Empty: no trade can be recognised as filled." Else VAdd res, "FILL_STATUS", "PASS", FILL_STATUS
    If ID_SUFFIX_LEN < 0 Or ID_SUFFIX_LEN > 6 Then VAdd res, "ID_SUFFIX_LEN", "FAIL", "Expected 0..6." Else VAdd res, "ID_SUFFIX_LEN", "PASS", CStr(ID_SUFFIX_LEN)
    If Len(DATE_FORMAT_TEXT) > 0 Then
        e = ValidateDateFormat(DATE_FORMAT_TEXT)
        If Len(e) > 0 Then VAdd res, "DATE_FORMAT_TEXT", "FAIL", e Else VAdd res, "DATE_FORMAT_TEXT", "PASS", DATE_FORMAT_TEXT
    End If
    If DealerMapCount() = 0 Then VAdd res, "DEALER_MAP", "WARN", "Empty: different spellings of one bank count as different dealers." Else VAdd res, "DEALER_MAP", "PASS", DealerMapCount() & " aliases."

    WriteValidation res, nFail, nWarn
    ValidateConfigCore = (nFail = 0)
    WriteLog "Validate Config: " & nFail & " FAIL, " & nWarn & " WARN.", IIf(nFail = 0, "INFO", "WARN")
    Exit Function
EH:
    VAdd res, "Validation", "FAIL", "Error: " & Err.Description
    WriteValidation res, nFail, nWarn
End Function

Private Sub VAdd(res As Collection, ByVal chk As String, ByVal lvl As String, ByVal detail As String)
    res.Add Array(chk, lvl, detail)
End Sub

Private Sub WriteValidation(res As Collection, ByRef nFail As Long, ByRef nWarn As Long)
    Dim ws As Worksheet, r As Long, it As Variant, c As Long
    Set ws = GetSheet(SH_CONFIG)
    ws.Range(ws.Cells(CFG_HDR_ROW - 1, CFG_COL_CHK), ws.Cells(200, CFG_COL_CHK + 2)).Clear
    ws.Cells(CFG_HDR_ROW - 1, CFG_COL_CHK).Value = "Validation " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(CFG_HDR_ROW, CFG_COL_CHK).Resize(1, 3).Value = Array("Check", "Result", "Detail")
    Hdr ws.Cells(CFG_HDR_ROW, CFG_COL_CHK).Resize(1, 3)
    nFail = 0: nWarn = 0: r = CFG_HDR_ROW
    For Each it In res
        r = r + 1
        ws.Cells(r, CFG_COL_CHK).Value = it(0)
        ws.Cells(r, CFG_COL_CHK + 1).Value = it(1)
        ws.Cells(r, CFG_COL_CHK + 2).Value = "'" & it(2)
        Select Case it(1)
            Case "PASS": c = RGB(198, 239, 206)
            Case "WARN": c = RGB(255, 235, 156): nWarn = nWarn + 1
            Case Else:   c = RGB(255, 199, 206): nFail = nFail + 1
        End Select
        ws.Cells(r, CFG_COL_CHK + 1).Interior.Color = c
    Next
End Sub

Private Function FolderProblem(ByVal p As String) As String
    Dim fso As Object, f As String, h As Integer
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(p) Then FolderProblem = "Folder not found: " & p: Exit Function
    f = p & "\~blot_write_test.tmp"
    On Error GoTo NoWrite
    h = FreeFile
    Open f For Output As #h
    Close #h
    Kill f
    Exit Function
NoWrite:
    FolderProblem = "Folder not writable: " & p
End Function

Public Sub EnsureFolder(ByVal p As String)
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(p) Then
        If fso.FolderExists(fso.GetParentFolderName(p)) Then fso.CreateFolder p
    End If
End Sub

Private Function DealerMapCount() As Long
    DealerMapCount = LoadDealerMap().Count
End Function

'==========================================================================================
'                                    SMALL PARSERS (locale-free)
'==========================================================================================
Public Function ParseTimeStr(ByVal s As String, ByRef t As Date) As Boolean
    Dim p As Variant, h As Long, m As Long, sec As Long
    p = Split(Trim$(s), ":")
    If UBound(p) < 1 Or UBound(p) > 2 Then Exit Function
    If Not IsDigits(CStr(p(0))) Or Not IsDigits(CStr(p(1))) Then Exit Function
    h = CLng(p(0)): m = CLng(p(1))
    If UBound(p) = 2 Then
        If Not IsDigits(CStr(p(2))) Then Exit Function
        sec = CLng(p(2))
    End If
    If h > 23 Or m > 59 Or sec > 59 Then Exit Function
    t = TimeSerial(h, m, sec)
    ParseTimeStr = True
End Function

Public Function IsDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Not (Mid$(s, i, 1) Like "#") Then Exit Function
    Next
    IsDigits = True
End Function

Public Function IsPlainNumber(ByVal s As String) As Boolean
    Dim p As Long
    s = Trim$(s)
    If Left$(s, 1) = "-" Then s = Mid$(s, 2)
    p = InStr(1, s, ".")
    If p = 0 Then
        IsPlainNumber = IsDigits(s)
    Else
        IsPlainNumber = IsDigits(Left$(s, p - 1)) And IsDigits(Mid$(s, p + 1))
    End If
End Function

Public Function InvariantNum(ByVal x As Double) As String
    Dim s As String
    If x = Int(x) And Abs(x) < 1E+15 Then
        InvariantNum = Format$(x, "0")
    Else
        s = Trim$(Str$(x))
        If Left$(s, 1) = "." Then s = "0" & s
        If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
        InvariantNum = s
    End If
End Function
