Attribute VB_Name = "modConfig"
Option Explicit
'==========================================================================================
' BLOT Best-Execution monitor - CONFIG sheet loader / validator / business-day calendar
'------------------------------------------------------------------------------------------
'   LoadConfig        reads the CONFIG sheet (Key | Value | Description) into the public
'                     variables below; called at the start of every public entry point
'   LoadDefaults      same, ignoring the sheet (used by the self-tests)
'   ApplyConfig d     defaults overridden by the keys in dictionary d
'   Cfg(key, def)     raw value of one key on the CONFIG sheet
'   ValidateConfig    checks the configuration and writes a PASS/WARN/FAIL list on CONFIG
'==========================================================================================

'------------------------------------ SHEETS ---------------------------------------------
Public Const SH_README As String = "README"
Public Const SH_DASH   As String = "DASHBOARD"
Public Const SH_CONFIG As String = "CONFIG"
Public Const SH_MAP    As String = "DEALER_MAP"
Public Const SH_DEALS  As String = "DEALS"
Public Const SH_HIST   As String = "HIST"
Public Const SH_HD     As String = "HIST_DAILY"
Public Const SH_RAW    As String = "RAW"
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

'------------------------------------ CONFIG VALUES --------------------------------------
Public SOURCE_MODE As String, BBG_TITLE_PREFIX As String, BBG_PANEL As String, BBG_WINDOW As String
Public BBG_PROCESS As String, EXPORT_CANDIDATES As String, CALIB_WAIT_S As Long
Public OPEN_SEQ As String, EXPORT_SEQ As String, STEP_GAP_S As Double
Public EXPORT_FOLDER As String, ARCHIVE_DIR As String
Public OUTLOOK_SUBJECT As String, OUTLOOK_FOLDER As String, CSV_DELIM As String
Public CLOSE_EXPORT As Boolean, TIMEOUT_S As Long, POLL_S As Long, TEST_EXPORT_RUNS As Long
Public RUN_AT As String, HOLIDAYS_REF As String
Public ALERT_TO As String, SHOW_POPUPS As Boolean, SEND_OK_EMAIL As Boolean
Public MIN_Q As Long, DEALER_MODE As String, GROUP_LEVEL As String, LIST_DELIMS As String
Public EXCL_STATUS As String, DATE_FORMAT_TEXT As String
Public HDR_DEALERS As String, HDR_RFQ As String, HDR_ORDER As String, HDR_ID As String
Public HDR_SEC As String, HDR_DATE As String, HDR_SIDE As String, HDR_QTY As String
Public HDR_STATUS As String, HDR_TIME As String, HDR_TRADER As String

Public HolidaySet As Object          ' Long(date serial) -> holiday name
Public gCfgIssues As String          ' values that could not be parsed by the last ApplyConfig

Private gDefs As Collection          ' Array(section, key, default, description, flag)
Private gDefIdx As Object            ' key -> index in gDefs

'==========================================================================================
'                                    DEFAULTS
'==========================================================================================
' flag: "U" = must be filled / calibrated by the operator (highlighted), "" = sensible default
Private Sub BuildDefs()
    If Not gDefs Is Nothing Then Exit Sub
    Set gDefs = New Collection
    Set gDefIdx = NewDict()

    AddDef "SOURCE", "SOURCE_MODE", "TERMINAL", "TERMINAL (drive the Bloomberg Terminal) | OUTLOOK (scheduled report by e-mail) | FOLDER (file drop / SFTP share).", ""
    AddDef "SOURCE", "BBG_TITLE_PREFIX", "1-BLOOMBERG", "Exact start of the Bloomberg panel window title (panel 1 = '1-BLOOMBERG').", ""
    AddDef "SOURCE", "BBG_PANEL", "<blp-1>", "DDE panel prefix. Must match the title number: '1-' <-> '<blp-1>'.", ""
    AddDef "SOURCE", "OPEN_SEQ", "KEYS:BLOT 2<GO>|WAIT:3|KEYS:BLOT -1<GO>|WAIT:5", "Steps that open the BLOT view of the previous day. Tokens: KEYS:text (DDE, <GO>=GO key) | SEND:keys (SendKeys) | CLICK:x,y | RCLICK:x,y | DCLICK:x,y | WAIT:s. {DAYS_BACK} = calendar days back to the previous business day (Mon=3).", ""
    AddDef "SOURCE", "EXPORT_SEQ", "", "Steps that export the BLOT grid. Fill it with the 'Record Export' button (do the export once by hand). Accept only after Test Export succeeds 5/5.", "U"
    AddDef "SOURCE", "EXPORT_CANDIDATES", "KEYS:2<GO>|WAIT:2|KEYS:2<GO>|WAIT:5 || KEYS:2<GO>|WAIT:2|SEND:{DOWN}{DOWN}~|WAIT:5 || KEYS:2<GO>|WAIT:2|SEND:{DOWN}~|WAIT:5", "Sequences tried by 'Calibrate Export' (separated by ||). Default: Export button 2<GO>, then the 2nd item 'XLS - Excel format'.", ""
    AddDef "SOURCE", "CALIB_WAIT_S", "25", "Seconds Calibrate Export waits for the export of each candidate.", ""
    AddDef "SOURCE", "BBG_PROCESS", "wintrv*.exe|bplus*.exe|blp*.exe|bloomberg*.exe|bbg*.exe|bterm*.exe", "Executable names of the Bloomberg Terminal, used to find its window when the title does not match BBG_TITLE_PREFIX.", ""
    AddDef "SOURCE", "BBG_WINDOW", "", "Bloomberg panel size when EXPORT_SEQ was recorded (set by Record Export, e.g. 1936x1056,MAX). Clicks need the same size.", ""
    AddDef "SOURCE", "STEP_GAP_S", "0.7", "Seconds between steps.", ""
    AddDef "SOURCE", "EXPORT_FOLDER", "", "FOLDER mode: drop folder. TERMINAL mode: optional, watched too when the export is saved to a file.", "U"
    AddDef "SOURCE", "ARCHIVE_DIR", "", "Folder for the dated copy of every raw export (regulatory evidence). Required.", "U"
    AddDef "SOURCE", "OUTLOOK_SUBJECT", "BLOT", "OUTLOOK mode: subject filter of the scheduled report e-mail.", ""
    AddDef "SOURCE", "OUTLOOK_FOLDER", "", "OUTLOOK mode: subfolder under Inbox (e.g. 'Reports\BLOT'); empty = Inbox.", ""
    AddDef "SOURCE", "CSV_DELIM", "AUTO", "CSV delimiter: AUTO | ; | , | TAB.", ""
    AddDef "SOURCE", "CLOSE_EXPORT", "TRUE", "Close the export workbook after parsing (and quit its Excel instance if it is then empty).", ""
    AddDef "SOURCE", "TIMEOUT_S", "300", "Seconds to wait for the export before a TIMEOUT alert.", ""
    AddDef "SOURCE", "POLL_S", "3", "Seconds between checks for the export.", ""
    AddDef "SOURCE", "TEST_EXPORT_RUNS", "5", "Consecutive cycles run by 'Test Export' (the acceptance criterion is 5/5).", ""

    AddDef "SCHEDULE", "RUN_AT", "08:30:00", "Daily run time (hh:mm:ss). The workbook must be open (or opened by Windows Task Scheduler before this time).", ""
    AddDef "SCHEDULE", "HOLIDAYS", "CONFIG!$I$5:$I$400", "Range with the holiday dates (runs are skipped; used for the previous business day). Default: the table in column I of this sheet.", ""

    AddDef "ALERTS", "ALERT_TO", "", "E-mail recipients (';'-separated). Outlook e-mail on alert, timeout or error. Empty = no e-mail.", "U"
    AddDef "ALERTS", "SHOW_POPUPS", "TRUE", "Popups on alerts. Set FALSE for unattended runs (a popup blocks Excel until dismissed).", "U"
    AddDef "ALERTS", "SEND_OK_EMAIL", "FALSE", "TRUE = also e-mail a short summary when everything is OK (daily heartbeat).", ""

    AddDef "RULE", "MIN_Q", "3", "Regulatory minimum of counterparties asked per deal (Banco de Portugal: 3).", ""
    AddDef "RULE", "DEALER_MODE", "AUTO", "AUTO | ROWS (one row per dealer) | LIST (dealers in one cell) | COUNT (numeric field, not auditable). Pin it after calibration.", "U"
    AddDef "RULE", "GROUP_LEVEL", "RFQ", "Unit of analysis: RFQ (each inquiry) | ORDER (all inquiries of one order, needs an order-id column). Compliance decision.", "U"
    AddDef "RULE", "LIST_DELIMS", ";,/|", "LIST mode: characters separating dealer names in one cell (line breaks always separate).", ""
    AddDef "RULE", "EXCL_STATUS", "cancel|cxl|reject|expired", "Statuses meaning 'not executed' ('|'-separated, substring match). A deal is dropped only if ALL its rows match.", "U"
    AddDef "RULE", "DATE_FORMAT_TEXT", "", "Format of text dates in the export, e.g. MM/DD/YYYY, DD/MM/YYYY, DD-MMM-YY, YYYYMMDD. Empty = only real dates and ISO yyyy-mm-dd.", "U"

    AddDef "HEADERS", "HDR_DEALERS", "dealers in comp|dealers in competition|dealers queried|num dealers|# dealers|no. of dealers|nb dealers|dealers|dealer|counterparties|counterparty", "Column with the counterparties ASKED (not quotes received). Put the exact export header first. Lowercase, '|'-separated.", "U"
    AddDef "HEADERS", "HDR_RFQ", "inquiry id|inquiry number|rfq id|rfq number|request id|quote request id", "RFQ / inquiry id (preferred trade key).", "U"
    AddDef "HEADERS", "HDR_ORDER", "order id|order number|parent order id|order", "Order id (used only when GROUP_LEVEL = ORDER).", ""
    AddDef "HEADERS", "HDR_ID", "trade id|ticket number|ticket|tkt|order number|order id|id", "Ticket / trade id (used as key only if every dealer row carries it).", ""
    AddDef "HEADERS", "HDR_SEC", "isin|security|description|ticker|instrument", "Security (ISIN first: used for integrity and the composite key).", ""
    AddDef "HEADERS", "HDR_DATE", "trade date|tradedate|date", "Trade date.", ""
    AddDef "HEADERS", "HDR_SIDE", "side|b/s|buy/sell", "Side.", ""
    AddDef "HEADERS", "HDR_QTY", "quantity|notional|amount|qty|size", "Quantity / notional.", ""
    AddDef "HEADERS", "HDR_STATUS", "status|order status|state", "Status.", ""
    AddDef "HEADERS", "HDR_TIME", "inquiry time|rfq time|request time|create time|time", "RFQ SEND time (common to all dealer rows; not the quote time).", ""
    AddDef "HEADERS", "HDR_TRADER", "trader|trader name|trader uuid|user", "Trader.", ""
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
    SOURCE_MODE = UCase$(CfgStr(d, "SOURCE_MODE"))
    BBG_TITLE_PREFIX = CfgStr(d, "BBG_TITLE_PREFIX")
    BBG_PANEL = CfgStr(d, "BBG_PANEL")
    BBG_WINDOW = CfgStr(d, "BBG_WINDOW")
    BBG_PROCESS = CfgStr(d, "BBG_PROCESS")
    EXPORT_CANDIDATES = CfgStr(d, "EXPORT_CANDIDATES")
    CALIB_WAIT_S = CLng(CfgNum(d, "CALIB_WAIT_S"))
    OPEN_SEQ = CfgStr(d, "OPEN_SEQ")
    EXPORT_SEQ = CfgStr(d, "EXPORT_SEQ")
    STEP_GAP_S = CfgNum(d, "STEP_GAP_S")
    EXPORT_FOLDER = StripSlash(CfgStr(d, "EXPORT_FOLDER"))
    ARCHIVE_DIR = StripSlash(CfgStr(d, "ARCHIVE_DIR"))
    OUTLOOK_SUBJECT = CfgStr(d, "OUTLOOK_SUBJECT")
    OUTLOOK_FOLDER = CfgStr(d, "OUTLOOK_FOLDER")
    CSV_DELIM = UCase$(CfgStr(d, "CSV_DELIM"))
    CLOSE_EXPORT = CfgBool(d, "CLOSE_EXPORT")
    TIMEOUT_S = CLng(CfgNum(d, "TIMEOUT_S"))
    POLL_S = CLng(CfgNum(d, "POLL_S"))
    TEST_EXPORT_RUNS = CLng(CfgNum(d, "TEST_EXPORT_RUNS"))
    RUN_AT = CfgStr(d, "RUN_AT")
    HOLIDAYS_REF = CfgStr(d, "HOLIDAYS")
    ALERT_TO = CfgStr(d, "ALERT_TO")
    SHOW_POPUPS = CfgBool(d, "SHOW_POPUPS")
    SEND_OK_EMAIL = CfgBool(d, "SEND_OK_EMAIL")
    MIN_Q = CLng(CfgNum(d, "MIN_Q"))
    DEALER_MODE = UCase$(CfgStr(d, "DEALER_MODE"))
    GROUP_LEVEL = UCase$(CfgStr(d, "GROUP_LEVEL"))
    LIST_DELIMS = CfgStr(d, "LIST_DELIMS")
    EXCL_STATUS = LCase$(CfgStr(d, "EXCL_STATUS"))
    DATE_FORMAT_TEXT = UCase$(CfgStr(d, "DATE_FORMAT_TEXT"))
    HDR_DEALERS = LCase$(CfgStr(d, "HDR_DEALERS"))
    HDR_RFQ = LCase$(CfgStr(d, "HDR_RFQ"))
    HDR_ORDER = LCase$(CfgStr(d, "HDR_ORDER"))
    HDR_ID = LCase$(CfgStr(d, "HDR_ID"))
    HDR_SEC = LCase$(CfgStr(d, "HDR_SEC"))
    HDR_DATE = LCase$(CfgStr(d, "HDR_DATE"))
    HDR_SIDE = LCase$(CfgStr(d, "HDR_SIDE"))
    HDR_QTY = LCase$(CfgStr(d, "HDR_QTY"))
    HDR_STATUS = LCase$(CfgStr(d, "HDR_STATUS"))
    HDR_TIME = LCase$(CfgStr(d, "HDR_TIME"))
    HDR_TRADER = LCase$(CfgStr(d, "HDR_TRADER"))
    Set HolidaySet = LoadHolidays(HOLIDAYS_REF)
End Sub

' key (UPPERCASE) -> raw cell value, for every key-like row of the CONFIG sheet
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
    Dim i As Long, c As String
    If Len(k) = 0 Then Exit Function
    For i = 1 To Len(k)
        c = Mid$(k, i, 1)
        If Not (c Like "[A-Z0-9_]") Then Exit Function
    Next
    IsConfigKeyName = True
End Function

' Raw value of one key on the CONFIG sheet (def when missing)
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
            If k = "RUN_AT" And CDbl(x) >= 0 And CDbl(x) < 1 Then
                ValToStr = Format$(CDate(x), "hh:nn:ss")
            Else
                ValToStr = InvariantNum(CDbl(x))
            End If
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
        Case "TRUE", "VERDADEIRO", "1", "YES", "Y", "SIM", "S"
            CfgBool = True
        Case "FALSE", "FALSO", "0", "NO", "N", "NAO"
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
'                                    CONFIG SHEET (Setup)
'==========================================================================================
' Creates the CONFIG layout. Existing values are never overwritten; missing keys are added.
Public Sub WriteConfigSheet()
    Dim ws As Worksheet, d As Variant, have As Object, lr As Long, i As Long, r As Long
    Dim lastSec As String, fresh As Boolean, kv As Variant

    BuildDefs
    Set ws = GetSheet(SH_CONFIG)
    fresh = (Application.WorksheetFunction.CountA(ws.Cells) = 0)

    ws.Rows(1).RowHeight = 34
    ws.Range("A2").Value = "CONFIG - the only sheet to edit (plus DEALER_MAP). Yellow = to fill / calibrate."
    ws.Range("A2").Font.Bold = True: ws.Range("A2").Font.Size = 13
    ws.Range("A3").Value = "Edit column B, then click 'Validate Config'. Values are text: booleans TRUE/FALSE, decimals with '.' or ','."
    ws.Range("A3").Font.Italic = True
    ws.Cells(CFG_HDR_ROW, 1).Resize(1, 3).Value = Array("Key", "Value", "Description")
    ws.Cells(CFG_HDR_ROW, 1).Resize(1, 3).Font.Bold = True
    ws.Cells(CFG_HDR_ROW, 1).Resize(1, 3).Interior.Color = RGB(217, 225, 242)

    Set have = NewDict()
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr > CFG_HDR_ROW Then
        For i = CFG_HDR_ROW + 1 To lr
            kv = ws.Cells(i, 1).Value
            If Not IsError(kv) Then
                If Len(Trim$(CStr(kv))) > 0 Then have(UCase$(Trim$(CStr(kv)))) = i
            End If
        Next
    End If
    r = Application.WorksheetFunction.Max(lr, CFG_HDR_ROW) + 1
    If Not fresh Then
        For Each d In gDefs
            If Not have.Exists(d(1)) Then
                If lastSec <> "ADDED" Then
                    r = r + 1
                    ws.Cells(r, 1).Value = "[ADDED BY SETUP]": ws.Cells(r, 1).Font.Bold = True
                    r = r + 1: lastSec = "ADDED"
                End If
                WriteCfgRow ws, r, d: r = r + 1
            Else
                ws.Cells(have(d(1)), CFG_COL_DESC).Value = d(3)          ' refresh description only
            End If
        Next
    Else
        r = CFG_HDR_ROW + 1
        For Each d In gDefs
            If d(0) <> lastSec Then
                If lastSec <> "" Then r = r + 1
                ws.Cells(r, 1).Value = "[" & d(0) & "]"
                ws.Cells(r, 1).Font.Bold = True: ws.Cells(r, 1).Font.Color = RGB(68, 84, 106)
                r = r + 1: lastSec = d(0)
            End If
            WriteCfgRow ws, r, d: r = r + 1
        Next
    End If

    ws.Columns(1).ColumnWidth = 20
    ws.Columns(2).ColumnWidth = 58
    ws.Columns(3).ColumnWidth = 90
    ws.Columns(3).WrapText = False
    ws.Columns(4).ColumnWidth = 2
    ws.Columns(5).ColumnWidth = 30: ws.Columns(6).ColumnWidth = 8: ws.Columns(7).ColumnWidth = 70
    ws.Columns(8).ColumnWidth = 2
    ws.Columns(9).ColumnWidth = 12: ws.Columns(10).ColumnWidth = 38

    WriteHolidayTable ws
    AddSheetButtons ws, Array("Validate Config|ValidateConfig", "Calibrate Export|CalibrateExport", "Record Export|RecordExport", "Test Export|TestExport", _
                              "Run Now|RunBlotReport", "List Windows|ListWindows", "Import Settings|ImportSettings")
End Sub

Private Sub WriteCfgRow(ws As Worksheet, ByVal r As Long, d As Variant)
    ws.Cells(r, CFG_COL_KEY).Value = d(1)
    ws.Cells(r, CFG_COL_VAL).NumberFormat = "@"
    ws.Cells(r, CFG_COL_VAL).Value = d(2)
    ws.Cells(r, CFG_COL_DESC).Value = d(3)
    If d(4) = "U" Then ws.Cells(r, CFG_COL_VAL).Interior.Color = RGB(255, 242, 204)
End Sub

' Writes one CONFIG value (as text)
Public Sub SetConfigValue(ByVal key As String, ByVal v As String)
    Dim ws As Worksheet, f As Range
    Set ws = GetSheet(SH_CONFIG)
    Set f = ws.Columns(CFG_COL_KEY).Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If f Is Nothing Then Err.Raise vbObjectError + 2, , "Key " & key & " not found on CONFIG (run Setup)."
    ws.Cells(f.Row, CFG_COL_VAL).NumberFormat = "@"
    ws.Cells(f.Row, CFG_COL_VAL).Value = v
End Sub

' Applies a site file (KEY=value per line, '#' = comment) to CONFIG. Used by the build. Returns the count.
Public Function ApplySiteConfig(ByVal path As String) As Long
    Dim txt As String, ln As Variant, p As Long, k As String, v As String, mine As Object, kk As Variant, n As Long
    Set mine = NewDict()
    For Each kk In ConfigKeys(): mine(kk) = 1: Next
    txt = Replace(Replace(ReadTextFile(path), vbCrLf, vbLf), vbCr, vbLf)
    For Each ln In Split(txt, vbLf)
        ln = Trim$(CStr(ln))
        If Len(ln) > 0 And Left$(ln, 1) <> "#" Then
            p = InStr(1, ln, "=")
            If p > 1 Then
                k = UCase$(Trim$(Left$(ln, p - 1))): v = Trim$(Mid$(ln, p + 1))
                If Not mine.Exists(k) Then Err.Raise vbObjectError + 3, , "Unknown CONFIG key in site file: " & k
                SetConfigValue k, v: n = n + 1
            End If
        End If
    Next
    WriteLog "Site configuration applied: " & n & " value(s).", "INFO"
    ApplySiteConfig = n
End Function

' Copies CONFIG values, holidays and DEALER_MAP from another (older) BLOT_BestEx workbook
Public Sub ImportSettings()
    Dim f As Variant, wb As Workbook, src As Worksheet, v As Variant, i As Long, lr As Long, k As String, n As Long
    Dim mine As Object, skipped As String, sm As Worksheet, dm As Worksheet
    f = Application.GetOpenFilename("BLOT workbook (*.xlsm),*.xlsm", , "Old BLOT_BestEx workbook to copy the settings from")
    If VarType(f) = vbBoolean Then Exit Sub
    If StrComp(CStr(f), ThisWorkbook.FullName, vbTextCompare) = 0 Then MsgBox "Select the OLD workbook, not this one.", vbExclamation: Exit Sub
    Set mine = NewDict()
    For Each v In ConfigKeys(): mine(v) = 1: Next
    Application.EnableEvents = False
    Set wb = Workbooks.Open(CStr(f), UpdateLinks:=0, ReadOnly:=True)
    Application.EnableEvents = True
    On Error GoTo Done
    Set src = wb.Worksheets(SH_CONFIG)
    lr = src.Cells(src.Rows.Count, CFG_COL_KEY).End(xlUp).Row
    For i = CFG_HDR_ROW + 1 To lr
        k = UCase$(Trim$(CStr(src.Cells(i, CFG_COL_KEY).Value)))
        If IsConfigKeyName(k) Then
            If mine.Exists(k) Then
                SetConfigValue k, ValToStr(src.Cells(i, CFG_COL_VAL).Value, k): n = n + 1
            Else
                skipped = skipped & k & " "
            End If
        End If
    Next
    v = src.Range(src.Cells(CFG_HDR_ROW + 1, CFG_COL_HOL), src.Cells(400, CFG_COL_HOL + 1)).Value
    GetSheet(SH_CONFIG).Range(GetSheet(SH_CONFIG).Cells(CFG_HDR_ROW + 1, CFG_COL_HOL), GetSheet(SH_CONFIG).Cells(400, CFG_COL_HOL + 1)).Value = v
    Set sm = wb.Worksheets(SH_MAP): Set dm = GetSheet(SH_MAP)
    lr = sm.Cells(sm.Rows.Count, 1).End(xlUp).Row
    If lr >= 2 Then
        dm.Range("A2:B" & dm.Rows.Count).ClearContents
        dm.Range("A2:B" & lr).NumberFormat = "@"
        dm.Range("A2:B" & lr).Value = sm.Range("A2:B" & lr).Value
    End If
Done:
    wb.Close False
    WriteLog "Import Settings from " & f & ": " & n & " CONFIG value(s), holidays and DEALER_MAP copied." & IIf(Len(skipped) > 0, " Not copied (unknown keys): " & skipped, "") & IIf(Err.Number <> 0, " ERROR: " & Err.Description, ""), IIf(Err.Number <> 0, "ERROR", "INFO")
    If Application.Visible Then MsgBox n & " CONFIG value(s), holidays and DEALER_MAP copied from" & vbLf & f & vbLf & vbLf & "Now click Validate Config.", vbInformation, "Import Settings"
End Sub

' Appends a step token to a sequence key (used by CaptureClickPoint)
Public Sub AppendStepToConfig(ByVal key As String, ByVal tok As String)
    Dim ws As Worksheet, f As Range, cur As String
    Set ws = GetSheet(SH_CONFIG)
    Set f = ws.Columns(CFG_COL_KEY).Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If f Is Nothing Then Err.Raise vbObjectError + 1, , "Key " & key & " not found on CONFIG (run Setup)."
    cur = Trim$(CStr(f.Offset(0, CFG_COL_VAL - CFG_COL_KEY).Value))
    f.Offset(0, CFG_COL_VAL - CFG_COL_KEY).NumberFormat = "@"
    If Len(cur) = 0 Then
        f.Offset(0, CFG_COL_VAL - CFG_COL_KEY).Value = tok
    Else
        f.Offset(0, CFG_COL_VAL - CFG_COL_KEY).Value = cur & "|WAIT:1|" & tok
    End If
End Sub

'==========================================================================================
'                                    HOLIDAYS / BUSINESS DAYS
'==========================================================================================
Public Function EasterSunday(ByVal y As Long) As Date
    Dim a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long
    Dim h As Long, i As Long, k As Long, l As Long, m As Long, mo As Long, dy As Long
    a = y Mod 19: b = y \ 100: c = y Mod 100: d = b \ 4: e = b Mod 4
    f = (b + 8) \ 25: g = (b - f + 1) \ 3: h = (19 * a + b - d - g + 15) Mod 30
    i = c \ 4: k = c Mod 4: l = (32 + 2 * e + 2 * i - h - k) Mod 7
    m = (a + 11 * h + 22 * l) \ 451
    mo = (h + l - 7 * m + 114) \ 31: dy = ((h + l - 7 * m + 114) Mod 31) + 1
    EasterSunday = DateSerial(y, mo, dy)
End Function

' PT national holidays + Lisbon municipal holiday + TARGET2 closing days (union), one year
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

Private Sub WriteHolidayTable(ws As Worksheet)
    Dim y As Long, h As Variant, r As Long, arr() As Variant, n As Long, i As Long, j As Long, t As Variant
    ws.Cells(CFG_HDR_ROW - 1, CFG_COL_HOL).Value = "Holidays (edit freely; weekends are always skipped)"
    ws.Cells(CFG_HDR_ROW - 1, CFG_COL_HOL).Font.Italic = True
    ws.Cells(CFG_HDR_ROW, CFG_COL_HOL).Resize(1, 2).Value = Array("Holiday", "Name / calendar")
    ws.Cells(CFG_HDR_ROW, CFG_COL_HOL).Resize(1, 2).Font.Bold = True
    ws.Cells(CFG_HDR_ROW, CFG_COL_HOL).Resize(1, 2).Interior.Color = RGB(217, 225, 242)
    If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(CFG_HDR_ROW + 1, CFG_COL_HOL), ws.Cells(400, CFG_COL_HOL))) > 0 Then Exit Sub

    ReDim arr(1 To 4 * 15, 1 To 2)
    For y = Year(Date) To Year(Date) + 3
        For Each h In HolidaysForYear(y)
            If Weekday(h(0), vbMonday) <= 5 Then n = n + 1: arr(n, 1) = h(0): arr(n, 2) = h(1)
        Next
    Next
    For i = 1 To n - 1                                         ' sort by date
        For j = i + 1 To n
            If arr(j, 1) < arr(i, 1) Then
                t = arr(i, 1): arr(i, 1) = arr(j, 1): arr(j, 1) = t
                t = arr(i, 2): arr(i, 2) = arr(j, 2): arr(j, 2) = t
            End If
        Next
    Next
    r = CFG_HDR_ROW + 1
    ws.Cells(r, CFG_COL_HOL).Resize(n, 1).NumberFormat = "yyyy-mm-dd"
    For i = 1 To n
        ws.Cells(r + i - 1, CFG_COL_HOL).Value = arr(i, 1)
        ws.Cells(r + i - 1, CFG_COL_HOL + 1).Value = arr(i, 2)
    Next
End Sub

Public Function LoadHolidays(ByVal ref As String) As Object
    Dim d As Object, rg As Range, c As Range, dt As Date
    Set d = NewDict()
    Set rg = ResolveRange(ref)
    If Not rg Is Nothing Then
        For Each c In rg.Cells
            If Not IsEmpty(c.Value2) Then
                If ParseDateCell(c.Value2, "", dt) Then d(CLng(dt)) = 1
            End If
        Next
    End If
    Set LoadHolidays = d
End Function

' "SHEET!A1:A10", "'SHEET'!A1:A10" or a workbook name
Public Function ResolveRange(ByVal ref As String) As Range
    Dim p As Long, sh As String, addr As String
    ref = Trim$(ref)
    If Len(ref) = 0 Then Exit Function
    On Error Resume Next
    p = InStrRev(ref, "!")
    If p > 0 Then
        sh = Replace(Left$(ref, p - 1), "'", "")
        addr = Mid$(ref, p + 1)
        Set ResolveRange = ThisWorkbook.Worksheets(sh).Range(addr)
    Else
        Set ResolveRange = ThisWorkbook.Names(ref).RefersToRange
    End If
    On Error GoTo 0
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
        MsgBox "Validate Config: " & nF & " FAIL, " & nW & " WARN." & vbLf & vbLf & _
               IIf(ok, IIf(nW = 0, "All green.", "No blocking issue; review the WARN lines (column E:G)."), _
                   "Fix the FAIL lines (column E:G) before running."), _
               IIf(ok, vbInformation, vbExclamation), "BLOT Best Execution"
    End If
End Sub

Public Function ValidateConfigCore(Optional ByRef nFail As Long, Optional ByRef nWarn As Long) As Boolean
    Dim res As Collection, sheetCfg As Object, k As Variant, s As String, e As String, msg As String
    Dim t As Date, n As Long, okDde As Boolean, ol As Object, fld As Object, ws As Worksheet
    Dim dup As String, nMap As Long, nyH As Long, cyH As Long, h As Variant, r As Long, x As Variant

    On Error GoTo EH
    Set res = New Collection
    Set sheetCfg = ReadConfigSheet()
    LoadConfig

    '--- presence / parse ---
    s = ""
    For Each k In ConfigKeys()
        If Not sheetCfg.Exists(k) Then s = s & k & " "
    Next
    If Len(s) > 0 Then VAdd res, "Keys present", "FAIL", "Missing: " & s & "(run Setup)" Else VAdd res, "Keys present", "PASS", "All keys found on CONFIG."
    If Len(gCfgIssues) > 0 Then VAdd res, "Values parse", "FAIL", gCfgIssues Else VAdd res, "Values parse", "PASS", "Numbers / booleans readable."

    '--- enums ---
    VEnum res, "SOURCE_MODE", SOURCE_MODE, "TERMINAL|OUTLOOK|FOLDER"
    VEnum res, "GROUP_LEVEL", GROUP_LEVEL, "RFQ|ORDER"
    If GROUP_LEVEL = "ORDER" Then VAdd res, "GROUP_LEVEL policy", "WARN", "ORDER = all RFQs of one order counted together (compliance decision; needs an order-id column in the export, see LOG 'Header mapping')."
    VEnum res, "CSV_DELIM", CSV_DELIM, "AUTO|;|,|TAB"
    If InStr(1, "|AUTO|ROWS|LIST|COUNT|", "|" & DEALER_MODE & "|") = 0 Then
        VAdd res, "DEALER_MODE", "FAIL", "'" & DEALER_MODE & "' is not AUTO|ROWS|LIST|COUNT."
    ElseIf DEALER_MODE = "AUTO" Then
        VAdd res, "DEALER_MODE", "WARN", "AUTO: pin ROWS/LIST/COUNT once the real export layout is known (see LOG 'mode=')."
    ElseIf DEALER_MODE = "COUNT" Then
        VAdd res, "DEALER_MODE", "WARN", "COUNT is not auditable (no dealer names). Prefer ROWS."
    Else
        VAdd res, "DEALER_MODE", "PASS", "Pinned to " & DEALER_MODE & "."
    End If

    '--- numbers ---
    If MIN_Q < 1 Then
        VAdd res, "MIN_Q", "FAIL", "Must be >= 1."
    ElseIf MIN_Q <> 3 Then
        VAdd res, "MIN_Q", "WARN", "Regulatory threshold is 3; current value " & MIN_Q & "."
    Else
        VAdd res, "MIN_Q", "PASS", "3 counterparties."
    End If
    If POLL_S < 1 Or TIMEOUT_S <= POLL_S Then VAdd res, "TIMEOUT_S / POLL_S", "FAIL", "Need POLL_S >= 1 and TIMEOUT_S > POLL_S." Else VAdd res, "TIMEOUT_S / POLL_S", "PASS", TIMEOUT_S & "s / " & POLL_S & "s."
    If STEP_GAP_S < 0 Then VAdd res, "STEP_GAP_S", "FAIL", "Must be >= 0." Else VAdd res, "STEP_GAP_S", "PASS", CStr(STEP_GAP_S) & "s."
    If TEST_EXPORT_RUNS < 5 Then VAdd res, "TEST_EXPORT_RUNS", "WARN", "Acceptance needs 5 consecutive successes." Else VAdd res, "TEST_EXPORT_RUNS", "PASS", TEST_EXPORT_RUNS & " cycles."
    If ParseTimeStr(RUN_AT, t) Then VAdd res, "RUN_AT", "PASS", Format$(t, "hh:nn:ss") Else VAdd res, "RUN_AT", "FAIL", "'" & RUN_AT & "' is not hh:mm[:ss]."

    '--- parsing rules ---
    If Len(DATE_FORMAT_TEXT) = 0 Then
        VAdd res, "DATE_FORMAT_TEXT", "WARN", "Empty: text dates other than ISO yyyy-mm-dd fall back to the previous business day (logged). Set it if the export has text dates."
    Else
        e = ValidateDateFormat(DATE_FORMAT_TEXT)
        If Len(e) > 0 Then
            VAdd res, "DATE_FORMAT_TEXT", "FAIL", e
        ElseIf DATE_FORMAT_TEXT Like "*[!A-Z]*" Then
            VAdd res, "DATE_FORMAT_TEXT", "PASS", DATE_FORMAT_TEXT
        Else
            VAdd res, "DATE_FORMAT_TEXT", "WARN", "'" & DATE_FORMAT_TEXT & "' has no separators: it only matches dates written like 11102026. For 11/10/2026 use MM/DD/YYYY; leave empty if the export has real date cells."
        End If
    End If
    If Len(Trim$(Replace(HDR_DEALERS, "|", ""))) = 0 Then VAdd res, "HDR_DEALERS", "FAIL", "Empty." Else VAdd res, "HDR_DEALERS", "PASS", "First alias: '" & Split(HDR_DEALERS, "|")(0) & "'."
    If Len(Trim$(Replace(HDR_RFQ, "|", ""))) = 0 Then VAdd res, "HDR_RFQ", "WARN", "Empty: deals will be keyed by ticket or composite key." Else VAdd res, "HDR_RFQ", "PASS", "First alias: '" & Split(HDR_RFQ, "|")(0) & "'."
    If GROUP_LEVEL = "ORDER" And Len(Trim$(Replace(HDR_ORDER, "|", ""))) = 0 Then VAdd res, "HDR_ORDER", "FAIL", "GROUP_LEVEL=ORDER needs HDR_ORDER."
    If Len(Trim$(Replace(EXCL_STATUS, "|", ""))) = 0 Then
        VAdd res, "EXCL_STATUS", "WARN", "Empty: every deal is treated as executed."
    ElseIf InStr(1, "|" & EXCL_STATUS & "|", "||") > 0 Then
        VAdd res, "EXCL_STATUS", "WARN", "Contains an empty token (ignored)."
    Else
        VAdd res, "EXCL_STATUS", "PASS", EXCL_STATUS
    End If

    '--- steps / terminal ---
    e = ValidateSteps(OPEN_SEQ)
    If Len(e) > 0 Then VAdd res, "OPEN_SEQ", "FAIL", e Else VAdd res, "OPEN_SEQ", "PASS", "Syntax OK."
    If SOURCE_MODE = "TERMINAL" Then
        If Len(EXPORT_SEQ) = 0 Then
            VAdd res, "EXPORT_SEQ", "FAIL", "Empty: record it with Capture Click Point or numbered items, then Test Export 5/5."
        Else
            e = ValidateSteps(EXPORT_SEQ)
            If Len(e) > 0 Then
                VAdd res, "EXPORT_SEQ", "FAIL", e
            ElseIf InStr(1, EXPORT_SEQ, "CLICK:", vbTextCompare) > 0 And Len(BBG_WINDOW) = 0 Then
                VAdd res, "EXPORT_SEQ", "WARN", "Contains clicks but BBG_WINDOW is empty: record it with Record Export so the panel size is checked before clicking."
            Else
                VAdd res, "EXPORT_SEQ", "PASS", "Syntax OK (confirm with Test Export 5/5)."
            End If
        End If
        e = PanelTitleMismatch(BBG_PANEL, BBG_TITLE_PREFIX)
        If Len(e) > 0 Then VAdd res, "BBG_PANEL / TITLE", "FAIL", e Else VAdd res, "BBG_PANEL / TITLE", "PASS", BBG_PANEL & " <-> " & BBG_TITLE_PREFIX
        okDde = DdeAvailable(msg)
        VAdd res, "DDE link", IIf(okDde, "PASS", "FAIL"), msg
        If BbgHwnd() = 0 Then
            s = WindowCandidates(True)
            VAdd res, "Bloomberg window", "FAIL", "No visible window title starts with '" & BBG_TITLE_PREFIX & "'. " & _
                IIf(Len(s) > 0, "Bloomberg-like windows: " & s & "-> put the start of the right one in BBG_TITLE_PREFIX (or use Record Export).", "No Bloomberg-like window found (click List Windows).")
        Else
            VAdd res, "Bloomberg window", "PASS", "Found."
        End If
    End If

    '--- folders ---
    If SOURCE_MODE = "FOLDER" And Len(EXPORT_FOLDER) = 0 Then
        VAdd res, "EXPORT_FOLDER", "FAIL", "Required in FOLDER mode."
    ElseIf Len(EXPORT_FOLDER) > 0 Then
        e = FolderProblem(EXPORT_FOLDER)
        If Len(e) > 0 Then VAdd res, "EXPORT_FOLDER", "FAIL", e Else VAdd res, "EXPORT_FOLDER", "PASS", EXPORT_FOLDER
    Else
        VAdd res, "EXPORT_FOLDER", "PASS", "Not used."
    End If
    If Len(ARCHIVE_DIR) = 0 Then
        VAdd res, "ARCHIVE_DIR", "FAIL", "Required: dated copy of every raw export (regulatory evidence)."
    Else
        EnsureFolder ARCHIVE_DIR
        e = FolderProblem(ARCHIVE_DIR)
        If Len(e) > 0 Then VAdd res, "ARCHIVE_DIR", "FAIL", e Else VAdd res, "ARCHIVE_DIR", "PASS", ARCHIVE_DIR
    End If

    '--- outlook ---
    If SOURCE_MODE = "OUTLOOK" Or Len(ALERT_TO) > 0 Then
        On Error Resume Next
        Set ol = Nothing: Set ol = CreateObject("Outlook.Application")
        On Error GoTo EH
        If ol Is Nothing Then
            VAdd res, "Outlook", "FAIL", "Outlook.Application not available (needed for OUTLOOK mode / ALERT_TO)."
        Else
            VAdd res, "Outlook", "PASS", "Available."
            If SOURCE_MODE = "OUTLOOK" Then
                Set fld = OutlookFolder(ol, OUTLOOK_FOLDER)
                If fld Is Nothing Then VAdd res, "OUTLOOK_FOLDER", "FAIL", "Folder '" & OUTLOOK_FOLDER & "' not found under Inbox." Else VAdd res, "OUTLOOK_FOLDER", "PASS", fld.Name
                If Len(OUTLOOK_SUBJECT) = 0 Then VAdd res, "OUTLOOK_SUBJECT", "FAIL", "Empty." Else VAdd res, "OUTLOOK_SUBJECT", "PASS", OUTLOOK_SUBJECT
            End If
        End If
    End If
    If Len(ALERT_TO) = 0 Then
        VAdd res, "ALERT_TO", "WARN", "No e-mail alerts (only LOG / popups)."
    Else
        e = BadEmails(ALERT_TO)
        If Len(e) > 0 Then VAdd res, "ALERT_TO", "FAIL", "Invalid: " & e Else VAdd res, "ALERT_TO", "PASS", ALERT_TO
    End If
    If SHOW_POPUPS Then VAdd res, "SHOW_POPUPS", "WARN", "TRUE: an alert popup blocks Excel until dismissed. Set FALSE for unattended runs." Else VAdd res, "SHOW_POPUPS", "PASS", "FALSE (unattended)."

    '--- holidays ---
    If ResolveRange(HOLIDAYS_REF) Is Nothing Then
        VAdd res, "HOLIDAYS", "FAIL", "Range '" & HOLIDAYS_REF & "' not found."
    Else
        For Each h In HolidaySet.Keys
            If Year(CDate(h)) = Year(Date) Then cyH = cyH + 1
            If Year(CDate(h)) = Year(Date) + 1 Then nyH = nyH + 1
        Next
        If cyH = 0 Then
            VAdd res, "HOLIDAYS", "WARN", "No holiday in " & Year(Date) & " (" & HolidaySet.Count & " dates loaded)."
        ElseIf nyH = 0 And Month(Date) >= 11 Then
            VAdd res, "HOLIDAYS", "WARN", "Add the " & (Year(Date) + 1) & " holidays."
        Else
            VAdd res, "HOLIDAYS", "PASS", HolidaySet.Count & " dates; previous business day = " & Format$(PrevBusinessDay(Date), "yyyy-mm-dd") & " (DAYS_BACK=" & DaysBack(Date) & ")."
        End If
    End If

    '--- dealer map ---
    dup = DealerMapConflicts(nMap)
    If Len(dup) > 0 Then
        VAdd res, "DEALER_MAP", "FAIL", "Alias mapped to different entities: " & dup
    ElseIf nMap = 0 Then
        VAdd res, "DEALER_MAP", "WARN", "Empty: aliases like 'JPM' / 'JPM LDN' / 'JPMS' would count as different counterparties."
    Else
        VAdd res, "DEALER_MAP", "PASS", nMap & " aliases."
    End If

    '--- workbook ---
    If ThisWorkbook.FileFormat <> 52 Then VAdd res, "Workbook format", "WARN", "Not saved as .xlsm." Else VAdd res, "Workbook format", "PASS", ThisWorkbook.FullName
    s = LastSelfTestResult()
    If Left$(s, 4) = "PASS" Then VAdd res, "Self-tests", "PASS", s Else VAdd res, "Self-tests", "WARN", IIf(Len(s) = 0, "Not run yet (button Run Self-Tests).", s)

    WriteValidation res, nFail, nWarn
    ValidateConfigCore = (nFail = 0)
    WriteLog "Validate Config: " & nFail & " FAIL, " & nWarn & " WARN.", IIf(nFail = 0, "INFO", "WARN")
    Exit Function
EH:
    VAdd res, "Validation", "FAIL", "Error: " & Err.Description
    WriteValidation res, nFail, nWarn
    ValidateConfigCore = False
End Function

Private Sub VAdd(res As Collection, ByVal chk As String, ByVal lvl As String, ByVal detail As String)
    res.Add Array(chk, lvl, detail)
End Sub

Private Sub VEnum(res As Collection, ByVal k As String, ByVal v As String, ByVal allowed As String)
    If InStr(1, "|" & allowed & "|", "|" & v & "|") = 0 Or Len(v) = 0 Then
        VAdd res, k, "FAIL", "'" & v & "' is not " & allowed & "."
    Else
        VAdd res, k, "PASS", v
    End If
End Sub

Private Sub WriteValidation(res As Collection, ByRef nFail As Long, ByRef nWarn As Long)
    Dim ws As Worksheet, r As Long, it As Variant, c As Long
    Set ws = GetSheet(SH_CONFIG)
    ws.Range(ws.Cells(CFG_HDR_ROW - 1, CFG_COL_CHK), ws.Cells(200, CFG_COL_CHK + 2)).Clear
    ws.Cells(CFG_HDR_ROW - 1, CFG_COL_CHK).Value = "Validation " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(CFG_HDR_ROW - 1, CFG_COL_CHK).Font.Italic = True
    ws.Cells(CFG_HDR_ROW, CFG_COL_CHK).Resize(1, 3).Value = Array("Check", "Result", "Detail")
    ws.Cells(CFG_HDR_ROW, CFG_COL_CHK).Resize(1, 3).Font.Bold = True
    ws.Cells(CFG_HDR_ROW, CFG_COL_CHK).Resize(1, 3).Interior.Color = RGB(217, 225, 242)
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
    f = p & "\~blot_write_test_" & Format$(Now, "hhnnss") & ".tmp"
    On Error GoTo NoWrite
    h = FreeFile
    Open f For Output As #h
    Print #h, "test"
    Close #h
    Kill f
    Exit Function
NoWrite:
    FolderProblem = "Folder not writable: " & p & " (" & Err.Description & ")"
End Function

Private Function BadEmails(ByVal s As String) As String
    Dim a As Variant, t As String
    For Each a In Split(Replace(s, ",", ";"), ";")
        t = Trim$(CStr(a))
        If Len(t) > 0 Then
            If Not (t Like "?*@?*.?*") Or InStr(1, t, " ") > 0 Then BadEmails = BadEmails & t & " "
        End If
    Next
End Function

Private Function PanelTitleMismatch(ByVal panel As String, ByVal title As String) As String
    Dim n As String
    If Not (LCase$(panel) Like "<blp-#>" Or LCase$(panel) Like "<blp-##>") Then
        PanelTitleMismatch = "BBG_PANEL '" & panel & "' is not <blp-N>.": Exit Function
    End If
    n = Mid$(panel, 6, Len(panel) - 6)
    If Not (title Like "#-*" Or title Like "##-*") Then Exit Function      ' title without panel number: cannot check
    If Left$(title, Len(n) + 1) <> n & "-" Then
        PanelTitleMismatch = "BBG_PANEL " & panel & " expects a title starting with '" & n & "-', got '" & title & "'."
    End If
End Function

' Returns "alias(entityA/entityB) ..." for aliases mapped to two different entities
Private Function DealerMapConflicts(ByRef nMap As Long) As String
    Dim ws As Worksheet, mv As Variant, lr As Long, i As Long, d As Object, a As String, e As String
    Set d = NewDict()
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(SH_MAP): On Error GoTo 0
    If ws Is Nothing Then Exit Function
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr < 2 Then Exit Function
    mv = ws.Range("A1:B" & lr).Value
    For i = 2 To UBound(mv, 1)
        If Not IsError(mv(i, 1)) And Not IsError(mv(i, 2)) Then
            a = NormName(CStr(mv(i, 1))): e = NormName(CStr(mv(i, 2)))
            If Len(a) > 0 And Len(e) > 0 Then
                nMap = nMap + 1
                If d.Exists(a) Then
                    If d(a) <> e Then DealerMapConflicts = DealerMapConflicts & a & "(" & d(a) & "/" & e & ") "
                Else
                    d(a) = e
                End If
            End If
        End If
    Next
End Function

Private Function LastSelfTestResult() As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_TESTS)
    If Not ws Is Nothing Then LastSelfTestResult = CStr(ws.Range("B2").Value)
End Function

Public Function OutlookFolder(ol As Object, ByVal path As String) As Object
    Dim fld As Object, p As Variant
    On Error GoTo Fail
    Set fld = ol.GetNamespace("MAPI").GetDefaultFolder(6)
    For Each p In Split(Replace(path, "/", "\"), "\")
        If Len(Trim$(CStr(p))) > 0 Then Set fld = fld.Folders(Trim$(CStr(p)))
    Next
    Set OutlookFolder = fld
    Exit Function
Fail:
    Set OutlookFolder = Nothing
End Function

'==========================================================================================
'                                    SMALL PARSERS (locale-free)
'==========================================================================================
' "hh:mm" or "hh:mm:ss"
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

' -?digits(.digits)?   (no thousands separators, '.' decimal)
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

' Locale-independent number to text ('.' decimal, no grouping)
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
