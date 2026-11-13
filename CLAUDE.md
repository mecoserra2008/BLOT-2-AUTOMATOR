# BLOT Best-Execution Monitor — Handoff for Claude Code

> Save this file as `CLAUDE.md` in the repo root, next to `src/modBlotBestEx.bas`.
> Claude Code loads `CLAUDE.md` automatically. This file is the specification and the source of truth for intent.
> `src/modBlotBestEx.bas` (v3) is the current code baseline.

---

## 0. Mission

Deliver **`dist/BLOT_BestEx.xlsm`**: a single, ready-to-use Excel workbook that runs **unattended** every business day and does the following:

1. Drives the Bloomberg Terminal (`BLOT 2 <GO>` → `BLOT -1 <GO>`), which shows the previous day's executed trades.
2. Exports the BLOT grid **with no human click**, or ingests a Bloomberg-delivered report (email or file drop).
3. Identifies each trade and counts the **distinct counterparties asked for a quote** on that trade.
4. Flags every deal with **fewer than 3 counterparties asked**. This is a Banco de Portugal requirement: the desk must request quotes from ≥ 3 counterparties.
5. Keeps an **auditable history** (per deal and per day), archives the raw export, logs every run, and emails alerts.

**Definition of "completely ready":** the user opens the workbook, fills a few cells on a `CONFIG` sheet (paths, email, export steps), clicks **Validate Config**, then **Run Now**. Nothing in the VBA editor needs to be touched.

The user is a quant trader at BiG (Lisbon). The machine is Windows with Excel 365 and a logged-in Bloomberg Terminal. Excel may be Portuguese-localized (see §9).

---

## 1. Requirements (from the user)

| # | Requirement |
|---|---|
| R1 | Open the Bloomberg Terminal (already running), type `BLOT 2` + GO, then `BLOT -1` + GO. This shows yesterday's executed trades. |
| R2 | Export that grid automatically. **100% automated; no manual step.** |
| R3 | Dashboard: (1) No. of trades; (2) No. of counterparties asked for a specific deal (lookup); (3) % of trades with < 3 counterparties asked. |
| R4 | Distribution of the number of counterparties asked per deal, plus an **alert** if any deal has < 3. |
| R5 | History is of **extreme importance** (regulatory evidence). |
| R6 | The count must be **per trade** (same trade), not per security/symbol. |

---

## 2. History of the work so far

### v1: first version (superseded)
- Opened the BLOT via **DDE** (`DDEInitiate("WINBLP","BBK")` + `DDEExecute "<blp-1>BLOT 2<GO>"`), with SendKeys as fallback.
- Export was **manual by design**: `AUTO_EXPORT_KEYS` was empty. The macro polled with `Application.OnTime` (non-blocking) for a new workbook in the *same* Excel instance.
- It grouped deals by the first column matching "trade id/ticket…". If none matched, each row became its own trade (`"ROW"&i`).

**Field result:** the Terminal opened but **no export happened**. Two root causes:
1. Nothing ever triggered the export.
2. Bloomberg's "Export to Excel" often opens in a **separate Excel process**, which v1 could not see.

### v2: trade-identification fixes (logic review)
Problems found in v1's trade identification and counting:

| Problem | Effect |
|---|---|
| Ticket ≠ competition. The RFQ/inquiry holds the dealer competition; only the winning dealer row may carry a ticket. | False alerts |
| Each row treated as a trade when no ID column exists | False alerts |
| Status filter applied to *dealer* rows. A dealer row marked Expired/Passed means the dealer **was asked** but those rows were dropped. | False alerts |
| No normalization of dealer names (`JPM`, `JPM LDN`, `JPMS` counted as 3) | **False compliance (dangerous direction)** |
| Tree-style exports leave the parent ID/ISIN blank on child rows | False alerts |

Fixes, all present in v3:
- Key hierarchy: RFQ ID → Ticket ID → composite key.
- Forward-fill of the key on child rows.
- An integrity check on each key.
- Status evaluated at trade level.
- A `DEALER_MAP` table that maps aliases to legal entities.

### v3: full automation of the export (current baseline: `src/modBlotBestEx.bas`)
- A **step engine** drives the Terminal. Tokens: `KEYS:` (DDE), `SEND:` (SendKeys), `CLICK:/RCLICK:/DCLICK:x,y` (Win32 mouse, coordinates relative to the Bloomberg window), `WAIT:s`.
- **All running Excel instances are enumerated** through `AccessibleObjectFromWindow` (XLMAIN → XLDESK → EXCEL7), so an export that opens in another process is detected.
- `SOURCE_MODE` takes one of three values:
  - `TERMINAL`: keys/clicks.
  - `OUTLOOK`: today's email attachment, matched by subject.
  - `FOLDER`: newest file in a drop folder.
- A stability check waits until the export's row count is unchanged across 2 polls before parsing.
- `LOG` sheet, Outlook email alerts, popups that can be turned off.
- Calibration macros: `TestDDE`, `TestExport`, `CaptureClickPoint` (records `CLICK:x,y`), `ImportActiveWorkbook`, `ImportFromFile`.
- Daily scheduler via `OnTime` (`ScheduleDaily` / `ScheduledRun` / `CancelSchedule`), hooked from `ThisWorkbook`.

**What is still missing:** see §7. In short: configuration lives in VBA `Const`s instead of a sheet, there is no build script, no tests, and several robustness gaps remain.

---

## 3. Architecture

```
                  +------------------- SOURCE_MODE --------------------+
                  |                                                     |
 TERMINAL:  OPEN_SEQ (DDE) -> EXPORT_SEQ (keys/clicks) -> new workbook in ANY Excel instance
 OUTLOOK :  today's mail matching OUTLOOK_SUBJECT -> attachment saved to TEMP -> open
 FOLDER  :  newest *.xls*/*.csv in EXPORT_FOLDER since today 00:00 -> open
                  |
                  v
        PollExport (OnTime every POLL_S, until TIMEOUT_S)
                  |
                  v
        FindDataBlock: header row (first 40 rows) containing a dealer alias -> data block
                  |
                  v
        ProcessExport:
          header mapping (aliases) -> DetectMode -> per-row TradeKey -> live/exclusion
          -> integrity (ISIN, side) -> distinct normalized dealers per key
                  |
                  v
   RAW (raw copy) | DEALS (today, per deal) | HIST (upsert per date+key) | HIST_DAILY (KPIs)
                  |
                  v
        DASHBOARD (KPIs, distribution + chart, deal lookup, exceptions) + LOG + email/popup
```

### Sheets

| Sheet | Content |
|---|---|
| `CONFIG` | **(to build)** all inputs; the only place the user edits |
| `DASHBOARD` | KPIs, status banner, distribution table + chart, per-deal lookup (dropdown), exception list |
| `DEALS` | the latest run, one row per deal, sorted by NDealers ascending |
| `RAW` | raw export of the latest run |
| `HIST` | audit trail, one row per (TradeDate, TradeKey); upsert on re-run |
| `HIST_DAILY` | one row per trade date: NTrades, NBelowMin, Pct, Avg/Min/Max dealers |
| `DEALER_MAP` | alias as exported → legal entity |
| `LOG` | timestamp, level, message for every run, alert, timeout and error |

Column schema of `DEALS` and `HIST`: `TradeDate | TradeKey | Security | Side | Qty | NDealers | Status | LoadedAt`.
- `TradeKey` keeps its prefix (`RFQ:`, `TKT:` or `CMP:`) so that it is always visible how the deal was grouped.
- `Status` is one of `OK`, `< 3` or `KEY?`.

---

## 4. Trade identification: the core compliance logic (do not weaken)

### 4.1 Unit of analysis
One **competition event** = one RFQ/inquiry. The metric for each deal is the **number of distinct legal-entity counterparties asked** on that RFQ.

> **Policy question still open for compliance:** if one order is worked through several RFQs (e.g. RFQ to 2 dealers, no fill, then RFQ to 2 others), is the unit the RFQ or the order? Default: **RFQ**.
> To switch to order level, put the order-ID header first in `HDR_RFQ`. Make this a CONFIG option `GROUP_LEVEL = RFQ | ORDER`.

### 4.2 Key hierarchy (`TradeKey`)
1. `RFQ:<inquiry id>`: preferred.
2. `TKT:<ticket/trade id>`: valid only if every dealer row carries it.
3. `CMP:<TradeDate|ISIN|Side|Qty|RFQ send time|Trader>`: fallback. Each field separates a different kind of collision:

   | Field | Separates |
   |---|---|
   | ISIN | different bonds on the same day |
   | Side, Qty | two trades in the same bond on the same day |
   | RFQ **send** time | re-trades of the same size. It must be the inquiry timestamp, which is common to all dealer rows. A quote timestamp would split the group. |
   | Trader | two desks working the same bond |

### 4.3 Row rules
- **Forward-fill.** A row with a dealer but no RFQ, ticket or ISIN inherits the previous key (tree exports).
- **Liveness.** A key is kept if **any** of its rows has a status that does not match `EXCL_STATUS`.
  - Dealer-level `Expired` / `Passed` / `Rejected` rows still count as **asked**.
  - A deal is dropped only if all its rows are cancelled or not executed.
- **Integrity.** Within a key, the non-blank ISIN and Side values must be unique. Otherwise the deal gets `Status = KEY?` and is listed as an exception.
- **Normalization.** `UCase(Trim(collapse spaces))`, then looked up in `DEALER_MAP`.

### 4.4 Counting modes (`DEALER_MODE`)

| Mode | Export shape | Count |
|---|---|---|
| `ROWS` (preferred; auditable) | one row per dealer, same RFQ ID repeated | distinct normalized dealer names per key |
| `LIST` | one row per deal, dealers delimited (`;,/\|` or newline) | distinct tokens, unioned across rows of the key |
| `COUNT` | Bloomberg numeric field (e.g. "Dealers in Comp") | max over rows of the key; **not auditable** |
| `AUTO` | detects the mode | numeric column → COUNT; repeated key → ROWS; otherwise → LIST |

Once the real export layout is known, `AUTO` must be **pinned** to a fixed mode in CONFIG.

### 4.5 Field caveat
The dealer field must measure **counterparties asked**, not **quotes received**. Using quotes received biases the metric downward and creates false alerts. Trades done by voice or outside Bloomberg never appear in the BLOT, so this control does not cover them.

---

## 5. Export automation: options, in order of preference

| Option | Mechanism | Pros | Cons |
|---|---|---|---|
| **A. Bloomberg-side delivery** (`OUTLOOK` / `FOLDER`) | Ask HELP HELP / the account rep whether the BLOT view (with the RFQ ID and dealers-asked columns) can be scheduled by email or SFTP/share drop. **Availability is unconfirmed.** | No GUI; works on a locked screen; survives UI changes | Depends on Bloomberg offering it |
| **B. Numbered toolbar items via DDE** (`KEYS:96<GO>\|KEYS:n<GO>`) | Many Bloomberg screens show `NN) Label` toolbar/menu items; typing `NN<GO>` triggers them | Independent of window position | Depends on the export item being numbered |
| **C. Recorded clicks** (`CLICK:x,y`) | `CaptureClickPoint` records coordinates relative to the Bloomberg window | Always possible | Panel must be maximized on the same monitor at the same size; screen unlocked; no human input during the run |

`TestExport` must succeed **5 out of 5** consecutive times before the step sequence is accepted.

---

## 6. Inputs: become the `CONFIG` sheet

| Key | Default | Notes |
|---|---|---|
| `SOURCE_MODE` | `TERMINAL` | `TERMINAL` \| `OUTLOOK` \| `FOLDER` |
| `BBG_TITLE_PREFIX` | `1-BLOOMBERG` | exact panel window title |
| `BBG_PANEL` | `<blp-1>` | must match the title (`1-` ↔ `<blp-1>`) |
| `OPEN_SEQ` | `KEYS:BLOT 2<GO>\|WAIT:3\|KEYS:BLOT -1<GO>\|WAIT:5` | the user's manual keystrokes |
| `EXPORT_SEQ` | *(empty, must be filled)* | from §5 B or C |
| `STEP_GAP_S` | `0.7` | delay between steps |
| `EXPORT_FOLDER` | *(path)* | FOLDER mode, or TERMINAL when the export saves to a file |
| `ARCHIVE_DIR` | *(path)* | **new.** Dated copy of every raw export (regulatory evidence) |
| `OUTLOOK_SUBJECT` / `OUTLOOK_FOLDER` | `BLOT` / *(empty = Inbox)* | OUTLOOK mode |
| `CLOSE_EXPORT` | `TRUE` | close the export workbook; quit its Excel instance if it is empty |
| `TIMEOUT_S` / `POLL_S` | `300` / `3` | |
| `RUN_AT` | `08:30:00` | daily run time |
| `HOLIDAYS` | *(range)* | **new.** PT/TARGET holidays: skip those runs and use them in `WORKDAY` |
| `ALERT_TO` | *(emails, `;`-separated)* | Outlook email on alert, timeout or error |
| `SHOW_POPUPS` | `TRUE` | set `FALSE` for unattended runs |
| `MIN_Q` | `3` | regulatory threshold |
| `DEALER_MODE` | `AUTO` | pin it after calibration |
| `GROUP_LEVEL` | `RFQ` | **new.** `RFQ` \| `ORDER` |
| `LIST_DELIMS` | `;,/\|` | |
| `HDR_DEALERS`, `HDR_RFQ`, `HDR_ID`, `HDR_SEC`, `HDR_DATE`, `HDR_SIDE`, `HDR_QTY`, `HDR_STATUS`, `HDR_TIME`, `HDR_TRADER` | alias lists (see `.bas`) | exact export headers first, lowercase, `\|`-separated |
| `EXCL_STATUS` | `cancel\|cxl\|reject\|expired` | trade-level statuses that mean "not executed" |
| `DATE_FORMAT_TEXT` | *(empty)* | **new.** Explicit format for text dates (e.g. `MM/DD/YYYY`) to avoid PT locale misparsing |

**Inputs the user must supply** (Claude Code: ask for them; never guess):
1. Header row of one real BLOT export (plus 2–3 anonymized rows).
2. The exact status values in that export.
3. `EXPORT_SEQ`: numbered items or captured clicks.
4. The list of dealer aliases for `DEALER_MAP`.
5. The paths (`EXPORT_FOLDER`, `ARCHIVE_DIR`), `ALERT_TO`, and the answer on `GROUP_LEVEL`.

---

## 7. Work plan for Claude Code

### 7.1 Repository layout
```
blot-bestex/
  CLAUDE.md                     <- this file
  src/
    modBlotBestEx.bas           <- v3 baseline (refactor here)
    modConfig.bas               <- CONFIG loader/validator (new)
    modSelfTest.bas             <- fixture-driven tests (new)
    ThisWorkbook.vba            <- event code, injected by the build (new)
  build/
    build.ps1                   <- creates dist/BLOT_BestEx.xlsm from src/ (new)
  tests/
    fixtures/*.csv              <- synthetic exports (new, §8)
    expected/*.json             <- expected results per fixture (new)
    reference_impl.py           <- pandas re-implementation of §4 for cross-checking (new)
  dist/
    BLOT_BestEx.xlsm            <- final deliverable
```

### 7.2 Tasks (in order)
1. **Config to a sheet.**
   - Replace every config `Const` with module-level variables filled by `LoadConfig()`, called at the start of every public entry point.
   - Read the `CONFIG` sheet (columns `Key | Value | Description`) via `Cfg(key, default)`.
   - `ValidateConfig()` checks: required keys present; paths exist and are writable; `EXPORT_SEQ` is non-empty in TERMINAL mode; `BBG_PANEL` is consistent with the title; the DDE option is enabled (try `DDEInitiate`). It writes a pass/fail list to `CONFIG`.
2. **Robustness gaps in v3 to fix:**
   - **Dates.** Parse from `Value2` when numeric. Parse text dates with `DATE_FORMAT_TEXT`. Never round-trip through a locale string.
   - **Multi-day exports.** `HIST_DAILY` must aggregate **per trade date**; v3 uses the max date only.
   - **CSV opening.** Use `Workbooks.Open(..., Local:=True)` for `.csv` (PT locale `;` separator).
   - **Archive.** Save the raw export to `ARCHIVE_DIR\BLOT_<tradedate>_<yyyymmdd_hhnnss>.xlsx` before parsing.
   - **Holidays.** Skip non-business days in `ScheduledRun`; use `HOLIDAYS` in `WORKDAY`.
   - **Re-entrancy.** Prevent overlapping runs. Add `AbortRun` to reset `gBusy` and cancel pending `PollExport`.
   - **`GROUP_LEVEL = ORDER`** support.
   - **Mode pinning.** When `DEALER_MODE` is `AUTO`, log the detected mode and the mapped column for every header (audit).
   - **Data-quality warning.** Warn if no RFQ/ticket column is found, i.e. composite keys are in use.
3. **Dashboard buttons** (form controls assigned to macros): `Run Now`, `Validate Config`, `Test Export`, `Capture Click Point`, `Import File`, `Run Self-Tests`, `Abort`.
4. **Self-tests** (`modSelfTest.bas`):
   - `RunSelfTests` loads every fixture from `tests/fixtures`, or from a hidden `FIXTURES` sheet embedded at build time.
   - It runs the parsing core (no Terminal) and asserts the results against `expected`.
   - It writes results to a `TESTS` sheet and returns a pass flag, with **no MsgBox**, so it can be run headless from the build.
   - Refactor `ProcessExport` into a **pure core** (array in, deals array out) plus I/O wrappers, so tests never touch `RAW`/`HIST`.
5. **Reference implementation** (`tests/reference_impl.py`, pandas): the same §4 logic, run on the same fixtures. Results must match the VBA output exactly. This is independent verification of the compliance logic.
6. **Build** (`build/build.ps1`, see §9): produce `dist/BLOT_BestEx.xlsm`, run `RunSelfTests` headless, and fail the build if any test fails.
7. **Documentation:** fill a `README` sheet in the workbook (1-page operator guide: what to fill, how to calibrate, what alerts mean).

### 7.3 Coding constraints (VBA)
- `Option Explicit` in every module. **Late binding only**: no references to set (Scripting, Outlook).
- Win32 `Declare PtrSafe` with `LongPtr`, which works on 32- and 64-bit Office.
- **ASCII-only source, CRLF line endings.** The VBA import uses the ANSI code page and corrupts UTF-8.
- English UI strings in code, to avoid encoding issues in the PT environment.
- Never use `localStorage`-style hacks. All state goes in sheets.
- Keep `HIST` append/upsert **idempotent**: re-running the same day must not duplicate rows.
- Never silently drop a deal. Anything ambiguous becomes `KEY?` and is reported.

---

## 8. Test fixtures to create (`tests/fixtures`)

| Fixture | Shape | Content | Expected |
|---|---|---|---|
| `f1_rows_flat.csv` | ROWS, RFQ ID on every row | RFQ-001 bond A Buy: JPM (Done), BNPP (Expired), SG (Passed). RFQ-002 bond B Sell: `JPM LDN`, `JPMS`, CITI (map both JPM aliases → JPMORGAN). RFQ-003: all rows `Cancelled`. RFQ-004: rows with **two different ISINs**, 3 dealers. | live = {001, 002, 004}. 001 → 3 OK. 002 → 2 `< 3`. 004 → `KEY?`. nB = 1, pct = 33.3% |
| `f2_rows_tree.csv` | ROWS, child rows with blank RFQ/ISIN | 2 parents, 4 and 2 child dealer rows | 4 OK, 2 `< 3` (forward-fill works) |
| `f3_count.csv` | COUNT | numeric "Dealers in Comp" 5, 2, 3; one trade split over 2 allocation rows (3, 3) | max per key; one `< 3` |
| `f4_list.csv` | LIST | `"JPM; BNPP / SG"`, `"JPM;JPM LDN"` (map), newline-separated cell | 3, 1 (alias collapse), n |
| `f5_composite.csv` | no IDs | same ISIN, same side, same qty, same day, **different RFQ send times**, plus one different trader | 3 distinct deals |
| `f6_pt_locale.csv` | `;`-separated, text dates `11/10/2026` (US) | with `DATE_FORMAT_TEXT = MM/DD/YYYY` | dates parsed as 10 Nov 2026 |
| `f7_rerun.csv` | same as f1, loaded twice | | HIST row count unchanged on the 2nd load |

`DEALER_MAP` fixture: `JPM LDN → JPMORGAN`, `JPMS → JPMORGAN`, `JPM → JPMORGAN`.

---

## 9. Build notes (Windows)

- Building requires Excel installed and *Trust Center → Macro Settings → **Trust access to the VBA project object model*** enabled.
  - Registry equivalent: `HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1`.
  - Corporate GPO may block this. If so, fall back to the **manual assembly steps** below and say so.
- **Localization gotcha:** in PT Excel the `ThisWorkbook` component is not named "ThisWorkbook". Always resolve it via `$wb.CodeName`.

Skeleton:
```powershell
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$out  = Join-Path $root 'dist\BLOT_BestEx.xlsm'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb  = $xl.Workbooks.Add()
  $vbp = $wb.VBProject
  foreach ($m in 'modBlotBestEx.bas','modConfig.bas','modSelfTest.bas') {
    $vbp.VBComponents.Import((Join-Path $root "src\$m")) | Out-Null
  }
  $tw = $vbp.VBComponents.Item($wb.CodeName)                 # localized-safe
  $tw.CodeModule.AddFromString((Get-Content (Join-Path $root 'src\ThisWorkbook.vba') -Raw))
  $xl.Run("'$($wb.Name)'!Setup")                              # sheets, CONFIG defaults, buttons, README
  $xl.Run("'$($wb.Name)'!EmbedFixtures", (Join-Path $root 'tests\fixtures'))
  $ok = $xl.Run("'$($wb.Name)'!RunSelfTestsHeadless")
  if (-not $ok) { throw 'Self-tests failed - see TESTS sheet' }
  if (Test-Path $out) { Remove-Item $out }
  $wb.SaveAs($out, 52)                                        # 52 = xlOpenXMLWorkbookMacroEnabled
} finally { if ($wb) { $wb.Close($false) }; $xl.Quit() }
```

**Manual fallback** if VBOM access is blocked:
1. New workbook, then Alt+F11.
2. Import the three `.bas` files.
3. Paste `ThisWorkbook.vba` into the ThisWorkbook module.
4. Run `Setup`, then `RunSelfTests`.
5. Save as `.xlsm`.

`ThisWorkbook.vba`:
```vba
Private Sub Workbook_Open()
    ScheduleDaily
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    CancelSchedule
End Sub
```

---

## 10. Environment prerequisites (operator machine)
1. Excel → Options → Advanced → General: **uncheck** "Ignore other applications that use Dynamic Data Exchange (DDE)".
2. Place the workbook in a **Trusted Location**, or have IT sign it.
3. Bloomberg Terminal logged in. For TERMINAL mode: Windows session **unlocked** at `RUN_AT`, panel maximized (CLICK mode), and no keyboard or mouse use during the roughly 15 s run.
4. Outlook open for OUTLOOK mode or `ALERT_TO`. Programmatic-access prompts must be allowed by IT.
5. The workbook stays open for the `OnTime` scheduler, or Windows Task Scheduler opens it before `RUN_AT`.

---

## 11. Acceptance criteria (definition of done)
- [ ] `build.ps1` produces `dist/BLOT_BestEx.xlsm` with no compile errors; all self-tests pass headless.
- [ ] `tests/reference_impl.py` reproduces the VBA results on every fixture.
- [ ] The user edits **only** `CONFIG` (and `DEALER_MAP`). `Validate Config` reports all green.
- [ ] `TestExport` succeeds 5/5 on the real Terminal.
- [ ] A scheduled run completes unattended:
  - `RAW`, `DEALS`, `HIST` and `HIST_DAILY` are updated;
  - the archive file is written;
  - a `LOG` entry is written;
  - an email is sent if any deal has < 3 counterparties or `KEY?`.
- [ ] Re-running the same day is idempotent.
- [ ] 5–10 real deals are checked by hand against the BLOT inquiry detail and match `NDealers`.

---

## 12. Open questions to raise with the user
1. Real export header row, status values, and whether the layout is ROWS, LIST or COUNT.
2. Does Bloomberg offer scheduled delivery of this BLOT view (option A)?
3. `GROUP_LEVEL`: RFQ or order, per compliance.
4. Are voice/off-platform trades in scope? If yes, they need a separate data source.
5. Holiday calendar to use (PT, TARGET2, or both).

---

## 13. Implementation status (2026-11-11, v4)

**Built:** `build/build.ps1` -> `dist/BLOT_BestEx.xlsm`. 249 VBA self-tests pass headless on 9 fixtures (12 cases), and `tests/reference_impl.py` matches the VBA deal by deal (`tests/vba_out/`). An end-to-end FOLDER-mode run (CSV and typed .xlsx) was verified on a scratch copy: archive, RAW/DEALS/HIST/HIST_DAILY/DASHBOARD/LOG, re-run idempotence, Validate Config, clean failure of TERMINAL mode without Bloomberg, scheduler.

**Build notes:** `build.ps1` enables `AccessVBOM` for the build only and restores the previous value after the build's own Excel instance has exited (Excel writes its trust settings back on exit). Options: `-SkipReference`, `-KeepVbomSetting`. In PT Excel the ThisWorkbook component is `EsteLivro` (resolved via `$wb.CodeName`). VBA limit: at most 24 line continuations per statement.

**Deliberate deviations from the spec (all stricter or additive):**
- CSV is parsed as **text** by VBA (`ParseCsvText`), not `Workbooks.Open(Local:=True)`: Excel converts text dates using the PT locale on open (`11/10/2026` -> 11 Oct), so `DATE_FORMAT_TEXT` could never apply. Delimiter: `CSV_DELIM` = AUTO|;|,|TAB.
- `DEALS`/`HIST` have 2 extra columns after the 8 specified ones: `Dealers` (normalized names, audit) and `Source` (archive file). `HIST_DAILY` has `NKeyIssues` in column I and is **recomputed per date from HIST**.
- `HDR_SEC` lists `isin` first (integrity + composite key use the ISIN).
- TKT rule enforced: if the ticket column is blank on some dealer rows, tickets are not used as keys (composite key instead). Fixture f9.
- Keys `ORD:` for `GROUP_LEVEL=ORDER` (`HDR_ORDER`). Fixture f8 shows RFQ vs ORDER.
- AUTO mode = COUNT if numeric majority; ROWS if any key (after forward-fill) has >= 2 dealer rows; else LIST.
- `{DAYS_BACK}` placeholder in step sequences. **Coverage-gap alert**: business days between the last loaded trade date and the new one with no data. Reason: if `BLOT -1` means one *calendar* day back, Monday runs would miss Friday (and post-holiday runs the day before the holiday).
- A run whose export has no executed trades raises an ALERT and archives the export, but writes no 0-trade HIST_DAILY row (its trade date cannot be known for sure).
- Holidays: table on CONFIG column I, prefilled (current year + 3) with PT national + Lisbon (13 Jun) + TARGET2 days. Edit to the policy chosen.
- Extra CONFIG keys: `CSV_DELIM`, `TEST_EXPORT_RUNS`, `SEND_OK_EMAIL`, `HDR_ORDER`.

**Still needs the operator / real Terminal (cannot be done offline):**
1. Paths: `ARCHIVE_DIR` (required), `EXPORT_FOLDER` (if used); `ALERT_TO`; `SHOW_POPUPS=FALSE`.
2. `EXPORT_SEQ` (numbered items or Capture Click Point), then Test Export 5/5.
3. The real header row: check the LOG "Header mapping" line; put exact headers first in `HDR_*`, the real status values in `EXCL_STATUS`, pin `DEALER_MODE`, and set `DATE_FORMAT_TEXT` if dates are text.
4. `DEALER_MAP` aliases (unmapped names are logged after every run).
5. Compliance answers: `GROUP_LEVEL`, holiday calendar, `BLOT -1` semantics (calendar vs business day), voice trades.
6. Acceptance: hand-check 5-10 real deals vs the BLOT inquiry detail.

---

## 14. v5 (2026-11-13): lean rebuild on the real BLOT export (supersedes §3-§4, §6, §8 where they differ)

**Real export layout** (one row per DEALER LEG): `Security | Side | Block | Status | Qty (M) | Price | Yield | BrkrName | Net | Curncy | Account | Trade Dt | SetDt | ISIN | Ac/c | Factor | ID`. Statuses seen so far: `Rejected`, `C/Cancel`, `Passed` (a Passed row can carry a price). The status of an executed row is not yet confirmed: `FILL_STATUS` defaults to `done|filled|accepted|traded|executed|complete`; check the "Statuses seen" line on the DASHBOARD.

**Rule (AnalyzeBlot, pure core):**
- request = rows whose `ID` is equal except the last `ID_SUFFIX_LEN` (2) characters: `FITDNA:6AF5A884011800|03|05|06`. No ID: `date|ISIN|side|qty`.
- integrity: one ISIN, one side, one trade date per request, else `KEY?`. Cross-check: if every dealer row of the same bond/side/day/qty becomes its own request, a warning says to check `ID_SUFFIX_LEN`.
- unit = request (`GROUP_LEVEL=INQUIRY`) or all requests for the same date|ISIN|side (`ORDER`; the export has no order id).
- asked = distinct dealers after `DEALER_MAP` (all statuses count: Rejected / C/Cancel / Passed dealers were asked); quoted = dealers with a price; filled = status matches `FILL_STATUS`, winner = that dealer.
- check: `BREACH` (filled, asked < MIN_Q) | `LOW` (not filled, asked < MIN_Q) | `KEY?` | `OK`. ALERT e-mail on BREACH / KEY?.

**Export automation (confirmed on the Terminal):** `OPEN_SEQ = KEYS:BLOT 2<GO>|WAIT:3|KEYS:BLOT -1<GO>|WAIT:5`, `EXPORT_SEQ = KEYS:2<GO>|WAIT:2|KEYS:2<GO>|WAIT:3|KEYS:1<GO>|WAIT:5` (Export -> XLS - Excel format -> 1) Export). Bloomberg saves the XLS into `EXPORT_FOLDER`, which is watched (this workbook, `~$` files and `BLOT_########_########_######*` archive copies are excluded). Step engine v5: KEYS (DDE), SEND, WAIT only.

**Workbook:** sheets DASHBOARD, INQUIRIES (latest units), QUOTES (latest legs), PASTE (Ctrl+V + Process Pasted Data), HIST (upsert per TradeDate+Key), HIST_DAILY (recomputed from HIST), CONFIG (24 keys), DEALER_MAP, LOG; TESTS/FIXTURES hidden. Buttons: Run Now, Process Pasted Data, Import File, Test Export, Validate Config, Abort. Removed in v5: README sheet, DEALS/RAW, Outlook/FOLDER source modes, clicks/recorder/calibration, LIST/COUNT/ROWS modes.

**Tests:** synthetic fixtures `tests/fixtures/b1..b3` in the real layout (no real trade data in the public repo); 196 VBA checks; `reference_impl.py` re-implements the rule and matches the VBA unit by unit. Site values: `config/site.local.ini` (git-ignored) -> `dist/BLOT_BestEx_SITE.xlsm`.

**One button + schedule (v5.1):** DASHBOARD button **RUN BLOT CHECK** = OPEN_SEQ -> EXPORT_SEQ -> wait for the export (new workbook in any Excel, or new XLS in EXPORT_FOLDER / workbook folder / Downloads / Documents / C:\blp\data; the folder where it is found is saved as EXPORT_FOLDER) -> analysis -> dashboard -> workbook saved. Progress shows in DASHBOARD!B2. `SCHEDULE_MODE`: `TASK` (CONFIG button **Install Daily Task** writes `BLOT_RunDaily.vbs` next to the workbook and registers the Windows task "BLOT Best Execution", weekdays at RUN_AT, StartWhenAvailable; the script opens Excel/the workbook if needed and calls `ScheduledRun`) | `WORKBOOK` (OnTime while the workbook is open) | `OFF`. `ScheduledRun` runs once per business day (hidden name `BLOT_LAST_OK`), one retry after `RETRY_MIN` if the export fails. Install Daily Task has not been exercised on the dev PC (task creation is a machine change); verify it on the Bloomberg PC.

**Export dialog (v5.2):** `KEYS:` goes through DDE into the panel command line and cannot press a button of the "Export Blotter" dialog, so the file was never saved. Steps now: `FOCUS:<title part>` (brings that window to the front, verified; falls back to the Bloomberg panel found by title prefix or by process wintrv/bplus/blp/bloomberg/bbg/bterm), `SEND:<keys>` (only after a verified focus; never into another app), `IFWIN:<title part>:<keys>` (optional prompts, e.g. overwrite). `EXPORT_SEQ` variants separated by `||` are tried in turn (30 s each, Esc on the dialog in between); the one that delivers the file moves first on CONFIG. Default: `...|FOCUS:Export Blotter|SEND:1~|WAIT:5 || ...|FOCUS:Export Blotter|SEND:~|WAIT:5`. On timeout LOG lists the window in front, windows opened during the run and the Bloomberg windows. Exports opened in Protected View are read through `ProtectedViewWindows`. Keystrokes need an unlocked Windows session.
