# Excel → STRAFE Testing Calculator Automation — Plan

## Goal

Automate the manual fill-in / read-back loop on the STRAFE **Testing Calculator**
(`https://strafe.stewart.com/TestingCalculator`). For every endorsement test row in
the workbook, the script types the row's inputs into the calculator, clicks
**Calculate**, reads the resulting **Fee Amount**, and writes it back into the
row's "Actual Value" column. The user then reviews Actual vs. Expected for QA.

No STRAFE API access → we drive the real web UI with Playwright.

```
Test Sheet (xlsx)  →  Playwright  →  existing STRAFE tab in Edge
     →  enter row data  →  click Calculate  →  read Fee Amount
     →  write back into the sheet  →  next row
```

## Workbook

- File: `202x xx xx UT STGC Ends - Test Sheet -1.xlsx` (passed to the script as a
  CLI argument — no hardcoded path).
- Target worksheet: **`STRAFE`** (confirmed by user). Not `STRAFE (2)`,
  `STRAFE Endorsements`, or any other tab.
- The `STRAFE` sheet is currently a template: the header is in place but columns
  **A, G, H, I, J, V, AD** are blank on the data rows. **The user populates those
  before running the script.**

### Global inputs (read once, top of sheet, column E)

| Cell | Meaning | STRAFE field |
|------|---------|--------------|
| `E2` | State (e.g. `UT`) | **State** (Basic Information) |
| `E5` | Underwriter (e.g. `STGC`) | **Underwriter** |
| `E8` | Effective Date — stored as an Excel date serial (e.g. `46272` = `09/07/2026`) | **Closing Date** — overwrite the page default, format `mm/dd/yyyy` |

### Per-row inputs

- Header row: **21**. Sub-header (amount labels): **22**. Data rows: **23** down
  to the last row whose column **B** (Quick Ref) is non-empty.
- Row 22 holds the dollar amounts per Expected/Actual column:
  `G22 = 1000`, `H22 = 4000000`, `I22 = 30000000`, `J22 = 51000000`
  (and the same values repeat under `K22..N22`).

| Excel column (data row) | Value | STRAFE field |
|-------------------------|-------|--------------|
| `B` — Quick Ref | e.g. `UTSTGC0000003` | **Fee Item Code** |
| `A` — Liability Type | exact dropdown wording, provided in the sheet | **Select Liability Type** dropdown — row 1 of the Coverage group |
| `AD` — Coverage Type Code | e.g. `BASIC` | **Coverage Type Code** — row 1 of the Coverage group |
| `V` — Endorsement Question | typed **verbatim** (whether `5` or `Q000005`); only if non-empty | **Boolean Question** — row 1 of the Endorsement Questions section |
| `W` — Boolean Answer | `Yes` / `No`; only when `V` is present | **Select Answer** dropdown next to Boolean Question (row 1) |
| `G` / `H` / `I` / `J` — Expected Value | presence = "run this pass"; the cell's own number is the expected result, used only for later QA | drives the **Amount** passes (see below) |

### The Amount / Fee Amount passes (the core loop)

For each data row, for each column in `G, H, I, J`:

1. If the row's cell in that column is **empty → skip** this pass (leave the
   matching `K`/`L`/`M`/`N` cell untouched).
2. If it has a value → type the corresponding **row-22 amount** into the STRAFE
   **Amount** field (row 1 of the Coverage group):
   `G → 1000`, `H → 4000000`, `I → 30000000`, `J → 51000000`.
3. Click **Calculate**.
4. Read the **Fee Amount** result. Write it into the parallel Actual Value cell,
   same row: `G→K`, `H→L`, `I→M`, `J→N`.
5. If the site shows a **Response/Error Msg** instead of a Fee Amount, write that
   error text into the Actual Value cell instead (per user).

Liability Type + Coverage Type Code + Boolean Question + Boolean Answer are set
**once per row** (row 1 of each repeated UI group); only the Amount field changes
between the up-to-4 passes. Confirmed by user: `G22/H22/I22/J22` header amounts
are what goes into **Amount** — the number in the `G`/`H`/`I`/`J` data cell is
only a "run this pass" flag plus the expected value for later QA.

### Fields left at page default (not touched)

Fee Item Code aside, everything the table above doesn't list: Closing Type Code,
Negotiated Pricing Code, Municipal Tax Code, County, City, Simultaneous Criteria,
Transaction Type, Reissue Information, the 2nd–4th Coverage rows, the 2nd–5th
Endorsement Question rows, Value Question/Answer, Sales Tax Information.

### Row skip logic

- Column `B` empty → not a data row, stop / skip.
- All of `G, H, I, J` empty → nothing to test, skip the row and write
  `No Expected Value` into `K` for that row.
- `A` (Liability Type) or `AD` (Coverage Type Code) empty, **or** `A` not one of
  the known dropdown options → skip the row (no site interaction) and write the
  reason into column `K` for that row, e.g.
  `Skipped: Liability Type missing`, `Skipped: Coverage Type Code missing`,
  `Skipped: Liability Type missing; Coverage Type Code missing`,
  `Skipped: Liability Type "<value>" not a valid option`.

### Select Liability Type — valid dropdown options (exact `<option>` text)

Column `A` must contain one of these strings verbatim; the script selects the
`<option>` whose text equals `A`:

```
Binder Amount
Commitment Amount
Construction Loan Amount
Construction Premium So Far
Guarantee Amount
Leasehold Amount
Leasehold Loan Amount
LMP Premium So Far
Loan Amount
LOP Premium So Far
Mortgage Premium So Far
Owner's Premium So Far
Project Loan Amount
Project Premium So Far
Sales Price
True/Full Rate CMP Premium
True/Full Rate LMP Premium
True/Full Rate Mortgage Premium
True/Full Rate Owners Premium
True/Full Rate PMP Premium
Unpaid Principal Balance
```

## Browser attach (Edge + CDP)

- The script **never launches, closes, reloads, or navigates** a browser.
- One-time setup each session, done by the user:
  1. Close all Edge windows.
  2. Relaunch: `msedge.exe --remote-debugging-port=9222`
  3. Log in to STRAFE, open the **Testing Calculator** page, leave it open.
- The script connects with Playwright `connect_over_cdp("http://localhost:9222")`,
  finds the open `strafe.stewart.com/TestingCalculator` page among the contexts,
  and drives that one tab for the whole run.
- Fails fast with a clear message if `localhost:9222` is unreachable or no
  matching tab is found.
- Because we attach over CDP, **no `playwright install` browser download is
  needed** — we use the user's real Edge.

## Tech stack

- **Python 3.11** (already installed by IT at `C:\Program Files\Python311`).
- **Virtual environment** in the project folder (user's choice):
  - `python -m venv .venv`
  - `.venv\Scripts\python -m pip install playwright openpyxl`
  - run with `.venv\Scripts\python fill_calculator.py "...xlsx"`
- Playwright **sync API** for the browser; **openpyxl** to read/write the xlsx
  (preserves the rest of the workbook's formatting and formulas — pandas would
  not).
- Excel date serial → `mm/dd/yyyy` conversion for `E8` (openpyxl returns a
  `datetime` when the cell is date-formatted; fall back to serial math otherwise).

## Script behavior

- Invocation: `.venv\Scripts\python fill_calculator.py "<path to xlsx>" [--limit N]`
- `--limit N` — process only the first N eligible rows (for a safe trial run).
- Safety copy `<name> (backup).xlsx` made just before writing and **deleted again
  when the run finishes OK** — it survives only if the run fails. `--no-backup` skips it.
- Refuse to start if the xlsx is open/locked in Excel — clear message.
- All CSS/label selectors defined as constants at the top of the file
  (`get_by_label("Underwriter")`, `get_by_label("Fee Item Code")`,
  `get_by_label("Amount")`, the Liability Type `<select>`, `get_by_label("Fee Amount")`,
  `get_by_label("Response/Error Msg")`, the Calculate button, …). These are
  best-effort from the screenshot and may need a tweak against the live DOM —
  one place to fix each.
- After filling Amount, blur / trigger the calc, wait for Fee Amount (or the
  error field) to update before reading. Fee Amount is written as a **number**
  (cell format `"$"#,##0.00`); a "no fee" row gets the message **text**.
- Waits for the site's `#loader-wrapper` overlay to clear before every field
  interaction and after Calculate (up to 60 s).
- **Resume:** a row is skipped when its Actual Value cells already hold a value,
  so rerunning the same command continues from where a stopped run left off.
  `--overwrite` redoes all; `--retry-messages` redoes only text-message cells.
- **Errors:** each row is wrapped; on failure the script saves progress, prints a
  plain-language reason (`explain_error()`) and the row it stopped on, keeps the
  backup, exits non-zero. `--keep-going` skips the failing row instead.
- Per-row / per-pass console progress + an end-of-run summary
  (processed / already-filled / skipped / failed / no-fee-message).
- Save the workbook once at the end (and in the `finally`, so progress survives a crash).

## Confirmed by user (round 2)

- Columns `A` (Liability Type) and `AD` (Coverage Type Code) are now filled in.
- Amount input = **row-22 header value** (not the `G/H/I/J` cell number).
- One Calculate pass per filled `G/H/I/J` column → `G→K`, `H→L`, `I→M`, `J→N`.
- Column `W` (Boolean Answer) holds `Yes` / `No` → pick it in the site's
  "Select Answer" dropdown, only when `V` has a question.
- Manual single-row test worked: filled Effective Date, Underwriter, State,
  Liability Type, Coverage Type, Amount → Calculate → Fee Amount `$10`, which
  belongs in `K23`.

## Open items / assumptions to confirm

1. ~~Skip a row when `A` or `AD` is blank~~ → **Resolved:** skip it, and write the
   reason into column `K` for that row.
2. ~~Exact option text of the Select Liability Type dropdown~~ → **Resolved:**
   list captured above.
3. Whether any row will need the 2nd–4th Coverage rows or 2nd–5th Endorsement
   Question rows (current scope: row 1 only).
4. Does clicking **Calculate** once return the Fee Amount for the single filled
   Coverage row? (manual test says yes — verify on first scripted run)
5. Confirm `E8` is genuinely a date cell (so the serial→date conversion is right).
6. Does the site keep the previously computed Fee Amount visible between passes?
   The script clears/re-reads Amount and waits for Fee Amount to change before
   recording, to avoid recording a stale value.

## Prerequisites checklist

- [ ] IT/user: nothing more for Python (installed).
- [ ] `python -m venv .venv` in `c:\Excel to Strafe Automation`.
- [ ] `.venv\Scripts\python -m pip install playwright openpyxl` (pip **inside the
      venv** — no admin, no Program Files write).
- [ ] User fills columns A, G–J, V, AD on the `STRAFE` sheet.
- [ ] User launches `msedge.exe --remote-debugging-port=9222`, logs in, opens the
      Testing Calculator.
- [ ] Trial run with `--limit 3`, review K–N, then full run.

## Round 3 additions (GUI + checkpoint + resume/change-detection)

- **Engine refactor:** `fill_calculator.py` core is `run_engine(cfg, hooks)`.
  `EngineConfig` holds the options; `Hooks` receives events
  (`start/row_begin/pass_done/progress/row_skip/row_done_already/row_changed/
  row_failed/checkpoint/resumed/saved/stopped/aborted/finished`) and answers
  `should_stop()` / `wait_at_checkpoint()`. `ConsoleHooks` = CLI; `GuiHooks`
  (in the GUI file) marshals events over a `queue.Queue` to the Tk thread.
  Setup problems raise `SetupError` instead of `sys.exit`.
- **10% checkpoint:** after `ceil(0.10 * rows-to-process)` processed rows (only
  when that set is ≥ 5), save workbook + state, emit `checkpoint`, block on
  `wait_at_checkpoint()`. On "continue" the engine **reloads the workbook from
  disk** (picks up user edits) and re-enters globals if they changed. Once/run.
  `--no-checkpoint` disables. GUI opens the saved file in Excel via `os.startfile`.
- **Clean stop:** `hooks.should_stop()` checked at the top of each row; the
  current row finishes, then save + `stopped` + return. GUI Stop sets a
  `threading.Event`. CLI stop = type `s` at the checkpoint (or Ctrl+C anywhere).
- **End-of-run "Results look good?"** (`Hooks.ask_satisfied(summary)`): called
  once after a completed run (`run_ok`). Yes → delete `<name> (state).json`
  (batch done); No → keep it. `ConsoleHooks` prompts `[y/N]` (non-interactive =
  keep); `GuiHooks` shows a green bar with Yes / No, blocking the worker on a
  `done_reply` queue like the checkpoint. Not asked on stop/abort.

- **Change detection / resume:** sidecar `<name> (state).json`
  `{ "rows": { "23": {"sig": "<sha1[:16]>", "cols": ["K","N"]} } }`. `sig` hashes
  the row inputs (A/AD/V/W/G–J) + globals sig. Per pass, a cell is (re)done when:
  empty, or `--overwrite`, or **inputs changed vs stored sig** (even if filled;
  the stale cell is cleared first), or `--retry-messages` and the cell holds text.
  Rows that are skipped-as-done still get their current sig written, so a later
  edit is detectable from the next run on. Also saves every 10 processed rows.
- **GUI** `fill_calculator_gui.py` (pure Tkinter, no extra deps so it bundles
  cleanly): 3 gated step "cards" (number badge + title + description + coloured
  status line with ✓/✗), flat coloured buttons, dark header band, styled
  (`clam`) progress bar, scrolling log, Stop button, amber checkpoint bar with
  Continue / Stop-&-review. Worker thread runs `run_engine`; `root.after()` pumps
  the event queue. `_enable_dpi_awareness()` (SetProcessDpiAwareness) + tk scaling
  → crisp on high-DPI. `Run GUI.bat` launches it with `pythonw.exe`.
- **`launch_debug_edge()` / `verify_cdp()`** in `fill_calculator.py`: find
  `msedge.exe`, start it with `--remote-debugging-port=9222 --user-data-dir=
  .edge-automation <url>`, then poll `/json/version`. Used by the GUI's
  Connect Edge button; the `.bat` still works too.

## Packaging to a single .exe (round 4)

- `build_exe.bat` + `STRAFE Fee Calculator.spec` → PyInstaller one-file build of
  `fill_calculator_gui.py` → `dist\STRAFE Fee Calculator.exe`.
- Spec uses `collect_all` for `playwright` (bundles the **Node driver** — needed
  even for CDP-attach; no browser is bundled), `openpyxl`, `et_xmlfile`.
  `console=False`; `_main_guarded()` writes `last-error.log` + a dialog on a
  startup crash.
- **Edge launch reuses the approved .bat:** `launch_debug_edge()` runs
  `launch_edge_debug.bat` (via `cmd /c`, hidden) when that file sits next to the
  script/exe — so the GUI/exe inherit IT's existing approval and share the one
  `.edge-automation\` profile the .bat creates. Only if the .bat is absent does it
  fall back to launching `msedge.exe` directly. The .bat itself was reverted to
  its IT-approved bytes (editing it had changed its hash and re-triggered the
  prompt). For the packaged .exe, keep `launch_edge_debug.bat` beside it.
- **Frozen-safe paths:** `base_dir()` = exe folder when frozen else script folder.
  `app_data_dir()` = `%LOCALAPPDATA%\STRAFE Fee Calculator\` holds the error log
  and the *fallback* Edge profile. `(state).json` / `(backup).xlsx` still derive
  from the chosen workbook path.
- Known caveats: ~150–200 MB, slow first launch (self-extract), and unsigned
  PyInstaller exes often trip corporate AV / SmartScreen — IT allow-list or
  code-sign; fallback is a `--onedir` folder, zipped.
- Target machine needs Microsoft Edge installed.

## Documentation & drive packaging (round 7)

- `docs/User Guide.md` — end-user walkthrough (the 3 steps, checkpoint, stop/
  resume, "Results look good?", user-level troubleshooting).
- `docs/Build and Deploy Guide.md` — maintainer: dev setup, `build_exe.bat`
  output & the standalone-exe facts, the AV/SmartScreen caveat + onedir fallback,
  the shared-drive folder layout, releasing a new version, fixing selectors,
  a "maintainer contact" blank to fill in.
- `docs/drive-readme.txt` — the `README.txt` dropped at the package root.
- `requirements.txt` — playwright, openpyxl, pyinstaller.
- `make_drive_package.bat` — assembles `drive package\STRAFE Fee Calculator\`
  with `App\` (exe + launch_edge_debug.bat + User Guide), `Source\`,
  `Sample workbook\`, `docs\`, `README.txt`; copies the exe from `dist\` if built,
  else writes an "EXE MISSING" placeholder.
- `README.md` gained a Documentation table + a Version log.

## Files

- `planning.md` — this document.
- `fill_calculator.py` — engine + CLI.
- `fill_calculator_gui.py` — Tkinter front end.  `Run GUI.bat` — its launcher.
- `build_exe.bat`, `STRAFE Fee Calculator.spec` — build the standalone .exe.
- `make_drive_package.bat`, `requirements.txt`, `docs/` — full packaging + docs.
- `make_app_folder.bat` — lightweight package: one flat folder
  `STRAFE Fee Calculator\` with just the exe + `User Guide.pdf` +
  `launch_edge_debug.bat` (rebuilds the PDF each time). This is what actually
  got pushed to the drive (round 8) - user wanted a single simple folder rather
  than the full Source/docs/Sample-workbook tree.
- `launch_edge_debug.bat` — start the debug Edge (dedicated `.edge-automation\`
  profile, login remembered) and open the calculator.
- `<name> (state).json` — per-row input fingerprints for resume/change-detection.
- `<name> (backup).xlsx` — safety copy; present only after a failed run.
- `202x xx xx UT STGC Ends - Test Sheet -1.xlsx` — data.

## Confirmed live DOM ids (from --list-fields)

| Field | Element id |
|-------|-----------|
| Closing Date | `#ClosingDate` (`<input type=date>` → needs `YYYY-MM-DD`) |
| Underwriter | `#Underwriter` |
| State (Basic Info) | `#PropertyInformation_PropertyState` (Sales-Tax State is `#PropertyInformation_SalesTaxState` — not used) |
| Fee Item Code | `#feeItemCodeInput` |
| Select Liability Type (row 1) | `select#BasisInfo[0].IdAmountName` |
| Coverage Type Code (row 1) | `#BasisInfo[0].BaseTitleInsTypeCode` |
| Amount (row 1) | `#BasisInfo[0].ValueStr` |
| Boolean Question (row 1) | `#QuestionList[0].Code` |
| Boolean Answer (row 1) | `select#QuestionList[0].AnswerType` (`Select Answer` / `Yes` / `No`) |
| Calculate | `button#btnTestToolCalculate` (Clear is `#btnTestToolClear`) |
| Fee Amount (result) | `#txtFeeAmount` (shows `$0.00` until a calc succeeds) |
| Response/Error Msg | `textarea#taResponseErrorMsg` (also shows the endorsement **name** on a *successful* calc — so it is not purely an error box) |

**Result rule:** read `#txtFeeAmount`; if it parses to a non-zero number, write
that **number** into `K/L/M/N` (cell format `"$"#,##0.00`). If it is `$0.00` /
blank, write the `#taResponseErrorMsg` **text** instead.

## If a selector is wrong

The label/xpath selectors in `SELECTORS` are best-effort (script author has no
access to the live site). If a field can't be found:

    .venv\Scripts\python fill_calculator.py "<xlsx>" --list-fields

dumps every input/select/button on the page with its id/name/placeholder/nearby
label — send that output back and the one wrong `SELECTORS[...]` entry gets fixed.
`--dry-run` prints the per-row plan from the sheet without touching the browser.