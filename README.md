# Excel → STRAFE Testing Calculator Automation

Fills the STRAFE **Testing Calculator** (`https://strafe.stewart.com/TestingCalculator`)
from the **STRAFE** worksheet of the endorsement test workbook, clicks **Calculate**,
and writes the resulting **Fee Amount** back into the sheet's Actual Value columns.

No STRAFE API access → it drives the real web page with Playwright, attaching to an
Edge tab you open and stay logged into.

Two ways to run it:

- **`fill_calculator_gui.py`** — a small window with 3 buttons (for non-technical users).
- **`fill_calculator.py`** — the same engine from the command line, with options.

## Documentation

| Doc | For whom |
|-----|----------|
| [`docs/User Guide.md`](docs/User%20Guide.md) | end users — how to use the app / the `.exe` |
| [`docs/Build and Deploy Guide.md`](docs/Build%20and%20Deploy%20Guide.md) | maintainer — dev setup, building the `.exe`, the shared-drive layout, updating selectors |
| this file (`README.md`) | technical reference — field mapping, all CLI flags, resume/checkpoint behaviour |
| [`planning.md`](planning.md) | design history / decisions / confirmed DOM ids |

### Version log

- **2026-09** — initial build; GUI + single-`.exe` packaging; 10% checkpoint,
  clean stop/resume, input-change detection, end-of-run "Results look good?".

---

## Files in this folder

| File | What it is |
|------|------------|
| `fill_calculator_gui.py` | the click-through window |
| `Run GUI.bat` | double-click to open that window (needs the venv) |
| `build_exe.bat` + `STRAFE Fee Calculator.spec` | build the standalone `dist\STRAFE Fee Calculator.exe` |
| `fill_calculator.py` | the engine / command-line version |
| `launch_edge_debug.bat` | starts Edge in the mode the script can attach to (the GUI's "Connect Edge" does the same) |
| `202x xx xx UT STGC Ends - Test Sheet -1.xlsx` | the workbook (data in + results out) |
| `... (backup).xlsx` | temporary safety copy; present only if the last run failed |
| `... (state).json` | per-row input fingerprints used for resume / change-detection (safe to delete; forces a fresh full run) |
| `requirements.txt` | Python packages for dev / building the exe |
| `make_drive_package.bat` | assembles `drive package\` for uploading to the shared drive |
| `docs/` | User Guide + Build & Deploy Guide |
| `planning.md` | design notes / field mapping / confirmed DOM ids |
| `.venv/` | the Python virtual environment (created during setup) |
| `.edge-automation/` | the dedicated Edge profile (keeps your STRAFE login) |

---

## One-time setup

Python 3.11 is already installed. In a terminal in this folder:

```powershell
python -m venv .venv
.venv\Scripts\python -m pip install playwright openpyxl
```

No `playwright install` / browser download is needed — it uses your real Edge.

---

## Single .exe (for people without Python)

Build once on a machine that has the `.venv`:

```powershell
.\build_exe.bat
```

That produces **`dist\STRAFE Fee Calculator.exe`** — one file, no Python/venv/pip
needed to run it. Give that file to users.

- The target machine must have **Microsoft Edge** (the .exe drives the user's Edge;
  it does not bundle a browser).
- First launch is slow (~10–20 s) — a one-file exe unpacks itself to a temp folder.
- Size is ~150–200 MB (it bundles Playwright's Node driver).
- Keep **`launch_edge_debug.bat` next to the .exe** — the exe runs it for "Connect
  Edge" so it reuses IT's approval and one Edge profile. Without it, the exe launches
  Edge directly and uses a profile under `%LOCALAPPDATA%\STRAFE Fee Calculator\`.
- Error log goes to `%LOCALAPPDATA%\STRAFE Fee Calculator\last-error.log`; the
  `(state).json` and `(backup).xlsx` sit next to the chosen workbook, same as before.
- **Antivirus / SmartScreen:** an unsigned PyInstaller exe often triggers a warning
  or quarantine on managed machines. IT may need to allow-list it (or code-sign it).
  If that's a problem, build a folder instead of one file — remove the `EXE(...)`
  `console`/one-file packaging isn't the issue; use `--onedir` and zip `dist\STRAFE Fee Calculator\`.

The .exe *is* the GUI below — same three steps.

## Using the GUI

Double-click **`Run GUI.bat`** (or `.venv\Scripts\python fill_calculator_gui.py`),
or the built **`STRAFE Fee Calculator.exe`**.

1. **Connect Edge** – runs `launch_edge_debug.bat` (the IT-approved launcher, if it's
   next to the app) to open Edge on the Testing Calculator with the debugging port on,
   then checks it. The first time, **log in to STRAFE** in that window. When it goes
   green, step 2 unlocks.
2. **Choose Excel file** – native file picker, `.xlsx` only. It checks the file has a
   `STRAFE` sheet with data rows. When OK, step 3 unlocks.
3. **Run it** – fills the calculator row by row. A progress bar and log show what's
   happening. **Stop** halts after the current row. Tick *"Redo everything"* to ignore
   saved results and rerun every row.

**Checkpoint:** after the first ~10% of rows, it saves, **opens the workbook in Excel
for you**, and waits. Review it, close Excel, then click **Continue** (or **Stop &
review**). Anything you corrected in the sheet is picked up.

---

## Using the command line

```powershell
# Edge first (once per session):
.\launch_edge_debug.bat            # then log in, keep the Testing Calculator open

# close the workbook in Excel, then:
.venv\Scripts\python fill_calculator.py "202x xx xx UT STGC Ends - Test Sheet -1.xlsx" --limit 3
.venv\Scripts\python fill_calculator.py "202x xx xx UT STGC Ends - Test Sheet -1.xlsx"
```

CLI checkpoint: it prints `CHECKPOINT ...` and waits — press **Enter** to continue, or
type **s** then Enter to stop.

---

## Resume & change-detection

Every run records each row's inputs (Liability Type, Coverage Code, Question, Answer,
G–J, and the globals) in `... (state).json`. On the next run it:

- **skips** rows whose Actual Value cells are already filled *and* whose inputs are unchanged,
- **re-runs** a row whose inputs you edited — even if its result was already filled,
- **fills** any result cell you cleared,
- **resumes** from wherever a stopped/crashed run left off (progress is saved before exit,
  every 10 rows, and at the checkpoint).

So the normal loop is: run → stop when you want → fix rows in Excel → run again.

**When a run finishes** it asks *"Results look good?"*
- **Yes** → the `(state).json` file is deleted (this batch is done, folder stays clean).
- **No** → it's kept, so you can edit rows and Run again with change-detection intact.

(CLI: it prompts `Results look good? … [y/N]`. Non-interactive runs keep the file.)

---

## What it reads and writes (sheet "STRAFE")

**Set once, from column E at the top:** `E2` → State, `E5` → Underwriter,
`E8` → Closing Date (Excel date → `YYYY-MM-DD`).

**Per data row (row 23 down, while column B has a Quick Ref):**

| Column | → STRAFE field |
|--------|----------------|
| `B` Quick Ref | Fee Item Code |
| `A` Liability Type | "Select Liability Type" dropdown (row 1) — must exactly match a dropdown option |
| `AD` Coverage Type Code | Coverage Type Code (row 1) |
| `V` Endorsement Question | Boolean Question (row 1) — only if filled |
| `W` Boolean Answer (`Yes`/`No`) | Select Answer dropdown (row 1) — only if `V` filled |

**Amount passes** — for each of `G / H / I / J` with a value, the **row-22 header
amount** for that column (`G→1000  H→4,000,000  I→30,000,000  J→51,000,000`) is typed
into **Amount**, Calculate is clicked, and the Fee Amount is written as a **number**
(currency-formatted) to `G→K  H→L  I→M  J→N` in the same row.

If the calculator returns no fee (e.g. a "Call for pricing" endorsement), the message
**text** is written there instead — check those rows by hand.

**Rows are skipped** (reason written into column `K`) when `A` or `AD` is blank, `A`
is not a valid dropdown option, or none of `G/H/I/J` has a value.

---

## Backups

Just before writing, the workbook is copied to `... (backup).xlsx`. If the run
**finishes OK** that copy is **deleted**. It is **kept only if the run fails**, so you
can restore. `--no-backup` skips it.

---

## Command-line options

| Flag | Meaning |
|------|---------|
| `--limit N` | process at most N rows this run |
| `--start-row R` | begin at sheet row R (default 23) |
| `--sheet NAME` | worksheet name (default `STRAFE`) |
| `--dry-run` | print the per-row plan; don't touch the browser |
| `--overwrite` | redo every row, ignore existing values / state |
| `--retry-messages` | redo only cells holding a text message instead of a `$` amount |
| `--keep-going` | on a row error, log and continue instead of stopping |
| `--no-checkpoint` | don't pause at ~10% |
| `--no-backup` | don't make the `(backup)` safety copy |
| `--list-fields` | dump every field on the page with its id (use if a selector breaks) |
| `--cdp-url URL` | CDP endpoint (default `http://127.0.0.1:9222`) |
| `--settle-ms MS` | wait budget for the result after Calculate (default 4000) |

---

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| `could not attach to Edge` | Close all Edge, click **Connect Edge** again (or run `launch_edge_debug.bat`), check `127.0.0.1:9222/json/version` shows JSON. |
| `no open tab is on the Testing Calculator` | In the debug Edge window, log in and open the Testing Calculator page. |
| `... open in Excel` | Close the workbook. |
| `Could not locate field 'X'` / values in the wrong box | The site DOM changed. Run `fill_calculator.py --list-fields`, send the output, selectors get repinned in `SELECTORS`. |
| `STOPPED at row N` / loading spinner never cleared | The calculator hung. Run again — it resumes at row N. If it repeats, add `--settle-ms 8000`. |
| A row wrote text instead of a `$` amount | The calculator returned no fee — check it by hand. After fixing the sheet, rerun with `--retry-messages` to redo just those. |