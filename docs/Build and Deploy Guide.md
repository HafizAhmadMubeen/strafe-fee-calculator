# STRAFE Fee Calculator — Build & Deploy Guide

For whoever maintains the tool. End users only need the `.exe` and the
[User Guide](User%20Guide.md).

---

## 1. What it is

- `fill_calculator.py` — the engine + a command-line interface.
- `fill_calculator_gui.py` — the Tkinter window (this is what the `.exe` wraps).
- It drives the STRAFE **Testing Calculator** web page with **Playwright**,
  attaching over the DevTools protocol to an Edge the user already has open.
  It never bundles or launches its own browser.
- It reads/writes the **STRAFE** sheet of the test workbook with **openpyxl**.

Full field mapping and design notes: [`../planning.md`](../planning.md).

---

## 2. Developer setup (one time, on your machine)

```powershell
cd <the source folder>
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
```

Run from source while developing:

```powershell
.venv\Scripts\python fill_calculator_gui.py                 # the GUI
.venv\Scripts\python fill_calculator.py "<workbook>.xlsx"   # the CLI
```

CLI options: `--limit N`, `--dry-run`, `--overwrite`, `--retry-messages`,
`--keep-going`, `--no-checkpoint`, `--no-backup`, `--list-fields`,
`--settle-ms MS`, `--start-row R`, `--sheet NAME`, `--cdp-url URL`.

---

## 3. Building the single .exe

```powershell
.\build_exe.bat
```

This installs PyInstaller into `.venv` and runs it against
`STRAFE Fee Calculator.spec`. Output:

```
dist\STRAFE Fee Calculator.exe
```

### Yes — that .exe is fully standalone

- **No Python, no venv, no pip** is needed to run it. Copy that one file anywhere
  and double-click.
- It bundles the Playwright Node driver, so it's **~150–200 MB**.
- **First launch is slow (~10–20 s)** — a one-file exe unpacks itself to a temp
  folder every time it starts. (If that's annoying, build an unpacked folder
  instead — see below.)
- The machine running it **must have Microsoft Edge** (every corporate PC does).
- It writes `… (state).json` and `… (backup).xlsx` **next to the workbook the
  user picks**, and its Edge profile + error log under
  `%LOCALAPPDATA%\STRAFE Fee Calculator\`.

### The catch: antivirus / SmartScreen

An **unsigned PyInstaller one-file exe is very likely to be quarantined or
blocked** on managed Windows machines. Before you hand it out, get one of:

1. **IT code-signs** `STRAFE Fee Calculator.exe` (best — no user warnings), or
2. **IT allow-lists** its SHA-256 hash in the endpoint protection tool, or
3. Ship an **unpacked folder** instead of one file — sometimes tolerated better,
   and it starts instantly. To do that, edit `STRAFE Fee Calculator.spec`: after
   the `EXE(...)` block add a `COLLECT(...)` and set the `EXE` `exclude_binaries=True`
   (standard PyInstaller onedir form), rebuild, then zip `dist\STRAFE Fee Calculator\`.

Until one of those is done, "anyone can download and run it" is **not** true in
practice — it will be stopped by security tooling.

### Keep `launch_edge_debug.bat` next to the .exe

The exe runs that (already IT-approved) script for the "Connect Edge" button, so
it inherits the approval and uses one shared Edge profile. If the .bat isn't
there, the exe launches Edge directly (which may need its own approval).

**Do not edit `launch_edge_debug.bat`.** Any change alters its hash and
re-triggers IT's approval prompt.

---

## 4. Folder layout on the shared drive

First (re)build the user-guide PDF if you changed `docs\User Guide.md`:

```powershell
.venv\Scripts\python docs\md_to_pdf.py "docs\User Guide.md"
```

(`md_to_pdf.py` is a tiny no-dependency Markdown→PDF converter that renders via
headless Edge. Needs Edge, nothing else.)

Then run **`make_drive_package.bat`** — it assembles this into `drive package\`,
including the built exe if `dist\` has one, and the User Guide PDF if it exists:

```
STRAFE Fee Calculator\
│
├─ README.txt                     "start here" - points to the folders below
│
├─ App\                           >>> give this folder to end users <<<
│   ├─ STRAFE Fee Calculator.exe
│   ├─ launch_edge_debug.bat
│   └─ User Guide.pdf             (built by docs\md_to_pdf.py)
│
├─ Source\                        for the maintainer / rebuilding
│   ├─ fill_calculator.py
│   ├─ fill_calculator_gui.py
│   ├─ build_exe.bat
│   ├─ STRAFE Fee Calculator.spec
│   ├─ Run GUI.bat
│   ├─ requirements.txt
│   ├─ README.md
│   └─ planning.md
│
└─ docs\
    ├─ User Guide.md
    ├─ Build and Deploy Guide.md   (this file)
```

Then upload the whole `STRAFE Fee Calculator\` folder to the drive. Point users
at **`App\`** only.

---

## 5. Releasing a new version

1. Make the code change.
2. Bump a version note at the top of `README.md` (date + what changed).
3. `.\build_exe.bat`
4. Smoke-test the exe on a machine **without** Python.
5. `.\make_drive_package.bat`
6. Replace the folder on the drive. Tell users to grab the new `App\` folder
   (the exe keeps the same name, so their shortcut still works).

---

## 6. If the STRAFE page changes and fields stop matching

The web element ids are pinned in `SELECTORS` near the top of
`fill_calculator.py`. If a field can't be found:

```powershell
.venv\Scripts\python fill_calculator.py "<workbook>.xlsx" --list-fields
```

That dumps every input/select/button on the live page with its `id`. Update the
one wrong entry in `SELECTORS`, rebuild, redeploy.

---

## 7. Maintainer contact

_Fill in:_ name / email / Teams — the person to contact for changes, a broken
selector, or a new rate-update sheet layout.
amubeen3636@gmail.com
