# STRAFE Fee Calculator — User Guide

A small Windows app that fills the STRAFE **Testing Calculator** from your test
workbook and writes the calculated **Fee Amount** back into the sheet, so you
don't have to type ~260 rows by hand.

You need: **Microsoft Edge** (already on every work PC) and the test **workbook**.
Nothing to install — it's one program file.

---

## 1. Start the app

Double-click **`STRAFE Fee Calculator.exe`**.

- The first launch takes ~15 seconds (it unpacks itself). Later launches are quick.
- If Windows shows a blue "Windows protected your PC" box: click **More info →
  Run anyway**. (If it's fully blocked, contact IT — see the note at the end.)

You'll see a window with three numbered steps. Each step lights up green when it's
done and unlocks the next one.

---

## 2. Step 1 — Connect Edge

Click **Connect Edge**.

- A new Edge window opens on the Testing Calculator page.
- **Log in to STRAFE** in that window if it asks. (It remembers you next time.)
- Leave that window on the Testing Calculator page. **Don't close it, don't click
  around in it** while the app runs.

When step 1 turns green ("Connected"), move on.

> If it fails: close **every** Edge window (check the taskbar and the system tray),
> then click **Connect Edge** again.

---

## 3. Step 2 — Choose Excel file

Click **Choose file…**, browse to your test workbook, and select it.

The app checks the file has a **STRAFE** sheet with data rows. When step 2 turns
green (showing the file name and row count), move on.

> **Close the workbook in Excel before running.** The app can't write to it while
> Excel has it open.

---

## 4. Step 3 — Run it

Click **Run it**.

- Watch the Edge window fill in each row. The log in the app shows every value it
  writes (e.g. `row 27:  G → K27 = 100.0`), and the progress bar shows `12 / 40`.
- **Stop** button: stops cleanly **after the current row**. You can pick up later
  (see "Stopping and continuing" below).
- **"Redo everything"** checkbox: tick it before running to re-do every row from
  scratch, ignoring anything already filled in.

### The checkpoint (automatic)

After about **10%** of the rows, the app **pauses**, saves, and **opens the
workbook in Excel** so you can sanity-check the first batch.

1. Look over the filled-in Actual Value cells.
2. **Close Excel.**
3. Click **Continue** (or **Stop & review** if something's wrong).

Anything you corrected in the sheet during the pause is picked up when it continues.

### When it finishes

It asks **"Results look good?"**

- **Yes – all done** → it tidies up (removes the small helper file it uses to
  track progress). Use this when the batch is finished.
- **No – keep it** → keeps that helper file, so you can fix rows in Excel and
  click **Run again** — it will only redo the rows you changed.

---

## What ends up in the sheet

For each row it writes the Fee Amount as a **number** into columns **K / L / M / N**
(matching the Expected Value columns G / H / I / J).

If the calculator gives **no fee** for a row (e.g. a "Call for pricing"
endorsement), it writes the **message text** into that cell instead — those rows
need a manual look. The end-of-run summary and the log tell you which rows.

Rows are **skipped** (with the reason written into column K) when the Liability
Type or Coverage Type Code is missing, the Liability Type isn't a real dropdown
option, or there's no Expected Value.

---

## Stopping and continuing

You can stop any time (the **Stop** button, or **Stop & review** at the
checkpoint, or just closing the app). Nothing is lost — progress is saved as it
goes.

To continue: start the app again, do steps 1 and 2, and click **Run it**. It
**skips every row that's already filled** and carries on from where it stopped.
If you **edited an input** on a row that was already done, that row gets redone
automatically.

---

## Backups

Just before it writes, the app makes a copy of your workbook next to it called
`… (backup).xlsx`. If the run finishes normally, that copy is **deleted**. It's
only left behind if a run fails — so you can restore from it.

---

## If something goes wrong

| What you see | What to do |
|---|---|
| "could not attach to Edge" | Close all Edge windows, click **Connect Edge** again. |
| "no open tab is on the Testing Calculator" | In the Edge window, log in and open the Testing Calculator page. |
| "… open in Excel" | Close the workbook in Excel. |
| It stopped with "STOPPED at row N" | Read the reason shown. Fix it, then **Run again** — it resumes at row N. |
| A cell has text instead of a dollar amount | The calculator returned no fee for that row — check it by hand. |
| The app won't start / antivirus blocks it | Send this to IT: *Please allow-list (or code-sign) `STRAFE Fee Calculator.exe`* |

Questions or a selector stopped matching the site → contact the maintainer listed
in the Build & Deploy Guide ( amubeen3636@gmail.com).
