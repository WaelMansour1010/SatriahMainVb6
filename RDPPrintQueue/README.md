# RDP PDF Auto-Print Watcher

Local Windows fallback for RDP printing: PDFs that the VB6 server app drops on
the client (via an RDP-redirected folder) are printed automatically with
SumatraPDF.

## Folder layout

```
C:\RDPPrintQueue\
    RdpPdfAutoPrint.ps1         <- main script
    Start-RdpPdfAutoPrint.bat   <- production launcher (must keep running)
    Test-RdpPdfAutoPrint.bat    <- prints a single known PDF (smoke test only)
    Install-AutoStart.bat       <- register per-user Scheduled Task at logon
    Uninstall-AutoStart.bat     <- remove the Scheduled Task
    Check-AutoStart.bat         <- show task state, last run, running process
    inbox\                      <- new PDFs arrive here
    processing\                 <- held while printing
    archive\                    <- printed OK
    failed\                     <- printing failed
    logs\                       <- auto-print-YYYY-MM-DD.log
    state\                      <- printed-jobs.csv (dedup db), watcher.lock
```

## Root cause of the previous failure

The old launcher passed `%LOCALAPPDATA%\SumatraPDF\SumatraPDF.exe` to the
PowerShell script. `%LOCALAPPDATA%` expands to whatever user token is active
at launch, so:

- When the script ran under a different session (different logon, scheduled
  task with a "run-as" account, RDP reconnect, SYSTEM), `LOCALAPPDATA` pointed
  at a profile that does **not** have SumatraPDF installed — hence
  "SumatraPDF not found using %LOCALAPPDATA%".
- Even if Sumatra had been found, `-print-to-default` relies on the *user's*
  default printer. Default printer is a per-user / per-session setting; under
  a non-interactive or wrong-user context the default printer is unset or
  different from what the tech sees in Settings.

Manual execution as `Wael` worked because that user happened to have both
Sumatra in LOCALAPPDATA and a usable default printer — neither is guaranteed
on other machines or other sessions.

## What the refactor fixes

1. **`Resolve-SumatraPath`** — explicit `-SumatraPath` → Program Files →
   Program Files (x86) → `%LOCALAPPDATA%`. Machine-wide installs are preferred.
2. **`Resolve-Printer`** — smart heuristic (see below). Passes printer name
   explicitly as `-print-to "<name>"` instead of `-print-to-default`.
3. **Startup diagnostics** — logs `whoami`, `USERNAME`, `USERPROFILE`,
   `COMPUTERNAME`, resolved Sumatra, resolved printer, full printer list.
4. **Test Mode (`-TestFile`)** — prints one PDF and exits with a specific code.
5. **DryRun mode (`-DryRun`)** — resolves everything, logs, does not print.
6. **Clean separation** — each concern lives in its own function.
7. **Queue + dedup preserved** — `printed-jobs.csv` keyed on SHA256 + size.

## Printer selection heuristic

When `-PrinterName` is **not** specified, `Resolve-Printer` picks in this order:

| Priority | Condition |
|----------|-----------|
| 1 | Redirected printer that is **also** the default |
| 2 | Any redirected printer — name contains `(redirected N)` or PortName starts with `TS`/`Client` |
| 3 | Windows default printer that is **not** excluded |
| 4 | First acceptable (non-excluded) printer |
| 5 | `WScript.Network.EnumPrinterConnections()` fallback — odd-indexed entries only (even = port names, odd = printer names) |

**Excluded virtual printers** (never selected automatically):
`Microsoft Print to PDF`, `Microsoft XPS Document Writer`, `Fax`,
any `OneNote` variant, `Adobe PDF`, `PDFCreator`, `CutePDF Writer`.

When `-PrinterName` **is** specified:
- Exact name match first.
- If not found, loose prefix match — so `HP LaserJet` matches
  `HP LaserJet (redirected 3)` across RDP reconnects.
- Fails clearly if neither matches.

**Printer re-resolution:** the printer is resolved live before every print
attempt (not only at startup). If the redirected printer was renamed during a
reconnect, the next attempt picks up the new name automatically.

**No-printer resilience:** in watcher mode, an unavailable printer does **not**
stop the process. The file stays in `inbox`, a throttled warning is logged
(at most once per 60 s), and the watcher retries automatically once a printer
appears.

## Clipboard backup (every job)

Every PDF is copied to the Windows clipboard as a **file-drop object** so
that, if printing fails for any device-specific reason, the operator can
press `Ctrl+V` in Explorer (or in the Print dialog of any PDF viewer) and
paste the file to print it manually.

Implementation: `Set-ClipboardFileDrop` in `RdpPdfAutoPrint.ps1` calls
`[System.Windows.Forms.Clipboard]::SetFileDropList()`. Because that API
requires an STA apartment, a dedicated STA runspace is created per call so
the watcher itself does not need to be launched with `-STA`.

**Timing (two clipboard updates per job, both unconditional):**

1. **Before the print attempt** — clipboard points to the file in
   `processing\`. This is the "safety net" so the user has the PDF ready
   even if the print hangs or times out.
2. **After the final move** — clipboard is re-pointed to the stable path
   (`archive\<file>.pdf` on success, `failed\<file>.pdf` on failure).

The second update is necessary because Windows file-drop clipboard stores
**paths, not bytes**. Once step 6 of the pipeline moves the file out of
`processing\`, the earlier clipboard path would be stale.

**Logged lines (all at INFO/WARN level):**

```
Clipboard set       : 20260418_104501_123_SaleReport545.pdf   (success)
Clipboard timed out : <filename>                              (STA runspace hung >5 s)
Clipboard error     : <message>                               (SetFileDropList threw)
Clipboard failed    : <message>                               (outer exception)
```

**Known limitations:**

- Clipboard is clobbered twice per job. If an operator copied unrelated
  content while a job was being processed, that content is overwritten.
- If the watcher runs in a different interactive session than the user's
  visible desktop (unlikely with the per-user Scheduled Task, but possible
  under `runas` or unusual tools), the clipboard target is the session in
  which the watcher runs — not necessarily what `Ctrl+V` will paste in the
  operator's visible session.
- If `System.Windows.Forms` cannot load (trimmed OS image, Server Core
  without the Desktop Experience feature), clipboard set will fail. The
  failure is logged as `WARN` and printing continues normally.

## Per-job failure categories (logged in `Result` column of printed-jobs.csv)

| Category | Meaning |
|----------|---------|
| `OK` | Printed successfully |
| `DUPLICATE` | Already printed (SHA256+size match), moved to archive |
| `NOPRINTER` | No acceptable printer visible at attempt time |
| `TIMEOUT` | Sumatra process exceeded `PrintTimeoutSeconds` |
| `SUMATRA_NONZERO` | Sumatra exited with non-zero code |
| `EXCEPTION` | Unexpected PowerShell exception |

## How to run it

### Recommended: Scheduled Task at user logon

- Action: `C:\RDPPrintQueue\Start-RdpPdfAutoPrint.bat`
- Trigger: *At log on* of the interactive user
- Run only when user is logged on (**do not** use "Run whether user is logged
  on or not" and **do not** run as SYSTEM — see below)
- Highest privileges: **not** required
- Start in: `C:\RDPPrintQueue`

This gives the watcher the same user token as the RDP session, so it sees the
same default printer and redirected printers the user sees.

### Acceptable alternative: Startup folder

Put a shortcut to `Start-RdpPdfAutoPrint.bat` in
`shell:startup`. Simpler to deploy, but harder to restart cleanly if the
watcher dies. Scheduled Task is preferred.

### Why NOT a Windows Service

- Default printer and RDP-redirected printers are **per-user / per-session**.
  A service running as `LocalSystem` or a different service account sees a
  different (usually empty) printer list.
- Sumatra's `-print-to` relies on the calling user's print spooler context.
- For RDP redirected printers specifically, they exist only inside the RDP
  session and are invisible to session 0 (where services live).

If a service is mandatory (e.g. no interactive user), you must:
- install a machine-wide printer (not a redirected one),
- run the service under a domain user that has that printer,
- pass `-PrinterName` explicitly.
Even then, a per-user Scheduled Task is simpler and less fragile.

## Auto-start on every reboot / logon

Use the supplied helpers. They create a **per-user Scheduled Task** named
`RdpPdfAutoPrintWatcher` that triggers at logon and runs the production
launcher; no Windows Service, no SYSTEM account.

```
C:\RDPPrintQueue\Install-AutoStart.bat      :: register (idempotent)
C:\RDPPrintQueue\Check-AutoStart.bat        :: verify state / last result
C:\RDPPrintQueue\Uninstall-AutoStart.bat    :: remove
```

- `Install-AutoStart.bat` calls `schtasks /Create /SC ONLOGON /RL LIMITED /F`,
  with a 30s logon delay (gives the user's print spooler time to publish
  redirected printers), and a `cmd /c cd /d C:\RDPPrintQueue && ...` wrapper
  so the task's working directory is `C:\RDPPrintQueue`. `/F` makes it safe
  to re-run — it overwrites the existing task instead of erroring.
- The task runs as `%USERDOMAIN%\%USERNAME%` — the user who installed it.
  Install it while logged on as the user who actually prints.
- If you need to install on a machine that has a shared kiosk account, log
  on as that kiosk user and run the installer.

### Which file goes where

| File                        | Use                                            |
|-----------------------------|------------------------------------------------|
| `Start-RdpPdfAutoPrint.bat` | **Production** — run by the Scheduled Task     |
| `Test-RdpPdfAutoPrint.bat`  | Manual smoke test only; do **not** auto-start  |
| `Install-AutoStart.bat`     | One-time per user, to register the task        |
| `Check-AutoStart.bat`       | Diagnostic — shows task state + last run code  |
| `Uninstall-AutoStart.bat`   | Clean removal of the task                      |

### Rollout recommendation

1. Copy `C:\RDPPrintQueue\` to the client.
2. Install SumatraPDF machine-wide if possible.
3. Log on as the operator who will print, open an **elevated** command
   prompt (only so `schtasks` can write to the task store — the task itself
   still runs non-elevated), and run `Install-AutoStart.bat`.
4. Run `Check-AutoStart.bat` — confirm `Status: Ready`, `Run As User` is the
   operator, `Scheduled Task State: Enabled`.
5. Run `Test-RdpPdfAutoPrint.bat` once to validate printing end-to-end and
   review the latest file under `C:\RDPPrintQueue\logs\`.
6. Log off / log on (or reboot). Re-run `Check-AutoStart.bat` — `Last Run
   Time` should be populated and `Last Result` should be `0`.
7. (Optional fallback) If for any reason the Scheduled Task cannot be used,
   drop a shortcut to `Start-RdpPdfAutoPrint.bat` into
   `shell:startup`. This is strictly a backup path.

## Deployment checklist per client PC

1. Copy the whole `RDPPrintQueue` folder to `C:\RDPPrintQueue`.
2. Install SumatraPDF machine-wide
   (`C:\Program Files\SumatraPDF\SumatraPDF.exe`). Per-user installs still
   work thanks to the LOCALAPPDATA fallback, but machine-wide is strongly
   preferred.
3. Confirm the target printer is visible to the logged-on user
   (`Get-Printer` or Control Panel → Devices and Printers).
4. Smoke test:
   ```
   C:\RDPPrintQueue\Test-RdpPdfAutoPrint.bat
   ```
   This prints `C:\RDPPrintQueue\failed\SaleReport545.pdf` and leaves the
   diagnostics window open.
5. Review the newest file under `C:\RDPPrintQueue\logs\`. Confirm
   `whoami`, `Resolved Sumatra`, and `Resolved Printer` are what you expect.
6. Register the Scheduled Task (see above) and reboot / re-logon to
   validate that the watcher comes up automatically.

## Useful command-line variants

```powershell
# One pass then exit (for "run every minute" scheduled task style)
powershell -NoProfile -ExecutionPolicy Bypass -File .\RdpPdfAutoPrint.ps1 -Once

# Dry run (resolves everything, logs, does not print)
powershell -NoProfile -ExecutionPolicy Bypass -File .\RdpPdfAutoPrint.ps1 -DryRun

# Force a specific printer (bypasses default-printer ambiguity)
powershell -NoProfile -ExecutionPolicy Bypass -File .\RdpPdfAutoPrint.ps1 `
    -PrinterName "HP LaserJet Pro M404"

# Smoke-print one file
powershell -NoProfile -ExecutionPolicy Bypass -File .\RdpPdfAutoPrint.ps1 `
    -TestFile "C:\RDPPrintQueue\failed\SaleReport545.pdf"
```

## Exit codes

| Code | Meaning                                             |
|------|-----------------------------------------------------|
| 0    | Success (watcher exited cleanly or TestMode OK)     |
| 2    | Sumatra or printer could not be resolved at startup |
| 3    | `-TestFile` not found                               |
| 4    | TestMode timed out                                  |
| 5    | TestMode: Sumatra returned non-zero                 |
| 6    | TestMode: unhandled exception                       |
