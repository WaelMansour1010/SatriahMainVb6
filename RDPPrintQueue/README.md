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
   Program Files (x86) → `%LOCALAPPDATA%`. Machine-wide installs are preferred,
   so Sumatra is found regardless of which user runs the watcher.
2. **`Resolve-Printer`** — validates `-PrinterName` or resolves the true
   Windows default printer and passes it **explicitly** to Sumatra with
   `-print-to "<name>"`. No more ambiguity around `-print-to-default`.
3. **Diagnostics block** — logs `whoami`, `USERNAME`, `USERPROFILE`,
   `COMPUTERNAME`, resolved Sumatra, resolved printer, and the full installed
   printer list on every start. Makes "which user / which printer" obvious in
   the log.
4. **Test Mode (`-TestFile`)** — prints one PDF and exits with a specific exit
   code, so the test `.bat` is a pure smoke test.
5. **DryRun mode (`-DryRun`)** — resolves everything, logs the exact
   Sumatra invocation, but does not print.
6. **Clean separation** — watcher loop, file readiness, printer discovery,
   Sumatra discovery, print execution, and logging each live in their own
   function.
7. **Queue + dedup preserved** — `printed-jobs.csv` keyed on SHA256 + size.

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
