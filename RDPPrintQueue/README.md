# RDP PDF Auto-Print Watcher

Local Windows fallback for RDP printing: PDFs that the VB6 server app drops on
the client (via an RDP-redirected folder) are printed automatically with
SumatraPDF.

## Folder layout

```
C:\RDPPrintQueue\
    RdpPdfAutoPrint.ps1         <- main script
    Start-RdpPdfAutoPrint.bat   <- production launcher
    Test-RdpPdfAutoPrint.bat    <- prints a single known PDF and exits
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
