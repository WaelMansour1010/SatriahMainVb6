<#
.SYNOPSIS
    RDP PDF Auto-Print Watcher.

.DESCRIPTION
    Watches C:\RDPPrintQueue\inbox for PDFs and prints them via SumatraPDF.
    Designed as a fallback for RDP printer redirection.

    Folder layout (under -QueueRoot):
        inbox        - new PDFs arrive here
        processing   - held while printing
        archive      - printed successfully
        failed       - printing failed after all retries
        logs         - rolling daily log files
        state        - printed-jobs.csv (dedup), watcher.lock

    Printer selection order (when -PrinterName is NOT specified):
        1. Redirected printer that is also the default
        2. Any redirected printer   (name contains "(redirected N)" or Port starts TS/Client)
        3. Default printer that is not a virtual/excluded printer
        4. First non-virtual/non-excluded printer
        5. WScript.Network fallback (odd-indexed values = actual printer names)
    Virtual printers excluded: Microsoft Print to PDF, XPS Document Writer,
        Fax, OneNote variants, Adobe PDF, PDFCreator, CutePDF Writer.

    Printer is re-resolved before every print attempt, not only at startup,
    so redirected-printer renames across reconnects are handled automatically.
    In watcher mode a missing printer does NOT stop the process; it stays
    alive, logs a throttled warning, and recovers when a printer appears.

.PARAMETER QueueRoot
    Root of the queue. Default: C:\RDPPrintQueue

.PARAMETER SumatraPath
    Optional explicit path to SumatraPDF.exe.

.PARAMETER PrinterName
    Optional explicit printer name.  If the exact name is not found,
    a loose prefix-match is tried to tolerate "(redirected N)" suffixes.

.PARAMETER PollIntervalSeconds
    How often to poll inbox. Default: 3.

.PARAMETER StableChecks
    Consecutive stable-size/timestamp samples before a file is considered
    ready. Default: 2.

.PARAMETER StableDelayMs
    Milliseconds between stability samples. Default: 750.

.PARAMETER MaxPrintRetries
    Retries per file on failure. Default: 3.

.PARAMETER PrintTimeoutSeconds
    Hard timeout per Sumatra invocation. Default: 60.

.PARAMETER TestFile
    Print this single file, log full diagnostics, then exit.
    Printer must be available; fails strictly (no retry loop).

.PARAMETER DryRun
    Resolve everything and log what would happen, but do not print.

.PARAMETER Once
    One inbox pass then exit (useful with a "run every N minutes" trigger).
#>

[CmdletBinding()]
param(
    [string] $QueueRoot           = 'C:\RDPPrintQueue',
    [string] $SumatraPath,
    [string] $PrinterName,
    [int]    $PollIntervalSeconds = 3,
    [int]    $StableChecks        = 2,
    [int]    $StableDelayMs       = 750,
    [int]    $MaxPrintRetries     = 3,
    [int]    $PrintTimeoutSeconds = 60,
    [string] $TestFile,
    [switch] $DryRun,
    [switch] $Once
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
$Script:Paths = [ordered]@{
    Root       = $QueueRoot
    Inbox      = Join-Path $QueueRoot 'inbox'
    Processing = Join-Path $QueueRoot 'processing'
    Archive    = Join-Path $QueueRoot 'archive'
    Failed     = Join-Path $QueueRoot 'failed'
    Logs       = Join-Path $QueueRoot 'logs'
    State      = Join-Path $QueueRoot 'state'
}
$Script:PrintedJobsDb = Join-Path $Script:Paths.State 'printed-jobs.csv'
$Script:LockFile      = Join-Path $Script:Paths.State 'watcher.lock'
$Script:LogFile       = Join-Path $Script:Paths.Logs  ("auto-print-{0:yyyy-MM-dd}.log" -f (Get-Date))

# Throttle: suppress repeated "no printer" warnings to once per 60 s.
$Script:LastNoPrinterWarnTicks = 0

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG','DIAG')]
        [string]$Level = 'INFO'
    )
    $line = '{0} [{1,-5}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Level, $Message
    try { Add-Content -LiteralPath $Script:LogFile -Value $line -Encoding UTF8 } catch { }
    $color = switch ($Level) {
        'ERROR' { 'Red' }   'WARN' { 'Yellow' }
        'DEBUG' { 'DarkGray' }  'DIAG' { 'Cyan' }
        default { 'Gray' }
    }
    Write-Host $line -ForegroundColor $color
}

function Initialize-QueueFolders {
    foreach ($p in $Script:Paths.Values) {
        if (-not (Test-Path -LiteralPath $p)) {
            New-Item -ItemType Directory -Path $p -Force | Out-Null
        }
    }
    if (-not (Test-Path -LiteralPath $Script:PrintedJobsDb)) {
        'Timestamp,FileName,Sha256,SizeBytes,Printer,Result' |
            Out-File -LiteralPath $Script:PrintedJobsDb -Encoding UTF8
    }
}

# ---------------------------------------------------------------------------
# Startup diagnostics (runs once at launch)
# ---------------------------------------------------------------------------
function Write-StartupDiagnostics {
    param([string]$ResolvedSumatra, [string]$ResolvedPrinter)
    Write-Log '--- Startup Diagnostics ---' DIAG
    try   { Write-Log ("whoami              : {0}" -f (& whoami 2>&1)) DIAG } catch { Write-Log "whoami              : <n/a>" DIAG }
    Write-Log ("USERNAME            : {0}" -f $env:USERNAME)     DIAG
    Write-Log ("USERPROFILE         : {0}" -f $env:USERPROFILE)  DIAG
    Write-Log ("LOCALAPPDATA        : {0}" -f $env:LOCALAPPDATA) DIAG
    Write-Log ("COMPUTERNAME        : {0}" -f $env:COMPUTERNAME) DIAG
    Write-Log ("SESSIONNAME         : {0}" -f $env:SESSIONNAME)  DIAG
    Write-Log ("PowerShell          : {0}" -f $PSVersionTable.PSVersion) DIAG
    Write-Log ("QueueRoot           : {0}" -f $Script:Paths.Root) DIAG
    $ss = if ([string]::IsNullOrWhiteSpace($ResolvedSumatra)) { '<unresolved>' } else { $ResolvedSumatra }
    $ps = if ([string]::IsNullOrWhiteSpace($ResolvedPrinter)) { '<unresolved>' } else { $ResolvedPrinter }
    Write-Log ("Resolved Sumatra    : {0}" -f $ss) DIAG
    Write-Log ("Resolved Printer    : {0}" -f $ps) DIAG
    Write-PrinterList
    Write-Log '---------------------------' DIAG
}

function Write-PrinterList {
    try {
        $printers = @(Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop)
        if ($printers.Count -eq 0) {
            Write-Log 'Printer list        : <none visible to this user>' DIAG
        } else {
            foreach ($p in $printers) {
                Write-Log ("Printer             : [{0,-40}]  Port={1,-12}  Default={2}  Net={3}" -f `
                    $p.Name, $p.PortName, $p.Default, $p.Network) DIAG
            }
        }
    } catch {
        Write-Log ("Printer list error  : {0}" -f $_.Exception.Message) WARN
    }
}

# ---------------------------------------------------------------------------
# Sumatra discovery  (unchanged except -Quiet flag)
# ---------------------------------------------------------------------------
function Resolve-SumatraPath {
    param([string]$Explicit, [switch]$Quiet)
    $candidates = @()
    if ($Explicit)         { $candidates += $Explicit }
    $candidates += 'C:\Program Files\SumatraPDF\SumatraPDF.exe'
    $candidates += 'C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe'
    if ($env:LOCALAPPDATA) { $candidates += (Join-Path $env:LOCALAPPDATA 'SumatraPDF\SumatraPDF.exe') }

    foreach ($c in $candidates) {
        if ([string]::IsNullOrWhiteSpace($c)) { continue }
        if (Test-Path -LiteralPath $c) {
            if (-not $Quiet) { Write-Log ("Sumatra resolved    : {0}" -f $c) INFO }
            return $c
        }
        if (-not $Quiet) { Write-Log ("Sumatra candidate missing: {0}" -f $c) DEBUG }
    }
    throw "SumatraPDF.exe not found. Install under 'C:\Program Files\SumatraPDF' or pass -SumatraPath."
}

# ---------------------------------------------------------------------------
# Printer helpers
# ---------------------------------------------------------------------------

# Returns all Win32_Printer objects visible to this user, or empty array.
function Get-AllPrinters {
    try {
        $list = @(Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop)
        return $list
    } catch {
        Write-Log ("Win32_Printer query failed: {0}" -f $_.Exception.Message) WARN
        return @()
    }
}

# Filters out virtual / excluded printers that are never physical targets.
$Script:ExcludedPatterns = @(
    '^Microsoft Print to PDF$',
    '^Microsoft XPS Document Writer$',
    '^Fax$',
    'OneNote',
    'Send to OneNote',
    '^Adobe PDF$',
    '^PDFCreator$',
    'CutePDF'
)
function Get-AcceptablePrinters {
    param([object[]]$All)
    if (-not $All -or $All.Count -eq 0) { return @() }
    $All | Where-Object {
        $n = $_.Name
        -not ($Script:ExcludedPatterns | Where-Object { $n -match $_ })
    }
}

# Detects if a Win32_Printer is an RDP-redirected printer.
# Redirected printers carry "(redirected N)" in the name, or their
# port starts with "TS" (Terminal Services) or "Client".
function Test-IsRedirected {
    param([object]$Printer)
    if (-not $Printer) { return $false }
    if ($Printer.Name    -match '\(redirected\s+\d+\)')  { return $true }
    if ($Printer.PortName -match '^(TS|Client)')         { return $true }
    return $false
}

# ---------------------------------------------------------------------------
# Printer resolution  (re-callable at any time during the watcher loop)
# ---------------------------------------------------------------------------
function Resolve-Printer {
    param(
        [string] $Explicit,  # value of -PrinterName param (may be empty)
        [switch] $Quiet      # suppress normal-level log when called inside loop
    )

    $installed = Get-AllPrinters

    # --- Explicit name ---
    if (-not [string]::IsNullOrWhiteSpace($Explicit)) {

        # Exact match first.
        $match = $installed | Where-Object { $_.Name -ieq $Explicit } | Select-Object -First 1
        if ($match) {
            if (-not $Quiet) { Write-Log ("Printer selected    : [{0}]  (explicit exact)" -f $match.Name) INFO }
            return $match.Name
        }

        # Loose prefix match: "HP LaserJet" → "HP LaserJet (redirected 3)"
        # This tolerates RDP renaming where the session appends "(redirected N)".
        $loose = $installed | Where-Object { $_.Name -like ("{0}*" -f $Explicit) } | Select-Object -First 1
        if ($loose) {
            Write-Log ("Printer selected    : [{0}]  (explicit-loose; asked for '{1}')" -f $loose.Name, $Explicit) WARN
            return $loose.Name
        }

        $avail = ($installed | ForEach-Object { $_.Name }) -join '; '
        throw "NOPRINTER: explicit '$Explicit' not found. Visible: $avail"
    }

    # --- Heuristic auto-selection ---
    $candidates = @(Get-AcceptablePrinters -All $installed)

    if ($candidates.Count -eq 0) {
        if ($installed.Count -gt 0) {
            # All visible printers are virtual; use the default as last resort.
            $def = $installed | Where-Object { $_.Default } | Select-Object -First 1
            if ($def) {
                Write-Log ("Printer selected    : [{0}]  (only-virtual fallback)" -f $def.Name) WARN
                return $def.Name
            }
        }
        throw "NOPRINTER: no printers visible to user '$env:USERNAME' on '$env:COMPUTERNAME'."
    }

    # Priority 1 – redirected AND default.
    $pick = $candidates | Where-Object { (Test-IsRedirected $_) -and $_.Default } | Select-Object -First 1
    if ($pick) {
        if (-not $Quiet) { Write-Log ("Printer selected    : [{0}]  (redirected+default)" -f $pick.Name) INFO }
        return $pick.Name
    }

    # Priority 2 – any redirected.
    $pick = $candidates | Where-Object { Test-IsRedirected $_ } | Select-Object -First 1
    if ($pick) {
        if (-not $Quiet) { Write-Log ("Printer selected    : [{0}]  (redirected)" -f $pick.Name) INFO }
        return $pick.Name
    }

    # Priority 3 – default non-virtual.
    $pick = $candidates | Where-Object { $_.Default } | Select-Object -First 1
    if ($pick) {
        if (-not $Quiet) { Write-Log ("Printer selected    : [{0}]  (default non-virtual)" -f $pick.Name) INFO }
        return $pick.Name
    }

    # Priority 4 – first acceptable.
    $pick = $candidates | Select-Object -First 1
    if ($pick) {
        if (-not $Quiet) { Write-Log ("Printer selected    : [{0}]  (first acceptable)" -f $pick.Name) INFO }
        return $pick.Name
    }

    # Priority 5 – WScript.Network fallback.
    # EnumPrinterConnections() returns alternating [port, printerName, port, printerName, ...].
    # Odd-indexed elements (1, 3, 5, …) are the actual printer names.
    try {
        $net  = New-Object -ComObject WScript.Network
        $enum = @($net.EnumPrinterConnections())
        $names = for ($i = 1; $i -lt $enum.Count; $i += 2) { $enum[$i] }
        $names = @($names | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($names.Count -gt 0) {
            Write-Log ("Printer selected    : [{0}]  (WScript.Network fallback)" -f $names[0]) WARN
            return $names[0]
        }
    } catch {
        Write-Log ("WScript.Network fallback failed: {0}" -f $_.Exception.Message) DEBUG
    }

    throw "NOPRINTER: no printer could be resolved for user '$env:USERNAME'."
}

# Throttled "no printer" warning — at most one log entry per 60 s.
function Write-NoPrinterThrottled {
    param([string]$Message)
    $now = [DateTime]::UtcNow.Ticks
    $threshold = [TimeSpan]::FromSeconds(60).Ticks
    if (($now - $Script:LastNoPrinterWarnTicks) -gt $threshold) {
        Write-Log ("Printer unavailable : {0}" -f $Message) WARN
        Write-Log ("Watcher is alive and will keep retrying. Current printer list:") WARN
        Write-PrinterList
        $Script:LastNoPrinterWarnTicks = $now
    }
}

# ---------------------------------------------------------------------------
# File readiness / hashing / dedup  (unchanged)
# ---------------------------------------------------------------------------
function Test-FileReady {
    param([Parameter(Mandatory)][string]$Path)
    $last = $null
    for ($i = 0; $i -lt $StableChecks; $i++) {
        try {
            $fi     = Get-Item -LiteralPath $Path -ErrorAction Stop
            $sample = '{0}|{1}' -f $fi.Length, $fi.LastWriteTimeUtc.Ticks
            if ($null -ne $last -and $sample -ne $last) { return $false }
            $last = $sample
        } catch { return $false }
        Start-Sleep -Milliseconds $StableDelayMs
    }
    try {
        $fs = [System.IO.File]::Open($Path, 'Open', 'Read', 'None')
        $fs.Close(); $fs.Dispose()
        return $true
    } catch { return $false }
}

function Get-FileSha256 { param([Parameter(Mandatory)][string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
}

function Test-AlreadyPrinted {
    param([Parameter(Mandatory)][string]$Sha256, [Parameter(Mandatory)][long]$Size)
    if (-not (Test-Path -LiteralPath $Script:PrintedJobsDb)) { return $false }
    try {
        $rows = Import-Csv -LiteralPath $Script:PrintedJobsDb -ErrorAction Stop
        foreach ($r in $rows) {
            if ($r.Sha256 -eq $Sha256 -and [long]$r.SizeBytes -eq $Size -and $r.Result -eq 'OK') {
                return $true
            }
        }
    } catch { }
    return $false
}

function Add-PrintedJob {
    param([string]$FileName, [string]$Sha256, [long]$Size, [string]$Printer, [string]$Result)
    $row = [pscustomobject]@{
        Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        FileName  = $FileName
        Sha256    = $Sha256
        SizeBytes = $Size
        Printer   = $Printer
        Result    = $Result
    }
    ($row | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1) |
        Add-Content -LiteralPath $Script:PrintedJobsDb -Encoding UTF8
}

# ---------------------------------------------------------------------------
# Clipboard — copy PDF as a file-drop object so the user can Ctrl+V it
# ---------------------------------------------------------------------------
function Set-ClipboardFileDrop {
    param([Parameter(Mandatory)][string]$Path)
    # [System.Windows.Forms.Clipboard] requires an STA apartment thread.
    # Rather than forcing -STA on the whole process, spin a dedicated STA
    # runspace for the clipboard call and dispose it immediately.
    try {
        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.ApartmentState = [System.Threading.ApartmentState]::STA
        $rs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $rs.Open()
        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.Runspace = $rs
        [void]$ps.AddScript({
            param([string]$FilePath)
            Add-Type -AssemblyName System.Windows.Forms
            $col = New-Object System.Collections.Specialized.StringCollection
            [void]$col.Add($FilePath)
            [System.Windows.Forms.Clipboard]::SetFileDropList($col)
        }).AddParameter('FilePath', $Path)
        $async = $ps.BeginInvoke()
        if (-not $async.AsyncWaitHandle.WaitOne(5000)) {
            Write-Log ("Clipboard timed out : {0}" -f (Split-Path -Leaf $Path)) WARN
        } elseif ($ps.HadErrors) {
            $errs = ($ps.Streams.Error | ForEach-Object { $_.ToString() }) -join '; '
            Write-Log ("Clipboard error     : {0}" -f $errs) WARN
        } else {
            Write-Log ("Clipboard set       : {0}" -f (Split-Path -Leaf $Path)) INFO
        }
        $ps.Dispose()
        $rs.Close(); $rs.Dispose()
    } catch {
        # Never allow clipboard failure to abort the print pipeline.
        Write-Log ("Clipboard failed    : {0}" -f $_.Exception.Message) WARN
    }
}

# ---------------------------------------------------------------------------
# Print execution  — logs full per-job context before every attempt
# ---------------------------------------------------------------------------
function Invoke-SumatraPrint {
    param(
        [Parameter(Mandatory)][string]$SumatraExe,
        [Parameter(Mandatory)][string]$Printer,
        [Parameter(Mandatory)][string]$PdfPath,
        [string]$RequestedPrinter     # the -PrinterName param value, for logging
    )

    $sumatraArgs = @('-print-to', $Printer, '-silent', $PdfPath)

    # Per-job diagnostic block — visible in every log line.
    Write-Log '-- job diag --' DIAG
    try { Write-Log ("  identity          : {0}" -f (& whoami 2>&1)) DIAG } catch { }
    Write-Log ("  session           : {0}" -f $env:SESSIONNAME)   DIAG
    Write-Log ("  sumatra           : {0}" -f $SumatraExe)        DIAG
    $reqShow = if ([string]::IsNullOrWhiteSpace($RequestedPrinter)) { '<auto>' } else { $RequestedPrinter }
    Write-Log ("  requested printer : {0}" -f $reqShow)           DIAG
    Write-Log ("  selected printer  : {0}" -f $Printer)           DIAG
    # Brief visible-printer snapshot at the moment of printing.
    try {
        $snap = @(Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop)
        foreach ($p in $snap) {
            Write-Log ("  visible printer   : [{0}]  Port={1}  Def={2}" -f $p.Name, $p.PortName, $p.Default) DIAG
        }
    } catch { Write-Log "  visible printer   : <query failed>" DIAG }
    Write-Log ("  pdf               : {0}" -f $PdfPath)           DIAG
    Write-Log ("  sumatra args      : {0}" -f ($sumatraArgs -join ' | ')) DIAG

    if ($DryRun) {
        Write-Log '  DryRun: skipping Start-Process.' WARN
        return @{ Category = 'DRYRUN'; ExitCode = 0; TimedOut = $false }
    }

    $proc = Start-Process -FilePath $SumatraExe -ArgumentList $sumatraArgs `
                          -PassThru -WindowStyle Hidden

    if (-not $proc.WaitForExit($PrintTimeoutSeconds * 1000)) {
        try { $proc.Kill() } catch { }
        Write-Log ("  result            : TIMEOUT after {0}s" -f $PrintTimeoutSeconds) ERROR
        return @{ Category = 'TIMEOUT'; ExitCode = -1; TimedOut = $true }
    }

    $ec = $proc.ExitCode
    if ($ec -eq 0) {
        Write-Log ("  result            : OK (exit 0)") INFO
        return @{ Category = 'OK'; ExitCode = 0; TimedOut = $false }
    }
    Write-Log ("  result            : SUMATRA_NONZERO (exit {0})" -f $ec) WARN
    return @{ Category = 'SUMATRA_NONZERO'; ExitCode = $ec; TimedOut = $false }
}

# ---------------------------------------------------------------------------
# Per-file pipeline  — re-resolves printer live on every attempt
# ---------------------------------------------------------------------------
function Invoke-PrintPipeline {
    param(
        [Parameter(Mandatory)][string]$SourcePath,
        [Parameter(Mandatory)][string]$SumatraExe
    )

    $fileName = Split-Path -Leaf $SourcePath

    # 1. File stability check.
    if (-not (Test-FileReady -Path $SourcePath)) {
        Write-Log ("Not ready yet       : {0}" -f $fileName) DEBUG
        return
    }

    # 2. Resolve printer live — before touching the file.
    #    If no printer, leave file in inbox; watcher loop will retry next poll.
    $livePrinter = $null
    try {
        $livePrinter = Resolve-Printer -Explicit $PrinterName -Quiet
    } catch {
        $msg = $_.Exception.Message
        if ($msg -like 'NOPRINTER:*') {
            Write-NoPrinterThrottled -Message ($msg -replace '^NOPRINTER:\s*','')
        } else {
            Write-Log ("Printer resolution error: {0}" -f $msg) ERROR
        }
        return   # leave file in inbox, retry next poll
    }

    # 3. Hash + dedup BEFORE moving.
    $size = (Get-Item -LiteralPath $SourcePath).Length
    $hash = Get-FileSha256 -Path $SourcePath

    if (Test-AlreadyPrinted -Sha256 $hash -Size $size) {
        Write-Log ("Duplicate (already OK), archiving: {0}" -f $fileName) WARN
        $dest = Join-Path $Script:Paths.Archive ("DUP_{0:yyyyMMdd_HHmmss}_{1}" -f (Get-Date), $fileName)
        Move-Item -LiteralPath $SourcePath -Destination $dest -Force
        Add-PrintedJob -FileName $fileName -Sha256 $hash -Size $size -Printer $livePrinter -Result 'DUPLICATE'
        return
    }

    # 4. Move to processing.
    $stamp    = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
    $procName = '{0}_{1}' -f $stamp, $fileName
    $procPath = Join-Path $Script:Paths.Processing $procName
    try {
        Move-Item -LiteralPath $SourcePath -Destination $procPath -Force
    } catch {
        Write-Log ("Failed to move to processing [{0}]: {1}" -f $fileName, $_.Exception.Message) ERROR
        return
    }

    # 4b. Clipboard backup — BEFORE the print attempt, as a safety net.
    # If printing fails or times out, the user can already Ctrl+V the PDF
    # into Explorer or a print dialog and print it manually. Clipboard set
    # is wrapped so its failure never blocks the print pipeline.
    Set-ClipboardFileDrop -Path $procPath

    # 5. Print with retries — re-resolve printer on each attempt.
    $attempt   = 0
    $succeeded = $false
    $lastCat   = 'UNKNOWN'
    $lastErr   = ''

    while ($attempt -lt $MaxPrintRetries -and -not $succeeded) {
        $attempt++

        # Re-resolve printer on attempt > 1 (first attempt already resolved above).
        if ($attempt -gt 1) {
            try {
                $livePrinter = Resolve-Printer -Explicit $PrinterName
            } catch {
                $lastCat = 'NOPRINTER'
                $lastErr = $_.Exception.Message -replace '^NOPRINTER:\s*',''
                Write-Log ("Attempt {0}/{1} NOPRINTER: {2}" -f $attempt, $MaxPrintRetries, $lastErr) WARN
                $backoff = [math]::Min(30, [math]::Pow(2, $attempt))
                Write-Log ("Backing off {0}s." -f $backoff) DEBUG
                Start-Sleep -Seconds $backoff
                continue
            }
        }

        Write-Log ("Print attempt {0}/{1} : {2}  → [{3}]" -f $attempt, $MaxPrintRetries, $procName, $livePrinter) INFO

        try {
            $r = Invoke-SumatraPrint -SumatraExe $SumatraExe `
                                     -Printer    $livePrinter `
                                     -PdfPath    $procPath `
                                     -RequestedPrinter $PrinterName

            if ($r.Category -eq 'OK' -or $r.Category -eq 'DRYRUN') {
                $succeeded = $true; break
            }
            $lastCat = $r.Category
            $lastErr = "ExitCode=$($r.ExitCode)"
            Write-Log ("Attempt {0} FAILED [{1}]: {2}" -f $attempt, $lastCat, $lastErr) WARN
        } catch {
            $lastCat = 'EXCEPTION'
            $lastErr = $_.Exception.Message
            Write-Log ("Attempt {0} EXCEPTION: {1}" -f $attempt, $lastErr) ERROR
        }

        if (-not $succeeded -and $attempt -lt $MaxPrintRetries) {
            $backoff = [math]::Min(30, [math]::Pow(2, $attempt))
            Write-Log ("Backing off {0}s before retry." -f $backoff) DEBUG
            Start-Sleep -Seconds $backoff
        }
    }

    # 6. Archive or fail.
    if ($succeeded) {
        $dest = Join-Path $Script:Paths.Archive $procName
        Move-Item -LiteralPath $procPath -Destination $dest -Force
        Add-PrintedJob -FileName $fileName -Sha256 $hash -Size $size -Printer $livePrinter -Result 'OK'
        Write-Log ("Printed OK          : {0}" -f $fileName) INFO
    } else {
        $dest = Join-Path $Script:Paths.Failed $procName
        Move-Item -LiteralPath $procPath -Destination $dest -Force
        Add-PrintedJob -FileName $fileName -Sha256 $hash -Size $size -Printer $livePrinter -Result ("FAIL:{0}:{1}" -f $lastCat, $lastErr)
        Write-Log ("Print FAILED        : {0}  [{1}] {2}" -f $fileName, $lastCat, $lastErr) ERROR
    }

    # Re-point the clipboard at the final stable path ($dest — either archive
    # or failed). The earlier 4b call pointed at the transient processing path;
    # since Windows file-drop clipboard stores paths (not bytes), that path
    # becomes stale once step 6 above moves the file. This second call ensures
    # the clipboard always references a file that still exists on disk, no
    # matter when the user chooses to paste.
    Set-ClipboardFileDrop -Path $dest
}

# ---------------------------------------------------------------------------
# Single-instance lock  (unchanged)
# ---------------------------------------------------------------------------
function Enter-SingletonLock {
    try {
        $stream = [System.IO.File]::Open($Script:LockFile, 'OpenOrCreate', 'ReadWrite', 'None')
    } catch {
        throw "Another watcher instance is already running (lock: $Script:LockFile)."
    }
    $bytes = [Text.Encoding]::UTF8.GetBytes(("{0}|{1}|{2}" -f $PID, $env:COMPUTERNAME, (Get-Date -Format 'o')))
    $stream.SetLength(0); $stream.Write($bytes, 0, $bytes.Length); $stream.Flush()
    return $stream
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Initialize-QueueFolders

Write-Log '==========================================================' INFO
Write-Log ('RdpPdfAutoPrint starting (PID {0})' -f $PID) INFO
Write-Log '==========================================================' INFO

# Sumatra must exist before we can do anything.
$resolvedSumatra = $null
try {
    $resolvedSumatra = Resolve-SumatraPath -Explicit $SumatraPath
} catch {
    Write-StartupDiagnostics -ResolvedSumatra $null -ResolvedPrinter $null
    Write-Log ("Sumatra not found: {0}" -f $_.Exception.Message) ERROR
    exit 2
}

# Try printer at startup just for the diagnostics block.
# In watcher mode a failure here is NOT fatal; the loop re-tries per file.
$startupPrinter = $null
try {
    $startupPrinter = Resolve-Printer -Explicit $PrinterName
} catch {
    Write-Log ("Printer not available at startup: {0}" -f ($_.Exception.Message -replace '^NOPRINTER:\s*','')) WARN
    if ($TestFile) {
        # TestMode needs a printer right now.
        Write-StartupDiagnostics -ResolvedSumatra $resolvedSumatra -ResolvedPrinter $null
        Write-Log 'TestMode requires a printer at startup. Aborting.' ERROR
        exit 2
    }
    Write-Log 'Watcher will keep running and retry printer resolution per file.' WARN
}
Write-StartupDiagnostics -ResolvedSumatra $resolvedSumatra -ResolvedPrinter $startupPrinter

# ---- Test mode ----
if ($TestFile) {
    Write-Log ('TEST MODE: {0}' -f $TestFile) WARN
    if (-not (Test-Path -LiteralPath $TestFile)) {
        Write-Log ('TestFile not found: {0}' -f $TestFile) ERROR
        exit 3
    }
    try {
        $r = Invoke-SumatraPrint -SumatraExe $resolvedSumatra `
                                 -Printer    $startupPrinter `
                                 -PdfPath    $TestFile `
                                 -RequestedPrinter $PrinterName
        switch ($r.Category) {
            'DRYRUN'          { Write-Log 'TEST MODE (DryRun) OK.'          INFO;  exit 0 }
            'OK'              { Write-Log 'TEST MODE success.'               INFO;  exit 0 }
            'TIMEOUT'         { Write-Log 'TEST MODE timed out.'             ERROR; exit 4 }
            'SUMATRA_NONZERO' { Write-Log ("TEST MODE failed, exit={0}" -f $r.ExitCode) ERROR; exit 5 }
            default           { Write-Log ("TEST MODE unexpected: {0}" -f $r.Category)  ERROR; exit 5 }
        }
    } catch {
        Write-Log ('TEST MODE exception: {0}' -f $_.Exception.Message) ERROR
        exit 6
    }
}

# ---- Watcher loop ----
$lock = $null
try {
    $lock = Enter-SingletonLock
    Write-Log ('Watching            : {0}' -f $Script:Paths.Inbox) INFO
    Write-Log ('DryRun              : {0}' -f $DryRun.IsPresent)   INFO
    Write-Log ('Once                : {0}' -f $Once.IsPresent)     INFO

    while ($true) {
        try {
            $pdfs = @(Get-ChildItem -LiteralPath $Script:Paths.Inbox -Filter '*.pdf' -File -ErrorAction SilentlyContinue |
                      Sort-Object LastWriteTimeUtc)
            foreach ($f in $pdfs) {
                try {
                    Invoke-PrintPipeline -SourcePath $f.FullName -SumatraExe $resolvedSumatra
                } catch {
                    Write-Log ("Pipeline exception [{0}]: {1}" -f $f.Name, $_.Exception.Message) ERROR
                }
            }
        } catch {
            Write-Log ("Inbox scan error: {0}" -f $_.Exception.Message) ERROR
        }

        if ($Once) { break }
        Start-Sleep -Seconds $PollIntervalSeconds
    }
} finally {
    if ($lock) { try { $lock.Close(); $lock.Dispose() } catch { } }
    try { if (Test-Path -LiteralPath $Script:LockFile) { Remove-Item -LiteralPath $Script:LockFile -Force } } catch { }
    Write-Log 'RdpPdfAutoPrint stopped.' INFO
}
