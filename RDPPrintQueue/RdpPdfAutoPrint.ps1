<#
.SYNOPSIS
    RDP PDF Auto-Print Watcher.

.DESCRIPTION
    Watches a local queue folder (inbox) for PDFs dropped by a remote side of an
    RDP session and prints them using SumatraPDF. Designed as a fallback for RDP
    printer redirection, keeping queue/archive/failed/log folders for audit.

    Folder layout (under -QueueRoot, default C:\RDPPrintQueue):
        inbox        - new PDFs arrive here
        processing   - files moved here before printing
        archive      - printed successfully
        failed       - failed prints
        logs         - rolling log files
        state        - printed-jobs.csv (dedup db), watcher.lock

.PARAMETER QueueRoot
    Root of the queue. Default: C:\RDPPrintQueue

.PARAMETER SumatraPath
    Optional explicit path to SumatraPDF.exe. If omitted, Resolve-SumatraPath
    searches Program Files, Program Files (x86), then LOCALAPPDATA.

.PARAMETER PrinterName
    Optional explicit printer name. If omitted, the current user's default
    printer is resolved and used explicitly (-print-to "Name"), which is more
    reliable than -print-to-default under service/scheduled-task contexts.

.PARAMETER PollIntervalSeconds
    How often to poll inbox. Default: 3.

.PARAMETER StableChecks
    Number of consecutive identical size/timestamp samples required before a
    file is considered stable. Default: 2.

.PARAMETER StableDelayMs
    Delay between stability samples in milliseconds. Default: 750.

.PARAMETER MaxPrintRetries
    Retries per file on print failure. Default: 3.

.PARAMETER PrintTimeoutSeconds
    Hard timeout for a single Sumatra invocation. Default: 60.

.PARAMETER TestFile
    If set, prints that single file, logs diagnostics, then exits. Ignores the
    watcher loop and queue folders.

.PARAMETER DryRun
    Resolve Sumatra/Printer and log everything, but do not actually print.

.PARAMETER Once
    Run one pass over inbox and exit. Useful for Scheduled Task "on interval".

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File RdpPdfAutoPrint.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File RdpPdfAutoPrint.ps1 -TestFile "C:\RDPPrintQueue\failed\SaleReport545.pdf"

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File RdpPdfAutoPrint.ps1 -DryRun
#>

[CmdletBinding()]
param(
    [string]   $QueueRoot           = 'C:\RDPPrintQueue',
    [string]   $SumatraPath,
    [string]   $PrinterName,
    [int]      $PollIntervalSeconds = 3,
    [int]      $StableChecks        = 2,
    [int]      $StableDelayMs       = 750,
    [int]      $MaxPrintRetries     = 3,
    [int]      $PrintTimeoutSeconds = 60,
    [string]   $TestFile,
    [switch]   $DryRun,
    [switch]   $Once
)

# ---------------------------------------------------------------------------
# Strict mode + globals
# ---------------------------------------------------------------------------
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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
    try {
        # Append is safe across concurrent runs; the lock file prevents multiple loops anyway.
        Add-Content -LiteralPath $Script:LogFile -Value $line -Encoding UTF8
    } catch {
        # Logging must never take down the watcher.
    }
    $color = switch ($Level) {
        'ERROR' { 'Red' }
        'WARN'  { 'Yellow' }
        'DEBUG' { 'DarkGray' }
        'DIAG'  { 'Cyan' }
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
# Diagnostics
# ---------------------------------------------------------------------------
function Write-Diagnostics {
    param(
        [string]$ResolvedSumatra,
        [string]$ResolvedPrinter
    )
    Write-Log '--- Diagnostics ---' DIAG
    try { Write-Log ("whoami              : {0}" -f (whoami)) DIAG } catch { Write-Log "whoami              : <n/a>" DIAG }
    Write-Log ("USERNAME            : {0}" -f $env:USERNAME)        DIAG
    Write-Log ("USERPROFILE         : {0}" -f $env:USERPROFILE)     DIAG
    Write-Log ("LOCALAPPDATA        : {0}" -f $env:LOCALAPPDATA)    DIAG
    Write-Log ("COMPUTERNAME        : {0}" -f $env:COMPUTERNAME)    DIAG
    Write-Log ("SESSIONNAME         : {0}" -f $env:SESSIONNAME)     DIAG
    Write-Log ("PowerShell          : {0}" -f $PSVersionTable.PSVersion) DIAG
    Write-Log ("QueueRoot           : {0}" -f $Script:Paths.Root)   DIAG
    $sumShow = if ([string]::IsNullOrWhiteSpace($ResolvedSumatra)) { '<unresolved>' } else { $ResolvedSumatra }
    $prtShow = if ([string]::IsNullOrWhiteSpace($ResolvedPrinter)) { '<unresolved>' } else { $ResolvedPrinter }
    Write-Log ("Resolved Sumatra    : {0}" -f $sumShow) DIAG
    Write-Log ("Resolved Printer    : {0}" -f $prtShow) DIAG
    try {
        $printers = Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop
        if ($printers) {
            foreach ($p in $printers) {
                Write-Log ("Printer             : {0}  Default={1}  Network={2}  PortName={3}" -f `
                    $p.Name, $p.Default, $p.Network, $p.PortName) DIAG
            }
        } else {
            Write-Log 'Printer             : <none installed>' DIAG
        }
    } catch {
        Write-Log ("Printer enumeration failed: {0}" -f $_.Exception.Message) WARN
    }
    Write-Log '-------------------' DIAG
}

# ---------------------------------------------------------------------------
# Sumatra discovery
# ---------------------------------------------------------------------------
function Resolve-SumatraPath {
    param([string]$Explicit)

    $candidates = @()
    if ($Explicit)         { $candidates += $Explicit }
    $candidates += 'C:\Program Files\SumatraPDF\SumatraPDF.exe'
    $candidates += 'C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe'
    if ($env:LOCALAPPDATA) { $candidates += (Join-Path $env:LOCALAPPDATA 'SumatraPDF\SumatraPDF.exe') }

    foreach ($c in $candidates) {
        if ([string]::IsNullOrWhiteSpace($c)) { continue }
        if (Test-Path -LiteralPath $c) {
            Write-Log ("Sumatra resolved    : {0}" -f $c) INFO
            return $c
        } else {
            Write-Log ("Sumatra candidate missing: {0}" -f $c) DEBUG
        }
    }
    throw "SumatraPDF.exe not found. Install Sumatra under 'C:\Program Files\SumatraPDF' or pass -SumatraPath."
}

# ---------------------------------------------------------------------------
# Printer discovery
# ---------------------------------------------------------------------------
function Resolve-Printer {
    param([string]$Explicit)

    $installed = @()
    try {
        $installed = Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop
    } catch {
        throw "Failed to enumerate printers via Win32_Printer: $($_.Exception.Message)"
    }

    if ($Explicit) {
        $match = $installed | Where-Object { $_.Name -ieq $Explicit } | Select-Object -First 1
        if (-not $match) {
            $names = ($installed | ForEach-Object { $_.Name }) -join ', '
            throw "PrinterName '$Explicit' not found. Installed: $names"
        }
        Write-Log ("Printer resolved    : {0} (explicit)" -f $match.Name) INFO
        return $match.Name
    }

    $default = $installed | Where-Object { $_.Default -eq $true } | Select-Object -First 1
    if (-not $default) {
        # Fallback: WScript.Network default (works under interactive user token).
        try {
            $net = New-Object -ComObject WScript.Network
            $name = $net.EnumPrinterConnections() | Where-Object { $_ } | Select-Object -First 1
            if ($name) {
                Write-Log ("Printer resolved    : {0} (WScript.Network fallback)" -f $name) WARN
                return $name
            }
        } catch { }
        $names = ($installed | ForEach-Object { $_.Name }) -join ', '
        throw "No default printer for user '$env:USERNAME' on '$env:COMPUTERNAME'. Installed: $names. Set a default printer or pass -PrinterName."
    }
    Write-Log ("Printer resolved    : {0} (default)" -f $default.Name) INFO
    return $default.Name
}

# ---------------------------------------------------------------------------
# File readiness / hashing
# ---------------------------------------------------------------------------
function Test-FileReady {
    param([Parameter(Mandatory)][string]$Path)
    # A file is "ready" when two consecutive samples report identical length and
    # LastWriteTime, AND we can open it exclusively (no other writer has it).
    $last = $null
    for ($i = 0; $i -lt $StableChecks; $i++) {
        try {
            $fi = Get-Item -LiteralPath $Path -ErrorAction Stop
            $sample = '{0}|{1}' -f $fi.Length, $fi.LastWriteTimeUtc.Ticks
            if ($null -ne $last -and $sample -ne $last) { return $false }
            $last = $sample
        } catch {
            return $false
        }
        Start-Sleep -Milliseconds $StableDelayMs
    }
    try {
        $fs = [System.IO.File]::Open($Path, 'Open', 'Read', 'None')
        $fs.Close(); $fs.Dispose()
        return $true
    } catch {
        return $false
    }
}

function Get-FileSha256 {
    param([Parameter(Mandatory)][string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
}

function Test-AlreadyPrinted {
    param(
        [Parameter(Mandatory)][string]$Sha256,
        [Parameter(Mandatory)][long]  $Size
    )
    if (-not (Test-Path -LiteralPath $Script:PrintedJobsDb)) { return $false }
    try {
        $rows = Import-Csv -LiteralPath $Script:PrintedJobsDb -ErrorAction Stop
    } catch { return $false }
    foreach ($r in $rows) {
        if ($r.Sha256 -eq $Sha256 -and [long]$r.SizeBytes -eq $Size -and $r.Result -eq 'OK') {
            return $true
        }
    }
    return $false
}

function Add-PrintedJob {
    param(
        [string]$FileName,[string]$Sha256,[long]$Size,[string]$Printer,[string]$Result
    )
    $row = [pscustomobject]@{
        Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        FileName  = $FileName
        Sha256    = $Sha256
        SizeBytes = $Size
        Printer   = $Printer
        Result    = $Result
    }
    # CSV append without re-emitting the header.
    ($row | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1) |
        Add-Content -LiteralPath $Script:PrintedJobsDb -Encoding UTF8
}

# ---------------------------------------------------------------------------
# Print execution
# ---------------------------------------------------------------------------
function Invoke-SumatraPrint {
    param(
        [Parameter(Mandatory)][string]$SumatraExe,
        [Parameter(Mandatory)][string]$Printer,
        [Parameter(Mandatory)][string]$PdfPath
    )

    # Build args as array -> Start-Process quotes each element safely,
    # so spaces in printer names or PDF paths are handled correctly.
    $sumatraArgs = @('-print-to', $Printer, '-silent', $PdfPath)

    Write-Log ("Sumatra exe         : {0}" -f $SumatraExe)          DIAG
    Write-Log ("Sumatra args        : {0}" -f ($sumatraArgs -join ' | ')) DIAG
    Write-Log ("Target PDF          : {0}" -f $PdfPath)             DIAG

    if ($DryRun) {
        Write-Log 'DryRun enabled - skipping actual Start-Process.' WARN
        return @{ ExitCode = 0; TimedOut = $false; Dry = $true }
    }

    $proc = Start-Process -FilePath $SumatraExe `
                          -ArgumentList $sumatraArgs `
                          -PassThru -WindowStyle Hidden

    if (-not $proc.WaitForExit($PrintTimeoutSeconds * 1000)) {
        try { $proc.Kill() } catch { }
        Write-Log ("Sumatra timed out after {0}s - killed." -f $PrintTimeoutSeconds) ERROR
        return @{ ExitCode = -1; TimedOut = $true; Dry = $false }
    }

    return @{ ExitCode = $proc.ExitCode; TimedOut = $false; Dry = $false }
}

# ---------------------------------------------------------------------------
# Per-file pipeline
# ---------------------------------------------------------------------------
function Invoke-PrintPipeline {
    param(
        [Parameter(Mandatory)][string]$SourcePath,
        [Parameter(Mandatory)][string]$SumatraExe,
        [Parameter(Mandatory)][string]$Printer
    )

    $fileName = Split-Path -Leaf $SourcePath

    # 1. Wait until the file is stable (RDP copy may still be streaming).
    if (-not (Test-FileReady -Path $SourcePath)) {
        Write-Log ("Not ready yet       : {0}" -f $fileName) DEBUG
        return
    }

    # 2. Hash + dedupe BEFORE moving, so that a duplicate re-drop from the
    #    server doesn't produce a second archive copy.
    $size = (Get-Item -LiteralPath $SourcePath).Length
    $hash = Get-FileSha256 -Path $SourcePath

    if (Test-AlreadyPrinted -Sha256 $hash -Size $size) {
        Write-Log ("Duplicate (already printed OK), archiving: {0}" -f $fileName) WARN
        $dupTarget = Join-Path $Script:Paths.Archive ("DUP_{0:yyyyMMdd_HHmmss}_{1}" -f (Get-Date), $fileName)
        Move-Item -LiteralPath $SourcePath -Destination $dupTarget -Force
        Add-PrintedJob -FileName $fileName -Sha256 $hash -Size $size -Printer $Printer -Result 'DUPLICATE'
        return
    }

    # 3. Move to processing with a unique prefix so concurrent drops never clash.
    $stamp      = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
    $procName   = '{0}_{1}' -f $stamp, $fileName
    $procPath   = Join-Path $Script:Paths.Processing $procName
    try {
        Move-Item -LiteralPath $SourcePath -Destination $procPath -Force
    } catch {
        Write-Log ("Failed to move to processing ({0}): {1}" -f $fileName, $_.Exception.Message) ERROR
        return
    }

    # 4. Print with retries.
    $attempt   = 0
    $succeeded = $false
    $lastErr   = ''
    while ($attempt -lt $MaxPrintRetries -and -not $succeeded) {
        $attempt++
        Write-Log ("Print attempt {0}/{1} : {2}" -f $attempt, $MaxPrintRetries, $procName) INFO
        try {
            $r = Invoke-SumatraPrint -SumatraExe $SumatraExe -Printer $Printer -PdfPath $procPath
            if ($r.Dry) {
                $succeeded = $true; break
            }
            if (-not $r.TimedOut -and $r.ExitCode -eq 0) {
                Write-Log ("Sumatra exit code   : 0 (success)") INFO
                $succeeded = $true; break
            }
            $lastErr = "ExitCode=$($r.ExitCode) TimedOut=$($r.TimedOut)"
            Write-Log ("Attempt {0} failed   : {1}" -f $attempt, $lastErr) WARN
        } catch {
            $lastErr = $_.Exception.Message
            Write-Log ("Attempt {0} exception: {1}" -f $attempt, $lastErr) WARN
        }
        if (-not $succeeded -and $attempt -lt $MaxPrintRetries) {
            $backoff = [math]::Min(30, [math]::Pow(2, $attempt))
            Write-Log ("Backing off {0}s before retry." -f $backoff) DEBUG
            Start-Sleep -Seconds $backoff
        }
    }

    # 5. Outcome: archive or failed.
    if ($succeeded) {
        $dest = Join-Path $Script:Paths.Archive $procName
        Move-Item -LiteralPath $procPath -Destination $dest -Force
        Add-PrintedJob -FileName $fileName -Sha256 $hash -Size $size -Printer $Printer -Result 'OK'
        Write-Log ("Printed OK          : {0}" -f $fileName) INFO
    } else {
        $dest = Join-Path $Script:Paths.Failed $procName
        Move-Item -LiteralPath $procPath -Destination $dest -Force
        Add-PrintedJob -FileName $fileName -Sha256 $hash -Size $size -Printer $Printer -Result "FAIL:$lastErr"
        Write-Log ("Print FAILED        : {0} ({1})" -f $fileName, $lastErr) ERROR
    }
}

# ---------------------------------------------------------------------------
# Single-instance lock
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

$resolvedSumatra = $null
$resolvedPrinter = $null
try {
    $resolvedSumatra = Resolve-SumatraPath -Explicit $SumatraPath
    $resolvedPrinter = Resolve-Printer     -Explicit $PrinterName
} catch {
    Write-Diagnostics -ResolvedSumatra $resolvedSumatra -ResolvedPrinter $resolvedPrinter
    Write-Log ("Startup resolution failed: {0}" -f $_.Exception.Message) ERROR
    exit 2
}
Write-Diagnostics -ResolvedSumatra $resolvedSumatra -ResolvedPrinter $resolvedPrinter

# ---- Test mode: print single file and exit ----
if ($TestFile) {
    Write-Log ('TEST MODE: {0}' -f $TestFile) WARN
    if (-not (Test-Path -LiteralPath $TestFile)) {
        Write-Log ('TestFile not found: {0}' -f $TestFile) ERROR
        exit 3
    }
    try {
        $r = Invoke-SumatraPrint -SumatraExe $resolvedSumatra -Printer $resolvedPrinter -PdfPath $TestFile
        if ($r.Dry)                 { Write-Log 'TEST MODE (DryRun) OK.' INFO;  exit 0 }
        if ($r.TimedOut)            { Write-Log 'TEST MODE timed out.'   ERROR; exit 4 }
        if ($r.ExitCode -eq 0)      { Write-Log 'TEST MODE success.'     INFO;  exit 0 }
        Write-Log ('TEST MODE failed, exit={0}' -f $r.ExitCode) ERROR
        exit 5
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
                    Invoke-PrintPipeline -SourcePath $f.FullName `
                                         -SumatraExe $resolvedSumatra `
                                         -Printer    $resolvedPrinter
                } catch {
                    Write-Log ("Pipeline exception on {0}: {1}" -f $f.Name, $_.Exception.Message) ERROR
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
