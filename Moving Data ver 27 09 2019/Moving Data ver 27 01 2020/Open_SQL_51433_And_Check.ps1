# ============================================================
# Set default local SQL Server instance to fixed port 51433
# + Open Windows Firewall inbound/outbound
# + Restart SQL Server service
# ============================================================

$ErrorActionPreference = 'Stop'

$Port = 51433
$SqlServiceName = 'MSSQLSERVER'
$InRuleName  = "SQL Default Instance TCP $Port IN"
$OutRuleName = "SQL Default Instance TCP $Port OUT"

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " Configure Default SQL Server Instance on Port $Port" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

# ------------------------------------------------------------
# 1) Check Admin
# ------------------------------------------------------------
$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "[ERROR] ·«“„  ‘€· PowerShell ﬂ„”ƒÊ· Run as Administrator" -ForegroundColor Red
    exit 1
}

# ------------------------------------------------------------
# 2) Detect SQL default instance ID
# ------------------------------------------------------------
$instanceNamesPath = 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL'

if (-not (Test-Path $instanceNamesPath)) {
    Write-Host "[ERROR] ·„ Ì „ «·⁄ÀÊ— ⁄·Ï „”«— Instance Names «·Œ«’ »‹ SQL Server" -ForegroundColor Red
    exit 1
}

$instanceProps = Get-ItemProperty -Path $instanceNamesPath

if (-not ($instanceProps.PSObject.Properties.Name -contains 'MSSQLSERVER')) {
    Write-Host "[ERROR] ·„ Ì „ «·⁄ÀÊ— ⁄·Ï Default Instance »«”„ MSSQLSERVER" -ForegroundColor Red
    Write-Host "ﬁœ  ﬂÊ‰ «·‰”Œ… Named Instance Ê·Ì”  Default Instance." -ForegroundColor Yellow
    exit 1
}

$instanceId = $instanceProps.MSSQLSERVER
Write-Host "Detected SQL Instance ID: $instanceId" -ForegroundColor Green

# ------------------------------------------------------------
# 3) Build registry paths
# ------------------------------------------------------------
$tcpRootPath = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$instanceId\MSSQLServer\SuperSocketNetLib\Tcp"
$ipAllPath   = Join-Path $tcpRootPath 'IPAll'

if (-not (Test-Path $tcpRootPath)) {
    Write-Host "[ERROR] ·„ Ì „ «·⁄ÀÊ— ⁄·Ï „”«— TCP «·Œ«’ »‹ SQL Server:" -ForegroundColor Red
    Write-Host $tcpRootPath -ForegroundColor Yellow
    exit 1
}

if (-not (Test-Path $ipAllPath)) {
    Write-Host "[ERROR] ·„ Ì „ «·⁄ÀÊ— ⁄·Ï IPAll path:" -ForegroundColor Red
    Write-Host $ipAllPath -ForegroundColor Yellow
    exit 1
}

# ------------------------------------------------------------
# 4) Show current TCP settings before change
# ------------------------------------------------------------
Write-Host ""
Write-Host "Current SQL TCP settings before update:" -ForegroundColor Cyan
try {
    $tcpBefore = Get-ItemProperty -Path $tcpRootPath
    $ipAllBefore = Get-ItemProperty -Path $ipAllPath

    Write-Host ("Tcp Enabled       : " + ($tcpBefore.Enabled | Out-String).Trim())
    Write-Host ("IPAll TcpPort     : " + ($ipAllBefore.TcpPort | Out-String).Trim())
    Write-Host ("IPAll DynamicPort : " + ($ipAllBefore.TcpDynamicPorts | Out-String).Trim())
}
catch {
    Write-Host "Could not read existing TCP settings." -ForegroundColor Yellow
}

# ------------------------------------------------------------
# 5) Enable TCP and set fixed port
# ------------------------------------------------------------
Write-Host ""
Write-Host "Applying SQL TCP settings..." -ForegroundColor Cyan

# Enable TCP at root
New-ItemProperty -Path $tcpRootPath -Name 'Enabled' -Value 1 -PropertyType DWord -Force | Out-Null

# Remove dynamic ports / set fixed port
New-ItemProperty -Path $ipAllPath -Name 'TcpDynamicPorts' -Value '' -PropertyType String -Force | Out-Null
New-ItemProperty -Path $ipAllPath -Name 'TcpPort' -Value "$Port" -PropertyType String -Force | Out-Null

Write-Host "SQL Server TCP/IP enabled." -ForegroundColor Green
Write-Host "Dynamic Ports cleared." -ForegroundColor Green
Write-Host "Fixed TCP Port set to $Port." -ForegroundColor Green

# ------------------------------------------------------------
# 6) Open Firewall rules
# ------------------------------------------------------------
Write-Host ""
Write-Host "Configuring Windows Firewall..." -ForegroundColor Cyan

if (-not (Get-NetFirewallRule -DisplayName $InRuleName -ErrorAction SilentlyContinue)) {
    New-NetFirewallRule `
        -DisplayName $InRuleName `
        -Direction Inbound `
        -Action Allow `
        -Protocol TCP `
        -LocalPort $Port | Out-Null
    Write-Host "Inbound firewall rule created." -ForegroundColor Green
}
else {
    Write-Host "Inbound firewall rule already exists." -ForegroundColor Yellow
}

if (-not (Get-NetFirewallRule -DisplayName $OutRuleName -ErrorAction SilentlyContinue)) {
    New-NetFirewallRule `
        -DisplayName $OutRuleName `
        -Direction Outbound `
        -Action Allow `
        -Protocol TCP `
        -LocalPort $Port | Out-Null
    Write-Host "Outbound firewall rule created." -ForegroundColor Green
}
else {
    Write-Host "Outbound firewall rule already exists." -ForegroundColor Yellow
}

# ------------------------------------------------------------
# 7) Restart SQL Server service
# ------------------------------------------------------------
Write-Host ""
Write-Host "Restarting SQL Server service ($SqlServiceName)..." -ForegroundColor Cyan

$svc = Get-Service -Name $SqlServiceName -ErrorAction SilentlyContinue
if (-not $svc) {
    Write-Host "[ERROR] Œœ„… SQL Server default instance €Ì— „ÊÃÊœ…: $SqlServiceName" -ForegroundColor Red
    exit 1
}

if ($svc.Status -eq 'Running') {
    Restart-Service -Name $SqlServiceName -Force
} else {
    Start-Service -Name $SqlServiceName
}

Start-Sleep -Seconds 6

$svc = Get-Service -Name $SqlServiceName
Write-Host "SQL Service Status: $($svc.Status)" -ForegroundColor Green

# ------------------------------------------------------------
# 8) Verify settings after restart
# ------------------------------------------------------------
Write-Host ""
Write-Host "SQL TCP settings after update:" -ForegroundColor Cyan
$tcpAfter = Get-ItemProperty -Path $tcpRootPath
$ipAllAfter = Get-ItemProperty -Path $ipAllPath

Write-Host ("Tcp Enabled       : " + ($tcpAfter.Enabled | Out-String).Trim())
Write-Host ("IPAll TcpPort     : " + ($ipAllAfter.TcpPort | Out-String).Trim())
Write-Host ("IPAll DynamicPort : " + ($ipAllAfter.TcpDynamicPorts | Out-String).Trim())

# ------------------------------------------------------------
# 9) Check if port is listening
# ------------------------------------------------------------
Write-Host ""
Write-Host "Checking if SQL is listening on port $Port ..." -ForegroundColor Cyan

$listen = Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue | Where-Object { $_.LocalPort -eq $Port }

if ($listen) {
    Write-Host "SUCCESS: Port $Port is now LISTENING." -ForegroundColor Green
    $listen | Format-Table LocalAddress, LocalPort, OwningProcess -AutoSize
} else {
    Write-Host "WARNING: Port $Port is not listening yet." -ForegroundColor Yellow
    Write-Host "ﬁœ  Õ «Ã  ‰ Ÿ— ÀÊ«‰Ì ≈÷«›Ì… √Ê  —«Ã⁄ ≈⁄œ«œ«  SQL Server Configuration Manager." -ForegroundColor Yellow
}

# ------------------------------------------------------------
# 10) Final info
# ------------------------------------------------------------
Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "Done." -ForegroundColor Green
Write-Host "Use this server format from remote machine:" -ForegroundColor Cyan
Write-Host "    ServerName,$Port"
Write-Host "or"
Write-Host "    IPAddress,$Port"
Write-Host "==================================================" -ForegroundColor Cyan