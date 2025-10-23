# Revision : 1.0
# Description : Run periodic internet speed tests (Ookla CLI if available) and append results to a pinned CSV log file for a specified duration and interval. Rev 1.0
# Author : Jason Lamb (with help from ChatGPT)
# Created Date : 2025-10-21
# Modified Date : 2025-10-21

param(
    [int]$RunForMinutes = 60,                       # total runtime window
    [int]$IntervalMinutes = 5,                      # wait between tests
    [string]$LogPath = "C:\temp\powershell-exports\speedtest-network.csv",  # pinned log file (CSV)
    [switch]$AutoInstall                            # attempt winget install of Ookla CLI if not found
)

# --- Prep: folders & header ---
$logFolder = Split-Path -Path $LogPath -Parent
if (-not (Test-Path $logFolder)) {
    New-Item -Path $logFolder -ItemType Directory -Force | Out-Null
}
if (-not (Test-Path $LogPath)) {
    # Create with header
    @'
Timestamp,ComputerName,LatencyMs,DownloadMbps,UploadMbps,PacketLoss,ISP,ServerName,ServerLocation,ResultUrl
'@ | Out-File -FilePath $LogPath -Encoding UTF8 -Force
    Write-Host "Created log file $LogPath : with CSV header"
}

# --- Locate or install Ookla Speedtest CLI ---
function Get-OklaSpeedtestPath {
    $cmd = Get-Command speedtest -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    if ($AutoInstall) {
        Write-Host "speedtest.exe not found, attempting winget install : Ookla.Speedtest.CLI"
        try {
            winget install --id Ookla.Speedtest.CLI -e --accept-package-agreements --accept-source-agreements | Out-Null
            Start-Sleep -Seconds 3
            $cmd = Get-Command speedtest -ErrorAction SilentlyContinue
            if ($cmd) { return $cmd.Source }
        } catch {
            Write-Host "Auto-install failed on this system : $_"
        }
    }

    return $null
}

$speedtestPath = Get-OklaSpeedtestPath
if (-not $speedtestPath) {
    Write-Host "ERROR : speedtest.exe (Ookla CLI) not found. Install from https://www.speedtest.net/apps/cli or run with -AutoInstall (requires winget)."
    return
}

# --- Core test function (Ookla JSON -> CSV line) ---
function Invoke-NetworkSpeedTest {
    try {
        # Accept license/GDPR non-interactively and emit JSON
        $raw = & $speedtestPath --accept-license --accept-gdpr -f json 2>$null
        if ([string]::IsNullOrWhiteSpace($raw)) {
            throw "speedtest returned no output."
        }
        $j = $raw | ConvertFrom-Json

        # Ookla CLI bandwidth fields are bytes/sec; convert to Mbps (decimal)
        $dlMbps = if ($j.download.bandwidth) { [math]::Round(($j.download.bandwidth * 8) / 1000000, 2) } else { $null }
        $ulMbps = if ($j.upload.bandwidth)   { [math]::Round(($j.upload.bandwidth   * 8) / 1000000, 2) } else { $null }
        $lat    = $j.ping.latency
        $loss   = if ($null -ne $j.packetLoss) { $j.packetLoss } else { $null }
        $isp    = $j.isp
        $srvN   = $j.server.name
        $srvL   = $j.server.location
        $url    = $j.result.url
        $ts     = (Get-Date).ToString("s")

        # Write to CSV (append)
        [pscustomobject]@{
            Timestamp       = $ts
            ComputerName    = $env:COMPUTERNAME
            LatencyMs       = $lat
            DownloadMbps    = $dlMbps
            UploadMbps      = $ulMbps
            PacketLoss      = $loss
            ISP             = $isp
            ServerName      = $srvN
            ServerLocation  = $srvL
            ResultUrl       = $url
        } | Export-Csv -Path $LogPath -Append -NoTypeInformation -UseQuotes Always

        Write-Host "Logged to $LogPath : $ts  DL ${dlMbps}Mbps  UL ${ulMbps}Mbps  Ping ${lat}ms"
    }
    catch {
        $ts = (Get-Date).ToString("s")
        Write-Host "Speed test failed at $ts : $_"
        # Append an error row (with blanks for speeds)
        [pscustomobject]@{
            Timestamp       = $ts
            ComputerName    = $env:COMPUTERNAME
            LatencyMs       = $null
            DownloadMbps    = $null
            UploadMbps      = $null
            PacketLoss      = $null
            ISP             = "ERROR"
            ServerName      = $null
            ServerLocation  = $null
            ResultUrl       = $null
        } | Export-Csv -Path $LogPath -Append -NoTypeInformation -UseQuotes Always
    }
}

# --- Scheduler loop (run for window with interval) ---
$endTime = (Get-Date).AddMinutes($RunForMinutes)
do {
    $start = Get-Date
    Invoke-NetworkSpeedTest

    # compute next tick without overshooting beyond end window
    $nextPlanned = $start.AddMinutes($IntervalMinutes)
    $now = Get-Date
    $sleepMs = [int]([math]::Max(0, ($nextPlanned - $now).TotalMilliseconds))
    if ($sleepMs -gt 0 -and (Get-Date) -lt $endTime) {
        Start-Sleep -Milliseconds $sleepMs
    }
} while ((Get-Date) -lt $endTime)

Write-Host "Completed run window of ${RunForMinutes} minute(s) : Log file $LogPath"

<# =========================
CHANGELOG / WHAT CHANGED (Rev 1.0)
- New script to run Ookla speed tests on a loop for a specified duration.
- Appends results to a single pinned CSV log at C:\temp\powershell-exports\speedtest-network.csv.
- Auto-detects speedtest.exe; optional -AutoInstall via winget.
- CSV columns: Timestamp, ComputerName, LatencyMs, DownloadMbps, UploadMbps, PacketLoss, ISP, ServerName, ServerLocation, ResultUrl.
========================= #>

<# =========================
USAGE EXAMPLES (dot-source then call)

. .\Run-SpeedTest-Logger.ps1

# 1) Default : run for 60 minutes, test every 5 minutes, pinned log
Run-SpeedTest-Logger -RunForMinutes 60 -IntervalMinutes 5

# 2) Custom interval every 1 minute for 15 minutes
Run-SpeedTest-Logger -RunForMinutes 15 -IntervalMinutes 1

# 3) Specify a custom pinned log path
Run-SpeedTest-Logger -RunForMinutes 30 -IntervalMinutes 5 -LogPath "C:\temp\powershell-exports\office-speed.csv"

# 4) Attempt auto-install of Ookla CLI via winget if missing
Run-SpeedTest-Logger -RunForMinutes 30 -IntervalMinutes 5 -AutoInstall

#> 

# --- Turn the script into a callable function when dot-sourced ---
function Run-SpeedTest-Logger {
    param(
        [int]$RunForMinutes = $RunForMinutes,
        [int]$IntervalMinutes = $IntervalMinutes,
        [string]$LogPath = $LogPath,
        [switch]$AutoInstall = $AutoInstall
    )
    & $PSCommandPath -RunForMinutes $RunForMinutes -IntervalMinutes $IntervalMinutes -LogPath $LogPath @($AutoInstall.IsPresent ? '-AutoInstall' : $null)
}
