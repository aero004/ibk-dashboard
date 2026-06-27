# IBK Dashboard — avto-yangilanish, server va cloudflared nazorati

param(
    [string]$RepoDir      = "C:\servers\ibk-dashboard",
    [string]$ScriptName   = "ibk_dashboard.py",
    [string]$CFExe        = "C:\cloudflared.exe",
    [string]$CFConfig     = "C:\ProgramData\Cloudflare\config.yml",
    [int]   $IntervalSec  = 60
)

function Get-IBKProcess {
    Get-CimInstance Win32_Process -Filter "Name='python.exe' OR Name='python3.exe'" |
        Where-Object { $_.CommandLine -like "*$ScriptName*" }
}

function Get-CFProcess {
    Get-CimInstance Win32_Process -Filter "Name='cloudflared.exe'"
}

Write-Host "[auto-update] Ishga tushdi. Papka: $RepoDir | Interval: ${IntervalSec}s"

while ($true) {
    try {
        Set-Location $RepoDir

        # --- Git: yangilanish bormi? ---
        $before = git rev-parse HEAD 2>&1
        $fetchOut = git fetch origin main 2>&1
        if ($fetchOut) { Write-Host "[auto-update] git fetch: $fetchOut" }

        $resetOut = git reset --hard origin/main 2>&1
        $after  = git rev-parse HEAD 2>&1

        $ts = Get-Date -Format "HH:mm:ss"

        if ($before -ne $after) {
            $short = ($after -replace '\s','').Substring(0, [Math]::Min(7, ($after -replace '\s','').Length))
            Write-Host "[auto-update] $ts Ozgarish topildi ($short). Qayta ishga tushirilmoqda..."

            $old = Get-IBKProcess
            if ($old) {
                $old | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
                Start-Sleep -Seconds 2
            }

            Start-Process python -ArgumentList "$RepoDir\$ScriptName" -WorkingDirectory $RepoDir -WindowStyle Hidden
            Write-Host "[auto-update] Server qayta ishga tushirildi."
        } else {
            Write-Host "[auto-update] $ts Ozgarish yoq. ($($before.ToString().Trim().Substring(0,7)))"
        }

        if (-not (Get-IBKProcess)) {
            Write-Host "[auto-update] $ts Server ishlamayapti. Ishga tushirilmoqda..."
            Start-Process python -ArgumentList "$RepoDir\$ScriptName" -WorkingDirectory $RepoDir -WindowStyle Hidden
            Write-Host "[auto-update] Server ishga tushirildi."
        }

        # --- Cloudflare tunnel nazorati ---
        if (-not (Get-CFProcess)) {
            Write-Host "[auto-update] $ts Cloudflared ishlamayapti. Qayta ishga tushirilmoqda..."
            Start-Process $CFExe -ArgumentList "tunnel","--config",$CFConfig,"run" -WindowStyle Hidden
            Write-Host "[auto-update] Cloudflared ishga tushirildi."
        }

    } catch {
        Write-Host "[auto-update] Xato: $($_.ToString())"
    }

    Start-Sleep -Seconds $IntervalSec
}
