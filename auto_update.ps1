# IBK Dashboard — avto-yangilanish va qayta ishga tushirish

param(
    [string]$RepoDir    = "C:\servers\ibk-dashboard",
    [string]$ScriptName = "ibk_dashboard.py",
    [int]   $IntervalSec = 60
)

function Get-IBKProcess {
    Get-CimInstance Win32_Process -Filter "Name='python.exe' OR Name='python3.exe'" |
        Where-Object { $_.CommandLine -like "*$ScriptName*" }
}

Write-Host "[auto-update] Ishga tushdi. Papka: $RepoDir | Interval: ${IntervalSec}s"

while ($true) {
    try {
        Set-Location $RepoDir

        $before = git rev-parse HEAD 2>$null
        git fetch --quiet origin main 2>$null
        git reset --hard origin/main --quiet 2>$null
        $after  = git rev-parse HEAD 2>$null

        $ts = Get-Date -Format "HH:mm:ss"

        if ($before -ne $after) {
            $short = $after.Substring(0, 7)
            Write-Host "[auto-update] $ts Ozgarish topildi ($short). Qayta ishga tushirilmoqda..."

            $old = Get-IBKProcess
            if ($old) {
                $old | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
                Start-Sleep -Seconds 2
            }

            Start-Process python -ArgumentList "$RepoDir\$ScriptName" -WorkingDirectory $RepoDir -WindowStyle Hidden
            Write-Host "[auto-update] Server qayta ishga tushirildi."
        } else {
            Write-Host "[auto-update] $ts Ozgarish yoq."
        }

        if (-not (Get-IBKProcess)) {
            Write-Host "[auto-update] Server ishlamayapti. Ishga tushirilmoqda..."
            Start-Process python -ArgumentList "$RepoDir\$ScriptName" -WorkingDirectory $RepoDir -WindowStyle Hidden
        }

    } catch {
        Write-Host "[auto-update] Xato: $($_.ToString())"
    }

    Start-Sleep -Seconds $IntervalSec
}
