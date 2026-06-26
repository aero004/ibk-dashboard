# IBK Dashboard — avto-yangilanish va qayta ishga tushirish
# Server kompda bu skript fon rejimida ishlaydi.
# Har 60 soniyada GitHub dan o'zgarish borligini tekshiradi.
# O'zgarish bo'lsa — eski serverni to'xtatib, yangisini ishga tushiradi.

param(
    [string]$RepoDir   = "C:\servers\ibk-dashboard",
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

        # Git pull qilish
        $before = git rev-parse HEAD 2>$null
        git fetch --quiet origin main 2>$null
        git reset --hard origin/main --quiet 2>$null
        $after  = git rev-parse HEAD 2>$null

        if ($before -ne $after) {
            Write-Host "[auto-update] $(Get-Date -f 'HH:mm:ss') O'zgarish topildi ($($after.Substring(0,7))). Qayta ishga tushirilmoqda..."

            # Eski jarayonni to'xtatish
            $old = Get-IBKProcess
            if ($old) {
                $old | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
                Start-Sleep -Seconds 2
            }

            # Yangi jarayonni boshlash
            Start-Process python -ArgumentList "$RepoDir\$ScriptName" -WorkingDirectory $RepoDir -WindowStyle Hidden
            Write-Host "[auto-update] Server qayta ishga tushirildi."
        } else {
            Write-Host "[auto-update] $(Get-Date -f 'HH:mm:ss') O'zgarish yo'q."
        }

        # Server ishlamayotgan bo'lsa, ishga tushirish
        if (-not (Get-IBKProcess)) {
            Write-Host "[auto-update] Server ishlamayapti — ishga tushirilmoqda..."
            Start-Process python -ArgumentList "$RepoDir\$ScriptName" -WorkingDirectory $RepoDir -WindowStyle Hidden
        }

    } catch {
        Write-Host "[auto-update] Xato: $_"
    }

    Start-Sleep -Seconds $IntervalSec
}
