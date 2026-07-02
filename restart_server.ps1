# IBK Dashboard serverni tuzatish va qayta ishga tushirish
# ISHLATISH: "ibk-dashboard" papkasi ichida turib shu skriptni ishga tushiring:
#   .\restart_server.ps1

Write-Host "1) Eski serverni qidirilmoqda va to'xtatilmoqda..." -ForegroundColor Cyan
$old = Get-CimInstance Win32_Process -Filter "Name = 'python.exe' OR Name = 'pythonw.exe'" |
    Where-Object { $_.CommandLine -like "*ibk_dashboard.py*" }
if ($old) {
    foreach ($p in $old) {
        Write-Host "   Eski jarayon o'chirilmoqda: PID $($p.ProcessId)"
        Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 2
} else {
    Write-Host "   Ishlab turgan eski jarayon topilmadi (ehtimol allaqachon o'chgan)."
}

Write-Host "2) Yangi kod yuklab olinmoqda (git pull)..." -ForegroundColor Cyan
git pull

Write-Host "3) Server qayta ishga tushirilmoqda..." -ForegroundColor Cyan
Start-Process -FilePath python -ArgumentList "ibk_dashboard.py" -WindowStyle Hidden `
  -RedirectStandardOutput "server_out.txt" `
  -RedirectStandardError "server_err.txt"

Write-Host "4) Tekshirilmoqda..." -ForegroundColor Cyan
Start-Sleep -Seconds 4
try {
    $r = Invoke-WebRequest -Uri "http://localhost:8788/api/server_info" -UseBasicParsing -TimeoutSec 8
    if ($r.StatusCode -eq 200) {
        Write-Host "TAYYOR: Server muvaffaqiyatli ishga tushdi (localhost:8788 javob berdi)." -ForegroundColor Green
    }
} catch {
    Write-Host "DIQQAT: Server hali javob bermayapti." -ForegroundColor Red
    Write-Host "server_err.txt faylini oching va oxirgi qatorlarni ko'ring - u yerda xato sababi yozilgan bo'ladi." -ForegroundColor Yellow
}
