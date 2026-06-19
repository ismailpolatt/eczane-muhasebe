# ============================================================
#  RxEys .back Otomatik Yukleyici  (eczane PC'sinde calisir)
# ------------------------------------------------------------
#  RxEys'in .back yedek klasorunu izler, yeni dosyalari ev
#  sunucusuna (ecz.polathome.com) otomatik gonderir.
#
#  Bulut/Drive gerekmez. Ev sunucusu eczaneden erisilebildigi
#  surece (web arayuzunu actiginiz adres) calisir.
#
#  Kurulum: en altta "GOREV ZAMANLAYICI" bolumune bakin.
# ============================================================

# ── AYARLAR ─────────────────────────────────────────────────
# RxEys'in .back dosyalarini kaydettigi klasor (DEGISTIRIN):
$WatchFolder = "C:\RxEysYedek"

# Ev sunucu adresi. Gerekirse http:// veya :5001 ekleyin.
$ServerUrl   = "https://ecz.polathome.com"
# ────────────────────────────────────────────────────────────

$Endpoint  = "$ServerUrl/api/import-back"
$StateFile = Join-Path $PSScriptRoot "uploaded.json"
$curl      = "$env:SystemRoot\System32\curl.exe"

# Daha once yuklenen dosyalari hatirla (ad + degisiklik zamani)
$state = @()
if (Test-Path $StateFile) {
    try { $state = @(Get-Content $StateFile -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { $state = @() }
}

if (-not (Test-Path $WatchFolder)) {
    Write-Host "HATA: Klasor bulunamadi -> $WatchFolder"
    exit 1
}
if (-not (Test-Path $curl)) {
    Write-Host "HATA: curl.exe bulunamadi. Windows 10/11 gerekir."
    exit 1
}

$uploaded = 0
$files = Get-ChildItem -Path $WatchFolder -Filter *.back -File | Sort-Object LastWriteTime

foreach ($f in $files) {
    $key = "$($f.Name)|$($f.LastWriteTime.Ticks)"
    if ($state -contains $key) { continue }

    Write-Host "Yukleniyor: $($f.Name) ..."
    $result = & $curl -s -S -X POST -F "file=@$($f.FullName)" $Endpoint 2>&1

    if ($LASTEXITCODE -eq 0 -and "$result" -match '"ok"\s*:\s*true') {
        Write-Host "  OK  -> $result"
        $state += $key
        $uploaded++
    } else {
        Write-Host "  HATA -> $result"
    }
}

# Durumu kaydet
$state | ConvertTo-Json | Set-Content -Path $StateFile -Encoding UTF8
Write-Host "Tamamlandi. $uploaded yeni dosya yuklendi. ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))"

# ============================================================
#  GOREV ZAMANLAYICI (Task Scheduler) ile otomatik calistirma
# ------------------------------------------------------------
#  1. "Gorev Zamanlayici" (taskschd.msc) ac
#  2. "Temel Gorev Olustur" -> ad: "Eczane Back Yukleyici"
#  3. Tetikleyici: "Gunluk" -> Baslama saati: 19:00
#     (eczane 19:00'da yedek aldigi icin gunde bir kez yeterli)
#  4. Eylem: "Program baslat"
#       Program : powershell.exe
#       Arguman : -NoProfile -ExecutionPolicy Bypass -File "C:\yol\eczane_uploader.ps1"
#  5. Bitir. Artik her gun 19:00'da yeni .back otomatik sunucuya gider.
#
#  Not: 19:00'da PC kapali/uykuda olabilir diyorsaniz, gorev
#  ozelliklerinde "Kacirilan gorevi mumkun oldugunca calistir"
#  secenegini isaretleyin (PC acilinca calisir).
# ============================================================
