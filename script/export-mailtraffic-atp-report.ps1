# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Menyimpan output skrip ke file CSV dinamis di folder skrip.
# =========================================================================

# Variabel Global dan Output
$scriptName = "Export_ATP_Mail_Traffic_Report" 
$scriptOutput = @() 

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"
$outputFilePath = Join-Path -Path $PSScriptRoot -ChildPath $outputFileName


## -----------------------------------------------------------------------
## 1. PRASYARAT DAN INSTALASI MODUL
## -----------------------------------------------------------------------

Write-Host "--- 1. Memeriksa dan Menyiapkan Lingkungan PowerShell ---" -ForegroundColor Blue

# 1.1. Mengatur Execution Policy
Write-Host "1.1. Mengatur Execution Policy ke RemoteSigned..." -ForegroundColor Cyan
try {
    Set-ExecutionPolicy RemoteSigned -Scope Process -Force -ErrorAction Stop
    Write-Host " Execution Policy berhasil diatur." -ForegroundColor Green
} catch {
    Write-Error "Gagal mengatur Execution Policy: $($_.Exception.Message)"
    exit 1
}

# 1.2. Fungsi Pembantu untuk Cek dan Instal Modul
function CheckAndInstallModule {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ModuleName
    )

    Write-Host "1.$(++$global:moduleStep). Memeriksa Modul '$ModuleName'..." -ForegroundColor Cyan

    if (Get-Module -Name $ModuleName -ListAvailable) {
        Write-Host " Modul '$ModuleName' sudah terinstal. Melewati instalasi." -ForegroundColor Green
    } else {
        Write-Host " Modul '$ModuleName' belum ditemukan. Memulai instalasi..." -ForegroundColor Yellow
        try {
            Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
            Write-Host " Modul '$ModuleName' berhasil diinstal." -ForegroundColor Green
        } catch {
            Write-Error "Gagal menginstal modul '$ModuleName'. Pastikan PowerShellGet sudah terinstal dan koneksi internet tersedia."
            exit 1
        }
    }
}

$global:moduleStep = 1
CheckAndInstallModule -ModuleName "PowerShellGet"
CheckAndInstallModule -ModuleName "ExchangeOnlineManagement"

## -----------------------------------------------------------------------
## 2. KONEKSI WAJIB (EXCHANGE ONLINE)
## -----------------------------------------------------------------------

Write-Host "`n--- 2. Membangun Koneksi ke Exchange Online ---" -ForegroundColor Blue

if (-not (Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange"})) {
    Write-Host "Menghubungkan ke Exchange Online. Anda mungkin diminta untuk login..." -ForegroundColor Yellow
    try {
        Connect-ExchangeOnline -ShowProgress $false -ErrorAction Stop | Out-Null
        Write-Host "Koneksi ke Exchange Online berhasil dibuat." -ForegroundColor Green
    } catch {
        Write-Error "Gagal terhubung ke Exchange Online. $($_.Exception.Message)"
        exit 1
    }
} else {
    Write-Host "Sesi Exchange Online sudah aktif. Melewati koneksi ulang." -ForegroundColor Green
}

## -----------------------------------------------------------------------
## 3. LOGIKA UTAMA SCRIPT
## -----------------------------------------------------------------------

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

try {
    Write-Host "Mengambil data ATP Report (Lalu lintas terbaru)..." -ForegroundColor Cyan
    
    # Menarik data dari Exchange Online
    $atpData = Get-MailTrafficATPReport -ErrorAction Stop
    
    if ($null -ne $atpData) {
        $atpList = @($atpData)
        $totalItems = $atpList.Count
        $indexCount = 0

        # Kolom yang ingin dihilangkan (sesuai permintaan Bapak)
        $excludedFields = "SummarizeBy", "PivotBy", "StartDate", "EndDate", "AggregateBy", "Index"

        foreach ($report in $atpList) {
            $indexCount++
            Write-Host "-> [$indexCount/$totalItems] Memproses Laporan: $($report.Date) . . . Event: $($report.Event)" -ForegroundColor White
            
            # Memfilter objek untuk membuang kolom yang tidak diinginkan
            $filteredReport = $report | Select-Object * -ExcludeProperty $excludedFields
            
            # Memasukkan ke array output
            $scriptOutput += $filteredReport
        }
    } else {
        Write-Host "Tidak ada data ATP ditemukan untuk periode ini." -ForegroundColor Yellow
    }
} catch {
    Write-Host "   Gagal mengambil laporan ATP: $($_.Exception.Message)" -ForegroundColor Red
}

## -----------------------------------------------------------------------
## 4. CLEANUP, DISCONNECT, DAN EKSPOR HASIL
## -----------------------------------------------------------------------

Write-Host "`n--- 4. Cleanup, Memutus Koneksi, dan Ekspor Hasil ---" -ForegroundColor Blue

# 4.1. Ekspor Hasil
if ($scriptOutput.Count -gt 0) {
    Write-Host "Mengekspor $($scriptOutput.Count) baris data hasil skrip..." -ForegroundColor Yellow
    try {
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction Stop
        Write-Host " Data berhasil diekspor ke:" -ForegroundColor Green
        Write-Host " $outputFilePath" -ForegroundColor Green
    }
    catch {
        Write-Error "Gagal mengekspor data ke CSV: $($_.Exception.Message)"
    }
} else {
    Write-Host " Tidak ada data yang dikumpulkan (\$scriptOutput kosong). Melewati ekspor." -ForegroundColor DarkYellow
}

# 4.2. Memutus koneksi Exchange Online
if (Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange"}) {
    Write-Host "Memutuskan koneksi dari Exchange Online..." -ForegroundColor DarkYellow
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host " Koneksi Exchange Online diputus." -ForegroundColor Green
}

Write-Host "`nSkrip $($scriptName) selesai dieksekusi." -ForegroundColor Yellow