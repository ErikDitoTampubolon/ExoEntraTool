<#
.SYNOPSIS
Mengekspor laporan kuota (Total, Terpakai, Tersedia) semua lisensi Microsoft 365 yang disubskripsikan menggunakan Microsoft Graph.

.DESCRIPTION
Skrip ini mengambil data Get-MgSubscribedSku, menghitung kuota yang tersedia untuk setiap jenis lisensi, dan menampilkan hasilnya di konsol dan mengekspornya ke CSV.

.AUTHOR
AI PowerShell Expert

.VERSION
1.0 (License Quota Report - Built on User Framework)

.PREREQUISITES
Modul Microsoft.Graph harus terinstal. Diperlukan koneksi internet dan akun administrator Microsoft 365 untuk login dengan scope 'Organization.Read.All'.

.NOTES
Output file akan disimpan di direktori yang sama dengan skrip ini, dengan nama Output_ExportLicenseQuotaReport_[Timestamp].csv.
#>
# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Menyimpan output skrip ke file CSV dinamis di folder skrip.
# =========================================================================

# Variabel Global dan Output
$scriptName = "ExportLicenseQuotaReport" # Nama skrip yang sebenarnya
$scriptOutput = @() # Array tempat semua data hasil skrip dikumpulkan

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
# Menggunakan $PSScriptRoot memastikan file disimpan di folder yang sama dengan skrip
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"

# Penanganan kasus $PSScriptRoot tidak ada saat dijalankan dari konsol
$scriptDir = if ($PSScriptRoot) {$PSScriptRoot} else {(Get-Location).Path}
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

## -----------------------------------------------------------------------
## 2. KONEKSI KE MICROSOFT GRAPH
## -----------------------------------------------------------------------

$requiredScopes = "User.ReadWrite.All", "Organization.Read.All"
Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Graph ---" -ForegroundColor Blue

if (Get-MgContext -ErrorAction SilentlyContinue) {
    Write-Host "Sesi Graph yang ada akan diputus untuk koneksi ulang dengan scopes baru." -ForegroundColor DarkYellow
    Disconnect-MgGraph
}

Write-Host "Anda akan diminta untuk login. Pastikan Anda menyetujui scopes berikut:" -ForegroundColor Cyan
Write-Host $requiredScopes -ForegroundColor Yellow

try {
    Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop | Out-Null
    Write-Host "✅ Koneksi ke Microsoft Graph berhasil!" -ForegroundColor Green
} catch {
    Write-Error "Gagal terhubung ke Microsoft Graph. Pastikan Anda memiliki kredensial dan hak akses yang benar."
    exit 1
}


## -----------------------------------------------------------------------
## 3. LOGIKA UTAMA SCRIPT ANDA DI SINI
## -----------------------------------------------------------------------

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

# >>> GANTI BAGIAN INI DENGAN KODE UTAMA SKRIP ANDA <<<

Write-Host "3.1. Mengambil detail semua Lisensi yang Disubskripsikan (SKU)..." -ForegroundColor Cyan

try {
    # Ambil semua SKU yang disubskripsikan
    $subscribedSkus = Get-MgSubscribedSku -ErrorAction Stop

    $totalSkus = $subscribedSkus.Count
    Write-Host "   ✅ Ditemukan $($totalSkus) SKU Lisensi Aktif." -ForegroundColor Green
    
    $i = 0
    foreach ($sku in $subscribedSkus) {
        $i++
        
        Write-Progress -Activity "Collecting License Quota Data" `
                       -Status "Processing License $i of ${totalSkus}: $($sku.SkuName)" `
                       -PercentComplete ([int](($i / $totalSkus) * 100))

        # Hitung Kuota
        $totalUnits = $sku.PrepaidUnits.Enabled
        $consumedUnits = $sku.ConsumedUnits
        $availableUnits = $totalUnits - $consumedUnits
        
        # Bangun objek kustom untuk diekspor
        $scriptOutput += [PSCustomObject]@{
            LicenseName = $sku.SkuName
            SkuPartNumber = $sku.SkuPartNumber
            CapabilityStatus = $sku.CapabilityStatus
            TotalUnits = $totalUnits
            ConsumedUnits = $consumedUnits
            AvailableUnits = $availableUnits
        }
    }
    
    Write-Progress -Activity "Collecting License Data Complete" -Status "Exporting Results" -Completed

    # Tampilkan di Konsol (Wajib Sesuai Permintaan)
    Write-Host "`n--- Hasil Laporan Kuota Lisensi ---" -ForegroundColor Blue
    if ($scriptOutput.Count -gt 0) {
        $scriptOutput | Format-Table -AutoSize
    } else {
        Write-Host "Tidak ada data yang tersedia untuk ditampilkan." -ForegroundColor DarkYellow
    }
    Write-Host "--------------------------------------------------------" -ForegroundColor Blue

}
catch {
    $reason = "Gagal fatal saat mengambil data Lisensi dari Microsoft Graph. Pastikan Anda memiliki scope 'Organization.Read.All' yang aktif. Error: $($_.Exception.Message)"
    Write-Error $reason
    # Tambahkan error fatal ke output jika terjadi
    $scriptOutput += [PSCustomObject]@{
        LicenseName = "FATAL ERROR"; SkuPartNumber = "N/A"; CapabilityStatus = "FAIL";
        TotalUnits = "N/A"; ConsumedUnits = "N/A"; AvailableUnits = "N/A"
    }
}

# >>> AKHIR DARI KODE UTAMA SKRIP ANDA <<<

## -----------------------------------------------------------------------
## 4. CLEANUP, DISCONNECT, DAN EKSPOR HASIL
## -----------------------------------------------------------------------

Write-Host "`n--- 4. Cleanup, Memutus Koneksi, dan Ekspor Hasil ---" -ForegroundColor Blue

# 4.1. Ekspor Hasil
# Hanya ekspor jika ada data yang valid (bukan hanya error fatal)
if ($scriptOutput.Count -gt 0 -and ($scriptOutput | Where-Object {$_.LicenseName -ne "FATAL ERROR"}).Count -gt 0) {
    Write-Host "Mengekspor $($scriptOutput.Count) baris data hasil skrip..." -ForegroundColor Yellow
    try {
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -ErrorAction Stop
        Write-Host " ✅ Data berhasil diekspor ke:" -ForegroundColor Green
        Write-Host " $outputFilePath" -ForegroundColor Green
    }
    catch {
        Write-Error "Gagal mengekspor data ke CSV: $($_.Exception.Message)"
    }
} else {
    Write-Host " ⚠️ Tidak ada data lisensi valid yang dikumpulkan. Melewati ekspor." -ForegroundColor DarkYellow
}

# 4.2. Memutus koneksi Microsoft Graph
# MODIFIKASI: Menggunakan Disconnect-MgGraph
if (Get-MgContext -ErrorAction SilentlyContinue) {
    Write-Host "Memutuskan koneksi dari Microsoft Graph..." -ForegroundColor DarkYellow
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Write-Host " ✅ Koneksi Microsoft Graph diputus." -ForegroundColor Green
}

Write-Host "`nSkrip $($scriptName) selesai dieksekusi." -ForegroundColor Yellow