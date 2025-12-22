# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Nama Skrip: Export-AllEntraDevices
# Deskripsi: Mengambil semua data perangkat dari Microsoft Entra ID.
# =========================================================================

# Variabel Global dan Output
$scriptName = "ExportAllEntraDevices" 
$scriptOutput = [System.Collections.ArrayList]::new() 

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName


## -----------------------------------------------------------------------
## 2. KONEKSI WAJIB (MICROSOFT ENTRA)
## -----------------------------------------------------------------------

Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Entra ---" -ForegroundColor Blue

try {
    Write-Host "Menghubungkan ke Microsoft Entra. Selesaikan login pada pop-up..." -ForegroundColor Yellow
    
    # Menangani potensi konflik DLL dengan mencoba Disconnect terlebih dahulu
    Disconnect-Entra -ErrorAction SilentlyContinue
    
    # Koneksi utama
    Connect-Entra -Scopes 'Device.Read.All' -ErrorAction Stop
    Write-Host "Koneksi ke Microsoft Entra berhasil dibuat." -ForegroundColor Green
} catch {
    Write-Error "Gagal terhubung: $($_.Exception.Message)"
    Write-Host "`nTIP: Jika error library berlanjut, tutup SEMUA jendela PowerShell lalu buka kembali." -ForegroundColor Yellow
    exit 1
}


## -----------------------------------------------------------------------
## 3. LOGIKA UTAMA SCRIPT
## -----------------------------------------------------------------------

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

try {
    Write-Host "Mengambil data semua perangkat (Get-EntraDevice -All)..." -ForegroundColor Cyan
    
    # Menjalankan logika utama sesuai instruksi Anda
    $devices = Get-EntraDevice -All | Select-Object AccountEnabled, DeviceId, OperatingSystem, ApproximateLastSignInDateTime, DisplayName
    
    # Menampilkan hasil di konsol dalam bentuk tabel
    $devices | Format-Table -AutoSize
    
    # Menyimpan ke variabel output untuk keperluan ekspor CSV
    foreach ($dev in $devices) {
        [void]$scriptOutput.Add($dev)
    }

} catch {
    Write-Error "Terjadi kesalahan saat mengambil data perangkat: $($_.Exception.Message)"
}

## -----------------------------------------------------------------------
## 4. CLEANUP, DISCONNECT, DAN EKSPOR HASIL
## -----------------------------------------------------------------------

Write-Host "`n--- 4. Cleanup, Memutus Koneksi, dan Ekspor Hasil ---" -ForegroundColor Blue

# 4.1. Ekspor Hasil
if ($scriptOutput.Count -gt 0) {
    Write-Host "Mengekspor $($scriptOutput.Count) data perangkat..." -ForegroundColor Yellow
    try {
        # Mengekspor hasil ke file CSV dengan pemisah titik koma
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction Stop
        Write-Host " Data berhasil diekspor ke: $outputFilePath" -ForegroundColor Green
    }
    catch {
        Write-Error "Gagal mengekspor data ke CSV: $($_.Exception.Message)"
    }
} else {
    Write-Host " Tidak ada data perangkat yang ditemukan." -ForegroundColor DarkYellow
}

# 4.2. Memutus koneksi Entra
Write-Host "Memutuskan sesi Microsoft Entra..." -ForegroundColor DarkYellow
Disconnect-Entra -ErrorAction SilentlyContinue

Write-Host "`nSkrip $($scriptName) selesai dieksekusi." -ForegroundColor Yellow