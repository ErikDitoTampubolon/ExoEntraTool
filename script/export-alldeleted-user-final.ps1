# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.1)
# Nama Skrip: Export-EntraDeletedUsers
# Deskripsi: Mengambil daftar pengguna yang dihapus dengan UI Progress.
# =========================================================================

# Variabel Global dan Output
$scriptName = "ExportEntraDeletedUsers" 
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
    Connect-Entra -Scopes 'User.Read.All' -ErrorAction Stop
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
    Write-Host "Mengambil data pengguna yang dihapus..." -ForegroundColor Cyan
    
    $rawDeleted = Get-EntraDeletedUser -All -ErrorAction Stop
    $totalData = $rawDeleted.Count
    
    if ($totalData -gt 0) {
        $i = 0
        foreach ($user in $rawDeleted) {
            $i++
            
            # OUTPUT PROGRES BARIS TUNGGAL SESUAI PERMINTAAN
            # Menggunakan -NoNewline dan `r untuk menjaga di satu baris
            $statusText = "-> [$i/$totalData] Memproses: $($user.UserPrincipalName) . . ."
            Write-Host "`r$statusText" -ForegroundColor White -NoNewline
            
            # Mapping data ke objek hasil
            $obj = [PSCustomObject]@{
                Id                 = $user.Id
                UserPrincipalName  = $user.UserPrincipalName
                DisplayName        = $user.DisplayName
                AccountEnabled     = $user.AccountEnabled
                DeletedDateTime    = $user.DeletedDateTime
                DeletionAgeInDays  = $user.DeletionAgeInDays
                UserType           = $user.UserType
            }
            [void]$scriptOutput.Add($obj)
        }
        Write-Host "`n`nBerhasil memproses $totalData pengguna." -ForegroundColor Green
    } else {
        Write-Host "Tidak ditemukan pengguna yang dihapus." -ForegroundColor Yellow
    }
} catch {
    Write-Error "Terjadi kesalahan: $($_.Exception.Message)"
}

## -----------------------------------------------------------------------
## 4. CLEANUP, DISCONNECT, DAN EKSPOR HASIL
## -----------------------------------------------------------------------

Write-Host "`n--- 4. Cleanup, Memutus Koneksi, dan Ekspor Hasil ---" -ForegroundColor Blue

if ($scriptOutput.Count -gt 0) {
    Write-Host "Mengekspor data ke CSV..." -ForegroundColor Yellow
    try {
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction Stop
        Write-Host "Data tersimpan di: $outputFilePath" -ForegroundColor Green
    }
    catch {
        Write-Error "Gagal ekspor: $($_.Exception.Message)"
    }
}

Disconnect-Entra -ErrorAction SilentlyContinue
Write-Host "`nSkrip selesai dieksekusi." -ForegroundColor Yellow