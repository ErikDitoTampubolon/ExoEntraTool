# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.1)
# Nama Skrip: Export-EntraInactiveGuestUsers
# Deskripsi: Mengambil daftar Guest User yang tidak aktif > 30 hari.
# =========================================================================

# Variabel Global dan Output
$scriptName = "ExportEntraInactiveGuestUsers" 
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
    Connect-Entra -Scopes 'AuditLog.Read.All','User.Read.All' -ErrorAction Stop
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
    Write-Host "Menganalisis seluruh pengguna yang tidak aktif > 30 hari..." -ForegroundColor Cyan
    
    # Menjalankan logika utama: Get-EntraInactiveSignInUser -LastSignInBeforeDaysAgo 30 -All
    $inactiveUsers = Get-EntraInactiveSignInUser -LastSignInBeforeDaysAgo 30 -All -ErrorAction Stop
    $totalData = $inactiveUsers.Count
    
    if ($totalData -gt 0) {
        $i = 0
        foreach ($user in $inactiveUsers) {
            $i++
            
            # Output progres baris tunggal (UI Refresh) sesuai permintaan sebelumnya
            $statusText = "-> [$i/$totalData] Memproses: $($user.UserPrincipalName) . . ."
            Write-Host "`r$statusText" -ForegroundColor Green -NoNewline
            
            # Mapping data ke objek hasil untuk ekspor CSV
            $obj = [PSCustomObject]@{
                DisplayName              = $user.DisplayName
                UserPrincipalName        = $user.UserPrincipalName
                UserType                 = $user.UserType
                AccountEnabled           = $user.AccountEnabled
                LastSignInDateTime       = $user.SignInActivity.LastSignInDateTime
                LastNonInteractiveSignIn = $user.SignInActivity.LastNonInteractiveSignInDateTime
                Id                       = $user.Id
            }
            [void]$scriptOutput.Add($obj)
        }
        Write-Host "`n`nBerhasil memproses $totalData pengguna tidak aktif." -ForegroundColor Green
    } else {
        Write-Host "Tidak ditemukan pengguna yang tidak aktif dalam 30 hari terakhir." -ForegroundColor Yellow
    }
} catch {
    Write-Error "Terjadi kesalahan saat mengambil data: $($_.Exception.Message)"
}

## -----------------------------------------------------------------------
## 4. CLEANUP, DISCONNECT, DAN EKSPOR HASIL
## -----------------------------------------------------------------------

Write-Host "`n--- 4. Cleanup, Memutus Koneksi, dan Ekspor Hasil ---" -ForegroundColor Blue

# 4.1. Ekspor Hasil
if ($scriptOutput.Count -gt 0) {
    Write-Host "Mengekspor data ke file CSV..." -ForegroundColor Yellow
    try {
        # Menggunakan titik koma (;) sebagai delimiter agar rapi saat dibuka di Excel
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction Stop
        Write-Host " Data berhasil diekspor ke: $outputFilePath" -ForegroundColor Green
    }
    catch {
        Write-Error "Gagal mengekspor data ke CSV: $($_.Exception.Message)"
    }
} else {
    Write-Host " Tidak ada data yang dikumpulkan. Melewati ekspor." -ForegroundColor DarkYellow
}

# 4.2. Memutus koneksi Entra
Write-Host "Memutuskan koneksi dari Microsoft Entra..." -ForegroundColor DarkYellow
Disconnect-Entra -ErrorAction SilentlyContinue
Write-Host " Sesi telah ditutup." -ForegroundColor Green

Write-Host "`nSkrip $($scriptName) selesai dieksekusi." -ForegroundColor Yellow