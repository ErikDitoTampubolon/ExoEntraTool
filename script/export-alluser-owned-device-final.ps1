# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.1)
# Nama Skrip: Export-AllUserOwnedDevice
# Deskripsi: Mengekspor daftar semua perangkat milik semua pengguna Entra ID.
# =========================================================================

# Variabel Global dan Output
$scriptName = "ExportAllUserOwnedDevice" 
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
    Write-Host "Mengambil daftar semua pengguna..." -ForegroundColor Cyan
    $allUsers = Get-EntraUser -All
    $totalUsers = $allUsers.Count
    $counter = 0

    Write-Host "Ditemukan $totalUsers pengguna. Memulai pemindaian perangkat..." -ForegroundColor White

    foreach ($user in $allUsers) {
        $counter++
        
        # Tampilan progres di satu baris (UI Refresh)
        $statusText = "-> [$counter/$totalUsers] Memproses: $($user.UserPrincipalName) . . ."
        Write-Host "`r$statusText" -ForegroundColor Green -NoNewline

        try {
            # Mengambil perangkat milik user tersebut
            $ownedDevices = Get-EntraUserOwnedDevice -UserId $user.Id -All -ErrorAction SilentlyContinue

            if ($ownedDevices) {
                foreach ($device in $ownedDevices) {
                    $obj = [PSCustomObject]@{
                        UserPrincipalName = $user.UserPrincipalName
                        UserDisplayName   = $user.DisplayName
                        DeviceDisplayName = $device.DisplayName
                        DeviceId          = $device.DeviceId
                        OperatingSystem   = $device.OperatingSystem
                        OSVersion         = $device.OperatingSystemVersion
                        AccountEnabled    = $device.AccountEnabled
                        TrustType         = $device.TrustType
                    }
                    [void]$scriptOutput.Add($obj)
                }
            }
        } catch {
            # Mengabaikan error jika satu user gagal diakses (misal: permission)
            continue
        }
    }
    Write-Host "`n`nPemrosesan selesai." -ForegroundColor Green
} catch {
    Write-Error "Terjadi kesalahan sistem: $($_.Exception.Message)"
}

## -----------------------------------------------------------------------
## 4. CLEANUP, DISCONNECT, DAN EKSPOR HASIL
## -----------------------------------------------------------------------

Write-Host "`n--- 4. Cleanup, Memutus Koneksi, dan Ekspor Hasil ---" -ForegroundColor Blue

# 4.1. Ekspor Hasil
if ($scriptOutput.Count -gt 0) {
    Write-Host "Mengekspor $($scriptOutput.Count) data perangkat pengguna..." -ForegroundColor Yellow
    try {
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction Stop
        Write-Host " Data berhasil diekspor ke: $outputFilePath" -ForegroundColor Green
    }
    catch {
        Write-Error "Gagal mengekspor data ke CSV: $($_.Exception.Message)"
    }
} else {
    Write-Host " Tidak ada perangkat ditemukan (\$scriptOutput kosong). Melewati ekspor." -ForegroundColor DarkYellow
}

# 4.2. Memutus koneksi Entra
Write-Host "Memutuskan koneksi dari Microsoft Entra..." -ForegroundColor DarkYellow
Disconnect-Entra -ErrorAction SilentlyContinue

Write-Host "`nSkrip $($scriptName) selesai dieksekusi." -ForegroundColor Yellow