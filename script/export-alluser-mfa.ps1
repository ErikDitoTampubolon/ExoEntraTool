# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.1)
# Deskripsi: Mendapatkan status Per-User MFA dengan tampilan Progres Konsol
# =========================================================================

# Variabel Global dan Output
$scriptName = "Export-AllUser-MFA" 
$scriptOutput = [System.Collections.ArrayList]::new() 

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"
$outputFilePath = Join-Path -Path $PSScriptRoot -ChildPath $outputFileName


## -----------------------------------------------------------------------
## 2. KONEKSI WAJIB (MICROSOFT ENTRA)
## -----------------------------------------------------------------------

Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Entra ---" -ForegroundColor Blue

try {
    Write-Host "Menghubungkan ke Microsoft Entra. Selesaikan login pada pop-up..." -ForegroundColor Yellow
    
    # Menangani potensi konflik DLL dengan mencoba Disconnect terlebih dahulu
    Disconnect-Entra -ErrorAction SilentlyContinue
    
    # Koneksi utama
    Connect-Entra -Scopes 'User.Read.All', 'UserAuthenticationMethod.Read.All' -ErrorAction Stop
    Write-Host "Koneksi ke Microsoft Entra berhasil dibuat." -ForegroundColor Green
} catch {
    Write-Error "Gagal terhubung: $($_.Exception.Message)"
    Write-Host "`nTIP: Jika error library berlanjut, tutup SEMUA jendela PowerShell lalu buka kembali." -ForegroundColor Yellow
    exit 1
}


## -----------------------------------------------------------------------
## 3. LOGIKA UTAMA SCRIPT
## -----------------------------------------------------------------------

Write-Host "`n--- 3. Memulai Logika Utama Skrip ---" -ForegroundColor Magenta

try {
    $users = Get-EntraUser -All -Select Id, UserPrincipalName, DisplayName
    $totalUsers = $users.Count
    $counter = 1

    foreach ($u in $users) {
        # FORMAT TAMPILAN SESUAI PERMINTAAN: -> [1/24] Memproses: email@domain.com
        Write-Host "-> [$($counter)/$($totalUsers)] Memproses: $($u.UserPrincipalName)" -ForegroundColor Green
        
        $mfaState = "Unknown"
        try {
            $mfaReq = Get-EntraBetaUserAuthenticationRequirement -UserId $u.Id -ErrorAction Stop
            $mfaState = if ($null -ne $mfaReq.PerUserMFAState) { $mfaReq.PerUserMFAState } else { "Disabled" }
        } catch {
            $mfaState = "None/Disabled"
        }

        $userProperties = [PSCustomObject]@{
            Id                = $u.Id
            DisplayName       = $u.DisplayName
            UserPrincipalName = $u.UserPrincipalName
            PerUserMFAState   = $mfaState
        }

        [void]$scriptOutput.Add($userProperties)
        $counter++
    }
} catch {
    Write-Error "Terjadi kesalahan: $($_.Exception.Message)"
}

## -----------------------------------------------------------------------
## 4. CLEANUP DAN EKSPOR HASIL
## -----------------------------------------------------------------------

Write-Host "`n--- 4. Cleanup dan Ekspor Hasil ---" -ForegroundColor Blue

if ($scriptOutput.Count -gt 0) {
    # Menampilkan tabel ringkasan singkat di akhir
    $scriptOutput | Select-Object DisplayName, UserPrincipalName, PerUserMFAState | ft -AutoSize

    try {
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction Stop
        Write-Host "`nData berhasil diekspor ke: $outputFilePath" -ForegroundColor Green
    } catch {
        Write-Error "Gagal mengekspor CSV."
    }
}

Disconnect-Entra -ErrorAction SilentlyContinue
Write-Host "`nSkrip selesai dieksekusi." -ForegroundColor Yellow