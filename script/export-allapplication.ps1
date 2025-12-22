# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Nama Skrip: Export-AllAppsOwnersReport
# Deskripsi: Mengambil daftar pemilik (owners) untuk SEMUA aplikasi di Entra.
# =========================================================================

# Variabel Global dan Output
$scriptName = "AllAppsOwnersReport" 
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
    Connect-Entra -Scopes 'Application.Read.All' -ErrorAction Stop
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
    Write-Host "Mengambil daftar semua aplikasi..." -ForegroundColor Cyan
    $allApps = Get-EntraApplication -All
    $totalApps = $allApps.Count
    $counter = 0

    Write-Host "Ditemukan $totalApps aplikasi. Memulai pengambilan data pemilik..." -ForegroundColor White

    foreach ($app in $allApps) {
        $counter++
        # Output progres hijau satu baris sesuai style yang Anda sukai
        Write-Host "-> [$counter/$totalApps] Memproses: $($app.DisplayName)" -ForegroundColor Green
        
        try {
            # Mengambil owner untuk aplikasi saat ini
            $owners = Get-EntraApplicationOwner -ApplicationId $app.Id -All -ErrorAction SilentlyContinue

            if ($owners) {
                foreach ($owner in $owners) {
                    $obj = [PSCustomObject]@{
                        ApplicationName   = $app.DisplayName
                        ApplicationId     = $app.AppId
                        OwnerObjectId     = $owner.Id
                        OwnerDisplayName  = $owner.DisplayName
                        UserPrincipalName = $owner.UserPrincipalName
                        CreatedDateTime   = $owner.CreatedDateTime
                        UserType          = $owner.UserType
                        AccountEnabled    = $owner.AccountEnabled
                    }
                    [void]$scriptOutput.Add($obj)
                }
            } else {
                # Jika aplikasi tidak memiliki owner, tetap catat dengan keterangan kosong
                $obj = [PSCustomObject]@{
                    ApplicationName   = $app.DisplayName
                    ApplicationId     = $app.AppId
                    OwnerObjectId     = "NO OWNER"
                    OwnerDisplayName  = "-"
                    UserPrincipalName = "-"
                    CreatedDateTime   = "-"
                    UserType          = "-"
                    AccountEnabled    = "-"
                }
                [void]$scriptOutput.Add($obj)
            }
        } catch {
            Write-Host "   Gagal mengambil owner untuk aplikasi: $($app.DisplayName)" -ForegroundColor Red
        }
    }
} catch {
    Write-Error "Terjadi kesalahan sistem: $($_.Exception.Message)"
}

## -----------------------------------------------------------------------
## 4. CLEANUP DAN EKSPOR HASIL
## -----------------------------------------------------------------------

Write-Host "`n--- 4. Cleanup dan Ekspor Hasil ---" -ForegroundColor Blue

if ($scriptOutput.Count -gt 0) {
    Write-Host "Mengekspor $($scriptOutput.Count) baris data ke CSV..." -ForegroundColor Yellow
    try {
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction Stop
        Write-Host "Data berhasil diekspor ke: $outputFilePath" -ForegroundColor Green
        
        # Tampilkan 10 baris pertama di konsol sebagai sampel
        Write-Host "`nSampel Hasil:" -ForegroundColor Gray
        $scriptOutput | Select-Object -First 10 | Format-Table -AutoSize
    } catch {
        Write-Error "Gagal mengekspor CSV."
    }
}

Disconnect-Entra -ErrorAction SilentlyContinue
Write-Host "`nSkrip selesai dieksekusi." -ForegroundColor Yellow