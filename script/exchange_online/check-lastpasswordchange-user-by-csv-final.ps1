# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.3)
# =========================================================================

# Variabel Global dan Output
$scriptName = "PasswordChangeByCSVReport" 
$scriptOutput = @() 
$inputFileName = "UserPrincipalName.csv"

# Penanganan Jalur Aman
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$inputFilePath = Join-Path -Path $scriptDir -ChildPath $inputFileName

# Tentukan jalur output
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFileName = "Output_${scriptName}_${timestamp}.csv"
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

# ==========================================================================
#                               INFORMASI SCRIPT                
# ==========================================================================
Write-Host "`n================================================" -ForegroundColor Yellow
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Yellow
Write-Host " Nama Skrip        : PasswordChangeReport" -ForegroundColor Yellow
Write-Host " Field Kolom       : [UserPrincipalName]
                     [DisplayName]
                     [LastPasswordChangeWIB]
                     [Status]
                     [Reason]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk membuat laporan tanggal terakhir pengguna mengganti password berdasarkan daftar UPN dari file CSV tanpa header. Script akan menghubungkan ke Microsoft Graph, mengambil atribut LastPasswordChangeDateTime, mengonversinya ke zona waktu WIB, menampilkan progres eksekusi di konsol, serta mengekspor hasil ke file CSV." -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Yellow

# ==========================================================================
#                               KONFIRMASI EKSEKUSI
# ==========================================================================
$confirmation = Read-Host "Apakah Anda ingin menjalankan skrip ini? (Y/N)"

if ($confirmation -ne "Y") {
    Write-Host "`nEksekusi skrip dibatalkan oleh pengguna." -ForegroundColor Red
    return
}

## ==========================================================================
##                          KONEKSI KE MICROSOFT GRAPH
## ==========================================================================


if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
    Write-Host "`n--- Membangun Koneksi ke Microsoft Graph ---" -ForegroundColor Blue
    try {
        Connect-MgGraph -Scopes "User.Read.All" -ErrorAction Stop | Out-Null
    } catch {
        Write-Error "Gagal terhubung ke Microsoft Graph."
        exit 1
    }
}

## ==========================================================================
##                              LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 3. Memulai Logika Utama Skrip: ${scriptName} ---" -ForegroundColor Magenta

if (-not (Test-Path -Path $inputFilePath)) {
    Write-Error "File input CSV tidak ditemukan!"
} else {
    $users = Import-Csv -Path $inputFilePath -Header "UserPrincipalName" -ErrorAction SilentlyContinue
    $totalUsers = $users.Count
    $userCount = 0
    
    foreach ($entry in $users) {
        $userCount++
        $upn = if ($entry.UserPrincipalName) { $entry.UserPrincipalName.Trim() } else { $null }
        if ([string]::IsNullOrWhiteSpace($upn)) { continue }

        Write-Progress -Activity "Generating Report" -Status "User ${userCount}/${totalUsers}"

        # TAMPILAN KONSOL (SESUAI GAMBAR)
        Write-Host "-> [${userCount}/${totalUsers}] Memproses: ${upn}..." -ForegroundColor White
        
        try {
            $user = Get-MgUser -UserId $upn -Property DisplayName, LastPasswordChangeDateTime -ErrorAction Stop
            
            $lastChangeRaw = $user.LastPasswordChangeDateTime
            $lastChangeWIB = "N/A"

            if ($lastChangeRaw) {
                $dateTimeUTC = [System.DateTime]::SpecifyKind($lastChangeRaw, [System.DateTimeKind]::Utc)
                $wibTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("SE Asia Standard Time")
                $dateTimeWIB = [System.TimeZoneInfo]::ConvertTimeFromUtc($dateTimeUTC, $wibTimeZone)
                $lastChangeWIB = $dateTimeWIB.ToString("yyyy-MM-dd HH:mm:ss")
            }
            
            # OUTPUT BARIS KEDUA (WARNA HIJAU SESUAI GAMBAR)
            Write-Host "Last Password Changes: ${lastChangeWIB}" -ForegroundColor Green
            
            $scriptOutput += [PSCustomObject]@{
                UserPrincipalName     = $upn
                DisplayName           = $user.DisplayName
                LastPasswordChangeWIB = $lastChangeWIB
                Status                = "SUCCESS"
            }
        } 
        catch {
            Write-Host "   Gagal mengambil data." -ForegroundColor Red
            $scriptOutput += [PSCustomObject]@{
                UserPrincipalName = $upn; Status = "FAIL"; Reason = $_.Exception.Message
            }
        }
    }
}

## ==========================================================================
##                              EKSPOR HASIL
## ==========================================================================

if ($scriptOutput.Count -gt 0) {
    # 1. Tentukan nama folder
    $exportFolderName = "exported_data"
    
    # 2. Ambil jalur dua tingkat di atas direktori skrip
    # Contoh: Jika skrip di C:\Users\Erik\Project\Scripts, maka ini ke C:\Users\Erik\
    $parentDir = (Get-Item $scriptDir).Parent.Parent.FullName
    
    # 3. Gabungkan menjadi jalur folder ekspor
    $exportFolderPath = Join-Path -Path $parentDir -ChildPath $exportFolderName

    # 4. Cek apakah folder 'exported_data' sudah ada di lokasi tersebut, jika belum buat baru
    if (-not (Test-Path -Path $exportFolderPath)) {
        New-Item -Path $exportFolderPath -ItemType Directory | Out-Null
        Write-Host "`nFolder '$exportFolderName' berhasil dibuat di: $parentDir" -ForegroundColor Yellow
    }

    # 5. Tentukan nama file dan jalur lengkap
    $outputFileName = "Output_$($scriptName)_$($timestamp).csv"
    $resultsFilePath = Join-Path -Path $exportFolderPath -ChildPath $outputFileName
    
    # 6. Ekspor data
    $scriptOutput | Export-Csv -Path $resultsFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
    
    Write-Host "`nSemua proses selesai!" -ForegroundColor Green
    Write-Host "Laporan tersimpan di: ${resultsFilePath}" -ForegroundColor Cyan
}