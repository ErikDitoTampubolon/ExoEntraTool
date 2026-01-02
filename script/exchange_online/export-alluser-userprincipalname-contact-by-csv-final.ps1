# =========================================================================
# AUTHOR: Erik Dito Tampubolon - TelkomSigma
# VERSION: 2.9 (UI Enhanced Output)
# Deskripsi: Fix ParserError & Support No Header CSV dengan Output Progres Hijau.
# =========================================================================

# Variabel Global dan Output
$scriptName = "AllActiveUsersDnUpnContactByCSVReport" 
$scriptOutput = [System.Collections.ArrayList]::new() 
$global:moduleStep = 1

# Konfigurasi File Input
$inputFileName = "UserPrincipalName.csv"

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptDir = if ($PSScriptRoot) {$PSScriptRoot} else {(Get-Location).Path}
$inputFilePath = Join-Path -Path $scriptDir -ChildPath $inputFileName
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

# ==========================================================================
#                            INFORMASI SCRIPT                
# ==========================================================================

Write-Host "`n================================================" -ForegroundColor Yellow
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Yellow
Write-Host " Nama Skrip        : UserContactReport_Final_Fixed" -ForegroundColor Yellow
Write-Host " Field Kolom       : [InputUser]
                     [DisplayName]
                     [UPN]
                     [Contact]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk membuat laporan kontak pengguna Microsoft Entra ID berdasarkan daftar UPN dari file CSV tanpa header. Script menampilkan progres eksekusi di konsol, mengambil informasi DisplayName, UPN, serta nomor telepon (BusinessPhones dan MobilePhone), lalu mengekspor hasil ke file CSV." -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Yellow

## ==========================================================================
##                          KONFIRMASI EKSEKUSI
## ==========================================================================

$confirmation = Read-Host "Apakah Anda ingin menjalankan skrip ini? (Y/N)"

if ($confirmation -ne "Y") {
    Write-Host "`nEksekusi skrip dibatalkan oleh pengguna." -ForegroundColor Red
    return
}

## ==========================================================================
##                          PRASYARAT DAN KONEKSI
## ==========================================================================

Write-Host "--- 1. Menyiapkan Lingkungan ---" -ForegroundColor Blue 
if (-not (Get-MgContext -ErrorAction SilentlyContinue)) { 
    Connect-MgGraph -Scopes "User.Read.All" -ErrorAction Stop | Out-Null
}

## ==========================================================================
##                          LOGIKA UTAMA SCRIPT
## ==========================================================================

Write-Host "`n--- 2. Memulai Logika Utama Skrip ---" -ForegroundColor Magenta 

if (-not (Test-Path $inputFilePath)) {
    Write-Host " ERROR: File '$inputFileName' tidak ditemukan!" -ForegroundColor Red
    exit 1
}

# Membaca CSV dengan Header manual karena file asli tidak memiliki judul kolom
$csvData = Import-Csv -Path $inputFilePath -Header "Email" -ErrorAction SilentlyContinue

if ($null -eq $csvData -or $csvData.Count -eq 0) {
    Write-Host " ERROR: File CSV kosong." -ForegroundColor Red
    exit 1
}

$totalData = $csvData.Count
$i = 0

foreach ($row in $csvData) {
    $i++
    
    # Ambil nilai email
    $targetUser = if ($row.Email) { $row.Email.Trim() } else { $null }
    
    if ([string]::IsNullOrWhiteSpace($targetUser)) { continue }

    # FORMAT OUTPUT SESUAI PERMINTAAN: -> [i/total] Memproses: email@domain.com
    Write-Host "-> [$($i)/$($totalData)] Memproses: $($targetUser)" -ForegroundColor White

    try {
        $userObj = Get-MgUser -UserId $targetUser -Property "UserPrincipalName","DisplayName","BusinessPhones","MobilePhone" -ErrorAction Stop
        
        $phones = @()
        if ($userObj.BusinessPhones) { $phones += ($userObj.BusinessPhones -join ", ") }
        if ($userObj.MobilePhone) { $phones += $userObj.MobilePhone }
        
        $contactInfo = if ($phones.Count -gt 0) { $phones -join " | " } else { "-" }

        [void]$scriptOutput.Add([PSCustomObject]@{
            InputUser   = $targetUser
            DisplayName = $userObj.DisplayName
            UPN         = $userObj.UserPrincipalName
            Contact     = $contactInfo
        })
    }
    catch {
        [void]$scriptOutput.Add([PSCustomObject]@{
            InputUser   = $targetUser
            DisplayName = "NOT FOUND"
            UPN         = "-"
            Contact     = "-"
        })
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