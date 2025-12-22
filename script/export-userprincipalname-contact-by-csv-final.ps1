# =========================================================================
# AUTHOR: Erik Dito Tampubolon - TelkomSigma
# VERSION: 2.9 (UI Enhanced Output)
# Deskripsi: Fix ParserError & Support No Header CSV dengan Output Progres Hijau.
# =========================================================================

# Variabel Global dan Output
$scriptName = "UserContactReport_Final_Fixed" 
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

## -----------------------------------------------------------------------
## 1. PRASYARAT DAN KONEKSI
## -----------------------------------------------------------------------

Write-Host "--- 1. Menyiapkan Lingkungan ---" -ForegroundColor Blue 
if (-not (Get-MgContext -ErrorAction SilentlyContinue)) { 
    Connect-MgGraph -Scopes "User.Read.All" -ErrorAction Stop | Out-Null
}

## -----------------------------------------------------------------------
## 2. LOGIKA UTAMA SCRIPT
## -----------------------------------------------------------------------

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

## -----------------------------------------------------------------------
## 3. EKSPOR HASIL
## -----------------------------------------------------------------------

Write-Host "`n--- 3. Ekspor Hasil ---" -ForegroundColor Blue 

if ($scriptOutput.Count -gt 0) { 
    try {
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction Stop
        Write-Host "Berhasil! File tersimpan di: $outputFilePath" -ForegroundColor Green 
    } catch {
        Write-Error "Gagal ekspor: $($_.Exception.Message)"
    }
} 

# Disconnect otomatis tetap dijalankan melalui skrip utama (main app)
Write-Host "`nSkrip Selesai." -ForegroundColor Yellow