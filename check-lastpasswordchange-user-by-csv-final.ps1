<#
.SYNOPSIS
Membuat laporan tentang tanggal terakhir pengguna mengganti password berdasarkan daftar UserPrincipalName dari file CSV TANPA HEADER.
.DESCRIPTION
Skrip dimodifikasi untuk membaca file CSV yang langsung berisi data email di baris pertama.
#>

# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.1 - No Header)
# =========================================================================

# Variabel Global dan Output
$scriptName = "PasswordChangeReport" 
$scriptOutput = @() 
$inputFileName = "UserPrincipalName.csv"

# Penanganan Jalur Aman (Fix: Empty Path Error)
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$inputFilePath = Join-Path -Path $scriptDir -ChildPath $inputFileName

# Tentukan jalur output
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

## -----------------------------------------------------------------------
## 2. KONEKSI WAJIB (MICROSOFT GRAPH)
## -----------------------------------------------------------------------

Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Graph ---" -ForegroundColor Blue

if (-not (Get-MgContext)) {
    Write-Host "Menghubungkan ke Microsoft Graph (Scope: User.Read.All)..." -ForegroundColor Yellow
    try {
        Connect-MgGraph -Scopes "User.Read.All" -ErrorAction Stop | Out-Null
        Write-Host "Koneksi ke Microsoft Graph berhasil dibuat." -ForegroundColor Green
    } catch {
        Write-Error "Gagal terhubung ke Microsoft Graph."
        exit 1
    }
}

## -----------------------------------------------------------------------
## 3. LOGIKA UTAMA SCRIPT
## -----------------------------------------------------------------------

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

if (-not (Test-Path -Path $inputFilePath)) {
    Write-Error "File input CSV tidak ditemukan di: $inputFilePath"
} else {
    Write-Host "Memuat data dari '$inputFileName' (Mode: No Header)..." -ForegroundColor Cyan
    
    # MODIFIKASI: Menggunakan -Header untuk membaca file yang tidak punya judul kolom
    # Ini membuat baris pertama di CSV dianggap sebagai data 'UserPrincipalName'
    $users = Import-Csv -Path $inputFilePath -Header "UserPrincipalName" -ErrorAction SilentlyContinue
    
    $totalUsers = $users.Count
    $userCount = 0
    
    Write-Host "Total $($totalUsers) data ditemukan." -ForegroundColor Yellow

    foreach ($entry in $users) {
        $userCount++
        
        # Ambil data dan bersihkan spasi jika ada
        $upn = if ($entry.UserPrincipalName) { $entry.UserPrincipalName.Trim() } else { $null }
        
        # Skip jika baris kosong
        if ([string]::IsNullOrWhiteSpace($upn)) { continue }

        Write-Progress -Activity "Generating Password Change Report" `
                       -Status "Processing User $userCount of ${totalUsers}: $($upn)" `
                       -PercentComplete ([int](($userCount / $totalUsers) * 100))
        
        Write-Host "-> Memproses ${userCount}: $($upn)..." -ForegroundColor White
        
        try {
            $user = Get-MgUser -UserId $upn -Property DisplayName, LastPasswordChangeDateTime -ErrorAction Stop
            
            $lastChangeRaw = $user.LastPasswordChangeDateTime
            $lastChangeWIBString = "N/A"
            $lastChangeUTCString = "N/A"
            $lastChangeWIB = "N/A"

            if ($lastChangeRaw) {
                $dateTimeUTC = [System.DateTime]::SpecifyKind($lastChangeRaw, [System.DateTimeKind]::Utc)
                $wibTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("SE Asia Standard Time")
                $dateTimeWIB = [System.TimeZoneInfo]::ConvertTimeFromUtc($dateTimeUTC, $wibTimeZone)
                
                $lastChangeWIBString = $dateTimeWIB.ToString("yyyy-MM-dd HH:mm:ss") + " WIB"
                $lastChangeWIB = $dateTimeWIB.ToString("yyyy-MM-dd HH:mm:ss")
                $lastChangeUTCString = $dateTimeUTC.ToString("yyyy-MM-dd HH:mm:ss")
            }
            
            $scriptOutput += [PSCustomObject]@{
                UserPrincipalName     = $upn
                DisplayName           = $user.DisplayName
                LastPasswordChangeUTC = $lastChangeUTCString
                LastPasswordChangeWIB = $lastChangeWIB
                Status                = "SUCCESS"
            }
            Write-Host "Last Logon berhasil dikonversi ke WIB." -ForegroundColor DarkGreen
        } 
        catch {
            $errMsg = $_.Exception.Message
            Write-Host "   ERROR: $($errMsg)" -ForegroundColor Red
            $scriptOutput += [PSCustomObject]@{
                UserPrincipalName = $upn; Status = "FAIL"; Reason = $errMsg
            }
        }
    }
}

## -----------------------------------------------------------------------
## 4. EKSPOR HASIL
## -----------------------------------------------------------------------

if ($scriptOutput.Count -gt 0) {
    Write-Host "`nMengekspor hasil ke CSV..." -ForegroundColor Yellow
    $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
    Write-Host "Selesai! File: $outputFilePath" -ForegroundColor Green
}

# Logout otomatis jika diperlukan
# Disconnect-MgGraph -ErrorAction SilentlyContinue 

Write-Host "`nSkrip Selesai." -ForegroundColor Yellow