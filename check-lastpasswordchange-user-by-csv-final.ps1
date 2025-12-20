<#
.SYNOPSIS
Membuat laporan tentang tanggal terakhir pengguna mengganti password berdasarkan daftar UPN dari file CSV TANPA HEADER.
#>

# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.3)
# =========================================================================

# Variabel Global dan Output
$scriptName = "PasswordChangeReport" 
$scriptOutput = @() 
$inputFileName = "UserPrincipalName.csv"

# Penanganan Jalur Aman
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$inputFilePath = Join-Path -Path $scriptDir -ChildPath $inputFileName

# Tentukan jalur output
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFileName = "Output_${scriptName}_${timestamp}.csv"
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

## -----------------------------------------------------------------------
## 2. KONEKSI WAJIB (MICROSOFT GRAPH)
## -----------------------------------------------------------------------

if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
    Write-Host "`n--- Membangun Koneksi ke Microsoft Graph ---" -ForegroundColor Blue
    try {
        Connect-MgGraph -Scopes "User.Read.All" -ErrorAction Stop | Out-Null
    } catch {
        Write-Error "Gagal terhubung ke Microsoft Graph."
        exit 1
    }
}


## -----------------------------------------------------------------------
## 3. LOGIKA UTAMA SCRIPT
## -----------------------------------------------------------------------

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
            Write-Host "   ❌ Gagal mengambil data." -ForegroundColor Red
            $scriptOutput += [PSCustomObject]@{
                UserPrincipalName = $upn; Status = "FAIL"; Reason = $_.Exception.Message
            }
        }
    }
}



## -----------------------------------------------------------------------
## 4. EKSPOR HASIL
## -----------------------------------------------------------------------

if ($scriptOutput.Count -gt 0) {
    try {
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
        Write-Host "`n✅ Laporan Berhasil Diekspor: $outputFilePath" -ForegroundColor Cyan
    } catch {
        Write-Host "`n⚠️ Gagal mengekspor CSV." -ForegroundColor Yellow
    }
}

Write-Host "`nSkrip Selesai." -ForegroundColor Yellow