# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.1)
# Menyimpan laporan detail semua Pengguna Aktif & No Telepon ke CSV.
# =========================================================================

# Variabel Global dan Output
$scriptName = "ExportActiveUsersContactReport" 
$scriptOutput = @() 
$global:moduleStep = 0

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptDir = if ($PSScriptRoot) {$PSScriptRoot} else {(Get-Location).Path}
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

# 1.2. Fungsi Pembantu untuk Cek dan Instal Modul 
function CheckAndInstallModule { 
    param( [string]$ModuleName ) 
    Write-Host "1.$(++$global:moduleStep). Memeriksa Modul '$ModuleName'..." -ForegroundColor Cyan 
    if (Get-Module -Name $ModuleName -ListAvailable) { 
        Write-Host " ✅ Modul '$ModuleName' sudah terinstal." -ForegroundColor Green 
    } else { 
        Write-Host " ⚠️ Modul '$ModuleName' belum ditemukan. Menginstal..." -ForegroundColor Yellow 
        Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
        Write-Host " ✅ Modul '$ModuleName' berhasil diinstal." -ForegroundColor Green 
    } 
} 

CheckAndInstallModule -ModuleName "Microsoft.Graph" 

## -----------------------------------------------------------------------
## 2. KONEKSI WAJIB (MICROSOFT GRAPH)
## -----------------------------------------------------------------------

Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Graph ---" -ForegroundColor Blue 

if (-not (Get-MgContext -ErrorAction SilentlyContinue)) { 
    Write-Host "Menghubungkan ke Microsoft Graph..." -ForegroundColor Yellow 
    try { 
        # Menambahkan scopes yang diperlukan
        $scopes = @("User.Read.All", "Directory.Read.All", "Organization.Read.All")
        Connect-MgGraph -Scopes $scopes -ErrorAction Stop | Out-Null
        Write-Host "✅ Koneksi Berhasil." -ForegroundColor Green 
    } catch { 
        Write-Error "Gagal login: $($_.Exception.Message)" 
        exit 1 
    } 
} else { 
    Write-Host "✅ Sesi sudah aktif." -ForegroundColor Green 
} 


## -----------------------------------------------------------------------
## 3. LOGIKA UTAMA SCRIPT: EXPORT ACTIVE USERS + CONTACTS
## -----------------------------------------------------------------------

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

try {
    # 3.1.1. Muat SkuMap untuk Lisensi
    $skuMap = @{}
    $allSkus = Get-MgSubscribedSku -ErrorAction SilentlyContinue
    if ($allSkus) { foreach ($sku in $allSkus) { $skuMap.Add($sku.SkuPartNumber, $sku.SkuName) } }

    # 3.1.2. Ambil Pengguna dengan Properti Telepon
    # Menambahkan BusinessPhones dan MobilePhone ke pemilihan properti
    $selectProperties = @(
        "Id", "UserPrincipalName", "DisplayName", "Mail", "JobTitle", 
        "Department", "UsageLocation", "AccountEnabled", "CreatedDateTime",
        "BusinessPhones", "MobilePhone"
    )

    Write-Host "Mengambil data pengguna dan nomor telepon..." -ForegroundColor Cyan
    $activeUsers = Get-MgUser -Filter "accountEnabled eq true" -All -Property $selectProperties -ErrorAction Stop
    $totalUsers = $activeUsers.Count
    
    $i = 0
    foreach ($user in $activeUsers) {
        $i++
        Write-Progress -Activity "Collecting Contact Info" -Status "User: $($user.UserPrincipalName)" -PercentComplete ([int](($i / $totalUsers) * 100))

        # Proses Lisensi
        $licensesString = "None Assigned"
        try {
            $userLicenses = Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction SilentlyContinue
            $friendlyNames = foreach ($lic in $userLicenses) { if ($skuMap.ContainsKey($lic.SkuPartNumber)) { $skuMap[$lic.SkuPartNumber] } else { $lic.SkuPartNumber } }
            if ($friendlyNames) { $licensesString = $friendlyNames -join "; " }
        } catch { $licensesString = "N/A" }

        # Proses Nomor Telepon (BusinessPhones adalah array, jadi kita gabung dengan koma)
        $officePhone = if ($user.BusinessPhones) { $user.BusinessPhones -join ", " } else { "-" }
        $mobilePhone = if ($user.MobilePhone) { $user.MobilePhone } else { "-" }

        # Masukkan ke Output
        $scriptOutput += [PSCustomObject]@{
            UserPrincipalName  = $user.UserPrincipalName
            DisplayName        = $user.DisplayName
            EmailAddress       = $user.Mail
            OfficePhone        = $officePhone  # Kolom Baru
            MobilePhone        = $mobilePhone  # Kolom Baru
            JobTitle           = $user.JobTitle
            Department         = $user.Department
            UsageLocation      = $user.UsageLocation
            LicenseSKUs        = $licensesString
            AccountCreatedDate = if ($user.CreatedDateTime) { $user.CreatedDateTime.ToShortDateString() } else { "N/A" }
        }
    }
}
catch {
    Write-Error "Terjadi kesalahan: $($_.Exception.Message)"
}


## -----------------------------------------------------------------------
## 4. CLEANUP & EKSPOR
## -----------------------------------------------------------------------

Write-Host "`n--- 4. Ekspor Hasil ---" -ForegroundColor Blue

if ($scriptOutput.Count -gt 0) {
    try {
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction Stop
        Write-Host " ✅ Laporan Kontak berhasil disimpan: $outputFilePath" -ForegroundColor Green
    }
    catch { Write-Error "Gagal simpan CSV: $($_.Exception.Message)" }
}

if (Get-MgContext -ErrorAction SilentlyContinue) {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Write-Host " ✅ Sesi Microsoft Graph diputus." -ForegroundColor Green
}

Write-Host "`nSkrip Selesai." -ForegroundColor Yellow