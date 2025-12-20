# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.5)
# Deskripsi: Mengambil DisplayName, UPN, dan Kontak dari SEMUA User Aktif.
# =========================================================================

# Variabel Global dan Output
$scriptName = "AllActiveUsersContactReport" 
$scriptOutput = @() 
$global:moduleStep = 1

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptDir = if ($PSScriptRoot) {$PSScriptRoot} else {(Get-Location).Path}
$outputFileName = "Output_${scriptName}_${timestamp}.csv"
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

## -----------------------------------------------------------------------
## 1. PRASYARAT DAN INSTALASI MODUL
## -----------------------------------------------------------------------
Write-Host "--- 1. Memeriksa dan Menyiapkan Lingkungan PowerShell ---" -ForegroundColor Blue 

# 1.1. Mengatur Execution Policy 
Write-Host "1.1. Mengatur Execution Policy ke RemoteSigned..." -ForegroundColor Cyan 
try { 
    Set-ExecutionPolicy RemoteSigned -Scope Process -Force -ErrorAction Stop 
    Write-Host " ✅ Execution Policy berhasil diatur." -ForegroundColor Green 
} catch { 
    Write-Error "Gagal mengatur Execution Policy: $($_.Exception.Message)" 
    exit 1 
} 

# 1.2. Fungsi Pembantu Modul 
function CheckAndInstallModule { 
    param( [Parameter(Mandatory=$true)] [string]$ModuleName ) 
    Write-Host "1.$(++$global:moduleStep). Memeriksa Modul '$ModuleName'..." -ForegroundColor Cyan 
    if (Get-Module -Name $ModuleName -ListAvailable) { 
        Write-Host " ✅ Modul '$ModuleName' sudah terinstal." -ForegroundColor Green 
    } else { 
        Write-Host " ⚠️ Modul '$ModuleName' belum ditemukan. Memulai instalasi..." -ForegroundColor Yellow 
        try { 
            Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop 
            Write-Host " ✅ Modul '$ModuleName' berhasil diinstal." -ForegroundColor Green 
        } catch { 
            Write-Error "Gagal menginstal modul '$ModuleName'." 
            exit 1 
        } 
    } 
} 

CheckAndInstallModule -ModuleName "Microsoft.Graph" 

## -----------------------------------------------------------------------
## 2. KONEKSI WAJIB (MICROSOFT GRAPH)
## -----------------------------------------------------------------------
Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Graph ---" -ForegroundColor Blue 

if (-not (Get-MgContext -ErrorAction SilentlyContinue)) { 
    Write-Host "Menghubungkan ke Microsoft Graph. Silakan login..." -ForegroundColor Yellow 
    try { 
        $scopes = @("User.Read.All")
        # Ditambahkan | Out-Null agar tidak tumpah detail login ke konsol
        Connect-MgGraph -Scopes $scopes -ErrorAction Stop | Out-Null
        Write-Host "✅ Koneksi Berhasil." -ForegroundColor Green 
    } catch { 
        Write-Error "Gagal terhubung: $($_.Exception.Message)" 
        exit 1 
    } 
} else { 
    Write-Host "✅ Sesi Microsoft Graph sudah aktif." -ForegroundColor Green 
} 

## -----------------------------------------------------------------------
## 3. LOGIKA UTAMA SCRIPT
## -----------------------------------------------------------------------
Write-Host "`n--- 3. Memulai Logika Utama Skrip: ${scriptName} ---" -ForegroundColor Magenta 

try {
    Write-Host "Mengambil data seluruh pengguna aktif..." -ForegroundColor Cyan
    
    $selectProperties = @("UserPrincipalName", "DisplayName", "BusinessPhones", "MobilePhone", "AccountEnabled")

    # Ambil semua user aktif
    $allUsers = Get-MgUser -Filter "accountEnabled eq true" -All -Property $selectProperties -ErrorAction Stop
    
    $totalUsers = $allUsers.Count
    Write-Host " ✅ Berhasil menemukan ${totalUsers} pengguna aktif." -ForegroundColor Green
    
    $i = 0
    foreach ($user in $allUsers) {
        $i++
        Write-Progress -Activity "Memproses Data Kontak" -Status "User: $($user.UserPrincipalName)" -PercentComplete ([int](($i / $totalUsers) * 100))

        # Gabungkan Nomor Telepon
        $phones = @()
        if ($user.BusinessPhones) { $phones += ($user.BusinessPhones -join ", ") }
        if ($user.MobilePhone) { $phones += $user.MobilePhone }
        
        $contactInfo = if ($phones.Count -gt 0) { $phones -join " | " } else { "-" }

        # MODIFIKASI: Urutan kolom sesuai permintaan: DisplayName, UPN, Contact
        $scriptOutput += [PSCustomObject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            Contact           = $contactInfo
        }
    }
}
catch {
    Write-Error "Kesalahan saat pengambilan data: $($_.Exception.Message)"
}

## -----------------------------------------------------------------------
## 4. EKSPOR HASIL
## -----------------------------------------------------------------------
Write-Host "`n--- 4. Cleanup dan Ekspor Hasil ---" -ForegroundColor Blue 

if ($scriptOutput.Count -gt 0) { 
    Write-Host "Mengekspor $($scriptOutput.Count) baris data..." -ForegroundColor Yellow 
    try { 
        # Delimiter titik koma (;) agar otomatis rapi saat dibuka di Excel Indonesia
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction Stop 
        Write-Host " ✅ Laporan berhasil disimpan di:" -ForegroundColor Green 
        Write-Host " $outputFilePath" -ForegroundColor Cyan 
    } 
    catch { 
        Write-Error "Gagal ekspor CSV: $($_.Exception.Message)" 
    } 
}

# Sesi tetap dibuka untuk integrasi menu utama Anda
Write-Host "`nSkrip ${scriptName} selesai dieksekusi." -ForegroundColor Yellow