# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Nama Skrip: Get-EntraAuthorizationPolicies
# Deskripsi: Mengambil detail kebijakan otorisasi Microsoft Entra ID.
# =========================================================================

# Variabel Global dan Output
$scriptName = "GetEntraAuthPolicy" 
$scriptOutput = @() 

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"
$outputFilePath = Join-Path -Path $PSScriptRoot -ChildPath $outputFileName


## -----------------------------------------------------------------------
## 1. PRASYARAT DAN INSTALASI MODUL
## -----------------------------------------------------------------------

Write-Host "--- 1. Memeriksa dan Menyiapkan Lingkungan PowerShell ---" -ForegroundColor Blue

# 1.1. Mengatur Execution Policy
Write-Host "1.1. Mengatur Execution Policy ke RemoteSigned..." -ForegroundColor Cyan
try {
    Set-ExecutionPolicy RemoteSigned -Scope Process -Force -ErrorAction Stop
    Write-Host " Execution Policy berhasil diatur." -ForegroundColor Green
} catch {
    Write-Error "Gagal mengatur Execution Policy: $($_.Exception.Message)"
    exit 1
}

# 1.2. Fungsi Pembantu untuk Cek dan Instal Modul
function CheckAndInstallModule {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ModuleName
    )

    Write-Host "1.$(++$global:moduleStep). Memeriksa Modul '$ModuleName'..." -ForegroundColor Cyan

    if (Get-Module -Name $ModuleName -ListAvailable) {
        Write-Host " Modul '$ModuleName' sudah terinstal. Melewati instalasi." -ForegroundColor Green
    } else {
        Write-Host " Modul '$ModuleName' belum ditemukan. Memulai instalasi..." -ForegroundColor Yellow
        try {
            Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
            Write-Host " Modul '$ModuleName' berhasil diinstal." -ForegroundColor Green
        } catch {
            Write-Error "Gagal menginstal modul '$ModuleName'."
            exit 1
        }
    }
}

$global:moduleStep = 1
CheckAndInstallModule -ModuleName "PowerShellGet"
# Menambahkan modul Microsoft.Entra sesuai kebutuhan cmdlet
CheckAndInstallModule -ModuleName "Microsoft.Entra"

## -----------------------------------------------------------------------
## 2. KONEKSI WAJIB (MICROSOFT ENTRA)
## -----------------------------------------------------------------------

Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Entra ---" -ForegroundColor Blue

try {
    Write-Host "Menghubungkan ke Microsoft Entra. Selesaikan login pada pop-up..." -ForegroundColor Yellow
    Connect-Entra -Scopes 'Policy.Read.All' -ErrorAction Stop
    Write-Host "Koneksi ke Microsoft Entra berhasil dibuat." -ForegroundColor Green
} catch {
    Write-Error "Gagal terhubung ke Microsoft Entra. $($_.Exception.Message)"
    exit 1
}

## -----------------------------------------------------------------------
## 3. LOGIKA UTAMA SCRIPT
## -----------------------------------------------------------------------

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

try {
    Write-Host "Mengambil data Authorization Policy..." -ForegroundColor Cyan
    
    # Menjalankan perintah utama
    $policies = Get-EntraAuthorizationPolicy -ErrorAction Stop
    
    if ($policies) {
        $total = $policies.Count
        $counter = 0

        foreach ($policy in $policies) {
            $counter++
            # Visualisasi progres baris tunggal
            Write-Host "`r-> [$counter/$total] Memproses: $($policy.DisplayName) . . ." -ForegroundColor Green -NoNewline
            
            # Memasukkan data ke array output
            $scriptOutput += [PSCustomObject]@{
                Id                           = $policy.Id
                DisplayName                  = $policy.DisplayName
                Description                  = $policy.Description
                AllowUserConsentForApp       = $policy.AllowedToUseSSPR
                BlockMsolIdp                 = $policy.BlockMsolIdp
                DefaultUserRolePermissions   = $policy.DefaultUserRolePermissions | ConvertTo-Json -Compress
            }
        }
        Write-Host "`nData berhasil dikumpulkan." -ForegroundColor Green
    }
} catch {
    Write-Error "Terjadi kesalahan saat mengambil kebijakan: $($_.Exception.Message)"
}

## -----------------------------------------------------------------------
## 4. CLEANUP, DISCONNECT, DAN EKSPOR HASIL
## -----------------------------------------------------------------------

Write-Host "`n--- 4. Cleanup, Memutus Koneksi, dan Ekspor Hasil ---" -ForegroundColor Blue

# 4.1. Ekspor Hasil
if ($scriptOutput.Count -gt 0) {
    Write-Host "Mengekspor $($scriptOutput.Count) baris data hasil skrip..." -ForegroundColor Yellow
    try {
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction Stop
        Write-Host " Data berhasil diekspor ke: $outputFilePath" -ForegroundColor Green
    }
    catch {
        Write-Error "Gagal mengekspor data ke CSV: $($_.Exception.Message)"
    }
} else {
    Write-Host " Tidak ada data yang dikumpulkan (\$scriptOutput kosong). Melewati ekspor." -ForegroundColor DarkYellow
}

# 4.2. Memutus koneksi Entra
Write-Host "Memutuskan koneksi dari Microsoft Entra..." -ForegroundColor DarkYellow
Disconnect-Entra -ErrorAction SilentlyContinue
Write-Host " Sesi Microsoft Entra diputus." -ForegroundColor Green

Write-Host "`nSkrip $($scriptName) selesai dieksekusi." -ForegroundColor Yellow