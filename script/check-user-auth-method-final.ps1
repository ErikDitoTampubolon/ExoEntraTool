# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Nama Skrip: Get-EntraUserAuthMethods
# Deskripsi: Menarik daftar metode autentikasi user dari CSV.
# =========================================================================

# Variabel Global dan Output
$scriptName = "GetAuthMethods" 
$scriptOutput = New-Object System.Collections.Generic.List[PSCustomObject]

# Konfigurasi File Input (Pastikan file ini ada di folder yang sama dengan skrip)
$inputFileName = "UserPrincipalName.csv"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$inputFilePath = Join-Path -Path $scriptDir -ChildPath $inputFileName

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName


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
    param([string]$ModuleName)
    Write-Host "1.$(++$global:moduleStep). Memeriksa Modul '$ModuleName'..." -ForegroundColor Cyan
    if (Get-Module -Name $ModuleName -ListAvailable) {
        Write-Host " Modul '$ModuleName' sudah terinstal." -ForegroundColor Green
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
CheckAndInstallModule -ModuleName "Microsoft.Entra"

## -----------------------------------------------------------------------
## 2. KONEKSI WAJIB (MICROSOFT ENTRA)
## -----------------------------------------------------------------------

Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Entra ---" -ForegroundColor Blue

try {
    Write-Host "Menghubungkan ke Microsoft Entra..." -ForegroundColor Yellow
    # Menghubungkan dengan scope yang diperlukan untuk membaca metode autentikasi
    Connect-Entra -Scopes 'UserAuthenticationMethod.Read.All' -ErrorAction Stop
    Write-Host "Koneksi ke Microsoft Entra berhasil." -ForegroundColor Green
} catch {
    Write-Error "Gagal terhubung ke Microsoft Entra. $($_.Exception.Message)"
    exit 1
}

## -----------------------------------------------------------------------
## 3. LOGIKA UTAMA SCRIPT
## -----------------------------------------------------------------------

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

if (Test-Path $inputFilePath) {
    # Mengimpor CSV tanpa header (menggunakan TempColumn) atau pastikan header bernama 'UserPrincipalName'
    $users = Import-Csv $inputFilePath -Header "TempColumn" | Where-Object { $_.TempColumn -ne $null -and $_.TempColumn.Trim() -ne "" }
    $totalUsers = $users.Count
    $counter = 0

    foreach ($row in $users) {
        $counter++
        $upn = $row.TempColumn.Trim()
        
        # UI Progres Baris Tunggal
        $statusText = "-> [$counter/$totalUsers] Memproses: $upn . . ."
        Write-Host "`r$statusText" -ForegroundColor Green -NoNewline

        try {
            # Mengambil metode autentikasi untuk user terkait
            $authMethods = Get-EntraUserAuthenticationMethod -UserId $upn -ErrorAction Stop
            
            foreach ($method in $authMethods) {
                $obj = [PSCustomObject]@{
                    UserPrincipalName        = $upn
                    AuthenticationMethodId   = $method.Id
                    DisplayName              = $method.DisplayName
                    AuthenticationMethodType = $method.AuthenticationMethodType
                    Status                   = "SUCCESS"
                }
                $scriptOutput.Add($obj)
            }
            
            # Jika user tidak memiliki metode yang terdaftar sama sekali
            if ($null -eq $authMethods) {
                $scriptOutput.Add([PSCustomObject]@{
                    UserPrincipalName = $upn
                    Status            = "NO_METHODS_FOUND"
                })
            }

        } catch {
            $scriptOutput.Add([PSCustomObject]@{
                UserPrincipalName = $upn
                Status            = "FAILED"
                ErrorMessage      = $_.Exception.Message
            })
        }
    }
    Write-Host "`n`nPemrosesan data selesai." -ForegroundColor Green
} else {
    Write-Host "ERROR: File '$inputFileName' tidak ditemukan di folder skrip!" -ForegroundColor Red
}

## -----------------------------------------------------------------------
## 4. CLEANUP, DISCONNECT, DAN EKSPOR HASIL
## -----------------------------------------------------------------------

Write-Host "`n--- 4. Cleanup dan Ekspor Hasil ---" -ForegroundColor Blue

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
    Write-Host " Tidak ada data yang dikumpulkan." -ForegroundColor DarkYellow
}

# 4.2. Memutus koneksi Entra
Write-Host "Memutuskan koneksi dari Microsoft Entra..." -ForegroundColor DarkYellow
Disconnect-Entra -ErrorAction SilentlyContinue
Write-Host " Sesi Microsoft Entra diputus." -ForegroundColor Green

Write-Host "`nSkrip $($scriptName) selesai dieksekusi." -ForegroundColor Yellow