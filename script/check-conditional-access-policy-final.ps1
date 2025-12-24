# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Nama Skrip: Get-EntraConditionalAccessPolicies
# Deskripsi: Menarik semua kebijakan Conditional Access ke file CSV.
# =========================================================================

# Variabel Global dan Output
$scriptName = "GetEntraCAPolicies" 
$scriptOutput = New-Object System.Collections.Generic.List[PSCustomObject]

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
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
        Write-Host " Modul '$ModuleName' belum ditemukan. Menginstal..." -ForegroundColor Yellow
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
    Write-Host "Menghubungkan ke Microsoft Entra dengan scope Policy.Read.All..." -ForegroundColor Yellow
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
    Write-Host "Mengambil data Conditional Access Policies..." -ForegroundColor Cyan
    $policies = Get-EntraConditionalAccessPolicy -ErrorAction Stop
    $total = $policies.Count
    $counter = 0

    foreach ($policy in $policies) {
        $counter++
        # Tampilan progres baris tunggal
        $statusText = "-> [$counter/$total] Memproses: $($policy.DisplayName) . . ."
        Write-Host "`r$statusText" -ForegroundColor Green -NoNewline

        # Mengumpulkan data ke dalam objek
        $obj = [PSCustomObject]@{
            Id           = $policy.Id
            DisplayName  = $policy.DisplayName
            State        = $policy.State
            CreatedTime  = $policy.CreatedDateTime
            ModifiedTime = $policy.ModifiedDateTime
        }
        $scriptOutput.Add($obj)
    }
    Write-Host "`n`nData berhasil dikumpulkan." -ForegroundColor Green
} catch {
    Write-Error "Gagal mengambil kebijakan: $($_.Exception.Message)"
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
    Write-Host " Tidak ada data yang ditemukan." -ForegroundColor DarkYellow
}

# 4.2. Memutus koneksi Entra
Write-Host "Memutuskan koneksi dari Microsoft Entra..." -ForegroundColor DarkYellow
Disconnect-Entra -ErrorAction SilentlyContinue
Write-Host " Sesi Microsoft Entra diputus." -ForegroundColor Green

Write-Host "`nSkrip $($scriptName) selesai dieksekusi." -ForegroundColor Yellow