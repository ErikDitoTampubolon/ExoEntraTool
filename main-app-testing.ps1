# =========================================================================
# AUTHOR: Erik Dito Tampubolon - TelkomSigma
# VERSION: 2.3 (Fixed Path Binding & Integrated Framework)
# DESKRIPSI: Skrip Utama Exchange Online - Solusi Error 'Path is Null'.
# =========================================================================

# -------------------------------------------------------------------------
# 1. PENANGANAN JALUR (MENCEGAH ERROR BIND ARGUMENT)
# -------------------------------------------------------------------------
# Baris ini WAJIB di paling atas untuk memastikan $scriptDir tidak NULL
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

# Validasi tambahan: Jika masih null (sangat jarang), paksa ke folder aktif
if ([string]::IsNullOrWhiteSpace($scriptDir)) { $scriptDir = "." }

## -----------------------------------------------------------------------
## 2. PRASYARAT DAN INSTALASI MODUL
## -----------------------------------------------------------------------

Write-Host "--- 1. Memeriksa Lingkungan PowerShell ---" -ForegroundColor Blue

# Set Execution Policy
Set-ExecutionPolicy RemoteSigned -Scope Process -Force -ErrorAction SilentlyContinue

# Fungsi Helper Instalasi
function Check-Module {
    param($ModuleName)
    Write-Host "Memeriksa Modul '$ModuleName'..." -ForegroundColor Cyan
    if (Get-Module -Name $ModuleName -ListAvailable) {
        Write-Host "Terinstal." -ForegroundColor Green
    } else {
        Write-Host "Belum ada. Menginstal..." -ForegroundColor Yellow
        Install-Module $ModuleName -Force -AllowClobber -Scope CurrentUser -ErrorAction SilentlyContinue
    }
}

Check-Module -ModuleName "PowerShellGet"
Check-Module -ModuleName "ExchangeOnlineManagement"
Check-Module -ModuleName "Microsoft.Graph"

## -----------------------------------------------------------------------
## 3. KONEKSI EXCHANGE ONLINE
## -----------------------------------------------------------------------

Write-Host "`n--- 2. Membangun Koneksi ke Exchange Online ---" -ForegroundColor Blue
if (-not (Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange"})) {
    try {
        Connect-ExchangeOnline -ShowProgress $false -ErrorAction Stop | Out-Null
        Write-Host "Koneksi Berhasil." -ForegroundColor Green
    } catch {
        Write-Error "Gagal terhubung: $($_.Exception.Message)"
        Pause; exit 1
    }
} else {
    Write-Host "Sesi Aktif." -ForegroundColor Green
}

## -----------------------------------------------------------------------
## 4. FUNGSI MENU & DISPLAY
## -----------------------------------------------------------------------

function Get-ExUser {
    try {
        $conn = Get-ConnectionInformation | Select-Object -First 1
        return if ($null -ne $conn) { $conn.UserPrincipalName } else { "Not Connected" }
    } catch { return "Not Connected" }
}

function Show-Menu {
    Clear-Host
    Write-Host "=============================================" -ForegroundColor Cyan
    Write-Host "Author : Erik Dito Tampubolon - TelkomSigma" -ForegroundColor White
    Write-Host "Version: 2.3 (Path Integrity Fixed)" -ForegroundColor White
    Write-Host "=============================================" -ForegroundColor Cyan
    
    Write-Host "Location     : $scriptDir" -ForegroundColor Gray
    Write-Host "Time         : $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')" -ForegroundColor Gray  
    Write-Host "---------------------------------------------" -ForegroundColor Cyan
    Write-Host "Menu Pilihan:" -ForegroundColor Yellow
    Write-Host "  1. Assign or Remove License User by .csv"
    Write-Host "  2. Export License Availability"
    Write-Host "  3. Export User Last Password Changes by .csv"
    Write-Host "  4. Export User Last Logon by .csv"
    Write-Host "  5. Export User OneDrive Storage by .csv"
    Write-Host "  6. Export All Active User (UPN and Contact)"
    Write-Host "  7. Export Active User (UPN and Contact) by .csv"
    Write-Host "  8. Export All Mailbox"
    Write-Host "  9. Export All Active User"
    Write-Host ""
    Write-Host "  10. Keluar & Putus Koneksi" -ForegroundColor Red
    Write-Host "=============================================" -ForegroundColor Cyan
}

## -----------------------------------------------------------------------
## 5. LOGIKA EKSEKUSI LOOP
## -----------------------------------------------------------------------

$running = $true
while ($running) {
    Show-Menu
    $choice = Read-Host "Pilih nomor menu"

    # Menyusun Path Script Anak dengan validasi $scriptDir
    $scripts = @{
        "1" = Join-Path $scriptDir "assign-or-remove-license-user-by-csv-final.ps1"
        "2" = Join-Path $scriptDir "check-license-name-and-quota-final.ps1"
        "3" = Join-Path $scriptDir "check-lastpasswordchange-user-by-csv-final.ps1"
        "4" = Join-Path $scriptDir "check-lastlogon-user-by-csv-final.ps1"
        "5" = Join-Path $scriptDir "check-storage-user-by-csv-final.ps1"
        "6" = Join-Path $scriptDir "export-alluser-userprincipalname-contact-final.ps1"
        "7" = Join-Path $scriptDir "export-userprincipalname-contact-by-csv-final.ps1"
        "8" = Join-Path $scriptDir "export-mailbox-final.ps1"
        "9" = Join-Path $scriptDir "export-active-users-final.ps1"
    }

    if ($choice -eq "10") {
        Write-Host "`nClosing session..." -ForegroundColor Cyan
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        $running = $false
    }
    elseif ($scripts.ContainsKey($choice)) {
        $targetScript = $scripts[$choice]
        if (Test-Path $targetScript) {
            Write-Host "`n[Mengeksekusi: $(Split-Path $targetScript -Leaf)...]" -ForegroundColor Green
            & $targetScript
        } else {
            Write-Host "`nError: File tidak ditemukan di $targetScript" -ForegroundColor Red
        }
        Pause
    }
    else {
        Write-Host "`nPilihan tidak valid!" -ForegroundColor Red
        Start-Sleep -Seconds 1
    }
}