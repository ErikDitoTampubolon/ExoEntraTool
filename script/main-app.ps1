# =========================================================================
# AUTHOR   : Erik Dito Tampubolon - TelkomSigma
# VERSION  : 1.1
# DESKRIPSI: Skrip Utama dengan ASCII Art Header "ExoEntraTool"
# =========================================================================

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
if ([string]::IsNullOrWhiteSpace($scriptDir)) { $scriptDir = "." }

# --- 0. FUNGSI PEMERIKSAAN INTEGRITAS FILE ---
function Test-ScriptIntegrity {
    Write-Host "--- 0. Memeriksa Integritas File Aplikasi ---" -ForegroundColor Blue
    
    # Daftar semua file yang harus ada
    $requiredScripts = @(
        "script\exchange_online\assign-or-remove-license-user-by-csv-final.ps1",
        "script\exchange_online\check-license-name-and-quota-final.ps1",
        "script\exchange_online\check-mailbox-storage-user-by-csv-final.ps1",
        "script\exchange_online\export-active-users-final.ps1",
        "script\exchange_online\export-alluser-userprincipalname-contact-final.ps1",
        "script\exchange_online\check-lastpasswordchange-user-by-csv-final.ps1",
        "script\exchange_online\export-alluser-userprincipalname-contact-by-csv-final.ps1",
        "script\exchange_online\check-mailbox-storage-user-by-csv-final.ps1",
        "script\exchange_online\check-lastlogon-user-by-csv-final.ps1",
        "script\exchange_online\check-transport-rules-final.ps1",
        "script\entra\enable-or-disable-mfa-user-by-csv-final.ps1",
        "script\entra\force-password-change-alluser-by-csv-final.ps1",
        "script\entra\export-alluser-mfa-final.ps1",
        "script\entra\export-alldevice-final.ps1",
        "script\entra\export-alluser-owned-device-final.ps1",
        "script\entra\export-allapplication-final.ps1",
        "script\entra\export-alldeleted-user-final.ps1",
        "script\entra\export-alluser-inactive-30days-final.ps1",
        "script\entra\export-list-alldevice-final.ps1",
        "script\entra\check-conditional-access-policy-final.ps1",
        "script\entra\check-user-auth-method-final.ps1",
        "script\entra\check-permission-grant-policy-final.ps1",
        "script\entra\check-entra-auth-policy-final.ps1",
        "script\entra\check-entra-identity-provider-final.ps1"
    )

    $missingFiles = @()

    foreach ($relPath in $requiredScripts) {
        $fullPath = Join-Path $scriptDir $relPath
        if (-not (Test-Path $fullPath)) {
            $missingFiles += $relPath
        }
    }

    if ($missingFiles.Count -gt 0) {
        Write-Host "PERINGATAN: Terdapat $($missingFiles.Count) file skrip yang hilang!" -ForegroundColor Red
        foreach ($file in $missingFiles) {
            Write-Host " [!] Hilang: $file" -ForegroundColor Yellow
        }
        Write-Host "`nSilakan hubungi Team Modern Work - Telkomsigma untuk perbaikan file." -ForegroundColor Cyan
        
        $Host.UI.RawUI.FlushInputBuffer()
        $choice = Read-Host "Apakah Anda ingin tetap menjalankan aplikasi dengan fitur terbatas? (Y/N)"
        if ($choice.ToUpper() -ne 'Y') {
            Write-Host "Aplikasi dihentikan. Sampai jumpa!" -ForegroundColor Red
            exit
        }
    } else {
        Write-Host "Integritas file OK. Semua modul skrip ditemukan.`n" -ForegroundColor Green
    }
}

# Jalankan Health Check sebelum mulai
Clear-Host
Test-ScriptIntegrity

# --- 1. Memeriksa Lingkungan PowerShell ---
Set-ExecutionPolicy RemoteSigned -Scope Process -Force -ErrorAction SilentlyContinue

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

Write-Host "--- 1. Prasyarat Modul ---" -ForegroundColor Blue
Check-Module -ModuleName "PowerShellGet"
Check-Module -ModuleName "ExchangeOnlineManagement"
Check-Module -ModuleName "Microsoft.Graph"
Check-Module -ModuleName "Microsoft.Entra"
Check-Module -ModuleName "Microsoft.Entra.Beta"

# --- 2. Membangun Koneksi Multi-Service ---
$requiredScopes = "User.ReadWrite.All", "Organization.Read.All"
Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Graph ---" -ForegroundColor Blue

if (Get-MgContext -ErrorAction SilentlyContinue) {
    Write-Host "Sesi Graph yang ada akan diputus untuk koneksi ulang dengan scopes baru." -ForegroundColor DarkYellow
    Disconnect-MgGraph
}

Write-Host "Anda akan diminta untuk login. Pastikan Anda menyetujui scopes berikut:" -ForegroundColor Cyan
Write-Host $requiredScopes -ForegroundColor Yellow

try {
    Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop | Out-Null
    Write-Host "Koneksi ke Microsoft Graph berhasil!" -ForegroundColor Green
} catch {
    Write-Error "Gagal terhubung ke Microsoft Graph."
    exit 1
}

# 2.2 KONEKSI ENTRA
Write-Host "Menghubungkan ke Microsoft Entra..." -ForegroundColor Yellow
try {
    Connect-Entra -scope 'User.Read.All', 'UserAuthenticationMethod.Read.All' -ErrorAction Stop
    Write-Host "Koneksi Entra Berhasil." -ForegroundColor Green
} catch {
    Write-Host "Peringatan: Gagal terkoneksi ke Entra." -ForegroundColor Yellow
}

# 2.3 KONEKSI EXCHANGE ONLINE
if (-not (Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange"})) {
    Write-Host "Menghubungkan ke Exchange Online..." -ForegroundColor Yellow
    Connect-ExchangeOnline -ShowProgress $false -ErrorAction SilentlyContinue
}

Write-Host "`nSemua layanan terhubung. Memuat antarmuka..." -ForegroundColor Green
Start-Sleep -Seconds 1

## -----------------------------------------------------------------------
## FUNGSI HEADER DENGAN ASCII ART
## -----------------------------------------------------------------------
function Show-Header {
    Clear-Host
    $ascii = @'
  _____              ______       _               _______           _ 
 |  ___|            |  ____|     | |             |__   __|         | | 
 | |__  __  _____   | |__   _ __ | |_ _ __ __ _     | | ___   ___  | |
 |  __| \ \/ / _ \  |  __| | '_ \| __| '__/ _` |    | |/ _ \ / _ \ | |
 | |___  >  < (_) | | |____| | | | |_| | | (_| |    | | (_) | (_) || |
 \____/ /_/\_\___/  |______|_| |_|\__|_|  \__,_|    |_|\___/ \___/ |_|
'@

    Write-Host $ascii -ForegroundColor Cyan
    Write-Host "======================================================================" -ForegroundColor DarkCyan
    Write-Host " Author   : Erik Dito Tampubolon - TelkomSigma" -ForegroundColor White
    Write-Host " Version  : 1.1 (ExoEntraTool Suite)" -ForegroundColor White
    Write-Host "----------------------------------------------------------------------" -ForegroundColor DarkCyan
    Write-Host " Location : ${scriptDir}" -ForegroundColor Gray
    Write-Host " Time     : $(Get-Date -Format 'dd-MM-yyyy HH:mm:ss')" -ForegroundColor Gray  
    Write-Host "======================================================================" -ForegroundColor DarkCyan
    Write-Host ""
}

## -----------------------------------------------------------------------
## LOGIKA EKSEKUSI LOOP
## -----------------------------------------------------------------------
$mainRunning = $true
while ($mainRunning) {
    Show-Header
    Write-Host "Menu Utama:" -ForegroundColor Yellow
    Write-Host "  1. Microsoft Exchange Online"
    Write-Host "  2. Microsoft Entra (Entra ID)"
    Write-Host ""
    Write-Host "  10. Keluar & Putus Koneksi" -ForegroundColor Red
    Write-Host "======================================================================" -ForegroundColor DarkCyan
    
    $mainChoice = Read-Host "Pilih nomor menu"

    switch ($mainChoice) {
        "1" { 
            $subRunning = $true
            while ($subRunning) {
                Show-Header
                Write-Host "Sub-Menu: Microsoft Exchange Online" -ForegroundColor Yellow
                Write-Host "  1. Assign or Remove License User by .csv"
                Write-Host "  2. Export List License Availability"
                Write-Host "  3. Export List All Mailbox"
                Write-Host "  4. Export List All Active User"
                Write-Host "  5. Export List All Active User (DisplayName,UPN,Contact)"
                Write-Host "  6. Export List Active User Last Password Changes by .csv"
                Write-Host "  7. Export List Active User (DisplayName,UPN,Contact) by .csv"
                Write-Host "  8. Export List Mailbox Storage Usage by .csv"
                Write-Host "  9. Export List Active User Last Logon by .csv"
                Write-Host "  10. Export List Transport Rules"
                Write-Host ""
                Write-Host "  B. Kembali ke Menu Utama" -ForegroundColor Yellow
                Write-Host "---------------------------------------------" -ForegroundColor DarkCyan
                
                $subChoice = Read-Host "Pilih nomor menu"
                if ($subChoice.ToUpper() -eq "B") { $subRunning = $false }
                elseif ($subChoice -eq "1") { & (Join-Path $scriptDir "script\exchange_online\assign-or-remove-license-user-by-csv-final.ps1"); Pause }
                elseif ($subChoice -eq "2") { & (Join-Path $scriptDir "script\exchange_online\check-license-name-and-quota-final.ps1"); Pause }
                elseif ($subChoice -eq "3") { & (Join-Path $scriptDir "script\exchange_online\export-mailbox-final.ps1"); Pause }
                elseif ($subChoice -eq "4") { & (Join-Path $scriptDir "script\exchange_online\export-active-users-final.ps1"); Pause }
                elseif ($subChoice -eq "5") { & (Join-Path $scriptDir "script\exchange_online\export-alluser-userprincipalname-contact-final.ps1"); Pause }
                elseif ($subChoice -eq "6") { & (Join-Path $scriptDir "script\exchange_online\check-lastpasswordchange-user-by-csv-final.ps1"); Pause }
                elseif ($subChoice -eq "7") { & (Join-Path $scriptDir "script\exchange_online\export-alluser-userprincipalname-contact-by-csv-final.ps1"); Pause }
                elseif ($subChoice -eq "8") { & (Join-Path $scriptDir "script\exchange_online\check-mailbox-storage-user-by-csv-final.ps1"); Pause }
                elseif ($subChoice -eq "9") { & (Join-Path $scriptDir "script\exchange_online\check-lastlogon-user-by-csv-final.ps1"); Pause }
                elseif ($subChoice -eq "10") { & (Join-Path $scriptDir "script\exchange_online\check-transport-rules-final.ps1"); Pause }
            }
        }
        "2" { 
            $subRunning = $true
            while ($subRunning) {
                Show-Header
                Write-Host "Sub-Menu: Microsoft Entra (Entra ID)" -ForegroundColor Yellow
                Write-Host "  1. Enable or Disable MFA User by .csv"
                Write-Host "  2. Force Change Password User by .csv"
                Write-Host "  3. Export List All User MFA Status"
                Write-Host "  4. Export List All Device"
                Write-Host "  5. Export List All User Owned Device"
                Write-Host "  6. Export List All Application"
                Write-Host "  7. Export List All Deleted User"
                Write-Host "  8. Export List All Inactive User (30 days)"
                Write-Host "  9. Export List Duplicate Device"
                Write-Host "  10. Export List Conditional Access Policy"
                Write-Host "  11. Export List User Auth Method"
                Write-Host "  12. Export List Permission Grant Policy"
                Write-Host "  13. Export List Entra Auth Policy"
                Write-Host "  14. Export List Entra Identity Provider"
                Write-Host ""
                Write-Host "  B. Kembali ke Menu Utama" -ForegroundColor Yellow
                Write-Host "---------------------------------------------" -ForegroundColor DarkCyan
                
                $subChoice = Read-Host "Pilih nomor menu"
                if ($subChoice.ToUpper() -eq "B") { $subRunning = $false }
	elseif ($subChoice -eq "1") { & (Join-Path $scriptDir "script\entra\enable-or-disable-mfa-user-by-csv-final.ps1"); Pause }
	elseif ($subChoice -eq "2") { & (Join-Path $scriptDir "script\entra\force-password-change-alluser-by-csv-final.ps1"); Pause }
                elseif ($subChoice -eq "3") { & (Join-Path $scriptDir "script\entra\export-alluser-mfa-final.ps1"); Pause }
                elseif ($subChoice -eq "4") { & (Join-Path $scriptDir "script\entra\export-alldevice-final.ps1"); Pause }
                elseif ($subChoice -eq "5") { & (Join-Path $scriptDir "script\entra\export-alluser-owned-device-final.ps1"); Pause }
                elseif ($subChoice -eq "6") { & (Join-Path $scriptDir "script\entra\export-allapplication-final.ps1"); Pause }
                elseif ($subChoice -eq "7") { & (Join-Path $scriptDir "script\entra\export-alldeleted-user-final.ps1"); Pause }
                elseif ($subChoice -eq "8") { & (Join-Path $scriptDir "script\entra\export-alluser-inactive-30days-final.ps1"); Pause }
                elseif ($subChoice -eq "9") { & (Join-Path $scriptDir "script\entra\export-list-alldevice-final.ps1"); Pause }
                elseif ($subChoice -eq "10") { & (Join-Path $scriptDir "script\entra\check-conditional-access-policy-final.ps1"); Pause }
                elseif ($subChoice -eq "11") { & (Join-Path $scriptDir "script\entra\check-user-auth-method-final.ps1"); Pause }
                elseif ($subChoice -eq "12") { & (Join-Path $scriptDir "script\entra\check-permission-grant-policy-final.ps1"); Pause }
                elseif ($subChoice -eq "13") { & (Join-Path $scriptDir "script\entra\check-entra-auth-policy-final.ps1"); Pause }
                elseif ($subChoice -eq "14") { & (Join-Path $scriptDir "script\entra\check-entra-identity-provider-final.ps1"); Pause }
            }
        }
        "10" {
            Write-Host "`nClosing sessions..." -ForegroundColor Cyan
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            Disconnect-Entra -ErrorAction SilentlyContinue
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            $mainRunning = $false
        }
        default { 
            Write-Host "Pilihan tidak valid!" -ForegroundColor Red
            Start-Sleep -Seconds 1 
        }
    }
}