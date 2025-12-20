<#
.SYNOPSIS
Membuat laporan tentang tanggal Last Logon kotak surat berdasarkan daftar email dari file CSV TANPA HEADER.
.DESCRIPTION
Skrip menggunakan parameter -Header pada Import-Csv sehingga baris pertama file CSV langsung dianggap sebagai data.
#>

# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.1 - No Header)
# =========================================================================

# Variabel Global dan Output
$scriptName = "MailboxLastLogonReport" 
$scriptOutput = @() 
$inputFileName = "UserPrincipalName.csv"

# Penanganan Jalur Aman (Fix: Empty Path Error)
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$inputFilePath = Join-Path -Path $scriptDir -ChildPath $inputFileName

# Tentukan jalur output
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFileName = "Output_${scriptName}_${timestamp}.csv"
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

## -----------------------------------------------------------------------
## 3. LOGIKA UTAMA SCRIPT
## -----------------------------------------------------------------------

Write-Host "`n--- 3. Memulai Logika Utama Skrip: ${scriptName} ---" -ForegroundColor Magenta

if (-not (Test-Path -Path $inputFilePath)) {
    Write-Error "File input CSV tidak ditemukan di: $inputFilePath"
} else {
    Write-Host "Memuat data dari '${inputFileName}' (Mode: No Header)..." -ForegroundColor Cyan
    
    # MODIFIKASI: Menggunakan -Header "UserPrincipalName" karena CSV tidak memiliki judul kolom
    $users = Import-Csv -Path $inputFilePath -Header "UserPrincipalName" -ErrorAction SilentlyContinue
    
    $totalUsers = $users.Count
    $userCount = 0

    if ($totalUsers -eq 0) {
        Write-Host "⚠️ File CSV kosong." -ForegroundColor Yellow
    }
    
    Write-Host "Total ${totalUsers} pengguna ditemukan." -ForegroundColor Yellow

    foreach ($entry in $users) {
        $userCount++
        
        # Trim email untuk membersihkan spasi yang mungkin ada
        $upn = if ($entry.UserPrincipalName) { $entry.UserPrincipalName.Trim() } else { $null }
        
        # Skip jika baris kosong
        if ([string]::IsNullOrWhiteSpace($upn)) { continue }

        # FIX: Menggunakan ${} untuk menghindari error 'Variable reference is not valid'
        Write-Progress -Activity "Generating Last Logon Report" `
                       -Status "Processing User ${userCount} of ${totalUsers}: ${upn}" `
                       -PercentComplete ([int](($userCount / $totalUsers) * 100))
        
        Write-Host "-> [${userCount}/${totalUsers}] Memproses: ${upn}..." -ForegroundColor White
        
        try {
            # 3.2.1. Validasi Keberadaan Mailbox
            $recipient = Get-Recipient -Identity $upn -ErrorAction Stop | Select-Object RecipientType, DisplayName

            if ($recipient.RecipientType -like "*UserMailbox*") {
                
                # 3.2.2. Ambil Statistik Mailbox (Output ditangkap agar tidak tumpah ke konsol)
                $stats = Get-MailboxStatistics -Identity $upn -ErrorAction Stop | Select-Object LastLogonTime

                $lastLogon = if ($stats.LastLogonTime) { 
                    $stats.LastLogonTime.ToString("yyyy-MM-dd HH:mm:ss") 
                } else { 
                    "N/A (Never Logged On)" 
                }
                
                $scriptOutput += [PSCustomObject]@{
                    UserPrincipalName = $upn
                    DisplayName       = $recipient.DisplayName
                    LastLogonTime     = $lastLogon
                    Status            = "SUCCESS"
                    Reason            = "Last Logon Time retrieved."
                }
                Write-Host "   ✅ Last Logon: ${lastLogon}" -ForegroundColor DarkGreen

            } else {
                $reason = "Recipient type is $($recipient.RecipientType) (Not a UserMailbox)."
                Write-Host "   ⚠️ Gagal: ${reason}" -ForegroundColor Yellow
                $scriptOutput += [PSCustomObject]@{
                    UserPrincipalName = $upn; DisplayName = $recipient.DisplayName; LastLogonTime = "N/A"; Status = "FAIL"; Reason = $reason
                }
            }
        } 
        catch {
            $errMsg = $_.Exception.Message
            $reason = if ($errMsg -like "*cannot be found*") { "Mailbox not found." } else { "Error: $errMsg" }
            
            Write-Host "   ❌ ERROR: ${reason}" -ForegroundColor Red
            $scriptOutput += [PSCustomObject]@{
                UserPrincipalName = $upn; DisplayName = ""; LastLogonTime = "N/A"; Status = "FAIL"; Reason = $reason
            }
        }
    }
    Write-Progress -Activity "Last Logon Report" -Completed
}

## -----------------------------------------------------------------------
## 4. EKSPOR HASIL
## -----------------------------------------------------------------------

if ($scriptOutput.Count -gt 0) {
    Write-Host "`nMengekspor hasil..." -ForegroundColor Yellow
    try {
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
        Write-Host " ✅ Data berhasil diekspor ke: $outputFilePath" -ForegroundColor Green
    }
    catch {
        Write-Error "Gagal ekspor: $($_.Exception.Message)"
    }
}

Write-Host "`nSkrip ${scriptName} selesai." -ForegroundColor Yellow