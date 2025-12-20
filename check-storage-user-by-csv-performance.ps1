<#
.SYNOPSIS
Membuat laporan detail penggunaan storage mailbox berdasarkan daftar email dari file CSV TANPA HEADER.
.DESCRIPTION
Skrip dioptimalkan untuk kecepatan dengan paralelisme (PS7+), batching, dan cmdlet EXO efisien.
#>

# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V5.1 - No Header - Optimized)
# =========================================================================

# Variabel Global dan Output
$scriptName = "MailboxStorageReport" 
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
## 3. LOGIKA UTAMA SCRIPT
## -----------------------------------------------------------------------

Write-Host "`n--- 3. Memulai Logika Utama Skrip: ${scriptName} ---" -ForegroundColor Magenta

# Saran: Jalankan Connect-ExchangeOnline -UseMultithreading sebelum script ini untuk mengaktifkan multithreading.

if (-not (Test-Path -Path $inputFilePath)) {
    Write-Error "File input CSV tidak ditemukan di: $inputFilePath"
    $scriptOutput += [PSCustomObject]@{ UserPrincipalName = $inputFileName; Status = "FAIL"; Reason = "Input file not found." }
} else {
    Write-Host "Memuat data dari '${inputFileName}' (Mode: No Header)..." -ForegroundColor Cyan
    
    # MODIFIKASI: Menggunakan -Header "UserPrincipalName" karena CSV tidak memiliki judul kolom
    $users = Import-Csv -Path $inputFilePath -Header "UserPrincipalName" -ErrorAction SilentlyContinue
    
    $totalUsers = $users.Count
    if ($totalUsers -eq 0) {
        Write-Host "⚠️ File CSV kosong." -ForegroundColor Yellow
    }
    
    Write-Host "Total ${totalUsers} pengguna ditemukan." -ForegroundColor Yellow

    # Kumpul UPN valid (trim dan filter kosong)
    $upns = $users | ForEach-Object { $_.UserPrincipalName.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    # Cek versi PowerShell untuk paralelisme
    $psVersion = $PSVersionTable.PSVersion.Major
    if ($psVersion -ge 7) {
        # Paralel di PS7+ dengan throttle limit 5 untuk hindari throttling
        $scriptOutput = $upns | ForEach-Object -Parallel {
            $upn = $_
            try {
                # Gunakan cmdlet EXO untuk kecepatan
                $recipient = Get-EXORecipient -Identity $upn -Properties RecipientType -ErrorAction Stop

                if ($recipient.RecipientType -like "*UserMailbox*") {
                    $stats = Get-EXOMailboxStatistics -Identity $upn -Properties TotalItemSize, ItemCount, LastLogonTime -ErrorAction Stop
                    $mailbox = Get-EXOMailbox -Identity $upn -Properties DisplayName, ProhibitSendQuota, ProhibitSendReceiveQuota, IssueWarningQuota -ErrorAction Stop
                    
                    # Konversi Size ke GB secara aman
                    $TotalItemSizeGB = try {
                        [math]::Round(($stats.TotalItemSize.Value.ToBytes() / 1GB), 2)
                    } catch {
                        $rawSize = $stats.TotalItemSize.ToString()
                        if ($rawSize -match '(\d+\.?\d*)\s*(\w+)') {
                            $val = [double]$Matches[1]
                            $unit = $Matches[2].ToUpper()
                            switch ($unit) {
                                "KB" { [math]::Round($val / 1MB, 4) }
                                "MB" { [math]::Round($val / 1024, 2) }
                                "GB" { $val }
                                default { $rawSize }
                            }
                        } else { $rawSize }
                    }
                    
                    [PSCustomObject]@{
                        UserPrincipalName = $upn
                        DisplayName       = $mailbox.DisplayName
                        TotalItemSizeGB   = $TotalItemSizeGB
                        ItemCount         = $stats.ItemCount
                        WarningQuota      = $mailbox.IssueWarningQuota.ToString()
                        SendQuota         = $mailbox.ProhibitSendQuota.ToString()
                        LastLogonTime     = $stats.LastLogonTime
                        Status            = "SUCCESS"
                        Reason            = "Storage data collected."
                    }
                } else {
                    [PSCustomObject]@{
                        UserPrincipalName = $upn; Status = "FAIL"; Reason = "Recipient type is $($recipient.RecipientType) (Not a UserMailbox)."
                    }
                }
            } catch {
                [PSCustomObject]@{
                    UserPrincipalName = $upn; Status = "FAIL"; Reason = $_.Exception.Message
                }
            }
        } -ThrottleLimit 5
    } else {
        # Fallback ke loop serial untuk PS5.1, dengan batching sederhana
        $userCount = 0
        foreach ($upn in $upns) {
            $userCount++
            Write-Progress -Activity "Generating Storage Report" `
                           -Status "Processing User ${userCount} of ${totalUsers}: ${upn}" `
                           -PercentComplete ([int](($userCount / $totalUsers) * 100))
            
            try {
                $recipient = Get-Recipient -Identity $upn -ErrorAction Stop | Select-Object RecipientType

                if ($recipient.RecipientType -like "*UserMailbox*") {
                    $stats = Get-MailboxStatistics -Identity $upn -ErrorAction Stop | Select-Object TotalItemSize, ItemCount, LastLogonTime
                    $mailbox = Get-Mailbox -Identity $upn -ErrorAction Stop | Select-Object DisplayName, ProhibitSendQuota, ProhibitSendReceiveQuota, IssueWarningQuota
                    
                    $TotalItemSizeGB = try {
                        [math]::Round(($stats.TotalItemSize.Value.ToBytes() / 1GB), 2)
                    } catch {
                        $rawSize = $stats.TotalItemSize.ToString()
                        if ($rawSize -match '(\d+\.?\d*)\s*(\w+)') {
                            $val = [double]$Matches[1]
                            $unit = $Matches[2].ToUpper()
                            switch ($unit) {
                                "KB" { [math]::Round($val / 1MB, 4) }
                                "MB" { [math]::Round($val / 1024, 2) }
                                "GB" { $val }
                                default { $rawSize }
                            }
                        } else { $rawSize }
                    }
                    
                    $scriptOutput += [PSCustomObject]@{
                        UserPrincipalName = $upn
                        DisplayName       = $mailbox.DisplayName
                        TotalItemSizeGB   = $TotalItemSizeGB
                        ItemCount         = $stats.ItemCount
                        WarningQuota      = $mailbox.IssueWarningQuota.ToString()
                        SendQuota         = $mailbox.ProhibitSendQuota.ToString()
                        LastLogonTime     = $stats.LastLogonTime
                        Status            = "SUCCESS"
                        Reason            = "Storage data collected."
                    }
                } else {
                    $scriptOutput += [PSCustomObject]@{
                        UserPrincipalName = $upn; Status = "FAIL"; Reason = "Recipient type is $($recipient.RecipientType) (Not a UserMailbox)."
                    }
                }
            } catch {
                $scriptOutput += [PSCustomObject]@{
                    UserPrincipalName = $upn; Status = "FAIL"; Reason = $_.Exception.Message
                }
            }
        }
        Write-Progress -Activity "Storage Report" -Completed
    }
}

## -----------------------------------------------------------------------
## 4. EKSPOR HASIL
## -----------------------------------------------------------------------

if ($scriptOutput.Count -gt 0) {
    Write-Host "`nMengekspor hasil..." -ForegroundColor Yellow
    try {
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
        Write-Host " ✅ Laporan berhasil diekspor ke: $outputFilePath" -ForegroundColor Green
    }
    catch {
        Write-Error "Gagal ekspor: $($_.Exception.Message)"
    }
}

Write-Host "`nSkrip ${scriptName} selesai dieksekusi." -ForegroundColor Yellow