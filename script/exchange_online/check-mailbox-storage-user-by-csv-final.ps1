# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V5.1 - No Header)
# =========================================================================

# Variabel Global dan Output
$scriptName = "MailboxStorageByCSVReport" 
$scriptOutput = @() 
$inputFileName = "UserPrincipalName.csv"

# Penanganan Jalur Aman
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$inputFilePath = Join-Path -Path $scriptDir -ChildPath $inputFileName

# Tentukan jalur output
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFileName = "Output_${scriptName}_${timestamp}.csv"
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

# ==========================================================
#                INFORMASI SCRIPT                
# ==========================================================
Write-Host "`n================================================" -ForegroundColor Yellow
Write-Host "                INFORMASI SCRIPT                " -ForegroundColor Yellow
Write-Host "================================================" -ForegroundColor Yellow
Write-Host " Nama Skrip        : MailboxStorageReport" -ForegroundColor Yellow
Write-Host " Field Kolom       : [UserPrincipalName]
                     [DisplayName]
                     [TotalItemSizeGB]
                     [ItemCount]
                     [WarningQuota]
                     [SendQuota]
                     [LastLogonTime]
                     [Status]
                     [Reason]" -ForegroundColor Yellow
Write-Host " Deskripsi Singkat : Script ini berfungsi untuk membuat laporan penggunaan storage mailbox berdasarkan daftar UPN dari file CSV tanpa header. Script akan memvalidasi mailbox, mengambil statistik ukuran, jumlah item, kuota peringatan dan kuota kirim, serta waktu logon terakhir, kemudian mengekspor hasil ke file CSV." -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Yellow

# ==========================================================
# KONFIRMASI EKSEKUSI
# ==========================================================
$confirmation = Read-Host "Apakah Anda ingin menjalankan skrip ini? (Y/N)"

if ($confirmation -ne "Y") {
    Write-Host "`nEksekusi skrip dibatalkan oleh pengguna." -ForegroundColor Red
    return
}

## -----------------------------------------------------------------------
## 3. LOGIKA UTAMA SCRIPT
## -----------------------------------------------------------------------

Write-Host "`n--- 3. Memulai Logika Utama Skrip: ${scriptName} ---" -ForegroundColor Magenta

if (-not (Test-Path -Path $inputFilePath)) {
    Write-Error "File input CSV tidak ditemukan di: $inputFilePath"
    $scriptOutput += [PSCustomObject]@{ UserPrincipalName = $inputFileName; Status = "FAIL"; Reason = "Input file not found." }
} else {
    Write-Host "Memuat data dari '${inputFileName}' (Mode: No Header)..." -ForegroundColor Cyan
    
    # MODIFIKASI: Menggunakan -Header "UserPrincipalName" karena CSV tidak memiliki judul kolom
    $users = Import-Csv -Path $inputFilePath -Header "UserPrincipalName" -ErrorAction SilentlyContinue
    
    $totalUsers = $users.Count
    $userCount = 0

    if ($totalUsers -eq 0) {
        Write-Host "File CSV kosong." -ForegroundColor Yellow
    }
    
    Write-Host "Total ${totalUsers} pengguna ditemukan." -ForegroundColor Yellow

    foreach ($entry in $users) {
        $userCount++
        
        # Trim email untuk membersihkan spasi
        $upn = if ($entry.UserPrincipalName) { $entry.UserPrincipalName.Trim() } else { $null }
        
        if ([string]::IsNullOrWhiteSpace($upn)) { continue }

        # Update Progress Bar dengan kurung kurawal ${}
        Write-Progress -Activity "Generating Storage Report" `
                       -Status "Processing User ${userCount} of ${totalUsers}: ${upn}" `
                       -PercentComplete ([int](($userCount / $totalUsers) * 100))
        
        Write-Host "-> [${userCount}/${totalUsers}] Memproses: ${upn}..." -ForegroundColor White
        
        try {
            # 3.2.1. Validasi Keberadaan Mailbox
            $recipient = Get-Recipient -Identity $upn -ErrorAction Stop | Select-Object RecipientType

            if ($recipient.RecipientType -like "*UserMailbox*") {
                # 3.2.2. Ambil Statistik dan Quota
                $stats = Get-MailboxStatistics -Identity $upn -ErrorAction Stop | Select-Object TotalItemSize, ItemCount, LastLogonTime
                $mailbox = Get-Mailbox -Identity $upn -ErrorAction Stop | Select-Object DisplayName, ProhibitSendQuota, ProhibitSendReceiveQuota, IssueWarningQuota
                
                # Konversi Size ke GB secara aman (Handle Deserialized Object)
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
                Write-Host "Size: ${TotalItemSizeGB} GB" -ForegroundColor DarkGreen

            } else {
                $reason = "Recipient type is $($recipient.RecipientType) (Not a UserMailbox)."
                Write-Host "Gagal: ${reason}" -ForegroundColor Yellow
                $scriptOutput += [PSCustomObject]@{
                    UserPrincipalName = $upn; Status = "FAIL"; Reason = $reason
                }
            }
        } 
        catch {
            $errMsg = $_.Exception.Message
            Write-Host "ERROR: ${errMsg}" -ForegroundColor Red
            $scriptOutput += [PSCustomObject]@{
                UserPrincipalName = $upn; Status = "FAIL"; Reason = $errMsg
            }
        }
    }
    Write-Progress -Activity "Storage Report" -Completed
}

## -----------------------------------------------------------------------
## 4. EKSPOR HASIL
## -----------------------------------------------------------------------
if ($scriptOutput.Count -gt 0) {
    # 1. Tentukan nama folder
    $exportFolderName = "exported_data"
    
    # 2. Ambil jalur dua tingkat di atas direktori skrip
    # Contoh: Jika skrip di C:\Users\Erik\Project\Scripts, maka ini ke C:\Users\Erik\
    $parentDir = (Get-Item $scriptDir).Parent.Parent.FullName
    
    # 3. Gabungkan menjadi jalur folder ekspor
    $exportFolderPath = Join-Path -Path $parentDir -ChildPath $exportFolderName

    # 4. Cek apakah folder 'exported_data' sudah ada di lokasi tersebut, jika belum buat baru
    if (-not (Test-Path -Path $exportFolderPath)) {
        New-Item -Path $exportFolderPath -ItemType Directory | Out-Null
        Write-Host "`nFolder '$exportFolderName' berhasil dibuat di: $parentDir" -ForegroundColor Yellow
    }

    # 5. Tentukan nama file dan jalur lengkap
    $outputFileName = "Output_$($scriptName)_$($timestamp).csv"
    $resultsFilePath = Join-Path -Path $exportFolderPath -ChildPath $outputFileName
    
    # 6. Ekspor data
    $scriptOutput | Export-Csv -Path $resultsFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
    
    Write-Host "`nSemua proses selesai!" -ForegroundColor Green
    Write-Host "Laporan tersimpan di: ${resultsFilePath}" -ForegroundColor Cyan
}