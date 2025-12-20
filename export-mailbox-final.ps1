<#
.SYNOPSIS
Mengekspor daftar semua Mailbox Exchange Online ke file CSV.

.DESCRIPTION
Skrip ini menggunakan Get-Mailbox untuk mengambil semua mailbox pengguna. Skrip ini akan MELEWATI koneksi jika sesi Exchange Online sudah aktif, dan yang PALING PENTING: TIDAK AKAN MEMUTUSKAN koneksi di akhir skrip (Bagian 4.2) sehingga sesi tetap aktif untuk skrip selanjutnya.

.AUTHOR
AI PowerShell Expert

.VERSION
3.3 (Export All Mailboxes to CSV - PERSISTENT SESSION)

.PREREQUISITES
Membutuhkan koneksi internet. Akun yang digunakan harus memiliki izin yang cukup untuk mengakses Exchange Online.

.NOTES
Output file akan disimpan di direktori yang sama dengan skrip ini, dengan nama Output_ExportAllMailboxesToCSV_[Timestamp].csv.
Koneksi ke Exchange Online AKAN TETAP AKTIF setelah skrip selesai.
#>
# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V3.3)
# Menyimpan output skrip ke file CSV dinamis di folder skrip.
# =========================================================================

# Variabel Global dan Output
$scriptName = "ExportAllMailboxesToCSV" # Nama skrip yang sebenarnya
$scriptOutput = @() # Array tempat semua data hasil skrip dikumpulkan

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
# Menggunakan $PSScriptRoot memastikan file disimpan di folder yang sama dengan skrip
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"

# Penanganan kasus $PSScriptRoot tidak ada saat dijalankan dari konsol
$scriptDir = if ($PSScriptRoot) {$PSScriptRoot} else {(Get-Location).Path}
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName


## -----------------------------------------------------------------------
## 3. LOGIKA UTAMA SCRIPT ANDA DI SINI
## -----------------------------------------------------------------------

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta

try {
    Write-Host "3.1. Mengambil semua Mailbox Pengguna..." -ForegroundColor Cyan
    
    # Logika inti: Mendapatkan mailbox pengguna dan memilih properti yang relevan.
    $mailboxData = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -ErrorAction Stop | 
                   Select-Object DisplayName, SamAccountName, PrimarySmtpAddress | 
                   Sort-Object PrimarySmtpAddress

    $scriptOutput = $mailboxData
    $totalMailboxes = $scriptOutput.Count
    Write-Host "  Ditemukan $($totalMailboxes) Mailbox untuk diekspor." -ForegroundColor Green
    
}
catch {
    $reason = "FATAL ERROR: Gagal mengambil Mailbox. Error: $($_.Exception.Message)"
    Write-Error $reason
    
    if ($scriptOutput.Count -eq 0) {
        $scriptOutput += [PSCustomObject]@{
            DisplayName = "ERROR"; SamAccountName = "N/A"; PrimarySmtpAddress = $reason
        }
    }
}

# >>> AKHIR DARI KODE UTAMA SKRIP ANDA <<<

## -----------------------------------------------------------------------
## 4. CLEANUP, DISCONNECT, DAN EKSPOR HASIL
## -----------------------------------------------------------------------

Write-Host "`n--- 4. Cleanup, Memutus Koneksi, dan Ekspor Hasil ---" -ForegroundColor Blue

# 4.1. Ekspor Hasil
if ($scriptOutput.Count -gt 0) {
    Write-Host "Mengekspor $($scriptOutput.Count) baris data hasil skrip..." -ForegroundColor Yellow
    try {
        # Menggunakan Export-Csv ke $outputFilePath
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Encoding UTF8 -Delimiter "," -ErrorAction Stop
        Write-Host " Data berhasil diekspor ke:" -ForegroundColor Green
        Write-Host " $outputFilePath" -ForegroundColor Green
    }
    catch {
        Write-Error "Gagal mengekspor data ke CSV: $($_.Exception.Message)"
    }
} else {
    Write-Host " Tidak ada data yang dikumpulkan (\$scriptOutput kosong). Melewati ekspor." -ForegroundColor DarkYellow
}

# 4.2. Memutus koneksi Exchange Online - BAGIAN INI DIKOMENTARI
#         UNTUK MEMPERTAHANKAN SESI TETAP AKTIF DI POWERSHELL.
if (Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange"}) {
#    Write-Host "Memutuskan koneksi dari Exchange Online..." -ForegroundColor DarkYellow
#    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
#    Write-Host " Koneksi Exchange Online diputus." -ForegroundColor Green
    Write-Host " Koneksi Exchange Online dipertahankan agar skrip berikutnya tidak perlu login ulang." -ForegroundColor DarkYellow
}


Write-Host "`nSkrip $($scriptName) selesai dieksekusi." -ForegroundColor Yellow