<#
.SYNOPSIS
Mengekspor semua Mail Flow (Transport) Rules dari Exchange Online ke file CSV.

.DESCRIPTION
Skrip ini menggunakan Get-TransportRule untuk mengambil semua aturan dan memproses properti kompleks (Conditions dan Actions) menjadi format string yang dapat dibaca di CSV. Bagian koneksi dan prasyarat dikomentari, **membutuhkan koneksi Exchange Online manual sebelum eksekusi**.

.AUTHOR
AI PowerShell Expert

.VERSION
2.0 (Export Transport Rules via Get-TransportRule - Manual Connection)

.PREREQUISITES
Sesi Exchange Online harus sudah aktif (menggunakan Connect-ExchangeOnline) sebelum menjalankan skrip ini.

.NOTES
Output file akan disimpan di direktori yang sama dengan skrip ini, dengan nama Output_ExportTransportRulesToCSV_[Timestamp].csv.
#>
# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.0)
# Menyimpan output skrip ke file CSV dinamis di folder skrip.
# =========================================================================

# Variabel Global dan Output
$scriptName = "ExportTransportRulesToCSV"
$scriptOutput = @()

# Tentukan jalur dan nama file output dinamis
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"

# Menentukan lokasi file output
$scriptDir = if ($PSScriptRoot) {$PSScriptRoot} else {(Get-Location).Path}
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName


## -----------------------------------------------------------------------
## 3. LOGIKA UTAMA SCRIPT ANDA DI SINI (EXPORT TRANSPORT RULES)
## -----------------------------------------------------------------------

Write-Host "`n--- 3. Memulai Logika Utama Skrip: $($scriptName) ---" -ForegroundColor Magenta
Write-Host "PENTING: Pastikan Anda telah menjalankan Connect-ExchangeOnline secara manual sebelum melanjutkan." -ForegroundColor Red

try {
    Write-Host "3.1. Mengambil semua Mail Flow Rules..." -ForegroundColor Cyan
    
    # Ambil semua Mail Flow Rules. Pilih properti yang relevan.
    $rules = Get-TransportRule | Select-Object Name, State, Priority, *Conditions, *Actions, Description, WhenCreated, WhenChanged, SenderRestrictions

    $totalRules = $rules.Count
    Write-Host "Ditemukan $($totalRules) Aturan Transport." -ForegroundColor Green
    
    $i = 0
    foreach ($rule in $rules) {
        $i++
        
        Write-Progress -Activity "Collecting Transport Rule Data" `
                       -Status "Processing ${i} of ${totalRules}: $($rule.Name)" `
                       -PercentComplete ([int](($i / $totalRules) * 100))
        
        # Inisialisasi variabel untuk properti yang kompleks
        $conditions = @()
        $actions = @()

        # Iterasi melalui semua properti untuk menemukan Conditions dan Actions
        $rule.PSObject.Properties | ForEach-Object {
            $propName = $_.Name
            $propValue = $_.Value
            
            if ($propValue -is [System.Collections.ICollection] -and $propValue.Count -gt 0) {
                # Properti adalah koleksi (misalnya: SentToMemberOf)
                $stringValue = $propValue -join "; "
            } elseif ($propValue -ne $null) {
                # Properti adalah nilai tunggal yang valid
                $stringValue = $propValue.ToString()
            } else {
                # Properti null atau kosong
                $stringValue = ""
            }

            if ($propName -like "*Conditions*") {
                # Hanya simpan Kondisi jika ada nilai
                if (-not [string]::IsNullOrEmpty($stringValue)) {
                    $conditions += "$propName : $stringValue"
                }
            } elseif ($propName -like "*Actions*") {
                 # Hanya simpan Actions jika ada nilai
                if (-not [string]::IsNullOrEmpty($stringValue)) {
                    $actions += "$propName : $stringValue"
                }
            }
        }
        
        # Gabungkan semua Conditions dan Actions menjadi satu string
        $conditionsString = $conditions -join "`r`n"
        $actionsString = $actions -join "`r`n"
        
        # Bangun objek kustom untuk diekspor
        $scriptOutput += [PSCustomObject]@{
            RuleName = $rule.Name
            State = $rule.State
            Priority = $rule.Priority
            WhenCreated = $rule.WhenCreated
            WhenChanged = $rule.WhenChanged
            SenderRestrictions = $rule.SenderRestrictions
            Description = $rule.Description
            
            # Kolom Conditions dan Actions
            AllConditions = $conditionsString
            AllActions = $actionsString
        }
    }

    Write-Progress -Activity "Collecting Transport Rule Data Complete" -Status "Exporting Results" -Completed

}
catch {
    $reason = "Gagal fatal saat mengambil Mail Flow Rules: $($_.Exception.Message)"
    Write-Error $reason
    $scriptOutput += [PSCustomObject]@{
        RuleName = "FATAL ERROR"; State = "N/A"; Priority = "N/A";
        AllConditions = $reason; AllActions = "N/A"
    }
}


## -----------------------------------------------------------------------
## 4. CLEANUP, DISCONNECT, DAN EKSPOR HASIL
## -----------------------------------------------------------------------

Write-Host "`n--- 4. Cleanup, Memutus Koneksi, dan Ekspor Hasil ---" -ForegroundColor Blue

# 4.1. Ekspor Hasil
if ($scriptOutput.Count -gt 0) {
    Write-Host "Mengekspor $($scriptOutput.Count) baris data hasil skrip..." -ForegroundColor Yellow
    try {
        # Menggunakan pemisah koma default (,)
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter "," -ErrorAction Stop
        Write-Host " Data berhasil diekspor ke:" -ForegroundColor Green
        Write-Host " $outputFilePath" -ForegroundColor Green
    }
    catch {
        Write-Error "Gagal mengekspor data ke CSV: $($_.Exception.Message)"
    }
} else {
    Write-Host " ⚠️ Tidak ada data yang dikumpulkan (\$scriptOutput kosong). Melewati ekspor." -ForegroundColor DarkYellow
}

# 4.2. Memutus koneksi Exchange Online
# Baris ini tetap tidak dikomentari agar Anda dapat membersihkan sesi yang Anda buat secara manual.
if (Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange"}) {
    Write-Host "Memutuskan koneksi dari Exchange Online..." -ForegroundColor DarkYellow
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host " Koneksi Exchange Online diputus." -ForegroundColor Green
}

Write-Host "`nSkrip $($scriptName) selesai dieksekusi." -ForegroundColor Yellow