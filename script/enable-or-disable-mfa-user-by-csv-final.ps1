# =========================================================================
# FRAMEWORK SCRIPT POWERSHELL DENGAN EKSPOR OTOMATIS (V2.4)
# Nama Skrip: Bulk-MFA-Manager-NoHeader
# Deskripsi: Mengelola MFA via CSV tanpa header menggunakan TempColumn.
# =========================================================================

# 1. Konfigurasi File Input & Path
$inputFileName = "UserPrincipalName.csv"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$inputFilePath = Join-Path -Path $scriptDir -ChildPath $inputFileName

# Variabel Global dan Output
$scriptName = "BulkMFAManager_NoHeader"
$scriptOutput = [System.Collections.ArrayList]::new() 
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFileName = "Output_$($scriptName)_$($timestamp).csv"
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

## -----------------------------------------------------------------------
## 2. PEMILIHAN OPERASI & KONEKSI
## -----------------------------------------------------------------------

Write-Host "`n--- 2. Pemilihan Operasi MFA ---" -ForegroundColor Blue
$operationChoice = Read-Host "Pilih operasi: (1) Enable MFA | (2) Disable MFA"

switch ($operationChoice) {
    "1" { 
        $targetMfaState = "enabled"
        $operationType = "ENABLE-MFA" 
    }
    "2" { 
        $targetMfaState = "disabled"
        $operationType = "DISABLE-MFA" 
    }
    default { 
        Write-Host "Pilihan tidak valid." -ForegroundColor Red
        return 
    }
}

try {
    Write-Host "`nMenghubungkan ke Microsoft Entra..." -ForegroundColor Yellow
    Connect-Entra -Scopes 'Policy.ReadWrite.AuthenticationMethod', 'User.ReadWrite.All' -ErrorAction Stop
    Write-Host "Koneksi Berhasil." -ForegroundColor Green
} catch {
    Write-Error "Gagal terhubung: $($_.Exception.Message)"
    exit 1
}

## -----------------------------------------------------------------------
## 3. LOGIKA UTAMA SCRIPT
## -----------------------------------------------------------------------

Write-Host "`n--- 3. Memproses Operasi: $operationType ---" -ForegroundColor Magenta

if (Test-Path $inputFilePath) {
    # MENGGUNAKAN -Header "TempColumn" karena file tidak memiliki header asli
    # Filter baris kosong untuk menghindari error
    $users = Import-Csv $inputFilePath -Header "TempColumn" | Where-Object { $_.TempColumn -ne $null -and $_.TempColumn.Trim() -ne "" }
    $totalUsers = $users.Count
    $counter = 0

    if ($totalUsers -eq 0) {
        Write-Host "File '$inputFileName' kosong." -ForegroundColor Red
        exit
    }

    foreach ($user in $users) {
        $counter++
        
        # Mengambil nilai dari properti TempColumn
        $targetUser = $user.TempColumn.Trim()
        
        # Output progres baris tunggal sesuai permintaan
        $statusText = "-> [$counter/$totalUsers] Memproses: $targetUser . . ."
        Write-Host "`r$statusText" -ForegroundColor Green -NoNewline

        # Validasi format UPN/Email sederhana
        if ($targetUser -match "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$") {
            try {
                # Update MFA menggunakan Microsoft Entra Beta
                Update-EntraBetaUserAuthenticationRequirement -UserId $targetUser -PerUserMfaState $targetMfaState -ErrorAction Stop
                
                $res = [PSCustomObject]@{
                    Timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    UserAccount = $targetUser
                    Operation   = $operationType
                    Status      = "SUCCESS"
                    Message     = "Status set to $targetMfaState"
                }
            } catch {
                $res = [PSCustomObject]@{
                    Timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    UserAccount = $targetUser
                    Operation   = $operationType
                    Status      = "FAILED"
                    Message     = $_.Exception.Message
                }
            }
        } else {
            $res = [PSCustomObject]@{
                Timestamp   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                UserAccount = $targetUser
                Operation   = $operationType
                Status      = "FAILED"
                Message     = "Format Email/UPN tidak valid atau baris kosong."
            }
        }
        [void]$scriptOutput.Add($res)
    }
    Write-Host "`n`nPemrosesan Selesai." -ForegroundColor Green
} else {
    Write-Host "ERROR: File '$inputFileName' tidak ditemukan di $scriptDir" -ForegroundColor Red
}

## -----------------------------------------------------------------------
## 4. CLEANUP DAN EKSPOR HASIL
## -----------------------------------------------------------------------

Write-Host "`n--- 4. Cleanup dan Ekspor Hasil ---" -ForegroundColor Blue

if ($scriptOutput.Count -gt 0) {
    $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
    Write-Host "Laporan detail eksekusi: $outputFilePath" -ForegroundColor Green
}

Disconnect-Entra -ErrorAction SilentlyContinue
Write-Host "`nSkrip selesai dieksekusi." -ForegroundColor Yellow