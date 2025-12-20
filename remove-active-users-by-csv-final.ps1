# =========================================================================
# REMOVE ACTIVE USERS - MICROSOFT GRAPH SDK (V1.0)
# Deskripsi: Menghapus user secara massal berdasarkan list UPN di CSV.
# =========================================================================

# 1. Variabel Global dan Output
$scriptName = "RemoveActiveUsers"
$scriptOutput = @() 
$inputFileName = "UserPrincipalName.csv" # File CSV tanpa header berisi list UPN

# Penanganan Jalur Folder
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$inputFilePath = Join-Path -Path $scriptDir -ChildPath $inputFileName

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFileName = "Output_${scriptName}_${timestamp}.csv"
$outputFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName

## -----------------------------------------------------------------------
## 2. KONEKSI (SILENT MODE)
## -----------------------------------------------------------------------
Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Graph ---" -ForegroundColor Blue

if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
    try {
        # Scope User.ReadWrite.All diperlukan untuk proses penghapusan
        Connect-MgGraph -Scopes "User.ReadWrite.All" -ErrorAction Stop | Out-Null
        Write-Host "✅ Koneksi Berhasil." -ForegroundColor Green
    } catch {
        Write-Error "Gagal terhubung ke Microsoft Graph."
        exit 1
    }
} else {
    Write-Host "✅ Sesi Aktif." -ForegroundColor Green
}

## -----------------------------------------------------------------------
## 2. KONEKSI KE MICROSOFT GRAPH (SILENT MODE)
## -----------------------------------------------------------------------
$requiredScopes = "User.ReadWrite.All", "Organization.Read.All"
Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Graph ---" -ForegroundColor Blue

if (Get-MgContext -ErrorAction SilentlyContinue) {
    Write-Host "✅ Sesi Microsoft Graph aktif." -ForegroundColor Green
} else {
    Write-Host "Menghubungkan ke Microsoft Graph..." -ForegroundColor Cyan
    try {
        Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop | Out-Null
        Write-Host "✅ Koneksi Berhasil." -ForegroundColor Green
    } catch {
        Write-Error "Gagal terhubung ke Microsoft Graph."
        return
    }
}


## -----------------------------------------------------------------------
## 3. LOGIKA UTAMA (DELETE PROCESS)
## -----------------------------------------------------------------------
Write-Host "`n--- 3. Memulai Logika Utama Skrip: ${scriptName} ---" -ForegroundColor Magenta

if (-not (Test-Path -Path $inputFilePath)) {
    Write-Host "❌ ERROR: File ${inputFileName} tidak ditemukan di folder skrip!" -ForegroundColor Red
} else {
    # Membaca CSV TANPA HEADER
    $users = Import-Csv -Path $inputFilePath -Header "UserPrincipalName" -ErrorAction SilentlyContinue
    $totalUsers = $users.Count
    $userCount = 0

    Write-Host "Memproses ${totalUsers} pengguna untuk dihapus..." -ForegroundColor Yellow

    foreach ($entry in $users) {
        $userCount++
        $upn = if ($entry.UserPrincipalName) { $entry.UserPrincipalName.Trim() } else { $null }

        if ([string]::IsNullOrWhiteSpace($upn)) { continue }

        Write-Progress -Activity "Removing Active Users" `
                       -Status "Processing User ${userCount} of ${totalUsers}: ${upn}" `
                       -PercentComplete ([int](($userCount / $totalUsers) * 100))

        Write-Host "-> [${userCount}/${totalUsers}] Menghapus: ${upn}" -ForegroundColor White

        try {
            # 3.1 Verifikasi apakah user ada sebelum dihapus
            $existingUser = Get-MgUser -UserId $upn -ErrorAction Stop | Select-Object Id, DisplayName

            # 3.2 Eksekusi Penghapusan
            # Output di-capture ke $null agar tidak memenuhi layar
            $null = Remove-MgUser -UserId $existingUser.Id -ErrorAction Stop
            
            Write-Host "   ✅ SUCCESS: User berhasil dihapus." -ForegroundColor Green
            $scriptOutput += [PSCustomObject]@{ 
                UserPrincipalName = $upn; 
                DisplayName = $existingUser.DisplayName; 
                Status = "DELETED"; 
                Reason = "User removed successfully" 
            }
        }
        catch {
            $errMsg = $_.Exception.Message
            $reason = if ($errMsg -like "*Resource '"+$upn+"' does not exist*") { "User tidak ditemukan." } else { "Error: $errMsg" }
            
            Write-Host "   ❌ FAILED: ${reason}" -ForegroundColor Red
            $scriptOutput += [PSCustomObject]@{ 
                UserPrincipalName = $upn; 
                DisplayName = "N/A"; 
                Status = "FAIL"; 
                Reason = $reason 
            }
        }
    }
}

## -----------------------------------------------------------------------
## 4. EKSPOR HASIL
## -----------------------------------------------------------------------
if ($scriptOutput.Count -gt 0) {
    try {
        $scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
        Write-Host "`n✅ Laporan Penghapusan Tersimpan: ${outputFilePath}" -ForegroundColor Cyan
    } catch {
        Write-Error "Gagal menyimpan laporan CSV."
    }
}

Write-Host "`nSkrip ${scriptName} selesai dieksekusi." -ForegroundColor Yellow