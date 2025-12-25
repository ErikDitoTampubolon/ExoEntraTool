# =========================================================================
# LISENSI MICROSOFT GRAPH ASSIGNMENT/REMOVAL SCRIPT V19.4 (Fixed & Clean)
# AUTHOR: Erik Dito Tampubolon - TelkomSigma
# =========================================================================

# 1. Konfigurasi File Input & Path
$inputFileName = "UserPrincipalName.csv"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$inputFilePath = Join-Path -Path $scriptDir -ChildPath $inputFileName

$defaultUsageLocation = 'ID'
$operationType = "" 

## -----------------------------------------------------------------------
## 2. KONEKSI KE MICROSOFT GRAPH
## -----------------------------------------------------------------------
$requiredScopes = "User.ReadWrite.All", "Organization.Read.All"
Write-Host "`n--- 2. Membangun Koneksi ke Microsoft Graph ---" -ForegroundColor Blue

if (Get-MgContext -ErrorAction SilentlyContinue) {
    Write-Host "Sesi Microsoft Graph aktif." -ForegroundColor Green
} else {
    Write-Host "Menghubungkan ke Microsoft Graph..." -ForegroundColor Cyan
    try {
        # Menggunakan -ContextScope Process untuk stabilitas sesi di dalam EXE
        Connect-MgGraph -Scopes $requiredScopes -ContextScope Process -ErrorAction Stop | Out-Null
        Write-Host "Koneksi Berhasil." -ForegroundColor Green
    } catch {
        Write-Error "Gagal terhubung ke Microsoft Graph."
        return
    }
}

## -----------------------------------------------------------------------
## 3. PEMILIHAN OPERASI DAN LISENSI 
## -----------------------------------------------------------------------
Write-Host "`n--- 3. Pemilihan Operasi ---" -ForegroundColor Blue
Write-Host "1. Assign License"
Write-Host "2. Remove License"
$operationChoice = Read-Host "Pilih nomor menu"

switch ($operationChoice) {
    "1" { $operationType = "ASSIGN" }
    "2" { $operationType = "REMOVE" }
    default { Write-Host "Pilihan tidak valid." -ForegroundColor Red; return }
}

try {
    $availableLicenses = Get-MgSubscribedSku | Select-Object SkuPartNumber, SkuId -ErrorAction Stop
    Write-Host "`nLisensi yang Tersedia:" -ForegroundColor Yellow
    [int]$index = 1
    $promptOptions = @{}
    foreach ($lic in $availableLicenses) {
        Write-Host "${index}. $($lic.SkuPartNumber)" -ForegroundColor Magenta
        $promptOptions.Add($index, $lic)
        $index++
    }
    
    $choiceInput = Read-Host "`nMasukkan nomor lisensi"
    if (-not $promptOptions.ContainsKey([int]$choiceInput)) { throw "Nomor tidak valid." }
    
    $selectedLicense = $promptOptions[[int]$choiceInput]
    $skuPartNumberTarget = $selectedLicense.SkuPartNumber
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    return
}

## -----------------------------------------------------------------------
## 4. PROSES LOGIKA UTAMA
## -----------------------------------------------------------------------
$allResults = @()
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"

if (-not (Test-Path -Path $inputFilePath)) {
    Write-Host "File ${inputFileName} tidak ditemukan di ${scriptDir}!" -ForegroundColor Red
    return
}

# Import CSV tanpa header sesuai instruksi sebelumnya
$users = Import-Csv -Path $inputFilePath -Header "UserPrincipalName" -ErrorAction SilentlyContinue
$totalUsers = $users.Count
$userCount = 0 

Write-Host "`n--- 4. Memproses ${totalUsers} Pengguna ---" -ForegroundColor Blue

foreach ($entry in $users) {
    $userCount++
    $userUpn = if ($entry.UserPrincipalName) { $entry.UserPrincipalName.Trim() } else { $null }
    if ([string]::IsNullOrWhiteSpace($userUpn)) { continue }

    # Output tampilan baris tunggal (Putih)
    Write-Host "-> [${userCount}/${totalUsers}] Memproses: ${userUpn} . . ." -ForegroundColor White

    try {
        # Ambil User Property
        $user = Get-MgUser -UserId $userUpn -Property 'Id', 'DisplayName', 'UsageLocation' -ErrorAction Stop
        
        # Atur UsageLocation jika ASSIGN dan belum ada (Syarat lisensi M365)
        if ($operationType -eq "ASSIGN" -and -not $user.UsageLocation) {
            $null = Update-MgUser -UserId $user.Id -UsageLocation $defaultUsageLocation -ErrorAction Stop
            $user.UsageLocation = $defaultUsageLocation
        }

        # Cek apakah user sudah punya lisensi tersebut
        $userLicense = Get-MgUserLicenseDetail -UserId $user.Id | Where-Object { $_.SkuId -eq $selectedLicense.SkuId }

        if ($operationType -eq "ASSIGN") {
            if ($userLicense) {
                $status = "ALREADY_ASSIGNED"; $reason = "Sudah memiliki lisensi ini."
            } else {
                $null = Set-MgUserLicense -UserId $user.Id -AddLicenses @(@{ SkuId = $selectedLicense.SkuId }) -RemoveLicenses @() -ErrorAction Stop
                $status = "SUCCESS"; $reason = "Lisensi berhasil diberikan."
            }
        } else {
            if (-not $userLicense) {
                $status = "ALREADY_REMOVED"; $reason = "User memang tidak memiliki lisensi ini."
            } else {
                $null = Set-MgUserLicense -UserId $user.Id -RemoveLicenses @($selectedLicense.SkuId) -AddLicenses @() -ErrorAction Stop
                $status = "SUCCESS_REMOVED"; $reason = "Lisensi berhasil dihapus."
            }
        }

        $allResults += [PSCustomObject]@{
            UserPrincipalName = $userUpn
            DisplayName       = $user.DisplayName
            Status            = $status
            Reason            = $reason
        }
    }
    catch {
        Write-Host "   Gagal: $($_.Exception.Message)" -ForegroundColor Red
        $allResults += [PSCustomObject]@{
            UserPrincipalName = $userUpn
            DisplayName       = "Error/Not Found"
            Status            = "FAIL"
            Reason            = $_.Exception.Message
        }
    }
}

## -----------------------------------------------------------------------
## 5. EKSPOR HASIL
## -----------------------------------------------------------------------
if ($allResults.Count -gt 0) {
    $outputFileName = "${operationType}_License_Results_${timestamp}.csv"
    # PERBAIKAN TYPO: ChildPath (sebelumnya ChsildPatsh)
    $resultsFilePath = Join-Path -Path $scriptDir -ChildPath $outputFileName
    
    $allResults | Export-Csv -Path $resultsFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
    
    Write-Host "`nSemua proses selesai!" -ForegroundColor Green
    Write-Host "Laporan tersimpan di: ${resultsFilePath}" -ForegroundColor Cyan
}

# Sesi tetap terbuka untuk kembali ke menu utama EXE