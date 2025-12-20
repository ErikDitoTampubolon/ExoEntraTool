# =========================================================================
# CREATE ACTIVE USERS - MICROSOFT GRAPH SDK (V6.2 - Domain Fix)
# =========================================================================

$scriptName = "CreateActiveUsers"
$scriptOutput = @() 
$inputFileName = "ActiveUser.csv" 

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$inputFilePath = Join-Path -Path $scriptDir -ChildPath $inputFileName
$outputFilePath = Join-Path -Path $scriptDir -ChildPath "Output_${scriptName}_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

## -----------------------------------------------------------------------
## 2. KONEKSI (SILENT MODE)
## -----------------------------------------------------------------------
if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
    Connect-MgGraph -Scopes "User.ReadWrite.All", "Domain.Read.All" -ErrorAction Stop | Out-Null
}

## -----------------------------------------------------------------------
## 3. LOGIKA UTAMA
## -----------------------------------------------------------------------
Write-Host "`n--- Memulai Pembuatan User ---" -ForegroundColor Magenta

if (-not (Test-Path -Path $inputFilePath)) {
    Write-Error "File $inputFileName tidak ditemukan!"
} else {
    # Membaca CSV dengan Header sesuai contoh Bapak
    $users = Import-Csv -Path $inputFilePath -ErrorAction SilentlyContinue
    
    # Ambil daftar domain yang valid di tenant untuk verifikasi awal
    $verifiedDomains = (Get-MgDomain | Where-Object { $_.IsVerified -eq $true }).Id

    foreach ($user in $users) {
        # CLEANING: Trim semua spasi di setiap kolom
        $fName    = $user.FirstName.Trim()
        $lName    = $user.LastName.Trim()
        $uName    = $user.Username.Trim()
        $dom      = $user.Domain.Trim()
        $password = $user.InitialPassword.Trim()
        $dName    = if ($user.DisplayName) { $user.DisplayName.Trim() } else { "${fName} ${lName}" }
        
        $upn = "${uName}@${dom}"

        Write-Host "-> Memproses: ${upn}" -ForegroundColor White

        # --- 3.1 Verifikasi Domain ---
        if ($verifiedDomains -notcontains $dom) {
            $reason = "Domain '${dom}' TIDAK VALID/TIDAK TERVERIFIKASI di tenant ini."
            Write-Host "   ❌ ERROR: ${reason}" -ForegroundColor Red
            $scriptOutput += [PSCustomObject]@{ UPN = $upn; Status = "FAIL"; Reason = $reason }
            continue
        }

        # --- 3.2 Pembuatan User ---
        try {
            $params = @{
                AccountEnabled = $true
                DisplayName    = $dName
                MailNickname   = $uName
                UserPrincipalName = $upn
                UsageLocation  = "ID" 
                GivenName      = $fName
                Surname        = $lName
                PasswordProfile = @{
                    ForceChangePasswordNextSignIn = $true
                    Password = $password
                }
            }

            $null = New-MgUser @params -ErrorAction Stop
            Write-Host "   ✅ SUCCESS: User Dibuat." -ForegroundColor Green
            $scriptOutput += [PSCustomObject]@{ UPN = $upn; Status = "CREATED"; Reason = "Success" }
        }
        catch {
            $errMsg = $_.Exception.Message
            Write-Host "   ❌ ERROR: ${errMsg}" -ForegroundColor Red
            $scriptOutput += [PSCustomObject]@{ UPN = $upn; Status = "FAIL"; Reason = $errMsg }
        }
    }
}

# Ekspor Hasil
$scriptOutput | Export-Csv -Path $outputFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
Write-Host "`n✅ Proses selesai. Laporan: $outputFilePath" -ForegroundColor Cyan