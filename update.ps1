[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$runspace1Script = {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    function Update-Windows {
        Write-Host "[INFO] Initializing Windows Update search..." -ForegroundColor Cyan

        $NoPatches = $false
        $Criteria = "IsInstalled=0 or IsHidden=0"

        $Searcher = New-Object -ComObject Microsoft.Update.Searcher
        $searchResult = $Searcher.Search($Criteria).Updates
        Write-Host "[INFO] Searching for applicable updates..." -ForegroundColor Cyan

        if ($searchResult.Count -eq 0) {
            Write-Host "[WARNING] No new updates found." -ForegroundColor Yellow
            return
        }

        if ($searchResult.Count -gt 0) {
            foreach ($update in $searchResult) {
                Write-Host "[UPDATE] Installing: $($update.Title)" -ForegroundColor Magenta
            }
        } else {
            $NoPatches = $true
        }

        if (-not $NoPatches) {
            $Session = New-Object -ComObject Microsoft.Update.Session
            $Downloader = $Session.CreateUpdateDownloader()
            $Downloader.Updates = $searchResult

            $Installer = New-Object -ComObject Microsoft.Update.Installer
            $Installer.Updates = $searchResult
            $InstallResult = $Installer.Install()

            if ($InstallResult.HResult -eq 0) {
                Write-Host "[SUCCESS] Updates installed successfully!" -ForegroundColor Green
            } else {
                Write-Host "[ERROR] Updates failed to install with error code $($InstallResult.HResult)" -ForegroundColor Red
            }
        }
    }

    Update-Windows
    Write-Host "[INFO] You may now close this window." -ForegroundColor Cyan 
}

# Path to the script or commands for the second runspace
$runspace2Script = {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    function Update-Apps {
        Write-Host "[INFO] Checking for Microsoft Store app updates..." -ForegroundColor Cyan
        winget upgrade --source msstore --all --accept-source-agreements --accept-package-agreements
        winget upgrade --all --accept-source-agreements --accept-package-agreements

        Write-Host "[SUCCESS] Completed all App updates!" -ForegroundColor Green
    }
    Update-Apps
    Write-Host "[INFO] You may now close this window." -ForegroundColor Cyan
}

# Write scripts to temp files
$TempFile = $env:TEMP
$runspace1File = "$TempFile\win-updates.ps1"
$runspace2File = "$TempFile\app-updates.ps1"
Set-Content -Path $runspace1File -Value ($runspace1Script | Out-String)
Set-Content -Path $runspace2File -Value ($runspace2Script | Out-String)

# Start new terminals for each runspace
Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-File", $runspace1File
Start-Process -FilePath "powershell.exe" -ArgumentList "-NoExit", "-File", $runspace2File

$job1 = Start-Job -ScriptBlock ([scriptblock]::Create((Get-Content $runspace1File -Raw)))
$job2 = Start-Job -ScriptBlock ([scriptblock]::Create((Get-Content $runspace2File -Raw)))

Write-Host "[INFO] Waiting for updates to complete..." -ForegroundColor Cyan
Wait-Job -Job $job1, $job2


Remove-Job -Job $job1, $job2
Remove-Item -Path $runspace1File, $runspace2File -Force


Write-Host "[WARNING] Restart device to complete updates." -ForegroundColor Yellow
Write-Host 
Write-Host "Updates Complete!" -ForegroundColor Green
Set-ExecutionPolicy Restricted -Force -Confirm:$false -Scope CurrentUser
