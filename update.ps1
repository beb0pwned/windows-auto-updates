[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Update-Windows {
    Write-Host "Initializing Windows Update search..." -ForegroundColor Cyan

    $NoPatches = $false
    $Criteria = "IsInstalled=0 or IsHidden=0"

    $Searcher = New-Object -ComObject Microsoft.Update.Searcher
    $searchResult = $Searcher.Search($Criteria).Updates
    Write-Host "Searching for applicable updates..." -ForegroundColor Cyan
    
    # Search for updates that are not installed
    if ($searchResult.Updates.Count -eq 0) {
        Write-Host "No new updates found." -ForegroundColor Green
    }

    if ($searchResult.count -gt 0) {
        foreach($item in $searchResult) {
            Write-Host "Preparing to install" $item.Title -ForegroundColor Yellow
        }
    } 
    else {
        $NoPatches = $true
    }

    if($NoPatches -eq $false) {
        $Session = New-Object -ComObject Microsoft.Update.Session
        $Downloader = $Session.CreateUpdateDownloader()
        $Downloader.Updates = $SearchResult
        $DownloadResult = $Downloader.Download()

        #Install updates
        $Installer = New-Object -ComObject Microsoft.Update.Installer
        $Installer.Updates = $SearchResult
        $InstallResult = $Installer.Install()

        if($InstallResult.HResult -eq 0)
        {
            Write-Host "Patches installed successfully" -ForegroundColor Green
        }
        else
        {
            Write-Host "Patches failed to install with error code" $InstallResult.HResult -ForegroundColor Red
        }
    }
}

function Update-Apps {
    Write-Host "Checking for MS Store Updates..." -ForegroundColor Cyan
    winget upgrade --all --accept-source-agreements --accept-package-agreements
} 

Update-Windows
Update-Apps

Set-ExecutionPolicy Restricted -Confirm:$false -Force

Write-Host "Remember to restart device to complete updates!" -ForegroundColor Magenta
Write-Host
Write-Host "Finished Updates!" -ForegroundColor Green
