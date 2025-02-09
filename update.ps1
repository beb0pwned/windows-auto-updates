[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Update-Windows {
    Write-Host "Installing Dependencies..." -ForegroundColor Cyan

    Write-Host "Installing NuGet..." -ForegroundColor Cyan
    if (-not (Get-Module -ListAvailable -Name NuGet)) {
        Install-PackageProvider -Name NuGet -Force -Confirm:$false
        Install-Module -Name NuGet -Force -Confirm:$false
        Write-Host "Successfully Installed NuGet! Continuing Updates..." -ForegroundColor Magenta
    }

    else {
        Write-Host "NuGet is already installed. Continuing Updates..." -ForegroundColor Magenta
    }

    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    Write-Host "Checking for Windows Updates..." -ForegroundColor Cyan
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate )) {
        Install-Module -Name PSWindowsUpdate -Force -Confirm:$false -SkipPublisherCheck
    }
    
    Get-WindowsUpdate -Install -AcceptAll -AutoReboot
    Write-Host "Windows Updates completed!" -ForegroundColor Cyan


}

function Update-Apps {
    Write-Host "Checking for MS Store Updates..." -ForegroundColor Cyan
    winget upgrade --all --accept-source-agreements --accept-package-agreements
} 


Update-Windows
Update-Apps

Set-ExecutionPolicy Restricted -Confirm:$false -Force

Write-Host "Remember to restart device to complete updates!" -ForegroundColor Magenta
Write-Host ""
Write-Host "Finished Updates!" -ForegroundColor Green
