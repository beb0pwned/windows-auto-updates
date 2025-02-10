[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[console]::OutputEncoding = [System.Text.Encoding]::UTF8

$ANSIColors = @{
    'Reset' = "`e[0m"
    'Red' = "`e[31m"
    'Green' = "`e[32m"
    'Yellow' = "`e[33m"
    'Magenta' = "`e[35m"
    'Cyan' = "`e[36m"
}

function Write-ColoredOutput {
    param (
        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter(Mandatory)]
        [ValidateSet('Red', 'Green', 'Yellow', 'Magenta', 'Cyan')]
        [string]$Color
    )

    $colorCode = $ANSIColors[$Color]
    $resetCode = $ANSIColors['Reset']
    Write-Output "$colorCode $Message$resetCode"
}


Add-Type -TypeDefinition @"
using System.Collections.Concurrent;
"@

$outputQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

$runspacePool = [runspaceFactory]::CreateRunspacePool(1, [Environment]::ProcessorCount)
$runspacePool.Open()

# Windows Updates Runspace
$runspace1 = [powershell]::Create().AddScript({
    param ($queue)

    function Update-Windows {
        $queue.Enqueue("[INFO] Initializing Windows Update search...|Cyan")

        $NoPatches = $false
        $Criteria = "IsInstalled=0 or IsHidden=0"

        $Searcher = New-Object -ComObject Microsoft.Update.Searcher
        $searchResult = $Searcher.Search($Criteria).Updates
        $queue.Enqueue("[INFO] Searching for applicable updates...|Cyan")
        # Search for updates that are not installed

        if ($searchResult.Updates.Count -eq 0) {
            $queue.Enqueue("[WARNING] No new updates found.|Yellow")
            return
        }

        if ($searchResult.count -gt 0) {
            foreach($update in $searchResult) {
                $queue.Enqueue("[UPDATE] Installing: $($update.Title)|Magenta")
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
                $queue.Enqueue("[SUCCESS] Patches installed successfully|Green")
            }
            else
            {
                $queue.Enqueue("[ERROR] Patches failed to install with error code $($InstallResult.HResult)|Red")
            }
        }
    }
    Update-Windows
}).AddArgument($outputQueue)
$runspace1.RunspacePool = $runspacePool
$runspace1Output = $runspace1.BeginInvoke()

# App Updates Runspace
$runspace2 = [powershell]::Create().Addscript({
    param ($queue)

    function Update-Apps {
        $queue.Enqueue("[INFO] Checking for Microsoft Store app updates...|Cyan")
        winget upgrade --all --accept-source-agreements --accept-package-agreements
    }
    Update-Apps
}).AddArgument($outputQueue)
$runspace2.RunspacePool = $runspacePool
$runspace2Output = $runspace2.BeginInvoke()

function Process-Queue {
    param (
        [Parameter(Mandatory)]
        [System.Collections.Concurrent.ConcurrentQueue[string]]$Queue
    )
    $line = $null
    while ($Queue.TryDequeue([ref]$line)) {
        $parts = $line -split '\|'
        $message = $parts[0]
        $color = $parts[1]
        Write-ColoredOutput -Message $message -Color $color
    }
}


while (-not ($runspace1Output.IsCompleted -and $runspace2Output.IsCompleted)) {
    Process-Queue -Queue $outputQueue
    Start-Sleep -Milliseconds 500
}


Process-Queue -Queue $outputQueue



# Collect Outs
$runspace1.EndInvoke($runspace1Output)
$runspace2.EndInvoke($runspace2Output)

$runspace1.Dispose()
$runspace2.Dispose()
$runspacePool.Close()

Set-ExecutionPolicy Restricted -Confirm:$false -Force

Write-Host "Remember to restart device to complete updates." -ForegroundColor Magenta

Write-Host "Finished Updates!" -ForegroundColor Green
