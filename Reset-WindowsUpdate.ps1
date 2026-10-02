#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Resets Windows Update components, then scans, downloads and installs updates.

.DESCRIPTION
    Built to run unattended from MDM (as SYSTEM, in 64-bit PowerShell).

    1. Checks the WU services are not disabled and no reboot is pending
    2. Removes backups left by previous runs and checks free space
    3. Stops the Windows Update related services
    4. Renames the SoftwareDistribution and catroot2 folders
    5. Re-registers the Windows Update DLLs
    6. Starts the services again
    7. With -Repair: runs DISM /RestoreHealth and sfc /scannow
    8. Scans, downloads and installs updates through the Microsoft.Update.Session COM object

    Exit codes: 0 = success, 1 = failure, RebootExitCode = success but a reboot is required.

.PARAMETER MinFreeGB
    Minimum free space (GB) required on the system drive. Default: 10.

.PARAMETER LogPath
    Folder for the log file. Default: C:\Windows\Logs\WUReset.

.PARAMETER SkipInstall
    Only reset and scan; do not download or install updates.

.PARAMETER IgnorePendingReboot
    Run even if Windows already has a reboot pending.

.PARAMETER Repair
    Run DISM /RestoreHealth and sfc /scannow after the reset. Adds 15-30+ minutes.

.PARAMETER RebootExitCode
    Exit code when updates installed but a reboot is required. Default: 0.
    Use 3010 when deploying as a Win32 app so the MDM treats it as a soft reboot.

.EXAMPLE
    .\Reset-WindowsUpdate.ps1

.EXAMPLE
    .\Reset-WindowsUpdate.ps1 -MinFreeGB 20 -SkipInstall
#>

[CmdletBinding()]
param(
    [int]    $MinFreeGB      = 10,
    [string] $LogPath        = "$env:SystemRoot\Logs\WUReset",
    [switch] $SkipInstall,
    [switch] $IgnorePendingReboot,
    [switch] $Repair,
    [int]    $RebootExitCode = 0
)


#region Settings ==============================================================

$ErrorActionPreference = 'Stop'

$RunTimestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogFilePath  = Join-Path $LogPath "WUReset_$RunTimestamp.log"

# Services that keep the update cache folders locked.
# All of them must be stopped before the folders can be renamed.
$WUServiceNames            = @('wuauserv', 'bits', 'cryptsvc', 'msiserver')
$ServiceStopTimeoutSeconds = 60

# The update cache folders. Windows recreates both, empty, on next use.
$WUCacheFolders = @(
    "$env:SystemRoot\SoftwareDistribution"   # update database and downloads
    "$env:SystemRoot\System32\catroot2"      # signature catalogs
)
$RenameAttempts = 3

# DLLs used by Windows Update, BITS and the cryptographic services.
# Many only exist on older Windows versions; missing ones are skipped.
$WUDllNames = @(
    'atl.dll',      'urlmon.dll',   'mshtml.dll',   'shdocvw.dll',  'browseui.dll',
    'jscript.dll',  'vbscript.dll', 'scrrun.dll',   'msxml.dll',    'msxml3.dll',
    'msxml6.dll',   'actxprxy.dll', 'softpub.dll',  'wintrust.dll', 'dssenh.dll',
    'rsaenh.dll',   'gpkcsp.dll',   'sccbase.dll',  'slbcsp.dll',   'cryptdlg.dll',
    'oleaut32.dll', 'ole32.dll',    'shell32.dll',  'initpki.dll',  'wuapi.dll',
    'wuaueng.dll',  'wuaueng1.dll', 'wucltui.dll',  'wups.dll',     'wups2.dll',
    'wuweb.dll',    'qmgr.dll',     'qmgrprxy.dll', 'wucltux.dll',  'muweb.dll',
    'wuwebv.dll'
)

# Result codes returned by the Windows Update API (OperationResultCode).
$WUResultNames = @{
    0 = 'NotStarted'
    1 = 'InProgress'
    2 = 'Succeeded'
    3 = 'SucceededWithErrors'
    4 = 'Failed'
    5 = 'Aborted'
}
$WUResultSucceeded = 2

# Reboot behaviour of an update (InstallationBehavior.RebootBehavior).
$WURebootBehaviorNames = @{
    0 = 'No reboot'
    1 = 'Always needs reboot'
    2 = 'May need reboot'
}

# regsvr32 exit codes (FAIL_ARGS .. FAIL_REG in its source).
$RegsvrExitMeanings = @{
    1 = 'invalid arguments passed to regsvr32'
    2 = 'OLE/COM could not be initialised in regsvr32'
    3 = 'DLL could not be loaded: corrupt, wrong architecture, or a dependency is missing'
    4 = 'DLL has no registration entry point: it does not support regsvr32, harmless'
    5 = 'DLL registration ran but failed: usually registry access denied or a damaged DLL'
}
$RegsvrNotRegistrable = 4

# Set by any step that fails, and by the install when Windows asks for a reboot.
# Together they decide the exit code at the end of the script.
$script:Failed         = $false
$script:RebootRequired = $false

#endregion


#region Helpers ===============================================================

<#
.SYNOPSIS
    Writes one line to the log file and to the console.
.DESCRIPTION
    Each line is a timestamp followed by the message.
    The console output is what the MDM captures; the log file stays on the
    device for troubleshooting afterwards.
#>
function Write-Log {
    param(
        [string] $Message
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logLine   = "$timestamp $Message"

    Add-Content -Path $LogFilePath -Value $logLine
    Write-Host $logLine
}

<#
.SYNOPSIS
    Runs a program hidden, waits for it to finish and returns its exit code.
.DESCRIPTION
    Used for regsvr32, DISM and SFC, which don't need any input.
#>
function Invoke-HiddenProcess {
    param(
        [string] $FilePath,
        [string] $Arguments
    )

    $processOptions = @{
        FilePath     = $FilePath
        ArgumentList = $Arguments
        Wait         = $true
        PassThru     = $true
        WindowStyle  = 'Hidden'
    }

    $process = Start-Process @processOptions
    return $process.ExitCode
}

<#
.SYNOPSIS
    Formats a COM HResult as 0x8024xxxx.
.DESCRIPTION
    The Windows Update API reports errors as a negative Int32. Shown in hex,
    the code can be looked up in Microsoft's Windows Update error code list.
#>
function Format-HResult {
    param(
        [int] $HResult
    )

    return '0x{0:X8}' -f $HResult
}

#endregion


#region Pre-checks ============================================================

<#
.SYNOPSIS
    Returns $false if any Windows Update service is set to Disabled.
.DESCRIPTION
    A disabled service is usually set on purpose by GPO, MDM policy or an
    "update blocker" tool. The script can't start it again, and a reset would
    leave the device without working updates, so the script stops before
    changing anything.
#>
function Test-ServiceStartup {
    $allServicesEnabled = $true

    foreach ($serviceName in $WUServiceNames) {
        $startType = (Get-Service -Name $serviceName).StartType
        Write-Log "Service $serviceName start type: $startType"

        if ($startType -eq 'Disabled') {
            Write-Log "Service $serviceName is disabled (check GPO/MDM policy)."
            $allServicesEnabled = $false
        }
    }

    return $allServicesEnabled
}

<#
.SYNOPSIS
    Returns $true if Windows is already waiting for a reboot.
.DESCRIPTION
    Checks the registry keys that the servicing stack (CBS) and Windows Update
    create when a reboot is needed. Installing more updates on top of a pending
    reboot often fails, so the script stops unless -IgnorePendingReboot is used.
#>
function Test-PendingReboot {
    $rebootRegistryKeys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    )

    foreach ($registryKey in $rebootRegistryKeys) {
        if (Test-Path $registryKey) {
            Write-Log "Pending reboot detected: $registryKey"
            return $true
        }
    }

    Write-Log 'No pending reboot detected.'
    return $false
}

<#
.SYNOPSIS
    Returns $false if the system drive has less than -MinFreeGB free.
.DESCRIPTION
    Updates are downloaded to and installed on the system drive. Cumulative and
    feature updates need several GB, and running out mid-install can leave
    Windows in a bad state.

    Uses .NET DriveInfo instead of WMI/CIM: devices with broken Windows Update
    often have a broken WMI repository too, and this check must still work there.
#>
function Test-DriveSpace {
    $systemDrive = [System.IO.DriveInfo]::new($env:SystemDrive)
    $freeGB      = [math]::Round($systemDrive.AvailableFreeSpace / 1GB, 2)

    Write-Log "Free space on $env:SystemDrive : $freeGB GB (minimum $MinFreeGB GB)"

    if ($freeGB -lt $MinFreeGB) {
        Write-Log "Not enough free space on $env:SystemDrive."
        return $false
    }

    return $true
}

#endregion


#region Services ==============================================================

<#
.SYNOPSIS
    Kills the process of a service stuck in "Stop pending".
.DESCRIPTION
    Last resort, only used after the stop timeout.
    Several Windows services can share one svchost.exe process. Killing a shared
    one would take unrelated services down with it, so the process is only
    killed when this service is the only one running in it.
#>
function Stop-ServiceProcess {
    param(
        [string] $ServiceName
    )

    $serviceProcessId = (Get-CimInstance Win32_Service -Filter "Name='$ServiceName'").ProcessId
    if (-not $serviceProcessId) {
        return
    }

    $servicesInProcess = @(Get-CimInstance Win32_Service -Filter "ProcessId=$serviceProcessId")
    if ($servicesInProcess.Count -gt 1) {
        Write-Log "$ServiceName shares process $serviceProcessId with other services, not killing it"
        return
    }

    try {
        Stop-Process -Id $serviceProcessId -Force -ErrorAction Stop
        Write-Log "Killed process $serviceProcessId for $ServiceName"
    } catch {
        Write-Log "Could not kill process $serviceProcessId for $ServiceName : $($_.Exception.Message)"
    }
}

<#
.SYNOPSIS
    Stops the Windows Update services, waiting up to the stop timeout for each.
.DESCRIPTION
    -Force also stops services that depend on them. A service that doesn't stop
    in time is handed to Stop-ServiceProcess.
    Services that are already stopped are skipped, so this is safe to call
    again (Reset-WUFolders does that between rename attempts).
#>
function Stop-WUServices {
    $stopTimeout = [TimeSpan]::FromSeconds($ServiceStopTimeoutSeconds)

    foreach ($serviceName in $WUServiceNames) {
        $service = Get-Service -Name $serviceName
        if ($service.Status -eq 'Stopped') {
            continue
        }

        Write-Log "Stopping service $serviceName"

        try {
            Stop-Service -Name $serviceName -Force -NoWait -ErrorAction Stop
            $service.WaitForStatus('Stopped', $stopTimeout)
        } catch {
            Write-Log "$serviceName did not stop within $ServiceStopTimeoutSeconds s: $($_.Exception.Message)"
            Stop-ServiceProcess -ServiceName $serviceName
        }
    }
}

<#
.SYNOPSIS
    Starts the Windows Update services again after the reset.
.DESCRIPTION
    A service that won't start marks the run as failed, because updates can't
    be scanned or installed without it.
#>
function Start-WUServices {
    foreach ($serviceName in $WUServiceNames) {
        Write-Log "Starting service $serviceName"

        try {
            Start-Service -Name $serviceName -ErrorAction Stop
        } catch {
            Write-Log "Could not start $serviceName : $($_.Exception.Message)"
            $script:Failed = $true
        }
    }
}

#endregion


#region Cache folders =========================================================

<#
.SYNOPSIS
    Deletes the backup folders left by earlier runs.
.DESCRIPTION
    Removes SoftwareDistribution.bak_* and catroot2.bak_*. Each backup can be
    several GB, so this runs before the drive space check. The backup made by
    the current run is kept until the next run.
#>
function Remove-OldBackups {
    foreach ($cacheFolder in $WUCacheFolders) {
        $parentFolder  = Split-Path $cacheFolder -Parent
        $backupPattern = "$(Split-Path $cacheFolder -Leaf).bak_*"
        $oldBackups    = Get-ChildItem -Path $parentFolder -Directory -Filter $backupPattern

        foreach ($backupFolder in $oldBackups) {
            try {
                Remove-Item -Path $backupFolder.FullName -Recurse -Force -ErrorAction Stop
                Write-Log "Removed old backup $($backupFolder.FullName)"
            } catch {
                Write-Log "Could not remove $($backupFolder.FullName) : $($_.Exception.Message)"
            }
        }
    }
}

<#
.SYNOPSIS
    Renames SoftwareDistribution and catroot2 to <name>.bak_<timestamp>.
.DESCRIPTION
    This is the actual reset: Windows recreates both folders empty when the
    services start, which clears a corrupt update database or catalog.

    Windows can restart a service between the stop and the rename, which locks
    the folder again. Each rename is therefore tried several times, with the
    services stopped again in between. If a folder still can't be renamed, the
    run is marked as failed.
#>
function Reset-WUFolders {
    foreach ($cacheFolder in $WUCacheFolders) {
        if (-not (Test-Path $cacheFolder)) {
            Write-Log "$cacheFolder does not exist, nothing to rename"
            continue
        }

        $backupName = "$(Split-Path $cacheFolder -Leaf).bak_$RunTimestamp"
        $isRenamed  = $false

        for ($attempt = 1; $attempt -le $RenameAttempts -and -not $isRenamed; $attempt++) {
            try {
                Rename-Item -Path $cacheFolder -NewName $backupName -ErrorAction Stop
                Write-Log "Renamed $cacheFolder -> $backupName"
                $isRenamed = $true
            } catch {
                Write-Log "Rename attempt $attempt of $cacheFolder failed: $($_.Exception.Message)"
                Start-Sleep -Seconds 5
                Stop-WUServices
            }
        }

        if (-not $isRenamed) {
            Write-Log "Could not rename $cacheFolder, reset is incomplete."
            $script:Failed = $true
        }
    }
}

#endregion


#region DLL registration ======================================================

<#
.SYNOPSIS
    Re-registers the Windows Update related DLLs with regsvr32 /s (silent).
.DESCRIPTION
    Repairs broken COM registrations, a cause of errors like 0x80070002 or
    "class not registered".

    Every non-zero regsvr32 exit code is logged with its meaning. Code 4 is
    expected for several DLLs on Windows 10/11 and is harmless; codes 3 and 5
    point at a damaged or blocked DLL and are worth looking into (try -Repair).
    Ends with a count of each outcome.
#>
function Register-WUDlls {
    $system32Path        = "$env:SystemRoot\System32"
    $registeredCount     = 0
    $notRegistrableCount = 0
    $failedCount         = 0
    $missingCount        = 0

    foreach ($dllName in $WUDllNames) {
        $dllPath = Join-Path $system32Path $dllName

        if (-not (Test-Path $dllPath)) {
            $missingCount++
            continue
        }

        $exitCode = Invoke-HiddenProcess -FilePath 'regsvr32.exe' -Arguments "/s `"$dllPath`""

        if ($exitCode -eq 0) {
            Write-Log "Registered $dllName"
            $registeredCount++
            continue
        }

        if ($RegsvrExitMeanings.ContainsKey($exitCode)) {
            $meaning = $RegsvrExitMeanings[$exitCode]
        } else {
            $meaning = 'unknown regsvr32 exit code'
        }
        Write-Log "regsvr32 $dllName returned $exitCode - $meaning"

        if ($exitCode -eq $RegsvrNotRegistrable) {
            $notRegistrableCount++
        } else {
            $failedCount++
        }
    }

    Write-Log ("DLLs: {0} registered, {1} not registrable (harmless), {2} failed, {3} not present" -f
        $registeredCount, $notRegistrableCount, $failedCount, $missingCount)
}

#endregion


#region Repair ================================================================

<#
.SYNOPSIS
    Runs DISM /RestoreHealth and sfc /scannow. Only used with -Repair.
.DESCRIPTION
    Fixes corruption in Windows itself, which a reset alone can't fix
    (typical errors: 0x800f081f, 0x80073712).

    - DISM /RestoreHealth repairs the component store, downloading clean files
      from Windows Update. That's why it runs after the services are started.
    - sfc /scannow then repairs protected system files from the component store.

    A DISM failure marks the run as failed. SFC's exit code is only logged,
    because SFC doesn't reliably report problems through it.
#>
function Invoke-Repair {
    $system32Path = "$env:SystemRoot\System32"

    Write-Log 'Running DISM /RestoreHealth (details in C:\Windows\Logs\DISM\dism.log)...'
    $dismExitCode = Invoke-HiddenProcess -FilePath "$system32Path\dism.exe" -Arguments '/Online /Cleanup-Image /RestoreHealth'
    Write-Log "DISM exit code: $dismExitCode"

    if ($dismExitCode -ne 0) {
        $script:Failed = $true
    }

    Write-Log 'Running sfc /scannow (details in C:\Windows\Logs\CBS\CBS.log)...'
    $sfcExitCode = Invoke-HiddenProcess -FilePath "$system32Path\sfc.exe" -Arguments '/scannow'
    Write-Log "SFC exit code: $sfcExitCode"
}

#endregion


#region Windows Update ========================================================

<#
.SYNOPSIS
    Scans for updates and returns the ones to install.
.DESCRIPTION
    Searches for software updates that are not installed and not hidden, logs
    the details of each one (KB, category, severity, size, reboot behaviour)
    and accepts licence agreements so the install doesn't stall.
    Returns a Microsoft.Update.UpdateColl, which may be empty.
#>
function Find-PendingUpdates {
    param(
        $UpdateSession
    )

    $searchCriteria = "IsInstalled=0 and Type='Software' and IsHidden=0"
    $searcher       = $UpdateSession.CreateUpdateSearcher()

    Write-Log "Scanning for updates (criteria: $searchCriteria)..."

    $stopwatch    = [Diagnostics.Stopwatch]::StartNew()
    $searchResult = $searcher.Search($searchCriteria)
    $scanSeconds  = [int]$stopwatch.Elapsed.TotalSeconds
    $scanResult   = $WUResultNames[[int]$searchResult.ResultCode]

    Write-Log "Scan finished in $scanSeconds s, result: $scanResult"
    Write-Log "Found $($searchResult.Updates.Count) update(s)"

    $pendingUpdates = New-Object -ComObject Microsoft.Update.UpdateColl
    $totalSizeMB    = 0
    $updateNumber   = 0
    $updateCount    = $searchResult.Updates.Count

    foreach ($update in $searchResult.Updates) {
        $updateNumber++

        $kbNumbers      = (@($update.KBArticleIDs) | ForEach-Object { "KB$_" }) -join ', '
        $categories     = (@($update.Categories) | ForEach-Object { $_.Name }) -join ', '
        $sizeMB         = [math]::Round($update.MaxDownloadSize / 1MB, 1)
        $rebootBehavior = $WURebootBehaviorNames[[int]$update.InstallationBehavior.RebootBehavior]

        if ($update.MsrcSeverity) {
            $severity = $update.MsrcSeverity
        } else {
            $severity = 'n/a'
        }

        Write-Log "[$updateNumber/$updateCount] $($update.Title)"
        Write-Log "    KB: $kbNumbers | Category: $categories | Severity: $severity"
        Write-Log "    Size: $sizeMB MB | Downloaded: $($update.IsDownloaded) | Reboot: $rebootBehavior"

        if (-not $update.EulaAccepted) {
            Write-Log '    Accepting licence agreement'
            $update.AcceptEula()
        }

        [void]$pendingUpdates.Add($update)
        $totalSizeMB += $sizeMB
    }

    if ($updateCount -gt 0) {
        Write-Log "Total download size (max): $totalSizeMB MB"
    }

    # The leading comma stops PowerShell from unrolling the collection.
    return , $pendingUpdates
}

<#
.SYNOPSIS
    Downloads the updates one at a time and returns the ones that downloaded.
.DESCRIPTION
    Downloading one by one gives a log line per update with its result, time
    taken and HResult, so a slow or failing update is easy to spot.
    Updates that are already downloaded are skipped. Any failed download marks
    the run as failed.
    Returns a Microsoft.Update.UpdateColl, which may be empty.
#>
function Save-PendingUpdates {
    param(
        $UpdateSession,
        $Updates
    )

    $downloader        = $UpdateSession.CreateUpdateDownloader()
    $downloadedUpdates = New-Object -ComObject Microsoft.Update.UpdateColl
    $failedCount       = 0

    for ($index = 0; $index -lt $Updates.Count; $index++) {
        $update      = $Updates.Item($index)
        $updateLabel = "[$($index + 1)/$($Updates.Count)] $($update.Title)"

        if ($update.IsDownloaded) {
            Write-Log "Already downloaded: $updateLabel"
            [void]$downloadedUpdates.Add($update)
            continue
        }

        Write-Log "Downloading: $updateLabel"

        $singleUpdate = New-Object -ComObject Microsoft.Update.UpdateColl
        [void]$singleUpdate.Add($update)
        $downloader.Updates = $singleUpdate

        $stopwatch = [Diagnostics.Stopwatch]::StartNew()
        try {
            $downloadResult = $downloader.Download()
            $resultName     = $WUResultNames[[int]$downloadResult.ResultCode]
            $seconds        = [int]$stopwatch.Elapsed.TotalSeconds
            $hResult        = Format-HResult $downloadResult.HResult

            Write-Log "    $resultName in $seconds s (HResult $hResult)"
        } catch {
            Write-Log "    Download error: $($_.Exception.Message)"
        }

        if ($update.IsDownloaded) {
            [void]$downloadedUpdates.Add($update)
        } else {
            $failedCount++
            $script:Failed = $true
        }
    }

    Write-Log "Downloads: $($downloadedUpdates.Count) ready, $failedCount failed"

    # The leading comma stops PowerShell from unrolling the collection.
    return , $downloadedUpdates
}

<#
.SYNOPSIS
    Installs the downloaded updates in one batch and logs each result.
.DESCRIPTION
    One batch lets Windows install prerequisites, such as servicing stack
    updates, in the right order.
    Logs the overall result, then the result, HResult and reboot need of each
    update, then a summary. A failed install marks the run as failed; a reboot
    request sets $script:RebootRequired.
#>
function Install-DownloadedUpdates {
    param(
        $UpdateSession,
        $Updates
    )

    Write-Log "Installing $($Updates.Count) update(s), this can take a while..."

    $installer         = $UpdateSession.CreateUpdateInstaller()
    $installer.Updates = $Updates

    $stopwatch     = [Diagnostics.Stopwatch]::StartNew()
    $installResult = $installer.Install()
    $minutes       = [int]$stopwatch.Elapsed.TotalMinutes
    $resultName    = $WUResultNames[[int]$installResult.ResultCode]
    $hResult       = Format-HResult $installResult.HResult

    Write-Log "Install finished in $minutes min, result: $resultName (HResult $hResult)"

    if ([int]$installResult.ResultCode -ne $WUResultSucceeded) {
        $script:Failed = $true
    }

    $installedCount = 0
    $failedCount    = 0

    for ($index = 0; $index -lt $Updates.Count; $index++) {
        $updateResult     = $installResult.GetUpdateResult($index)
        $updateResultCode = [int]$updateResult.ResultCode
        $updateTitle      = $Updates.Item($index).Title

        if ($updateResultCode -eq $WUResultSucceeded) {
            $installedCount++
        } else {
            $failedCount++
        }

        Write-Log "[$($index + 1)/$($Updates.Count)] $($WUResultNames[$updateResultCode]): $updateTitle"
        Write-Log "    HResult $(Format-HResult $updateResult.HResult) | Reboot required: $($updateResult.RebootRequired)"
    }

    Write-Log "Install summary: $installedCount installed, $failedCount failed"

    if ($installResult.RebootRequired) {
        Write-Log 'A reboot is required to finish installing updates.'
        $script:RebootRequired = $true
    }
}

<#
.SYNOPSIS
    Scans, downloads and installs updates through the Windows Update Agent API.
.DESCRIPTION
    Uses the Microsoft.Update.Session COM object, the same API Windows Update
    itself uses. Stops after the scan when -SkipInstall is used.
#>
function Invoke-WUInstall {
    $agentVersion = (Get-Item "$env:SystemRoot\System32\wuaueng.dll").VersionInfo.ProductVersion
    Write-Log "Windows Update Agent version: $agentVersion"

    $updateSession = New-Object -ComObject Microsoft.Update.Session
    $updateSession.ClientApplicationID = 'Reset-WindowsUpdate'

    $pendingUpdates = Find-PendingUpdates -UpdateSession $updateSession
    if ($pendingUpdates.Count -eq 0) {
        return
    }

    if ($SkipInstall) {
        Write-Log 'SkipInstall set, not downloading or installing.'
        return
    }

    $downloadedUpdates = Save-PendingUpdates -UpdateSession $updateSession -Updates $pendingUpdates
    if ($downloadedUpdates.Count -eq 0) {
        Write-Log 'No updates were downloaded, nothing to install.'
        $script:Failed = $true
        return
    }

    Install-DownloadedUpdates -UpdateSession $updateSession -Updates $downloadedUpdates
}

#endregion


#region Main ==================================================================

# Order matters: checks that change nothing run first, so a device this script
# can't fix is left untouched. Any unexpected error restarts the services so the
# device isn't left without Windows Update.

New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
Write-Log "=== Windows Update reset started on $env:COMPUTERNAME as $env:USERNAME ==="

try {
    # Checks: stop before changing anything if the device can't be fixed now.
    if (-not (Test-ServiceStartup)) {
        exit 1
    }

    if ((Test-PendingReboot) -and -not $IgnorePendingReboot) {
        Write-Log 'Reboot the device first, or run with -IgnorePendingReboot.'
        exit 1
    }

    Remove-OldBackups

    if (-not (Test-DriveSpace)) {
        exit 1
    }

    # Reset.
    Stop-WUServices
    Reset-WUFolders
    Register-WUDlls
    Start-WUServices

    if ($Repair) {
        Invoke-Repair
    }

    # Update.
    Invoke-WUInstall
} catch {
    Write-Log "Unexpected error: $($_.Exception.Message)"
    $script:Failed = $true
    Start-WUServices
}

# The last line is what most MDM consoles show as the script output.
if ($script:Failed) {
    Write-Log 'RESULT: Failed, see log for details.'
    exit 1
}

if ($script:RebootRequired) {
    Write-Log 'RESULT: Success, reboot required.'
    exit $RebootExitCode
}

Write-Log 'RESULT: Success.'
exit 0

#endregion
