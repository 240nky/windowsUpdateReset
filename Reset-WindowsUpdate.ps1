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
    .\Reset-WindowsUpdate.ps1 -MinFreeGB 20 -SkipInstall
#>
[CmdletBinding()]
param(
    [int]$MinFreeGB = 10,
    [string]$LogPath = "$env:SystemRoot\Logs\WUReset",
    [switch]$SkipInstall,
    [switch]$IgnorePendingReboot,
    [switch]$Repair,
    [int]$RebootExitCode = 0
)

$ErrorActionPreference = 'Stop'
$RunTimestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogFilePath  = Join-Path $LogPath "WUReset_$RunTimestamp.log"

# Services that hold the Windows Update cache folders open. All four must be
# stopped before the folders can be renamed.
$WUServiceNames            = 'wuauserv', 'bits', 'cryptsvc', 'msiserver'
$ServiceStopTimeoutSeconds = 60

# SoftwareDistribution holds the update database and downloads,
# catroot2 holds the signature catalogs. Both are rebuilt by Windows on next use.
$WUCacheFolders = "$env:SystemRoot\SoftwareDistribution", "$env:SystemRoot\System32\catroot2"

# DLLs used by Windows Update, BITS and cryptographic services.
# Many only exist on older Windows versions; missing ones are skipped.
$WUDllNames = @(
    'atl.dll', 'urlmon.dll', 'mshtml.dll', 'shdocvw.dll', 'browseui.dll',
    'jscript.dll', 'vbscript.dll', 'scrrun.dll', 'msxml.dll', 'msxml3.dll',
    'msxml6.dll', 'actxprxy.dll', 'softpub.dll', 'wintrust.dll', 'dssenh.dll',
    'rsaenh.dll', 'gpkcsp.dll', 'sccbase.dll', 'slbcsp.dll', 'cryptdlg.dll',
    'oleaut32.dll', 'ole32.dll', 'shell32.dll', 'initpki.dll', 'wuapi.dll',
    'wuaueng.dll', 'wuaueng1.dll', 'wucltui.dll', 'wups.dll', 'wups2.dll',
    'wuweb.dll', 'qmgr.dll', 'qmgrprxy.dll', 'wucltux.dll', 'muweb.dll',
    'wuwebv.dll'
)

# Set by any step that fails, and by the install when Windows asks for a reboot.
# Together they decide the exit code at the end of the script.
$script:Failed         = $false
$script:RebootRequired = $false

# ---------------------------------------------------------------- logging ---
<#
    Write-Log
    Writes one line to the log file and to the console: a timestamp and the message.
    Console output is what the MDM captures, the file is kept on the device for
    troubleshooting afterwards.
#>
function Write-Log {
    param([string]$Message)
    $line = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Add-Content -Path $LogFilePath -Value $line
    Write-Host $line
}

# --------------------------------------------------------- pre-checks -------
<#
    Test-ServiceStartup
    Returns $false if any of the Windows Update services is set to Disabled.
    A disabled service is usually set on purpose by GPO, MDM policy or an
    "update blocker" tool. The script can't start it again, and a reset would
    leave the device without working updates, so we stop before changing anything.
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
    Test-PendingReboot
    Returns $true if Windows is already waiting for a reboot.
    Checks the two registry keys that the servicing stack (CBS) and Windows
    Update create when a reboot is needed. Installing more updates on top of a
    pending reboot often fails, so the script stops unless -IgnorePendingReboot
    is used.
#>
function Test-PendingReboot {
    $rebootRegistryKeys = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
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
    Test-DriveSpace
    Returns $false if the system drive has less than -MinFreeGB free.
    Updates are downloaded to and installed on the system drive; cumulative and
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

# --------------------------------------------------------------- services ---
<#
    Stop-ServiceProcess
    Last resort for a service stuck in "Stop pending": kills its process.
    Several Windows services can share one svchost.exe process. Killing a shared
    one would take unrelated services down with it, so this only kills the
    process when the service is the only one running in it.
#>
function Stop-ServiceProcess {
    param([string]$ServiceName)
    $serviceProcessId = (Get-CimInstance Win32_Service -Filter "Name='$ServiceName'").ProcessId
    if (-not $serviceProcessId) { return }
    if (@(Get-CimInstance Win32_Service -Filter "ProcessId=$serviceProcessId").Count -gt 1) {
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
    Stop-WUServices
    Stops each Windows Update service and waits up to $ServiceStopTimeoutSeconds for it.
    -Force also stops services that depend on it. If a service doesn't stop in
    time, Stop-ServiceProcess is used. Services already stopped are skipped, so
    this is safe to call again (Reset-WUFolders does that between retries).
#>
function Stop-WUServices {
    foreach ($serviceName in $WUServiceNames) {
        $service = Get-Service -Name $serviceName
        if ($service.Status -eq 'Stopped') { continue }
        Write-Log "Stopping service $serviceName"
        try {
            Stop-Service -Name $serviceName -Force -NoWait -ErrorAction Stop
            $service.WaitForStatus('Stopped', [TimeSpan]::FromSeconds($ServiceStopTimeoutSeconds))
        } catch {
            Write-Log "$serviceName did not stop within $ServiceStopTimeoutSeconds s: $($_.Exception.Message)"
            Stop-ServiceProcess $serviceName
        }
    }
}

<#
    Start-WUServices
    Starts the Windows Update services again after the reset.
    A service that won't start marks the run as failed, because updates can't
    be scanned or installed without it.
#>
function Start-WUServices {
    foreach ($serviceName in $WUServiceNames) {
        try {
            Write-Log "Starting service $serviceName"
            Start-Service -Name $serviceName -ErrorAction Stop
        } catch {
            Write-Log "Could not start $serviceName : $($_.Exception.Message)"
            $script:Failed = $true
        }
    }
}

# ------------------------------------------------------- cache folders ------
<#
    Remove-OldBackups
    Deletes SoftwareDistribution.bak_* and catroot2.bak_* folders left by earlier
    runs. Each backup can be several GB, so this runs before the drive space
    check. The backup made by the current run is kept until the next run.
#>
function Remove-OldBackups {
    foreach ($cacheFolder in $WUCacheFolders) {
        $backupPattern = "$(Split-Path $cacheFolder -Leaf).bak_*"
        foreach ($backupFolder in Get-ChildItem -Path (Split-Path $cacheFolder -Parent) -Directory -Filter $backupPattern) {
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
    Reset-WUFolders
    Renames SoftwareDistribution and catroot2 to <name>.bak_<timestamp>.
    This is the actual "reset": Windows recreates both folders empty when the
    services start, which clears a corrupt update database or catalog.
    A service can be restarted by Windows between the stop and the rename, which
    locks the folder, so each rename is tried 3 times with the services stopped
    again in between. If a folder still can't be renamed the run is marked failed.
#>
function Reset-WUFolders {
    foreach ($cacheFolder in $WUCacheFolders) {
        if (-not (Test-Path $cacheFolder)) {
            Write-Log "$cacheFolder does not exist, nothing to rename"
            continue
        }
        $backupName = "$(Split-Path $cacheFolder -Leaf).bak_$RunTimestamp"
        $isRenamed = $false
        for ($attempt = 1; $attempt -le 3 -and -not $isRenamed; $attempt++) {
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

# ---------------------------------------------------------------- dlls ------
<#
    Register-WUDlls
    Re-registers the Windows Update related DLLs with regsvr32 /s (silent).
    This repairs broken COM registrations, a cause of errors like 0x80070002 or
    "class not registered".
    Every non-zero regsvr32 exit code is logged with what it means. Code 4 is
    expected for several DLLs on Windows 10/11 and is harmless; codes 3 and 5
    point at a damaged or blocked DLL and are worth looking into (try -Repair).
    Ends with a count of each outcome.
#>
function Register-WUDlls {
    # regsvr32 exit codes (from its source: FAIL_ARGS .. FAIL_REG).
    $regsvrExitMeanings = @{
        1 = 'invalid arguments passed to regsvr32'
        2 = 'OLE/COM could not be initialised in regsvr32'
        3 = 'DLL could not be loaded: corrupt, wrong architecture, or a dependency is missing'
        4 = 'DLL has no registration entry point: it does not support regsvr32, harmless'
        5 = 'DLL registration ran but failed: usually registry access denied or a damaged DLL'
    }
    $system32Path = "$env:SystemRoot\System32"
    $dllCounts    = @{ Registered = 0; NotRegistrable = 0; Failed = 0; Missing = 0 }

    foreach ($dllName in $WUDllNames) {
        $dllPath = Join-Path $system32Path $dllName
        if (-not (Test-Path $dllPath)) {
            $dllCounts.Missing++
            continue
        }
        $regsvrProcess = Start-Process regsvr32.exe -ArgumentList "/s `"$dllPath`"" -Wait -PassThru -WindowStyle Hidden
        $exitCode = $regsvrProcess.ExitCode
        if ($exitCode -eq 0) {
            Write-Log "Registered $dllName"
            $dllCounts.Registered++
            continue
        }
        $meaning = if ($regsvrExitMeanings.ContainsKey($exitCode)) { $regsvrExitMeanings[$exitCode] } else { 'unknown regsvr32 exit code' }
        Write-Log "regsvr32 $dllName returned $exitCode - $meaning"
        if ($exitCode -eq 4) { $dllCounts.NotRegistrable++ } else { $dllCounts.Failed++ }
    }
    Write-Log ("DLLs: {0} registered, {1} not registrable (harmless), {2} failed, {3} not present" -f
        $dllCounts.Registered, $dllCounts.NotRegistrable, $dllCounts.Failed, $dllCounts.Missing)
}

# --------------------------------------------------------------- repair -----
<#
    Invoke-Repair
    Only runs with -Repair. Fixes corruption in Windows itself, which a reset
    alone can't fix (typical errors: 0x800f081f, 0x80073712).
    - DISM /RestoreHealth repairs the component store, downloading clean files
      from Windows Update. That's why it runs after the services are started.
    - sfc /scannow then repairs protected system files from the component store.
    A DISM failure marks the run as failed. SFC's exit code is only logged,
    because it doesn't reliably report problems through it.
#>
function Invoke-Repair {
    Write-Log 'Running DISM /RestoreHealth (details in C:\Windows\Logs\DISM\dism.log)...'
    $dismProcess = Start-Process "$env:SystemRoot\System32\dism.exe" -ArgumentList '/Online /Cleanup-Image /RestoreHealth' -Wait -PassThru -WindowStyle Hidden
    Write-Log "DISM exit code: $($dismProcess.ExitCode)"
    if ($dismProcess.ExitCode -ne 0) { $script:Failed = $true }

    Write-Log 'Running sfc /scannow (details in C:\Windows\Logs\CBS\CBS.log)...'
    $sfcProcess = Start-Process "$env:SystemRoot\System32\sfc.exe" -ArgumentList '/scannow' -Wait -PassThru -WindowStyle Hidden
    Write-Log "SFC exit code: $($sfcProcess.ExitCode)"
}

# -------------------------------------------------------- update via COM ----
<#
    Format-HResult
    Turns a COM HResult (a negative Int32) into the usual 0x8024xxxx form that
    can be looked up in Microsoft's Windows Update error code list.
#>
function Format-HResult {
    param([int]$HResult)
    return '0x{0:X8}' -f $HResult
}

<#
    Invoke-WUInstall
    Uses the Windows Update Agent COM API (Microsoft.Update.Session) to:
    1. Scan for software updates that are not installed and not hidden
    2. Log the details of each update found (KB, category, size, reboot behaviour)
    3. Download each update one at a time, logging the result and error code
    4. Install everything that downloaded in one batch, so Windows can order
       prerequisites such as servicing stack updates correctly
    5. Log the result, error code and reboot need of each update, then a summary
    Stops after step 2 when -SkipInstall is used.
    Any failed download or install marks the run as failed. A reboot request
    sets $script:RebootRequired.
#>
function Invoke-WUInstall {
    # OperationResultCode values returned by the WU API.
    $resultNames = @{ 0 = 'NotStarted'; 1 = 'InProgress'; 2 = 'Succeeded'; 3 = 'SucceededWithErrors'; 4 = 'Failed'; 5 = 'Aborted' }
    # InstallationBehavior.RebootBehavior values.
    $rebootBehaviorNames = @{ 0 = 'No reboot'; 1 = 'Always needs reboot'; 2 = 'May need reboot' }

    $agentVersion = (Get-Item "$env:SystemRoot\System32\wuaueng.dll").VersionInfo.ProductVersion
    Write-Log "Windows Update Agent version: $agentVersion"

    $updateSession = New-Object -ComObject Microsoft.Update.Session
    $updateSession.ClientApplicationID = 'Reset-WindowsUpdate'
    $searcher      = $updateSession.CreateUpdateSearcher()

    # ---- 1. scan
    $searchCriteria = "IsInstalled=0 and Type='Software' and IsHidden=0"
    Write-Log "Scanning for updates (criteria: $searchCriteria)..."
    $stopwatch    = [Diagnostics.Stopwatch]::StartNew()
    $searchResult = $searcher.Search($searchCriteria)
    Write-Log "Scan finished in $([int]$stopwatch.Elapsed.TotalSeconds) s, result: $($resultNames[[int]$searchResult.ResultCode])"
    Write-Log "Found $($searchResult.Updates.Count) update(s)"
    if ($searchResult.Updates.Count -eq 0) { return }

    # ---- 2. list what was found
    $updatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
    $totalSizeMB      = 0
    $updateNumber     = 0
    foreach ($update in $searchResult.Updates) {
        $updateNumber++
        $kbNumbers      = (@($update.KBArticleIDs) | ForEach-Object { "KB$_" }) -join ', '
        $categories     = (@($update.Categories) | ForEach-Object { $_.Name }) -join ', '
        $sizeMB         = [math]::Round($update.MaxDownloadSize / 1MB, 1)
        $severity       = if ($update.MsrcSeverity) { $update.MsrcSeverity } else { 'n/a' }
        $rebootBehavior = $rebootBehaviorNames[[int]$update.InstallationBehavior.RebootBehavior]
        $totalSizeMB   += $sizeMB

        Write-Log "[$updateNumber/$($searchResult.Updates.Count)] $($update.Title)"
        Write-Log "    KB: $kbNumbers | Category: $categories | Severity: $severity"
        Write-Log "    Size: $sizeMB MB | Downloaded: $($update.IsDownloaded) | Reboot: $rebootBehavior"

        if (-not $update.EulaAccepted) {
            Write-Log '    Accepting licence agreement'
            $update.AcceptEula()
        }
        [void]$updatesToInstall.Add($update)
    }
    Write-Log "Total download size (max): $totalSizeMB MB"

    if ($SkipInstall) {
        Write-Log 'SkipInstall set, not downloading or installing.'
        return
    }

    # ---- 3. download, one update at a time for per-update progress
    $downloader = $updateSession.CreateUpdateDownloader()
    $downloadedUpdates = New-Object -ComObject Microsoft.Update.UpdateColl
    $failedDownloadCount = 0
    for ($index = 0; $index -lt $updatesToInstall.Count; $index++) {
        $update      = $updatesToInstall.Item($index)
        $updateLabel = "[$($index + 1)/$($updatesToInstall.Count)] $($update.Title)"
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
            $resultCode     = [int]$downloadResult.ResultCode
            Write-Log "    $($resultNames[$resultCode]) in $([int]$stopwatch.Elapsed.TotalSeconds) s (HResult $(Format-HResult $downloadResult.HResult))"
        } catch {
            Write-Log "    Download error: $($_.Exception.Message)"
        }
        if ($update.IsDownloaded) {
            [void]$downloadedUpdates.Add($update)
        } else {
            $failedDownloadCount++
            $script:Failed = $true
        }
    }
    Write-Log "Downloads: $($downloadedUpdates.Count) ready, $failedDownloadCount failed"
    if ($downloadedUpdates.Count -eq 0) {
        Write-Log 'No updates were downloaded, nothing to install.'
        $script:Failed = $true
        return
    }

    # ---- 4. install, as one batch
    Write-Log "Installing $($downloadedUpdates.Count) update(s), this can take a while..."
    $installer = $updateSession.CreateUpdateInstaller()
    $installer.Updates = $downloadedUpdates
    $stopwatch     = [Diagnostics.Stopwatch]::StartNew()
    $installResult = $installer.Install()
    $resultCode    = [int]$installResult.ResultCode
    Write-Log "Install finished in $([int]$stopwatch.Elapsed.TotalMinutes) min, result: $($resultNames[$resultCode]) (HResult $(Format-HResult $installResult.HResult))"
    if ($resultCode -ne 2) { $script:Failed = $true }

    # ---- 5. per-update results and summary
    $installedCount     = 0
    $failedInstallCount = 0
    for ($index = 0; $index -lt $downloadedUpdates.Count; $index++) {
        $updateResult     = $installResult.GetUpdateResult($index)
        $updateResultCode = [int]$updateResult.ResultCode
        if ($updateResultCode -eq 2) { $installedCount++ } else { $failedInstallCount++ }
        Write-Log "[$($index + 1)/$($downloadedUpdates.Count)] $($resultNames[$updateResultCode]): $($downloadedUpdates.Item($index).Title)"
        Write-Log "    HResult $(Format-HResult $updateResult.HResult) | Reboot required: $($updateResult.RebootRequired)"
    }
    Write-Log "Summary: $installedCount installed, $failedInstallCount failed, $failedDownloadCount not downloaded"

    if ($installResult.RebootRequired) {
        Write-Log 'A reboot is required to finish installing updates.'
        $script:RebootRequired = $true
    }
}

# ------------------------------------------------------------------- main ---
# Order matters: checks that change nothing run first, so a device that can't
# be fixed by this script is left untouched. Any unexpected error restarts the
# services so the device isn't left without Windows Update.
New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
Write-Log "=== Windows Update reset started on $env:COMPUTERNAME as $env:USERNAME ==="

try {
    if (-not (Test-ServiceStartup)) { exit 1 }
    if ((Test-PendingReboot) -and -not $IgnorePendingReboot) {
        Write-Log 'Reboot the device first, or run with -IgnorePendingReboot.'
        exit 1
    }
    Remove-OldBackups
    if (-not (Test-DriveSpace)) { exit 1 }

    Stop-WUServices
    Reset-WUFolders
    Register-WUDlls
    Start-WUServices
    if ($Repair) { Invoke-Repair }
    Invoke-WUInstall
} catch {
    Write-Log "Unexpected error: $($_.Exception.Message)"
    $script:Failed = $true
    Start-WUServices
}

# Last line is what most MDM consoles show as the script output.
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
