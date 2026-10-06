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

# Log files kept in $LogPath, including this run's.
$LogsToKeep = 10

# Services that lock the WU cache folders.
$WUServiceNames        = 'wuauserv', 'bits', 'cryptsvc', 'msiserver'
$ServiceTimeoutSeconds = 60

# Update database/downloads and signature catalogs. Windows rebuilds both.
$WUCacheFolders = "$env:SystemRoot\SoftwareDistribution", "$env:SystemRoot\System32\catroot2"

# WU-related DLLs. Missing ones are skipped.
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

# These two decide the exit code.
$script:Failed         = $false
$script:RebootRequired = $false

# ---------------------------------------------------------------- logging ---
# Timestamped line to the log file and console.
function Write-Log {
    param([string]$Message)
    $line = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Add-Content -Path $LogFilePath -Value $line
    Write-Host $line
}

# Deletes all but the newest $LogsToKeep log files.
function Remove-OldLogs {
    Get-ChildItem -Path $LogPath -Filter 'WUReset_*.log' |
        Sort-Object Name -Descending |
        Select-Object -Skip $LogsToKeep |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

# Writes the RESULT line (MDM consoles show the last line) and exits with the matching code.
function Exit-WithResult {
    param([switch]$Failure)
    if ($Failure) { $script:Failed = $true }
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
}

# --------------------------------------------------------- pre-checks -------
# False if a WU service is disabled (usually by policy).
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

# True if Windows is already waiting for a reboot.
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

# False if the system drive is below -MinFreeGB. Uses DriveInfo so broken WMI can't block it.
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
# True when every WU service is in the given state; logs any that aren't.
function Test-ServicesInState {
    param([string]$State)
    $allInState = $true
    foreach ($serviceName in $WUServiceNames) {
        $status = (Get-Service -Name $serviceName).Status
        if ($status -ne $State) {
            Write-Log "Service $serviceName is $status, expected $State"
            $allInState = $false
        }
    }
    return $allInState
}

# Mike F Robbins' method: ends the process of a service stuck in Stop Pending, unless it's a shared svchost.
function Stop-HungService {
    param([string]$ServiceName)
    $hungService = Get-CimInstance Win32_Service -Filter "Name='$ServiceName' AND State='Stop Pending'"
    if (-not $hungService) {
        Write-Log "$ServiceName is not in Stop Pending, leaving it alone"
        return
    }
    $serviceProcessId = $hungService.ProcessId
    if (@(Get-CimInstance Win32_Service -Filter "ProcessId=$serviceProcessId").Count -gt 1) {
        Write-Log "$ServiceName shares process $serviceProcessId with other services, not ending it"
        return
    }
    try {
        Stop-Process -Id $serviceProcessId -Force -ErrorAction Stop
        Write-Log "Ended process $serviceProcessId of hung service $ServiceName"
        (Get-Service -Name $ServiceName).WaitForStatus('Stopped', [TimeSpan]::FromSeconds(15))
    } catch {
        Write-Log "Could not end process $serviceProcessId of $ServiceName : $($_.Exception.Message)"
    }
}

# Stops the WU services and waits for each. True only when all are stopped. Safe to re-run.
function Stop-WUServices {
    foreach ($serviceName in $WUServiceNames) {
        $service = Get-Service -Name $serviceName
        if ($service.Status -eq 'Stopped') { continue }
        Write-Log "Stopping service $serviceName"
        try {
            Stop-Service -Name $serviceName -Force -NoWait -ErrorAction Stop
            $service.WaitForStatus('Stopped', [TimeSpan]::FromSeconds($ServiceTimeoutSeconds))
            Write-Log "Service $serviceName stopped"
        } catch {
            Write-Log "$serviceName did not stop within $ServiceTimeoutSeconds s: $($_.Exception.Message)"
            Stop-HungService $serviceName
        }
    }
    return Test-ServicesInState -State 'Stopped'
}

# Starts the WU services and waits for each. True only when all are running.
function Start-WUServices {
    foreach ($serviceName in $WUServiceNames) {
        $service = Get-Service -Name $serviceName
        if ($service.Status -eq 'Running') { continue }
        Write-Log "Starting service $serviceName"
        try {
            if ($service.Status -ne 'StartPending') { $service.Start() }
            $service.WaitForStatus('Running', [TimeSpan]::FromSeconds($ServiceTimeoutSeconds))
            Write-Log "Service $serviceName running"
        } catch {
            Write-Log "$serviceName did not start within $ServiceTimeoutSeconds s: $($_.Exception.Message)"
        }
    }
    return Test-ServicesInState -State 'Running'
}

# ------------------------------------------------------- cache folders ------
# Deletes .bak_* folders from earlier runs.
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

# Renames the cache folders (the actual reset), retrying if a service relocks them.
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
                $null = Stop-WUServices
            }
        }
        if (-not $isRenamed) {
            Write-Log "Could not rename $cacheFolder, reset is incomplete."
            $script:Failed = $true
        }
    }
}

# ---------------------------------------------------------------- dlls ------
# Re-registers the WU DLLs and explains any regsvr32 failure. Code 4 is harmless.
function Register-WUDlls {
    # regsvr32 exit codes.
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
# DISM then SFC (-Repair only). Runs after the reset because DISM needs Windows Update.
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
# HResult as 0x8024xxxx, for looking up WU error codes.
function Format-HResult {
    param([int]$HResult)
    return '0x{0:X8}' -f $HResult
}

# Scan, download (one at a time) and install (one batch) via the WU COM API.
function Invoke-WUInstall {
    # WU API result codes.
    $resultNames = @{ 0 = 'NotStarted'; 1 = 'InProgress'; 2 = 'Succeeded'; 3 = 'SucceededWithErrors'; 4 = 'Failed'; 5 = 'Aborted' }
    # Update reboot behaviour.
    $rebootBehaviorNames = @{ 0 = 'No reboot'; 1 = 'Always needs reboot'; 2 = 'May need reboot' }

    $agentVersion = (Get-Item "$env:SystemRoot\System32\wuaueng.dll" -ErrorAction SilentlyContinue).VersionInfo.ProductVersion
    if (-not $agentVersion) { $agentVersion = 'unknown' }
    Write-Log "Windows Update Agent version: $agentVersion"

    $updateSession = New-Object -ComObject Microsoft.Update.Session
    $updateSession.ClientApplicationID = 'Reset-WindowsUpdate'
    $searcher      = $updateSession.CreateUpdateSearcher()

    # ---- scan
    $searchCriteria = "IsInstalled=0 and Type='Software' and IsHidden=0"
    Write-Log "Scanning for updates (criteria: $searchCriteria)..."
    $stopwatch    = [Diagnostics.Stopwatch]::StartNew()
    $searchResult = $searcher.Search($searchCriteria)
    Write-Log "Scan finished in $([int]$stopwatch.Elapsed.TotalSeconds) s, result: $($resultNames[[int]$searchResult.ResultCode])"
    if ([int]$searchResult.ResultCode -ne 2) {
        Write-Log 'Scan did not succeed, not installing anything.'
        $script:Failed = $true
        return
    }
    Write-Log "Found $($searchResult.Updates.Count) update(s)"
    if ($searchResult.Updates.Count -eq 0) { return }

    # ---- list
    $updatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
    $totalSizeMB      = 0
    $updateNumber     = 0
    $skippedCount     = 0
    foreach ($update in $searchResult.Updates) {
        $updateNumber++
        $kbNumbers      = (@($update.KBArticleIDs) | ForEach-Object { "KB$_" }) -join ', '
        $categories     = (@($update.Categories) | ForEach-Object { $_.Name }) -join ', '
        $sizeMB         = [math]::Round($update.MaxDownloadSize / 1MB, 1)
        $severity       = if ($update.MsrcSeverity) { $update.MsrcSeverity } else { 'n/a' }
        $rebootBehavior = $rebootBehaviorNames[[int]$update.InstallationBehavior.RebootBehavior]

        Write-Log "[$updateNumber/$($searchResult.Updates.Count)] $($update.Title)"
        Write-Log "    KB: $kbNumbers | Category: $categories | Severity: $severity"
        Write-Log "    Size: $sizeMB MB | Downloaded: $($update.IsDownloaded) | Reboot: $rebootBehavior"

        # Nobody can answer a prompt when running unattended as SYSTEM.
        if ($update.InstallationBehavior.CanRequestUserInput) {
            Write-Log '    Skipped: this update may ask for user input'
            $skippedCount++
            continue
        }
        if (-not $update.EulaAccepted) {
            Write-Log '    Accepting licence agreement'
            $update.AcceptEula()
        }
        [void]$updatesToInstall.Add($update)
        $totalSizeMB += $sizeMB
    }
    Write-Log "To install: $($updatesToInstall.Count), skipped: $skippedCount, total download size (max): $totalSizeMB MB"
    if ($updatesToInstall.Count -eq 0) { return }

    if ($SkipInstall) {
        Write-Log 'SkipInstall set, not downloading or installing.'
        return
    }

    # ---- download, one at a time
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

    # ---- install, one batch
    Write-Log "Installing $($downloadedUpdates.Count) update(s), this can take a while..."
    $installer = $updateSession.CreateUpdateInstaller()
    $installer.Updates = $downloadedUpdates
    $stopwatch     = [Diagnostics.Stopwatch]::StartNew()
    $installResult = $installer.Install()
    $resultCode    = [int]$installResult.ResultCode
    Write-Log "Install finished in $([int]$stopwatch.Elapsed.TotalMinutes) min, result: $($resultNames[$resultCode]) (HResult $(Format-HResult $installResult.HResult))"
    if ($resultCode -ne 2) { $script:Failed = $true }

    # ---- results
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
# Checks first, so a device that can't be fixed is left untouched.
New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
Write-Log "=== Windows Update reset started on $env:COMPUTERNAME as $env:USERNAME ==="

try {
    Remove-OldLogs
    if (-not (Test-ServiceStartup)) { Exit-WithResult -Failure }
    if ((Test-PendingReboot) -and -not $IgnorePendingReboot) {
        Write-Log 'Reboot the device first, or run with -IgnorePendingReboot.'
        Exit-WithResult -Failure
    }
    Remove-OldBackups
    if (-not (Test-DriveSpace)) { Exit-WithResult -Failure }

    # Only touch the cache folders once every service is confirmed stopped.
    # Retried because Windows can restart a service right after it stops.
    $allStopped = $false
    for ($attempt = 1; $attempt -le 3 -and -not $allStopped; $attempt++) {
        $allStopped = Stop-WUServices
        if (-not $allStopped) { Start-Sleep -Seconds 5 }
    }
    if (-not $allStopped) {
        Write-Log 'Not all services stopped, skipping the reset.'
        $null = Start-WUServices
        Exit-WithResult -Failure
    }
    Reset-WUFolders
    Register-WUDlls

    # Repair and updates need every service confirmed running.
    if (-not (Start-WUServices)) {
        Write-Log 'Not all services started, skipping repair and updates.'
        Exit-WithResult -Failure
    }
    if ($Repair) { Invoke-Repair }
    Invoke-WUInstall
} catch {
    Write-Log "Unexpected error: $($_.Exception.Message)"
    $script:Failed = $true
    $null = Start-WUServices
}

Exit-WithResult
