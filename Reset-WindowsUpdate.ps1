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
$Stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogFile = Join-Path $LogPath "WUReset_$Stamp.log"

# Services that hold the Windows Update cache folders open. All four must be
# stopped before the folders can be renamed.
$Services          = 'wuauserv', 'bits', 'cryptsvc', 'msiserver'
$ServiceTimeoutSec = 60

# SoftwareDistribution holds the update database and downloads,
# catroot2 holds the signature catalogs. Both are rebuilt by Windows on next use.
$WUFolders = "$env:SystemRoot\SoftwareDistribution", "$env:SystemRoot\System32\catroot2"

# DLLs used by Windows Update, BITS and cryptographic services.
# Many only exist on older Windows versions; missing ones are skipped.
$Dlls = @(
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
    Add-Content -Path $LogFile -Value $line
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
    $ok = $true
    foreach ($name in $Services) {
        $startType = (Get-Service -Name $name).StartType
        Write-Log "Service $name start type: $startType"
        if ($startType -eq 'Disabled') {
            Write-Log "Service $name is disabled (check GPO/MDM policy)."
            $ok = $false
        }
    }
    return $ok
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
    $keys = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    foreach ($key in $keys) {
        if (Test-Path $key) {
            Write-Log "Pending reboot detected: $key"
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
    $disk   = [System.IO.DriveInfo]::new($env:SystemDrive)
    $freeGB = [math]::Round($disk.AvailableFreeSpace / 1GB, 2)
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
    param([string]$Name)
    $procId = (Get-CimInstance Win32_Service -Filter "Name='$Name'").ProcessId
    if (-not $procId) { return }
    if (@(Get-CimInstance Win32_Service -Filter "ProcessId=$procId").Count -gt 1) {
        Write-Log "$Name shares process $procId with other services, not killing it"
        return
    }
    try {
        Stop-Process -Id $procId -Force -ErrorAction Stop
        Write-Log "Killed process $procId for $Name"
    } catch {
        Write-Log "Could not kill process $procId for $Name : $($_.Exception.Message)"
    }
}

<#
    Stop-WUServices
    Stops each Windows Update service and waits up to $ServiceTimeoutSec for it.
    -Force also stops services that depend on it. If a service doesn't stop in
    time, Stop-ServiceProcess is used. Services already stopped are skipped, so
    this is safe to call again (Reset-WUFolders does that between retries).
#>
function Stop-WUServices {
    foreach ($name in $Services) {
        $svc = Get-Service -Name $name
        if ($svc.Status -eq 'Stopped') { continue }
        Write-Log "Stopping service $name"
        try {
            Stop-Service -Name $name -Force -NoWait -ErrorAction Stop
            $svc.WaitForStatus('Stopped', [TimeSpan]::FromSeconds($ServiceTimeoutSec))
        } catch {
            Write-Log "$name did not stop within $ServiceTimeoutSec s: $($_.Exception.Message)"
            Stop-ServiceProcess $name
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
    foreach ($name in $Services) {
        try {
            Write-Log "Starting service $name"
            Start-Service -Name $name -ErrorAction Stop
        } catch {
            Write-Log "Could not start $name : $($_.Exception.Message)"
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
    foreach ($folder in $WUFolders) {
        $pattern = "$(Split-Path $folder -Leaf).bak_*"
        foreach ($dir in Get-ChildItem -Path (Split-Path $folder -Parent) -Directory -Filter $pattern) {
            try {
                Remove-Item -Path $dir.FullName -Recurse -Force -ErrorAction Stop
                Write-Log "Removed old backup $($dir.FullName)"
            } catch {
                Write-Log "Could not remove $($dir.FullName) : $($_.Exception.Message)"
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
    foreach ($folder in $WUFolders) {
        if (-not (Test-Path $folder)) {
            Write-Log "$folder does not exist, nothing to rename"
            continue
        }
        $newName = "$(Split-Path $folder -Leaf).bak_$Stamp"
        $renamed = $false
        for ($attempt = 1; $attempt -le 3 -and -not $renamed; $attempt++) {
            try {
                Rename-Item -Path $folder -NewName $newName -ErrorAction Stop
                Write-Log "Renamed $folder -> $newName"
                $renamed = $true
            } catch {
                Write-Log "Rename attempt $attempt of $folder failed: $($_.Exception.Message)"
                Start-Sleep -Seconds 5
                Stop-WUServices
            }
        }
        if (-not $renamed) {
            Write-Log "Could not rename $folder, reset is incomplete."
            $script:Failed = $true
        }
    }
}

# ---------------------------------------------------------------- dlls ------
<#
    Register-WUDlls
    Re-registers the Windows Update related DLLs with regsvr32 /s (silent).
    This repairs broken COM registrations, a cause of errors like 0x80070002 or
    "class not registered". Some DLLs on newer Windows don't support
    registration and return a non-zero code; that's logged and harmless.
#>
function Register-WUDlls {
    $sys32 = "$env:SystemRoot\System32"
    foreach ($dll in $Dlls) {
        $path = Join-Path $sys32 $dll
        if (-not (Test-Path $path)) { continue }
        $proc = Start-Process regsvr32.exe -ArgumentList "/s `"$path`"" -Wait -PassThru -WindowStyle Hidden
        if ($proc.ExitCode -eq 0) {
            Write-Log "Registered $dll"
        } else {
            Write-Log "regsvr32 $dll returned $($proc.ExitCode)"
        }
    }
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
    $dism = Start-Process "$env:SystemRoot\System32\dism.exe" -ArgumentList '/Online /Cleanup-Image /RestoreHealth' -Wait -PassThru -WindowStyle Hidden
    Write-Log "DISM exit code: $($dism.ExitCode)"
    if ($dism.ExitCode -ne 0) { $script:Failed = $true }

    Write-Log 'Running sfc /scannow (details in C:\Windows\Logs\CBS\CBS.log)...'
    $sfc = Start-Process "$env:SystemRoot\System32\sfc.exe" -ArgumentList '/scannow' -Wait -PassThru -WindowStyle Hidden
    Write-Log "SFC exit code: $($sfc.ExitCode)"
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
    Get-UpdateSource
    Describes where this device gets its updates from, for the log.
    If WSUS is configured by policy, scans go to that server instead of
    Microsoft, which matters when a scan fails or finds nothing.
#>
function Get-UpdateSource {
    $policy = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    $wsus   = (Get-ItemProperty -Path $policy -Name WUServer -ErrorAction SilentlyContinue).WUServer
    $useIt  = (Get-ItemProperty -Path "$policy\AU" -Name UseWUServer -ErrorAction SilentlyContinue).UseWUServer
    if ($wsus -and $useIt -eq 1) { return "WSUS ($wsus)" }
    return 'Windows Update / Windows Update for Business'
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
    $resultText = @{ 0 = 'NotStarted'; 1 = 'InProgress'; 2 = 'Succeeded'; 3 = 'SucceededWithErrors'; 4 = 'Failed'; 5 = 'Aborted' }
    # InstallationBehavior.RebootBehavior values.
    $rebootText = @{ 0 = 'No reboot'; 1 = 'Always needs reboot'; 2 = 'May need reboot' }

    $agentVersion = (Get-Item "$env:SystemRoot\System32\wuaueng.dll").VersionInfo.ProductVersion
    Write-Log "Windows Update Agent version: $agentVersion"
    Write-Log "Update source: $(Get-UpdateSource)"

    $session  = New-Object -ComObject Microsoft.Update.Session
    $session.ClientApplicationID = 'Reset-WindowsUpdate'
    $searcher = $session.CreateUpdateSearcher()

    # ---- 1. scan
    $criteria = "IsInstalled=0 and Type='Software' and IsHidden=0"
    Write-Log "Scanning for updates (criteria: $criteria)..."
    $timer  = [Diagnostics.Stopwatch]::StartNew()
    $search = $searcher.Search($criteria)
    Write-Log "Scan finished in $([int]$timer.Elapsed.TotalSeconds) s, result: $($resultText[[int]$search.ResultCode])"
    Write-Log "Found $($search.Updates.Count) update(s)"
    if ($search.Updates.Count -eq 0) { return }

    # ---- 2. list what was found
    $toInstall = New-Object -ComObject Microsoft.Update.UpdateColl
    $totalMB   = 0
    $n = 0
    foreach ($update in $search.Updates) {
        $n++
        $kb       = (@($update.KBArticleIDs) | ForEach-Object { "KB$_" }) -join ', '
        $category = (@($update.Categories) | ForEach-Object { $_.Name }) -join ', '
        $sizeMB   = [math]::Round($update.MaxDownloadSize / 1MB, 1)
        $severity = if ($update.MsrcSeverity) { $update.MsrcSeverity } else { 'n/a' }
        $reboot   = $rebootText[[int]$update.InstallationBehavior.RebootBehavior]
        $totalMB += $sizeMB

        Write-Log "[$n/$($search.Updates.Count)] $($update.Title)"
        Write-Log "    KB: $kb | Category: $category | Severity: $severity"
        Write-Log "    Size: $sizeMB MB | Downloaded: $($update.IsDownloaded) | Reboot: $reboot"

        if (-not $update.EulaAccepted) {
            Write-Log '    Accepting licence agreement'
            $update.AcceptEula()
        }
        [void]$toInstall.Add($update)
    }
    Write-Log "Total download size (max): $totalMB MB"

    if ($SkipInstall) {
        Write-Log 'SkipInstall set, not downloading or installing.'
        return
    }

    # ---- 3. download, one update at a time for per-update progress
    $downloader = $session.CreateUpdateDownloader()
    $downloaded = New-Object -ComObject Microsoft.Update.UpdateColl
    $failedDownloads = 0
    for ($i = 0; $i -lt $toInstall.Count; $i++) {
        $update = $toInstall.Item($i)
        $label  = "[$($i + 1)/$($toInstall.Count)] $($update.Title)"
        if ($update.IsDownloaded) {
            Write-Log "Already downloaded: $label"
            [void]$downloaded.Add($update)
            continue
        }
        Write-Log "Downloading: $label"
        $single = New-Object -ComObject Microsoft.Update.UpdateColl
        [void]$single.Add($update)
        $downloader.Updates = $single
        $timer = [Diagnostics.Stopwatch]::StartNew()
        try {
            $result = $downloader.Download()
            $code   = [int]$result.ResultCode
            Write-Log "    $($resultText[$code]) in $([int]$timer.Elapsed.TotalSeconds) s (HResult $(Format-HResult $result.HResult))"
        } catch {
            Write-Log "    Download error: $($_.Exception.Message)"
        }
        if ($update.IsDownloaded) {
            [void]$downloaded.Add($update)
        } else {
            $failedDownloads++
            $script:Failed = $true
        }
    }
    Write-Log "Downloads: $($downloaded.Count) ready, $failedDownloads failed"
    if ($downloaded.Count -eq 0) {
        Write-Log 'No updates were downloaded, nothing to install.'
        $script:Failed = $true
        return
    }

    # ---- 4. install, as one batch
    Write-Log "Installing $($downloaded.Count) update(s), this can take a while..."
    $installer = $session.CreateUpdateInstaller()
    $installer.Updates = $downloaded
    $timer   = [Diagnostics.Stopwatch]::StartNew()
    $install = $installer.Install()
    $code    = [int]$install.ResultCode
    Write-Log "Install finished in $([int]$timer.Elapsed.TotalMinutes) min, result: $($resultText[$code]) (HResult $(Format-HResult $install.HResult))"
    if ($code -ne 2) { $script:Failed = $true }

    # ---- 5. per-update results and summary
    $installed      = 0
    $failedInstalls = 0
    for ($i = 0; $i -lt $downloaded.Count; $i++) {
        $r    = $install.GetUpdateResult($i)
        $rc   = [int]$r.ResultCode
        if ($rc -eq 2) { $installed++ } else { $failedInstalls++ }
        Write-Log "[$($i + 1)/$($downloaded.Count)] $($resultText[$rc]): $($downloaded.Item($i).Title)"
        Write-Log "    HResult $(Format-HResult $r.HResult) | Reboot required: $($r.RebootRequired)"
    }
    Write-Log "Summary: $installed installed, $failedInstalls failed, $failedDownloads not downloaded"

    if ($install.RebootRequired) {
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
