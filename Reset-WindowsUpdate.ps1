#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Resets Windows Update components, then scans, downloads and installs updates.

.DESCRIPTION
    Built to run unattended from MDM (as SYSTEM).

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

# MDM agents often start 32-bit PowerShell, where System32 is redirected to SysWOW64.
# Relaunch in 64-bit PowerShell so the right folders and DLLs are touched.
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
    foreach ($p in $PSBoundParameters.GetEnumerator()) {
        if ($p.Value -is [switch]) {
            if ($p.Value) { $argList += "-$($p.Key)" }
        } else {
            $argList += "-$($p.Key)", "$($p.Value)"
        }
    }
    & "$env:SystemRoot\SysNative\WindowsPowerShell\v1.0\powershell.exe" @argList
    exit $LASTEXITCODE
}

$ErrorActionPreference = 'Stop'
$Stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogFile = Join-Path $LogPath "WUReset_$Stamp.log"

$Services          = 'wuauserv', 'bits', 'cryptsvc', 'msiserver'
$ServiceTimeoutSec = 60
$WUFolders         = "$env:SystemRoot\SoftwareDistribution", "$env:SystemRoot\System32\catroot2"

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

$script:Failed         = $false
$script:RebootRequired = $false

# ---------------------------------------------------------------- logging ---
function Write-Log {
    param([string]$Message)
    $line = '{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Add-Content -Path $LogFile -Value $line
    Write-Host $line
}

# --------------------------------------------------------- pre-checks -------
function Test-ServiceStartup {
    $ok = $true
    foreach ($name in $Services) {
        if ((Get-Service -Name $name).StartType -eq 'Disabled') {
            Write-Log "Service $name is disabled (check GPO/MDM policy)."
            $ok = $false
        }
    }
    return $ok
}

function Test-PendingReboot {
    $keys = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    foreach ($key in $keys) {
        if (Test-Path $key) {
            Write-Log "Pending reboot detected: $key"
            return $true
        }
    }
    return $false
}

function Test-DriveSpace {
    $disk   = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$env:SystemDrive'"
    $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
    Write-Log "Free space on $env:SystemDrive : $freeGB GB (minimum $MinFreeGB GB)"
    if ($freeGB -lt $MinFreeGB) {
        Write-Log "Not enough free space on $env:SystemDrive."
        return $false
    }
    return $true
}

# --------------------------------------------------------------- services ---
function Stop-ServiceProcess {
    param([string]$Name)
    $procId = (Get-CimInstance Win32_Service -Filter "Name='$Name'").ProcessId
    if (-not $procId) { return }
    # Never kill a shared svchost, it would take other services down with it.
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

function Reset-WUFolders {
    foreach ($folder in $WUFolders) {
        if (-not (Test-Path $folder)) { continue }
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
                Stop-WUServices   # something may have restarted a service
            }
        }
        if (-not $renamed) {
            Write-Log "Could not rename $folder, reset is incomplete."
            $script:Failed = $true
        }
    }
}

# ---------------------------------------------------------------- dlls ------
function Register-WUDlls {
    $sys32 = "$env:SystemRoot\System32"
    foreach ($dll in $Dlls) {
        $path = Join-Path $sys32 $dll
        if (-not (Test-Path $path)) { continue }   # many are absent on newer Windows
        $proc = Start-Process regsvr32.exe -ArgumentList "/s `"$path`"" -Wait -PassThru -WindowStyle Hidden
        if ($proc.ExitCode -eq 0) {
            Write-Log "Registered $dll"
        } else {
            Write-Log "regsvr32 $dll returned $($proc.ExitCode)"
        }
    }
}

# --------------------------------------------------------------- repair -----
function Invoke-Repair {
    # Runs after the reset so DISM can pull repair files from Windows Update.
    Write-Log 'Running DISM /RestoreHealth (details in C:\Windows\Logs\DISM\dism.log)...'
    $dism = Start-Process "$env:SystemRoot\System32\dism.exe" -ArgumentList '/Online /Cleanup-Image /RestoreHealth' -Wait -PassThru -WindowStyle Hidden
    Write-Log "DISM exit code: $($dism.ExitCode)"
    if ($dism.ExitCode -ne 0) { $script:Failed = $true }

    Write-Log 'Running sfc /scannow (details in C:\Windows\Logs\CBS\CBS.log)...'
    $sfc = Start-Process "$env:SystemRoot\System32\sfc.exe" -ArgumentList '/scannow' -Wait -PassThru -WindowStyle Hidden
    Write-Log "SFC exit code: $($sfc.ExitCode)"
}

# -------------------------------------------------------- update via COM ----
function Invoke-WUInstall {
    $resultText = @{ 0 = 'NotStarted'; 1 = 'InProgress'; 2 = 'Succeeded'; 3 = 'SucceededWithErrors'; 4 = 'Failed'; 5 = 'Aborted' }

    $session  = New-Object -ComObject Microsoft.Update.Session
    $searcher = $session.CreateUpdateSearcher()

    Write-Log 'Scanning for updates...'
    $search = $searcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
    Write-Log "Found $($search.Updates.Count) update(s)"
    if ($search.Updates.Count -eq 0) { return }

    $toInstall = New-Object -ComObject Microsoft.Update.UpdateColl
    foreach ($update in $search.Updates) {
        Write-Log "  - $($update.Title)"
        if (-not $update.EulaAccepted) { $update.AcceptEula() }
        [void]$toInstall.Add($update)
    }

    if ($SkipInstall) {
        Write-Log 'SkipInstall set, not downloading or installing.'
        return
    }

    Write-Log 'Downloading updates...'
    $downloader = $session.CreateUpdateDownloader()
    $downloader.Updates = $toInstall
    $download = $downloader.Download()
    $code = [int]$download.ResultCode
    Write-Log "Download result: $($resultText[$code])"
    if ($code -ne 2) { $script:Failed = $true }

    $downloaded = New-Object -ComObject Microsoft.Update.UpdateColl
    foreach ($update in $toInstall) {
        if ($update.IsDownloaded) { [void]$downloaded.Add($update) }
    }
    if ($downloaded.Count -eq 0) {
        Write-Log 'No updates were downloaded, nothing to install.'
        $script:Failed = $true
        return
    }

    Write-Log "Installing $($downloaded.Count) update(s)..."
    $installer = $session.CreateUpdateInstaller()
    $installer.Updates = $downloaded
    $install = $installer.Install()
    $code = [int]$install.ResultCode
    Write-Log "Install result: $($resultText[$code])"
    if ($code -ne 2) { $script:Failed = $true }

    for ($i = 0; $i -lt $downloaded.Count; $i++) {
        $code = [int]$install.GetUpdateResult($i).ResultCode
        Write-Log "  $($resultText[$code]): $($downloaded.Item($i).Title)"
    }

    if ($install.RebootRequired) {
        Write-Log 'A reboot is required to finish installing updates.'
        $script:RebootRequired = $true
    }
}

# ------------------------------------------------------------------- main ---
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
