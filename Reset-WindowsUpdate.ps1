#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Resets Windows Update components, then scans, downloads and installs updates.

.DESCRIPTION
    1. Checks free space on the system drive
    2. Stops the Windows Update related services
    3. Renames the SoftwareDistribution and catroot2 folders
    4. Re-registers the Windows Update DLLs
    5. Starts the services again
    6. Scans, downloads and installs updates through the Microsoft.Update.Session COM object

.PARAMETER MinFreeGB
    Minimum free space (GB) required on the system drive. Default: 10.

.PARAMETER LogPath
    Folder for the log file. Default: C:\Windows\Logs\WUReset.

.PARAMETER SkipInstall
    Only reset and scan; do not download or install updates.

.EXAMPLE
    .\Reset-WindowsUpdate.ps1
    .\Reset-WindowsUpdate.ps1 -MinFreeGB 20 -SkipInstall
#>
[CmdletBinding()]
param(
    [int]$MinFreeGB = 10,
    [string]$LogPath = "$env:SystemRoot\Logs\WUReset",
    [switch]$SkipInstall
)

$ErrorActionPreference = 'Stop'
$Stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogFile = Join-Path $LogPath "WUReset_$Stamp.log"

$Services = 'wuauserv', 'bits', 'cryptsvc', 'msiserver'

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

# ---------------------------------------------------------------- logging ---
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -Path $LogFile -Value $line
    switch ($Level) {
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        default { Write-Host $line }
    }
}

# ------------------------------------------------------------- disk space ---
function Test-DriveSpace {
    $disk   = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$env:SystemDrive'"
    $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
    Write-Log "Free space on $env:SystemDrive : $freeGB GB (minimum $MinFreeGB GB)"
    if ($freeGB -lt $MinFreeGB) {
        Write-Log "Not enough free space on $env:SystemDrive." 'ERROR'
        return $false
    }
    return $true
}

# --------------------------------------------------------------- services ---
function Stop-WUServices {
    foreach ($name in $Services) {
        try {
            Write-Log "Stopping service $name"
            Stop-Service -Name $name -Force -ErrorAction Stop
        } catch {
            Write-Log "Could not stop $name : $($_.Exception.Message)" 'WARN'
        }
    }
}

function Start-WUServices {
    foreach ($name in $Services) {
        try {
            Write-Log "Starting service $name"
            Start-Service -Name $name -ErrorAction Stop
        } catch {
            Write-Log "Could not start $name : $($_.Exception.Message)" 'WARN'
        }
    }
}

# ------------------------------------------------------- cache folders ------
function Reset-WUFolders {
    $folders = "$env:SystemRoot\SoftwareDistribution", "$env:SystemRoot\System32\catroot2"
    foreach ($folder in $folders) {
        if (Test-Path $folder) {
            $newName = "$(Split-Path $folder -Leaf).bak_$Stamp"
            try {
                Write-Log "Renaming $folder -> $newName"
                Rename-Item -Path $folder -NewName $newName -ErrorAction Stop
            } catch {
                Write-Log "Could not rename $folder : $($_.Exception.Message)" 'WARN'
            }
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
            Write-Log "regsvr32 $dll returned $($proc.ExitCode)" 'WARN'
        }
    }
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
    Write-Log "Download result: $($resultText[[int]$download.ResultCode])"

    $downloaded = New-Object -ComObject Microsoft.Update.UpdateColl
    foreach ($update in $toInstall) {
        if ($update.IsDownloaded) { [void]$downloaded.Add($update) }
    }
    if ($downloaded.Count -eq 0) {
        Write-Log 'No updates were downloaded, nothing to install.' 'WARN'
        return
    }

    Write-Log "Installing $($downloaded.Count) update(s)..."
    $installer = $session.CreateUpdateInstaller()
    $installer.Updates = $downloaded
    $install = $installer.Install()
    Write-Log "Install result: $($resultText[[int]$install.ResultCode])"

    for ($i = 0; $i -lt $downloaded.Count; $i++) {
        $code = [int]$install.GetUpdateResult($i).ResultCode
        Write-Log "  $($resultText[$code]): $($downloaded.Item($i).Title)"
    }

    if ($install.RebootRequired) {
        Write-Log 'A reboot is required to finish installing updates.' 'WARN'
    }
}

# ------------------------------------------------------------------- main ---
New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
Write-Log "=== Windows Update reset started on $env:COMPUTERNAME ==="

try {
    if (-not (Test-DriveSpace)) { exit 1 }

    Stop-WUServices
    Reset-WUFolders
    Register-WUDlls
    Start-WUServices
    Invoke-WUInstall

    Write-Log '=== Windows Update reset finished ==='
    exit 0
} catch {
    Write-Log "Unexpected error: $($_.Exception.Message)" 'ERROR'
    Start-WUServices
    exit 1
}
