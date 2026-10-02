# windowsUpdateReset

PowerShell script that resets Windows Update and then installs pending updates.

## What it does

1. Checks free space on the system drive (stops if below `-MinFreeGB`)
2. Stops `wuauserv`, `bits`, `cryptsvc`, `msiserver`
3. Renames `SoftwareDistribution` and `catroot2` to `*.bak_<timestamp>`
4. Re-registers the Windows Update DLLs (skips any not present)
5. Starts the services again
6. Scans, downloads and installs updates using the `Microsoft.Update.Session` COM object

Every step is written to a timestamped log in `C:\Windows\Logs\WUReset\`.

## Usage

Run from an elevated PowerShell prompt:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\Reset-WindowsUpdate.ps1
```

| Parameter      | Default                  | Description                          |
| -------------- | ------------------------ | ------------------------------------ |
| `-MinFreeGB`   | `10`                     | Minimum free GB on the system drive  |
| `-LogPath`     | `C:\Windows\Logs\WUReset` | Folder for the log file             |
| `-SkipInstall` | off                      | Reset and scan only, no install      |

Exit code is `0` on success and `1` on failure. Check the log for whether a reboot is required.
The old `*.bak_*` folders can be deleted once updates are working again.
