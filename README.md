# windowsUpdateReset

PowerShell script that resets Windows Update and then installs pending updates.
Built to run unattended from MDM as SYSTEM.

## What it does

1. Stops if a WU service is disabled or a reboot is already pending
2. Deletes `*.bak_*` folders left by previous runs, then checks free space on the system drive
3. Stops `wuauserv`, `bits`, `cryptsvc`, `msiserver` (60 s timeout, then kills the process if it isn't a shared svchost)
4. Renames `SoftwareDistribution` and `catroot2` to `*.bak_<timestamp>` (3 attempts)
5. Re-registers the Windows Update DLLs (skips any not present)
6. Starts the services again
7. Scans, downloads and installs updates using the `Microsoft.Update.Session` COM object

Every step is written to a timestamped log in `C:\Windows\Logs\WUReset\`.
The last line of output is a `RESULT:` summary, which MDM consoles show as the script output.

## Usage

Run from an elevated PowerShell prompt:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\Reset-WindowsUpdate.ps1
```

| Parameter              | Default                   | Description                                      |
| ---------------------- | ------------------------- | ------------------------------------------------ |
| `-MinFreeGB`           | `10`                      | Minimum free GB on the system drive              |
| `-LogPath`             | `C:\Windows\Logs\WUReset` | Folder for the log file                          |
| `-SkipInstall`         | off                       | Reset and scan only, no install                  |
| `-IgnorePendingReboot` | off                       | Run even if a reboot is already pending          |
| `-RebootExitCode`      | `0`                       | Exit code when a reboot is required              |

## Exit codes

| Code               | Meaning                                    |
| ------------------ | ------------------------------------------ |
| `0`                | Success                                    |
| `1`                | Failed (see log)                           |
| `-RebootExitCode`  | Success, reboot required (default `0`)     |

`-RebootExitCode` defaults to `0` because MDM platform scripts treat any non-zero code as a
failure and retry, which would reset Windows Update again. Use `3010` when deploying as a
Win32 app so it's reported as a soft reboot.

## MDM notes

- Run as SYSTEM (Intune: "Run this script using the logged on credentials" = **No**).
- 32-bit PowerShell is detected and the script relaunches itself in 64-bit PowerShell.
- Downloading and installing can take a long time; MDM script timeouts may cut it off on slow devices.
- Don't run it over remote PowerShell (`Invoke-Command` / `Enter-PSSession`): the update COM
  object refuses to download or install from a remote session.
