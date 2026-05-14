# winboot-rescue

> **Professional Windows 10/11 boot recovery utility for WinRE/WinPE.**  
> Run it from a bootable USB stick after power loss, abrupt shutdown, BCD corruption, or missing EFI partition.

---

## Features

- **Environment detection** — identifies WinRE/WinPE vs standard Windows; locates the USB drive automatically
- **Full disk & partition inventory** — enumerates all disks, volumes, GPT vs MBR, EFI partition detection
- **Multi-installation support** — scans all drive letters for valid Windows installs; never assumes `C:`
- **Smart boot mode detection** — UEFI vs Legacy BIOS via firmware API + bcdedit cross-check
- **BCD backup before any modification** — always exports a timestamped `.bak` before touching boot data
- **UEFI/GPT repair** — mounts EFI partition (assigns temp letter if needed), rebuilds via `bcdboot /f UEFI`
- **BIOS/MBR repair** — `bootrec /fixmbr` → `/fixboot` → `/scanos` → `/rebuildbcd` + `bcdboot /f BIOS` fallback
- **Offline CHKDSK** — scan and fix file system errors on the offline Windows volume
- **Offline SFC** — `sfc /scannow /offbootdir /offwindir` with CBS.log collection
- **Structured logging** — timestamped main log, raw command outputs, full PowerShell transcript
- **Final verdict** — clear outcome summary with actionable next steps

---

## Files

| File | Description |
|---|---|
| `boot-repair.ps1` | Main recovery script (~60 KB, 1400+ lines, 20 functions) |
| `boot-repair.cmd` | Minimal launcher — admin check + `ExecutionPolicy Bypass` |

---

## Interactive Menu

```
  [1] Collect diagnostics only
  [2] Quick repair (auto-detect UEFI or BIOS)
  [3] Full repair (diagnostics + repair + CHKDSK)
  [4] Rebuild EFI boot files (UEFI/GPT systems)
  [5] BIOS/MBR boot repair (Legacy BIOS systems)
  [6] Run CHKDSK on Windows volume
  [7] Run offline SFC (System File Checker)
  [8] Export full diagnostic report
  [9] Re-select Windows installation
  [0] Exit
```

---

## Quick Start

### 1. Copy to USB

Copy both files to the root of your WinRE/WinPE USB stick:

```
E:\boot-repair.ps1
E:\boot-repair.cmd
```

Logs are created automatically under:

```
E:\Logs\BootRepair\
    boot-repair_<timestamp>.log
    transcript_<timestamp>.txt
    raw_<timestamp>\          <- individual command outputs
    bcd_backup_<timestamp>.bak
    boot-repair-report_<timestamp>.txt
```

### 2. Boot from USB

Boot the target PC from your WinRE/WinPE USB stick.

### 3. Run the tool

From the Command Prompt in WinRE/WinPE:

```bat
E:\boot-repair.cmd
```

Replace `E:` with the actual USB drive letter assigned in the recovery environment.

Alternatively, run PowerShell directly:

```powershell
powershell.exe -ExecutionPolicy Bypass -File E:\boot-repair.ps1
```

### 4. Recommended first step

Always start with **[1] Collect diagnostics only** — it is safe, makes no changes, and shows you exactly what the tool detected (firmware mode, partition style, EFI drive, Windows installations found).

Then choose the appropriate repair:

| Scenario | Option |
|---|---|
| UEFI/GPT system, EFI boot missing | `[4]` Rebuild EFI boot files |
| BIOS/MBR system, BCD/MBR corrupt | `[5]` BIOS/MBR boot repair |
| Unknown / let the tool decide | `[2]` Quick repair |
| Full recovery including disk check | `[3]` Full repair |

---

## Architecture

```
boot-repair.ps1
├── Section 0:  Global state & script-level variables
├── Section 1:  Logging (Write-Log, Invoke-LoggedCommand, Save-RawOutput)
├── Section 2:  Environment detection (WinRE/WinPE, USB drive)
├── Section 3:  Disk & partition inventory (Get-Disk, diskpart fallback)
│   └──         EFI partition detection + temp letter assignment
├── Section 4:  Windows installation scanning (all drives, no C: assumption)
├── Section 5:  Diagnostic collection (bcdedit, mountvol, reagentc, CIM)
├── Section 6:  BCD backup (bcdedit /export + file copy fallback)
├── Section 7:  UEFI/GPT repair (bcdboot /f UEFI + /f ALL fallback)
├── Section 8:  BIOS/MBR repair (bootrec sequence + bcdboot fallback)
├── Section 9:  CHKDSK (offline, user-confirmed, /F and /R options)
├── Section 10: Offline SFC (/offbootdir + /offwindir, CBS.log collection)
├── Section 11: Utilities (Get-UserConfirmation, Export-Report, Get-FinalVerdict)
├── Section 12: Interactive menu (Show-MainMenu)
└── Section 13: Main entry point
```

---

## Verdicts

At the end of every session the tool outputs one of:

| Verdict | Meaning |
|---|---|
| `BOOT FILES REBUILT SUCCESSFULLY` | Repair completed — restart and verify |
| `DIAGNOSTICS ONLY` | No repair was performed this session |
| `EFI PARTITION MISSING OR INACCESSIBLE` | UEFI system needs an accessible EFI partition |
| `NO VALID WINDOWS INSTALLATION DETECTED` | Possible hardware failure or severely corrupt disk |
| `REPAIR ATTEMPTED BUT UNCERTAIN` | Review logs; consider CHKDSK + SFC |

---

## WinRE/WinPE Limitations

- `sfc /scannow` requires `/offbootdir` and `/offwindir` in WinRE — the script handles this automatically
- `wmic` may be absent in minimal WinPE — CIM (`Get-CimInstance`) is used as primary with wmic as fallback
- `reagentc` may not be available in all PE builds — gracefully skipped if missing
- Drive letters are assigned dynamically in WinPE — `C:` is never assumed
- `bootrec /fixboot` returns "Access is denied" on UEFI systems — handled explicitly with logged explanation
- Transcript may fail if the log path is on a read-only volume — falls back to `%TEMP%`

---

## Requirements

- Windows 10/11 WinRE or WinPE (x64)
- PowerShell 5.1 (included in all Windows 10/11 PE builds)
- Run as Administrator (default in WinRE/WinPE)
- No internet connection required
- No external PowerShell modules required

---

## License

MIT
