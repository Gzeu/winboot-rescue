# 🛠️ winboot-rescue

**Professional Windows 10/11 Boot Recovery Utility for WinRE/WinPE**

A robust, modular PowerShell-based tool for diagnosing and repairing Windows boot failures caused by power outages, abrupt shutdowns, BCD corruption, missing/misconfigured EFI partitions, corrupt MBR/boot sectors, or damaged system files.

> ⚠️ **Run from a bootable WinRE/WinPE USB stick as Administrator.**

---

## Features

- ✅ **UEFI/GPT** and **Legacy BIOS/MBR** full support — correct tool selected automatically
- ✅ **Auto-detects firmware mode** (UEFI vs BIOS) using multiple methods including `GetFirmwareEnvironmentVariable` API
- ✅ **Scans ALL drives** for Windows installations — never assumes `C:`
- ✅ **Interactive multi-install selection** if multiple Windows copies are found
- ✅ **BCD backup** before any modification
- ✅ **EFI partition auto-mount** — assigns a temporary drive letter if ESP has none
- ✅ **Structured logging** — transcript + timestamped log + raw command output saved to USB stick
- ✅ **Offline SFC** (`sfc /scannow /offbootdir /offwindir`) — correct WinRE syntax
- ✅ **CHKDSK** with optional `/R` bad sector recovery
- ✅ **Graceful degradation** — if a command isn't available in the PE environment, it skips cleanly
- ✅ **Final verdict** — clear pass/fail/warning summary at end of session

---

## Quick Start

### 1. Put it on your USB stick

Copy both files to the **root** of your bootable WinRE/WinPE USB:

```
E:\
├── boot-repair.ps1
└── boot-repair.cmd
```

The `Logs\BootRepair\` folder will be created automatically on the USB stick when the tool runs.

### 2. Boot from USB

- **WinPE USB**: Boot normally, open Command Prompt
- **WinRE** (from Windows Setup): Choose *Repair your computer → Troubleshoot → Command Prompt*

### 3. Run

```cmd
E:\boot-repair.cmd
```

Or directly:

```powershell
powershell.exe -ExecutionPolicy Bypass -NoProfile -File E:\boot-repair.ps1
```

### 4. What to do first

Select **Option 1 (Collect diagnostics)** on your first run. This runs all diagnostic commands and saves raw output without touching anything. Review the summary, then choose the repair path.

---

## Menu Options

| Option | Description |
|--------|-------------|
| `[1]` | **Collect diagnostics only** — safe, read-only, saves all data |
| `[2]` | **Quick repair** — auto-detects UEFI/BIOS and runs the right repair |
| `[3]` | **Full repair** — diagnostics + repair + CHKDSK + export report |
| `[4]` | **Rebuild EFI boot files** — `bcdboot` targeting EFI System Partition |
| `[5]` | **BIOS/MBR boot repair** — `bootrec /fixmbr`, `/fixboot`, `/rebuildbcd` |
| `[6]` | **Run CHKDSK** on the Windows volume (optional `/R`) |
| `[7]` | **Run offline SFC** — System File Checker in offline mode |
| `[8]` | **Export full report** — saves a human-readable `.txt` summary |
| `[9]` | **Re-select Windows installation** |
| `[0]` | Exit |

---

## What Gets Repaired

### UEFI/GPT Systems
- Mounts the EFI System Partition (assigns a temp drive letter if needed)
- Verifies `\EFI\Microsoft\Boot\` exists
- Runs `bcdboot <WinDir> /s <EFI>: /f UEFI` as primary repair
- Falls back to `bcdboot /f ALL` if primary fails
- **Does NOT use `bootrec /fixboot` on UEFI** — this is intentional (it returns "Access is denied" on pure UEFI and is the wrong tool)

### BIOS/MBR Systems
- `bootrec /fixmbr` — rewrites the MBR
- `bootrec /fixboot` — rewrites the Volume Boot Record
- `bootrec /scanos` — scans for OS entries
- `bootrec /rebuildbcd` — rebuilds the BCD store (interactive)
- Falls back to `bcdboot <WinDir> /f BIOS` if rebuildbcd fails

---

## Log Files

All sessions create:

```
<USB>:\Logs\BootRepair\
├── boot-repair_YYYYMMDD_HHmmss.log      # Structured timestamped log
├── transcript_YYYYMMDD_HHmmss.txt       # Full PowerShell transcript
├── bcd_backup_YYYYMMDD_HHmmss.bak       # BCD export before repair
├── boot-repair-report_*.txt             # Human-readable final report
└── raw_YYYYMMDD_HHmmss\
    ├── bcdedit_enum_all.txt
    ├── diskpart_list.txt
    ├── bootrec_fixmbr.txt
    ├── bcdboot_uefi.txt
    ├── chkdsk_output.txt
    ├── sfc_offline.txt
    ├── CBS.log (copied from Windows)
    └── ...
```

If the USB is read-only, logs fall back to `%TEMP%\BootRepair\`.

---

## Requirements

| Requirement | Notes |
|-------------|-------|
| PowerShell 5.1+ | Available in WinRE/WinPE on Windows 10/11 |
| Administrator rights | Required for all boot repair operations |
| WinRE or WinPE | Recommended; will warn if run from a live OS |
| No internet required | Fully offline |
| No external modules | Pure PowerShell + built-in Windows tools |

---

## Compatibility

- ✅ Windows 10 (1903+)
- ✅ Windows 11 (all versions)
- ✅ WinRE (Recovery Environment)
- ✅ WinPE (Custom boot USB)
- ✅ UEFI + GPT
- ✅ Legacy BIOS + MBR
- ⚠️ Secure Boot: You may need to disable Secure Boot temporarily to boot custom WinPE

---

## WinRE/WinPE Limitations

The following limitations are handled in code:

- `sfc /scannow` in WinRE **requires** `/offbootdir` and `/offwindir` — always used correctly
- `wmic` may be absent in minimal WinPE — PowerShell CIM cmdlets used as fallback
- `reagentc` may not be available in all PE builds — skipped gracefully
- Drive letters are **dynamically assigned** in WinPE — `C:` is never assumed
- `bootrec /fixboot` returns *Access is denied* on UEFI systems — handled explicitly with a clear explanation
- Transcript may fail on read-only volumes — falls back to `%TEMP%`

---

## Security Notes

- **Never formats or deletes partitions**
- **Never marks volumes active without confirmation**
- **Always backs up BCD before modification**
- Asks for explicit Y/N confirmation before slow or destructive operations
- All operations are logged with full command arguments and exit codes

---

## License

MIT License — free to use, modify, and distribute.

---

## Contributing

PRs welcome. Please test changes in a WinPE environment before submitting.

Common test scenarios:
- VM with GPT/UEFI + deleted EFI partition
- VM with MBR + corrupted BCD
- VM with multiple Windows installations on separate drives
- Minimal WinPE (missing `wmic`, `reagentc`)