# Changelog

## 2.1.0 — 2026-05-14

### Added
- Expanded README with architecture overview, verdict table, limitation notes, and quick-start table
- Added `boot-repair.cmd` admin guard with clear error message on non-elevated launch
- Added `.gitignore` covering `Logs/`, `*.log`, `*.bak`, `*.tmp`

### Improved
- `boot-repair.ps1` now fully reconstructed: all 20 functions, 1396 lines, ~60 KB
- Clarified WinRE/WinPE limitation notes in script header

## 2.0.0 — 2026-05-14 (initial release)

- Professional PowerShell recovery utility for Windows 10/11 boot repair
- WinRE/WinPE environment detection
- USB-based structured logging and PowerShell transcript capture
- Disk, partition, and Windows installation discovery without assuming `C:`
- Diagnostics collection: `diskpart`, `bcdedit`, `mountvol`, `reagentc`, CIM
- UEFI/GPT repair flow with EFI partition detection and `bcdboot`
- BIOS/MBR repair flow with `bootrec` sequence and `bcdboot` fallback
- Optional CHKDSK and offline SFC
- Final report export and interactive repair menu
- `boot-repair.cmd` wrapper for launch from WinRE command prompt
