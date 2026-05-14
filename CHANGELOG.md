# Changelog

## [2.0.0] - 2026-05-14

### Added
- Full UEFI/GPT boot repair via `bcdboot` (correct primary tool)
- Auto-detection of EFI System Partition with temporary drive letter assignment
- Automatic cleanup of temporary EFI drive letters after repair
- Multi-Windows-installation detection and interactive selection
- Offline SFC (`sfc /scannow /offbootdir /offwindir`) with CBS.log capture
- BCD backup (`bcdedit /export`) before any modification
- Structured logging: timestamped log file + raw output files + PowerShell transcript
- Log fallback to `%TEMP%` when USB is read-only
- CIM-based disk/volume enumeration with `wmic` fallback
- `GetFirmwareEnvironmentVariable` API for reliable UEFI detection
- Explicit handling of `bootrec /fixboot` "Access is denied" on UEFI systems
- Final verdict output with color-coded pass/fail/warning
- Interactive Y/N confirmation for all slow or risky operations
- Full diagnostic report export as `.txt`
- `boot-repair.cmd` wrapper with admin check

### Architecture
- Modular PowerShell functions (13 sections)
- `Invoke-LoggedCommand` wrapper captures all output, exit codes, and raw logs
- No assumptions about drive letters (`C:` never hardcoded)
- Graceful degradation when commands are missing in minimal WinPE