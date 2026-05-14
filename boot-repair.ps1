#Requires -Version 5.1
<#
.SYNOPSIS
    Windows Boot Recovery Utility - Professional Boot Repair Tool
    For use in WinRE / WinPE environments from a bootable USB drive.

.DESCRIPTION
    Diagnoses and repairs common Windows 10/11 boot failures including:
    - BCD corruption after power loss or abrupt shutdown
    - Missing/unconfigured EFI System Partition
    - Corrupt MBR/boot sector (Legacy BIOS)
    - Missing or damaged boot files
    - Corrupt system image (via DISM offline repair)

.NOTES
    === WinRE/WinPE LIMITATIONS ===
    - 'sfc /scannow' requires /offbootdir and /offwindir parameters in WinRE
    - 'wmic' may be absent in minimal WinPE; PowerShell CIM cmdlets used as fallback
    - 'reagentc' may not be available in all PE builds
    - Drive letters are assigned dynamically in WinPE; C: is NOT assumed
    - Some bootrec operations (fixboot) may fail with "Access is denied" on UEFI systems - handled explicitly
    - Transcript may fail if the log path is on a read-only volume
    - DISM /RestoreHealth requires a valid Windows source (WIM/ESD or Windows Update)

    AUTHOR  : Boot Repair Utility v2.1
    COMPAT  : Windows 10/11 WinRE / WinPE (x64)
    USAGE   : Launch via boot-repair.cmd wrapper, or:
              powershell.exe -ExecutionPolicy Bypass -File boot-repair.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================
# SECTION 0: GLOBAL STATE & INITIALIZATION
# ============================================================

$Script:LogDir         = $null
$Script:LogFile        = $null
$Script:RawLogDir      = $null
$Script:ScriptDrive    = Split-Path -Qualifier $MyInvocation.MyCommand.Path
$Script:TranscriptPath = $null
$Script:BootMode       = $null   # 'UEFI' or 'BIOS'
$Script:PartStyle      = $null   # 'GPT' or 'MBR'
$Script:WinInstalls    = @()     # Array of detected Windows installations
$Script:SelectedWin    = $null   # Selected Windows installation object
$Script:EfiDrive       = $null   # Letter assigned to EFI partition (may be temporary)
$Script:EfiTempLetter  = $null   # If we assigned a temp letter, store it here for cleanup
$Script:SystemDisk     = $null   # Disk number hosting Windows
$Script:DiagComplete   = $false
$Script:RepairDone     = $false
$Script:WimSourcePath  = $null   # Optional: path to install.wim/esd for DISM source

# ============================================================
# SECTION 1: LOGGING INFRASTRUCTURE
# ============================================================

function Initialize-Logging {
    <#
    .SYNOPSIS Initializes the log directory on the USB drive and starts transcript.
    #>
    $logRoot = Join-Path $Script:ScriptDrive 'Logs\BootRepair'

    try {
        if (-not (Test-Path $logRoot)) {
            New-Item -ItemType Directory -Path $logRoot -Force | Out-Null
        }
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $Script:LogDir        = $logRoot
        $Script:LogFile       = Join-Path $logRoot "boot-repair_$timestamp.log"
        $Script:RawLogDir     = Join-Path $logRoot "raw_$timestamp"
        $Script:TranscriptPath = Join-Path $logRoot "transcript_$timestamp.txt"
        New-Item -ItemType Directory -Path $Script:RawLogDir -Force | Out-Null

        try { Start-Transcript -Path $Script:TranscriptPath -Append | Out-Null } catch { <# transcript not critical #> }

        Write-Log "=== Windows Boot Repair Utility v2.1 ==="
        Write-Log "Session started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        Write-Log "Script location: $($Script:ScriptDrive)"
        Write-Log "Log directory:   $($Script:LogDir)"
        Write-Log "Log file:        $($Script:LogFile)"
        Write-Log "Raw logs:        $($Script:RawLogDir)"
        Write-Log "========================================"
    } catch {
        $fallback = Join-Path $env:TEMP 'BootRepair'
        New-Item -ItemType Directory -Path $fallback -Force -ErrorAction SilentlyContinue | Out-Null
        $Script:LogDir   = $fallback
        $Script:LogFile  = Join-Path $fallback 'boot-repair.log'
        $Script:RawLogDir = Join-Path $fallback 'raw'
        New-Item -ItemType Directory -Path $Script:RawLogDir -Force -ErrorAction SilentlyContinue | Out-Null
        Write-Host "[WARN] Could not create log on USB ($logRoot). Falling back to: $fallback" -ForegroundColor Yellow
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','STEP','RAW','SECTION')]
        [string]$Level = 'INFO',
        [switch]$NoConsole
    )

    $ts   = Get-Date -Format 'HH:mm:ss'
    $line = "[$ts][$Level] $Message"

    if ($Script:LogFile) {
        try { Add-Content -Path $Script:LogFile -Value $line -Encoding UTF8 } catch { <# best effort #> }
    }

    if ($NoConsole) { return }

    switch ($Level) {
        'INFO'    { Write-Host $line -ForegroundColor Cyan }
        'WARN'    { Write-Host $line -ForegroundColor Yellow }
        'ERROR'   { Write-Host $line -ForegroundColor Red }
        'SUCCESS' { Write-Host $line -ForegroundColor Green }
        'STEP'    { Write-Host "`n$line" -ForegroundColor Magenta }
        'SECTION' { Write-Host "`n$('='*60)`n$line`n$('='*60)" -ForegroundColor White }
        'RAW'     { Write-Host $line -ForegroundColor DarkGray }
        default   { Write-Host $line }
    }
}

function Invoke-LoggedCommand {
    <#
    .SYNOPSIS Executes a command, captures all output, logs it, and returns structured result.
    .RETURNS Object with .Output (string[]), .ExitCode (int), .Success (bool), .Skipped (bool)
    #>
    param(
        [Parameter(Mandatory)][string]$Command,
        [string[]]$Arguments   = @(),
        [string]  $Description = '',
        [string]  $SaveRawAs   = '',
        [switch]  $IgnoreExit
    )

    $desc = if ($Description) { $Description } else { "$Command $($Arguments -join ' ')" }
    Write-Log "Executing: $desc" -Level STEP

    $cmdPath = Get-Command $Command -ErrorAction SilentlyContinue
    if (-not $cmdPath) {
        Write-Log "Command not found: $Command - skipping." -Level WARN
        return [PSCustomObject]@{ Output = @(); ExitCode = -1; Success = $false; Skipped = $true }
    }

    $result = [PSCustomObject]@{ Output = @(); ExitCode = 0; Success = $false; Skipped = $false }

    try {
        $output = & $Command @Arguments 2>&1
        $result.ExitCode = $LASTEXITCODE
        $result.Output   = @($output | ForEach-Object { "$_" })
        $result.Success  = ($result.ExitCode -eq 0)

        foreach ($line in $result.Output) {
            Write-Log $line -Level RAW -NoConsole:($result.Output.Count -gt 50)
        }

        if ($result.Output.Count -gt 50) {
            Write-Host "  [Output: $($result.Output.Count) lines - see log]" -ForegroundColor DarkGray
        } else {
            $result.Output | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }
        }

        if ($result.ExitCode -ne 0 -and -not $IgnoreExit) {
            Write-Log "Command exited with code $($result.ExitCode): $desc" -Level WARN
        } elseif ($result.Success) {
            Write-Log "Command completed successfully: $desc" -Level SUCCESS
        }

        if ($SaveRawAs -and $Script:RawLogDir) {
            $rawPath = Join-Path $Script:RawLogDir $SaveRawAs
            $result.Output | Set-Content -Path $rawPath -Encoding UTF8 -ErrorAction SilentlyContinue
        }

    } catch {
        $result.ExitCode = -1
        $result.Output   = @("EXCEPTION: $_")
        Write-Log "Exception running '$Command': $_" -Level ERROR
    }

    Write-Log "Exit code: $($result.ExitCode)" -Level INFO
    return $result
}

function Save-RawOutput {
    param([string]$FileName, [string]$Content)
    if ($Script:RawLogDir) {
        $path = Join-Path $Script:RawLogDir $FileName
        try { Set-Content -Path $path -Value $Content -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
    }
}

# ============================================================
# SECTION 2: ENVIRONMENT DETECTION
# ============================================================

function Get-Environment {
    <#
    .SYNOPSIS Detects whether running in WinRE/WinPE and identifies the USB drive letter.
    #>
    Write-Log "Detecting execution environment..." -Level SECTION

    $isWinPE = $false
    $isWinRE = $false

    if (Test-Path 'X:\Windows\System32\winpe.jpg' -ErrorAction SilentlyContinue) { $isWinPE = $true }
    if (Test-Path 'X:\' -ErrorAction SilentlyContinue) { $isWinPE = $true }

    if (Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\MiniNT' -ErrorAction SilentlyContinue) {
        $isWinPE = $true
        $isWinRE = $true
    }

    if ($env:WINPE -eq '1') { $isWinPE = $true }

    if ($isWinRE) {
        Write-Log "Environment: Windows Recovery Environment (WinRE)" -Level SUCCESS
    } elseif ($isWinPE) {
        Write-Log "Environment: Windows Preinstallation Environment (WinPE)" -Level SUCCESS
    } else {
        Write-Log "Environment: Standard Windows session (not WinRE/WinPE)" -Level WARN
        Write-Log "  WARNING: Some repairs work best from WinRE. Continue at your own risk." -Level WARN
    }

    Write-Log "Script running from drive: $Script:ScriptDrive"
    Write-Log "PowerShell version: $($PSVersionTable.PSVersion)"
    Write-Log "OS version: $([System.Environment]::OSVersion.VersionString)"

    # Auto-detect WIM source on the USB stick (sources\install.wim / install.esd)
    foreach ($wimName in @('install.wim','install.esd','sources\install.wim','sources\install.esd')) {
        $wimCandidate = Join-Path $Script:ScriptDrive $wimName
        if (Test-Path $wimCandidate -ErrorAction SilentlyContinue) {
            $Script:WimSourcePath = $wimCandidate
            Write-Log "  Found WIM/ESD source on USB: $wimCandidate" -Level SUCCESS
            break
        }
    }

    return [PSCustomObject]@{
        IsWinPE = $isWinPE
        IsWinRE = $isWinRE
        UsbDrive = $Script:ScriptDrive
    }
}

function Get-BootMode {
    <#
    .SYNOPSIS Detects whether the system firmware is UEFI or Legacy BIOS.
    .NOTES Uses multiple detection methods including GetFirmwareEnvironmentVariable API.
    #>
    Write-Log "Detecting firmware/boot mode (UEFI vs BIOS)..." -Level STEP

    $bootMode = 'BIOS'

    try {
        $uefiKey = Get-ItemProperty 'HKLM:\HARDWARE\UEFI' -ErrorAction SilentlyContinue
        if ($uefiKey) { $bootMode = 'UEFI' }
    } catch {}

    try {
        $bcdTest = & bcdedit.exe /enum firmware 2>&1
        if ($LASTEXITCODE -eq 0 -and $bcdTest -match 'firmware') {
            $bootMode = 'UEFI'
        }
    } catch {}

    # Most reliable: Win32 GetFirmwareEnvironmentVariableA
    # Error 998 (ERROR_NOACCESS) = UEFI present, Error 1 (ERROR_INVALID_FUNCTION) = BIOS only
    try {
        $sig = '[DllImport("kernel32.dll", SetLastError=true)] public static extern uint GetFirmwareEnvironmentVariableA(string lpName, string lpGuid, System.IntPtr pBuffer, uint nSize);'
        $type = Add-Type -MemberDefinition $sig -Name 'FirmwareCheck' -Namespace 'Win32' -PassThru -ErrorAction SilentlyContinue
        if ($type) {
            $null = $type::GetFirmwareEnvironmentVariableA('', '{00000000-0000-0000-0000-000000000000}', [System.IntPtr]::Zero, 0)
            $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
            if ($err -eq 998) { $bootMode = 'UEFI' }
            elseif ($err -eq 1) { $bootMode = 'BIOS' }
        }
    } catch {}

    $Script:BootMode = $bootMode
    Write-Log "Detected firmware mode: $bootMode" -Level SUCCESS
    return $bootMode
}

# ============================================================
# SECTION 3: DISK & PARTITION INVENTORY
# ============================================================

function Get-DiskInventory {
    <#
    .SYNOPSIS Enumerates all disks, volumes, and identifies partition style.
    .NOTES Also detects NVMe drives and warns if StorNVMe driver may be missing in WinPE.
    #>
    Write-Log "Enumerating disks and volumes..." -Level SECTION

    $disks = @()

    try {
        $diskObjects = Get-Disk -ErrorAction SilentlyContinue
        if ($diskObjects) {
            foreach ($disk in $diskObjects) {
                $partitions = Get-Partition -DiskNumber $disk.DiskNumber -ErrorAction SilentlyContinue
                $volumes    = @()
                foreach ($part in $partitions) {
                    $vol = Get-Volume -Partition $part -ErrorAction SilentlyContinue
                    $volumes += [PSCustomObject]@{
                        PartitionNumber = $part.PartitionNumber
                        DriveLetter     = $part.DriveLetter
                        Size            = [math]::Round($part.Size / 1GB, 2)
                        Type            = $part.Type
                        IsActive        = $part.IsActive
                        GptType         = $part.GptType
                        VolumeLabel     = if ($vol) { $vol.FileSystemLabel } else { '' }
                        FileSystem      = if ($vol) { $vol.FileSystem } else { '' }
                    }
                }
                $diskObj = [PSCustomObject]@{
                    DiskNumber        = $disk.DiskNumber
                    Model             = $disk.FriendlyName
                    Size              = [math]::Round($disk.Size / 1GB, 2)
                    PartitionStyle    = $disk.PartitionStyle
                    OperationalStatus = $disk.OperationalStatus
                    Partitions        = $volumes
                }
                $disks += $diskObj

                # NVMe detection: warn if WinPE may lack driver
                if ($disk.FriendlyName -match 'NVMe|NVME') {
                    Write-Log "  [NVMe] Detected NVMe drive: $($disk.FriendlyName)" -Level WARN
                    Write-Log "  [NVMe] If Windows cannot be found, your WinPE may lack the NVMe/StorNVMe driver." -Level WARN
                }

                Write-Log "  Disk $($disk.DiskNumber): $($disk.FriendlyName) | $($disk.PartitionStyle) | $([math]::Round($disk.Size/1GB,1)) GB | $($disk.OperationalStatus)"
                foreach ($v in $volumes) {
                    $letter = if ($v.DriveLetter) { "$($v.DriveLetter):" } else { '(no letter)' }
                    Write-Log "    Part $($v.PartitionNumber): $letter | $($v.FileSystem) | $([math]::Round($v.Size,2)) GB | Type=$($v.Type) | GptType=$($v.GptType)"
                }
            }
            $mainDisk = $disks | Where-Object { $_.PartitionStyle -ne 'RAW' } | Select-Object -First 1
            if ($mainDisk) {
                $Script:PartStyle = $mainDisk.PartitionStyle
                Write-Log "Primary partition style: $($Script:PartStyle)" -Level SUCCESS
            }
        } else {
            throw "Get-Disk returned no results"
        }
    } catch {
        Write-Log "Get-Disk failed ($_) - falling back to CIM/WMI..." -Level WARN
        try {
            $cimDisks = Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue
            foreach ($cd in $cimDisks) {
                Write-Log "  Disk: $($cd.DeviceID) | $($cd.Model) | $([math]::Round($cd.Size/1GB,1)) GB"
                $disks += [PSCustomObject]@{
                    DiskNumber = $cd.Index; Model = $cd.Model
                    Size = [math]::Round($cd.Size/1GB,2); PartitionStyle = 'Unknown'
                    Partitions = @()
                }
            }
        } catch {
            Write-Log "CIM disk enumeration also failed: $_" -Level ERROR
        }
    }

    $dpScript = "list disk`r`nlist volume`r`nlist partition"
    $dpFile    = Join-Path $env:TEMP 'dp_list.txt'
    $dpScript | Set-Content $dpFile -Encoding ASCII -ErrorAction SilentlyContinue
    Invoke-LoggedCommand -Command 'diskpart' -Arguments @("/s", $dpFile) `
        -Description "diskpart list disk/vol/partition" -SaveRawAs "diskpart_list.txt" -IgnoreExit | Out-Null
    Remove-Item $dpFile -Force -ErrorAction SilentlyContinue

    return $disks
}

function Get-EfiPartition {
    <#
    .SYNOPSIS Locates the EFI System Partition, assigns a drive letter if needed.
    .RETURNS Object with DriveLetter, DiskNumber, PartitionNumber, NeedsCleanup (bool), Found (bool).
    #>
    Write-Log "Locating EFI System Partition..." -Level STEP

    $efiInfo = [PSCustomObject]@{
        DriveLetter     = $null
        DiskNumber      = $null
        PartitionNumber = $null
        NeedsCleanup    = $false
        Found           = $false
    }

    $efiGuid = '{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}'

    try {
        # Only look on the same disk as the selected Windows installation if known
        $allDisks = Get-Disk -ErrorAction SilentlyContinue
        $searchDisks = if ($Script:SystemDisk -ne $null) {
            $allDisks | Where-Object { $_.DiskNumber -eq $Script:SystemDisk }
        } else { $allDisks }

        foreach ($disk in $searchDisks) {
            $parts = Get-Partition -DiskNumber $disk.DiskNumber -ErrorAction SilentlyContinue
            foreach ($p in $parts) {
                $gpt = "$($p.GptType)".ToLower()
                if ($gpt -eq $efiGuid -or $p.Type -eq 'System') {
                    $efiInfo.DiskNumber      = $disk.DiskNumber
                    $efiInfo.PartitionNumber = $p.PartitionNumber
                    $efiInfo.DriveLetter     = $p.DriveLetter
                    $efiInfo.Found           = $true
                    Write-Log "  Found EFI partition: Disk $($disk.DiskNumber), Partition $($p.PartitionNumber), Type=$($p.Type)"

                    if (-not $p.DriveLetter) {
                        Write-Log "  EFI partition has no drive letter. Attempting to assign one..." -Level WARN
                        $freeLetter = Get-AvailableDriveLetter
                        if ($freeLetter) {
                            try {
                                $dpAssign = "select disk $($disk.DiskNumber)`r`nselect partition $($p.PartitionNumber)`r`nassign letter=$freeLetter`r`n"
                                $dpFile = Join-Path $env:TEMP 'dp_assign.txt'
                                $dpAssign | Set-Content $dpFile -Encoding ASCII
                                $r = Invoke-LoggedCommand -Command 'diskpart' -Arguments @("/s", $dpFile) `
                                        -Description "Assign letter $freeLetter to EFI partition" -IgnoreExit
                                Remove-Item $dpFile -Force -ErrorAction SilentlyContinue
                                if ($r.Success -or ($r.Output -join '') -match 'successfully') {
                                    $efiInfo.DriveLetter  = $freeLetter
                                    $efiInfo.NeedsCleanup = $true
                                    $Script:EfiTempLetter = $freeLetter
                                    Write-Log "  Assigned temporary letter $freeLetter`: to EFI partition." -Level SUCCESS
                                }
                            } catch {
                                Write-Log "  Could not assign letter to EFI partition: $_" -Level ERROR
                            }
                        }
                    }
                    break
                }
            }
            if ($efiInfo.Found) { break }
        }
    } catch {
        Write-Log "Error enumerating EFI partition via PowerShell: $_ - trying diskpart..." -Level WARN
        $dpScript = "list disk"
        $dpFile   = Join-Path $env:TEMP 'dp_efi.txt'
        $dpScript | Set-Content $dpFile -Encoding ASCII
        Invoke-LoggedCommand -Command 'diskpart' -Arguments @("/s", $dpFile) -IgnoreExit | Out-Null
        Remove-Item $dpFile -Force -ErrorAction SilentlyContinue
    }

    if (-not $efiInfo.Found) {
        Write-Log "  EFI System Partition NOT found. System may be BIOS/MBR or EFI partition is damaged." -Level WARN
    } else {
        $Script:EfiDrive = $efiInfo.DriveLetter
        Write-Log "  EFI partition letter: $($efiInfo.DriveLetter)" -Level SUCCESS
    }

    return $efiInfo
}

function Remove-TempEfiLetter {
    <#
    .SYNOPSIS Removes the temporary drive letter assigned to the EFI partition during repair.
    #>
    if ($Script:EfiTempLetter) {
        Write-Log "Removing temporary EFI drive letter $($Script:EfiTempLetter):..." -Level STEP
        $dpScript = "select volume $($Script:EfiTempLetter)`r`nremove letter=$($Script:EfiTempLetter)`r`n"
        $dpFile   = Join-Path $env:TEMP 'dp_remove.txt'
        $dpScript | Set-Content $dpFile -Encoding ASCII
        Invoke-LoggedCommand -Command 'diskpart' -Arguments @("/s", $dpFile) `
            -Description "Remove temp EFI letter $($Script:EfiTempLetter)" -IgnoreExit | Out-Null
        Remove-Item $dpFile -Force -ErrorAction SilentlyContinue
        $Script:EfiTempLetter = $null
    }
}

function Get-AvailableDriveLetter {
    <#
    .SYNOPSIS Returns the first unused drive letter (starting from S to avoid common conflicts).
    #>
    $used = [System.IO.DriveInfo]::GetDrives() | ForEach-Object { $_.Name[0] }
    foreach ($l in @('S','T','U','V','W','R','Q','P','O','N','M','L','K','J','I')) {
        if ($l -notin $used) { return $l }
    }
    return $null
}

# ============================================================
# SECTION 4: WINDOWS INSTALLATION DETECTION
# ============================================================

function Get-WindowsInstallations {
    <#
    .SYNOPSIS Scans all accessible drive letters for valid Windows installations.
    .NOTES Does NOT assume C:. Checks for all key Windows files on every mounted drive.
    #>
    Write-Log "Scanning for Windows installations on all accessible drives..." -Level SECTION

    $installations = @()

    $drives = [System.IO.DriveInfo]::GetDrives() |
              Where-Object { $_.DriveType -in @('Fixed','Removable') -and $_.IsReady } |
              ForEach-Object { $_.Name.TrimEnd('\') }

    Write-Log "  Checking drives: $($drives -join ', ')"

    foreach ($drive in $drives) {
        if ($drive -eq $Script:ScriptDrive.TrimEnd('\')) {
            Write-Log "  Skipping script drive: $drive"
            continue
        }

        $systemConfig   = Join-Path $drive 'Windows\System32\Config\SYSTEM'
        $explorer       = Join-Path $drive 'Windows\explorer.exe'
        $winloadEfi     = Join-Path $drive 'Windows\System32\winload.efi'
        $winloadExe     = Join-Path $drive 'Windows\System32\winload.exe'
        $ntoskrnl       = Join-Path $drive 'Windows\System32\ntoskrnl.exe'

        $hasConfig   = Test-Path $systemConfig   -ErrorAction SilentlyContinue
        $hasExplorer = Test-Path $explorer        -ErrorAction SilentlyContinue
        $hasWinload  = (Test-Path $winloadEfi -ErrorAction SilentlyContinue) -or
                       (Test-Path $winloadExe -ErrorAction SilentlyContinue)
        $hasNtoskrnl = Test-Path $ntoskrnl       -ErrorAction SilentlyContinue

        if ($hasConfig -and $hasExplorer -and $hasNtoskrnl) {
            $winVer  = 'Unknown'
            $buildNum = ''
            try {
                $hivePath = Join-Path $drive 'Windows\System32\Config\SOFTWARE'
                if (Test-Path $hivePath) {
                    $hiveKey = "HKLM\OFFLINE_CHECK_$$"
                    $regLoad = & reg.exe load $hiveKey $hivePath 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        $psKey = "Registry::$hiveKey\Microsoft\Windows NT\CurrentVersion"
                        $ntProps = Get-ItemProperty $psKey -ErrorAction SilentlyContinue
                        if ($ntProps) {
                            $winVer   = "$($ntProps.ProductName)"
                            $buildNum = "$($ntProps.CurrentBuildNumber).$($ntProps.UBR)"
                        }
                        & reg.exe unload $hiveKey 2>&1 | Out-Null
                    }
                }
            } catch { <# version detection not critical #> }

            $winBootType = if ($hasWinload -and (Test-Path $winloadEfi -ErrorAction SilentlyContinue)) { 'UEFI' }
                           elseif ($hasWinload) { 'BIOS' }
                           else { 'Unknown' }

            # Track which disk this installation lives on
            try {
                $driveObj = Get-Partition | Where-Object { $_.DriveLetter -eq $drive[0] } | Select-Object -First 1
                if ($driveObj) { $Script:SystemDisk = $driveObj.DiskNumber }
            } catch {}

            $install = [PSCustomObject]@{
                Drive         = $drive
                WindowsDir    = Join-Path $drive 'Windows'
                Version       = $winVer
                Build         = $buildNum
                BootType      = $winBootType
                HasWinloadEfi = Test-Path $winloadEfi -ErrorAction SilentlyContinue
                HasWinloadExe = Test-Path $winloadExe -ErrorAction SilentlyContinue
            }
            $installations += $install
            Write-Log "  [FOUND] Windows at $drive\ | $winVer | Build $buildNum | BootType: $winBootType" -Level SUCCESS
        } else {
            Write-Log "  [$drive] Not a valid Windows installation (config=$hasConfig, explorer=$hasExplorer, ntoskrnl=$hasNtoskrnl)"
        }
    }

    if ($installations.Count -eq 0) {
        Write-Log "  No valid Windows installation found on any accessible drive!" -Level ERROR
    } else {
        Write-Log "  Total Windows installations found: $($installations.Count)" -Level SUCCESS
    }

    $Script:WinInstalls = $installations
    return $installations
}

function Select-WindowsInstallation {
    <#
    .SYNOPSIS Presents an interactive selection menu if multiple installations are detected.
    #>
    if ($Script:WinInstalls.Count -eq 0) {
        Write-Log "No Windows installations available for selection." -Level ERROR
        return $null
    }

    if ($Script:WinInstalls.Count -eq 1) {
        $Script:SelectedWin = $Script:WinInstalls[0]
        Write-Log "Auto-selected single Windows installation: $($Script:SelectedWin.Drive)" -Level SUCCESS
        return $Script:SelectedWin
    }

    Write-Host "`n$('='*60)" -ForegroundColor White
    Write-Host "  MULTIPLE WINDOWS INSTALLATIONS DETECTED" -ForegroundColor Yellow
    Write-Host "$('='*60)" -ForegroundColor White
    for ($i = 0; $i -lt $Script:WinInstalls.Count; $i++) {
        $inst = $Script:WinInstalls[$i]
        Write-Host "  [$($i+1)] $($inst.Drive)\ | $($inst.Version) | Build $($inst.Build) | $($inst.BootType)" -ForegroundColor Cyan
    }
    Write-Host ""

    do {
        $choice = Read-Host "Select installation (1-$($Script:WinInstalls.Count))"
        $idx    = [int]$choice - 1
    } while ($idx -lt 0 -or $idx -ge $Script:WinInstalls.Count)

    $Script:SelectedWin = $Script:WinInstalls[$idx]
    Write-Log "User selected Windows installation: $($Script:SelectedWin.Drive)\" -Level SUCCESS
    return $Script:SelectedWin
}

# ============================================================
# SECTION 5: DIAGNOSTIC COLLECTION
# ============================================================

function Collect-Diagnostics {
    <#
    .SYNOPSIS Runs all diagnostic commands and saves output to raw log files.
    #>
    Write-Log "Collecting full system diagnostics..." -Level SECTION

    Invoke-LoggedCommand -Command 'bcdedit' -Arguments @('/enum','all') `
        -Description "BCD Store (all entries)" -SaveRawAs "bcdedit_enum_all.txt" -IgnoreExit | Out-Null

    Invoke-LoggedCommand -Command 'bcdedit' -Arguments @('/enum','firmware') `
        -Description "BCD Firmware entries (UEFI)" -SaveRawAs "bcdedit_enum_firmware.txt" -IgnoreExit | Out-Null

    Invoke-LoggedCommand -Command 'mountvol' -Arguments @() `
        -Description "Mount points and volumes" -SaveRawAs "mountvol.txt" -IgnoreExit | Out-Null

    Invoke-LoggedCommand -Command 'reagentc' -Arguments @('/info') `
        -Description "Windows Recovery Agent info" -SaveRawAs "reagentc_info.txt" -IgnoreExit | Out-Null

    # DISM image health check (informational, non-destructive)
    if ($Script:SelectedWin) {
        Invoke-LoggedCommand -Command 'dism' `
            -Arguments @('/Image:'+$Script:SelectedWin.Drive, '/Get-Intl') `
            -Description "DISM get locale info for selected install" -SaveRawAs "dism_intl.txt" -IgnoreExit | Out-Null
    }

    try {
        $logicalDisks = Get-CimInstance -ClassName Win32_LogicalDisk -ErrorAction SilentlyContinue |
            Select-Object DeviceID, DriveType, Size, FreeSpace, FileSystem, VolumeName |
            Format-Table -AutoSize | Out-String
        Write-Log "Logical disks (CIM):" -Level STEP
        Write-Host $logicalDisks -ForegroundColor DarkGray
        Save-RawOutput -FileName "logical_disks_cim.txt" -Content $logicalDisks
    } catch {
        Write-Log "CIM logical disk query failed: $_ - trying wmic..." -Level WARN
        Invoke-LoggedCommand -Command 'wmic' -Arguments @('logicaldisk','get','deviceid,size,freespace,filesystem,volumename') `
            -Description "wmic logicaldisk fallback" -SaveRawAs "wmic_logicaldisk.txt" -IgnoreExit | Out-Null
    }

    Invoke-LoggedCommand -Command 'bootrec' -Arguments @('/scanos') `
        -Description "Scan OS entries (bootrec)" -SaveRawAs "bootrec_scanos.txt" -IgnoreExit | Out-Null

    try {
        $sysinfo = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue |
            Select-Object Manufacturer, Model, TotalPhysicalMemory | Format-List | Out-String
        Save-RawOutput -FileName "system_info.txt" -Content $sysinfo
    } catch {}

    try {
        $biosInfo = Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue |
            Select-Object Manufacturer, Version, ReleaseDate, SMBIOSBIOSVersion | Format-List | Out-String
        Save-RawOutput -FileName "bios_info.txt" -Content $biosInfo
    } catch {}

    $Script:DiagComplete = $true
    Write-Log "Diagnostic collection complete. Raw logs in: $Script:RawLogDir" -Level SUCCESS
}

function Show-DiagnosticSummary {
    <#
    .SYNOPSIS Prints a human-readable summary of the detected system state.
    #>
    Write-Host "`n$('='*60)" -ForegroundColor White
    Write-Host "  DIAGNOSTIC SUMMARY" -ForegroundColor Cyan
    Write-Host "$('='*60)" -ForegroundColor White
    Write-Host "  Firmware mode:       $Script:BootMode" -ForegroundColor $(if ($Script:BootMode) {'Green'} else {'Red'})
    Write-Host "  Partition style:     $Script:PartStyle" -ForegroundColor $(if ($Script:PartStyle) {'Green'} else {'Yellow'})
    Write-Host "  EFI drive letter:    $(if ($Script:EfiDrive) {$Script:EfiDrive+':'} else {'Not assigned / Not found'})" `
        -ForegroundColor $(if ($Script:EfiDrive) {'Green'} else {'Yellow'})
    Write-Host "  Windows found:       $($Script:WinInstalls.Count)" `
        -ForegroundColor $(if ($Script:WinInstalls.Count -gt 0) {'Green'} else {'Red'})
    if ($Script:SelectedWin) {
        Write-Host "  Selected install:    $($Script:SelectedWin.Drive)\ | $($Script:SelectedWin.Version) | Build $($Script:SelectedWin.Build)" -ForegroundColor Green
    }
    if ($Script:WimSourcePath) {
        Write-Host "  WIM/ESD source:      $Script:WimSourcePath" -ForegroundColor Green
    } else {
        Write-Host "  WIM/ESD source:      Not found on USB (DISM /RestoreHealth will use Windows Update if available)" -ForegroundColor Yellow
    }
    Write-Host "  Log directory:       $Script:LogDir" -ForegroundColor DarkGray
    Write-Host "$('='*60)`n" -ForegroundColor White
}

# ============================================================
# SECTION 6: BCD BACKUP
# ============================================================

function Backup-BcdStore {
    <#
    .SYNOPSIS Creates a backup of the current BCD store before any modifications.
    .RETURNS Path to backup file, or $null on failure.
    #>
    Write-Log "Backing up BCD store..." -Level STEP

    if (-not $Script:LogDir) {
        Write-Log "Log directory not initialized - cannot back up BCD." -Level WARN
        return $null
    }

    $backupPath = Join-Path $Script:LogDir "bcd_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').bak"

    $result = Invoke-LoggedCommand -Command 'bcdedit' -Arguments @('/export', $backupPath) `
                -Description "Export BCD store to $backupPath" -IgnoreExit

    if ($result.Success -or (Test-Path $backupPath)) {
        Write-Log "BCD backup saved: $backupPath" -Level SUCCESS
        return $backupPath
    }

    Write-Log "bcdedit /export failed. Attempting direct file copy..." -Level WARN

    $bcdPaths = @(
        'X:\Boot\BCD',
        'C:\Boot\BCD',
        "$($Script:EfiDrive):\EFI\Microsoft\Boot\BCD"
    )

    foreach ($src in $bcdPaths) {
        if (Test-Path $src -ErrorAction SilentlyContinue) {
            try {
                Copy-Item $src $backupPath -Force -ErrorAction Stop
                Write-Log "BCD file copied from $src to $backupPath" -Level SUCCESS
                return $backupPath
            } catch {
                Write-Log "Could not copy from $src : $_" -Level WARN
            }
        }
    }

    Write-Log "BCD backup failed - proceeding without backup (be careful!)" -Level ERROR
    return $null
}

# ============================================================
# SECTION 7: UEFI/GPT BOOT REPAIR
# ============================================================

function Repair-UefiBoot {
    <#
    .SYNOPSIS Repairs the UEFI boot environment for a GPT system.
    .NOTES Uses bcdboot as primary tool. bootrec /fixboot is NOT used on UEFI (returns Access Denied).
           Attempts /locale en-US to prevent locale-mismatch boot failures.
    #>
    Write-Log "Starting UEFI/GPT boot repair..." -Level SECTION

    if (-not $Script:SelectedWin) {
        Write-Log "No Windows installation selected. Cannot proceed." -Level ERROR
        return $false
    }

    $winDir = $Script:SelectedWin.WindowsDir
    Write-Log "Target Windows directory: $winDir"

    $efiInfo = Get-EfiPartition
    if (-not $efiInfo.Found) {
        Write-Log "EFI System Partition not found. Cannot perform UEFI boot repair." -Level ERROR
        Write-Log "  -> The EFI partition may be missing, deleted, or damaged." -Level ERROR
        Write-Log "  -> If disk is GPT, consider recreating EFI partition manually with diskpart." -Level ERROR
        Write-Log "  -> Alternatively, convert to MBR and use BIOS repair (destructive, not recommended)." -Level ERROR
        return $false
    }

    $efiLetter = $efiInfo.DriveLetter
    if (-not $efiLetter) {
        Write-Log "EFI partition found but has no drive letter - cannot proceed." -Level ERROR
        return $false
    }

    Write-Log "EFI partition accessible at $efiLetter`:" -Level SUCCESS

    $efiBootPath = "$efiLetter`:\EFI\Microsoft\Boot"
    $efiBootExists = Test-Path $efiBootPath -ErrorAction SilentlyContinue

    if ($efiBootExists) {
        Write-Log "  EFI\Microsoft\Boot already exists at $efiBootPath" -Level INFO
        $bcdPath = "$efiBootPath\BCD"
        if (Test-Path $bcdPath -ErrorAction SilentlyContinue) {
            Write-Log "  BCD file present at $bcdPath" -Level SUCCESS
        } else {
            Write-Log "  BCD file MISSING at $bcdPath - will rebuild" -Level WARN
        }
    } else {
        Write-Log "  EFI\Microsoft\Boot directory does NOT exist - will create via bcdboot" -Level WARN
    }

    Backup-BcdStore | Out-Null

    # Primary: bcdboot with explicit UEFI and locale (avoids locale mismatch boot failures)
    Write-Log "Running: bcdboot $winDir /s $efiLetter`: /f UEFI /l en-us" -Level STEP
    $bcdbootResult = Invoke-LoggedCommand -Command 'bcdboot' `
        -Arguments @($winDir, '/s', "$efiLetter`:", '/f', 'UEFI', '/l', 'en-us') `
        -Description "bcdboot UEFI /l en-us (primary)" -SaveRawAs "bcdboot_uefi.txt" -IgnoreExit

    if ($bcdbootResult.Success) {
        Write-Log "bcdboot UEFI completed successfully." -Level SUCCESS
        $Script:RepairDone = $true
        Remove-TempEfiLetter
        return $true
    }

    Write-Log "bcdboot UEFI failed (exit $($bcdbootResult.ExitCode)). Trying /f ALL..." -Level WARN

    $bcdbootAll = Invoke-LoggedCommand -Command 'bcdboot' `
        -Arguments @($winDir, '/s', "$efiLetter`:", '/f', 'ALL') `
        -Description "bcdboot ALL (fallback)" -SaveRawAs "bcdboot_all.txt" -IgnoreExit

    if ($bcdbootAll.Success) {
        Write-Log "bcdboot /f ALL completed successfully." -Level SUCCESS
        $Script:RepairDone = $true
        Remove-TempEfiLetter
        return $true
    }

    if (Test-Path "$efiBootPath\BCD" -ErrorAction SilentlyContinue) {
        Write-Log "BCD file present despite non-zero exit code. Boot may be repaired." -Level WARN
        $Script:RepairDone = $true
        Remove-TempEfiLetter
        return $true
    }

    Write-Log "NOTE: 'bootrec /fixboot' is NOT used here intentionally." -Level INFO
    Write-Log "  On UEFI systems, bootrec /fixboot returns 'Access is denied' - this is expected." -Level INFO
    Write-Log "  bcdboot is the correct tool for UEFI boot file repair." -Level INFO

    Write-Log "UEFI boot repair failed. Manual intervention may be required." -Level ERROR
    Remove-TempEfiLetter
    return $false
}

# ============================================================
# SECTION 8: BIOS/MBR BOOT REPAIR
# ============================================================

function Repair-BiosBoot {
    <#
    .SYNOPSIS Repairs the BIOS/MBR boot environment for an MBR-partitioned disk.
    .NOTES Runs bootrec in correct logical order with fallbacks.
    #>
    Write-Log "Starting BIOS/MBR boot repair..." -Level SECTION

    if (-not $Script:SelectedWin) {
        Write-Log "No Windows installation selected. Cannot proceed." -Level ERROR
        return $false
    }

    $winDir = $Script:SelectedWin.WindowsDir

    Backup-BcdStore | Out-Null

    $allSuccess = $true

    # Step 1: Fix the MBR
    Write-Log "Step 1/5: Fixing MBR with bootrec /fixmbr..." -Level STEP
    Write-Log "  WARNING: This will overwrite the Master Boot Record." -Level WARN
    $confirm = Get-UserConfirmation "Overwrite MBR with bootrec /fixmbr? [Y/N]"
    if ($confirm) {
        $r = Invoke-LoggedCommand -Command 'bootrec' -Arguments @('/fixmbr') `
            -Description "bootrec /fixmbr" -SaveRawAs "bootrec_fixmbr.txt" -IgnoreExit
        if (-not $r.Success) { $allSuccess = $false }
    } else {
        Write-Log "  /fixmbr skipped by user." -Level INFO
    }

    # Step 2: Fix the boot sector
    Write-Log "Step 2/5: Fixing boot sector with bootrec /fixboot..." -Level STEP
    Write-Log "  NOTE: On UEFI systems this may return 'Access is denied' - that is expected." -Level INFO
    $r = Invoke-LoggedCommand -Command 'bootrec' -Arguments @('/fixboot') `
        -Description "bootrec /fixboot" -SaveRawAs "bootrec_fixboot.txt" -IgnoreExit
    $fixbootOut = ($r.Output -join ' ').ToLower()
    if ($fixbootOut -match 'access is denied') {
        Write-Log "  bootrec /fixboot returned 'Access is denied' - normal on UEFI. Continuing..." -Level WARN
    } elseif (-not $r.Success) {
        $allSuccess = $false
        Write-Log "  bootrec /fixboot failed with exit code $($r.ExitCode)" -Level WARN
    }

    # Step 3: Scan for OS
    Write-Log "Step 3/5: Scanning for OS entries with bootrec /scanos..." -Level STEP
    Invoke-LoggedCommand -Command 'bootrec' -Arguments @('/scanos') `
        -Description "bootrec /scanos" -SaveRawAs "bootrec_scanos2.txt" -IgnoreExit | Out-Null

    # Step 4: Rebuild BCD
    Write-Log "Step 4/5: Rebuilding BCD with bootrec /rebuildbcd..." -Level STEP
    Write-Log "  This will scan all disks and offer to add OS entries to the BCD." -Level INFO
    Write-Log "  Answer 'Yes' (Y) for each valid Windows installation found." -Level INFO
    $r = Invoke-LoggedCommand -Command 'bootrec' -Arguments @('/rebuildbcd') `
        -Description "bootrec /rebuildbcd" -SaveRawAs "bootrec_rebuildbcd.txt" -IgnoreExit

    if (-not $r.Success) {
        Write-Log "  bootrec /rebuildbcd failed. Attempting bcdboot as fallback..." -Level WARN
        $bcdBootFallback = Invoke-LoggedCommand -Command 'bcdboot' `
            -Arguments @($winDir, '/f', 'BIOS') `
            -Description "bcdboot /f BIOS (fallback)" -SaveRawAs "bcdboot_bios.txt" -IgnoreExit
        if ($bcdBootFallback.Success) {
            Write-Log "  bcdboot BIOS fallback succeeded." -Level SUCCESS
        } else {
            $allSuccess = $false
            Write-Log "  Both rebuildbcd and bcdboot fallback failed." -Level ERROR
        }
    } else {
        Write-Log "  bootrec /rebuildbcd completed." -Level SUCCESS
    }

    # Step 5: bcdboot to ensure boot files are in place
    Write-Log "Step 5/5: Ensuring boot files with bcdboot $winDir /f BIOS..." -Level STEP
    $bcdBoot = Invoke-LoggedCommand -Command 'bcdboot' `
        -Arguments @($winDir, '/f', 'BIOS') `
        -Description "bcdboot BIOS final" -SaveRawAs "bcdboot_bios_final.txt" -IgnoreExit

    if ($bcdBoot.Success -or $allSuccess) {
        Write-Log "BIOS/MBR boot repair completed." -Level SUCCESS
        $Script:RepairDone = $true
        return $true
    }

    Write-Log "BIOS/MBR boot repair completed with some warnings. Check logs." -Level WARN
    $Script:RepairDone = $true
    return $false
}

# ============================================================
# SECTION 9: CHKDSK
# ============================================================

function Run-Chkdsk {
    <#
    .SYNOPSIS Runs chkdsk on the selected Windows volume.
    .NOTES In WinRE the volume should be offline/unmounted for best results.
    #>
    Write-Log "Running CHKDSK on Windows volume..." -Level SECTION

    if (-not $Script:SelectedWin) {
        Write-Log "No Windows installation selected." -Level ERROR
        return $false
    }

    $drive = $Script:SelectedWin.Drive
    $driveLetter = $drive.TrimEnd('\')

    Write-Host "`n[INFO] CHKDSK will check the file system on $driveLetter" -ForegroundColor Cyan
    Write-Host "  /F  = Fix errors automatically" -ForegroundColor Cyan
    Write-Host "  /R  = Locate bad sectors and recover data (adds significant time)" -ForegroundColor Cyan
    Write-Host ""

    $useR = Get-UserConfirmation "Include /R (bad sector scan)? Adds 30-90 minutes. [Y/N]"
    $chkArgs = @($driveLetter)
    if ($useR) {
        $chkArgs += '/R'
        Write-Log "CHKDSK will run with /F /R on $driveLetter" -Level INFO
    } else {
        $chkArgs += '/F'
        Write-Log "CHKDSK will run with /F on $driveLetter" -Level INFO
    }

    Write-Log "Running: chkdsk $($chkArgs -join ' ')" -Level STEP
    $result = Invoke-LoggedCommand -Command 'chkdsk' -Arguments $chkArgs `
        -Description "CHKDSK on $driveLetter" -SaveRawAs "chkdsk_output.txt" -IgnoreExit

    Write-Log "CHKDSK exit code: $($result.ExitCode)" -Level INFO

    # CHKDSK exit codes: 0=no errors, 1=errors fixed, 2=cleanup needed, 3=errors not fixed
    switch ($result.ExitCode) {
        0 { Write-Log "CHKDSK: No errors found." -Level SUCCESS }
        1 { Write-Log "CHKDSK: Errors found and fixed." -Level SUCCESS }
        2 { Write-Log "CHKDSK: Disk cleanup needs to be performed." -Level WARN }
        3 { Write-Log "CHKDSK: Errors found but not all were fixed." -Level ERROR }
        default { Write-Log "CHKDSK: Exit code $($result.ExitCode) - review output." -Level WARN }
    }

    return ($result.ExitCode -le 1)
}

# ============================================================
# SECTION 10: OFFLINE SFC
# ============================================================

function Run-OfflineSfc {
    <#
    .SYNOPSIS Runs System File Checker in offline mode against the selected Windows installation.
    .NOTES  sfc /scannow /offbootdir and /offwindir are required in WinRE.
            Full scan can take 20-60 minutes.
    #>
    Write-Log "Running offline SFC (System File Checker)..." -Level SECTION

    if (-not $Script:SelectedWin) {
        Write-Log "No Windows installation selected." -Level ERROR
        return $false
    }

    $winDir  = $Script:SelectedWin.WindowsDir
    $bootDir = $Script:SelectedWin.Drive

    Write-Host "`n[INFO] Offline SFC will verify and repair Windows system files." -ForegroundColor Cyan
    Write-Host "       This can take 20-60 minutes depending on disk speed." -ForegroundColor Cyan
    Write-Host "       Windows directory: $winDir" -ForegroundColor Cyan
    Write-Host "       Boot directory:    $bootDir" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "[WARN] SFC requires the Windows installation to be offline (not the running OS)." -ForegroundColor Yellow
    Write-Host "       In WinRE/WinPE, this is the correct mode." -ForegroundColor Yellow
    Write-Host ""

    $confirm = Get-UserConfirmation "Proceed with offline SFC? This may take a long time. [Y/N]"
    if (-not $confirm) {
        Write-Log "Offline SFC skipped by user." -Level INFO
        return $false
    }

    $sfcPath = Get-Command 'sfc' -ErrorAction SilentlyContinue
    if (-not $sfcPath) {
        Write-Log "sfc.exe not found in this WinPE environment. Cannot run SFC." -Level ERROR
        return $false
    }

    Write-Log "Running: sfc /scannow /offbootdir=$bootDir\ /offwindir=$winDir" -Level STEP
    $result = Invoke-LoggedCommand -Command 'sfc' `
        -Arguments @('/scannow', "/offbootdir=$bootDir\", "/offwindir=$winDir") `
        -Description "Offline SFC scan" -SaveRawAs "sfc_offline.txt" -IgnoreExit

    Write-Log "SFC exit code: $($result.ExitCode)" -Level INFO

    $sfcLogPath = Join-Path $winDir 'Logs\CBS\CBS.log'
    if (Test-Path $sfcLogPath -ErrorAction SilentlyContinue) {
        Write-Log "SFC/CBS log available at: $sfcLogPath" -Level INFO
        try {
            $cbsDest = Join-Path $Script:RawLogDir 'CBS.log'
            Copy-Item $sfcLogPath $cbsDest -Force -ErrorAction Stop
            Write-Log "CBS.log copied to: $cbsDest" -Level SUCCESS
        } catch {
            Write-Log "Could not copy CBS.log: $_" -Level WARN
        }
    }

    $output = $result.Output -join ' '
    if ($output -match 'did not find any integrity violations') {
        Write-Log "SFC: No integrity violations found." -Level SUCCESS
    } elseif ($output -match 'successfully repaired') {
        Write-Log "SFC: Corrupt files were found and repaired." -Level SUCCESS
    } elseif ($output -match 'unable to fix') {
        Write-Log "SFC: Some files could not be repaired. Recommend running DISM /RestoreHealth next." -Level ERROR
    } else {
        Write-Log "SFC: Scan completed. Review output above for details." -Level INFO
    }

    return $result.Success
}

# ============================================================
# SECTION 10b: DISM OFFLINE REPAIR
# ============================================================

function Run-DismRepair {
    <#
    .SYNOPSIS Runs DISM offline image repair against the selected Windows installation.
    .NOTES Requires either a WIM/ESD source on the USB or internet access via Windows Update.
            Can repair system image corruption that SFC cannot fix.
            Can take 30-60+ minutes.
    #>
    Write-Log "Running DISM offline repair (RestoreHealth)..." -Level SECTION

    if (-not $Script:SelectedWin) {
        Write-Log "No Windows installation selected." -Level ERROR
        return $false
    }

    $imagePath = $Script:SelectedWin.Drive  # e.g. D:

    Write-Host "`n[INFO] DISM will scan and repair the Windows component store." -ForegroundColor Cyan
    Write-Host "       This can take 30-60 minutes." -ForegroundColor Cyan
    if ($Script:WimSourcePath) {
        Write-Host "       Source file: $Script:WimSourcePath" -ForegroundColor Green
    } else {
        Write-Host "       No local WIM/ESD source found. DISM will attempt Windows Update (requires internet)." -ForegroundColor Yellow
        Write-Host "       In offline WinRE WITHOUT internet, this operation will likely fail." -ForegroundColor Yellow
        Write-Host "       Copy sources\install.wim or install.esd from a Windows ISO to the USB stick root." -ForegroundColor Yellow
    }
    Write-Host ""

    $confirm = Get-UserConfirmation "Proceed with DISM offline repair? [Y/N]"
    if (-not $confirm) {
        Write-Log "DISM repair skipped by user." -Level INFO
        return $false
    }

    $dismCmd = Get-Command 'dism' -ErrorAction SilentlyContinue
    if (-not $dismCmd) {
        Write-Log "dism.exe not found. Cannot run DISM repair." -Level ERROR
        return $false
    }

    # First: scan image health (fast, non-destructive)
    Write-Log "Step 1/3: Scanning image health..." -Level STEP
    $scanResult = Invoke-LoggedCommand -Command 'dism' `
        -Arguments @("/Image:$imagePath", '/Cleanup-Image', '/ScanHealth') `
        -Description "DISM ScanHealth" -SaveRawAs "dism_scanhealth.txt" -IgnoreExit

    $scanOut = ($scanResult.Output -join ' ').ToLower()
    if ($scanOut -match 'no component store corruption detected') {
        Write-Log "DISM ScanHealth: No corruption detected." -Level SUCCESS
    } else {
        Write-Log "DISM ScanHealth: Corruption or issues detected. Proceeding to RestoreHealth." -Level WARN
    }

    # Second: check image health
    Write-Log "Step 2/3: Checking image health..." -Level STEP
    Invoke-LoggedCommand -Command 'dism' `
        -Arguments @("/Image:$imagePath", '/Cleanup-Image', '/CheckHealth') `
        -Description "DISM CheckHealth" -SaveRawAs "dism_checkhealth.txt" -IgnoreExit | Out-Null

    # Third: restore image health
    Write-Log "Step 3/3: Running RestoreHealth..." -Level STEP
    $restoreArgs = @("/Image:$imagePath", '/Cleanup-Image', '/RestoreHealth')
    if ($Script:WimSourcePath) {
        $restoreArgs += "/Source:WIM:$($Script:WimSourcePath):1"
        $restoreArgs += '/LimitAccess'
        Write-Log "  Using local WIM source: $Script:WimSourcePath" -Level INFO
    }

    $restoreResult = Invoke-LoggedCommand -Command 'dism' `
        -Arguments $restoreArgs `
        -Description "DISM RestoreHealth" -SaveRawAs "dism_restorehealth.txt" -IgnoreExit

    $restoreOut = ($restoreResult.Output -join ' ').ToLower()
    if ($restoreResult.Success -or $restoreOut -match 'the restore operation completed successfully') {
        Write-Log "DISM RestoreHealth completed successfully." -Level SUCCESS
        return $true
    } elseif ($restoreOut -match 'source files could not be found') {
        Write-Log "DISM: Source files not found. Place install.wim/install.esd in the USB root and retry." -Level ERROR
    } else {
        Write-Log "DISM RestoreHealth failed (exit $($restoreResult.ExitCode)). Review dism_restorehealth.txt" -Level ERROR
    }

    return $false
}

# ============================================================
# SECTION 11: UTILITY FUNCTIONS
# ============================================================

function Get-UserConfirmation {
    <#
    .SYNOPSIS Prompts user for Y/N confirmation.
    .RETURNS $true if user confirmed, $fals