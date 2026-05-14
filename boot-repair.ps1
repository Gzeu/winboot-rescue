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

.NOTES
    WinRE/WinPE LIMITATIONS:
    - sfc /scannow requires /offbootdir and /offwindir parameters in WinRE
    - wmic may be absent in minimal WinPE; PowerShell CIM cmdlets used as fallback
    - reagentc may not be available in all PE builds
    - Drive letters are assigned dynamically in WinPE; C: IS NOT assumed
    - Some bootrec operations (/fixboot) may fail with "Access is denied" on UEFI systems -- handled explicitly
    - Transcript may fail if the log path is on a read-only volume

    AUTHOR  : Boot Repair Utility v2.0
    COMPAT  : Windows 10/11, WinRE, WinPE x64
    USAGE   : Launch via boot-repair.cmd wrapper, or:
              powershell.exe -ExecutionPolicy Bypass -File boot-repair.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================
# SECTION 0: GLOBAL STATE
# ============================================================
$Script:LogDir          = $null
$Script:LogFile         = $null
$Script:RawLogDir       = $null
$Script:ScriptDrive     = Split-Path -Qualifier $MyInvocation.MyCommand.Path
$Script:TranscriptPath  = $null
$Script:BootMode        = $null   # 'UEFI' or 'BIOS'
$Script:PartStyle       = $null   # 'GPT' or 'MBR'
$Script:WinInstalls     = @()     # Array of detected Windows installations
$Script:SelectedWin     = $null   # Selected Windows installation object
$Script:EfiDrive        = $null   # Letter assigned to EFI partition (may be temporary)
$Script:EfiTempLetter   = $null   # If we assigned a temp letter, store it here for cleanup
$Script:SystemDisk      = $null   # Disk number hosting Windows
$Script:DiagComplete    = $false
$Script:RepairDone      = $false
# FIX Bug3: track whether we are in WinPE/WinRE so the script-drive skip applies only there
$Script:IsWinPE         = $false

# ============================================================
# SECTION 1: LOGGING INFRASTRUCTURE
# ============================================================

function Initialize-Logging {
    <#
    .SYNOPSIS Initializes the log directory on the USB drive and starts transcript.
    #>

    # Determine log root -- prefer the drive this script is running from
    $logRoot = Join-Path $Script:ScriptDrive "Logs"
    try {
        if (-not (Test-Path $logRoot)) {
            New-Item -ItemType Directory -Path $logRoot -Force | Out-Null
        }
        $timestamp             = Get-Date -Format 'yyyyMMdd_HHmmss'
        $Script:LogDir         = $logRoot
        $Script:LogFile        = Join-Path $logRoot "boot-repair_$timestamp.log"
        $Script:RawLogDir      = Join-Path $logRoot "raw_$timestamp"
        $Script:TranscriptPath = Join-Path $logRoot "transcript_$timestamp.txt"
        New-Item -ItemType Directory -Path $Script:RawLogDir -Force | Out-Null

        # Start PowerShell transcript for full session capture
        try { Start-Transcript -Path $Script:TranscriptPath -Append | Out-Null } catch { <# transcript not critical #> }

        Write-Log "Windows Boot Repair Utility v2.0"
        Write-Log "Session started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        Write-Log "Script location: $Script:ScriptDrive"
        Write-Log "Log directory  : $Script:LogDir"
        Write-Log "Log file       : $Script:LogFile"
        Write-Log "Raw logs       : $Script:RawLogDir"
    }
    catch {
        # Fallback log to TEMP if USB is not writable
        $fallback = Join-Path $env:TEMP "BootRepair"
        New-Item -ItemType Directory -Path $fallback -Force -ErrorAction SilentlyContinue | Out-Null
        $Script:LogDir    = $fallback
        $Script:LogFile   = Join-Path $fallback "boot-repair.log"
        $Script:RawLogDir = Join-Path $fallback "raw"
        New-Item -ItemType Directory -Path $Script:RawLogDir -Force -ErrorAction SilentlyContinue | Out-Null
        Write-Host "[WARN] Could not create log on USB ($logRoot). Falling back to $fallback" -ForegroundColor Yellow
    }
}

function Write-Log {
    <#
    .SYNOPSIS Writes a timestamped entry to both the console and the log file.
    .PARAMETER Message The message to log.
    .PARAMETER Level   INFO (default), WARN, ERROR, SUCCESS, STEP, RAW
    .PARAMETER NoConsole Suppress console output for verbose raw data.
    #>
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','STEP','RAW','SECTION')]
        [string]$Level = 'INFO',
        [switch]$NoConsole
    )

    $ts   = Get-Date -Format 'HH:mm:ss'
    $line = "[$ts][$Level] $Message"

    # Write to log file always
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
    .SYNOPSIS    Executes a command, captures all output, logs it, and returns the output.
    .PARAMETER   Command      The executable name (e.g. 'bootrec').
    .PARAMETER   Arguments    Array of argument strings.
    .PARAMETER   Description  Human-readable description shown in logs.
    .PARAMETER   SaveRawAs    If provided, raw output is saved to a file in RawLogDir.
    .PARAMETER   IgnoreExit   If set, non-zero exit codes do not throw.
    .RETURNS     Object with .Output (string[]), .ExitCode (int), .Success (bool)

    FIX Bug1+Bug4: blank/whitespace-only lines from command output (diskpart, bcdedit,
    mountvol etc. emit decorative blank lines) are filtered BEFORE Write-Log and before
    console print.  This prevents the 'Cannot bind argument to parameter Message because
    it is an empty string' StrictMode exception.
    #>
    param(
        [Parameter(Mandatory)][string]$Command,
        [string[]]$Arguments    = @(),
        [string]$Description    = '',
        [string]$SaveRawAs      = '',
        [switch]$IgnoreExit
    )

    $desc = if ($Description) { $Description } else { "$Command $($Arguments -join ' ')" }
    Write-Log "Executing: $desc" -Level STEP

    # Check command availability
    $cmdPath = Get-Command $Command -ErrorAction SilentlyContinue
    if (-not $cmdPath) {
        Write-Log "Command not found: $Command -- skipping." -Level WARN
        return [PSCustomObject]@{ Output = @(); ExitCode = -1; Success = $false; Skipped = $true }
    }

    $result = [PSCustomObject]@{ Output = @(); ExitCode = 0; Success = $false; Skipped = $false }

    try {
        # Capture stdout+stderr; convert everything to string first
        $rawOutput       = & $Command $Arguments 2>&1
        $result.ExitCode = $LASTEXITCODE

        # FIX Bug1+Bug4: convert to strings then drop blank/whitespace-only lines
        $allLines        = @($rawOutput | ForEach-Object { "$_" })
        $result.Output   = @($allLines  | Where-Object { $_.Trim() -ne '' })
        $result.Success  = $result.ExitCode -eq 0

        # Log non-blank lines only
        foreach ($line in $result.Output) {
            Write-Log $line -Level RAW -NoConsole:($result.Output.Count -gt 50)
        }

        # Print abbreviated output to console if large
        if ($result.Output.Count -gt 50) {
            Write-Host "  [Output: $($result.Output.Count) lines -- see log]" -ForegroundColor DarkGray
        } else {
            # FIX Bug4: only print non-blank lines to console as well
            $result.Output | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }
        }

        if ($result.ExitCode -ne 0 -and -not $IgnoreExit) {
            Write-Log "Command exited with code $($result.ExitCode): $desc" -Level WARN
        } elseif ($result.Success) {
            Write-Log "Command completed successfully: $desc" -Level SUCCESS
        }
    }
    catch {
        $result.ExitCode = -1
        $result.Output   = @("EXCEPTION: $_")
        Write-Log "Exception running '$Command': $_" -Level ERROR
    }

    Write-Log "Exit code: $($result.ExitCode)" -Level INFO

    if ($SaveRawAs -and $Script:RawLogDir) {
        $rawPath = Join-Path $Script:RawLogDir $SaveRawAs
        $result.Output | Set-Content -Path $rawPath -Encoding UTF8 -ErrorAction SilentlyContinue
    }

    return $result
}

function Save-RawOutput {
    <#
    .SYNOPSIS Saves arbitrary string content to a raw log file.
    #>
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

    # X: drive is the WinPE ramdisk in most scenarios
    if (Test-Path "X:\Windows\System32" -ErrorAction SilentlyContinue) { $isWinPE = $true }
    if (Test-Path "X:\Sources"          -ErrorAction SilentlyContinue) { $isWinPE = $true }

    # WinRE marker
    if (Test-Path "HKLM:\SYSTEM\CurrentControlSet\Control\MiniNT" -ErrorAction SilentlyContinue) {
        $isWinPE = $true
        $isWinRE = $true
    }
    if ($env:WINPE -eq '1') { $isWinPE = $true }

    # FIX Bug3: store WinPE state globally so Get-WindowsInstallations can use it
    $Script:IsWinPE = $isWinPE

    if      ($isWinRE) { Write-Log "Environment: Windows Recovery Environment (WinRE)" -Level SUCCESS }
    elseif  ($isWinPE) { Write-Log "Environment: Windows Preinstallation Environment (WinPE)" -Level SUCCESS }
    else               {
        Write-Log "Environment: Standard Windows session (not WinRE/WinPE)" -Level WARN
        Write-Log "WARNING: Some repairs work best from WinRE. Continue at your own risk." -Level WARN
    }

    Write-Log "Script running from drive: $Script:ScriptDrive"
    Write-Log "PowerShell version: $($PSVersionTable.PSVersion)"
    Write-Log "OS version: $([System.Environment]::OSVersion.VersionString)"

    return [PSCustomObject]@{ IsWinPE = $isWinPE; IsWinRE = $isWinRE; UsbDrive = $Script:ScriptDrive }
}

# ============================================================
# SECTION 3: HARDWARE DETECTION
# ============================================================

function Get-BootMode {
    <#
    .SYNOPSIS Detects whether the system firmware is UEFI or Legacy BIOS.
    .NOTES    Uses multiple detection methods for reliability in WinPE.
    #>
    Write-Log "Detecting firmware/boot mode (UEFI vs BIOS)..." -Level STEP

    $bootMode = 'BIOS'   # default assumption

    # Method 1: Check HKLM (may be minimal in WinPE)
    try {
        $uefiKey = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State' -ErrorAction SilentlyContinue
        if ($uefiKey) { $bootMode = 'UEFI' }
    } catch {}

    # Method 2: bcdedit /enum firmware is UEFI-only and will fail on BIOS
    try {
        $bcdTest = & bcdedit.exe /enum firmware 2>&1
        if ($LASTEXITCODE -eq 0 -and ($bcdTest -match 'firmware')) { $bootMode = 'UEFI' }
    } catch {}

    # Method 3: Win32 API GetFirmwareEnvironmentVariable -- most reliable but requires elevation
    # Error 998 (ERROR_NOACCESS) = UEFI, Error 1 (ERROR_INVALID_FUNCTION) = BIOS
    try {
        $sig = @'
[DllImport("kernel32.dll", SetLastError=true)]
public static extern uint GetFirmwareEnvironmentVariableA(string lpName, string lpGuid, System.IntPtr pBuffer, uint nSize);
'@
        $type = Add-Type -MemberDefinition $sig -Name 'FirmwareCheck' -Namespace 'Win32' -PassThru -ErrorAction SilentlyContinue
        if ($type) {
            $null   = $type::GetFirmwareEnvironmentVariableA("", "{00000000-0000-0000-0000-000000000000}", [System.IntPtr]::Zero, 0)
            $err    = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
            if      ($err -eq 998) { $bootMode = 'UEFI' }
            elseif  ($err -eq 1)   { $bootMode = 'BIOS' }
        }
    } catch {}

    $Script:BootMode = $bootMode
    Write-Log "Detected firmware mode: $bootMode" -Level SUCCESS
    return $bootMode
}

function Get-DiskInventory {
    <#
    .SYNOPSIS Enumerates all disks, volumes, and identifies partition style.
    .RETURNS  Array of disk objects with partition details.
    #>
    Write-Log "Enumerating disks and volumes..." -Level SECTION

    $disks = @()

    try {
        # Use Get-Disk if available (not always in minimal WinPE)
        $diskObjects = Get-Disk -ErrorAction SilentlyContinue
        if ($diskObjects) {
            foreach ($disk in $diskObjects) {
                $partitions = Get-Partition -DiskNumber $disk.DiskNumber -ErrorAction SilentlyContinue
                $volumes    = @()
                foreach ($part in $partitions) {
                    $vol      = Get-Volume -Partition $part -ErrorAction SilentlyContinue
                    $volumes += [PSCustomObject]@{
                        PartitionNumber = $part.PartitionNumber
                        DriveLetter     = $part.DriveLetter
                        Size            = [math]::Round($part.Size/1GB, 2)
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
                    Size              = [math]::Round($disk.Size/1GB, 2)
                    PartitionStyle    = $disk.PartitionStyle
                    OperationalStatus = $disk.OperationalStatus
                    Partitions        = $volumes
                }
                $disks += $diskObj
                Write-Log "  Disk $($disk.DiskNumber): $($disk.FriendlyName) | $($disk.PartitionStyle) | $([math]::Round($disk.Size/1GB,1)) GB | $($disk.OperationalStatus)"
                foreach ($v in $volumes) {
                    $letter = if ($v.DriveLetter) { $v.DriveLetter } else { '(no letter)' }
                    Write-Log "    Part $($v.PartitionNumber): $letter | $($v.FileSystem) | $([math]::Round($v.Size,2)) GB | Type=$($v.Type) | GptType=$($v.GptType)"
                }
            }

            # Determine overall partition style from first non-RAW disk
            $mainDisk = $disks | Where-Object { $_.PartitionStyle -ne 'RAW' } | Select-Object -First 1
            if ($mainDisk) {
                $Script:PartStyle = $mainDisk.PartitionStyle
                Write-Log "Primary partition style: $Script:PartStyle" -Level SUCCESS
            } else {
                throw "Get-Disk returned no results"
            }
        }
    }
    catch {
        Write-Log "Get-Disk failed ($_) -- falling back to CIM/WMI..." -Level WARN
        try {
            $cimDisks = Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue
            foreach ($cd in $cimDisks) {
                Write-Log ("  Disk: " + $cd.DeviceID + " | " + $cd.Model + " | " + [math]::Round($cd.Size/1GB,1) + " GB")
                $disks += [PSCustomObject]@{
                    DiskNumber     = $cd.Index; Model = $cd.Model
                    Size           = [math]::Round($cd.Size/1GB,2); PartitionStyle = 'Unknown'
                    Partitions     = @()
                }
            }
        }
        catch { Write-Log "CIM disk enumeration also failed" -Level ERROR }
    }

    # Always run diskpart for raw log -- reliable in all PE variants
    $dpScript = "list disk`r`nlist volume`r`nlist partition`r`n"
    $dpFile   = Join-Path $env:TEMP "dplist.txt"
    $dpScript | Set-Content $dpFile -Encoding ASCII -ErrorAction SilentlyContinue
    $dpResult = Invoke-LoggedCommand -Command 'diskpart' -Arguments @('/s', $dpFile) `
        -Description "diskpart list disk/vol/partition" -SaveRawAs "diskpartlist.txt" -IgnoreExit
    Remove-Item $dpFile -Force -ErrorAction SilentlyContinue

    return $disks
}

function Get-EfiPartition {
    <#
    .SYNOPSIS Locates the EFI System Partition, assigns a drive letter if needed.
    .RETURNS  Object with DriveLetter, DiskNumber, PartitionNumber, and NeedsCleanup (bool).
    #>
    Write-Log "Locating EFI System Partition..." -Level STEP

    $efiInfo = [PSCustomObject]@{
        DriveLetter     = $null
        DiskNumber      = $null
        PartitionNumber = $null
        NeedsCleanup    = $false
        Found           = $false
    }

    # EFI GPT partition type GUID
    $efiGuid = 'c12a7328-f81f-11d2-ba4b-00a0c93ec93b'

    try {
        $allDisks = Get-Disk -ErrorAction SilentlyContinue
        foreach ($disk in $allDisks) {
            $parts = Get-Partition -DiskNumber $disk.DiskNumber -ErrorAction SilentlyContinue
            foreach ($p in $parts) {
                $gpt = $p.GptType.ToLower()
                if ($gpt -eq $efiGuid -or $p.Type -eq 'System') {
                    $efiInfo.DiskNumber      = $disk.DiskNumber
                    $efiInfo.PartitionNumber = $p.PartitionNumber
                    $efiInfo.DriveLetter     = $p.DriveLetter
                    $efiInfo.Found           = $true
                    Write-Log "Found EFI partition: Disk=$($disk.DiskNumber), Partition=$($p.PartitionNumber), Type=$($p.Type)"

                    # If no drive letter, assign a temporary one
                    if (-not $p.DriveLetter) {
                        Write-Log "EFI partition has no drive letter. Attempting to assign one..." -Level WARN
                        $freeLetter = Get-AvailableDriveLetter
                        if ($freeLetter) {
                            try {
                                $dpAssign = "select disk $($disk.DiskNumber)`r`nselect partition $($p.PartitionNumber)`r`nassign letter=$freeLetter`r`n"
                                $dpFile   = Join-Path $env:TEMP "dpassign.txt"
                                $dpAssign | Set-Content $dpFile -Encoding ASCII
                                $r = Invoke-LoggedCommand -Command 'diskpart' -Arguments @('/s', $dpFile) `
                                    -Description "Assign letter $freeLetter to EFI partition" -IgnoreExit
                                Remove-Item $dpFile -Force -ErrorAction SilentlyContinue
                                if ($r.Success -or ($r.Output -join ' ') -match 'successfully') {
                                    $efiInfo.DriveLetter     = $freeLetter
                                    $efiInfo.NeedsCleanup    = $true
                                    $Script:EfiTempLetter    = $freeLetter
                                    Write-Log "Assigned temporary letter $freeLetter to EFI partition." -Level SUCCESS
                                }
                            }
                            catch { Write-Log "Could not assign letter to EFI partition: $_" -Level ERROR }
                        }
                    }
                    break
                }
            }
            if ($efiInfo.Found) { break }
        }
    }
    catch {
        Write-Log "Error enumerating EFI partition via PowerShell: $_ -- trying diskpart..." -Level WARN
        # Fallback: parse diskpart output
        $dpScript = "list disk`r`n"
        $dpFile   = Join-Path $env:TEMP "dpefi.txt"
        $dpScript | Set-Content $dpFile -Encoding ASCII
        Invoke-LoggedCommand -Command 'diskpart' -Arguments @('/s', $dpFile) -IgnoreExit | Out-Null
        Remove-Item $dpFile -Force -ErrorAction SilentlyContinue
    }

    if (-not $efiInfo.Found) {
        Write-Log "EFI System Partition NOT found. System may be BIOS/MBR or EFI partition is damaged." -Level WARN
    } else {
        $Script:EfiDrive = $efiInfo.DriveLetter
        Write-Log "EFI partition letter: $($efiInfo.DriveLetter)" -Level SUCCESS
    }

    return $efiInfo
}

function Remove-TempEfiLetter {
    <#
    .SYNOPSIS Removes the temporary drive letter assigned to the EFI partition.
    #>
    if ($Script:EfiTempLetter) {
        Write-Log "Removing temporary EFI drive letter: $Script:EfiTempLetter..." -Level STEP
        $dpScript = "select volume $Script:EfiTempLetter`r`nremove letter=$Script:EfiTempLetter`r`n"
        $dpFile   = Join-Path $env:TEMP "dpremove.txt"
        $dpScript | Set-Content $dpFile -Encoding ASCII
        Invoke-LoggedCommand -Command 'diskpart' -Arguments @('/s', $dpFile) `
            -Description "Remove temp EFI letter $Script:EfiTempLetter" -IgnoreExit | Out-Null
        Remove-Item $dpFile -Force -ErrorAction SilentlyContinue
        $Script:EfiTempLetter = $null
    }
}

function Get-AvailableDriveLetter {
    <#
    .SYNOPSIS Returns the first unused drive letter (starting from S: to avoid conflicts).
    #>
    $used = [System.IO.DriveInfo]::GetDrives() | ForEach-Object { $_.Name[0] }
    foreach ($l in @('S','T','U','V','W','X','Y','Z','R','Q','P','O','N','M','L','K','J','I')) {
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
    .NOTES    Does NOT assume C:. Checks for all key Windows files.

    FIX Bug3: The script-drive exclusion only applies when running in WinPE/WinRE.
    In a standard Windows session the script may live on C: which is also where
    Windows lives -- skipping it would mean never finding any installation.
    In WinPE the script runs from the USB (e.g. E:) so excluding it is correct.
    #>
    Write-Log "Scanning for Windows installations on all accessible drives..." -Level SECTION

    $installations = @()

    # Get all available drive letters
    $drives = [System.IO.DriveInfo]::GetDrives() |
        Where-Object { $_.DriveType -in @('Fixed','Removable') -and $_.IsReady } |
        ForEach-Object { $_.Name.TrimEnd('\') }

    Write-Log "Checking drives: $($drives -join ', ')"

    foreach ($drive in $drives) {
        # FIX Bug3: skip script drive only when running inside WinPE/WinRE
        # In a normal Windows session the USB and Windows share the same drive space
        # so we must not exclude based on script location.
        if ($Script:IsWinPE -and ($drive -eq $Script:ScriptDrive.TrimEnd('\'))) {
            Write-Log "  Skipping script drive (WinPE mode): $drive"
            continue
        }

        $systemConfig = Join-Path $drive "Windows\System32\Config\SYSTEM"
        $explorer     = Join-Path $drive "Windows\explorer.exe"
        $winloadEfi   = Join-Path $drive "Windows\System32\winload.efi"
        $winloadExe   = Join-Path $drive "Windows\System32\winload.exe"
        $ntoskrnl     = Join-Path $drive "Windows\System32\ntoskrnl.exe"

        $hasConfig   = Test-Path $systemConfig -ErrorAction SilentlyContinue
        $hasExplorer = Test-Path $explorer     -ErrorAction SilentlyContinue
        $hasWinload  = (Test-Path $winloadEfi  -ErrorAction SilentlyContinue) -or
                       (Test-Path $winloadExe  -ErrorAction SilentlyContinue)
        $hasNtoskrnl = Test-Path $ntoskrnl     -ErrorAction SilentlyContinue

        # Required files to confirm a valid Windows installation
        if ($hasConfig -and $hasExplorer -and $hasNtoskrnl) {
            $winVer   = 'Unknown'
            $buildNum = ''

            try {
                $hivePath = Join-Path $drive "Windows\System32\Config\SOFTWARE"
                if (Test-Path $hivePath) {
                    # Determine Windows version from registry hive -- offline mount attempt
                    $hiveKey = 'HKLM\BOOTCHECK'
                    & reg.exe load $hiveKey $hivePath 2>&1 | Out-Null
                    if ($LASTEXITCODE -eq 0) {
                        $psKey   = "Registry::$hiveKey\Microsoft\Windows NT\CurrentVersion"
                        $ntProps = Get-ItemProperty $psKey -ErrorAction SilentlyContinue
                        if ($ntProps) {
                            $winVer   = $ntProps.ProductName
                            $buildNum = "$($ntProps.CurrentBuildNumber).$($ntProps.UBR)"
                        }
                        & reg.exe unload $hiveKey 2>&1 | Out-Null
                    }
                }
            } catch { <# version detection not critical #> }

            $winBootType = if ($hasWinload -and (Test-Path $winloadEfi -ErrorAction SilentlyContinue)) { 'UEFI' }
                           elseif ($hasWinload) { 'BIOS' }
                           else { 'Unknown' }

            $install = [PSCustomObject]@{
                Drive          = $drive
                WindowsDir     = Join-Path $drive "Windows"
                Version        = $winVer
                Build          = $buildNum
                BootType       = $winBootType
                HasWinloadEfi  = Test-Path $winloadEfi -ErrorAction SilentlyContinue
                HasWinloadExe  = Test-Path $winloadExe -ErrorAction SilentlyContinue
            }
            $installations += $install
            Write-Log "  FOUND Windows at $drive`: $winVer | Build $buildNum | BootType=$winBootType" -Level SUCCESS
        } else {
            Write-Log "  $drive`: Not a valid Windows installation (config=$hasConfig, explorer=$hasExplorer, ntoskrnl=$hasNtoskrnl)"
        }
    }

    if ($installations.Count -eq 0) {
        Write-Log "No valid Windows installation found on any accessible drive!" -Level ERROR
    } else {
        Write-Log "Total Windows installations found: $($installations.Count)" -Level SUCCESS
    }

    $Script:WinInstalls = $installations
    return $installations
}

function Select-WindowsInstallation {
    <#
    .SYNOPSIS Presents an interactive selection menu if multiple installations are detected.
    .RETURNS  Selected Windows installation object, or null if none.
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

    # Multiple installations -- prompt user
    Write-Host "`n$('='*60)" -ForegroundColor White
    Write-Host "  MULTIPLE WINDOWS INSTALLATIONS DETECTED" -ForegroundColor Yellow
    Write-Host "$('='*60)" -ForegroundColor White

    for ($i = 0; $i -lt $Script:WinInstalls.Count; $i++) {
        $inst = $Script:WinInstalls[$i]
        Write-Host ("  [" + ($i+1) + "] " + $inst.Drive + "\ | " + $inst.Version + " | Build " + $inst.Build + " | " + $inst.BootType) -ForegroundColor Cyan
    }

    do {
        $choice = Read-Host "Select installation (1-$($Script:WinInstalls.Count))"
        $idx    = [int]$choice - 1
    } while ($idx -lt 0 -or $idx -ge $Script:WinInstalls.Count)

    $Script:SelectedWin = $Script:WinInstalls[$idx]
    Write-Log "User selected Windows installation: $($Script:SelectedWin.Drive)" -Level SUCCESS
    return $Script:SelectedWin
}

# ============================================================
# SECTION 5: DIAGNOSTICS COLLECTION
# ============================================================

function Collect-Diagnostics {
    <#
    .SYNOPSIS Runs all diagnostic commands and saves output to raw log files.
    #>
    Write-Log "Collecting full system diagnostics..." -Level SECTION

    # Step 1: Backup BCD
    Backup-BcdStore | Out-Null

    # Step 2: System info
    try {
        $sysinfo = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue |
            Select-Object Manufacturer, Model, TotalPhysicalMemory | Format-List | Out-String
        Save-RawOutput -FileName "systeminfo.txt" -Content $sysinfo
    } catch {}

    # Step 3: BIOS/Firmware info
    try {
        $biosInfo = Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue |
            Select-Object Manufacturer, Version, ReleaseDate, SMBIOSBIOSVersion | Format-List | Out-String
        Save-RawOutput -FileName "biosinfo.txt" -Content $biosInfo
    } catch {}

    # Step 4: mountvol
    Invoke-LoggedCommand -Command 'mountvol' -Arguments @() `
        -Description "Mount points and volumes" -SaveRawAs "mountvol.txt" -IgnoreExit | Out-Null

    # Step 5: Logical disk info via CIM (wmic replacement)
    try {
        $logicalDisks = Get-CimInstance -ClassName Win32_LogicalDisk -ErrorAction SilentlyContinue |
            Select-Object DeviceID, DriveType, Size, FreeSpace, FileSystem, VolumeName |
            Format-Table -AutoSize | Out-String
        Write-Log "Logical disks (CIM):" -Level STEP
        Write-Host $logicalDisks -ForegroundColor DarkGray
        Save-RawOutput -FileName "logicaldisks_cim.txt" -Content $logicalDisks
    } catch {
        Write-Log "CIM logical disk query failed" -Level WARN
        # Fallback to wmic if CIM fails
        Invoke-LoggedCommand -Command 'wmic' `
            -Arguments @('logicaldisk','get','deviceid,size,freespace,filesystem,volumename') `
            -Description "wmic logicaldisk fallback" -SaveRawAs "wmic_logicaldisk.txt" -IgnoreExit | Out-Null
    }

    # Step 6: bcdedit /enum all
    Invoke-LoggedCommand -Command 'bcdedit' -Arguments @('/enum', 'all') `
        -Description "BCD Store -- all entries" -SaveRawAs "bcdedit_enum_all.txt" -IgnoreExit | Out-Null

    # Step 7: bcdedit /enum firmware (UEFI only)
    Invoke-LoggedCommand -Command 'bcdedit' -Arguments @('/enum', 'firmware') `
        -Description "BCD Firmware entries (UEFI)" -SaveRawAs "bcdedit_enum_firmware.txt" -IgnoreExit | Out-Null

    # Step 8: bootrec /scanos (BIOS-relevant but harmless on UEFI)
    Invoke-LoggedCommand -Command 'bootrec' -Arguments @('/scanos') `
        -Description "Scan OS entries (bootrec)" -SaveRawAs "bootrec_scanos.txt" -IgnoreExit | Out-Null

    # Step 9: reagentc /info (if available)
    Invoke-LoggedCommand -Command 'reagentc' -Arguments @('/info') `
        -Description "Windows Recovery Agent info" -SaveRawAs "reagentc_info.txt" -IgnoreExit | Out-Null

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
    Write-Host "  Firmware mode : $Script:BootMode"   -ForegroundColor $(if ($Script:BootMode)    { 'Green' } else { 'Red' })
    Write-Host "  Partition style: $Script:PartStyle" -ForegroundColor $(if ($Script:PartStyle)   { 'Green' } else { 'Yellow' })
    Write-Host "  EFI drive letter: $(if ($Script:EfiDrive) { $Script:EfiDrive } else { 'Not assigned / Not found' })" `
        -ForegroundColor $(if ($Script:EfiDrive) { 'Green' } else { 'Yellow' })
    Write-Host "  Windows found : $($Script:WinInstalls.Count)" `
        -ForegroundColor $(if ($Script:WinInstalls.Count -gt 0) { 'Green' } else { 'Red' })
    if ($Script:SelectedWin) {
        Write-Host ("  Selected install:    " + $Script:SelectedWin.Drive + "\ | " + $Script:SelectedWin.Version) -ForegroundColor Green
    }
    Write-Host "$('='*60)`n" -ForegroundColor White
}

# ============================================================
# SECTION 6: BCD BACKUP
# ============================================================

function Backup-BcdStore {
    <#
    .SYNOPSIS Creates a timestamped backup of the BCD store before any modification.
    .NOTES    Silently skips if log dir not initialized or BCD not found.

    FIX Bug2: Invoke-LoggedCommand is wrapped in try/catch so $exportResult is
    never $null when StrictMode is active.  Previously if the command threw before
    returning, $exportResult was unset and accessing .Success caused a StrictMode
    'property cannot be found' exception.
    #>
    Write-Log "Backing up BCD store..." -Level STEP

    if (-not $Script:LogDir) {
        Write-Log "Log directory not initialized -- cannot back up BCD." -Level WARN
        return $false
    }

    $ts      = Get-Date -Format 'yyyyMMdd_HHmmss'
    $bcdDest = Join-Path $Script:LogDir "BCD_backup_$ts"

    # FIX Bug2: explicit try/catch so $exportResult always has a value after the block
    $exportResult = $null
    try {
        $exportResult = Invoke-LoggedCommand -Command 'bcdedit' `
            -Arguments @('/export', $bcdDest) `
            -Description "BCD export backup" -IgnoreExit
    } catch {
        Write-Log "Invoke-LoggedCommand threw during BCD export: $_" -Level WARN
        $exportResult = [PSCustomObject]@{ Output = @(); ExitCode = -1; Success = $false; Skipped = $false }
    }

    if ($exportResult -and $exportResult.Success) {
        Write-Log "BCD exported to: $bcdDest" -Level SUCCESS
        return $true
    }

    try {
        # Fallback: copy BCD file directly from EFI or system partition
        $bcdPaths = @(
            'C:\Boot\BCD',
            "$(if ($Script:EfiDrive) { $Script:EfiDrive } else { 'S' }):\EFI\Microsoft\Boot\BCD",
            'X:\Boot\BCD'
        )
        foreach ($p in $bcdPaths) {
            if (Test-Path $p -ErrorAction SilentlyContinue) {
                $dest = Join-Path $Script:LogDir "BCD_raw_backup_$ts"
                Copy-Item $p $dest -Force -ErrorAction Stop
                Write-Log "BCD file copied from $p to $dest" -Level SUCCESS
                return $true
            }
        }
    } catch {}

    Write-Log "BCD backup failed -- proceeding without backup (be careful!)" -Level ERROR
    return $false
}

# ============================================================
# SECTION 7: UEFI BOOT REPAIR
# ============================================================

function Repair-UefiBoot {
    <#
    .SYNOPSIS Repairs the UEFI boot environment for a GPT system.
    .NOTES    This is the correct approach for UEFI systems.
              Do NOT use bootrec /fixboot as primary -- it targets MBR/VBR.
    #>
    Write-Log "Starting UEFI/GPT boot repair..." -Level SECTION

    if (-not $Script:SelectedWin) {
        Write-Log "No Windows installation selected. Cannot proceed." -Level ERROR
        return $false
    }

    $winDir = $Script:SelectedWin.WindowsDir   # e.g. D:\Windows

    Write-Log "Target Windows directory: $winDir"

    # Step 1: Locate and mount EFI partition
    Write-Log "Step 1/5: Locating EFI System Partition..." -Level STEP
    $efi = Get-EfiPartition
    if (-not $efi.Found -or -not $efi.DriveLetter) {
        Write-Log "EFI partition found but has no drive letter -- cannot proceed." -Level ERROR
        Write-Log "  Suggestion: Manually assign a letter with diskpart (assign letter=S)." -Level INFO
        return $false
    }

    $efiLetter   = $efi.DriveLetter
    $efiBootPath = "${efiLetter}:\EFI\Microsoft\Boot\BCD"
    $efiBootDir  = "${efiLetter}:\EFI\Microsoft\Boot"

    # Step 2: Verify EFI directory exists
    $efiBootExists = Test-Path $efiBootPath -ErrorAction SilentlyContinue
    if ($efiBootExists) {
        Write-Log "  EFI\Microsoft\Boot directory exists at $efiBootPath" -Level INFO
    } else {
        Write-Log "  EFI\Microsoft\Boot directory does NOT exist -- will create via bcdboot" -Level WARN
    }

    # Step 3: Run bcdboot (PRIMARY repair tool for UEFI)
    Write-Log "Step 2/5: Running bcdboot to rebuild EFI boot files..." -Level STEP
    Write-Log "  bcdboot $winDir /s ${efiLetter}: /f UEFI" -Level INFO
    $bcdbootResult = Invoke-LoggedCommand -Command 'bcdboot' `
        -Arguments @($winDir, '/s', "${efiLetter}:", '/f', 'UEFI') `
        -Description "bcdboot UEFI" -SaveRawAs "bcdboot_uefi.txt" -IgnoreExit

    if ($bcdbootResult.Success) {
        Write-Log "bcdboot /f UEFI completed successfully." -Level SUCCESS
        $Script:RepairDone = $true
        Remove-TempEfiLetter
        return $true
    }

    Write-Log "bcdboot /f UEFI failed (exit $($bcdbootResult.ExitCode)). Trying /f ALL..." -Level WARN

    # Step 4: Fallback -- /f ALL creates both UEFI and BIOS entries
    $bcdbootAll = Invoke-LoggedCommand -Command 'bcdboot' `
        -Arguments @($winDir, '/s', "${efiLetter}:", '/f', 'ALL') `
        -Description "bcdboot ALL fallback" -SaveRawAs "bcdboot_all.txt" -IgnoreExit
    if ($bcdbootAll.Success) {
        Write-Log "bcdboot /f ALL completed successfully." -Level SUCCESS
        $Script:RepairDone = $true
        Remove-TempEfiLetter
        return $true
    }

    # Step 5: Scan for OS entries with bootrec /scanos
    Write-Log "Step 4/5: Scanning for OS entries with bootrec /scanos..." -Level STEP
    $scanResult = Invoke-LoggedCommand -Command 'bootrec' -Arguments @('/scanos') `
        -Description "bootrec /scanos" -SaveRawAs "bootrec_scanos2.txt" -IgnoreExit

    # Step 5b: bootrec /rebuildbcd (use with caution -- interactive)
    Write-Log "Step 5/5: Rebuilding BCD with bootrec /rebuildbcd..." -Level STEP
    Write-Log "  NOTE: This is an interactive command -- it will prompt you to add entries." -Level WARN
    Invoke-LoggedCommand -Command 'bootrec' -Arguments @('/rebuildbcd') `
        -Description "bootrec /rebuildbcd" -SaveRawAs "bootrec_rebuildbcd.txt" -IgnoreExit | Out-Null

    # Step 6: Check if EFI now exists and BCD was created
    if (Test-Path $efiBootPath -ErrorAction SilentlyContinue) {
        Write-Log "BCD file created despite non-zero exit code. Boot may be repaired." -Level WARN
        $Script:RepairDone = $true
        Remove-TempEfiLetter
        return $true
    }

    Write-Log "UEFI boot repair could not fully complete. EFI\Microsoft\Boot\BCD not found." -Level ERROR
    Remove-TempEfiLetter
    return $false
}

# ============================================================
# SECTION 8: BIOS/MBR BOOT REPAIR
# ============================================================

function Repair-BiosBoot {
    <#
    .SYNOPSIS Repairs the BIOS/MBR boot environment for an MBR-partitioned disk.
    .NOTES    bootrec is the correct primary tool here (unlike UEFI where bcdboot leads).
    #>
    Write-Log "Starting BIOS/MBR boot repair..." -Level SECTION

    if (-not $Script:SelectedWin) {
        Write-Log "No Windows installation selected. Cannot proceed." -Level ERROR
        return $false
    }

    $winDir    = $Script:SelectedWin.WindowsDir
    $allSuccess = $true

    # Step 1: Backup BCD
    Write-Log "Step 1/6: Backing up BCD store..." -Level STEP
    Backup-BcdStore | Out-Null

    # Step 2: Fix the Master Boot Record
    Write-Log "Step 2/6: Fixing MBR with bootrec /fixmbr..." -Level STEP
    Write-Log "  This rewrites the Master Boot Record. Safe -- does not affect partitions." -Level INFO
    $fixmbrResult = Invoke-LoggedCommand -Command 'bootrec' -Arguments @('/fixmbr') `
        -Description "bootrec /fixmbr" -SaveRawAs "bootrec_fixmbr.txt" -IgnoreExit
    if (-not $fixmbrResult.Success) {
        Write-Log "bootrec /fixmbr failed. This is unusual -- check disk health." -Level WARN
        $allSuccess = $false
    }

    # Step 3: Fix the Volume Boot Record
    Write-Log "Step 3/6: Fixing Volume Boot Record with bootrec /fixboot..." -Level STEP
    Write-Log "  This rewrites the boot sector on the active partition." -Level INFO
    Write-Log "  NOTE: On UEFI systems this may return 'Access is denied' -- normal." -Level INFO
    $fixbootResult = Invoke-LoggedCommand -Command 'bootrec' -Arguments @('/fixboot') `
        -Description "bootrec /fixboot" -SaveRawAs "bootrec_fixboot.txt" -IgnoreExit

    $fixbootOutput = ($fixbootResult.Output -join ' ')
    if ($fixbootOutput -match 'Access is denied') {
        Write-Log "bootrec /fixboot: Access is denied. Possible causes:" -Level WARN
        Write-Log "  1. This is a UEFI system (bootrec /fixboot targets MBR/VBR)" -Level INFO
        Write-Log "  2. The active partition is not properly marked" -Level INFO
        Write-Log "  3. The boot volume is locked -- try from a different WinPE build" -Level INFO
        $allSuccess = $false
    } elseif (-not $fixbootResult.Success) {
        Write-Log "bootrec /fixboot failed (exit $($fixbootResult.ExitCode))" -Level WARN
        $allSuccess = $false
    }

    # Step 4: Scan for OS entries
    Write-Log "Step 4/5: Scanning for OS entries with bootrec /scanos..." -Level STEP
    Invoke-LoggedCommand -Command 'bootrec' -Arguments @('/scanos') `
        -Description "bootrec /scanos" -SaveRawAs "bootrec_scanos.txt" -IgnoreExit | Out-Null

    # Step 5: Rebuild BCD
    Write-Log "Step 5/5: Rebuilding BCD with bootrec /rebuildbcd..." -Level STEP
    Write-Log "  NOTE: This is an interactive command -- it will prompt you to add entries." -Level WARN
    $rebuildResult = Invoke-LoggedCommand -Command 'bootrec' -Arguments @('/rebuildbcd') `
        -Description "bootrec /rebuildbcd" -SaveRawAs "bootrec_rebuildbcd.txt" -IgnoreExit

    if (-not $rebuildResult.Success) {
        Write-Log "bootrec /rebuildbcd failed. Attempting bcdboot as fallback..." -Level WARN
        # Fallback: bcdboot for BIOS
        $bcdbootResult = Invoke-LoggedCommand -Command 'bcdboot' `
            -Arguments @($winDir, '/f', 'BIOS') `
            -Description "bcdboot /f BIOS fallback" -SaveRawAs "bcdboot_bios.txt" -IgnoreExit
        if ($bcdbootResult.Success) {
            Write-Log "bcdboot /f BIOS succeeded as fallback." -Level SUCCESS
        }
    }

    # Step 6: Use bcdboot to ensure Windows boot entry exists in BCD
    Write-Log "Step 6/6: Ensuring boot entry with bcdboot..." -Level STEP
    $bcdbootFinal = Invoke-LoggedCommand -Command 'bcdboot' `
        -Arguments @($winDir, '/f', 'BIOS') `
        -Description "bcdboot /f BIOS final" -SaveRawAs "bcdboot_bios_final.txt" -IgnoreExit
    if ($bcdbootFinal.Success) {
        Write-Log "Boot entry created/verified via bcdboot." -Level SUCCESS
        $Script:RepairDone = $true
    }

    if ($Script:RepairDone -or $allSuccess) {
        Write-Log "BIOS/MBR boot repair completed." -Level SUCCESS
        return $true
    } else {
        Write-Log "BIOS/MBR repair completed with warnings. Review log files." -Level WARN
        $Script:RepairDone = $true
        return $false
    }
}

# ============================================================
# SECTION 9: CHKDSK
# ============================================================

function Run-Chkdsk {
    <#
    .SYNOPSIS Runs CHKDSK on the Windows volume.
    .NOTES    /F flag requires exclusive access -- volume must be offline or scheduled.
              In WinRE, the Windows volume is offline, so /F should work.
    #>
    Write-Log "Running CHKDSK on Windows volume..." -Level SECTION

    if (-not $Script:SelectedWin) {
        Write-Log "No Windows installation selected." -Level ERROR
        return $false
    }

    $driveLetter = $Script:SelectedWin.Drive.TrimEnd('\')

    Write-Host "`n[INFO] CHKDSK will verify the file system on $driveLetter" -ForegroundColor Cyan
    Write-Host "       /F mode fixes file system errors." -ForegroundColor Cyan
    Write-Host "       /R mode ALSO finds bad sectors (can take HOURS)." -ForegroundColor Cyan
    Write-Host ""
    Write-Host "[WARN] If the volume is in use, CHKDSK may schedule a check on next reboot." -ForegroundColor Yellow
    Write-Host ""

    $runRepair  = Get-UserConfirmation "Run CHKDSK with /F (fix errors)? [Y/N]"
    $runRecover = $false
    if ($runRepair) {
        $runRecover = Get-UserConfirmation "Also run /R (bad sector recovery - VERY SLOW)? [Y/N]"
    }

    $chkArgs = @($driveLetter)
    if ($runRepair)  { $chkArgs += '/F' }
    if ($runRecover) { $chkArgs += '/R' }
    $chkArgs += '/X'   # Force dismount if needed

    Write-Log "Running: chkdsk $($chkArgs -join ' ')" -Level STEP
    $result = Invoke-LoggedCommand -Command 'chkdsk' `
        -Arguments $chkArgs -Description "CHKDSK on $driveLetter" `
        -SaveRawAs "chkdsk_$($driveLetter.TrimEnd(':')).txt" -IgnoreExit

    Write-Log "CHKDSK exit code: $($result.ExitCode)" -Level INFO

    switch ($result.ExitCode) {
        0 { Write-Log "CHKDSK: No errors found." -Level SUCCESS }
        1 { Write-Log "CHKDSK: Errors found and fixed." -Level SUCCESS }
        2 { Write-Log "CHKDSK: Disk cleanup needed (non-critical)." -Level WARN }
        3 { Write-Log "CHKDSK: Could not check disk -- may be in use or corrupted." -Level ERROR }
        default { Write-Log "CHKDSK: Unexpected exit code $($result.ExitCode)." -Level WARN }
    }

    return ($result.ExitCode -le 1)
}

# ============================================================
# SECTION 10: OFFLINE SFC
# ============================================================

function Run-OfflineSfc {
    <#
    .SYNOPSIS Runs System File Checker in offline mode against the selected Windows installation.
    .NOTES  sfc /scannow requires /offbootdir and /offwindir parameters in WinRE.
            Full scan can take 20-60 minutes.
    #>
    Write-Log "Running offline SFC (System File Checker)..." -Level SECTION

    if (-not $Script:SelectedWin) {
        Write-Log "No Windows installation selected." -Level ERROR
        return $false
    }

    $winDir    = $Script:SelectedWin.WindowsDir   # e.g. D:\Windows
    $bootDir   = $Script:SelectedWin.Drive        # e.g. D:

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

    # Check if sfc.exe exists
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

    # Check SFC log for results
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
        Write-Log "SFC: Some files could not be repaired. Manual intervention needed." -Level ERROR
    } else {
        Write-Log "SFC: Scan completed. Review output above for details." -Level INFO
    }

    return $result.Success
}

# ============================================================
# SECTION 11: UTILITY FUNCTIONS
# ============================================================

function Get-UserConfirmation {
    <#
    .SYNOPSIS Prompts user for Y/N confirmation.
    .RETURNS $true if user confirmed, $false otherwise.
    #>
    param([string]$Prompt)
    do {
        $response = Read-Host "`n$Prompt"
    } while ($response -notmatch '^[YyNn]$')
    $confirmed = $response -match '^[Yy]$'
    Write-Log "User responded: $response (to: $Prompt)"
    return $confirmed
}

function Export-Report {
    <#
    .SYNOPSIS Exports a comprehensive summary report of all actions taken.
    #>
    Write-Log "Generating final report..." -Level SECTION

    $reportPath = Join-Path $Script:LogDir "boot-repair-report_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $sb = [System.Text.StringBuilder]::new()

    $null = $sb.AppendLine("="*60)
    $null = $sb.AppendLine("WINDOWS BOOT REPAIR - FULL REPORT")
    $null = $sb.AppendLine("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $null = $sb.AppendLine("="*60)
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("SYSTEM INFORMATION")
    $null = $sb.AppendLine("-"*40)
    $null = $sb.AppendLine("Firmware Mode   : $Script:BootMode")
    $null = $sb.AppendLine("Partition Style : $Script:PartStyle")
    $null = $sb.AppendLine("EFI Drive       : $(if ($Script:EfiDrive) {"$Script:EfiDrive`:"} else {'Not found'})")
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("WINDOWS INSTALLATIONS FOUND: $($Script:WinInstalls.Count)")
    foreach ($inst in $Script:WinInstalls) {
        $null = $sb.AppendLine("  - " + $inst.Drive + "\ | " + $inst.Version + " | Build " + $inst.Build + " | " + $inst.BootType)
    }
    $null = $sb.AppendLine("")
    if ($Script:SelectedWin) {
        $null = $sb.AppendLine("SELECTED INSTALLATION")
        $null = $sb.AppendLine("-"*40)
        $null = $sb.AppendLine("Drive     : $($Script:SelectedWin.Drive)")
        $null = $sb.AppendLine("Windows   : $($Script:SelectedWin.WindowsDir)")
        $null = $sb.AppendLine("Version   : $($Script:SelectedWin.Version)")
        $null = $sb.AppendLine("Build     : $($Script:SelectedWin.Build)")
        $null = $sb.AppendLine("Boot Type : $($Script:SelectedWin.BootType)")
        $null = $sb.AppendLine("")
    }
    $null = $sb.AppendLine("REPAIR ACTIONS")
    $null = $sb.AppendLine("-"*40)
    $null = $sb.AppendLine("Diagnostics Collected : $Script:DiagComplete")
    $null = $sb.AppendLine("Repair Attempted      : $Script:RepairDone")
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("LOG FILES")
    $null = $sb.AppendLine("-"*40)
    $null = $sb.AppendLine("Main log     : $Script:LogFile")
    $null = $sb.AppendLine("Raw logs dir : $Script:RawLogDir")
    $null = $sb.AppendLine("Transcript   : $Script:TranscriptPath")
    $null = $sb.AppendLine("")

    $null = $sb.AppendLine("VERDICT")
    $null = $sb.AppendLine("-"*40)
    $verdict = Get-FinalVerdict
    $null = $sb.AppendLine($verdict)
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("="*60)
    $null = $sb.AppendLine("END OF REPORT")
    $null = $sb.AppendLine("="*60)

    $reportContent = $sb.ToString()

    try {
        Set-Content -Path $reportPath -Value $reportContent -Encoding UTF8
        Write-Log "Report saved: $reportPath" -Level SUCCESS
        Write-Host "`n$reportContent" -ForegroundColor White
    } catch {
        Write-Log "Could not save report: $_" -Level ERROR
        Write-Host "`n$reportContent" -ForegroundColor White
    }

    return $reportPath
}

function Get-FinalVerdict {
    <#
    .SYNOPSIS Returns a verdict string based on the current system state.
    #>
    if ($Script:WinInstalls.Count -eq 0) {
        return "[VERDICT] NO VALID WINDOWS INSTALLATION DETECTED -- possible hardware/storage issue or severely corrupt disk."
    }

    if (-not $Script:RepairDone) {
        return "[VERDICT] DIAGNOSTICS ONLY -- No repair was performed in this session."
    }

    if ($Script:BootMode -eq 'UEFI' -and -not $Script:EfiDrive) {
        return "[VERDICT] EFI PARTITION MISSING OR INACCESSIBLE -- UEFI boot cannot be fully repaired without an accessible EFI System Partition."
    }

    if ($Script:RepairDone) {
        return "[VERDICT] BOOT FILES REBUILT SUCCESSFULLY -- Restart the computer to verify. If boot still fails, run CHKDSK or check hardware."
    }

    return "[VERDICT] REPAIR ATTEMPTED BUT UNCERTAIN -- Review log files for details. Consider running CHKDSK and offline SFC."
}

# ============================================================
# SECTION 12: INTERACTIVE MENU
# ============================================================

function Show-MainMenu {
    <#
    .SYNOPSIS Displays the main interactive repair menu and handles user selection.
    #>
    $continue = $true

    while ($continue) {
        Write-Host "`n$('='*60)" -ForegroundColor White
        Write-Host "  WINDOWS BOOT REPAIR UTILITY v2.0" -ForegroundColor Cyan
        Write-Host "$('='*60)" -ForegroundColor White
        Write-Host "  Firmware : $Script:BootMode    |  Partition: $Script:PartStyle" -ForegroundColor DarkGray
        Write-Host ("  Windows  : " + $Script:WinInstalls.Count + " found   |  Selected : " + $(if ($Script:SelectedWin) {$Script:SelectedWin.Drive} else {"None"})) -ForegroundColor DarkGray
        Write-Host "$('-'*60)" -ForegroundColor DarkGray
        Write-Host "  [1] Collect diagnostics only" -ForegroundColor Yellow
        Write-Host "  [2] Quick repair (auto-detect and repair)" -ForegroundColor Green
        Write-Host "  [3] Full repair (diagnostics + quick repair + CHKDSK)" -ForegroundColor Green
        Write-Host "  [4] Rebuild EFI boot files (UEFI/GPT systems)" -ForegroundColor Cyan
        Write-Host "  [5] BIOS/MBR boot repair (Legacy BIOS systems)" -ForegroundColor Cyan
        Write-Host "  [6] Run CHKDSK on Windows volume" -ForegroundColor Magenta
        Write-Host "  [7] Run offline SFC (System File Checker)" -ForegroundColor Magenta
        Write-Host "  [8] Export full diagnostic report" -ForegroundColor White
        Write-Host "  [9] Re-select Windows installation" -ForegroundColor White
        Write-Host "  [0] Exit" -ForegroundColor Red
        Write-Host "$('='*60)" -ForegroundColor White

        $choice = Read-Host "`nSelect option"
        Write-Log "User selected menu option: $choice"

        switch ($choice) {
            '1' {
                Collect-Diagnostics
                Show-DiagnosticSummary
            }
            '2' {
                if (-not $Script:DiagComplete) { Collect-Diagnostics }
                if (-not $Script:SelectedWin) { Select-WindowsInstallation | Out-Null }
                if ($Script:SelectedWin) {
                    if ($Script:BootMode -eq 'UEFI') {
                        Repair-UefiBoot | Out-Null
                    } else {
                        Repair-BiosBoot | Out-Null
                    }
                    Show-DiagnosticSummary
                } else {
                    Write-Log "Cannot repair -- no Windows installation selected." -Level ERROR
                }
            }
            '3' {
                if (-not $Script:DiagComplete) { Collect-Diagnostics }
                if (-not $Script:SelectedWin) { Select-WindowsInstallation | Out-Null }
                if ($Script:SelectedWin) {
                    if ($Script:BootMode -eq 'UEFI') { Repair-UefiBoot | Out-Null }
                    else { Repair-BiosBoot | Out-Null }
                    Run-Chkdsk | Out-Null
                    Show-DiagnosticSummary
                    Export-Report | Out-Null
                } else {
                    Write-Log "Cannot repair -- no Windows installation selected." -Level ERROR
                }
            }
            '4' {
                if (-not $Script:SelectedWin) { Select-WindowsInstallation | Out-Null }
                Repair-UefiBoot | Out-Null
            }
            '5' {
                if (-not $Script:SelectedWin) { Select-WindowsInstallation | Out-Null }
                Repair-BiosBoot | Out-Null
            }
            '6' {
                if (-not $Script:SelectedWin) { Select-WindowsInstallation | Out-Null }
                Run-Chkdsk | Out-Null
            }
            '7' {
                if (-not $Script:SelectedWin) { Select-WindowsInstallation | Out-Null }
                Run-OfflineSfc | Out-Null
            }
            '8' {
                Export-Report | Out-Null
            }
            '9' {
                $Script:SelectedWin = $null
                Get-WindowsInstallations | Out-Null
                Select-WindowsInstallation | Out-Null
            }
            '0' {
                $continue = $false
                Write-Log "User exited menu." -Level INFO
            }
            default {
                Write-Host "  Invalid option. Please enter 0-9." -ForegroundColor Red
            }
        }
    }
}

# ============================================================
# SECTION 13: MAIN ENTRY POINT
# ============================================================

function Main {
    <#
    .SYNOPSIS Main entry point -- initializes and starts the repair tool.
    #>
    try { Clear-Host } catch {}

    Write-Host @"
$('='*60)
  WINDOWS BOOT REPAIR UTILITY v2.0
  Professional Boot Recovery Tool for WinRE/WinPE
$('='*60)
  !! Run as Administrator !!
  !! From WinRE or WinPE boot USB !!
$('='*60)
"@ -ForegroundColor Cyan

    Initialize-Logging

    # FIX Bug3: Get-Environment must run BEFORE Get-WindowsInstallations
    # so $Script:IsWinPE is set correctly before the drive-skip logic executes
    $env = Get-Environment

    Get-BootMode | Out-Null
    Get-DiskInventory | Out-Null
    Get-WindowsInstallations | Out-Null

    Show-DiagnosticSummary

    if ($Script:WinInstalls.Count -gt 0) {
        Select-WindowsInstallation | Out-Null
    } else {
        Write-Host "`n[CRITICAL] No Windows installation found on any accessible drive." -ForegroundColor Red
        Write-Host "  - Check disk connections" -ForegroundColor Yellow
        Write-Host "  - Drive letters may not be assigned -- run Option 1 for diagnostics" -ForegroundColor Yellow
        Write-Host "  - The disk may have severe corruption or hardware failure" -ForegroundColor Yellow
        Write-Host ""
    }

    Show-MainMenu

    Remove-TempEfiLetter

    $verdict = Get-FinalVerdict
    Write-Host "`n$('='*60)" -ForegroundColor White
    Write-Host $verdict -ForegroundColor $(
        if ($verdict -match 'SUCCESSFULLY') { 'Green' }
        elseif ($verdict -match 'MISSING|NO VALID|HARDWARE') { 'Red' }
        else { 'Yellow' }
    )
    Write-Host "$('='*60)`n" -ForegroundColor White

    Write-Log "Session ended: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level INFO
    Write-Log "Log files saved to: $Script:LogDir" -Level SUCCESS

    try { Stop-Transcript -ErrorAction SilentlyContinue } catch {}

    Write-Host "Log files saved to: $Script:LogDir" -ForegroundColor Green
    Write-Host "Press any key to exit..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# ============================================================
# EXECUTE
# ============================================================
Main
