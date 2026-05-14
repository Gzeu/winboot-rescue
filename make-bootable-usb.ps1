#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    make-bootable-usb.ps1 -- Creates a bootable WinPE/WinRE USB stick and copies
    the winboot-rescue toolkit onto it automatically.

.DESCRIPTION
    This script automates the full process of creating a bootable Windows
    recovery USB drive that includes the winboot-rescue tool:

    1. Detects all USB drives and lets the user select one
    2. Formats the USB as FAT32 (MBR or GPT selectable)
    3. Uses Windows ADK / WinPE if available, OR falls back to:
       - Copying WinRE from the current Windows installation
       - Using DISM + bcdboot to make it bootable
    4. Copies boot-repair.ps1 and boot-repair.cmd onto the USB root
    5. Verifies the result

.NOTES
    === REQUIREMENTS ===
    - Must run as Administrator
    - Windows 10/11 host (not WinPE -- run this on a working PC)
    - A USB drive of at least 1 GB (4 GB+ recommended for WinPE ADK)
    - To use the ADK path: Windows ADK + WinPE Add-on must be installed
      Download: https://learn.microsoft.com/en-us/windows-hardware/get-started/adk-install

    === MODES ===
    Mode A -- ADK/WinPE (full WinPE environment, best)
    Mode B -- WinRE copy from local machine (no ADK needed, uses existing WinRE)
    Mode C -- WinRE ISO mount (if you have a Windows ISO)

    COMPAT  : Windows 10/11 (x64), run from standard Windows session
    AUTHOR  : winboot-rescue toolkit
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================
# GLOBALS
# ============================================================
$Script:ToolkitDir  = $PSScriptRoot          # Where boot-repair.ps1/.cmd live
$Script:SelectedUSB = $null
$Script:LogFile     = $null
$Script:WorkDir     = $null
$Script:MountDir    = $null

# ============================================================
# LOGGING
# ============================================================
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','STEP','SECTION')]
        [string]$Level = 'INFO'
    )
    $ts   = Get-Date -Format 'HH:mm:ss'
    $line = "[$ts][$Level] $Message"
    if ($Script:LogFile) {
        try { Add-Content -Path $Script:LogFile -Value $line -Encoding UTF8 } catch {}
    }
    switch ($Level) {
        'INFO'    { Write-Host $line -ForegroundColor Cyan }
        'WARN'    { Write-Host $line -ForegroundColor Yellow }
        'ERROR'   { Write-Host $line -ForegroundColor Red }
        'SUCCESS' { Write-Host $line -ForegroundColor Green }
        'STEP'    { Write-Host "`n$line" -ForegroundColor Magenta }
        'SECTION' { Write-Host "`n$('='*60)`n$line`n$('='*60)" -ForegroundColor White }
    }
}

function Initialize-WorkDir {
    $Script:WorkDir  = Join-Path $env:TEMP 'WinBootUSB'
    $Script:MountDir = Join-Path $Script:WorkDir 'mount'
    $Script:LogFile  = Join-Path $Script:WorkDir "usb-creator_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    foreach ($d in @($Script:WorkDir, $Script:MountDir)) {
        if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
    }
    Write-Log "Work directory: $Script:WorkDir" -Level INFO
    Write-Log "Log file: $Script:LogFile" -Level INFO
}

# ============================================================
# SECTION 1: USB DRIVE SELECTION
# ============================================================
function Get-UsbDrives {
    <#
    .SYNOPSIS Returns all removable USB disks with size info.
    #>
    Write-Log "Scanning for USB drives..." -Level STEP

    $usbDisks = @()
    try {
        $allDisks = Get-Disk -ErrorAction SilentlyContinue | Where-Object { $_.BusType -eq 'USB' }
        foreach ($d in $allDisks) {
            $volumes = Get-Partition -DiskNumber $d.DiskNumber -ErrorAction SilentlyContinue |
                       ForEach-Object { Get-Volume -Partition $_ -ErrorAction SilentlyContinue }
            $usbDisks += [PSCustomObject]@{
                DiskNumber = $d.DiskNumber
                Model      = $d.FriendlyName
                SizeGB     = [math]::Round($d.Size / 1GB, 1)
                Status     = $d.OperationalStatus
                Volumes    = $volumes
                Letters    = ($volumes | Where-Object { $_.DriveLetter } | ForEach-Object { "$($_.DriveLetter):" }) -join ', '
            }
            Write-Log "  Found USB: Disk $($d.DiskNumber) | $($d.FriendlyName) | $([math]::Round($d.Size/1GB,1)) GB | Letters: $(if ($volumes) {($volumes|Where-Object{$_.DriveLetter}|ForEach-Object{"$($_.DriveLetter):"}) -join ', '} else {'none'})" -Level INFO
        }
    } catch {
        Write-Log "Error enumerating USB disks: $_" -Level ERROR
    }

    if ($usbDisks.Count -eq 0) {
        Write-Log "No USB drives found. Insert a USB drive and re-run." -Level ERROR
    }

    return $usbDisks
}

function Select-UsbDrive {
    <#
    .SYNOPSIS Shows USB drive list and asks user to select one.
    #>
    $usbs = Get-UsbDrives
    if ($usbs.Count -eq 0) { return $null }

    Write-Host "`n$('='*60)" -ForegroundColor White
    Write-Host "  SELECT USB DRIVE TO FORMAT" -ForegroundColor Yellow
    Write-Host "  !! ALL DATA ON SELECTED DRIVE WILL BE ERASED !!" -ForegroundColor Red
    Write-Host "$('='*60)" -ForegroundColor White

    for ($i = 0; $i -lt $usbs.Count; $i++) {
        $u = $usbs[$i]
        Write-Host "  [$($i+1)] Disk $($u.DiskNumber) -- $($u.Model) -- $($u.SizeGB) GB -- Letters: $(if ($u.Letters) {$u.Letters} else {'(none)'})" -ForegroundColor Cyan
    }
    Write-Host ""

    do {
        $choice = Read-Host "Select USB drive (1-$($usbs.Count)) or Q to quit"
        if ($choice -match '^[Qq]$') { return $null }
        $idx = [int]$choice - 1
    } while ($idx -lt 0 -or $idx -ge $usbs.Count)

    $selected = $usbs[$idx]
    Write-Log "User selected: Disk $($selected.DiskNumber) -- $($selected.Model) -- $($selected.SizeGB) GB" -Level SUCCESS

    Write-Host "`n[CONFIRM] You selected: Disk $($selected.DiskNumber) -- $($selected.Model) -- $($selected.SizeGB) GB" -ForegroundColor Yellow
    Write-Host "          ALL DATA WILL BE PERMANENTLY ERASED." -ForegroundColor Red
    $confirm = Read-Host "Type YES (uppercase) to confirm"
    if ($confirm -ne 'YES') {
        Write-Log "User cancelled USB selection." -Level WARN
        return $null
    }

    $Script:SelectedUSB = $selected
    return $selected
}

# ============================================================
# SECTION 2: FORMAT USB
# ============================================================
function Format-UsbDrive {
    <#
    .SYNOPSIS Wipes and formats the selected USB drive as FAT32 (GPT or MBR).
    .PARAMETER PartitionStyle  'GPT' (for UEFI-only) or 'MBR' (for BIOS+UEFI compat)
    #>
    param(
        [ValidateSet('GPT','MBR')][string]$PartitionStyle = 'MBR'
    )

    if (-not $Script:SelectedUSB) {
        Write-Log "No USB drive selected." -Level ERROR
        return $null
    }

    $diskNum = $Script:SelectedUSB.DiskNumber
    Write-Log "Formatting Disk $diskNum as $PartitionStyle / FAT32..." -Level SECTION

    # Build diskpart script
    if ($PartitionStyle -eq 'GPT') {
        $dpScript = @"
select disk $diskNum
clean
convert gpt
create partition primary
format fs=fat32 quick label="WinRescue"
assign
active
"@
    } else {
        # MBR -- most compatible (boots on BIOS and most UEFI with CSM)
        $dpScript = @"
select disk $diskNum
clean
convert mbr
create partition primary
format fs=fat32 quick label="WinRescue"
assign
active
"@
    }

    $dpFile = Join-Path $Script:WorkDir 'format_usb.txt'
    $dpScript | Set-Content $dpFile -Encoding ASCII

    Write-Log "Running diskpart format (this will take 10-60 seconds)..." -Level STEP
    $result = & diskpart.exe /s $dpFile 2>&1
    $exitCode = $LASTEXITCODE
    $result | ForEach-Object { Write-Log "  $_" -Level INFO }
    Remove-Item $dpFile -Force -ErrorAction SilentlyContinue

    if ($exitCode -ne 0) {
        Write-Log "diskpart exited with code $exitCode -- format may have failed." -Level WARN
    } else {
        Write-Log "Disk $diskNum formatted successfully as $PartitionStyle/FAT32." -Level SUCCESS
    }

    # Find the new drive letter
    Start-Sleep -Seconds 2
    $newDisk = Get-Disk -Number $diskNum -ErrorAction SilentlyContinue
    $newPart = Get-Partition -DiskNumber $diskNum -ErrorAction SilentlyContinue | Select-Object -First 1
    $newVol  = if ($newPart) { Get-Volume -Partition $newPart -ErrorAction SilentlyContinue } else { $null }

    if ($newVol -and $newVol.DriveLetter) {
        Write-Log "USB drive letter: $($newVol.DriveLetter):" -Level SUCCESS
        return "$($newVol.DriveLetter):"
    }

    # Fallback: assign letter via diskpart
    Write-Log "No drive letter assigned. Assigning via diskpart..." -Level WARN
    $freeLetter = Get-FirstFreeLetter
    if ($freeLetter) {
        $assignScript = "select disk $diskNum`r`nselect partition 1`r`nassign letter=$freeLetter`r`n"
        $assignFile   = Join-Path $Script:WorkDir 'assign_letter.txt'
        $assignScript | Set-Content $assignFile -Encoding ASCII
        & diskpart.exe /s $assignFile 2>&1 | Out-Null
        Remove-Item $assignFile -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
        Write-Log "Assigned letter $freeLetter to USB." -Level SUCCESS
        return "${freeLetter}:"
    }

    Write-Log "Could not determine USB drive letter after format." -Level ERROR
    return $null
}

function Get-FirstFreeLetter {
    $used = [System.IO.DriveInfo]::GetDrives() | ForEach-Object { $_.Name[0] }
    foreach ($l in 'F','G','H','I','J','K','L','M','N','O','P','R','S','T','U','V','W') {
        if ($l -notin $used) { return $l }
    }
    return $null
}

# ============================================================
# SECTION 3: DETECT ADK / WinPE
# ============================================================
function Find-AdkPath {
    <#
    .SYNOPSIS Finds the Windows ADK installation path.
    .RETURNS Path to ADK root or $null if not installed.
    #>
    $adkPaths = @(
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows Kits\Installed Roots',
        'HKLM:\SOFTWARE\Microsoft\Windows Kits\Installed Roots'
    )
    foreach ($regPath in $adkPaths) {
        $key = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
        if ($key -and $key.KitsRoot10) {
            $adkRoot = $key.KitsRoot10
            Write-Log "ADK found at: $adkRoot" -Level SUCCESS
            return $adkRoot
        }
    }

    # Common install paths
    $commonPaths = @(
        'C:\Program Files (x86)\Windows Kits\10',
        'C:\Program Files\Windows Kits\10'
    )
    foreach ($p in $commonPaths) {
        if (Test-Path $p) {
            Write-Log "ADK found (path scan): $p" -Level SUCCESS
            return $p
        }
    }

    Write-Log "Windows ADK not found." -Level WARN
    return $null
}

function Find-WinPeAddon {
    param([string]$AdkRoot)
    $wpeRoot = Join-Path $AdkRoot 'Assessment and Deployment Kit\Windows Preinstallation Environment'
    if (Test-Path $wpeRoot) {
        Write-Log "WinPE Add-on found: $wpeRoot" -Level SUCCESS
        return $wpeRoot
    }
    Write-Log "WinPE Add-on not found at: $wpeRoot" -Level WARN
    return $null
}

# ============================================================
# SECTION 4: MODE A -- ADK WinPE BUILD
# ============================================================
function Build-WinPeUsb {
    <#
    .SYNOPSIS Full WinPE USB creation using ADK copype + DISM + bcdboot.
    .PARAMETER UsbLetter  Drive letter of the formatted USB (e.g. 'F:')
    .PARAMETER AdkRoot    Path to ADK installation
    #>
    param([string]$UsbLetter, [string]$AdkRoot)

    Write-Log "Building WinPE environment using ADK..." -Level SECTION

    $copype  = Join-Path $AdkRoot 'Assessment and Deployment Kit\Windows Preinstallation Environment\copype.cmd'
    $wpePe   = Join-Path $Script:WorkDir 'WinPE_amd64'

    if (-not (Test-Path $copype)) {
        Write-Log "copype.cmd not found at: $copype" -Level ERROR
        return $false
    }

    # Run copype to generate WinPE working files
    Write-Log "Running copype amd64 $wpePe ..." -Level STEP
    $cpResult = & cmd.exe /c "`"$copype`" amd64 `"$wpePe`"" 2>&1
    $cpResult | ForEach-Object { Write-Log "  $_" -Level INFO }

    $wpeWim = Join-Path $wpePe 'media\sources\boot.wim'
    if (-not (Test-Path $wpeWim)) {
        Write-Log "copype failed -- boot.wim not found at $wpeWim" -Level ERROR
        return $false
    }
    Write-Log "copype completed. boot.wim: $wpeWim" -Level SUCCESS

    # Mount boot.wim, inject tools
    $mountDir = Join-Path $Script:WorkDir 'pe_mount'
    if (-not (Test-Path $mountDir)) { New-Item -ItemType Directory $mountDir -Force | Out-Null }

    Write-Log "Mounting boot.wim for customization..." -Level STEP
    $dismMount = & dism.exe /Mount-Image /ImageFile:"$wpeWim" /Index:1 /MountDir:"$mountDir" 2>&1
    $dismMount | ForEach-Object { Write-Log "  $_" -Level INFO }

    if ($LASTEXITCODE -ne 0) {
        Write-Log "DISM mount failed. Proceeding without customization." -Level WARN
    } else {
        # Copy our toolkit into the WinPE image
        $peToolDir = Join-Path $mountDir 'Windows\System32\winboot-rescue'
        New-Item -ItemType Directory -Path $peToolDir -Force | Out-Null
        $toolFiles = @('boot-repair.ps1', 'boot-repair.cmd')
        foreach ($f in $toolFiles) {
            $src = Join-Path $Script:ToolkitDir $f
            if (Test-Path $src) {
                Copy-Item $src $peToolDir -Force
                Write-Log "  Injected $f into WinPE image." -Level SUCCESS
            } else {
                Write-Log "  $f not found in $Script:ToolkitDir -- skipping injection." -Level WARN
            }
        }

        # Append boot hint to startnet.cmd (wpeinit must stay first)
        $startupHint = Join-Path $mountDir 'Windows\System32\startnet.cmd'
        try {
            $existing = Get-Content $startupHint -ErrorAction SilentlyContinue
            if ($existing -notmatch 'winboot-rescue') {
                Add-Content -Path $startupHint -Value "`r`necho." -Encoding ASCII
                Add-Content -Path $startupHint -Value "echo  winboot-rescue: X:\Windows\System32\winboot-rescue\boot-repair.cmd" -Encoding ASCII
            }
        } catch {}

        # Unmount and commit
        Write-Log "Committing WinPE image changes..." -Level STEP
        $dismUnmount = & dism.exe /Unmount-Image /MountDir:"$mountDir" /Commit 2>&1
        $dismUnmount | ForEach-Object { Write-Log "  $_" -Level INFO }
        if ($LASTEXITCODE -eq 0) {
            Write-Log "WinPE image customized and committed." -Level SUCCESS
        } else {
            Write-Log "DISM unmount had errors -- image may still work." -Level WARN
        }
    }

    # Copy WinPE files to USB
    Write-Log "Copying WinPE media to USB $UsbLetter ..." -Level STEP
    $mediaDir = Join-Path $wpePe 'media'
    $copyResult = & robocopy.exe "$mediaDir" "$UsbLetter\" /E /NFL /NDL /NJH /NJS 2>&1
    Write-Log "robocopy exit: $LASTEXITCODE (0-7 = OK)" -Level INFO

    # Run bcdboot for safety on MBR USB
    $windowsDir = Join-Path $UsbLetter 'Windows'
    if (Test-Path $windowsDir) {
        Write-Log "Running bcdboot to ensure boot sector..." -Level STEP
        & bcdboot.exe "$windowsDir" /s "$UsbLetter" /f ALL 2>&1 | ForEach-Object { Write-Log "  $_" -Level INFO }
    }

    # Copy toolkit to USB root for easy access
    Copy-ToolkitToUsb -UsbLetter $UsbLetter

    Write-Log "WinPE USB creation complete!" -Level SUCCESS
    return $true
}

# ============================================================
# SECTION 5: MODE B -- WinRE from local machine
# ============================================================
function Build-WinReUsb {
    <#
    .SYNOPSIS Uses the local machine's WinRE.wim to create a bootable USB.
    .PARAMETER UsbLetter  Drive letter of formatted USB
    .NOTES Does not require ADK. Uses reagentc to locate WinRE.wim.
    #>
    param([string]$UsbLetter)

    Write-Log "Building WinRE USB from local Windows Recovery Environment..." -Level SECTION

    # Find WinRE.wim
    $winReWim = Find-WinReWim
    if (-not $winReWim) {
        Write-Log "WinRE.wim not found. Cannot use Mode B." -Level ERROR
        return $false
    }

    Write-Log "WinRE.wim found: $winReWim" -Level SUCCESS

    # Create sources folder on USB
    $usbSources = Join-Path $UsbLetter 'sources'
    if (-not (Test-Path $usbSources)) { New-Item -ItemType Directory $usbSources -Force | Out-Null }

    # Copy WinRE.wim as boot.wim
    $bootWimDest = Join-Path $usbSources 'boot.wim'
    Write-Log "Copying WinRE.wim to USB as sources\boot.wim ..." -Level STEP
    try {
        Copy-Item $winReWim $bootWimDest -Force
        Write-Log "boot.wim copied successfully." -Level SUCCESS
    } catch {
        Write-Log "Failed to copy WinRE.wim: $_" -Level ERROR
        return $false
    }

    # Mount and inject toolkit
    $mountDir = Join-Path $Script:WorkDir 're_mount'
    if (-not (Test-Path $mountDir)) { New-Item -ItemType Directory $mountDir -Force | Out-Null }

    Write-Log "Mounting WinRE image to inject toolkit..." -Level STEP
    $dismMount = & dism.exe /Mount-Image /ImageFile:"$bootWimDest" /Index:1 /MountDir:"$mountDir" 2>&1
    if ($LASTEXITCODE -eq 0) {
        $reToolDir = Join-Path $mountDir 'Windows\System32\winboot-rescue'
        New-Item -ItemType Directory -Path $reToolDir -Force | Out-Null
        foreach ($f in @('boot-repair.ps1', 'boot-repair.cmd')) {
            $src = Join-Path $Script:ToolkitDir $f
            if (Test-Path $src) {
                Copy-Item $src $reToolDir -Force
                Write-Log "  Injected $f into WinRE image." -Level SUCCESS
            }
        }

        & dism.exe /Unmount-Image /MountDir:"$mountDir" /Commit 2>&1 | ForEach-Object { Write-Log "  $_" -Level INFO }
    } else {
        Write-Log "Could not mount WinRE.wim for injection -- toolkit will be on USB root instead." -Level WARN
        & dism.exe /Unmount-Image /MountDir:"$mountDir" /Discard 2>&1 | Out-Null
    }

    # Copy boot files from Windows installation
    $windowsBoot = 'C:\Windows\Boot'
    if (Test-Path $windowsBoot) {
        Write-Log "Copying Windows boot files to USB..." -Level STEP
        & robocopy.exe "$windowsBoot\EFI" "$UsbLetter\EFI" /E /NFL /NDL /NJH /NJS 2>&1 | Out-Null
        & robocopy.exe "$windowsBoot\PCAT" "$UsbLetter\Boot" /E /NFL /NDL /NJH /NJS 2>&1 | Out-Null
        Write-Log "Boot files copied." -Level SUCCESS
    }

    # Make USB bootable with bcdboot
    $systemRoot = $env:SystemRoot
    Write-Log "Running bcdboot to create boot entries on USB..." -Level STEP
    $bcdResult = & bcdboot.exe "$systemRoot" /s "$UsbLetter" /f ALL 2>&1
    $bcdResult | ForEach-Object { Write-Log "  $_" -Level INFO }

    if ($LASTEXITCODE -eq 0) {
        Write-Log "bcdboot completed successfully." -Level SUCCESS
    } else {
        Write-Log "bcdboot returned $LASTEXITCODE -- USB may still boot but verify." -Level WARN
    }

    # Bootsect for MBR/BIOS compatibility
    $bootsect = 'C:\Windows\System32\bootsect.exe'
    if (Test-Path $bootsect) {
        Write-Log "Running bootsect for MBR/BIOS compatibility..." -Level STEP
        & $bootsect /nt60 "$UsbLetter" /force /mbr 2>&1 | ForEach-Object { Write-Log "  $_" -Level INFO }
    }

    Copy-ToolkitToUsb -UsbLetter $UsbLetter
    Write-Log "WinRE USB creation complete!" -Level SUCCESS
    return $true
}

function Find-WinReWim {
    <#
    .SYNOPSIS Locates WinRE.wim on the local system.
    #>
    # Check reagentc first
    try {
        $reagentOut = & reagentc.exe /info 2>&1 | Out-String
        if ($reagentOut -match 'Windows RE location\s*:\s*(.+\.wim)') {
            $wimPath = $Matches[1].Trim()
            if (Test-Path $wimPath) { return $wimPath }
        }
    } catch {}

    # Common locations
    $candidates = @(
        'C:\Windows\System32\Recovery\WinRE.wim',
        'C:\Recovery\WindowsRE\WinRE.wim',
        'D:\Recovery\WindowsRE\WinRE.wim',
        "$env:SystemRoot\System32\Recovery\WinRE.wim"
    )
    foreach ($p in $candidates) {
        if (Test-Path $p -ErrorAction SilentlyContinue) { return $p }
    }

    # Search Recovery partitions
    try {
        $drives = [System.IO.DriveInfo]::GetDrives() | Where-Object { $_.IsReady }
        foreach ($d in $drives) {
            $candidate = Join-Path $d.RootDirectory.FullName 'Recovery\WindowsRE\WinRE.wim'
            if (Test-Path $candidate -ErrorAction SilentlyContinue) { return $candidate }
        }
    } catch {}

    return $null
}

# ============================================================
# SECTION 6: MODE C -- Windows ISO
# ============================================================
function Build-IsoUsb {
    <#
    .SYNOPSIS Uses a Windows ISO file to create a bootable USB.
    .PARAMETER UsbLetter   Drive letter of formatted USB
    .PARAMETER IsoPath     Path to Windows 10/11 ISO file
    #>
    param([string]$UsbLetter, [string]$IsoPath)

    Write-Log "Building USB from Windows ISO: $IsoPath" -Level SECTION

    if (-not (Test-Path $IsoPath)) {
        Write-Log "ISO file not found: $IsoPath" -Level ERROR
        return $false
    }

    # Mount ISO
    Write-Log "Mounting ISO..." -Level STEP
    try {
        $mountResult = Mount-DiskImage -ImagePath $IsoPath -PassThru
        $isoDrive    = ($mountResult | Get-Volume).DriveLetter + ':'
        Write-Log "ISO mounted at: $isoDrive" -Level SUCCESS
    } catch {
        Write-Log "Failed to mount ISO: $_" -Level ERROR
        return $false
    }

    try {
        # Copy all ISO contents to USB
        Write-Log "Copying ISO contents to USB $UsbLetter (this takes a few minutes)..." -Level STEP
        $roboCopy = & robocopy.exe "$isoDrive\" "$UsbLetter\" /E /NFL /NDL /NJH /NJS 2>&1
        Write-Log "robocopy exit: $LASTEXITCODE (0-7 = OK)" -Level INFO

        Copy-ToolkitToUsb -UsbLetter $UsbLetter

        # Run bootsect for BIOS compatibility
        $bootsect = "$isoDrive\boot\bootsect.exe"
        if (Test-Path $bootsect) {
            Write-Log "Running bootsect for BIOS boot support..." -Level STEP
            & $bootsect /nt60 "$UsbLetter" /force /mbr 2>&1 | ForEach-Object { Write-Log "  $_" -Level INFO }
        }

        Write-Log "ISO USB creation complete!" -Level SUCCESS
        return $true
    } finally {
        # Always unmount ISO
        try { Dismount-DiskImage -ImagePath $IsoPath | Out-Null } catch {}
        Write-Log "ISO unmounted." -Level INFO
    }
}

# ============================================================
# SECTION 7: COPY TOOLKIT TO USB
# ============================================================
function Copy-ToolkitToUsb {
    <#
    .SYNOPSIS Copies boot-repair.ps1 and boot-repair.cmd to the USB root.
    #>
    param([string]$UsbLetter)

    Write-Log "Copying winboot-rescue toolkit to USB root..." -Level STEP

    $toolFiles = @('boot-repair.ps1', 'boot-repair.cmd')
    $anyMissing = $false

    foreach ($f in $toolFiles) {
        $src  = Join-Path $Script:ToolkitDir $f
        $dest = Join-Path $UsbLetter $f

        if (Test-Path $src) {
            Copy-Item $src $dest -Force
            Write-Log "  Copied: $f -> $UsbLetter\$f" -Level SUCCESS
        } else {
            Write-Log "  NOT FOUND: $f (expected at $src)" -Level WARN
            $anyMissing = $true
        }
    }

    # Create a README on the USB
    $readmeContent = @"
WINBOOT-RESCUE TOOLKIT
======================

Files on this USB:
  boot-repair.cmd  <- Launch this from WinRE command prompt
  boot-repair.ps1  <- Main PowerShell script

HOW TO USE:
  1. Boot from this USB (change boot order in BIOS/UEFI)
  2. When WinRE/WinPE loads, open Command Prompt
  3. Find this USB drive letter (usually NOT C:)
     > diskpart
     > list volume
     > exit
  4. Navigate to this USB and run:
     > E:\boot-repair.cmd   (replace E: with your USB letter)
  5. Choose option [1] first to collect diagnostics
  6. Then choose [2] or [4]/[5] based on your system type

FIRST TIME: Always run option [1] (diagnostics) before attempting repair.

GitHub: https://github.com/Gzeu/winboot-rescue
"@
    $readmeDest = Join-Path $UsbLetter 'README-BOOT-REPAIR.txt'
    Set-Content -Path $readmeDest -Value $readmeContent -Encoding UTF8 -ErrorAction SilentlyContinue
    Write-Log "  Created README-BOOT-REPAIR.txt on USB." -Level INFO

    if (-not $anyMissing) {
        Write-Log "All toolkit files copied to USB." -Level SUCCESS
    }
}

# ============================================================
# SECTION 8: VERIFY USB
# ============================================================
function Test-UsbBoot {
    <#
    .SYNOPSIS Verifies the USB has required boot files.
    #>
    param([string]$UsbLetter)

    Write-Log "Verifying USB boot files..." -Level SECTION

    $checks = @(
        @{ Path = "$UsbLetter\sources\boot.wim";       Required = $true  ; Label = "WinPE/WinRE image (sources\boot.wim)" },
        @{ Path = "$UsbLetter\boot\bootmgr";            Required = $false ; Label = "BIOS boot manager (boot\bootmgr)" },
        @{ Path = "$UsbLetter\bootmgr";                 Required = $false ; Label = "BIOS bootmgr (root)" },
        @{ Path = "$UsbLetter\EFI\Boot\bootx64.efi";   Required = $false ; Label = "UEFI boot file (EFI\Boot\bootx64.efi)" },
        @{ Path = "$UsbLetter\boot-repair.cmd";         Required = $true  ; Label = "winboot-rescue launcher" },
        @{ Path = "$UsbLetter\boot-repair.ps1";         Required = $true  ; Label = "winboot-rescue main script" }
    )

    $allGood = $true
    foreach ($c in $checks) {
        $exists = Test-Path $c.Path -ErrorAction SilentlyContinue
        $status = if ($exists) { 'SUCCESS' } elseif ($c.Required) { 'ERROR'; $allGood = $false } else { 'WARN' }
        $icon   = if ($exists) { '[OK]' } else { '[MISSING]' }
        Write-Log "  $icon $($c.Label)" -Level $status
    }

    if ($allGood) {
        Write-Log "USB verification passed. Drive is ready." -Level SUCCESS
    } else {
        Write-Log "USB verification: some required files missing. USB may not boot correctly." -Level WARN
    }

    return $allGood
}

# ============================================================
# SECTION 9: MAIN MENU
# ============================================================
function Show-Menu {
    $adkRoot = Find-AdkPath
    $hasAdk  = $null -ne $adkRoot

    Write-Host "`n$('='*60)" -ForegroundColor White
    Write-Host "  WINBOOT-RESCUE -- USB CREATOR" -ForegroundColor Cyan
    Write-Host "$('='*60)" -ForegroundColor White
    Write-Host "  ADK installed: $(if ($hasAdk) {'YES -- Mode A available'} else {'NO -- Mode A unavailable'})" `
        -ForegroundColor $(if ($hasAdk) {'Green'} else {'Yellow'})
    Write-Host "$('-'*60)" -ForegroundColor DarkGray
    Write-Host "  [1] Mode A -- WinPE USB from ADK (best, full WinPE)$(if (-not $hasAdk){' [ADK required]'})" `
        -ForegroundColor $(if ($hasAdk) {'Green'} else {'DarkGray'})
    Write-Host "  [2] Mode B -- WinRE USB from local machine (no ADK needed)" -ForegroundColor Cyan
    Write-Host "  [3] Mode C -- USB from Windows ISO file" -ForegroundColor Cyan
    Write-Host "  [4] Only copy toolkit to existing bootable USB" -ForegroundColor Yellow
    Write-Host "  [5] Verify existing USB" -ForegroundColor White
    Write-Host "  [0] Exit" -ForegroundColor Red
    Write-Host "$('='*60)" -ForegroundColor White

    $choice = Read-Host "`nSelect option"
    Write-Log "User selected mode: $choice"

    switch ($choice) {
        '1' {
            if (-not $hasAdk) {
                Write-Log "ADK not installed. Download from: https://learn.microsoft.com/windows-hardware/get-started/adk-install" -Level ERROR
                Write-Host "`n  Download ADK + WinPE Add-on from Microsoft, then re-run." -ForegroundColor Yellow
                return
            }
            $usb = Select-UsbDrive
            if (-not $usb) { return }
            $usbLetter = Format-UsbDrive -PartitionStyle 'MBR'
            if (-not $usbLetter) { return }
            Build-WinPeUsb -UsbLetter $usbLetter -AdkRoot $adkRoot
            Test-UsbBoot -UsbLetter $usbLetter
        }
        '2' {
            $usb = Select-UsbDrive
            if (-not $usb) { return }
            $usbLetter = Format-UsbDrive -PartitionStyle 'MBR'
            if (-not $usbLetter) { return }
            Build-WinReUsb -UsbLetter $usbLetter
            Test-UsbBoot -UsbLetter $usbLetter
        }
        '3' {
            $isoPath = Read-Host "Enter full path to Windows ISO file"
            $isoPath = $isoPath.Trim('"')
            if (-not (Test-Path $isoPath)) {
                Write-Log "ISO not found: $isoPath" -Level ERROR
                return
            }
            $usb = Select-UsbDrive
            if (-not $usb) { return }
            $usbLetter = Format-UsbDrive -PartitionStyle 'MBR'
            if (-not $usbLetter) { return }
            Build-IsoUsb -UsbLetter $usbLetter -IsoPath $isoPath
            Test-UsbBoot -UsbLetter $usbLetter
        }
        '4' {
            $letter = Read-Host "Enter existing USB drive letter (e.g. E)"
            $usbLetter = "$($letter.Trim(':')):"
            if (-not (Test-Path $usbLetter)) {
                Write-Log "Drive $usbLetter not found." -Level ERROR
                return
            }
            Copy-ToolkitToUsb -UsbLetter $usbLetter
            Test-UsbBoot -UsbLetter $usbLetter
        }
        '5' {
            $letter = Read-Host "Enter USB drive letter to verify (e.g. E)"
            $usbLetter = "$($letter.Trim(':')):"
            Test-UsbBoot -UsbLetter $usbLetter
        }
        '0' { return }
        default { Write-Host "  Invalid option." -ForegroundColor Red }
    }

    # Show log location
    Write-Host "`n[INFO] Log file: $Script:LogFile" -ForegroundColor DarkGray
    Write-Host "Press any key to return to menu..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Show-Menu
}

# ============================================================
# ENTRY POINT
# ============================================================
function Main {
    try { Clear-Host } catch {}

    Write-Host @"
$('='*60)
  WINBOOT-RESCUE -- USB STICK CREATOR
  Creates a bootable WinPE/WinRE USB with boot-repair toolkit
$('='*60)
  !! Run as Administrator !!
  !! All data on selected USB will be ERASED !!
$('='*60)
"@ -ForegroundColor Cyan

    Initialize-WorkDir

    # Check admin
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Log "Not running as Administrator! Re-launch as Admin." -Level ERROR
        Write-Host "`nRight-click the script and choose 'Run as administrator'" -ForegroundColor Yellow
        Read-Host "Press Enter to exit"
        exit 1
    }

    Write-Log "Script directory: $Script:ToolkitDir"
    Write-Log "Checking for toolkit files..."

    foreach ($f in @('boot-repair.ps1','boot-repair.cmd')) {
        $path = Join-Path $Script:ToolkitDir $f
        if (Test-Path $path) {
            Write-Log "  [OK] $f found" -Level SUCCESS
        } else {
            Write-Log "  [MISSING] $f not found in $Script:ToolkitDir" -Level WARN
            Write-Log "  -> Place all files in the same folder before running." -Level WARN
        }
    }

    Show-Menu

    Write-Log "Session ended: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level INFO
}

Main
