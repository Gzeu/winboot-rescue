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
    FIXES   : copype.cmd path detection, ReadKey crash, ADK sub-path scan
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

# Safe "press any key" — works in all PS hosts including cmd-launched sessions
function Wait-KeyPress {
    param([string]$Prompt = 'Press Enter to continue...')
    Write-Host "`n$Prompt" -ForegroundColor DarkGray
    try {
        # Try ReadKey first (works in interactive terminal)
        if ($Host.UI.RawUI.KeyAvailable -ne $null) {
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
            return
        }
    } catch {}
    # Fallback: Read-Host works everywhere
    Read-Host | Out-Null
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
    .NOTES  Result is always wrapped in @() so .Count is safe under StrictMode
            even when only one USB drive is present.
    #>
    Write-Log "Scanning for USB drives..." -Level STEP

    $usbDisks = [System.Collections.Generic.List[object]]::new()
    try {
        $allDisks = @(Get-Disk -ErrorAction SilentlyContinue | Where-Object { $_.BusType -eq 'USB' })
        foreach ($d in $allDisks) {
            $volumes = @(Get-Partition -DiskNumber $d.DiskNumber -ErrorAction SilentlyContinue |
                       ForEach-Object { Get-Volume -Partition $_ -ErrorAction SilentlyContinue })
            $letters = ($volumes | Where-Object { $_.DriveLetter } | ForEach-Object { "$($_.DriveLetter):" }) -join ', '
            $usbDisks.Add([PSCustomObject]@{
                DiskNumber = $d.DiskNumber
                Model      = $d.FriendlyName
                SizeGB     = [math]::Round($d.Size / 1GB, 1)
                Status     = $d.OperationalStatus
                Volumes    = $volumes
                Letters    = $letters
            })
            Write-Log "  Found USB: Disk $($d.DiskNumber) | $($d.FriendlyName) | $([math]::Round($d.Size/1GB,1)) GB | Letters: $(if ($letters) {$letters} else {'none'})" -Level INFO
        }
    } catch {
        Write-Log "Error enumerating USB disks: $_" -Level ERROR
    }

    return @($usbDisks)
}

function Select-UsbDrive {
    $usbs = @(Get-UsbDrives)
    if ($usbs.Count -eq 0) {
        Write-Log "No USB drives found. Insert a USB drive and re-run." -Level ERROR
        return $null
    }

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
        [int]$idx = 0
        $validNum = [int]::TryParse($choice, [ref]$idx)
        $idx = $idx - 1
    } while (-not $validNum -or $idx -lt 0 -or $idx -ge $usbs.Count)

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
    param(
        [ValidateSet('GPT','MBR')][string]$PartitionStyle = 'MBR'
    )

    if (-not $Script:SelectedUSB) {
        Write-Log "No USB drive selected." -Level ERROR
        return $null
    }

    $diskNum = $Script:SelectedUSB.DiskNumber
    Write-Log "Formatting Disk $diskNum as $PartitionStyle / FAT32..." -Level SECTION

    if ($PartitionStyle -eq 'GPT') {
        $dpScript = @"
select disk $diskNum
clean
convert gpt
create partition primary
format fs=fat32 quick label="WinRescue"
assign
"@
    } else {
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

    # Find new drive letter
    Start-Sleep -Seconds 2
    $newPart = @(Get-Partition -DiskNumber $diskNum -ErrorAction SilentlyContinue) | Select-Object -First 1
    $newVol  = if ($newPart) { Get-Volume -Partition $newPart -ErrorAction SilentlyContinue } else { $null }

    if ($newVol -and $newVol.DriveLetter) {
        Write-Log "USB drive letter: $($newVol.DriveLetter):" -Level SUCCESS
        return "$($newVol.DriveLetter):"
    }

    # Fallback: assign letter
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
    .SYNOPSIS Finds the Windows ADK installation path and locates copype.cmd.
    .RETURNS Hashtable with AdkRoot and CopypePath, or $null if not found.
    .NOTES   ADK + WinPE Add-on must both be installed.
             copype.cmd is part of the WinPE Add-on, not base ADK.
             Scans registry first, then common install paths, then deep search.
    #>

    # Step 1: Try registry for ADK root
    $adkRoot = $null
    $adkRegPaths = @(
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows Kits\Installed Roots',
        'HKLM:\SOFTWARE\Microsoft\Windows Kits\Installed Roots'
    )
    foreach ($regPath in $adkRegPaths) {
        try {
            $key = Get-ItemProperty $regPath -ErrorAction SilentlyContinue
            if ($key -and $key.KitsRoot10) {
                $adkRoot = $key.KitsRoot10.TrimEnd('\')
                break
            }
        } catch {}
    }

    # Step 2: Fallback to common install directories
    if (-not $adkRoot) {
        $commonPaths = @(
            "${env:ProgramFiles(x86)}\Windows Kits\10",
            "${env:ProgramFiles}\Windows Kits\10",
            'C:\Program Files (x86)\Windows Kits\10',
            'C:\Program Files\Windows Kits\10'
        )
        foreach ($p in $commonPaths) {
            if (Test-Path $p -ErrorAction SilentlyContinue) {
                $adkRoot = $p
                break
            }
        }
    }

    if (-not $adkRoot) {
        Write-Log "Windows ADK not found." -Level WARN
        return $null
    }

    Write-Log "ADK root found: $adkRoot" -Level SUCCESS

    # Step 3: Locate copype.cmd — it ships with the WinPE Add-on
    # Standard path relative to ADK root:
    $copypeRelative = 'Assessment and Deployment Kit\Windows Preinstallation Environment\copype.cmd'
    $copypeStandard = Join-Path $adkRoot $copypeRelative
    if (Test-Path $copypeStandard -ErrorAction SilentlyContinue) {
        Write-Log "copype.cmd found: $copypeStandard" -Level SUCCESS
        return @{ AdkRoot = $adkRoot; CopypePath = $copypeStandard }
    }

    # Step 4: Deep search under ADK root for copype.cmd
    Write-Log "copype.cmd not at standard path. Searching under ADK root..." -Level WARN
    try {
        $found = Get-ChildItem -Path $adkRoot -Filter 'copype.cmd' -Recurse -ErrorAction SilentlyContinue |
                 Select-Object -First 1
        if ($found) {
            Write-Log "copype.cmd found via search: $($found.FullName)" -Level SUCCESS
            return @{ AdkRoot = $adkRoot; CopypePath = $found.FullName }
        }
    } catch {}

    # Step 5: Search whole Program Files
    Write-Log "Searching Program Files for copype.cmd..." -Level WARN
    foreach ($searchRoot in @(${env:ProgramFiles}, ${env:ProgramFiles(x86)})) {
        if (-not $searchRoot) { continue }
        try {
            $found = Get-ChildItem -Path $searchRoot -Filter 'copype.cmd' -Recurse -ErrorAction SilentlyContinue |
                     Select-Object -First 1
            if ($found) {
                Write-Log "copype.cmd found: $($found.FullName)" -Level SUCCESS
                return @{ AdkRoot = $adkRoot; CopypePath = $found.FullName }
            }
        } catch {}
    }

    Write-Log "ADK found at $adkRoot but copype.cmd is missing." -Level WARN
    Write-Log "Install the WinPE Add-on for ADK: https://learn.microsoft.com/windows-hardware/get-started/adk-install" -Level WARN
    # Return ADK found but no copype — caller can decide
    return @{ AdkRoot = $adkRoot; CopypePath = $null }
}

# ============================================================
# SECTION 4: MODE A -- ADK WinPE BUILD
# ============================================================
function Build-WinPeUsb {
    param([string]$UsbLetter, [hashtable]$AdkInfo)

    Write-Log "Building WinPE environment using ADK..." -Level SECTION

    $copype  = $AdkInfo.CopypePath
    $adkRoot = $AdkInfo.AdkRoot
    $wpePe   = Join-Path $Script:WorkDir 'WinPE_amd64'

    if (-not $copype -or -not (Test-Path $copype)) {
        Write-Log "copype.cmd not available. Cannot build WinPE." -Level ERROR
        Write-Log "Install the WinPE Add-on: https://learn.microsoft.com/windows-hardware/get-started/adk-install" -Level WARN
        return $false
    }

    # Run copype to generate WinPE working files
    Write-Log "Running: copype amd64 $wpePe ..." -Level STEP
    $cpResult = & cmd.exe /c "`"$copype`" amd64 `"$wpePe`"" 2>&1
    $cpResult | ForEach-Object { Write-Log "  $_" -Level INFO }

    $wpeWim = Join-Path $wpePe 'media\sources\boot.wim'
    if (-not (Test-Path $wpeWim)) {
        Write-Log "copype failed -- boot.wim not found at $wpeWim" -Level ERROR
        return $false
    }
    Write-Log "copype completed. boot.wim ready." -Level SUCCESS

    # Mount boot.wim and inject toolkit
    $mountDir = Join-Path $Script:WorkDir 'pe_mount'
    if (-not (Test-Path $mountDir)) { New-Item -ItemType Directory $mountDir -Force | Out-Null }

    Write-Log "Mounting WinPE image for customization..." -Level STEP
    $dismMount = & dism.exe /Mount-Image /ImageFile:"$wpeWim" /Index:1 /MountDir:"$mountDir" 2>&1
    $dismMount | ForEach-Object { Write-Log "  $_" -Level INFO }

    if ($LASTEXITCODE -ne 0) {
        Write-Log "DISM mount failed. Proceeding without customization." -Level WARN
    } else {
        $peToolDir = Join-Path $mountDir 'Windows\System32\winboot-rescue'
        New-Item -ItemType Directory -Path $peToolDir -Force | Out-Null
        foreach ($f in @('boot-repair.ps1', 'boot-repair.cmd')) {
            $src = Join-Path $Script:ToolkitDir $f
            if (Test-Path $src) {
                Copy-Item $src $peToolDir -Force
                Write-Log "  Injected $f into WinPE image." -Level SUCCESS
            } else {
                Write-Log "  $f not found -- skipping injection." -Level WARN
            }
        }

        # Add boot hint to startnet.cmd
        $startupHint = Join-Path $mountDir 'Windows\System32\startnet.cmd'
        try {
            $existing = Get-Content $startupHint -Raw -ErrorAction SilentlyContinue
            if ($existing -notmatch 'winboot-rescue') {
                Add-Content -Path $startupHint -Value "`r`necho." -Encoding ASCII
                Add-Content -Path $startupHint -Value 'echo  winboot-rescue: X:\Windows\System32\winboot-rescue\boot-repair.cmd' -Encoding ASCII
                Add-Content -Path $startupHint -Value 'echo  OR from USB root: navigate to USB letter and run boot-repair.cmd' -Encoding ASCII
            }
        } catch {}

        Write-Log "Committing WinPE image changes..." -Level STEP
        $dismUnmount = & dism.exe /Unmount-Image /MountDir:"$mountDir" /Commit 2>&1
        $dismUnmount | ForEach-Object { Write-Log "  $_" -Level INFO }
        if ($LASTEXITCODE -eq 0) {
            Write-Log "WinPE image customized and committed." -Level SUCCESS
        } else {
            Write-Log "DISM unmount had errors -- image may still work." -Level WARN
            & dism.exe /Unmount-Image /MountDir:"$mountDir" /Discard 2>&1 | Out-Null
        }
    }

    # Copy WinPE media files to USB
    Write-Log "Copying WinPE media to USB $UsbLetter ..." -Level STEP
    $mediaDir = Join-Path $wpePe 'media'
    & robocopy.exe "$mediaDir" "$UsbLetter\" /E /NFL /NDL /NJH /NJS 2>&1 | Out-Null
    Write-Log "robocopy exit: $LASTEXITCODE (0-7 = success)" -Level INFO

    # Run bcdboot for MBR boot sector
    $peWinDir = Join-Path $UsbLetter 'Windows'
    if (Test-Path $peWinDir) {
        Write-Log "Running bcdboot to write boot sector..." -Level STEP
        & bcdboot.exe "$peWinDir" /s "$UsbLetter" /f ALL 2>&1 | ForEach-Object { Write-Log "  $_" -Level INFO }
    }

    Copy-ToolkitToUsb -UsbLetter $UsbLetter
    Write-Log "WinPE USB creation complete!" -Level SUCCESS
    return $true
}

# ============================================================
# SECTION 5: MODE B -- WinRE from local machine
# ============================================================
function Build-WinReUsb {
    param([string]$UsbLetter)

    Write-Log "Building WinRE USB from local Windows Recovery Environment..." -Level SECTION

    $winReWim = Find-WinReWim
    if (-not $winReWim) {
        Write-Log "WinRE.wim not found. Cannot use Mode B." -Level ERROR
        return $false
    }
    Write-Log "WinRE.wim found: $winReWim" -Level SUCCESS

    $usbSources = Join-Path $UsbLetter 'sources'
    if (-not (Test-Path $usbSources)) { New-Item -ItemType Directory $usbSources -Force | Out-Null }

    $bootWimDest = Join-Path $usbSources 'boot.wim'
    Write-Log "Copying WinRE.wim to USB as sources\boot.wim ..." -Level STEP
    try {
        Copy-Item $winReWim $bootWimDest -Force
        Write-Log "boot.wim copied." -Level SUCCESS
    } catch {
        Write-Log "Failed to copy WinRE.wim: $_" -Level ERROR
        return $false
    }

    # Mount and inject toolkit
    $mountDir = Join-Path $Script:WorkDir 're_mount'
    if (-not (Test-Path $mountDir)) { New-Item -ItemType Directory $mountDir -Force | Out-Null }

    Write-Log "Mounting WinRE image to inject toolkit..." -Level STEP
    & dism.exe /Mount-Image /ImageFile:"$bootWimDest" /Index:1 /MountDir:"$mountDir" 2>&1 | Out-Null
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
        & dism.exe /Unmount-Image /MountDir:"$mountDir" /Commit 2>&1 |
            ForEach-Object { Write-Log "  $_" -Level INFO }
    } else {
        Write-Log "Could not mount WinRE.wim -- toolkit will be on USB root only." -Level WARN
        & dism.exe /Unmount-Image /MountDir:"$mountDir" /Discard 2>&1 | Out-Null
    }

    # Copy Windows boot files
    $windowsBoot = "$env:SystemRoot\Boot"
    if (Test-Path $windowsBoot) {
        Write-Log "Copying Windows boot files to USB..." -Level STEP
        & robocopy.exe "$windowsBoot\EFI"  "$UsbLetter\EFI"  /E /NFL /NDL /NJH /NJS 2>&1 | Out-Null
        & robocopy.exe "$windowsBoot\PCAT" "$UsbLetter\Boot" /E /NFL /NDL /NJH /NJS 2>&1 | Out-Null
        Write-Log "Boot files copied." -Level SUCCESS
    }

    # bcdboot for UEFI + BIOS
    $systemRoot = $env:SystemRoot
    Write-Log "Running bcdboot to create boot entries on USB..." -Level STEP
    $bcdResult = & bcdboot.exe "$systemRoot" /s "$UsbLetter" /f ALL 2>&1
    $bcdResult | ForEach-Object { Write-Log "  $_" -Level INFO }
    if ($LASTEXITCODE -eq 0) {
        Write-Log "bcdboot completed." -Level SUCCESS
    } else {
        Write-Log "bcdboot returned $LASTEXITCODE -- USB may still boot." -Level WARN
    }

    # bootsect for BIOS MBR
    $bootsect = "$env:SystemRoot\System32\bootsect.exe"
    if (Test-Path $bootsect) {
        Write-Log "Running bootsect for MBR/BIOS compatibility..." -Level STEP
        & $bootsect /nt60 "$UsbLetter" /force /mbr 2>&1 |
            ForEach-Object { Write-Log "  $_" -Level INFO }
    }

    Copy-ToolkitToUsb -UsbLetter $UsbLetter
    Write-Log "WinRE USB creation complete!" -Level SUCCESS
    return $true
}

function Find-WinReWim {
    # Try reagentc
    try {
        $reagentOut = & reagentc.exe /info 2>&1 | Out-String
        if ($reagentOut -match 'Windows RE location\s*:\s*(.+\.wim)') {
            $wimPath = $Matches[1].Trim()
            if (Test-Path $wimPath) { return $wimPath }
        }
    } catch {}

    $candidates = @(
        "$env:SystemRoot\System32\Recovery\WinRE.wim",
        'C:\Recovery\WindowsRE\WinRE.wim',
        'D:\Recovery\WindowsRE\WinRE.wim'
    )
    foreach ($p in $candidates) {
        if (Test-Path $p -ErrorAction SilentlyContinue) { return $p }
    }

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
    param([string]$UsbLetter, [string]$IsoPath)

    Write-Log "Building USB from Windows ISO: $IsoPath" -Level SECTION

    if (-not (Test-Path $IsoPath)) {
        Write-Log "ISO file not found: $IsoPath" -Level ERROR
        return $false
    }

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
        Write-Log "Copying ISO contents to USB $UsbLetter (this may take several minutes)..." -Level STEP
        & robocopy.exe "$isoDrive\" "$UsbLetter\" /E /NFL /NDL /NJH /NJS 2>&1 | Out-Null
        Write-Log "robocopy exit: $LASTEXITCODE (0-7 = OK)" -Level INFO

        Copy-ToolkitToUsb -UsbLetter $UsbLetter

        $bootsect = "$isoDrive\boot\bootsect.exe"
        if (Test-Path $bootsect) {
            Write-Log "Running bootsect for BIOS boot support..." -Level STEP
            & $bootsect /nt60 "$UsbLetter" /force /mbr 2>&1 |
                ForEach-Object { Write-Log "  $_" -Level INFO }
        }

        Write-Log "ISO USB creation complete!" -Level SUCCESS
        return $true
    } finally {
        try { Dismount-DiskImage -ImagePath $IsoPath | Out-Null } catch {}
        Write-Log "ISO unmounted." -Level INFO
    }
}

# ============================================================
# SECTION 7: COPY TOOLKIT TO USB
# ============================================================
function Copy-ToolkitToUsb {
    param([string]$UsbLetter)

    Write-Log "Copying winboot-rescue toolkit to USB root..." -Level STEP

    $anyMissing = $false
    foreach ($f in @('boot-repair.ps1', 'boot-repair.cmd')) {
        $src  = Join-Path $Script:ToolkitDir $f
        $dest = Join-Path $UsbLetter $f

        if (Test-Path $src) {
            Copy-Item $src $dest -Force
            Write-Log "  Copied: $f" -Level SUCCESS
        } else {
            Write-Log "  NOT FOUND: $f (expected at $src)" -Level WARN
            $anyMissing = $true
        }
    }

    $readmeContent = @"
WINBOOT-RESCUE TOOLKIT
======================

Files on this USB:
  boot-repair.cmd  <- Launch this from WinRE/WinPE command prompt
  boot-repair.ps1  <- Main PowerShell script

HOW TO USE:
  1. Boot from this USB (change boot order in BIOS/UEFI)
  2. When WinRE/WinPE loads, open Command Prompt
  3. Find this USB drive letter:
     > diskpart
     > list volume
     > exit
  4. Navigate to USB and run:
     > E:\boot-repair.cmd   (replace E: with your USB letter)
  5. Choose option [1] first to collect diagnostics
  6. Then choose [2] or [4]/[5] based on your system type

  If injected into WinPE image, also available at:
     X:\Windows\System32\winboot-rescue\boot-repair.cmd

FIRST TIME: Always run option [1] (diagnostics) before attempting repair.

GitHub: https://github.com/Gzeu/winboot-rescue
"@
    Set-Content -Path (Join-Path $UsbLetter 'README-BOOT-REPAIR.txt') `
        -Value $readmeContent -Encoding UTF8 -ErrorAction SilentlyContinue
    Write-Log "  Created README-BOOT-REPAIR.txt" -Level INFO

    if (-not $anyMissing) {
        Write-Log "All toolkit files copied to USB." -Level SUCCESS
    }
}

# ============================================================
# SECTION 8: VERIFY USB
# ============================================================
function Test-UsbBoot {
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
        $status = if ($exists) { 'SUCCESS' } elseif ($c.Required) { $allGood = $false; 'ERROR' } else { 'WARN' }
        $icon   = if ($exists) { '[OK]' } else { '[MISSING]' }
        Write-Log "  $icon $($c.Label)" -Level $status
    }

    if ($allGood) {
        Write-Log "USB verification passed. Drive appears ready." -Level SUCCESS
    } else {
        Write-Log "USB verification: some required files missing. USB may not boot correctly." -Level WARN
    }

    return $allGood
}

# ============================================================
# SECTION 9: MAIN MENU
# ============================================================
function Show-Menu {
    $adkInfo = Find-AdkPath
    $hasAdk  = $null -ne $adkInfo
    $hasCopype = $hasAdk -and $null -ne $adkInfo.CopypePath

    Write-Host "`n$('='*60)" -ForegroundColor White
    Write-Host "  WINBOOT-RESCUE -- USB CREATOR" -ForegroundColor Cyan
    Write-Host "$('='*60)" -ForegroundColor White

    if ($hasAdk -and $hasCopype) {
        Write-Host "  ADK installed: YES -- Mode A available" -ForegroundColor Green
    } elseif ($hasAdk -and -not $hasCopype) {
        Write-Host "  ADK installed: YES -- but WinPE Add-on MISSING (Mode A unavailable)" -ForegroundColor Yellow
        Write-Host "  Install Add-on: https://learn.microsoft.com/windows-hardware/get-started/adk-install" -ForegroundColor DarkGray
    } else {
        Write-Host "  ADK installed: NO -- Mode A unavailable" -ForegroundColor Yellow
    }

    Write-Host "$('-'*60)" -ForegroundColor DarkGray
    $modeAColor = if ($hasCopype) { 'Green' } else { 'DarkGray' }
    Write-Host "  [1] Mode A -- WinPE USB from ADK (best, full WinPE)$(if (-not $hasCopype){' [WinPE Add-on required]'})" -ForegroundColor $modeAColor
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
            if (-not $hasCopype) {
                Write-Log "WinPE Add-on not installed. Mode A unavailable." -Level ERROR
                Write-Host "`n  Download and install the WinPE Add-on for ADK from:" -ForegroundColor Yellow
                Write-Host "  https://learn.microsoft.com/windows-hardware/get-started/adk-install" -ForegroundColor Yellow
            } else {
                $usb = Select-UsbDrive
                if ($usb) {
                    $usbLetter = Format-UsbDrive -PartitionStyle 'MBR'
                    if ($usbLetter) {
                        Build-WinPeUsb -UsbLetter $usbLetter -AdkInfo $adkInfo | Out-Null
                        Test-UsbBoot   -UsbLetter $usbLetter | Out-Null
                    }
                }
            }
        }
        '2' {
            $usb = Select-UsbDrive
            if ($usb) {
                $usbLetter = Format-UsbDrive -PartitionStyle 'MBR'
                if ($usbLetter) {
                    Build-WinReUsb -UsbLetter $usbLetter | Out-Null
                    Test-UsbBoot   -UsbLetter $usbLetter | Out-Null
                }
            }
        }
        '3' {
            $isoPath = (Read-Host "Enter full path to Windows ISO file").Trim('"')
            if (Test-Path $isoPath) {
                $usb = Select-UsbDrive
                if ($usb) {
                    $usbLetter = Format-UsbDrive -PartitionStyle 'MBR'
                    if ($usbLetter) {
                        Build-IsoUsb -UsbLetter $usbLetter -IsoPath $isoPath | Out-Null
                        Test-UsbBoot -UsbLetter $usbLetter | Out-Null
                    }
                }
            } else {
                Write-Log "ISO not found: $isoPath" -Level ERROR
            }
        }
        '4' {
            $letter    = Read-Host "Enter existing USB drive letter (e.g. E)"
            $usbLetter = "$($letter.Trim(':')):"
            if (Test-Path $usbLetter) {
                Copy-ToolkitToUsb -UsbLetter $usbLetter
                Test-UsbBoot      -UsbLetter $usbLetter | Out-Null
            } else {
                Write-Log "Drive $usbLetter not found." -Level ERROR
            }
        }
        '5' {
            $letter    = Read-Host "Enter USB drive letter to verify (e.g. E)"
            $usbLetter = "$($letter.Trim(':')):"
            Test-UsbBoot -UsbLetter $usbLetter | Out-Null
        }
        '0' {
            Write-Log "User exited." -Level INFO
            return
        }
        default { Write-Host "  Invalid option. Enter 0-5." -ForegroundColor Red }
    }

    Write-Host "`n[INFO] Log file: $Script:LogFile" -ForegroundColor DarkGray
    Wait-KeyPress 'Press Enter to return to menu...'
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

    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Log "Not running as Administrator! Re-launch as Admin." -Level ERROR
        Write-Host "`nRight-click the script and choose 'Run as administrator'" -ForegroundColor Yellow
        Read-Host "Press Enter to exit" | Out-Null
        exit 1
    }

    Write-Log "Script directory: $Script:ToolkitDir" -Level INFO
    Write-Log "Checking for toolkit files..." -Level INFO

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
    Write-Host "`nDone. Log saved to: $Script:LogFile" -ForegroundColor Green
    Wait-KeyPress 'Press Enter to exit...'
}

Main
