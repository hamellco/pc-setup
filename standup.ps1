# =============================================================
# Script  : standup.ps1
# Author  : Hamellco
# Purpose : Personal PC standup script
# Version : See last Git commit date below
# =============================================================

# =============================================================
# --- CONFIGURATION - Fill in before running               ---
# =============================================================

$myToken = "github_pat_YOURTOKENHERE"

# =============================================================

# --- Capture raw arguments to support --dryrun style flag ---
param (
    [string[]]$Args = @()
)

# --- Check if --dryrun was passed ---
$DryRun = $Args -contains "--dryrun"

# --- Get the last Git commit date and use it as the version ---
# --- Falls back to today's date if git is not available yet on a new PC ---
try {
    $version = git log -1 --format="%cd" --date=format:"%Y-%m-%d" 2>$null
    if (-not $version) { $version = Get-Date -Format "yyyy-MM-dd" }
} catch {
    $version = Get-Date -Format "yyyy-MM-dd"
}

# --- Reusable function to print centered text ---
function Write-Centered($text, $color = "White") {
    $consoleWidth = $Host.UI.RawUI.WindowSize.Width
    $padding      = [math]::Max(0, [math]::Floor(($consoleWidth - $text.Length) / 2))
    $centered     = (" " * $padding) + $text
    Write-Host $centered -ForegroundColor $color
}

# --- Reusable function to print a section header ---
# --- Text is left aligned inside fixed width wings so ### always lines up ---
function Write-SectionHeader($title) {
    $innerWidth = 40
    $padded     = $title.PadRight($innerWidth)
    $header     = "### $padded ###"
    Write-Host ""
    Write-Host $header -ForegroundColor Cyan
    Write-Host ""
}

# --- Reusable function to save state after each completed section ---
function Save-State {
    if (-not $DryRun) {
        $state | ConvertTo-Json | Set-Content $stateFile
    }
}

# =============================================================
# --- Hamellco PC Standup Script Banner ---
# =============================================================

$bannerLine = "============================================="
$titleText  = "Hamellco PC Standup Script"
$verText    = "Version: $version"
$dryText    = "*** DRY RUN - No changes will be saved ***"

Write-Host ""
Write-Centered $bannerLine "Cyan"
Write-Centered $titleText "Yellow"
Write-Centered $verText "Yellow"
if ($DryRun) {
    Write-Centered $dryText "Red"
}
Write-Centered $bannerLine "Cyan"
Write-Host ""

# =============================================================
# --- Local Folder Setup ---
# --- All state and identity files live in C:\Temp\hamellco  ---
# =============================================================

$hamellcoFolder = "C:\Temp\hamellco"
$identityFile   = Join-Path $hamellcoFolder "pc-identity.json"
$stateFile      = Join-Path $hamellcoFolder "state.json"

# --- Create C:\Temp\hamellco if it doesn't exist ---
if (-not (Test-Path $hamellcoFolder)) {
    if ($DryRun) {
        Write-Host "  [DRY RUN] Would create folder: $hamellcoFolder" -ForegroundColor DarkYellow
        Write-Host ""
    } else {
        New-Item -ItemType Directory -Path $hamellcoFolder | Out-Null
        Write-Host "  Created folder: $hamellcoFolder" -ForegroundColor DarkGray
        Write-Host ""
    }
}

# =============================================================
# --- Sync Identity File from GitHub ---
# --- Pulls the latest pc-identity.json from the repo before  ---
# --- doing any identity checks so we always have the full    ---
# --- list of known machines                                  ---
# =============================================================

Write-SectionHeader "Syncing Identity File"

if (-not $myToken -or $myToken -eq "github_pat_YOURTOKENHERE") {
    # --- Token not filled in - skip sync and use local file ---
    Write-Host "  No token set - skipping sync, using local identity file." -ForegroundColor DarkYellow
    Write-Host ""
} elseif ($DryRun) {
    Write-Host "  [DRY RUN] Would pull pc-identity.json from GitHub to $identityFile" -ForegroundColor DarkYellow
    Write-Host ""
} else {
    try {
        $identityUrl     = "https://raw.githubusercontent.com/hamellco/pc-setup/main/standup/pc-identity.json"
        $identityHeaders = @{ Authorization = "token $myToken" }
        Invoke-WebRequest -Uri $identityUrl -Headers $identityHeaders -OutFile $identityFile
        Write-Host "  Identity file synced from GitHub." -ForegroundColor Green
        Write-Host ""
    } catch {
        Write-Host "  Could not sync identity file from GitHub - using local copy if available." -ForegroundColor Yellow
        Write-Host ""
    }
}

# --- Load existing state or start fresh ---
$state = [PSCustomObject]@{
    identityDone        = $false
    activationDone      = $false
    personalizationDone = $false
}

if (Test-Path $stateFile) {
    $loaded = Get-Content $stateFile | ConvertFrom-Json
    $state.identityDone        = $loaded.identityDone
    $state.activationDone      = $loaded.activationDone
    $state.personalizationDone = $loaded.personalizationDone
    Write-Host "  Resuming from previous session..." -ForegroundColor DarkYellow
    Write-Host ""
}

# =============================================================
# --- PC Identity ---
# =============================================================

if ($state.identityDone) {
    Write-SectionHeader "PC Identity"
    Write-Host "  Already completed - skipping." -ForegroundColor DarkGray
    Write-Host ""
} else {
    Write-SectionHeader "PC Identity"

    $pcName = $env:COMPUTERNAME

    $identityData = @()
    if (Test-Path $identityFile) {
        $identityData = Get-Content $identityFile | ConvertFrom-Json
    }

    $existingIdentity = $identityData | Where-Object { $_.windowsName -eq $pcName }

    if ($pcName -eq "Scheherazade") {
        if ($existingIdentity) {
            Write-Host "  Welcome back, Scheherazade." -ForegroundColor Green
            Write-Host ""
            $state.identityDone = $true
            Save-State
        } else {
            $confirm = Read-Host "  Are you setting up a new Scheherazade? (yes/no)"
            if ($confirm -eq "yes") {
                $newEntry = [PSCustomObject]@{
                    windowsName = $pcName
                    identity    = "Scheherazade"
                    type        = "Primary"
                    assigned    = (Get-Date -Format "yyyy-MM-dd")
                }
                if ($DryRun) {
                    Write-Host "  [DRY RUN] Would register Scheherazade with the following data:" -ForegroundColor DarkYellow
                    Write-Host ($newEntry | ConvertTo-Json) -ForegroundColor DarkYellow
                } else {
                    $identityData += $newEntry
                    $identityData | ConvertTo-Json -Depth 3 | Set-Content $identityFile
                    Write-Host "  Scheherazade has been registered." -ForegroundColor Green
                    $state.identityDone = $true
                    Save-State
                }
                Write-Host ""
            } else {
                Write-Host "  OK - no changes made." -ForegroundColor Yellow
                Write-Host ""
            }
        }
    } else {
        if ($existingIdentity) {
            Write-Host "  Welcome back, $($existingIdentity.identity)" -ForegroundColor Green
            Write-Host ""
            $state.identityDone = $true
            Save-State
        } else {
            $chassisTypes = (Get-WmiObject -Class Win32_SystemEnclosure).ChassisTypes
            $isLaptop     = $chassisTypes | Where-Object { $_ -in @(8,9,10,11,12,14,18,21) }
            $typeCode     = if ($isLaptop) { "LT" } else { "DT" }
            $date         = Get-Date -Format "yyyy-MM-dd"
            $newIdentity  = "Hamellco-$typeCode-$date"

            $newEntry = [PSCustomObject]@{
                windowsName = $pcName
                identity    = $newIdentity
                type        = if ($isLaptop) { "Laptop" } else { "Desktop" }
                assigned    = $date
            }

            if ($DryRun) {
                Write-Host "  [DRY RUN] Would assign new identity: $newIdentity" -ForegroundColor DarkYellow
                Write-Host "  [DRY RUN] Would save the following data:" -ForegroundColor DarkYellow
                Write-Host ($newEntry | ConvertTo-Json) -ForegroundColor DarkYellow
            } else {
                $identityData += $newEntry
                $identityData | ConvertTo-Json -Depth 3 | Set-Content $identityFile
                Write-Host "  New PC detected. Assigned identity: $newIdentity" -ForegroundColor Cyan
                $state.identityDone = $true
                Save-State
            }
            Write-Host ""
        }
    }
}

# =============================================================
# --- Windows Activation Status ---
# =============================================================

if ($state.activationDone) {
    Write-SectionHeader "Windows Activation Status"
    Write-Host "  Already completed - skipping." -ForegroundColor DarkGray
    Write-Host ""
} else {
    Write-SectionHeader "Windows Activation Status"

    function Get-WindowsActivationStatus {
        $activation = Get-CimInstance -ClassName SoftwareLicensingProduct |
            Where-Object { $_.Name -like "*Windows*" -and $_.PartialProductKey } |
            Select-Object -First 1
        return $activation.LicenseStatus
    }

    function Get-ActivationLabel($code) {
        switch ($code) {
            1 { return "Activated" }
            2 { return "Out of Box Grace Period" }
            3 { return "Out of Tolerance Grace Period" }
            4 { return "Non-Genuine Grace Period" }
            5 { return "Notification" }
            6 { return "Extended Grace Period" }
            default { return "Unknown / Not Activated" }
        }
    }

    $licenseStatus = Get-WindowsActivationStatus
    $statusLabel   = Get-ActivationLabel $licenseStatus

    if ($licenseStatus -eq 1) {
        Write-Host "  Windows Status : $statusLabel" -ForegroundColor Green
        Write-Host ""
        $state.activationDone = $true
        Save-State
    } else {
        Write-Host "  Windows Status : $statusLabel" -ForegroundColor Red
        Write-Host ""

        if ($DryRun) {
            Write-Host "  [DRY RUN] Would launch Microsoft Activation Scripts (MAS) to activate Windows." -ForegroundColor DarkYellow
            Write-Host ""
        } else {
            Write-Host "  Opening Microsoft Activation Scripts (MAS)..." -ForegroundColor Yellow
            Write-Host "  Please complete the activation in the window that opens." -ForegroundColor Yellow
            Write-Host "  This script will continue automatically once that window closes." -ForegroundColor Yellow
            Write-Host ""

            Start-Process -FilePath "cmd.exe" -ArgumentList "/c irm https://get.activated.win | iex" -Verb RunAs -PassThru -Wait | Out-Null

            Write-Host "  Checking activation status again..." -ForegroundColor Cyan
            Write-Host ""

            $licenseStatus = Get-WindowsActivationStatus
            $statusLabel   = Get-ActivationLabel $licenseStatus

            if ($licenseStatus -eq 1) {
                Write-Host "  Windows Status : $statusLabel" -ForegroundColor Green
                Write-Host ""
                $state.activationDone = $true
                Save-State
            } else {
                Write-Host "  Windows Status : $statusLabel" -ForegroundColor Red
                Write-Host ""
                Write-Host "  WARNING: Windows is still not activated. You may need to activate manually." -ForegroundColor Red
                Write-Host ""
            }
        }
    }
}

# =============================================================
# --- Personalization ---
# =============================================================

if ($state.personalizationDone) {
    Write-SectionHeader "Personalization"
    Write-Host "  Already completed - skipping." -ForegroundColor DarkGray
    Write-Host ""
} else {
    Write-SectionHeader "Personalization"

    $personalizationKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    $accentColorKey     = "HKCU:\SOFTWARE\Microsoft\Windows\DWM"
    $taskbarKey         = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    $searchKey          = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search"

    # --- Nord Blue #5E81AC converted to Windows ABGR format ---
    # --- RGB #5E81AC = R:5E G:81 B:AC reversed to ABGR = FF AC 81 5E ---
    $nordBlueABGR = 0xFFAC815E

    if ($DryRun) {
        Write-Host "  [DRY RUN] Would apply the following personalization settings:" -ForegroundColor DarkYellow
        Write-Host ""
        Write-Host "  Appearance" -ForegroundColor DarkYellow
        Write-Host "  - App Mode             : Dark" -ForegroundColor DarkYellow
        Write-Host "  - System Mode          : Dark" -ForegroundColor DarkYellow
        Write-Host "  - Transparency         : On" -ForegroundColor DarkYellow
        Write-Host "  - Accent Color         : Nord Blue (#5E81AC)" -ForegroundColor DarkYellow
        Write-Host "  - Accent on Taskbar    : On" -ForegroundColor DarkYellow
        Write-Host "  - Accent on Title Bars : On" -ForegroundColor DarkYellow
        Write-Host ""
        Write-Host "  Taskbar Items" -ForegroundColor DarkYellow
        Write-Host "  - Search               : Hide" -ForegroundColor DarkYellow
        Write-Host "  - Task View            : Off" -ForegroundColor DarkYellow
        Write-Host "  - Widgets              : Off" -ForegroundColor DarkYellow
        Write-Host ""
        Write-Host "  Taskbar Behaviors" -ForegroundColor DarkYellow
        Write-Host "  - Alignment            : Left" -ForegroundColor DarkYellow
        Write-Host "  - Auto Hide            : Off" -ForegroundColor DarkYellow
        Write-Host "  - Show Badges          : On" -ForegroundColor DarkYellow
        Write-Host "  - Show Flashing        : On" -ForegroundColor DarkYellow
        Write-Host "  - Share Window         : On" -ForegroundColor DarkYellow
        Write-Host "  - Show Desktop Corner  : On" -ForegroundColor DarkYellow
        Write-Host "  - Combine Buttons      : Always" -ForegroundColor DarkYellow
        Write-Host "  - Smaller Buttons      : When taskbar is full" -ForegroundColor DarkYellow
        Write-Host ""
    } else {
        # --- Appearance ---
        Set-ItemProperty -Path $personalizationKey -Name "AppsUseLightTheme"    -Value 0 -Type DWord
        Set-ItemProperty -Path $personalizationKey -Name "SystemUsesLightTheme" -Value 0 -Type DWord
        Set-ItemProperty -Path $personalizationKey -Name "EnableTransparency"   -Value 1 -Type DWord
        Set-ItemProperty -Path $personalizationKey -Name "ColorPrevalence"      -Value 1 -Type DWord
        Set-ItemProperty -Path $accentColorKey     -Name "ColorPrevalence"      -Value 1 -Type DWord
        Set-ItemProperty -Path $accentColorKey     -Name "AccentColor"          -Value $nordBlueABGR -Type DWord
        Set-ItemProperty -Path $accentColorKey     -Name "AccentColorMenu"      -Value $nordBlueABGR -Type DWord

        # --- Broadcast color change to Windows so it applies live without reboot ---
        Add-Type -TypeDefinition @"
            using System;
            using System.Runtime.InteropServices;
            public class WinAPI {
                [DllImport("user32.dll", SetLastError = true)]
                public static extern IntPtr SendMessageTimeout(
                    IntPtr hWnd, uint Msg, UIntPtr wParam, string lParam,
                    uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);
            }
"@
        $HWND_BROADCAST   = [IntPtr]0xffff
        $WM_SETTINGCHANGE = 0x001A
        $result           = [UIntPtr]::Zero
        [WinAPI]::SendMessageTimeout($HWND_BROADCAST, $WM_SETTINGCHANGE, [UIntPtr]::Zero, "ImmersiveColorSet", 2, 5000, [ref]$result) | Out-Null

        Write-Host "  Appearance" -ForegroundColor White
        Write-Host "  - App Mode             : Dark" -ForegroundColor Green
        Write-Host "  - System Mode          : Dark" -ForegroundColor Green
        Write-Host "  - Transparency         : On" -ForegroundColor Green
        Write-Host "  - Accent Color         : Nord Blue (#5E81AC)" -ForegroundColor Green
        Write-Host "  - Accent on Taskbar    : On" -ForegroundColor Green
        Write-Host "  - Accent on Title Bars : On" -ForegroundColor Green
        Write-Host ""

        # --- Taskbar Items ---
        Set-ItemProperty -Path $searchKey  -Name "SearchboxTaskbarMode" -Value 0 -Type DWord
        Set-ItemProperty -Path $taskbarKey -Name "ShowTaskViewButton"   -Value 0 -Type DWord
        Set-ItemProperty -Path $taskbarKey -Name "TaskbarDa"            -Value 0 -Type DWord

        Write-Host "  Taskbar Items" -ForegroundColor White
        Write-Host "  - Search               : Hidden" -ForegroundColor Green
        Write-Host "  - Task View            : Off" -ForegroundColor Green
        Write-Host "  - Widgets              : Off" -ForegroundColor Green
        Write-Host ""

        # --- Taskbar Behaviors ---
        Set-ItemProperty -Path $taskbarKey -Name "TaskbarAl"          -Value 0 -Type DWord
        Set-ItemProperty -Path $taskbarKey -Name "AutoHideTaskbar"    -Value 0 -Type DWord
        Set-ItemProperty -Path $taskbarKey -Name "TaskbarBadges"      -Value 1 -Type DWord
        Set-ItemProperty -Path $taskbarKey -Name "TaskbarFlashing"    -Value 1 -Type DWord
        Set-ItemProperty -Path $taskbarKey -Name "TaskbarSn"          -Value 1 -Type DWord
        Set-ItemProperty -Path $taskbarKey -Name "TaskbarSd"          -Value 1 -Type DWord
        Set-ItemProperty -Path $taskbarKey -Name "TaskbarGlomLevel"   -Value 0 -Type DWord
        Set-ItemProperty -Path $taskbarKey -Name "MMTaskbarGlomLevel" -Value 0 -Type DWord
        Set-ItemProperty -Path $taskbarKey -Name "TaskbarSmallIcons"  -Value 2 -Type DWord

        Write-Host "  Taskbar Behaviors" -ForegroundColor White
        Write-Host "  - Alignment            : Left" -ForegroundColor Green
        Write-Host "  - Auto Hide            : Off" -ForegroundColor Green
        Write-Host "  - Show Badges          : On" -ForegroundColor Green
        Write-Host "  - Show Flashing        : On" -ForegroundColor Green
        Write-Host "  - Share Window         : On" -ForegroundColor Green
        Write-Host "  - Show Desktop Corner  : On" -ForegroundColor Green
        Write-Host "  - Combine Buttons      : Always" -ForegroundColor Green
        Write-Host "  - Smaller Buttons      : When taskbar is full" -ForegroundColor Green
        Write-Host ""

        # --- Restart Explorer to apply taskbar changes immediately ---
        Write-Host "  Restarting Explorer to apply taskbar changes..." -ForegroundColor Yellow
        Stop-Process -Name explorer -Force
        Write-Host ""

        $state.personalizationDone = $true
        Save-State

        Write-Host "  Personalization complete." -ForegroundColor Green
        Write-Host ""
        $restart = Read-Host "  A restart is recommended to apply all changes. Restart now? (yes/no)"
        if ($restart -eq "yes") {
            Write-Host ""
            Write-Host "  Restarting in 5 seconds..." -ForegroundColor Yellow
            Write-Host "  Run standup.ps1 again after reboot to continue setup." -ForegroundColor Yellow
            Write-Host ""
            Start-Sleep -Seconds 5
            Restart-Computer -Force
        } else {
            Write-Host ""
            Write-Host "  Skipping restart. Run standup.ps1 again when you are ready to continue." -ForegroundColor Yellow
            Write-Host ""
        }
    }
}

# =============================================================
# --- All Sections Complete ---
# =============================================================

if ($state.identityDone -and $state.activationDone -and $state.personalizationDone) {
    Write-Host ""
    Write-Centered "=============================================" "Cyan"
    Write-Centered "       All setup steps complete!            " "Green"
    Write-Centered "=============================================" "Cyan"
    Write-Host ""

    if (-not $DryRun) {
        Remove-Item $stateFile -Force
        Write-Host "  State file cleared. Ready for next setup." -ForegroundColor DarkGray
        Write-Host ""
    }
}