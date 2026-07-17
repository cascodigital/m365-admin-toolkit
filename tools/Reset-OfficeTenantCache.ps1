<#
.SYNOPSIS
    Resets per-user Office, Outlook, OneDrive, and WAM/AAD tenant cache.

.DESCRIPTION
    Cleans the current Windows user's Microsoft 365 identity cache when Outlook
    or OneDrive keeps trying to authenticate against an old tenant after a
    domain or mailbox migration.

    This script is intentionally per-user. Run it while logged on as the
    affected Windows user. Running it elevated as another account cleans the
    wrong HKCU/AppData profile.

    What it removes:
    - Outlook profiles and local Outlook cache for the current user
    - Office identity cache under HKCU
    - OneAuth and IdentityCache folders
    - AAD Broker TokenBroker and LocalState folders
    - OneDrive Business account and tenant cache
    - Credential Manager entries matching Microsoft 365/Office/OneDrive terms

    What it does not remove:
    - The Office installation
    - Windows itself
    - System-wide Azure AD join state
    - User documents

.PARAMETER ExtraCredentialPattern
    Optional extra regex pattern used to delete Credential Manager targets,
    for example an old tenant name, domain, or tenant id.

.PARAMETER SkipOneDriveReset
    Do not run OneDrive.exe /reset at the end.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Reset-OfficeTenantCache.ps1

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Reset-OfficeTenantCache.ps1 -ExtraCredentialPattern "oldtenant|contoso.com|00000000-0000-0000-0000-000000000000"

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Reset-OfficeTenantCache.ps1 -WhatIf

.NOTES
    Author: Casco Digital
    Compatibility: Windows PowerShell 5.1+, Windows 10/11

    Recommended flow:
    1. Log on as the affected Windows user.
    2. Run this script in a normal PowerShell session.
    3. Reboot Windows.
    4. Remove stale entries from Settings > Accounts > Access work or school.
    5. Open OneDrive first, then create a fresh Outlook profile.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$ExtraCredentialPattern,
    [switch]$SkipOneDriveReset
)

$ErrorActionPreference = "SilentlyContinue"

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host $Message -ForegroundColor Cyan
}

function Remove-PathSafe {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }

    if (Test-Path -LiteralPath $Path) {
        if ($PSCmdlet.ShouldProcess($Path, "Remove path")) {
            Remove-Item -LiteralPath $Path -Recurse -Force
        }
        Write-Host "  - Removed path: $Path" -ForegroundColor Gray
    }
}

function Remove-PathPatternSafe {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }

    $items = Get-Item -Path $Path
    foreach ($item in $items) {
        if ($PSCmdlet.ShouldProcess($item.FullName, "Remove path")) {
            Remove-Item -LiteralPath $item.FullName -Recurse -Force
        }
        Write-Host "  - Removed path: $($item.FullName)" -ForegroundColor Gray
    }
}

function Remove-RegistryKeySafe {
    param([string]$Path)

    if (Test-Path -LiteralPath $Path) {
        if ($PSCmdlet.ShouldProcess($Path, "Remove registry key")) {
            Remove-Item -LiteralPath $Path -Recurse -Force
        }
        Write-Host "  - Removed registry key: $Path" -ForegroundColor Gray
    }
}

Write-Host "==== RESET OFFICE / OUTLOOK / ONEDRIVE TENANT CACHE ====" -ForegroundColor Yellow
Write-Host "Current user: $env:USERDOMAIN\$env:USERNAME"
Write-Host "Run as the affected Windows user. Admin-as-another-user cleans the wrong profile." -ForegroundColor Yellow

Write-Step "[1/7] Closing Microsoft 365 processes..."
$processes = @(
    "OUTLOOK",
    "OneDrive",
    "Teams",
    "ms-teams",
    "WINWORD",
    "EXCEL",
    "POWERPNT",
    "ONENOTE",
    "OfficeClickToRun",
    "Microsoft.AAD.BrokerPlugin"
)

foreach ($process in $processes) {
    Get-Process -Name $process | Stop-Process -Force
}

$oneDriveExecutables = @(
    "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe",
    "$env:ProgramFiles\Microsoft OneDrive\OneDrive.exe",
    "${env:ProgramFiles(x86)}\Microsoft OneDrive\OneDrive.exe"
) | Where-Object { Test-Path -LiteralPath $_ }

foreach ($oneDriveExe in $oneDriveExecutables) {
    if ($PSCmdlet.ShouldProcess($oneDriveExe, "Shutdown OneDrive")) {
        & $oneDriveExe /shutdown
    }
}

Start-Sleep -Seconds 3

Write-Step "[2/7] Exporting HKCU registry backups..."
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$backupDir = Join-Path $env:USERPROFILE "Desktop\office-tenant-cache-backup-$timestamp"

if ($PSCmdlet.ShouldProcess($backupDir, "Create backup directory")) {
    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
}

if ($PSCmdlet.ShouldProcess("HKCU\Software\Microsoft\OneDrive", "Export registry backup")) {
    reg export "HKCU\Software\Microsoft\OneDrive" (Join-Path $backupDir "onedrive-hkcu.reg") /y | Out-Null
}

if ($PSCmdlet.ShouldProcess("HKCU\Software\Microsoft\Office\16.0", "Export registry backup")) {
    reg export "HKCU\Software\Microsoft\Office\16.0" (Join-Path $backupDir "office16-hkcu.reg") /y | Out-Null
}

Write-Host "  - Backup folder: $backupDir" -ForegroundColor Gray

Write-Step "[3/7] Removing OneDrive Business account and tenant cache..."
$oneDriveRegistryKeys = @(
    "HKCU:\Software\Microsoft\OneDrive\Accounts",
    "HKCU:\Software\Microsoft\OneDrive\Tenants"
)

foreach ($key in $oneDriveRegistryKeys) {
    Remove-RegistryKeySafe -Path $key
}

$oneDrivePaths = @(
    "$env:LOCALAPPDATA\Microsoft\OneDrive\settings\Business*",
    "$env:LOCALAPPDATA\Microsoft\OneDrive\logs\Business*",
    "$env:LOCALAPPDATA\Microsoft\OneDrive\ListSync",
    "$env:LOCALAPPDATA\Microsoft\OneDrive\Business*"
)

foreach ($path in $oneDrivePaths) {
    Remove-PathPatternSafe -Path $path
}

Write-Step "[4/7] Removing Office identity, WAM, and AAD Broker cache..."
$identityPaths = @(
    "$env:LOCALAPPDATA\Microsoft\OneAuth",
    "$env:LOCALAPPDATA\Microsoft\IdentityCache",
    "$env:LOCALAPPDATA\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\AC\TokenBroker",
    "$env:LOCALAPPDATA\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\LocalState"
)

foreach ($path in $identityPaths) {
    Remove-PathSafe -Path $path
}

Write-Step "[5/7] Removing Outlook profiles and Office cache..."
$officeRegistryKeys = @(
    "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles",
    "HKCU:\Software\Microsoft\Office\16.0\Common\Identity",
    "HKCU:\Software\Microsoft\Office\16.0\Outlook\AutoDiscover",
    "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
)

foreach ($key in $officeRegistryKeys) {
    Remove-RegistryKeySafe -Path $key
}

$officePaths = @(
    "$env:LOCALAPPDATA\Microsoft\Outlook",
    "$env:APPDATA\Microsoft\Outlook",
    "$env:LOCALAPPDATA\Microsoft\Office\16.0\OfficeFileCache"
)

foreach ($path in $officePaths) {
    Remove-PathSafe -Path $path
}

Write-Step "[6/7] Removing matching Credential Manager entries..."
$patterns = @(
    "Office",
    "Outlook",
    "MicrosoftOffice",
    "ADAL",
    "MSOID",
    "Exchange",
    "OneDrive"
)

if (-not [string]::IsNullOrWhiteSpace($ExtraCredentialPattern)) {
    $patterns += $ExtraCredentialPattern
}

$credentialRegex = ($patterns | ForEach-Object { "($_)" }) -join "|"
$credentialTargets = (cmdkey /list) |
    Select-String -Pattern "^\s*(Target|Destino|Alvo):\s*(.+)$" |
    ForEach-Object { $_.Matches[0].Groups[2].Value.Trim() } |
    Where-Object { $_ -match $credentialRegex }

foreach ($target in $credentialTargets) {
    if ($PSCmdlet.ShouldProcess($target, "Delete Credential Manager target")) {
        cmdkey /delete:$target | Out-Null
    }
    Write-Host "  - Deleted credential: $target" -ForegroundColor Gray
}

Write-Step "[7/7] Resetting OneDrive..."
if (-not $SkipOneDriveReset) {
    $oneDriveExe = $oneDriveExecutables | Select-Object -First 1
    if ($oneDriveExe) {
        if ($PSCmdlet.ShouldProcess($oneDriveExe, "Run OneDrive reset")) {
            & $oneDriveExe /reset
        }
        Write-Host "  - OneDrive reset requested: $oneDriveExe" -ForegroundColor Gray
    } else {
        Write-Host "  - OneDrive.exe not found; skipped reset." -ForegroundColor DarkYellow
    }
} else {
    Write-Host "  - Skipped by parameter." -ForegroundColor DarkYellow
}

Write-Host ""
Write-Host "==== DONE ====" -ForegroundColor Green
Write-Host "Reboot Windows before opening OneDrive or Outlook." -ForegroundColor Yellow
Write-Host "Then review Settings > Accounts > Access work or school and remove stale tenant entries." -ForegroundColor Yellow
Write-Host "If the old tenant still appears, run: dsregcmd /status" -ForegroundColor Yellow
