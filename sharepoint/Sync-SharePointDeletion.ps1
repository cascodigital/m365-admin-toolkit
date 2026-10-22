<#
.SYNOPSIS
    Deletes files from SharePoint document libraries that no longer exist in the local source.

.DESCRIPTION
    Intended for use after a migration (for example with the SharePoint Migration Tool)
    to keep the destination in sync with the source (local folder, file share, etc.).

    For each mapping (document library -> source folder) the script lists every file in
    the library (all pages, all subfolders), checks whether the same relative path exists
    in the source, and plans the deletion of files that are missing from the source.

    By default the script only SIMULATES: it logs what would be deleted and changes
    nothing. Deletion happens only with -Execute. Deleted items go to the site recycle bin.

    Safety guards (a library is skipped, nothing is deleted, unless noted):
    - Every source root is validated before anything else runs. If any source root is
      unreachable, not a folder, or empty, the whole run is aborted.
    - If the planned deletions for a library exceed -MaxDeletePercent of the files in the
      library, or exceed -MaxDelete files, the library is skipped unless -Force is used.
    - -WhatIf is supported and prevents any deletion even with -Execute.

.PARAMETER TenantId
    Tenant name or ID, e.g. contoso.onmicrosoft.com. The SharePoint host name is derived
    from it (contoso.sharepoint.com) when -SharePointHost is not given.

.PARAMETER SiteRelUrl
    Server-relative URL of the site, e.g. /sites/Finance.

.PARAMETER SharePointHost
    Optional SharePoint host name, e.g. contoso.sharepoint.com. Use it when the tenant ID
    is a GUID or a custom domain.

.PARAMETER Mapping
    Array of hashtables with the keys Library (document library display name) and
    LocalPath (source folder), e.g.
    @(@{ Library = 'Documents'; LocalPath = '\\fileserver\share\Documents' })

.PARAMETER Execute
    Actually delete files. Without this switch the script only simulates.
    Alias: ModoReal (the original setting name).

.PARAMETER MaxDeletePercent
    Maximum share of the files in a library that may be deleted in one run (default 10).

.PARAMETER MaxDelete
    Maximum number of files that may be deleted per library in one run (default 500).

.PARAMETER Force
    Ignore -MaxDeletePercent and -MaxDelete. The source root checks are never bypassed.

.PARAMETER LogPath
    Folder for the log file. Defaults to the script folder; falls back to $env:TEMP
    if that folder is not writable.

.EXAMPLE
    .\Sync-SharePointDeletion.ps1 -TenantId contoso.onmicrosoft.com -SiteRelUrl /sites/Finance `
        -Mapping @(@{ Library = 'Documents'; LocalPath = '\\fileserver\share\Documents' })
    # Simulation: logs the files that would be deleted.

.EXAMPLE
    .\Sync-SharePointDeletion.ps1 -TenantId contoso.onmicrosoft.com -SiteRelUrl /sites/Finance `
        -Mapping @(@{ Library = 'Documents'; LocalPath = '\\fileserver\share\Documents' }) -Execute
    # Deletes files missing from the source, within the safety thresholds.

.NOTES
    Version       : 2.0
    Requirements  : PowerShell 5.1+, Microsoft.Graph.Authentication, Microsoft.Graph.Sites,
                    Microsoft.Graph.Files

    Permissions (delegated):
    - Simulation: Sites.Read.All
    - -Execute  : Sites.ReadWrite.All
    For app-only automation prefer Sites.Selected, granting write access only to the
    target site.

    Path comparison is done on relative paths. Files whose names were changed during
    migration (for example invalid characters replaced) will not match the source and
    will be planned for deletion - review the simulation log before using -Execute.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter()]
    [string]$TenantId = "contoso.onmicrosoft.com",

    [Parameter()]
    [string]$SiteRelUrl = "/sites/SiteName",

    [Parameter()]
    [string]$SharePointHost,

    [Parameter()]
    [hashtable[]]$Mapping = @(
        @{ Library = "Documents"; LocalPath = "\\fileserver\share\Documents" },
        @{ Library = "HR";        LocalPath = "\\fileserver\share\HR" }
    ),

    [Parameter()]
    [Alias('ModoReal')]
    [switch]$Execute,

    [Parameter()]
    [ValidateRange(0, 100)]
    [double]$MaxDeletePercent = 10,

    [Parameter()]
    [ValidateRange(0, [int]::MaxValue)]
    [int]$MaxDelete = 500,

    [Parameter()]
    [switch]$Force,

    [Parameter()]
    [string]$LogPath
)

# ============================================================
# Log file
# ============================================================

function Resolve-OutputDirectory {
    param([string]$Path)

    $candidates = @()
    if (-not [string]::IsNullOrWhiteSpace($Path)) { $candidates += $Path }
    elseif ($PSScriptRoot) { $candidates += $PSScriptRoot }
    else { $candidates += (Get-Location).Path }
    $candidates += $env:TEMP

    foreach ($candidate in $candidates) {
        try {
            if (-not (Test-Path -LiteralPath $candidate)) {
                New-Item -ItemType Directory -Path $candidate -Force -ErrorAction Stop -WhatIf:$false | Out-Null
            }
            $probe = Join-Path -Path $candidate -ChildPath (".write_test_{0}.tmp" -f [guid]::NewGuid())
            [System.IO.File]::WriteAllText($probe, '')
            Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue -WhatIf:$false
            return $candidate
        }
        catch {
            Write-Warning "Log folder '$candidate' is not writable. Trying the next location."
        }
    }
    throw "No writable log folder found."
}

$logDirectory = Resolve-OutputDirectory -Path $LogPath
$LogFile = Join-Path -Path $logDirectory -ChildPath "Sync-SharePointDeletion_$(Get-Date -Format 'yyyyMMdd_HHmm').log"

function Write-Log {
    param([string]$Message)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | $Message"
    Write-Host $line
    Add-Content -Path $LogFile -Value $line -WhatIf:$false
}

$modeText = if ($Execute -and -not $WhatIfPreference) { 'EXECUTE' } else { 'SIMULATION' }
Write-Log "Log file: $LogFile"
Write-Log "Mode: $modeText | MaxDeletePercent: $MaxDeletePercent | MaxDelete: $MaxDelete | Force: $([bool]$Force)"

# ============================================================
# Source validation (before connecting or listing anything)
# ============================================================

$sourceErrors = @()
foreach ($entry in $Mapping) {
    if (-not $entry.Library -or -not $entry.LocalPath) {
        $sourceErrors += "Invalid mapping entry (Library and LocalPath are required)."
        continue
    }

    $root = $entry.LocalPath
    if (-not (Test-Path -LiteralPath $root -PathType Container)) {
        $sourceErrors += "Source root not reachable or not a folder: $root (library '$($entry.Library)')"
        continue
    }

    $firstChild = Get-ChildItem -LiteralPath $root -Force -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -eq $firstChild) {
        $sourceErrors += "Source root is empty or cannot be read: $root (library '$($entry.Library)')"
    }
}

if ($sourceErrors.Count -gt 0) {
    foreach ($msg in $sourceErrors) { Write-Log "ABORT: $msg" }
    Write-Log "ABORT: Source validation failed. Nothing was changed."
    exit 2
}

# ============================================================
# Modules and connection
# ============================================================

foreach ($moduleName in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Sites', 'Microsoft.Graph.Files')) {
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Write-Log "ERROR: Module '$moduleName' is not installed. Run: Install-Module $moduleName -Scope CurrentUser"
        exit 1
    }
}

$scopes = if ($Execute) { @("Sites.ReadWrite.All") } else { @("Sites.Read.All") }

try {
    Connect-MgGraph -TenantId $TenantId -Scopes $scopes -ErrorAction Stop
    Write-Log "Connected to Microsoft Graph - tenant: $TenantId"
} catch {
    Write-Log "ERROR: Could not connect to Microsoft Graph - $_"
    exit 1
}

try {
    if ([string]::IsNullOrWhiteSpace($SharePointHost)) {
        $SharePointHost = ($TenantId -replace "\.onmicrosoft\.com$", "") + ".sharepoint.com"
    }
    $site = Get-MgSite -SiteId "${SharePointHost}:${SiteRelUrl}" -ErrorAction Stop
    Write-Log "Site found: $($site.DisplayName) [$($site.Id)]"
} catch {
    Write-Log "ERROR: Site not found - $_"
    Disconnect-MgGraph | Out-Null
    exit 1
}

# ============================================================
# Process each library
# ============================================================

$totalDeleted = 0
$totalPlanned = 0

foreach ($entry in $Mapping) {
    $libName   = $entry.Library
    $localRoot = $entry.LocalPath

    Write-Log "=== Starting: $libName ==="

    try {
        $drive = Get-MgSiteDrive -SiteId $site.Id -All -ErrorAction Stop |
                 Where-Object { $_.Name -eq $libName } |
                 Select-Object -First 1
        if (-not $drive) { throw "Library '$libName' not found." }
    } catch {
        Write-Log "ERROR locating library '$libName': $_"
        Write-Log "=== Skipping: $libName ==="
        continue
    }

    # List every file in the library (all pages, breadth-first over folders)
    try {
        $allFiles = New-Object System.Collections.Generic.List[object]
        $queue = [System.Collections.Queue]::new()

        $rootItems = Get-MgDriveItemChild -DriveId $drive.Id -DriveItemId "root" -All -ErrorAction Stop
        foreach ($i in $rootItems) { $queue.Enqueue(@{ Item = $i; Path = "" }) }

        while ($queue.Count -gt 0) {
            $queued   = $queue.Dequeue()
            $current  = $queued.Item
            $basePath = $queued.Path

            if ($null -ne $current.Folder) {
                if ($current.Folder.ChildCount -gt 0) {
                    $children = Get-MgDriveItemChild -DriveId $drive.Id -DriveItemId $current.Id -All -ErrorAction Stop
                    foreach ($c in $children) {
                        $queue.Enqueue(@{ Item = $c; Path = "$basePath\$($current.Name)" })
                    }
                }
            } elseif ($null -ne $current.File) {
                $allFiles.Add(@{ Id = $current.Id; Name = $current.Name; RelPath = "$basePath\$($current.Name)".TrimStart("\") })
            }
        }
    } catch {
        Write-Log "ERROR listing files in '$libName': $_"
        Write-Log "=== Skipping: $libName ==="
        continue
    }

    Write-Log "Files in the cloud ($libName): $($allFiles.Count)"

    # Re-check the source root right before comparing (share may have dropped)
    if (-not (Test-Path -LiteralPath $localRoot -PathType Container)) {
        Write-Log "ABORT: Source root '$localRoot' is no longer reachable. Nothing deleted in '$libName'."
        Write-Log "=== Skipping: $libName ==="
        continue
    }

    # Plan deletions
    $toDelete = @($allFiles | Where-Object { -not (Test-Path -LiteralPath (Join-Path $localRoot $_.RelPath)) })
    $totalPlanned += $toDelete.Count

    $percent = if ($allFiles.Count -gt 0) { [math]::Round(($toDelete.Count / $allFiles.Count) * 100, 2) } else { 0 }
    Write-Log "Planned deletions ($libName): $($toDelete.Count) of $($allFiles.Count) ($percent%)"

    if ($toDelete.Count -eq 0) {
        Write-Log "=== Completed: $libName ==="
        continue
    }

    # Threshold guard
    $overThreshold = ($percent -gt $MaxDeletePercent) -or ($toDelete.Count -gt $MaxDelete)
    if ($overThreshold -and -not $Force) {
        foreach ($f in $toDelete) { Write-Log "BLOCKED (over threshold) - would delete: $($f.RelPath)" }
        Write-Log "ABORT: $($toDelete.Count) deletions ($percent%) exceed the limits (MaxDelete $MaxDelete, MaxDeletePercent $MaxDeletePercent). Nothing deleted in '$libName'. Review the list and rerun with -Force if it is correct."
        Write-Log "=== Skipping: $libName ==="
        continue
    }

    foreach ($f in $toDelete) {
        if (-not $Execute) {
            Write-Log "SIMULATION - would delete: $($f.RelPath)"
            continue
        }

        if ($PSCmdlet.ShouldProcess("$libName\$($f.RelPath)", "Delete SharePoint file")) {
            try {
                Remove-MgDriveItem -DriveId $drive.Id -DriveItemId $f.Id -ErrorAction Stop
                $totalDeleted++
                Write-Log "DELETED: $($f.RelPath)"
            } catch {
                Write-Log "ERROR deleting '$($f.RelPath)': $_"
            }
        } else {
            Write-Log "WHATIF - would delete: $($f.RelPath)"
        }
    }

    Write-Log "=== Completed: $libName ==="
}

Disconnect-MgGraph | Out-Null

if ($Execute -and -not $WhatIfPreference) {
    Write-Log "Script finished. Mode: EXECUTE - $totalDeleted of $totalPlanned planned file(s) deleted."
} else {
    Write-Log "Script finished. Mode: SIMULATION - $totalPlanned file(s) would be deleted; nothing was changed."
}
