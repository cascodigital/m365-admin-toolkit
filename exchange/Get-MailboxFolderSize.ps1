#requires -Version 5.1
<#
.SYNOPSIS
    Size + archiving (MRM) diagnostics for one or more mailboxes.
    READ ONLY: no write cmdlets against the tenant.

.DESCRIPTION
    For each mailbox it shows:
      - primary vs quota, recoverable items, archive vs quota, auto-expanding
      - retention policy, the default tag (DPT) and what blocks MRM (holds, ElcProcessing)
      - folders pointing to a tag that no longer exists (the "Unable to retrieve..." WARNING)
      - TOP N folders (own size, subfolders not included) with: oldest item, effective
        tag and age limit, MRM status, estimated inflow GB/day and how big the folder would be
        with a 30-day tag.

    Reading the status:
      MRM pending    -> tag applied, the assistant has not run yet (cycle up to 7 days)
      OVERDUE        -> has an item older than the age limit + 7 days: investigate
      OK on schedule -> archiving has already done its job; the volume is recent mail.
                        Only a shorter tag will reduce it.

    GB/day is an estimate: size / days between the oldest and the newest item.

.PARAMETER Mailbox
    One or more email addresses. Without the parameter, it prompts (accepts several separated by commas).

.PARAMETER Top
    How many folders to show per mailbox. Default 5.

.PARAMETER AdminUPN
    Admin account used to connect to Exchange Online if no session is active.

.EXAMPLE
    .\Get-MailboxFolderSize.ps1 -Mailbox user1@domain.com,user2@domain.com -AdminUPN ADMIN@DOMAIN.COM

.NOTES
    An Exchange admin role (e.g. "Exchange Recipient Administrator", activated in PIM if you
    use it) must be active BEFORE running, otherwise Get-Mailbox answers "couldn't be found"
    instead of access denied.
    CSV output goes to the 'csv' subfolder next to the script ($env:TEMP if not writable).
#>
[CmdletBinding()]
param(
    [string[]]$Mailbox,
    [int]$Top = 5,
    [Parameter(Mandatory = $true)][string]$AdminUPN
)

$ErrorActionPreference = 'Stop'

function Convert-ExchangeSizeToBytes {
    param([AllowNull()]$Size)
    if ($null -eq $Size) { return [int64]0 }
    try {
        if ($Size.PSObject.Methods.Name -contains 'ToBytes') { return [int64]$Size.ToBytes() }
        if ($null -ne $Size.Value -and $Size.Value.PSObject.Methods.Name -contains 'ToBytes') { return [int64]$Size.Value.ToBytes() }
    } catch { }
    # In EXO REST the size comes back as TEXT "9.88 GB (10,608,936,405 bytes)".
    # Sorting that text directly gives the wrong (lexical) order; always convert first.
    if ([string]$Size -match '\(([0-9,\.\s]+)\s+bytes\)') { return [int64]($Matches[1] -replace '[^0-9]', '') }
    return [int64]0
}

function Convert-ToDays {
    param([AllowNull()]$Age)
    if ($null -eq $Age -or [string]::IsNullOrWhiteSpace([string]$Age)) { return $null }
    try { return [int]([TimeSpan]::Parse([string]$Age).TotalDays) } catch { return $null }
}

function GB { param([int64]$b) [math]::Round($b / 1GB, 2) }

# ---------- output: proves it is writable, not just that it exists ----------
$__root = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.ProviderPath }
$csvDir = Join-Path $__root 'csv'   # CSV always in the csv subfolder next to the script, even if the script is moved
if ((-not (Test-Path $csvDir)) -and (Test-Path (Join-Path (Split-Path $__root) 'csv'))) { $csvDir = Join-Path (Split-Path $__root) 'csv' }   # script in a subfolder: use the csv folder of the parent
try { if (-not (Test-Path $csvDir)) { $null = New-Item -ItemType Directory -Path $csvDir -ErrorAction Stop } } catch {}
try { $__p = Join-Path $csvDir "w$PID.tmp"; [IO.File]::WriteAllText($__p, 'x'); [IO.File]::Delete($__p) }
catch { $csvDir = $env:TEMP; Write-Host "WARNING: CSV output redirected to $csvDir" -ForegroundColor Yellow }

# ---------- input ----------
if (-not $Mailbox) {
    $Mailbox = (Read-Host 'FULL email address(es), separated by commas') -split '[,;\s]+' | Where-Object { $_ }
}
$Mailbox = @($Mailbox | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^[^\s@]+@[^\s@]+\.[^\s@]+$' })
if (-not $Mailbox) { throw 'No valid email address provided.' }

# ---------- connection: reuses the open session ----------
$connectedHere = $false
$connected = $false
try { $connected = [bool](Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object State -eq 'Connected') } catch { }
if (-not $connected) {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Write-Host "Sign in with: $AdminUPN" -ForegroundColor Yellow
    Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false -ErrorAction Stop
    $connectedHere = $true
}

$tagCache = @{}
function Get-TagInfo {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    if (-not $tagCache.ContainsKey($Name)) {
        $t = Get-RetentionPolicyTag -Identity $Name -ErrorAction SilentlyContinue
        $tagCache[$Name] = if ($t) {
            [PSCustomObject]@{
                Name    = $t.Name
                Type    = [string]$t.Type
                Action  = [string]$t.RetentionAction
                Days    = Convert-ToDays $t.AgeLimitForRetention
                Enabled = [bool]$t.RetentionEnabled
            }
        } else { $null }
    }
    return $tagCache[$Name]
}

$csv = New-Object System.Collections.Generic.List[object]
$now = Get-Date

try {
    foreach ($mbx in $Mailbox) {
        Write-Host "`n==================== $mbx ====================" -ForegroundColor Cyan
        try {
            $m = Get-Mailbox -Identity $mbx -ErrorAction Stop
        } catch {
            Write-Host "Could not find $mbx. Is an Exchange admin role (e.g. via PIM) active? ($($_.Exception.Message))" -ForegroundColor Red
            continue
        }

        $st = Get-MailboxStatistics -Identity $mbx -ErrorAction Stop
        $used = Convert-ExchangeSizeToBytes $st.TotalItemSize
        $quota = Convert-ExchangeSizeToBytes $m.ProhibitSendReceiveQuota
        $recov = Convert-ExchangeSizeToBytes $st.TotalDeletedItemSize

        $archTxt = 'NO ARCHIVE'
        if ([string]$m.ArchiveStatus -eq 'Active') {
            $ast = Get-MailboxStatistics -Identity $mbx -Archive -ErrorAction SilentlyContinue
            $aUsed = Convert-ExchangeSizeToBytes $ast.TotalItemSize
            $aQuota = Convert-ExchangeSizeToBytes $m.ArchiveQuota
            $aPct = if ($aQuota) { [math]::Round(100 * $aUsed / $aQuota) } else { '?' }
            $auto = if ($m.AutoExpandingArchiveEnabled) { 'auto-expanding ON' } else { 'auto-expanding off' }
            $archTxt = "$(GB $aUsed) GB of $(GB $aQuota) GB ($aPct%) | $auto"
        }

        # the policy's default archive tag (DPT)
        $dpt = $null
        $personalTags = @()
        if ($m.RetentionPolicy) {
            $pol = Get-RetentionPolicy -Identity ([string]$m.RetentionPolicy) -ErrorAction SilentlyContinue
            foreach ($link in @($pol.RetentionPolicyTagLinks)) {
                $ti = Get-TagInfo ([string]$link)
                if (-not $ti) { continue }
                if ($ti.Type -eq 'All' -and $ti.Action -eq 'MoveToArchive') { $dpt = $ti }
                if ($ti.Type -eq 'Personal' -and $ti.Enabled -and $ti.Days) { $personalTags += "$($ti.Name)" }
            }
        }
        $dptTxt = if ($dpt) { "$($dpt.Name) ($($dpt.Days)d)" } else { 'NONE - folders without a personal tag never archive' }

        $blockers = @()
        if ($m.ElcProcessingDisabled) { $blockers += 'ElcProcessingDisabled' }
        if ($m.RetentionHoldEnabled)  { $blockers += 'RetentionHold' }
        if ($m.LitigationHoldEnabled) { $blockers += 'LitigationHold (does not prevent moving, retains whatever is deleted)' }
        $holds = @($m.InPlaceHolds | Where-Object { $_ })
        if ($holds.Count) { $blockers += "InPlaceHolds: $($holds.Count)" }
        $blockersTxt = if ($blockers) { $blockers -join ', ' } else { 'none' }

        # folders; missing-tag WARNINGs are counted, not dumped to the screen
        $warnings = $null
        $fs = Get-MailboxFolderStatistics -Identity $mbx -IncludeOldestAndNewestItems -WarningAction SilentlyContinue -WarningVariable warnings -ErrorAction Stop
        $orphans = @($warnings | Where-Object { "$_" -match 'policy tag' }).Count

        $pctTxt = if ($quota) { "$([math]::Round(100 * $used / $quota))%" } else { 'quota ?' }
        Write-Host ("Primary    : {0} GB of {1} GB ({2}) | {3} items | Recoverable {4} GB" -f (GB $used), (GB $quota), $pctTxt, $st.ItemCount, (GB $recov))
        Write-Host ("Archive    : {0}" -f $archTxt)
        Write-Host ("Policy     : {0} | default (DPT): {1}" -f $m.RetentionPolicy, $dptTxt)
        if ($personalTags) { Write-Host ("Folder tags: {0}" -f ($personalTags -join ', ')) -ForegroundColor DarkGray }
        Write-Host ("Blockers   : {0}" -f $blockersTxt) -ForegroundColor $(if ($blockers) { 'Yellow' } else { 'Gray' })
        if ($orphans) { Write-Host ("Orphan tag : {0} folder(s) point to a tag that no longer exists in the tenant" -f $orphans) -ForegroundColor Yellow }

        $rows = foreach ($f in $fs) {
            if ($f.RecoverableItemsFolder -or [string]$f.FolderType -eq 'Root') { continue }
            $b = Convert-ExchangeSizeToBytes $f.FolderSize
            if ($b -le 0) { continue }

            $tag = if ($f.ArchivePolicy) { Get-TagInfo ([string]$f.ArchivePolicy) } else { $dpt }
            $tagTxt = if ($f.ArchivePolicy -and -not $tag) { "$($f.ArchivePolicy) (ORPHAN)" }
                      elseif (-not $tag) { '-' }
                      elseif ($f.ArchivePolicy) { "$($tag.Name)" }
                      else { "$($tag.Name) [DPT]" }

            $oldestDate = $f.OldestItemReceivedDate
            $newestDate  = $f.NewestItemReceivedDate
            $itemAge = if ($oldestDate) { [int]($now - [datetime]$oldestDate).TotalDays } else { $null }
            $email = ([string]$f.ContainerClass -eq 'IPF.Note')

            $status = '-'
            if ([string]$f.FolderType -eq 'DeletedItems') { $status = 'Deleted Items: age counts from deletion' }
            elseif (-not $email) { $status = 'not a mail folder' }
            elseif (-not $tag -or -not $tag.Enabled -or -not $tag.Days) { $status = 'NO ARCHIVING' }
            elseif ([string]$f.RetentionFlags -match 'NeedsRescan') { $status = 'MRM pending (new tag)' }
            elseif ($null -ne $itemAge -and $itemAge -gt ($tag.Days + 7)) { $status = "OVERDUE: item ${itemAge}d old" }
            else { $status = "OK on schedule ($($tag.Days)d)" }

            $perDay = $null; $in30 = $null
            if ($email -and $oldestDate -and $newestDate) {
                $span = [math]::Max(1, ([datetime]$newestDate - [datetime]$oldestDate).TotalDays)
                $perDay = [math]::Round(($b / 1GB) / $span, 2)
                $in30 = [math]::Min((GB $b), [math]::Round($perDay * 30, 1))
            }

            [PSCustomObject]@{
                Mailbox      = $mbx
                Folder        = $f.FolderPath
                Items        = $f.ItemsInFolder
                GB           = GB $b
                PctMailbox     = if ($used) { [math]::Round(100 * $b / $used) } else { 0 }
                Oldest   = if ($oldestDate) { ([datetime]$oldestDate).ToString('yyyy-MM-dd') } else { '' }
                AgeDays    = $itemAge
                Tag          = $tagTxt
                Status       = $status
                GBPerDay     = $perDay
                GBIn30d     = $in30
                Bytes        = $b
            }
        }

        $topList = @($rows | Sort-Object Bytes -Descending | Select-Object -First $Top)
        $sumTop = ($topList | Measure-Object Bytes -Sum).Sum
        Write-Host ("`nTOP {0} FOLDERS ({1} GB = {2}% of primary)" -f $topList.Count, (GB $sumTop), $(if ($used) { [math]::Round(100 * $sumTop / $used) } else { 0 })) -ForegroundColor Yellow
        $topList | Select-Object Folder, Items, GB, @{n='%';e={$_.PctMailbox}}, Oldest, Tag, Status, @{n='GB/day';e={$_.GBPerDay}}, @{n='GB w/30d';e={$_.GBIn30d}} |
            Format-Table -AutoSize | Out-String -Width 300 | Write-Host

        foreach ($t in $topList) { $csv.Add(($t | Select-Object * -ExcludeProperty Bytes)) }
    }

    if ($csv.Count) {
        $outFile = Join-Path $csvDir ("MailboxTopFolders_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
        $csv | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8
        Write-Host "CSV: $outFile" -ForegroundColor Green
    }
}
finally {
    if ($connectedHere) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }
}
