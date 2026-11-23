<#
.SYNOPSIS
    Removes corrupted inbox rules from a mailbox, by NAME.

.DESCRIPTION
    Corrupted rules are NOT returned by Get-InboxRule (they fail to parse and are
    skipped, emitting only a WARNING). That is why InError is always False and they
    can only be removed by NAME.

    1. Reads the mailbox (-IncludeHidden), captures every warning and extracts the
       rule name from it (or uses -CorruptRules if you pass the names yourself).
    2. Backs up the readable rules (XML + CSV). Corrupted ones cannot be backed up.
    3. Shows the plan and only removes if you type YES.
    4. Re-reads and compares the warning count BEFORE and AFTER. That count is the proof.

    Readable rules are never touched. A corrupted name that also exists as a readable
    rule is held as MANUAL. A warning with no rule name in it is MANUAL too: rerun
    with -CorruptRules <name>.

.PARAMETER MailboxAddress
    The mailbox that holds the corrupted rules.

.PARAMETER CorruptRules
    Optional. Names of the corrupted rules. If omitted, they are taken from the
    Get-InboxRule "contains errors" warnings.

.PARAMETER DryRun
    Reads and shows the plan. Deletes NOTHING. Alias: -WhatIf.

.PARAMETER NukeOutlookBlob
    LAST RESORT. Adds -AlwaysDeleteOutlookRulesBlob, which deletes Outlook's client-side
    rules blob. This DESTROYS client-only rules (play a sound, desktop alert,
    move to local PST) that are NOT in any backup because they never reach
    the server. Only use with the mailbox owner's explicit authorization.

.NOTES
    Requires an active Exchange Online session (Connect-ExchangeOnline) with an
    Exchange admin role. Backups and the result CSV go to the 'csv' subfolder next
    to the script ($env:TEMP if not writable).
    Rollback: none for the corrupted rules (unreadable). Readable rules are in
    csv\<stamp>.<mailbox>.Rules-BEFORE.xml

.EXAMPLE
    .\Remove-CorruptInboxRules.ps1 -MailboxAddress user@domain.com -DryRun
    .\Remove-CorruptInboxRules.ps1 -MailboxAddress user@domain.com
    .\Remove-CorruptInboxRules.ps1 -MailboxAddress user@domain.com -CorruptRules 'Rule A','Rule B'
#>

param(
    [Parameter(Mandatory = $true)][string]$MailboxAddress,
    [Alias('Rule')][string[]]$CorruptRules,
    [Alias('WhatIf')][switch]$DryRun,
    [switch]$NukeOutlookBlob
)

[Console]::OutputEncoding = [Text.Encoding]::UTF8
$OutputEncoding = [Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'

# ---------- output folder (write probe) ----------
$__root = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.ProviderPath }
$csvDir = Join-Path $__root 'csv'
$__parent = Split-Path $__root
if ($__parent -and (-not (Test-Path $csvDir)) -and (Test-Path (Join-Path $__parent 'csv'))) { $csvDir = Join-Path $__parent 'csv' }
try { if (-not (Test-Path $csvDir)) { $null = New-Item -ItemType Directory -Path $csvDir -ErrorAction Stop } } catch {}
try { $__p = Join-Path $csvDir "w$PID.tmp"; [IO.File]::WriteAllText($__p,'x'); [IO.File]::Delete($__p) }
catch { $csvDir = if ($env:TEMP) { $env:TEMP } else { [IO.Path]::GetTempPath() }; Write-Host "WARNING: CSV output redirected to $csvDir" -Foreground Yellow }
$stamp = Get-Date -Format 'yyyy.MM.dd.HHmmss'
$slug  = ($MailboxAddress.Split('@')[0]) -replace '[^\w.\-]', '_'

# ---------- EXO session ----------
$conn = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' -and $_.TokenStatus -eq 'Active' })
if (-not $conn.Count) {
    Write-Host 'No active EXO session. Run Connect-ExchangeOnline first (and confirm your Exchange admin role / PIM activation is active).' -Foreground Red
    exit 1
}

function Read-Rules {
    $w = $null
    $r = @(Get-InboxRule -Mailbox $MailboxAddress -IncludeHidden -WarningVariable w -WarningAction SilentlyContinue)
    [pscustomobject]@{ Rules = $r; Warnings = @($w | ForEach-Object { [string]$_ } | Where-Object { $_.Trim() } | Select-Object -Unique) }
}

# ---------- 1. READ ----------
$before = Read-Rules
$readable = @($before.Rules | ForEach-Object { [string]$_.Name })
Write-Host ''
Write-Host "Mailbox: $MailboxAddress" -Foreground Cyan
Write-Host ("Readable rules : {0}" -f $before.Rules.Count)
Write-Host ("Read warnings  : {0}" -f $before.Warnings.Count) -Foreground $(if ($before.Warnings.Count) { 'Yellow' } else { 'Green' })
$before.Warnings | ForEach-Object { Write-Host "    $_" -Foreground DarkYellow }

# rule names: from -CorruptRules if given, otherwise the first "quoted" text of each warning
$plan = New-Object System.Collections.Generic.List[object]
if ($CorruptRules) {
    foreach ($n in $CorruptRules) { $plan.Add([pscustomobject]@{ Rule = $n.Trim(); Source = '-CorruptRules'; Warning = ''; Status = 'REMOVE'; Detail = '' }) }
} else {
    foreach ($w in $before.Warnings) {
        if ($w -notmatch '(?i)rule') {
            $plan.Add([pscustomobject]@{ Rule = ''; Source = 'warning'; Warning = $w; Status = 'MANUAL'; Detail = 'warning is not about an inbox rule - read it' })
        } elseif ($w -match '"([^"]+)"') {
            $plan.Add([pscustomobject]@{ Rule = $Matches[1]; Source = 'warning'; Warning = $w; Status = 'REMOVE'; Detail = '' })
        } else {
            $plan.Add([pscustomobject]@{ Rule = ''; Source = 'warning'; Warning = $w; Status = 'MANUAL'; Detail = 'no rule name in the warning - rerun with -CorruptRules <name>' })
        }
    }
}
foreach ($p in $plan) {
    if ($p.Status -eq 'REMOVE' -and ($readable -contains $p.Rule)) { $p.Status = 'MANUAL'; $p.Detail = 'same name as a READABLE rule - not removed by name' }
}
$targets = @($plan | Where-Object Status -eq 'REMOVE')

if (-not $plan.Count) {
    Write-Host ''
    Write-Host 'NO CORRUPTED RULES (0 read warnings). NOTHING WAS CHANGED.' -Foreground Green
    return
}

# ---------- 2. BACKUP of the readable rules ----------
# 0 readable rules usually means wrong mailbox or no permission -> stop (unless the names were given)
if ($before.Rules.Count -eq 0 -and -not $CorruptRules) {
    Write-Host 'STOP: 0 readable rules on this mailbox (wrong address, or no permission?). If every rule is corrupted, run again with -CorruptRules <names>. NOTHING WAS CHANGED.' -Foreground Red
    return
}
$xml = Join-Path $csvDir "$stamp.$slug.Rules-BEFORE.xml"
$csv = Join-Path $csvDir "$stamp.$slug.Rules-BEFORE.csv"
try {
    $before.Rules | Export-Clixml -Path $xml
    $before.Rules | Select-Object Name, Enabled, Priority, InError, ErrorType, RuleIdentity, Description |
        Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
} catch {
    Write-Host "STOP: backup could not be written ($($_.Exception.Message)). NOTHING WAS CHANGED." -Foreground Red
    return
}

# ---------- 3. PLAN ----------
Write-Host ''
Write-Host "PLAN - $MailboxAddress" -Foreground Cyan
$plan | Format-Table Status, Rule, Source, Detail -AutoSize -Wrap
if ($NukeOutlookBlob) { Write-Host '-NukeOutlookBlob ACTIVE: client-only Outlook rules will be DESTROYED (no backup).' -Foreground Red }
Write-Host "Backup of readable rules: $xml" -Foreground DarkGray
Write-Host "TO REMOVE: $($targets.Count)  |  MANUAL: $(@($plan | Where-Object Status -eq 'MANUAL').Count)" -Foreground Cyan

if (-not $targets.Count) { Write-Host 'NOTHING TO REMOVE AUTOMATICALLY. NOTHING WAS CHANGED.' -Foreground Yellow; return }
if ($DryRun) { Write-Host 'DRY-RUN - NOTHING WAS CHANGED.' -Foreground Yellow; return }
$resp = (Read-Host "Type YES to remove $($targets.Count) rule(s) from $MailboxAddress").Trim()
if ($resp -ne 'YES') { Write-Host 'CANCELLED. NOTHING WAS CHANGED.' -Foreground Yellow; return }

# ---------- 4. APPLY ----------
foreach ($t in $targets) {
    $prm = @{ Mailbox = $MailboxAddress; Identity = $t.Rule; Confirm = $false; ErrorAction = 'Stop' }
    if ($NukeOutlookBlob) { $prm['AlwaysDeleteOutlookRulesBlob'] = $true }
    try   { Remove-InboxRule @prm; $t.Status = 'REMOVED'; Write-Host "  ok: $($t.Rule)" -Foreground Green }
    catch { $t.Status = 'FAILED'; $t.Detail = $_.Exception.Message; Write-Host "  FAILED: $($t.Rule) - $($t.Detail)" -Foreground Red }
}

# ---------- 5. VERIFY ----------
Start-Sleep -Seconds 5
$after = Read-Rules
$res = Join-Path $csvDir "$stamp.$slug.RemoveCorruptRules.csv"
$plan | Select-Object Status, Rule, Source, Detail, Warning | Export-Csv -Path $res -NoTypeInformation -Encoding UTF8

Write-Host ''
Write-Host ("Read warnings BEFORE : {0}" -f $before.Warnings.Count)
Write-Host ("Read warnings AFTER  : {0}" -f $after.Warnings.Count) -Foreground $(if ($after.Warnings.Count) { 'Yellow' } else { 'Green' })
Write-Host ("Readable rules       : {0} -> {1}" -f $before.Rules.Count, $after.Rules.Count)
if ($after.Warnings.Count -eq 0) {
    Write-Host 'VERIFIED: no corrupted rules left on the server. User must fully close Outlook (OUTLOOK.EXE gone) and reopen.' -Foreground Green
} else {
    Write-Host "PARTIAL: $($after.Warnings.Count) warning(s) left:" -Foreground Yellow
    $after.Warnings | ForEach-Object { Write-Host "    $_" -Foreground DarkYellow }
    Write-Host 'If they are the same ones, Remove-InboxRule cannot touch them: -NukeOutlookBlob, with the mailbox owner''s OK.' -Foreground Yellow
}
Write-Host "Result: $res" -Foreground DarkGray
