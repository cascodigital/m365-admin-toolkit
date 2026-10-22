<#
.SYNOPSIS
    Removes corrupted inbox rules from a mailbox, by NAME.

.DESCRIPTION
    Corrupted rules are NOT returned by Get-InboxRule (they fail to parse and are
    skipped, emitting only a WARNING). That is why InError is always False and they
    can only be removed by NAME.

    Success metric: the count of WARNINGs from Get-InboxRule must drop to 0.
    Get the rule names from those WARNINGs ("... contains errors") -- or run
    Audit-InboxRules.ps1 first, which lists them.
    The script measures this before and after, so you have proof, not guesswork.

.PARAMETER DryRun
    Only shows what it would do. Deletes NOTHING. Run it this way first.

.PARAMETER NukeOutlookBlob
    LAST RESORT. Adds -AlwaysDeleteOutlookRulesBlob, which deletes Outlook's client-side
    rules blob. This DESTROYS client-only rules (play a sound, desktop alert,
    move to local PST) that are NOT in any backup because they never reach
    the server. Only use with the mailbox owner's explicit authorization.

.PARAMETER MailboxAddress
    The mailbox that holds the corrupted rules.

.PARAMETER CorruptRules
    Names of the corrupted rules, copied from the Get-InboxRule "contains errors" WARNINGs.

.NOTES
    Requires an active Exchange Online session (Connect-ExchangeOnline) with an
    Exchange admin role. Backups (XML + CSV) and the result CSV go to the 'csv'
    subfolder next to the script ($env:TEMP if not writable).

.EXAMPLE
    .\Remove-CorruptInboxRules.ps1 -MailboxAddress user@domain.com -CorruptRules 'Rule A','Rule B' -DryRun
    .\Remove-CorruptInboxRules.ps1 -MailboxAddress user@domain.com -CorruptRules 'Rule A','Rule B'
#>

param(
    [Parameter(Mandatory = $true)][string]$MailboxAddress,
    # Names copied from the Get-InboxRule "contains errors" WARNINGs
    [Parameter(Mandatory = $true)][string[]]$CorruptRules,
    [switch]$DryRun,
    [switch]$NukeOutlookBlob
)

$ErrorActionPreference = 'Stop'

$stamp   = Get-Date -Format 'yyyyMMdd_HHmm'
$__root = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.ProviderPath }
$csvDir = Join-Path $__root 'csv'   # CSV always in the csv subfolder next to the script, even if the script is moved
if ((-not (Test-Path $csvDir)) -and (Test-Path (Join-Path (Split-Path $__root) 'csv'))) { $csvDir = Join-Path (Split-Path $__root) 'csv' }   # script in a subfolder: use the csv folder of the parent
try { if (-not (Test-Path $csvDir)) { $null = New-Item -ItemType Directory -Path $csvDir -ErrorAction Stop } } catch {}
try { $__p = Join-Path $csvDir "w$PID.tmp"; [IO.File]::WriteAllText($__p,'x'); [IO.File]::Delete($__p) }
catch { $csvDir = $env:TEMP; Write-Host "WARNING: CSV output redirected to $csvDir" -Foreground Yellow }

# ---------------------------------------------------------------------------
# 0. Live session?
# ---------------------------------------------------------------------------
$conn = Get-ConnectionInformation -ErrorAction SilentlyContinue |
        Where-Object { $_.State -eq 'Connected' -and $_.TokenStatus -eq 'Active' }
if (-not $conn) {
    Write-Error "No active EXO session. Connect first (and confirm your Exchange admin role / PIM activation is active)."
    exit 1
}

# ---------------------------------------------------------------------------
# 1. BASELINE - count the warnings BEFORE (this is the metric that matters)
# ---------------------------------------------------------------------------
Write-Host "`n=== BASELINE ===" -ForegroundColor Cyan

$before = Get-InboxRule -Mailbox $MailboxAddress -IncludeHidden `
                        -WarningVariable warnBefore -WarningAction SilentlyContinue

Write-Host ("Readable rules  : {0}" -f $before.Count)
Write-Host ("Rules in error  : {0}   <-- server warnings" -f $warnBefore.Count) -ForegroundColor Yellow
$warnBefore | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkYellow }

# ---------------------------------------------------------------------------
# 2. BACKUP of the readable ones (the corrupted ones are unsalvageable by definition)
# ---------------------------------------------------------------------------
$xml = Join-Path $csvDir "Rules_PreDelete_$stamp.xml"
$csv = Join-Path $csvDir "Rules_PreDelete_$stamp.csv"

$before | Export-Clixml -Path $xml
$before | Select-Object Name, Enabled, Priority, InError, ErrorType, RuleIdentity, Description |
    Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8

Write-Host "`nBackup saved:" -ForegroundColor Green
Write-Host "  $xml"
Write-Host "  $csv"

if ($before.Count -eq 0) {
    Write-Error "Empty backup. Aborting for safety."
    exit 1
}

# ---------------------------------------------------------------------------
# 3. Removal, one by one, with an individual report
# ---------------------------------------------------------------------------
Write-Host "`n=== REMOVAL ===" -ForegroundColor Cyan
if ($DryRun)          { Write-Host "DRY-RUN MODE: nothing will be deleted." -ForegroundColor Magenta }
if ($NukeOutlookBlob) { Write-Warning "-AlwaysDeleteOutlookRulesBlob ACTIVE: client-only rules will be destroyed." }

$results = foreach ($name in $CorruptRules) {

    if ($DryRun) {
        Write-Host "  [DRY] would remove: $name" -ForegroundColor Magenta
        [PSCustomObject]@{ Rule = $name; Status = 'DryRun'; Detail = '' }
        continue
    }

    $params = @{
        Mailbox     = $MailboxAddress
        Identity    = $name
        Confirm     = $false
        ErrorAction = 'Stop'
    }
    if ($NukeOutlookBlob) { $params['AlwaysDeleteOutlookRulesBlob'] = $true }

    try {
        Remove-InboxRule @params
        Write-Host "  [OK]   removed: $name" -ForegroundColor Green
        [PSCustomObject]@{ Rule = $name; Status = 'Removed'; Detail = '' }
    }
    catch {
        $msg = $_.Exception.Message
        Write-Host "  [FAIL] $name -> $msg" -ForegroundColor Red
        [PSCustomObject]@{ Rule = $name; Status = 'Failed'; Detail = $msg }
    }
}

$results | Export-Csv -Path (Join-Path $csvDir "Rules_RemovalResult_$stamp.csv") `
                      -NoTypeInformation -Encoding UTF8

if ($DryRun) {
    Write-Host "`nDry-run complete. Run without -DryRun to apply." -ForegroundColor Magenta
    exit 0
}

# ---------------------------------------------------------------------------
# 4. VERIFICATION - did the warnings drop?
# ---------------------------------------------------------------------------
Write-Host "`n=== VERIFICATION ===" -ForegroundColor Cyan
Start-Sleep -Seconds 5

$after = Get-InboxRule -Mailbox $MailboxAddress -IncludeHidden `
                       -WarningVariable warnAfter -WarningAction SilentlyContinue

Write-Host ("Rules in error BEFORE: {0}" -f $warnBefore.Count)
Write-Host ("Rules in error AFTER : {0}" -f $warnAfter.Count)
Write-Host ("Readable rules       : {0} -> {1}" -f $before.Count, $after.Count)

if ($warnAfter.Count -eq 0) {
    Write-Host "`nSUCCESS: no corrupted rules left on the server." -ForegroundColor Green
    Write-Host "Next step: the user needs to CLOSE Outlook completely" -ForegroundColor Cyan
    Write-Host "(confirm OUTLOOK.EXE is gone in Task Manager) and reopen it." -ForegroundColor Cyan
}
else {
    Write-Warning "$($warnAfter.Count) rule(s) still in error:"
    $warnAfter | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkYellow }
    Write-Host "If the same ones remain, Remove-InboxRule cannot touch them." -ForegroundColor Yellow
    Write-Host "Then the way forward is -NukeOutlookBlob, with the mailbox owner's authorization." -ForegroundColor Yellow
}
