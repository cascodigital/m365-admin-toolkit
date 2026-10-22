<#
.SYNOPSIS
    Makes a shared mailbox keep a copy of mail sent as / on behalf of it in its own Sent Items.

.DESCRIPTION
    Turns on MessageCopyForSentAsEnabled and MessageCopyForSendOnBehalfEnabled. Without them,
    a message sent "as" the shared mailbox only shows up in the sender's own Sent Items.
    Refuses to run if the target is not a SharedMailbox. Writes a BEFORE/AFTER verification CSV.

    SIMULATES BY DEFAULT. It only applies the change when the global APPLY key is set
    in the same command line (see examples). The key is consumed on every run.

.PARAMETER Mailbox
    Address of the shared mailbox. If omitted, the $DefaultTarget placeholder at the top
    of the script is used (edit it, or pass -Mailbox).

.EXAMPLE
    .\Enable-SharedSentItemsCopy.ps1 -Mailbox shared@domain.com
    Simulates.

.EXAMPLE
    $global:APPLY = $true; .\Enable-SharedSentItemsCopy.ps1 -Mailbox shared@domain.com
    Applies.

.NOTES
    IMPACT  : turns on the Sent Items copy for Send As and Send on Behalf. 1 mailbox. Reversible.
    ROLLBACK: Set-Mailbox <shared> -MessageCopyForSentAsEnabled $false -MessageCopyForSendOnBehalfEnabled $false
    Requires an already connected Exchange Online session (Connect-ExchangeOnline).
    The legacy key $global:EXECUTAR is still accepted as an alias of $global:APPLY.
    CSV output goes to the 'csv' subfolder next to the script ($env:TEMP if not writable).
#>
param([string]$Mailbox)

# ---------- TARGET (placeholder) ----------
$DefaultTarget = 'SHARED@DOMAIN.COM'   # <<< edit, or pass -Mailbox

# ---------- SAFETY PREAMBLE ----------
# This script SIMULATES by default. To APPLY, run everything in a single command:
#     $global:APPLY = $true; .\Enable-SharedSentItemsCopy.ps1 -Mailbox shared@domain.com
$Execute = $false                 # <<< FAIL CLOSED: NEVER remove this line
$__k  = Get-Variable APPLY    -Scope Global -ValueOnly -ErrorAction SilentlyContinue
$__k2 = Get-Variable EXECUTAR -Scope Global -ValueOnly -ErrorAction SilentlyContinue   # legacy key name
$Execute = (($__k -is [bool] -and $__k) -or ($__k2 -is [bool] -and $__k2))
$global:APPLY = $false            # consumes the key
$global:EXECUTAR = $false

$__root = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.ProviderPath }
$csvDir = Join-Path $__root 'csv'   # CSV always in the csv subfolder next to the script, even if the script is moved
if ((-not (Test-Path $csvDir)) -and (Test-Path (Join-Path (Split-Path $__root) 'csv'))) { $csvDir = Join-Path (Split-Path $__root) 'csv' }   # script in a subfolder: use the csv folder of the parent
try { if (-not (Test-Path $csvDir)) { $null = New-Item -ItemType Directory -Path $csvDir -ErrorAction Stop } } catch {}
try { $__p = Join-Path $csvDir "w$PID.tmp"; [IO.File]::WriteAllText($__p,'x'); [IO.File]::Delete($__p) }
catch { $csvDir = $env:TEMP; Write-Host "WARNING: CSV output redirected to $csvDir" -Foreground Yellow }

$target = if ($Mailbox) { $Mailbox } else { $DefaultTarget }
if ($target -eq 'SHARED@DOMAIN.COM') { throw "STOP: set -Mailbox (or edit `$DefaultTarget) - the placeholder is still in place." }
$fields = 'Name','PrimarySmtpAddress','RecipientTypeDetails','MessageCopyForSentAsEnabled','MessageCopyForSendOnBehalfEnabled'

$before = Get-Mailbox -Identity $target -ErrorAction Stop | Select-Object $fields
Write-Host "TARGETS: $(@($before).Count)" -Foreground Cyan
if ($before.RecipientTypeDetails -ne 'SharedMailbox') { throw "STOP: $target is not a SharedMailbox ($($before.RecipientTypeDetails))" }

$p = @{ Identity = $target; MessageCopyForSentAsEnabled = $true; MessageCopyForSendOnBehalfEnabled = $true; ErrorAction = 'Stop' }
if ($Execute) {
    try { Set-Mailbox @p } catch { Write-Host "FAILED: $($_.Exception.Message)" -Foreground Red }
} else {
    Write-Host "SIMULATION: Set-Mailbox $(($p.GetEnumerator() | Sort-Object Key | ForEach-Object { "-$($_.Key) '$($_.Value)'" }) -join ' ')"
}

# ============ VERIFICATION - LAST BLOCK, NOTHING AFTER ============
if (-not $before) { Write-Host "ATTENTION: `$before EMPTY - this table is NOT a verification" -Foreground Red }
$after = Get-Mailbox -Identity $target | Select-Object $fields
$tab = @(
  $before | Select-Object @{n='Phase';e={'BEFORE'}},*
  $after  | Select-Object @{n='Phase';e={'AFTER' }},*
)
Write-Host ""
if (-not $Execute) {
    Write-Host '### SIMULATION MODE - NOTHING WAS CHANGED ###' -Foreground Yellow
    Write-Host '### Run  $global:APPLY = $true; .\Enable-SharedSentItemsCopy.ps1 -Mailbox <shared>  to apply ###' -Foreground Yellow
} else {
    Write-Host '### APPLIED - check BEFORE vs AFTER ###' -Foreground Green
}
$tab | Format-Table -AutoSize
$csv = Join-Path $csvDir "verification-shared-sentcopy-$(Get-Date -f yyyyMMdd-HHmmss).csv"
$tab | Export-Csv $csv -NoTypeInformation -Encoding UTF8
if (Test-Path $csv) { Write-Host "Summary saved to: $csv" }
else { Write-Host "ATTENTION: the CSV was NOT written" -Foreground Red }
