<#
.SYNOPSIS
    Investigates "my email went to quarantine" complaints in Exchange Online.
    Interactive by default, parameterizable for automation.

.DESCRIPTION
    v3.0 - rewritten after a real quarantine case, which exposed
    the central trap of this investigation:

        QUARANTINE indexes the P2  (header From:, what the user sees)
        MESSAGE TRACE indexes the P1 (envelope MAIL FROM / Return-Path)

    When the sender is a shared mailbox, alias or automation, P1 != P2 and
    searching the trace by the reported address returns ZERO results - a false
    negative that looks like a conclusion. This version NEVER correlates by sender:
    it uses MessageId, which is immune to the divergence.

    Second trap covered: Status=Delivered in the trace does NOT mean clean
    delivery. A message that was quarantined and later released shows as Delivered.
    The proof of the hold is in the quarantine ReleaseStatus.

.PARAMETER Recipient
    Mailbox that complained. Optional if -Sender is provided.

.PARAMETER Sender
    Sender. Accepts 'user@domain.com', '@domain.com' or 'domain.com'.
    Even with a full address, the trace filter works by DOMAIN (anti-P1/P2).

.PARAMETER Days
    Look-back window in days. Default 7 (prompted in interactive mode).

.PARAMETER Detail
    Also pull per-message event detail (Get-MessageTraceDetailV2) for up to 25 messages.

.PARAMETER NonInteractive
    Do not prompt; use only the parameters given.

.PARAMETER UPN
    Admin account used to connect to Exchange Online if no session is active.

.PARAMETER OutRoot
    Root folder for the evidence package. Default: 'MailFlowForensics' subfolder next to the
    script. CSVs go to the 'csv' subfolder next to the script unless -OutRoot is given.

.OUTPUTS
    Evidence folder with SUMMARY.txt, ticket.txt (internal note + end-user reply draft),
    console.log and CSVs (correlated, quarantine, messagetrace, orphans, ...).

.EXAMPLE
    .\Invoke-MailFlowForensics.ps1 -UPN ADMIN@DOMAIN.COM

.EXAMPLE
    .\Invoke-MailFlowForensics.ps1 -UPN ADMIN@DOMAIN.COM -Recipient user@domain.com -Sender @external.com -Days 7

.NOTES
    Environment : Exchange Online
    Requires    : ExchangeOnlineManagement 3.x + an Exchange / Defender role with access to
                  quarantine and message trace (activated in PIM, if you use it)
    Message trace uses Get-MessageTraceV2; the legacy Get-MessageTrace is only used as a
    fallback when V2 is not available in the session.
    Version     : 3.0
#>

[CmdletBinding()]
param(
    [string] $Recipient,
    [string] $Sender,
    [int]    $Days,
    [switch] $Detail,
    [switch] $NonInteractive,
    [Parameter(Mandatory = $true)][string] $UPN,
    [string] $OutRoot                       # default = subfolder next to the script
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

# Output next to the script, not on the Desktop.
$__outRootGiven = [bool]$OutRoot
if (-not $OutRoot) {
    $base = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $OutRoot = Join-Path $base 'MailFlowForensics'
}

#region ---------------------------------------------------------- UI HELPERS
function Say  { param($m, $c = 'Gray') Write-Host $m -ForegroundColor $c }
function Head { param($m) Say ''; Say ('=' * 76) DarkCyan; Say "  $m" Cyan; Say ('=' * 76) DarkCyan }
function Ask {
    param([string]$Question, [string]$Default, [switch]$AllowEmpty)
    while ($true) {
        $p = if ($Default) { "$Question [$Default]" } else { $Question }
        $a = (Read-Host $p).Trim()
        if (-not $a -and $Default)    { return $Default }
        if (-not $a -and $AllowEmpty) { return '' }
        if ($a) { return $a }
        Say '   Empty answer.' DarkYellow
    }
}
function Norm-MsgId { param($v) if ($v) { $v.Trim().Trim('<','>').ToLower() } else { '' } }
#endregion

#region ------------------------------------------------- INPUT NORMALIZATION
function Resolve-SenderInput {
    param([string]$Raw)
    if (-not $Raw) { return $null }
    $v = $Raw.Trim().TrimStart('*').Trim()
    if ($v -match '^[^@\s]+@([^@\s]+\.[^@\s]+)$') {
        return @{ Mode='Address'; Address=$v.ToLower(); Domain=$Matches[1].ToLower(); Pattern="*@$($Matches[1].ToLower())" }
    }
    $d = $v.TrimStart('@').ToLower()
    if ($d -match '^[^@\s]+\.[^@\s]+$') {
        return @{ Mode='Domain'; Address=$null; Domain=$d; Pattern="*@$d" }
    }
    return $null
}
function Test-EmailAddress { param([string]$v) $v -match '^[^@\s]+@[^@\s]+\.[^@\s]+$' }
#endregion

Clear-Host
Head 'MAIL FLOW FORENSICS v3.0  -  quarantine / inbound filtering'
Say '  Exchange Online' DarkGray
Say '  Correlation by MessageId (immune to P1/P2 divergence)' DarkGray

#region ------------------------------------------------------ INPUT COLLECTION
if (-not $NonInteractive) {
    Say ''
    Say ' Leave BLANK whatever you do not know. At least one of the two is required.' DarkYellow
    Say ''
    if (-not $Recipient) { $Recipient = Ask 'RECIPIENT email (who complained)' -AllowEmpty }
    if (-not $Sender) {
        Say '   (accepts "user@domain.com" or just "@domain.com")' DarkGray
        $Sender = Ask 'SENDER email/domain' -AllowEmpty
    }
    if (-not $Days)   { $Days = [int](Ask 'Window in days' -Default '7') }
    if (-not $Detail) { $Detail = (Ask 'Pull per-message event detail? (y/N)' -Default 'N') -match '^[sSyY]' }
}
if (-not $Days) { $Days = 7 }

if ($Recipient) { $Recipient = $Recipient.Trim() }
$Snd = Resolve-SenderInput $Sender

if (-not $Recipient -and -not $Snd) { throw 'Provide at least the recipient OR the sender.' }
if ($Recipient -and -not (Test-EmailAddress $Recipient)) { throw "Invalid recipient: '$Recipient'" }
if ($Sender -and -not $Snd) { throw "Invalid sender: '$Sender'. Use 'user@domain.com' or '@domain.com'." }

if ($Snd -and $Snd.Mode -eq 'Address') {
    Say ''
    Say " [i] Full address provided: $($Snd.Address)" Cyan
    Say "     The trace filter will work by DOMAIN (*@$($Snd.Domain)), not by" DarkGray
    Say '     address. Reason: trace indexes the envelope (P1), which often differs from' DarkGray
    Say '     the header From: (P2). Filtering by address produces a false negative.' DarkGray
}

$End   = (Get-Date).ToUniversalTime()
$Start = $End.AddDays(-$Days)
$Stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$Tag   = if ($Snd) { $Snd.Domain } elseif ($Recipient) { ($Recipient -split '@')[0] } else { 'query' }
$OutDir = Join-Path $OutRoot "$Tag`_$Stamp"
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
# CSV always in the csv subfolder next to the script, even if the script is moved
# (an explicit -OutRoot still wins). File names get the case prefix so runs do not overwrite each other.
$__root    = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.ProviderPath }
$csvDir    = if ($__outRootGiven) { $OutDir } else { Join-Path $__root 'csv' }
if (-not $__outRootGiven -and (-not (Test-Path $csvDir)) -and (Test-Path (Join-Path (Split-Path $__root) 'csv'))) { $csvDir = Join-Path (Split-Path $__root) 'csv' }   # script in a subfolder: use the csv folder of the parent
$csvPrefix = if ($__outRootGiven) { '' } else { "$Tag`_$Stamp`_" }
try { if (-not (Test-Path $csvDir)) { $null = New-Item -ItemType Directory -Path $csvDir -ErrorAction Stop } } catch {}
try { $__p = Join-Path $csvDir "w$PID.tmp"; [IO.File]::WriteAllText($__p,'x'); [IO.File]::Delete($__p) }
catch { $csvDir = $env:TEMP; Write-Host "WARNING: CSV output redirected to $csvDir" -Foreground Yellow }
try { Start-Transcript -Path (Join-Path $OutDir 'console.log') -Force | Out-Null } catch {}

Head 'PARAMETERS'
Say ("  Recipient    : {0}" -f $(if ($Recipient) { $Recipient } else { '(not specified)' })) White
Say ("  Sender       : {0}   [mode {1}]" -f $(if ($Snd) { $Snd.Pattern } else { '(not specified)' }), $(if ($Snd) { $Snd.Mode } else { 'n/a' })) White
Say ("  UTC window   : {0}  ->  {1}   ({2}d)" -f $Start.ToString('u'), $End.ToString('u'), $Days) White
Say ("  Output       : {0}" -f $OutDir) White
Say ("  CSV          : {0}" -f $csvDir) White
#endregion

#region ---------------------------------------------------------- CONNECTION
Head 'CONNECTION'
if (-not (Get-Module ExchangeOnlineManagement -ListAvailable)) {
    throw 'ExchangeOnlineManagement module missing. Install-Module ExchangeOnlineManagement -Scope CurrentUser'
}
Import-Module ExchangeOnlineManagement -ErrorAction Stop
$conn = Get-ConnectionInformation -ErrorAction SilentlyContinue |
        Where-Object { $_.State -eq 'Connected' -and $_.TokenStatus -eq 'Active' }
if ($conn) { Say "[+] Active session: $($conn.UserPrincipalName)" Green }
else       { Say "[*] Connecting as $UPN ..." Yellow
             Connect-ExchangeOnline -UserPrincipalName $UPN -ShowBanner:$false
             Say '[+] Connected.' Green }
#endregion

#region -------------------------------------------------------- PRE-FLIGHT
Head 'PERMISSION PRE-FLIGHT'
$Can = [ordered]@{ Quarantine=$false; TraceV2=$false; TraceV1=$false; Detail=$false
                   JunkCfg=$false; InboxRule=$false; AllowBlock=$false }
$Denied = @()
function Probe {
    param([string]$Name, [scriptblock]$Test)
    try { & $Test | Out-Null; $script:Can[$Name] = $true; Say "  [OK]     $Name" Green }
    catch { $script:Denied += "$Name -> $($_.Exception.Message)"; Say "  [DENIED] $Name" Red }
}
Probe 'Quarantine' { Get-QuarantineMessage -PageSize 1 -ErrorAction Stop }
Probe 'TraceV2'    { Get-MessageTraceV2 -StartDate $Start -EndDate $End -ResultSize 1 -ErrorAction Stop }
if (-not $Can.TraceV2) { Probe 'TraceV1' { Get-MessageTrace -StartDate $Start -EndDate $End -PageSize 1 -ErrorAction Stop } }
$Can.Detail = [bool]($Can.TraceV2 -and (Get-Command Get-MessageTraceDetailV2 -ErrorAction SilentlyContinue))
if ($Recipient) {
    Probe 'JunkCfg'   { Get-MailboxJunkEmailConfiguration -Identity $Recipient -ErrorAction Stop }
    Probe 'InboxRule' { Get-InboxRule -Mailbox $Recipient -ErrorAction Stop }
}
Probe 'AllowBlock' { Get-TenantAllowBlockListItems -ListType Sender -ErrorAction Stop }

if (-not $Can.Quarantine -and -not $Can.TraceV2 -and -not $Can.TraceV1) {
    Say ' No access to any primary source. Check/re-activate your admin role; if it persists, request' Red
    Say ' "Security Reader" or "Hygiene Management" in Defender.' Yellow
    $Denied | ForEach-Object { Say "   $_" DarkGray }
    try { Stop-Transcript | Out-Null } catch {}
    throw 'Aborting.'
}
#endregion

#region ------------------------------------------- PHASE 1: QUARANTINE (P2)
Head 'PHASE 1/5  -  QUARANTINE   (source of P2 / header From)'
$Quar = @()
if ($Can.Quarantine) {
    $useDomain = (Get-Command Get-QuarantineMessage).Parameters.ContainsKey('Domain')
    $page = 1
    do {
        $p = @{ StartReceivedDate=$Start; EndReceivedDate=$End; Direction='Inbound'
                PageSize=1000; Page=$page; ErrorAction='Stop' }
        # Always by DOMAIN, never by address: the user reports the P2 and there may
        # be more than one P2 in the same domain (that was the case: 9 senders).
        if ($Snd -and $useDomain) { $p.Domain = $Snd.Domain }
        if ($Recipient)           { $p.RecipientAddress = $Recipient }
        try { $batch = Get-QuarantineMessage @p }
        catch {
            if ($p.ContainsKey('Domain')) { Say '    [!] -Domain rejected; filtering client-side.' DarkYellow
                                            $useDomain=$false; $page=1; $Quar=@(); continue }
            Say "    [!] $($_.Exception.Message)" Red; break
        }
        if ($batch) { $Quar += $batch; Say "    page $page -> $($batch.Count)" DarkGray }
        $page++
    } while ($batch -and $batch.Count -eq 1000 -and $page -le 200)

    if ($Snd) { $Quar = @($Quar | Where-Object { $_.SenderAddress -like $Snd.Pattern }) }

    $qHeld = @($Quar | Where-Object { $_.ReleaseStatus -ne 'RELEASED' })
    $qRel  = @($Quar | Where-Object { $_.ReleaseStatus -eq 'RELEASED' })
    Say "[+] $($Quar.Count) in quarantine  |  $($qHeld.Count) still held  |  $($qRel.Count) already released" Green
    if ($qRel.Count) {
        Say "    [!] $($qRel.Count) manual release(s) = false positives already acknowledged by the org itself." Yellow
    }
} else { Say '[!] Skipped (no permission) - without this, quarantine cannot be CONFIRMED.' Red }
#endregion

#region -------------------------------------- PHASE 2: TRACE (by MessageId)
function Get-Trace {
    param([string]$SenderAddress, [string]$RecipientAddress, [string]$MessageId)
    $out = @(); $chunk = 10; $cursor = $End
    while ($cursor -gt $Start) {
        $from = $cursor.AddDays(-$chunk); if ($from -lt $Start) { $from = $Start }
        $p = @{ StartDate=$from; EndDate=$cursor; ErrorAction='Stop' }
        if ($SenderAddress)    { $p.SenderAddress    = $SenderAddress }
        if ($RecipientAddress) { $p.RecipientAddress = $RecipientAddress }
        if ($MessageId)        { $p.MessageId        = $MessageId }
        try {
            $batch = if ($Can.TraceV2) { Get-MessageTraceV2 @p -ResultSize 5000 }
                     else              { Get-MessageTrace   @p -PageSize   5000 }
        } catch { Say "    [!] trace $($from.ToString('MM-dd')): $($_.Exception.Message)" Red; break }
        if ($batch) {
            $out += $batch
            if ($batch.Count -ge 5000) { Say '    [!] chunk saturated (5000). Reduce -Days.' Yellow }
        }
        $cursor = $from
    }
    return $out
}

Head 'PHASE 2/5  -  MESSAGE TRACE   (source of P1 / envelope sender)'
$Trace = @()
if ($Can.TraceV2 -or $Can.TraceV1) {

    # (a) ANCHOR: MessageId taken from quarantine. Immune to P1 != P2.
    $traceCmd = if ($Can.TraceV2) { 'Get-MessageTraceV2' } else { 'Get-MessageTrace' }
    $canById  = (Get-Command $traceCmd).Parameters.ContainsKey('MessageId')
    if (-not $canById) {
        Say "[!] $traceCmd does not expose -MessageId in this module version." DarkYellow
        Say '    The MessageId anchor will be skipped; correlation will depend on the' DarkYellow
        Say '    recipient sweep. Provide -Recipient so you are not flying blind.' DarkYellow
    }

    $ids = @($Quar | ForEach-Object { $_.MessageId } | Where-Object { $_ } | Sort-Object -Unique)
    if ($canById -and $ids.Count) {
        Say "[*] Anchoring by MessageId: $($ids.Count) unique id(s) from quarantine" Yellow
        $i = 0
        foreach ($id in ($ids | Select-Object -First 300)) {
            $i++; Write-Progress -Activity 'Trace by MessageId' -Status "$i/$($ids.Count)" -PercentComplete (100*$i/$ids.Count)
            $Trace += Get-Trace -MessageId $id
        }
        Write-Progress -Activity 'Trace by MessageId' -Completed
        Say "    -> $($Trace.Count) row(s)" DarkGray
    }

    # (b) SWEEP: complainant's mailbox, filtered by DOMAIN (never by address).
    if ($Recipient) {
        Say "[*] Sweeping mailbox $Recipient (client-side domain filter)" Yellow
        $all = Get-Trace -RecipientAddress $Recipient
        Say "    $($all.Count) total message(s) in the window" DarkGray
        $Trace += if ($Snd) { $all | Where-Object { $_.SenderAddress -like $Snd.Pattern } } else { $all }
    }

    # (c) COMPLEMENT: P1 senders already observed, to catch other recipients.
    $p1s = @($Trace | Select-Object -ExpandProperty SenderAddress -Unique | Where-Object { $_ })
    foreach ($s in $p1s) {
        if ($Snd -and $s -notlike $Snd.Pattern) { continue }
        Say "[*] Expanding by P1 sender: $s" Yellow
        $Trace += Get-Trace -SenderAddress $s
    }

    $Trace = @($Trace | Sort-Object MessageId, RecipientAddress, Received -Unique)
    Say "[+] $($Trace.Count) unique event(s)." Green
} else { Say '[!] Skipped (no permission).' Red }
#endregion

#region ------------------------------------- PHASE 3: CORRELATION AND DIVERGENCE
Head 'PHASE 3/5  -  P1 x P2 CORRELATION'

$qIndex = @{}
foreach ($q in $Quar) { $qIndex["$(Norm-MsgId $q.MessageId)|$($q.RecipientAddress.ToLower())"] = $q }

$Corr = foreach ($t in $Trace) {
    $k = "$(Norm-MsgId $t.MessageId)|$($t.RecipientAddress.ToLower())"
    $q = $qIndex[$k]
    $eff = if ($q) {
               if ($q.ReleaseStatus -eq 'RELEASED') { 'QUARANTINED -> released manually' } else { 'QUARANTINED - still held' }
           } elseif ($t.Status -match 'FilteredAsSpam') { 'DELIVERED to Junk (FilteredAsSpam)' }
           elseif ($t.Status -match 'Quarantined')      { 'QUARANTINED' }
           elseif ($t.Status -match 'Failed')           { 'FAILED at transport' }
           else                                          { $t.Status }
    [pscustomobject]@{
        Received             = $t.Received
        EnvelopeSender_P1    = $t.SenderAddress
        HeaderSender_P2      = if ($q) { $q.SenderAddress } else { '' }
        SenderDivergence     = if ($q -and $q.SenderAddress -and ($q.SenderAddress -ne $t.SenderAddress)) { 'YES' } else { '' }
        Recipient            = $t.RecipientAddress
        Subject              = $t.Subject
        TraceStatus          = $t.Status
        QuarantineType       = if ($q) { $q.Type } else { '' }
        QuarantinePolicy     = if ($q) { $q.PolicyName } else { '' }
        ReleaseStatus        = if ($q) { $q.ReleaseStatus } else { '' }
        EffectiveDisposition = $eff
        MessageId            = $t.MessageId
        MessageTraceId       = $t.MessageTraceId
        FromIP               = $t.FromIP
    }
}
$Corr = @($Corr)

# Quarantine items with no matching trace row (permission or window)
$traceKeys = @{}; foreach ($t in $Trace) { $traceKeys["$(Norm-MsgId $t.MessageId)|$($t.RecipientAddress.ToLower())"] = $true }
$Orphans = @($Quar | Where-Object { -not $traceKeys["$(Norm-MsgId $_.MessageId)|$($_.RecipientAddress.ToLower())"] })

$Diverg = @($Corr | Where-Object { $_.SenderDivergence -eq 'YES' })
$PairMap = $Diverg | Group-Object HeaderSender_P2, EnvelopeSender_P1 | ForEach-Object {
    [pscustomobject]@{ HeaderSender_P2 = $_.Group[0].HeaderSender_P2
                       EnvelopeSender_P1 = $_.Group[0].EnvelopeSender_P1
                       Messages = $_.Count }
}

if ($Diverg.Count) {
    Say "[!!] P1/P2 DIVERGENCE in $($Diverg.Count) message(s):" Red
    $PairMap | ForEach-Object { Say ("     header {0}  ->  envelope {1}   ({2} msg)" -f $_.HeaderSender_P2, $_.EnvelopeSender_P1, $_.Messages) DarkYellow }
    Say '     Consequence: an ADDRESS-level allow-list may not work - the filter' Yellow
    Say '     evaluates the envelope. Request the allow at DOMAIN level.' Yellow
} else { Say '[+] No P1/P2 divergence in the correlated messages.' Green }

if ($Orphans.Count) { Say "[!] $($Orphans.Count) quarantine item(s) with no trace row (see orphans.csv)." DarkYellow }
#endregion

#region ------------------------------------------------------ PHASE 4: DETAIL
Head 'PHASE 4/5  -  EVENT DETAIL'
$Details = @()
if ($Detail -and $Can.Detail -and $Trace.Count) {
    $targets = @($Corr | Where-Object { $_.EffectiveDisposition -notmatch '^Delivered$' } | Select-Object -First 25)
    if (-not $targets) { $targets = @($Corr | Select-Object -First 10) }
    Say "[*] Detailing $($targets.Count) message(s)..." Yellow
    foreach ($m in $targets) {
        try {
            $Details += Get-MessageTraceDetailV2 -MessageTraceId $m.MessageTraceId -RecipientAddress $m.Recipient -StartDate $Start -EndDate $End -ErrorAction Stop |
                Select-Object @{n='MessageId';e={$m.MessageId}}, @{n='Sender';e={$m.EnvelopeSender_P1}},
                              @{n='Recipient';e={$m.Recipient}}, Date, Event, Action, Detail
        } catch { Say "    [!] $($m.MessageId): $($_.Exception.Message)" DarkGray }
    }
    Say "[+] $($Details.Count) event(s)." Green
} elseif ($Detail) { Say '[!] Unavailable.' DarkYellow } else { Say '[-] Not requested (-Detail).' DarkGray }
#endregion

#region -------------------------------------- PHASE 5: USER AND TENANT
Head 'PHASE 5/5  -  USER AND TENANT SIDE'
$JunkCfg=$null; $Rules=@(); $AllowBlock=@(); $blocked=@(); $trusted=@(); $suspectRules=@()

if ($Can.JunkCfg) {
    $JunkCfg = Get-MailboxJunkEmailConfiguration -Identity $Recipient
    $blocked = @($JunkCfg.BlockedSendersAndDomains | Where-Object { -not $Snd -or $_ -like "*$($Snd.Domain)*" })
    $trusted = @($JunkCfg.TrustedSendersAndDomains | Where-Object { -not $Snd -or $_ -like "*$($Snd.Domain)*" })
    Say "  Junk filter ..........: $($JunkCfg.Enabled)" White
    Say "  Blocked senders ......: $($JunkCfg.BlockedSendersAndDomains.Count)" White
    if ($blocked) { Say "  [!!] SENDER ON THE USER'S BLOCKED LIST: $($blocked -join ', ')" Red }
    if ($trusted) { Say "  [+]  Already listed as trusted: $($trusted -join ', ')" Green }
}
if ($Can.InboxRule) {
    $Rules = @(Get-InboxRule -Mailbox $Recipient | Where-Object { $_.Enabled })
    $suspectRules = @($Rules | Where-Object {
        $t = ($_.Description + $_.From + $_.SubjectContainsWords + $_.FromAddressContainsWords) -join ' '
        (-not $Snd -or $t -like "*$($Snd.Domain)*") -and ($_.DeleteMessage -or $_.MoveToFolder -or $_.MarkAsRead)
    })
    Say "  Enabled inbox rules ..: $($Rules.Count)" White
    if ($suspectRules) { Say '  [!!] A RULE MAY BE DIVERTING MAIL:' Red
                      $suspectRules | ForEach-Object { Say "       - $($_.Name)" DarkYellow } }
}
if ($Can.AllowBlock -and $Snd) {
    foreach ($lt in 'Sender','Url') {
        try { $AllowBlock += Get-TenantAllowBlockListItems -ListType $lt -ErrorAction Stop |
                             Where-Object { $_.Value -like "*$($Snd.Domain)*" } } catch {}
    }
    if ($AllowBlock) { Say '  [!] Existing entry in the Tenant Allow/Block List:' Yellow
                       $AllowBlock | ForEach-Object { Say "       $($_.Value)  Action=$($_.Action)  Exp=$($_.ExpirationDate)" DarkYellow } }
    else             { Say '  Tenant Allow/Block ...: no entry for the domain' White }
}
#endregion

#region ---------------------------------------------------------- VERDICT
$TenantName = try { (Get-AcceptedDomain -ErrorAction Stop | Where-Object Default).DomainName } catch { '(unknown)' }
Head 'VERDICT'
$nQuar = $Quar.Count
$nHeld = @($Quar | Where-Object { $_.ReleaseStatus -ne 'RELEASED' }).Count
$nRel  = @($Quar | Where-Object { $_.ReleaseStatus -eq 'RELEASED' }).Count
$nJunk = @($Corr | Where-Object { $_.EffectiveDisposition -match 'Junk' }).Count
$nFail = @($Corr | Where-Object { $_.EffectiveDisposition -match 'FAILED' }).Count
$nClean= @($Corr | Where-Object { $_.EffectiveDisposition -eq 'Delivered' }).Count
$fpRate = if ($nQuar) { [math]::Round(100*$nRel/$nQuar) } else { 0 }

$Verdict, $Action =
    if ($nQuar -gt 0) {
        "CONFIRMED - SERVER-SIDE QUARANTINE. $nQuar message(s); $nHeld still held, $nRel manually released ($fpRate% false-positive rate).",
        $(if ($Diverg.Count) { 'Escalate to the security team. Request DOMAIN-level allow (envelope sender differs from header sender - address-level allow may not match).' }
          else               { 'Escalate to the security team with quarantine.csv + correlated.csv as evidence.' })
    }
    elseif ($nFail -gt 0)  { "CONFIRMED - TRANSPORT FAILURE. $nFail message(s) failed.", 'Not quarantine: investigate connector / edge / DBEB before escalating to security.' }
    elseif ($nJunk -gt 0)  { "CONFIRMED - JUNK FOLDER. $nJunk message(s) FilteredAsSpam.", 'Delivered to the Junk folder, not held. The fix is anti-spam policy, not release.' }
    elseif ($blocked)      { 'NOT CONFIRMED - USER BLOCKED SENDER LIST.', 'Caused by the user. Remove from Blocked Senders. No security escalation needed.' }
    elseif ($suspectRules)    { 'NOT CONFIRMED - INBOX RULE.', 'Outlook rule set by the user. Fix it together with the user.' }
    elseif ($nClean -gt 0) { "NOT CONFIRMED - DELIVERED. $nClean message(s) delivered.", 'Mail flow is clean. Check folder / Focused Inbox / client sync.' }
    elseif ($Trace.Count -eq 0) { 'NO DATA - nothing from this sender reached the tenant in the window.', 'Confirm the domain, widen -Days, or suspect a pre-trace block.' }
    else                   { 'INCONCLUSIVE.', 'Run with -Detail to see the event chain.' }

$color = if ($Verdict -like 'CONFIRMED*') { 'Red' } elseif ($Verdict -like 'NOT*') { 'Yellow' } else { 'DarkYellow' }
Say ''; Say "  $Verdict" $color; Say ''; Say "  ACTION: $Action" Cyan; Say ''
Say '  ----------------------------------------------------------------' DarkGray
Say ("  Total quarantine .....: {0}   (held {1} / released {2})" -f $nQuar,$nHeld,$nRel) White
Say ("  Correlated trace .....: {0}" -f $Corr.Count) White
Say ("  P1/P2 divergence .....: {0}" -f $Diverg.Count) White
$Corr | Group-Object EffectiveDisposition | Sort-Object Count -Descending |
    ForEach-Object { Say ("      {0,-38} {1}" -f $_.Name, $_.Count) DarkGray }
#endregion

#region ------------------------------------------------------------ EXPORT
Head 'EXPORT'
function Dump { param($data,$file,$sel)
    if ($data -and $data.Count) {
        if ($sel) { $data | Select-Object $sel | Export-Csv (Join-Path $csvDir ($csvPrefix + $file)) -NoTypeInformation -Encoding UTF8 }
        else      { $data | Export-Csv (Join-Path $csvDir ($csvPrefix + $file)) -NoTypeInformation -Encoding UTF8 }
        Say "  [+] $csvPrefix$file  ($($data.Count))" Green
    }
}
Dump $Corr 'correlated.csv'
Dump $Quar 'quarantine.csv' @('ReceivedTime','SenderAddress','RecipientAddress','Subject','Type','QuarantineTypes','PolicyName','PolicyType','Expires','Released','ReleaseStatus','MessageId','Identity')
Dump $Trace 'messagetrace.csv' @('Received','SenderAddress','RecipientAddress','Subject','Status','FromIP','ToIP','Size','MessageId','MessageTraceId')
Dump $Orphans 'orphans.csv' @('ReceivedTime','SenderAddress','RecipientAddress','Subject','ReleaseStatus','MessageId')
Dump $Details 'trace-detail.csv'
Dump $Rules 'inboxrules.csv' @('Name','Enabled','Priority','Description')
Dump $PairMap 'sender-divergence.csv'

$sb = [System.Text.StringBuilder]::new(); $n = { param($t) [void]$sb.AppendLine($t) }
& $n 'MAIL FLOW EVIDENCE PACKAGE'; & $n '=========================='; & $n ''
& $n "Recipient .......: $(if ($Recipient) { $Recipient } else { '(not specified)' })"
& $n "Sender ..........: $(if ($Snd) { $Snd.Pattern } else { '(not specified)' })"
& $n "Window (UTC) ....: $($Start.ToString('u'))  ->  $($End.ToString('u'))   ($Days days)"
& $n "Tenant ..........: $TenantName"
& $n "Collected by ....: $UPN"
& $n "Collected at ....: $((Get-Date).ToUniversalTime().ToString('u')) UTC"
& $n ''; & $n "VERDICT: $Verdict"; & $n "ACTION : $Action"; & $n ''
& $n 'QUARANTINE (server-side hold - message never reached the mailbox)'
& $n '-----------------------------------------------------------------'
& $n "  Total .................: $nQuar"
& $n "  Still held ............: $nHeld"
& $n "  Manually released .....: $nRel   ($fpRate% of total = confirmed false positives)"
$Quar | Group-Object Type       | Sort-Object Count -Descending | ForEach-Object { & $n ("  type/{0,-24} {1}" -f $_.Name,$_.Count) }
$Quar | Group-Object PolicyName | Sort-Object Count -Descending | ForEach-Object { & $n ("  policy/{0,-22} {1}" -f $_.Name,$_.Count) }
& $n ''
& $n 'EFFECTIVE DISPOSITION (quarantine x trace, correlated by MessageId)'
& $n '-------------------------------------------------------------------'
$Corr | Group-Object EffectiveDisposition | Sort-Object Count -Descending | ForEach-Object { & $n ("  {0,-40} {1}" -f $_.Name,$_.Count) }
& $n ''
& $n 'SENDER IDENTITY (P1 envelope vs P2 header)'
& $n '------------------------------------------'
if ($PairMap) {
    & $n '  DIVERGENCE DETECTED - the filter evaluates the ENVELOPE sender:'
    $PairMap | ForEach-Object { & $n "    header From: $($_.HeaderSender_P2)   ->   envelope: $($_.EnvelopeSender_P1)   ($($_.Messages) msg)" }
    & $n '  => Address-level allow entries may NOT match. Request DOMAIN-level allow.'
} else { & $n '  No divergence detected.' }
& $n ''
& $n 'AFFECTED RECIPIENTS'; & $n '-------------------'
$Quar | Group-Object RecipientAddress | Sort-Object Count -Descending | ForEach-Object { & $n ("  {0,-40} {1}" -f $_.Name,$_.Count) }
& $n ''
& $n 'OBSERVED SENDERS (header From)'; & $n '------------------------------'
$Quar | Group-Object SenderAddress | Sort-Object Count -Descending | ForEach-Object { & $n ("  {0,-46} {1}" -f $_.Name,$_.Count) }
& $n ''
& $n 'RECIPIENT-SIDE CHECKS'; & $n '---------------------'
if ($JunkCfg) {
    & $n "  Junk filter enabled ..: $($JunkCfg.Enabled)"
    & $n "  Blocked senders ......: $($JunkCfg.BlockedSendersAndDomains.Count)"
    if ($blocked) { & $n "  >> SENDER ON USER'S BLOCKED LIST: $($blocked -join ', ')" }
    if ($trusted) { & $n "  >> Sender on user's trusted list: $($trusted -join ', ')" }
} else { & $n '  (not collected)' }
if ($Rules)     { & $n "  Enabled inbox rules ..: $($Rules.Count)" }
if ($suspectRules) { $suspectRules | ForEach-Object { & $n "  >> SUSPECT RULE: $($_.Name)" } }
& $n ''
& $n 'TENANT ALLOW/BLOCK LIST'; & $n '-----------------------'
if ($AllowBlock) { $AllowBlock | ForEach-Object { & $n "  $($_.Value)  Action=$($_.Action)  Expires=$($_.ExpirationDate)" } }
else             { & $n '  (no entry for this domain)' }
& $n ''
if ($Denied) {
    & $n 'DATA SOURCES UNAVAILABLE (permission)'; & $n '-------------------------------------'
    $Denied | ForEach-Object { & $n "  $_" }
    & $n '  NOTE: verdict is based only on the sources above.'; & $n ''
}
& $n 'METHOD NOTES'; & $n '------------'
& $n '  * Quarantine reports the P2 (header From); message trace reports the P1'
& $n '    (envelope MAIL FROM). Correlation is done by MessageId, not by sender.'
& $n '  * Trace Status=Delivered does NOT mean clean delivery: a quarantined'
& $n '    message that was released later shows as Delivered. See ReleaseStatus.'
& $n ''
& $n 'Attachments: correlated.csv, quarantine.csv, messagetrace.csv,'
& $n '             sender-divergence.csv, orphans.csv, console.log'
$sb.ToString() | Set-Content (Join-Path $OutDir 'SUMMARY.txt') -Encoding UTF8
Say '  [+] SUMMARY.txt' Green

#region ------------------------------------------------------ TICKET.TXT
Head 'TICKET NOTES (ticket.txt)'

$isServerSide = $Verdict -like 'CONFIRMED*'
$isUserSide   = ($blocked -and $blocked.Count) -or ($suspectRules -and $suspectRules.Count)

$who = if ($Recipient) {
           $lp = ($Recipient -split '@')[0]
           (($lp -split '[._-]') | Where-Object { $_ } |
             ForEach-Object { $_.Substring(0,1).ToUpper() + $_.Substring(1).ToLower() }) -join ' '
       } else { '' }
$first = if ($who) { ($who -split ' ')[0] } else { 'there' }

$dom     = if ($Snd) { $Snd.Domain } else { 'the reported sender' }
$winTxt  = "{0}-{1} {2} (UTC)" -f $Start.ToString('dd'), $End.ToString('dd MMM yyyy'), ''
$winTxt  = "$($Start.ToString('dd MMM')) - $($End.ToString('dd MMM yyyy')), UTC"
$types   = (($Quar | Group-Object Type | Sort-Object Count -Descending |
             ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ', ')
$pol     = (($Quar | Group-Object PolicyName | Sort-Object Count -Descending |
             Select-Object -First 1).Name)
$polType = ($Quar | Select-Object -First 1).PolicyType
$nSnd    = @($Quar | Select-Object -ExpandProperty SenderAddress -Unique).Count
$rcpts   = @($Quar | Group-Object RecipientAddress | Sort-Object Count -Descending)
$rcptTxt = (($rcpts | Select-Object -First 6 | ForEach-Object { $_.Name }) -join ', ')
if ($rcpts.Count -gt 6) { $rcptTxt += " (+$($rcpts.Count - 6) more)" }
$oldest  = ($Quar | Where-Object { $_.ReleaseStatus -ne 'RELEASED' } |
            Sort-Object { [datetime]$_.ReceivedTime } | Select-Object -First 1)
$files   = @('SUMMARY.txt','correlated.csv') +
           $(if ($Quar.Count)    { 'quarantine.csv' }) +
           $(if ($PairMap)       { 'sender-divergence.csv' }) | Where-Object { $_ }

# ---------------------------------------------------------------- INTERNAL
$i = [System.Text.StringBuilder]::new(); $w = { param($t) [void]$i.AppendLine($t) }

& $w "$dom - inbound mail flow review"
& $w ''
& $w "Findings from a $Days-day review ($winTxt), tenant ${TenantName}:"
& $w ''
if ($nQuar -gt 0) {
    & $w "- $nQuar inbound message(s) from $dom quarantined."
    & $w "- Verdict: $types."
    & $w "- Policy: `"$pol`" ($polType), action = quarantine. Server-side hold,"
    & $w "  not Junk-folder filing."
    & $w "- $nSnd distinct sender address(es); $($rcpts.Count) recipient(s) affected: $rcptTxt."
    if ($nRel -gt 0) {
        & $w "- $nRel of $nQuar ($fpRate%) manually released by staff. $nHeld still held$(if ($oldest) { ", oldest since $([datetime]$oldest.ReceivedTime | Get-Date -Format 'dd MMM')" })."
    } else {
        & $w "- $nHeld still held$(if ($oldest) { ", oldest since $([datetime]$oldest.ReceivedTime | Get-Date -Format 'dd MMM')" })."
    }
}
if ($nJunk -gt 0)  { & $w "- $nJunk message(s) delivered to the Junk folder (FilteredAsSpam), not quarantined." }
if ($nFail -gt 0)  { & $w "- $nFail message(s) failed at transport." }
if ($nClean -gt 0) { & $w "- $nClean message(s) delivered normally." }
if ($nQuar -eq 0 -and $Trace.Count -eq 0) { & $w "- No inbound mail from $dom reached the tenant in the window." }
if ($blocked)   { & $w "- Sender is present on the recipient's own Blocked Senders list: $($blocked -join ', ')." }
if ($suspectRules) { & $w "- Recipient has an enabled inbox rule matching this sender: $(($suspectRules | ForEach-Object { $_.Name }) -join ', ')." }
if ($AllowBlock) { & $w "- Existing Tenant Allow/Block entry: $(($AllowBlock | ForEach-Object { "$($_.Value) [$($_.Action)]" }) -join ', ')." }

if ($PairMap) {
    $ex = $PairMap | Sort-Object Messages -Descending | Select-Object -First 1
    $exId = ($Diverg | Where-Object { $_.HeaderSender_P2 -eq $ex.HeaderSender_P2 } | Select-Object -First 1).MessageId
    & $w ''
    & $w 'One detail relevant to any remediation scope: header From and envelope'
    & $w "sender diverge. On Message-ID"
    & $w "$exId"
    & $w "the header From is $($ex.HeaderSender_P2) while the envelope sender is"
    & $w "$($ex.EnvelopeSender_P1). Message trace matches on the envelope, which is"
    & $w 'why a search on the reported address returns nothing.'
}

& $w ''
& $w "Attachments: $($files -join ', ')."
& $w 'Method: Get-QuarantineMessage + Get-MessageTrace, correlated by Message-ID.'
& $w 'Timestamps UTC.'
if ($isServerSide -and -not $isUserSide) { & $w ''; & $w 'Handing over to the security team for assessment.' }

# ---------------------------------------------------------------- CUSTOMER
$c = [System.Text.StringBuilder]::new(); $v = { param($t) [void]$c.AppendLine($t) }

& $v "Hi $first,"
& $v ''
& $v "I've reviewed the mail flow for $dom over the last $Days days. Here's what I found."
& $v ''

if ($nQuar -gt 0) {
    & $v "$nQuar message(s) from $dom were held in the server-side quarantine and"
    & $v 'classified as spam by our mail filter. They were stopped before reaching'
    & $v "your mailbox, which is why they weren't in your Junk folder either."
    & $v ''
    if ($nRel -gt 0 -and $nHeld -gt 0) {
        & $v "$nRel of them had already been released manually, so those should be in"
        & $v "your mailbox now. The remaining $nHeld are still being held - I've released"
        & $v 'nothing on my own; please let me know if you need them delivered.'
    } elseif ($nRel -gt 0) {
        & $v "All $nRel of them had already been released manually, so they should be in"
        & $v 'your mailbox now.'
    } else {
        & $v "All $nHeld are still being held. I have not released anything yet - let me"
        & $v 'know if you need them delivered while the underlying issue is reviewed.'
    }
    & $v ''
    & $v 'Releasing individual messages only fixes them one at a time, so I have'
    & $v 'passed the full findings to our security team, who own the mail'
    & $v 'filtering policy. Any permanent change to how this sender is treated'
    & $v 'will come from them.'
}
elseif ($nJunk -gt 0) {
    & $v "$nJunk message(s) from $dom were delivered, but our mail filter classified"
    & $v 'them as spam and filed them into your Junk Email folder. They are not'
    & $v 'blocked - please check that folder.'
    & $v ''
    & $v 'I have passed the findings to our security team, who own the mail'
    & $v 'filtering policy, so this can be looked at properly.'
}
elseif ($nFail -gt 0) {
    & $v "$nFail message(s) from $dom failed to be accepted by our mail system and"
    & $v 'never reached the mailbox. This is a delivery failure rather than a spam'
    & $v 'classification.'
    & $v ''
    & $v 'I have passed the findings to our security team for assessment.'
}
elseif ($blocked) {
    & $v "The messages are not being blocked by our mail filter. $dom is on the"
    & $v "Blocked Senders list inside your own mailbox, which is what's diverting"
    & $v 'them.'
    & $v ''
    & $v 'This one is fixable directly on your account - let me know and I can'
    & $v 'remove that entry for you. No security escalation is needed.'
}
elseif ($suspectRules) {
    & $v 'The messages are not being blocked by our mail filter. There is a rule in'
    & $v 'your Outlook that matches this sender and is moving the messages.'
    & $v ''
    & $v 'I can walk you through removing it, or remove it for you if you prefer.'
    & $v 'No security escalation is needed.'
}
elseif ($nClean -gt 0) {
    & $v "$nClean message(s) from $dom were delivered to your mailbox normally in"
    & $v 'this period. Nothing was quarantined or filtered.'
    & $v ''
    & $v "If you still can't see them, it's most likely a folder, rule or Focused"
    & $v 'Inbox issue on the client side - happy to take a look with you.'
}
else {
    & $v "I could not find any inbound mail from $dom reaching us in the last"
    & $v "$Days days - nothing quarantined and nothing delivered."
    & $v ''
    & $v 'Could you confirm the exact sender address, or forward me one of the'
    & $v 'messages they say was sent? That will let me widen the search correctly.'
}

& $v ''
& $v 'Best regards,'

# ---------------------------------------------------------------- WRITE
$j = [System.Text.StringBuilder]::new()
[void]$j.AppendLine('=============== INTERNAL COMMENT ==============='); [void]$j.AppendLine('')
[void]$j.Append($i.ToString()); [void]$j.AppendLine('')
[void]$j.AppendLine('=============== CUSTOMER COMMENT ==============='); [void]$j.AppendLine('')
[void]$j.Append($c.ToString())
$j.ToString() | Set-Content (Join-Path $OutDir 'ticket.txt') -Encoding UTF8
Say '  [+] ticket.txt   (internal + customer, ready to paste)' Green
Say ''
Write-Host $j.ToString() -ForegroundColor White
#endregion

try { Stop-Transcript | Out-Null } catch {}
Head 'DONE'
Say "  Evidence in: $OutDir" Cyan; Say "  CSV in     : $csvDir" Cyan; Say ''
if (-not $NonInteractive) {
    if ((Ask 'Open the folder? (Y/n)' -Default 'Y') -match '^[sSyY]') { Start-Process explorer.exe $OutDir }
}
