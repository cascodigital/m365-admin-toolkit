<#
.SYNOPSIS
    Full inbox rules audit: hidden rules, shadowing (dead rule),
    corruption, external forwarding and mailbox-level forwarding.

.DESCRIPTION
    Built from a real case where a simpler script returned
    "25 rules, none of them forward" -- a true and useless sentence that hid
    three real problems.

    WHAT IT FINDS THAT SIMPLER CHECKS MISS:

    [A] DEAD RULE BY SHADOWING  <<< the main finding
        Two rules with the SAME condition, the first one with 'stop processing more rules'.
        The second one never runs. In the real case: two rules both matched the same
        vendor noreply sender and sent to DIFFERENT folders. The second rule had been dead
        code for months and no human noticed by eyeballing the list.
        The script compares the CONDITION SIGNATURE, not the Description text.

    [B] HIDDEN USER RULE
        Get-InboxRule without -IncludeHidden hides rules that DO RUN. In the real case there was
        a user rule at priority 1, Enabled, with stop --
        invisible in Outlook's Rules and Alerts dialog. The user could not see,
        edit or turn it off. Junk E-mail Rule and Microsoft.Exchange.OOF.* are system rules
        and the script separates them as noise.

    [C] RULE THAT DELETES EXCHANGE NOTICES
        A rule that deletes messages from MicrosoftExchange<32hex>@ kills NDRs, quota
        warnings and system notifications. The user goes blind to exactly the channel that
        would explain their problem.

    [D] CORRUPTED RULES
        They fail to parse and are SKIPPED from the output -- the cmdlet only emits a WARNING. That is why
        InError is always False: the bad ones never even become objects. Count warnings, not rows.

    [E] EXTERNAL FORWARDING, both halves
        Inbox rule AND mailbox-level forwarding, both evaluated against the tenant's
        accepted domains (Get-AcceptedDomain), so no hardcoded domain list can go stale.

    KNOWN PITFALL:
    >> The Priority field is NOT stable between the call with and without -IncludeHidden. <<
    The same rule can show as Priority 3 without hidden and 4 with; another as 6 without
    and 5 with. Comparing priority across the two listings produces phantom "gaps"
    and false conclusions.
    The stable key is RuleIdentity. All ordering analysis here runs on ONE
    listing only -- the full one, with -IncludeHidden.

.PARAMETER MailboxAddress
    The mailbox to audit (UPN or primary SMTP address).

.PARAMETER AdminUPN
    Admin account used to connect to Exchange Online if no session is active.

.PARAMETER OutFolder
    Folder for the CSV output. Default: 'csv' subfolder next to the script
    (falls back to $env:TEMP if not writable).

.NOTES
    READ-ONLY: zero write cmdlets against the tenant.

    PREREQ: an Exchange admin role (e.g. "Exchange Recipient Administrator", activated in PIM
    if you use it) BEFORE connecting. Without it, Get-Mailbox reports
    "object couldn't be found" -- a false 404, it is permissions.

.EXAMPLE
    .\Audit-InboxRules.ps1 -MailboxAddress "user@domain.com" -AdminUPN "ADMIN@DOMAIN.COM"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$MailboxAddress,

    [Parameter(Mandatory = $true)]
    [string]$AdminUPN,

    [Parameter(Mandatory = $false)]
    [string]$OutFolder
)

# Internal domains: filled from Get-AcceptedDomain right after connecting (no hardcoded list).
$TrustedDomains = @()

# system rules that are hidden in EVERY mailbox -- noise, not a finding
$SystemHiddenPattern = '^Junk E-mail Rule$|^Microsoft\.Exchange\.OOF\.'

# Exchange system mailbox (NDR, quota, notifications)
$SystemMailboxPattern = 'MicrosoftExchange[0-9a-f]{32}@'

# ---------------------------------------------------------------------------
# 0. Output folder -- PROOF of write access, not Test-Path.
#    Fallback is TEMP, not Desktop: Desktop may be redirected by OneDrive
#    KFM, and a CSV with PII should not sync.
# ---------------------------------------------------------------------------
$__root = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.ProviderPath }
$csvDir = if ($OutFolder) { $OutFolder } else { Join-Path $__root 'csv' }   # CSV always in the csv subfolder next to the script, even if the script is moved (-OutFolder still wins)
if (-not $OutFolder -and (-not (Test-Path $csvDir)) -and (Test-Path (Join-Path (Split-Path $__root) 'csv'))) { $csvDir = Join-Path (Split-Path $__root) 'csv' }   # script in a subfolder: use the csv folder of the parent
try { if (-not (Test-Path $csvDir)) { $null = New-Item -ItemType Directory -Path $csvDir -ErrorAction Stop } } catch {}
try { $__p = Join-Path $csvDir "w$PID.tmp"; [IO.File]::WriteAllText($__p,'x'); [IO.File]::Delete($__p) }
catch { $csvDir = $env:TEMP; Write-Host "WARNING: CSV output redirected to $csvDir" -ForegroundColor Yellow }

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$slug  = ($MailboxAddress.Split('@')[0]) -replace '[^\w\.\-]','_'

# ---------------------------------------------------------------------------
# 1. Connection
# ---------------------------------------------------------------------------
try {
    if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
        Write-Host "Connecting to EXO as $AdminUPN ..." -ForegroundColor Cyan
        Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false
    } else { Write-Host "EXO session already active." -ForegroundColor Green }
    $TrustedDomains = @(Get-AcceptedDomain | ForEach-Object { [string]$_.DomainName })
    Write-Host "Internal domains: $($TrustedDomains -join ', ')" -ForegroundColor DarkGray
}
catch {
    Write-Host "CONNECTION FAILED: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "'Error Acquiring Token' + RuntimeBroker = WAM broker broken on the CLIENT," -ForegroundColor Yellow
    Write-Host "not PIM or a role -- the crash happens before any session exists." -ForegroundColor Yellow
    return
}

try { $mailbox = Get-Mailbox -Identity $MailboxAddress -ErrorAction Stop }
catch {
    Write-Host "Mailbox not found: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "ATTENTION: 'couldn't be found' is almost always inactive PIM disguised as a 404." -ForegroundColor Yellow
    return
}

Write-Host "`nMailbox: $($mailbox.DisplayName) <$($mailbox.PrimarySmtpAddress)>" -ForegroundColor Green
Write-Host "Type: $($mailbox.RecipientTypeDetails) | Output: $csvDir" -ForegroundColor DarkGray

$findings = [System.Collections.Generic.List[object]]::new()
function Add-Finding {
    param($Severity,$Source,$Item,$Detail,$Action)
    $findings.Add([PSCustomObject]@{
        Mailbox=$MailboxAddress; Severity=$Severity; Source=$Source
        Item=$Item; Detail=$Detail; Action=$Action
    })
}

# ---------------------------------------------------------------------------
# 2. Signatures -- compare the CONDITION, not the Description prose
#    (Description is localized text and changes shape without the rule changing)
# ---------------------------------------------------------------------------
$CondProps = @(
    'From','SentTo','FromAddressContainsWords','RecipientAddressContainsWords'
    'SubjectContainsWords','SubjectOrBodyContainsWords','BodyContainsWords'
    'HeaderContainsWords','MessageTypeMatches','SentOnlyToMe','MyNameInToBox'
    'MyNameInCcBox','MyNameInToOrCcBox','MyNameNotInToBox','HasAttachment'
    'WithImportance','WithSensitivity','FromSubscription','HasClassification'
    'FlaggedForAction','WithinSizeRangeMinimum','WithinSizeRangeMaximum'
    'ReceivedAfterDate','ReceivedBeforeDate'
    'ExceptIfFrom','ExceptIfSentTo','ExceptIfSubjectContainsWords'
)
$ActionProps = @(
    'MoveToFolder','CopyToFolder','DeleteMessage','PermanentDelete'
    'ForwardTo','ForwardAsAttachmentTo','RedirectTo','SendTextMessageNotificationTo'
    'MarkAsRead','MarkImportance','ApplyCategory','ApplySystemCategory'
)

function Get-Signature {
    param($Rule,[string[]]$Props)
    $parts = foreach ($p in $Props) {
        $pr = $Rule.PSObject.Properties[$p]
        if (-not $pr -or $null -eq $pr.Value) { continue }
        if ($pr.Value -is [bool] -and -not $pr.Value) { continue }   # $false is not a condition
        $vals = @($pr.Value | ForEach-Object { "$_".Trim() } | Where-Object { $_ }) | Sort-Object
        if ($vals.Count -eq 0) { continue }
        "$p=$($vals -join '|')"
    }
    ($parts | Sort-Object) -join ' && '
}

# ---------------------------------------------------------------------------
# 3. Read -- both listings. Warnings captured (corrupted rule).
#    No -ResultSize: the Unlimited sentinel is rejected by some EXO sessions.
# ---------------------------------------------------------------------------
$wVis = @()
try { $visible = @(Get-InboxRule -Mailbox $MailboxAddress -WarningVariable wVis -WarningAction SilentlyContinue -ErrorAction Stop) }
catch { Write-Host "Failed to read inbox rules: $($_.Exception.Message)" -ForegroundColor Red; return }

$hasHidden = (Get-Command Get-InboxRule).Parameters.Keys -contains 'IncludeHidden'
$wAll = @()
if ($hasHidden) {
    try { $allRules = @(Get-InboxRule -Mailbox $MailboxAddress -IncludeHidden -WarningVariable wAll -WarningAction SilentlyContinue -ErrorAction Stop) }
    catch { $allRules = $visible; $hasHidden = $false }
} else { $allRules = $visible }

$corruptCount = [Math]::Max(@($wVis).Count, @($wAll).Count)

# stable key: RuleIdentity. NEVER Priority (unstable between listings).
$visibleIds  = @($visible | ForEach-Object { "$($_.RuleIdentity)" })
$hiddenRules = @($allRules | Where-Object { "$($_.RuleIdentity)" -notin $visibleIds })

Write-Host "`n=== INVENTORY ===" -ForegroundColor Yellow
Write-Host ("Visible in Outlook ...: {0}" -f $visible.Count)
Write-Host ("Total incl. hidden ...: {0}{1}" -f $allRules.Count, $(if(-not $hasHidden){' (session does not expose -IncludeHidden)'}else{''}))
Write-Host ("Hidden ...............: {0}" -f $hiddenRules.Count)
Write-Host ("Corrupted ............: {0}" -f $corruptCount) -ForegroundColor $(if($corruptCount -gt 0){'Red'}else{'Green'})

foreach ($w in @($wAll)) {
    Add-Finding 'CORRUPT' 'InboxRule' "$w" 'Skipped by the parser -- content and destination unknown.' `
        'Run Remove-CorruptInboxRules.ps1 -DryRun first.'
}

# ---------------------------------------------------------------------------
# 4. Working table
# ---------------------------------------------------------------------------
$map = foreach ($rule in $allRules) {
    $rid      = "$($rule.RuleIdentity)"
    $isHidden = $rid -notin $visibleIds
    [PSCustomObject]@{
        Order        = [int]$rule.Priority
        RuleName     = "$($rule.Name)"
        Hidden       = $isHidden
        System       = ($isHidden -and "$($rule.Name)" -match $SystemHiddenPattern)
        Enabled      = $rule.Enabled
        Stop         = [bool]$rule.StopProcessingRules
        CondSig      = Get-Signature $rule $CondProps
        ActSig       = Get-Signature $rule $ActionProps
        MoveToFolder = "$($rule.MoveToFolder)"
        ForwardTo             = if ($rule.ForwardTo) { (@($rule.ForwardTo|%{"$_"}) -join '; ') } else { '' }
        ForwardAsAttachmentTo = if ($rule.ForwardAsAttachmentTo) { (@($rule.ForwardAsAttachmentTo|%{"$_"}) -join '; ') } else { '' }
        RedirectTo            = if ($rule.RedirectTo) { (@($rule.RedirectTo|%{"$_"}) -join '; ') } else { '' }
        TextNotifyTo          = if ($rule.SendTextMessageNotificationTo) { (@($rule.SendTextMessageNotificationTo|%{"$_"}) -join '; ') } else { '' }
        DeleteMessage = [bool]$rule.DeleteMessage
        RuleIdentity  = $rid
        Description   = (("$($rule.Description)" -replace "`r`n",' ') -replace '\s+',' ').Trim()
    }
}

# ---------------------------------------------------------------------------
# 5. [A] SHADOWING -- the dead rule
# ---------------------------------------------------------------------------
Write-Host "`n=== SHADOWING (rule that never runs) ===" -ForegroundColor Yellow
$deadCount = 0
foreach ($g in ($map | Where-Object { $_.CondSig } | Group-Object CondSig | Where-Object Count -gt 1)) {
    $ord = @($g.Group | Sort-Object Order)
    for ($i = 0; $i -lt $ord.Count - 1; $i++) {
        if (-not $ord[$i].Stop) { continue }
        $blocker = $ord[$i]
        for ($j = $i + 1; $j -lt $ord.Count; $j++) {
            $dead = $ord[$j]
            $sameAction = ($dead.ActSig -eq $blocker.ActSig)
            $sev  = if ($sameAction) { 'REDUNDANT' } else { 'DEAD RULE' }
            $action = if ($sameAction) {
                "Harmless duplicate. Delete '$($dead.RuleName)' to clean up the list."
            } else {
                "CONFIRM with the user which folder is correct, and delete the losing rule."
            }
            Add-Finding $sev 'InboxRule' "$($dead.RuleName)  [order $($dead.Order)]" `
                "Condition identical to '$($blocker.RuleName)' (order $($blocker.Order)), which has STOP. This rule's destination: '$($dead.MoveToFolder)' | winning rule's destination: '$($blocker.MoveToFolder)'." $action
            if (-not $sameAction) {
                $deadCount++
                Write-Host "  [DEAD]  '$($dead.RuleName)' never runs -- blocked by '$($blocker.RuleName)'" -ForegroundColor Red
                Write-Host "          expected in '$($dead.MoveToFolder)', goes to '$($blocker.MoveToFolder)'" -ForegroundColor Red
            } else {
                Write-Host "  [dup]   '$($dead.RuleName)' is redundant with '$($blocker.RuleName)' (same destination)" -ForegroundColor DarkYellow
            }
        }
        break
    }
}
if ($deadCount -eq 0) { Write-Host "No dead rule by identical condition." -ForegroundColor DarkGray }
Write-Host "Note: only detects IDENTICAL conditions. A broader rule swallowing a narrower one" -ForegroundColor DarkGray
Write-Host "      needs a human eye -- see the execution map at the end." -ForegroundColor DarkGray

# ---------------------------------------------------------------------------
# 6. [B] HIDDEN USER RULES
# ---------------------------------------------------------------------------
Write-Host "`n=== HIDDEN RULES ===" -ForegroundColor Yellow
$hiddenUser = @($map | Where-Object { $_.Hidden -and -not $_.System })
$hiddenSys  = @($map | Where-Object { $_.Hidden -and $_.System })
if ($hiddenSys.Count -gt 0) {
    Write-Host "System (normal, ignore): $(($hiddenSys.RuleName) -join ', ')" -ForegroundColor DarkGray
}
if ($hiddenUser.Count -eq 0) {
    Write-Host "No hidden user rules." -ForegroundColor Green
} else {
    foreach ($o in $hiddenUser) {
        $sev = if ($o.Enabled) { 'HIDDEN ACTIVE' } else { 'HIDDEN INACTIVE' }
        Write-Host "  [$sev] '$($o.RuleName)' (order $($o.Order), stop=$($o.Stop))" -ForegroundColor Red
        Write-Host "         $($o.Description)" -ForegroundColor DarkGray
        Add-Finding $sev 'InboxRule' "$($o.RuleName)  [order $($o.Order)]" `
            "Does not appear in Outlook's Rules and Alerts. Enabled=$($o.Enabled), Stop=$($o.Stop). $($o.Description)" `
            'Show the user and ask whether the behavior is expected. Can only be removed with Remove-InboxRule.'
    }
}

# ---------------------------------------------------------------------------
# 7. [C] RULE THAT DELETES EXCHANGE NOTICES
# ---------------------------------------------------------------------------
foreach ($m in $map) {
    if ($m.CondSig -match $SystemMailboxPattern -and ($m.DeleteMessage -or $m.MoveToFolder -match 'Deleted|Lixeira')) {   # 'Lixeira' = Portuguese-localized Deleted Items
        Add-Finding 'BLINDSPOT' 'InboxRule' "$($m.RuleName)  [order $($m.Order)]" `
            'Discards messages from the Exchange system mailbox: NDRs, quota warnings and failure notifications.' `
            'Inform the user -- they will not see any Exchange notice while this rule exists.'
    }
}

# ---------------------------------------------------------------------------
# 8. [E] FORWARDING -- mailbox AND rules, actually evaluated
# ---------------------------------------------------------------------------
Write-Host "`n=== FORWARDING ===" -ForegroundColor Yellow
function Test-Targets {
    param([string]$Source,[string]$Item,[string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return }
    $addrs = @([regex]::Matches($Raw,'[\w\.\-\+]+@[\w\.\-]+\.\w+') | ForEach-Object { $_.Value } | Select-Object -Unique)
    foreach ($a in $addrs) {
        if ($TrustedDomains -notcontains $a.Split('@')[1]) {
            Add-Finding 'EXTERNAL' $Source $Item "Destination outside the internal domains: $a" `
                'Likely cause of the anti-autoforward block. Alternative: shared mailbox or internal group.'
            Write-Host "  [EXTERNAL] $Item -> $a" -ForegroundColor Red
        }
    }
    # a target that does not resolve to SMTP (directory object) cannot be assumed safe
    if ($addrs.Count -eq 0 -or $Raw -match '\[EX:') {
        Add-Finding 'REVIEW' $Source $Item "Target did not resolve to SMTP: $Raw" 'Check manually where it points.'
    }
}

$mbxFwdAddr = "$($mailbox.ForwardingAddress)"
$mbxFwdSmtp = "$($mailbox.ForwardingSmtpAddress)"
if ($mbxFwdAddr -or $mbxFwdSmtp) {
    Write-Host "Mailbox-level forwarding ACTIVE:" -ForegroundColor Red
    Write-Host "  ForwardingAddress .....: $mbxFwdAddr"
    Write-Host "  ForwardingSmtpAddress .: $mbxFwdSmtp"
    Write-Host "  DeliverToMailboxAndForward: $($mailbox.DeliverToMailboxAndForward)"
    Test-Targets 'Mailbox' 'ForwardingAddress'     $mbxFwdAddr
    Test-Targets 'Mailbox' 'ForwardingSmtpAddress' $mbxFwdSmtp
    if (-not $mailbox.DeliverToMailboxAndForward) {
        Add-Finding 'REVIEW' 'Mailbox' 'DeliverToMailboxAndForward=False' `
            'Forwards WITHOUT keeping a local copy -- the user does not see the message in their own mailbox.' 'Confirm whether this is intentional.'
    }
} else { Write-Host "No mailbox-level forwarding." -ForegroundColor DarkGray }

$withFwd = @($map | Where-Object { $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo -or $_.TextNotifyTo })
if ($withFwd.Count -eq 0) { Write-Host "No rule forwards or redirects." -ForegroundColor DarkGray }
foreach ($f in $withFwd) {
    Test-Targets 'InboxRule' "$($f.RuleName) :: ForwardTo"       $f.ForwardTo
    Test-Targets 'InboxRule' "$($f.RuleName) :: ForwardAsAttach" $f.ForwardAsAttachmentTo
    Test-Targets 'InboxRule' "$($f.RuleName) :: RedirectTo"      $f.RedirectTo
    Test-Targets 'InboxRule' "$($f.RuleName) :: TextNotifyTo"    $f.TextNotifyTo
}

# duplicate names -- Outlook's stamp when a rule is duplicated
foreach ($d in ($map | Group-Object RuleName | Where-Object Count -gt 1)) {
    Add-Finding 'DUPLICATE' 'InboxRule' "$($d.Name)" "$($d.Count) rules with the identical name." 'Check which one to keep.'
}
foreach ($s in ($map | Where-Object { $_.RuleName -match '\(\d+\)\s*$' })) {
    Add-Finding 'DUPLICATE' 'InboxRule' "$($s.RuleName)" 'The "(N)" suffix is Outlook''s stamp when duplicating a rule.' 'Check whether the original still exists.'
}

# ---------------------------------------------------------------------------
# 9. Export -- ALWAYS, even when negative (keeps evidence of a clean result)
# ---------------------------------------------------------------------------
$csvMap      = Join-Path $csvDir "rules-map-$slug-$stamp.csv"
$csvFindings = Join-Path $csvDir "rules-findings-$slug-$stamp.csv"
$map | Sort-Object Order | Export-Csv -Path $csvMap -NoTypeInformation -Encoding UTF8
if ($findings.Count -gt 0) { $findings | Export-Csv -Path $csvFindings -NoTypeInformation -Encoding UTF8 }
else {
    [PSCustomObject]@{ Mailbox=$MailboxAddress; Severity='NOTHING FOUND'; Source='-'; Item='-'
        Detail="$($allRules.Count) rules, 0 corrupted, 0 hidden user rules, 0 dead, 0 external destinations."; Action='-'
    } | Export-Csv -Path $csvFindings -NoTypeInformation -Encoding UTF8
}

# ============ SUMMARY -- LAST BLOCK, NOTHING AFTER ============
$c = @{}
foreach ($k in 'CORRUPT','DEAD RULE','HIDDEN ACTIVE','EXTERNAL','BLINDSPOT','REDUNDANT','DUPLICATE','REVIEW') {
    $c[$k] = @($findings | Where-Object Severity -eq $k).Count
}

Write-Host "`n`n=============== EXECUTION MAP ===============" -ForegroundColor Cyan
Write-Host "Order in which the rules touch each message. 'Stop' cuts off the ones below" -ForegroundColor DarkGray
Write-Host "ONLY for the message that matched that rule." -ForegroundColor DarkGray
$map | Sort-Object Order | Select-Object Order,
    @{n='Hid';e={ if ($_.System) {'sys'} elseif ($_.Hidden) {'YES'} else {''} }},
    @{n='Stop';e={ if ($_.Stop) {'STOP'} else {''} }}, Enabled,
    @{n='Rule';e={ $_.RuleName.Substring(0,[Math]::Min(42,$_.RuleName.Length)) }},
    @{n='Folder';e={ $_.MoveToFolder.Substring(0,[Math]::Min(32,$_.MoveToFolder.Length)) }} |
    Format-Table -AutoSize

Write-Host "=================== SUMMARY ===================" -ForegroundColor Cyan
Write-Host ("Mailbox ..............: {0}" -f $MailboxAddress)
Write-Host ("Rules (visible/total) : {0} / {1}" -f $visible.Count, $allRules.Count)
Write-Host ("Corrupted ............: {0}" -f $c['CORRUPT'])       -ForegroundColor $(if($c['CORRUPT']){'Red'}else{'Green'})
Write-Host ("DEAD RULES ...........: {0}" -f $c['DEAD RULE'])     -ForegroundColor $(if($c['DEAD RULE']){'Red'}else{'Green'})
Write-Host ("Hidden user rules ....: {0}" -f $c['HIDDEN ACTIVE']) -ForegroundColor $(if($c['HIDDEN ACTIVE']){'Red'}else{'Green'})
Write-Host ("External destinations : {0}" -f $c['EXTERNAL'])      -ForegroundColor $(if($c['EXTERNAL']){'Red'}else{'Green'})
Write-Host ("Deletes Exchange notices: {0}" -f $c['BLINDSPOT'])   -ForegroundColor $(if($c['BLINDSPOT']){'Yellow'}else{'Green'})
Write-Host ("Redundant/duplicate ..: {0}" -f ($c['REDUNDANT'] + $c['DUPLICATE'])) -ForegroundColor DarkGray
Write-Host ("To review ............: {0}" -f $c['REVIEW'])        -ForegroundColor DarkGray

if ($findings.Count -gt 0) {
    Write-Host ""
    $findings | Where-Object { $_.Severity -notin 'REDUNDANT','DUPLICATE' } |
        Select-Object Severity,Item,Detail | Format-Table -AutoSize -Wrap
}

Write-Host ""
if ($c['CORRUPT'] -gt 0) {
    Write-Host "VERDICT: INCONCLUSIVE -- $($c['CORRUPT']) unreadable rule(s). Run Remove-CorruptInboxRules.ps1 -DryRun." -ForegroundColor Red
} elseif ($c['DEAD RULE'] -gt 0) {
    Write-Host "VERDICT: $($c['DEAD RULE']) rule(s) NEVER run. That is the 'my rules are not working' complaint." -ForegroundColor Red
} elseif ($c['HIDDEN ACTIVE'] -gt 0) {
    Write-Host "VERDICT: there is an active rule invisible in the user's Outlook. Show it to them before touching anything." -ForegroundColor Red
} elseif ($c['EXTERNAL'] -gt 0) {
    Write-Host "VERDICT: external destination -- cause of the anti-autoforward block." -ForegroundColor Red
} else {
    Write-Host "VERDICT: clean. No dead, hidden or corrupted rule and no external destination." -ForegroundColor Green
}

Write-Host ""
if (Test-Path $csvMap)      { Write-Host "Map .....: $csvMap" }      else { Write-Host "ATTENTION: map was NOT written" -ForegroundColor Red }
if (Test-Path $csvFindings) { Write-Host "Findings : $csvFindings" } else { Write-Host "ATTENTION: findings were NOT written" -ForegroundColor Red }
