<#
.SYNOPSIS
    Configures a catch-all mailbox for one or more domains in Microsoft 365 / Exchange Online.

.DESCRIPTION
    Implements a catch-all: mail sent to NON-EXISTENT addresses in the selected domains is
    redirected to one collector mailbox, while mail to any EXISTING recipient is delivered
    normally.

    What the script does:
      1. Removes the transport rule and exception group created by a previous run
         (same names / same group address) so it can be re-run safely.
      2. Sets each selected accepted domain to InternalRelay (required: with an
         Authoritative domain, Exchange rejects unknown recipients with NDR 5.1.10
         before any transport rule runs).
      3. Creates a dynamic distribution group (DDG) that contains EVERY mail-enabled
         recipient type in the tenant (see "How existing recipients are excluded").
      4. Creates a transport rule (priority 0) that redirects mail for the selected
         domains to the collector mailbox, EXCEPT when the recipient is a member of that
         DDG, or is one of the recipients listed in the rule's ExceptIfSentTo list.

    HOW EXISTING RECIPIENTS ARE EXCLUDED (the robust part)
    Earlier versions only excluded RecipientTypeDetails 'UserMailbox', so mail sent to
    shared mailboxes, groups, Microsoft 365 groups, room mailboxes, mail users and mail
    contacts was wrongly redirected to the catch-all. Now:

      a) The exception DDG uses a RecipientFilter that covers all mail-enabled types:
         user/shared/room/equipment/scheduling mailboxes, Microsoft 365 groups
         (GroupMailbox), distribution and mail-enabled security groups, mail users,
         guest mail users, mail contacts, mail-enabled public folders and team mailboxes.
         Because it is dynamic, recipients created later are covered automatically.

      b) Coverage is verified before the rule is created: the script compares
         Get-Recipient (all recipients) with the DDG preview (Get-Recipient
         -RecipientPreviewFilter). Any recipient the DDG does NOT cover - for example
         other dynamic distribution groups, which cannot be members of a DDG - is added
         explicitly to the rule's ExceptIfSentTo list and printed on screen.
         Items in (b) are a snapshot: if you later create a NEW dynamic distribution
         group in a catch-all domain, re-run the script (it is idempotent).

    This is the established catch-all pattern for Exchange Online (InternalRelay domain +
    dynamic group of existing recipients + transport rule with an exception), extended so
    the exception covers every recipient type and is verified at setup time.

.PARAMETER AdminUPN
    Admin UPN used for Connect-ExchangeOnline. Prompted if omitted.

.PARAMETER Domains
    One or more accepted domains to configure (comma-separated string or array). Prompted if omitted.

.PARAMETER CatchAllMailbox
    Mailbox that receives all mail sent to non-existent addresses. Prompted if omitted.

.PARAMETER ExceptionGroupAddress
    SMTP address for the exception dynamic distribution group. Prompted if omitted.

.EXAMPLE
    .\Set-CatchAllMailbox.ps1
    Fully interactive.

.EXAMPLE
    .\Set-CatchAllMailbox.ps1 -AdminUPN ADMIN@DOMAIN.COM -Domains domain.com,domain2.com -CatchAllMailbox catchall@domain.com -ExceptionGroupAddress catchall-exception@domain.com
    Any email sent to a non-existent address in domain.com or domain2.com is delivered to catchall@domain.com.

.EXAMPLE
    .\Set-CatchAllMailbox.ps1 -WhatIf
    Shows what would be removed, changed and created without changing anything.

.NOTES
    Version       : 3.1
    Compatibility : PowerShell 5.1+ / ExchangeOnlineManagement 3.x
    Requirement   : Exchange Administrator (or Global Administrator).

    IMPORTANT:
    - Propagation of the InternalRelay change can take up to 60 minutes.
      During that time 'Address not found' NDRs (5.1.10) are normal.
    - Dynamic group membership is recalculated periodically by Exchange Online; a brand-new
      recipient can take a short while to be recognized by the exception.
    - The catch-all mailbox itself is an existing recipient, so it is excluded (no loop).
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [string]$AdminUPN,
    [string[]]$Domains,
    [string]$CatchAllMailbox,
    [string]$ExceptionGroupAddress
)

$ErrorActionPreference = 'Stop'
$WarningPreference = 'SilentlyContinue'

# --- INTERACTIVE INPUT (only for what was not passed as a parameter) ---
if (-not $AdminUPN)              { $AdminUPN              = Read-Host "Administrator email (UPN)" }
if (-not $Domains)               { $Domains               = Read-Host "Target domains (comma-separated: dom1.com, dom2.com)" }
if (-not $CatchAllMailbox)       { $CatchAllMailbox       = Read-Host "Mailbox that will receive everything (catch-all mailbox)" }
if (-not $ExceptionGroupAddress) { $ExceptionGroupAddress = Read-Host "Email address for the exception dynamic group" }

$ruleName       = "Catch-All Global - MultiDomain"
$groupName      = "Catch-All Exception - All Recipients"
# Names used by earlier (Portuguese) versions of this script - cleaned up too, so a re-run never leaves duplicates.
$legacyRuleName = "Catch-All Global - MultiDominio"

$domainList = @($Domains | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
if ($domainList.Count -eq 0) { throw "No domain provided." }

# Every mail-enabled recipient type. DDGs cannot contain other DDGs - those are handled
# by the coverage check below (added to ExceptIfSentTo).
$exceptionFilter = "(RecipientTypeDetails -eq 'UserMailbox') -or " +
                   "(RecipientTypeDetails -eq 'SharedMailbox') -or " +
                   "(RecipientTypeDetails -eq 'RoomMailbox') -or " +
                   "(RecipientTypeDetails -eq 'EquipmentMailbox') -or " +
                   "(RecipientTypeDetails -eq 'SchedulingMailbox') -or " +
                   "(RecipientTypeDetails -eq 'TeamMailbox') -or " +
                   "(RecipientTypeDetails -eq 'GroupMailbox') -or " +
                   "(RecipientTypeDetails -eq 'MailUniversalDistributionGroup') -or " +
                   "(RecipientTypeDetails -eq 'MailUniversalSecurityGroup') -or " +
                   "(RecipientTypeDetails -eq 'MailNonUniversalGroup') -or " +
                   "(RecipientTypeDetails -eq 'RoomList') -or " +
                   "(RecipientTypeDetails -eq 'MailUser') -or " +
                   "(RecipientTypeDetails -eq 'GuestMailUser') -or " +
                   "(RecipientTypeDetails -eq 'MailContact') -or " +
                   "(RecipientTypeDetails -eq 'PublicFolder')"

Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false

try {
    # --- PHASE 1: CLEANUP (RESET) ---
    Write-Host "Removing previous configuration to avoid duplicates..." -ForegroundColor Yellow

    foreach ($rn in @($ruleName, $legacyRuleName)) {
        if (Get-TransportRule -Identity $rn -ErrorAction SilentlyContinue) {
            if ($PSCmdlet.ShouldProcess("Transport rule '$rn'", "Remove")) {
                Remove-TransportRule -Identity $rn -Confirm:$false
                Write-Host "Old rule '$rn' removed." -ForegroundColor Gray
            }
        }
    }

    if (Get-DynamicDistributionGroup -Identity $ExceptionGroupAddress -ErrorAction SilentlyContinue) {
        if ($PSCmdlet.ShouldProcess("Dynamic distribution group '$ExceptionGroupAddress'", "Remove")) {
            Remove-DynamicDistributionGroup -Identity $ExceptionGroupAddress -Confirm:$false
            Write-Host "Old group removed." -ForegroundColor Gray
        }
    }

    # --- PHASE 2: DOMAIN CONFIGURATION ---
    foreach ($dom in $domainList) {
        if ($PSCmdlet.ShouldProcess("Accepted domain '$dom'", "Set DomainType InternalRelay")) {
            Write-Host "Setting domain $dom to InternalRelay..." -ForegroundColor Cyan
            Set-AcceptedDomain -Identity $dom -DomainType InternalRelay
        }
    }

    # --- PHASE 3: EXCEPTION DYNAMIC GROUP ---
    if ($PSCmdlet.ShouldProcess("Dynamic distribution group '$ExceptionGroupAddress'", "Create (all mail-enabled recipients)")) {
        Write-Host "Creating exception group for all existing recipients..." -ForegroundColor Cyan
        New-DynamicDistributionGroup -Name $groupName `
            -PrimarySmtpAddress $ExceptionGroupAddress `
            -RecipientFilter $exceptionFilter `
            -ErrorAction Stop | Out-Null
    }

    # --- PHASE 4: COVERAGE CHECK ---
    # Any recipient with an address in a catch-all domain that the DDG does not cover is
    # added explicitly to ExceptIfSentTo, so it is never redirected.
    Write-Host "Checking that every existing recipient is covered by the exception..." -ForegroundColor Cyan
    $covered = @{}
    Get-Recipient -RecipientPreviewFilter $exceptionFilter -ResultSize Unlimited |
        ForEach-Object { $covered[[string]$_.PrimarySmtpAddress] = $true }

    $uncovered = @(Get-Recipient -ResultSize Unlimited | Where-Object {
        -not $covered.ContainsKey([string]$_.PrimarySmtpAddress) -and
        [string]$_.PrimarySmtpAddress -ne $ExceptionGroupAddress -and
        @($_.EmailAddresses | Where-Object {
            $addr = ([string]$_) -replace '^smtp:', ''
            $domainList -contains ($addr.Split('@')[-1])
        }).Count -gt 0
    })

    $explicitExceptions = @()
    if ($uncovered.Count -gt 0) {
        Write-Host "Recipients not covered by the dynamic group (added to the rule exception list):" -ForegroundColor Yellow
        $uncovered | ForEach-Object { Write-Host ("  - {0}  [{1}]" -f $_.PrimarySmtpAddress, $_.RecipientTypeDetails) -ForegroundColor Yellow }
        $explicitExceptions = @($uncovered | ForEach-Object { [string]$_.PrimarySmtpAddress })
    } else {
        Write-Host "All existing recipients in the selected domains are covered." -ForegroundColor Green
    }

    # --- PHASE 5: CATCH-ALL TRANSPORT RULE ---
    $ruleParams = @{
        Name                   = $ruleName
        RecipientDomainIs      = $domainList
        RedirectMessageTo      = $CatchAllMailbox
        ExceptIfSentToMemberOf = $ExceptionGroupAddress
        Priority               = 0
        StopRuleProcessing     = $true
    }
    if ($explicitExceptions.Count -gt 0) { $ruleParams.ExceptIfSentTo = $explicitExceptions }

    if ($PSCmdlet.ShouldProcess("Transport rule '$ruleName'", "Create (priority 0, redirect to $CatchAllMailbox)")) {
        Write-Host "Creating transport rule with priority 0..." -ForegroundColor Cyan
        New-TransportRule @ruleParams | Out-Null
    }

    Write-Host "`n===========================================================" -ForegroundColor Green
    if ($WhatIfPreference) {
        Write-Host "WHATIF: NOTHING WAS CHANGED" -ForegroundColor Yellow
    } else {
        Write-Host "CONFIGURATION COMPLETED" -ForegroundColor Green
    }
    Write-Host "Domains    : $($domainList -join ', ')" -ForegroundColor White
    Write-Host "Destination: $CatchAllMailbox" -ForegroundColor White
    Write-Host "Propagation: allow up to 60 minutes before final validation." -ForegroundColor Yellow
    Write-Host "===========================================================" -ForegroundColor Green
}
catch {
    Write-Host "EXECUTION ERROR: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "Session closed." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false
