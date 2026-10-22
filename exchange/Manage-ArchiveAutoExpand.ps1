<#
.SYNOPSIS
    Interactive parent script: audits and (optionally) enables Auto-Expanding
    Archive on Exchange Online mailboxes.

.DESCRIPTION
    Runs with no arguments at all. It asks for the email, checks everything and decides:

      [A] ALREADY ENABLED        -> reports and exits. Nothing to do.
      [B] NO LICENSE             -> warns and exits. Enabling would not help.
      [C] DISABLED + LICENSED    -> offers to enable (Y/N).
                                    If YES: runs it, waits 30s with a countdown,
                                    re-queries and shows the proof.
      [D] NO ARCHIVE             -> warns and gives you the command to provision it.

    When done, asks whether to process another mailbox. Loops until you exit.

    The license is inferred from ArchiveQuota (>= 90 GB = E3/E5 or EOA add-on),
    so the script works with an Exchange Online session only (no Microsoft Graph
    connection or license-read permission needed).

.PARAMETER User
    Optional. Skips the first prompt and processes this email right away.

.PARAMETER Admin
    Admin UPN used for Connect-ExchangeOnline.

.PARAMETER Force
    Skips the Y/N confirmation. Use with care.

.EXAMPLE
    .\Manage-ArchiveAutoExpand.ps1 -Admin ADMIN@DOMAIN.COM

.EXAMPLE
    .\Manage-ArchiveAutoExpand.ps1 -Admin ADMIN@DOMAIN.COM -User user@domain.com

.NOTES
    ROLE REQUIRED: "Exchange Recipient Administrator" (or Exchange Administrator).
    If you use PIM: Entra -> Identity Governance -> PIM -> My roles ->
    Microsoft Entra roles -> Activate.
    Without an active role Get-Mailbox misleads: it returns "object couldn't be found"
    instead of "access denied".

    Enable-Mailbox -AutoExpandingArchive is IRREVERSIBLE.
#>

[CmdletBinding()]
param(
    [string] $User,
    [Parameter(Mandatory = $true)][string] $Admin,
    [switch] $Force
)

$ErrorActionPreference = 'Stop'

# ================================================================ helpers ===
function Write-Step  { param($m) Write-Host "[*] $m" -ForegroundColor Cyan }
function Write-Ok    { param($m) Write-Host "[+] $m" -ForegroundColor Green }
function Write-Warn2 { param($m) Write-Host "[!] $m" -ForegroundColor Yellow }
function Write-Err2  { param($m) Write-Host "[X] $m" -ForegroundColor Red }

function Write-Banner {
    param($Text, $Color = 'Magenta')
    $line = '=' * 62
    Write-Host ''
    Write-Host $line -ForegroundColor $Color
    Write-Host ("  " + $Text) -ForegroundColor $Color
    Write-Host $line -ForegroundColor $Color
}

function ConvertTo-Gb {
    param($Size)
    if ($null -eq $Size) { return $null }
    $s = $Size.ToString()
    if ($s -match '\(([\d,\.]+)\s*bytes\)') {
        return [math]::Round(([double]($Matches[1] -replace '[,\.]','')) / 1GB, 2)
    }
    if ($s -match 'Unlimited') { return 'Unlimited' }
    return $s
}

function Start-Countdown {
    param([int] $Seconds = 30, [string] $Activity = 'Waiting for EXO propagation')
    for ($i = $Seconds; $i -gt 0; $i--) {
        $pct = (($Seconds - $i) / $Seconds) * 100
        Write-Progress -Activity $Activity -Status "$i seconds remaining..." -PercentComplete $pct
        Start-Sleep -Seconds 1
    }
    Write-Progress -Activity $Activity -Completed
}

function Read-YesNo {
    param([string] $Prompt)
    while ($true) {
        $a = (Read-Host "$Prompt [Y/N]").Trim().ToUpper()
        if ($a -in @('Y','YES','S','SIM')) { return $true  }   # S/SIM/NAO kept for backward compatibility
        if ($a -in @('N','NO','NAO')) { return $false }
        Write-Warn2 "Please answer Y or N."
    }
}

function Resolve-TargetMailbox {
    param([string] $Id)
    try {
        return Get-Mailbox -Identity $Id -ErrorAction Stop
    } catch {
        Write-Warn2 "Get-Mailbox failed for '$Id'. Trying Get-Recipient..."
    }
    $local = ($Id -split '@')[0]
    $rcpt  = @(Get-Recipient -Filter "EmailAddresses -like '*$local*'" -ErrorAction SilentlyContinue)

    if ($rcpt.Count -eq 0) {
        throw "No recipient found for '$Id'. Check your admin role (PIM) and the address."
    }
    if ($rcpt.Count -gt 1) {
        Write-Warn2 "Multiple recipients matched '$local':"
        $rcpt | Select-Object Name, PrimarySmtpAddress, RecipientType | Format-Table -AutoSize | Out-Host
        throw "Ambiguous. Run again with the exact PrimarySmtpAddress."
    }
    Write-Ok "Resolved via Get-Recipient -> Name = $($rcpt[0].Name)"
    return Get-Mailbox -Identity $rcpt[0].Name -ErrorAction Stop
}

# ============================================================ connection ====
function Connect-IfNeeded {
    try   { $conn = Get-ConnectionInformation -ErrorAction SilentlyContinue }
    catch { $conn = $null }

    if (-not $conn -or $conn.State -ne 'Connected') {
        Write-Step "No active EXO session. Connecting as $Admin ..."
        Write-Warn2 "The 'Exchange Recipient Administrator' role must be ACTIVE now (PIM, if used)."
        Connect-ExchangeOnline -UserPrincipalName $Admin -ShowBanner:$false
        Write-Ok "Connected."
    } else {
        Write-Ok "EXO session already active as $($conn.UserPrincipalName)."
    }
}

# ======================================================== per-user core ====
function Invoke-ArchiveCheck {
    param([string] $Target)

    Write-Banner "AUDITING: $Target"

    try {
        $mbx = Resolve-TargetMailbox -Id $Target
    } catch {
        Write-Err2 $_.Exception.Message
        return
    }

    # ArchiveStatus LIES for cloud mailboxes (on-prem legacy).
    # The truth is ArchiveState + ArchiveGuid.
    $hasArchive = ($mbx.ArchiveGuid -ne [guid]::Empty) -and ($mbx.ArchiveState -ne 'None')
    $quotaGb    = ConvertTo-Gb $mbx.ArchiveQuota
    $warnGb     = ConvertTo-Gb $mbx.ArchiveWarningQuota

    $usedGb = $null; $pct = $null; $items = $null
    if ($hasArchive) {
        $st = Get-MailboxStatistics -Identity $mbx.Identity -Archive -ErrorAction SilentlyContinue
        if ($st) {
            $usedGb = ConvertTo-Gb $st.TotalItemSize
            $items  = $st.ItemCount
            if (($usedGb -is [double]) -and ($quotaGb -is [double]) -and $quotaGb -gt 0) {
                $pct = [math]::Round(($usedGb / $quotaGb) * 100, 1)
            }
        }
    }

    $primary = Get-MailboxStatistics -Identity $mbx.Identity -ErrorAction SilentlyContinue

    $licensed = ($quotaGb -eq 'Unlimited') -or (($quotaGb -is [double]) -and ($quotaGb -ge 90))

    # ----------------------------------------------------------- panel ----
    Write-Host ''
    [pscustomobject]@{
        DisplayName                 = $mbx.DisplayName
        PrimarySmtpAddress          = $mbx.PrimarySmtpAddress
        RecipientTypeDetails        = $mbx.RecipientTypeDetails
        ArchiveState                = $mbx.ArchiveState
        HasArchive                  = $hasArchive
        AutoExpandingArchiveEnabled = $mbx.AutoExpandingArchiveEnabled
        ArchiveQuotaGB              = $quotaGb
        ArchiveWarningQuotaGB       = $warnGb
        ArchiveUsedGB               = $usedGb
        ArchivePercentUsed          = $pct
        ArchiveItemCount            = $items
        PrimarySizeGB               = (ConvertTo-Gb $primary.TotalItemSize)
        RetentionPolicy             = $mbx.RetentionPolicy
        LitigationHoldEnabled       = $mbx.LitigationHoldEnabled
        InPlaceHolds                = ($mbx.InPlaceHolds -join ', ')
    } | Format-List

    # ------------------------------------------------------ scenario [A] --
    if ($mbx.AutoExpandingArchiveEnabled -eq $true) {
        Write-Banner "ALREADY ENABLED" 'Green'
        Write-Ok "Auto-Expanding Archive = True. Nothing to do."
        Write-Host ''
        Write-Host " The archive grows automatically up to 1.5 TB."
        if ($pct) { Write-Host " Current usage: $usedGb GB ($pct% of the $quotaGb GB base quota)." }
        Write-Host ''
        Write-Host " If the user still complains about a full archive:"
        Write-Host "   - provisioning of the auxiliary archive takes up to 30 days;"
        Write-Host "   - or the retention policy is not moving any items"
        Write-Host "     (current RetentionPolicy: $($mbx.RetentionPolicy))."
        return
    }

    # ------------------------------------------------------ scenario [D] --
    if (-not $hasArchive) {
        Write-Banner "NO ARCHIVE PROVISIONED" 'Yellow'
        Write-Warn2 "ArchiveState = $($mbx.ArchiveState). This mailbox has no archive."
        Write-Host ''
        Write-Host " Auto-expand is not even on the table. Provision the archive first:"
        Write-Host "   Enable-Mailbox -Identity '$($mbx.PrimarySmtpAddress)' -Archive" -ForegroundColor Yellow
        Write-Host ''
        Write-Host " Then run this script again."
        return
    }

    # ------------------------------------------------------ scenario [B] --
    if (-not $licensed) {
        Write-Banner "NO ARCHIVING LICENSE" 'Red'
        Write-Err2 "ArchiveQuota = $quotaGb GB."
        Write-Host ''
        Write-Host " A ~50 GB quota indicates NO archiving license."
        Write-Host " Auto-Expanding Archive requires E3, E5 or the"
        Write-Host " Exchange Online Archiving (EOA) add-on."
        Write-Host ''
        Write-Warn2 "Enabling it here would NOT add any space. Fix the license first."
        Write-Host ''
        Write-Host " Next step: check the user's SKU in the portal"
        Write-Host " admin.microsoft.com -> Users -> Licenses and apps."
        return
    }

    # ------------------------------------------------------ scenario [C] --
    Write-Banner "DISABLED - BUT LICENSED. IT CAN BE ENABLED." 'Yellow'
    Write-Ok "License confirmed (ArchiveQuota = $quotaGb GB)."
    Write-Warn2 "AutoExpandingArchiveEnabled = False"
    if ($pct) {
        Write-Host ''
        Write-Host " Current usage: $usedGb GB of $quotaGb GB  ($pct%)"
        if ($warnGb -is [double] -and $usedGb -is [double] -and $usedGb -ge $warnGb) {
            Write-Warn2 " Already past the ArchiveWarningQuota of $warnGb GB - that is why the alert."
        }
    }

    if ($mbx.LitigationHoldEnabled -eq $true -or $mbx.InPlaceHolds) {
        Write-Host ''
        Write-Warn2 "WARNING: MAILBOX ON HOLD."
        Write-Host "   LitigationHold : $($mbx.LitigationHoldEnabled)"
        Write-Host "   InPlaceHolds   : $($mbx.InPlaceHolds -join ', ')"
        Write-Host "   Auxiliary archives have limitations in eDiscovery/Content Search."
        Write-Host "   Check with your security/legal team before proceeding." -ForegroundColor Yellow
    }

    Write-Host ''
    Write-Host ' THIS ACTION IS IRREVERSIBLE:' -ForegroundColor Red
    Write-Host '   - cannot be disabled later, not even with a Microsoft ticket' -ForegroundColor Red
    Write-Host '   - blocks moving the mailbox back on-prem and complicates cross-tenant moves' -ForegroundColor Red
    Write-Host '   - eDiscovery/Content Search limitations on auxiliary archives' -ForegroundColor Red
    Write-Host '   - client must be Outlook 2016+ or OWA' -ForegroundColor Red
    Write-Host ''

    if ($Force) {
        Write-Warn2 "-Force active: skipping confirmation."
    } elseif (-not (Read-YesNo "Enable Auto-Expanding Archive for $($mbx.PrimarySmtpAddress) NOW?")) {
        Write-Host ''
        Write-Warn2 "ABORTED. Nothing was changed."
        Write-Host " When you decide, the command is:"
        Write-Host "   Enable-Mailbox -Identity '$($mbx.PrimarySmtpAddress)' -AutoExpandingArchive" -ForegroundColor Yellow
        return
    }

    # ----------------------------------------------------------- execute --
    Write-Host ''
    Write-Step "Running Enable-Mailbox -AutoExpandingArchive ..."
    try {
        Enable-Mailbox -Identity $mbx.Identity -AutoExpandingArchive -ErrorAction Stop | Out-Null
        Write-Ok "Cmdlet ran without errors."
    } catch {
        Write-Err2 "FAILED: $($_.Exception.Message)"
        Write-Host ''
        Write-Host " Suspect 1: role/PIM activation expired. Reactivate and try again."
        Write-Host " Suspect 2: the operation may require the full Exchange Administrator role"
        Write-Host "            rather than Recipient Administrator. In that case, escalate."
        return
    }

    Write-Host ''
    Write-Step "Waiting 30s for EXO propagation before verifying..."
    Start-Countdown -Seconds 30

    # ------------------------------------------------------------ verify --
    Write-Step "Verifying result..."
    Write-Host ''
    Write-Host "Get-Mailbox `"$($mbx.PrimarySmtpAddress)`" | FL AutoExpandingArchiveEnabled,ArchiveState,ArchiveQuota" -ForegroundColor DarkGray
    Write-Host ''

    $after = Get-Mailbox -Identity $mbx.Identity -ErrorAction Stop
    $after | Format-List AutoExpandingArchiveEnabled, ArchiveState, ArchiveQuota

    if ($after.AutoExpandingArchiveEnabled -eq $true) {
        Write-Banner "CONFIRMED: AutoExpandingArchiveEnabled = True" 'Green'
        Write-Host ''
        Write-Host " The archive now grows automatically up to 1.5 TB."
        Write-Host ''
        Write-Warn2 "BUT: provisioning of the auxiliary archive is ASYNCHRONOUS and takes UP TO 30 DAYS."
        Write-Host " Until then the real headroom is whatever is left of the $quotaGb GB."
        if ($usedGb -is [double] -and $quotaGb -is [double]) {
            Write-Host " Remaining today: approx. $([math]::Round($quotaGb - $usedGb, 2)) GB."
        }
        Write-Host " If it hits the ceiling before that, the Managed Folder Assistant stops"
        Write-Host " moving items and the user gets errors."
        Write-Host ''
        Write-Host " Re-check in 7 days by running this script again." -ForegroundColor Yellow
    } else {
        Write-Banner "STILL FALSE - propagation pending" 'Yellow'
        Write-Warn2 "The cmdlet did not error, so it is most likely just EXO replication."
        Write-Host " Re-query in ~15 minutes:"
        Write-Host "   Get-Mailbox '$($mbx.PrimarySmtpAddress)' | FL AutoExpandingArchiveEnabled" -ForegroundColor Yellow
        Write-Host ''
        Write-Host " If it is still False after 1h, it was a silent permission issue: escalate."
    }
}

# ================================================================= main =====
Clear-Host
Write-Banner 'AUTO-EXPANDING ARCHIVE - AUDIT AND ENABLEMENT'
Write-Host '  Exchange Online' -ForegroundColor DarkGray
Write-Host ''
Write-Warn2 'Make sure the role "Exchange Recipient Administrator" is ACTIVE.'
Write-Host '  If you use PIM: Entra -> Identity Governance -> PIM -> My roles -> Entra roles -> Activate' -ForegroundColor DarkGray

Connect-IfNeeded

$target = $User
while ($true) {

    if (-not $target) {
        Write-Host ''
        $target = (Read-Host 'Mailbox email address (blank ENTER to exit)').Trim()
        if ([string]::IsNullOrWhiteSpace($target)) { break }
    }

    try {
        Invoke-ArchiveCheck -Target $target
    } catch {
        Write-Err2 "Unexpected error: $($_.Exception.Message)"
    }

    $target = $null
    Write-Host ''
    if (-not (Read-YesNo 'Check another mailbox?')) { break }
}

Write-Host ''
Write-Ok 'Done.'
Write-Host 'EXO session left OPEN. To close it:' -ForegroundColor DarkGray
Write-Host '  Disconnect-ExchangeOnline -Confirm:$false' -ForegroundColor DarkGray
