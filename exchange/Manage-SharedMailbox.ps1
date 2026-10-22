<#
.SYNOPSIS
    Manages access to a shared mailbox in Exchange Online.

.DESCRIPTION
    Connects to EXO with the admin account, asks which shared mailbox,
    lists who has Full Access / Send As / Send on Behalf, and offers a menu to
    add (or remove) users in any of the three permissions.

    PREREQUISITE: the "Exchange Recipient Administrator" role ACTIVE (via PIM, if used).
    Without it, EXO misleadingly says the mailbox does not exist.

.PARAMETER AdminUPN
    Admin account used to connect to Exchange Online if no session is active.

.PARAMETER SharedMailbox
    Optional. Shared mailbox address; if omitted, the script prompts for it.

.EXAMPLE
    .\Manage-SharedMailbox.ps1 -AdminUPN ADMIN@DOMAIN.COM
    Prompts for the shared mailbox and opens the menu.

.EXAMPLE
    .\Manage-SharedMailbox.ps1 -AdminUPN ADMIN@DOMAIN.COM -SharedMailbox finance@domain.com
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$AdminUPN,
    [string]$SharedMailbox
)

$ErrorActionPreference = 'Stop'

#region ---------- Helpers ----------

function Write-Head {
    param([string]$Text)
    Write-Host ''
    Write-Host ('=' * 78) -ForegroundColor DarkCyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host ('=' * 78) -ForegroundColor DarkCyan
}

function Write-Ok   { param($m) Write-Host "[ OK ] $m"  -ForegroundColor Green }
function Write-Warn { param($m) Write-Host "[WARN] $m"  -ForegroundColor Yellow }
function Write-Err  { param($m) Write-Host "[FAIL] $m"  -ForegroundColor Red }
function Write-Info { param($m) Write-Host "[INFO] $m"  -ForegroundColor Gray }

function Connect-EXOIfNeeded {
    param([string]$UPN)

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "ExchangeOnlineManagement module not installed. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    $conn = Get-ConnectionInformation -ErrorAction SilentlyContinue |
            Where-Object { $_.State -eq 'Connected' }

    if ($conn) {
        Write-Ok "Already connected to EXO as $($conn[0].UserPrincipalName)"
        return
    }

    Write-Info "Connecting to Exchange Online as $UPN ..."
    Connect-ExchangeOnline -UserPrincipalName $UPN -ShowBanner:$false
    Write-Ok "Connected."
}

# Resolves anything (email, alias, name) to a recipient object
function Resolve-Recipient {
    param(
        [Parameter(Mandatory)][string]$Identity,
        [switch]$MailboxOnly
    )

    $r = $null
    try { $r = Get-Recipient -Identity $Identity -ErrorAction Stop } catch { }

    if (-not $r) {
        try {
            $r = @(Get-Recipient -ResultSize 20 -Filter "EmailAddresses -like '*$Identity*'" -ErrorAction Stop)
        } catch { }
    }

    if (-not $r) {
        try {
            $r = @(Get-Recipient -ResultSize 20 -Filter "DisplayName -like '*$Identity*'" -ErrorAction Stop)
        } catch { }
    }

    $r = @($r)
    if ($r.Count -eq 0) {
        Write-Err "Nothing found for '$Identity'. (If EVERYTHING returns not found -> your admin role / PIM activation expired.)"
        return $null
    }

    if ($r.Count -gt 1) {
        Write-Warn "More than one result for '$Identity':"
        for ($i = 0; $i -lt $r.Count; $i++) {
            "{0,3}) {1,-38} {2,-45} {3}" -f ($i + 1), $r[$i].DisplayName, $r[$i].PrimarySmtpAddress, $r[$i].RecipientTypeDetails |
                Write-Host
        }
        $pick = Read-Host "Choose the number (Enter = cancel)"
        if ([string]::IsNullOrWhiteSpace($pick)) { return $null }
        $idx = 0
        if (-not [int]::TryParse($pick, [ref]$idx) -or $idx -lt 1 -or $idx -gt $r.Count) {
            Write-Err "Invalid choice."
            return $null
        }
        $r = $r[$idx - 1]
    } else {
        $r = $r[0]
    }

    if ($MailboxOnly -and $r.RecipientTypeDetails -notlike '*Mailbox*') {
        Write-Err "'$($r.PrimarySmtpAddress)' is a $($r.RecipientTypeDetails), not a mailbox."
        return $null
    }

    return $r
}

#endregion

#region ---------- Reading permissions ----------

# The EXO V3 REST cmdlets return booleans sometimes as [bool], sometimes as the string
# "False" -- and in PowerShell any non-empty string is truthy. Normalize here.
function ConvertTo-Bool {
    param($Value)
    if ($null -eq $Value)      { return $false }
    if ($Value -is [bool])     { return $Value }
    return ([string]$Value).Trim() -eq 'True'
}

function Get-SharedMailboxAccess {
    param([Parameter(Mandatory)]$Mailbox)

    $mbx = Get-Mailbox -Identity $Mailbox.Identity

    $full = @(
        Get-MailboxPermission -Identity $mbx.Identity |
        Where-Object {
            $_.User -notlike 'NT AUTHORITY\*' -and
            $_.User -notlike 'S-1-5-21-*'      -and
            -not (ConvertTo-Bool $_.IsInherited) -and
            $_.AccessRights -contains 'FullAccess'
        } |
        Select-Object @{n='User';e={[string]$_.User}},
                      @{n='Rights';e={($_.AccessRights -join ', ')}},
                      @{n='Deny';e={ConvertTo-Bool $_.Deny}}
    )

    $sendAs = @(
        Get-RecipientPermission -Identity $mbx.Identity -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Trustee -notlike 'NT AUTHORITY\*' -and
            $_.AccessRights -contains 'SendAs'
        } |
        Select-Object @{n='User';e={[string]$_.Trustee}},
                      @{n='Deny';e={[string]$_.AccessControlType -eq 'Deny'}}
    )

    $onBehalf = @(
        foreach ($g in $mbx.GrantSendOnBehalfTo) {
            $res = Get-Recipient -Identity $g -ErrorAction SilentlyContinue
            [pscustomobject]@{
                User = if ($res) { [string]$res.PrimarySmtpAddress } else { [string]$g }
            }
        }
    )

    [pscustomobject]@{
        Mailbox      = $mbx
        FullAccess   = $full
        SendAs       = $sendAs
        SendOnBehalf = $onBehalf
    }
}

function Show-SharedMailboxAccess {
    param([Parameter(Mandatory)]$Access)

    $mbx = $Access.Mailbox

    Write-Head "MAILBOX: $($mbx.DisplayName)  <$($mbx.PrimarySmtpAddress)>   [$($mbx.RecipientTypeDetails)]"

    Write-Host ''
    Write-Host '--- FULL ACCESS (open the mailbox) ---' -ForegroundColor White
    if ($Access.FullAccess.Count -eq 0) {
        Write-Host '    (nobody)' -ForegroundColor DarkGray
    } else {
        $Access.FullAccess | ForEach-Object {
            if ($_.Deny) {
                Write-Host ("    - {0} [DENY]" -f $_.User) -ForegroundColor Red
            } else {
                Write-Host ("    - {0}" -f $_.User) -ForegroundColor Gray
            }
        }
    }

    Write-Host ''
    Write-Host '--- SEND AS (sends AS the mailbox) ---' -ForegroundColor White
    if ($Access.SendAs.Count -eq 0) {
        Write-Host '    (nobody)' -ForegroundColor DarkGray
    } else {
        $Access.SendAs | ForEach-Object {
            if ($_.Deny) {
                Write-Host ("    - {0} [DENY]" -f $_.User) -ForegroundColor Red
            } else {
                Write-Host ("    - {0}" -f $_.User) -ForegroundColor Gray
            }
        }
    }

    Write-Host ''
    Write-Host '--- SEND ON BEHALF (sends ON BEHALF OF) ---' -ForegroundColor White
    if ($Access.SendOnBehalf.Count -eq 0) {
        Write-Host '    (nobody)' -ForegroundColor DarkGray
    } else {
        $Access.SendOnBehalf | ForEach-Object {
            Write-Host ("    - {0}" -f $_.User) -ForegroundColor Gray
        }
    }
    Write-Host ''
}

#endregion

#region ---------- Writing permissions ----------

function Add-FullAccess {
    param($Mailbox, $User, [bool]$AutoMapping = $true)
    try {
        Add-MailboxPermission -Identity $Mailbox.Identity `
                              -User $User.Identity `
                              -AccessRights FullAccess `
                              -InheritanceType All `
                              -AutoMapping:$AutoMapping `
                              -Confirm:$false -ErrorAction Stop | Out-Null
        Write-Ok "Full Access granted to $($User.PrimarySmtpAddress) (AutoMapping: $AutoMapping)"
    } catch {
        Write-Err "Full Access failed: $($_.Exception.Message)"
    }
}

function Add-SendAs {
    param($Mailbox, $User)
    try {
        Add-RecipientPermission -Identity $Mailbox.Identity `
                                -Trustee $User.Identity `
                                -AccessRights SendAs `
                                -Confirm:$false -ErrorAction Stop | Out-Null
        Write-Ok "Send As granted to $($User.PrimarySmtpAddress)"
    } catch {
        Write-Err "Send As failed: $($_.Exception.Message)"
    }
}

function Add-SendOnBehalf {
    param($Mailbox, $User)
    try {
        Set-Mailbox -Identity $Mailbox.Identity `
                    -GrantSendOnBehalfTo @{Add = $User.Identity} `
                    -Confirm:$false -ErrorAction Stop
        Write-Ok "Send on Behalf granted to $($User.PrimarySmtpAddress)"
    } catch {
        Write-Err "Send on Behalf failed: $($_.Exception.Message)"
    }
}

function Remove-AnyAccess {
    param($Mailbox, $User, [ValidateSet('Full','SendAs','OnBehalf')]$Kind)
    try {
        switch ($Kind) {
            'Full' {
                Remove-MailboxPermission -Identity $Mailbox.Identity -User $User.Identity `
                    -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop | Out-Null
                Write-Ok "Full Access removed from $($User.PrimarySmtpAddress)"
            }
            'SendAs' {
                Remove-RecipientPermission -Identity $Mailbox.Identity -Trustee $User.Identity `
                    -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                Write-Ok "Send As removed from $($User.PrimarySmtpAddress)"
            }
            'OnBehalf' {
                Set-Mailbox -Identity $Mailbox.Identity `
                    -GrantSendOnBehalfTo @{Remove = $User.Identity} -Confirm:$false -ErrorAction Stop
                Write-Ok "Send on Behalf removed from $($User.PrimarySmtpAddress)"
            }
        }
    } catch {
        Write-Err "Removal failed: $($_.Exception.Message)"
    }
}

# Asks about the extra rights after adding Full Access
function Invoke-SendRightsPrompt {
    param($Mailbox, $User)

    Write-Host ''
    Write-Host "Do you want $($User.PrimarySmtpAddress) to also be able to SEND?" -ForegroundColor Yellow
    Write-Host '   1) Send As        - the email goes out AS the shared mailbox'
    Write-Host '   2) Send on Behalf - goes out as "John Doe on behalf of Mailbox"'
    Write-Host '   3) Both          - NOT recommended: when the user has both,'
    Write-Host '                      Send As wins and on-behalf is never used'
    Write-Host '   4) None (Full Access only)'
    $opt = Read-Host 'Option [4]'

    switch ($opt) {
        '1' { Add-SendAs        -Mailbox $Mailbox -User $User }
        '2' { Add-SendOnBehalf  -Mailbox $Mailbox -User $User }
        '3' { Add-SendAs        -Mailbox $Mailbox -User $User
              Add-SendOnBehalf  -Mailbox $Mailbox -User $User }
        default { Write-Info 'No send permission added.' }
    }
}

#endregion

#region ---------- Main ----------

Write-Head 'SHARED MAILBOX MANAGER - EXCHANGE ONLINE'
Write-Warn 'Requires the "Exchange Recipient Administrator" role ACTIVE. Without it, EXO says the mailbox does not exist.'

Connect-EXOIfNeeded -UPN $AdminUPN

# ---- Which mailbox? (without assuming any domain) ----
$mbxObj = $null
while (-not $mbxObj) {
    if ([string]::IsNullOrWhiteSpace($SharedMailbox)) {
        $SharedMailbox = Read-Host "`nEnter the FULL email address of the shared mailbox (empty Enter = exit)"
    }
    if ([string]::IsNullOrWhiteSpace($SharedMailbox)) {
        Write-Info 'Exiting.'
        return
    }
    $mbxObj = Resolve-Recipient -Identity $SharedMailbox.Trim() -MailboxOnly
    $SharedMailbox = $null
}

$access = Get-SharedMailboxAccess -Mailbox $mbxObj
Show-SharedMailboxAccess -Access $access

# ---- Menu ----
do {
    Write-Host ''
    Write-Host '--------------------- WHAT DO YOU WANT TO DO? ---------------------' -ForegroundColor Cyan
    Write-Host '   1) Add user with FULL ACCESS (then I ask about sending)'
    Write-Host '   2) Add Send As ONLY'
    Write-Host '   3) Add Send on Behalf ONLY'
    Write-Host '   4) Remove a permission from a user'
    Write-Host '   5) Show the access list again'
    Write-Host '   6) Switch shared mailbox'
    Write-Host '   0) Exit'
    $choice = Read-Host 'Option'

    switch ($choice) {

        '1' {
            do {
                $u = Read-Host "`nEmail of the user to add (Enter = stop)"
                if ([string]::IsNullOrWhiteSpace($u)) { break }

                $usr = Resolve-Recipient -Identity $u.Trim()
                if (-not $usr) { continue }

                $am = Read-Host "AutoMapping (mailbox appears automatically in Outlook)? [Y/n]"
                $autoMap = -not ($am -match '^[nN]')

                Add-FullAccess -Mailbox $mbxObj -User $usr -AutoMapping $autoMap
                Invoke-SendRightsPrompt -Mailbox $mbxObj -User $usr

                $more = Read-Host "`nAdd ANOTHER user? [y/N]"
            } while ($more -match '^[sSyY]')

            $access = Get-SharedMailboxAccess -Mailbox $mbxObj
            Show-SharedMailboxAccess -Access $access
        }

        '2' {
            do {
                $u = Read-Host "`nEmail of the user for Send As (Enter = stop)"
                if ([string]::IsNullOrWhiteSpace($u)) { break }
                $usr = Resolve-Recipient -Identity $u.Trim()
                if ($usr) { Add-SendAs -Mailbox $mbxObj -User $usr }
                $more = Read-Host "Add another? [y/N]"
            } while ($more -match '^[sSyY]')

            $access = Get-SharedMailboxAccess -Mailbox $mbxObj
            Show-SharedMailboxAccess -Access $access
        }

        '3' {
            do {
                $u = Read-Host "`nEmail of the user for Send on Behalf (Enter = stop)"
                if ([string]::IsNullOrWhiteSpace($u)) { break }
                $usr = Resolve-Recipient -Identity $u.Trim()
                if ($usr) { Add-SendOnBehalf -Mailbox $mbxObj -User $usr }
                $more = Read-Host "Add another? [y/N]"
            } while ($more -match '^[sSyY]')

            $access = Get-SharedMailboxAccess -Mailbox $mbxObj
            Show-SharedMailboxAccess -Access $access
        }

        '4' {
            $u = Read-Host "`nEmail of the user to REMOVE (Enter = cancel)"
            if (-not [string]::IsNullOrWhiteSpace($u)) {
                $usr = Resolve-Recipient -Identity $u.Trim()
                if ($usr) {
                    Write-Host '   1) Full Access'
                    Write-Host '   2) Send As'
                    Write-Host '   3) Send on Behalf'
                    Write-Host '   4) ALL'
                    $k = Read-Host 'Remove which?'
                    switch ($k) {
                        '1' { Remove-AnyAccess -Mailbox $mbxObj -User $usr -Kind Full }
                        '2' { Remove-AnyAccess -Mailbox $mbxObj -User $usr -Kind SendAs }
                        '3' { Remove-AnyAccess -Mailbox $mbxObj -User $usr -Kind OnBehalf }
                        '4' { Remove-AnyAccess -Mailbox $mbxObj -User $usr -Kind Full
                              Remove-AnyAccess -Mailbox $mbxObj -User $usr -Kind SendAs
                              Remove-AnyAccess -Mailbox $mbxObj -User $usr -Kind OnBehalf }
                        default { Write-Info 'Cancelled.' }
                    }
                    $access = Get-SharedMailboxAccess -Mailbox $mbxObj
                    Show-SharedMailboxAccess -Access $access
                }
            }
        }

        '5' {
            $access = Get-SharedMailboxAccess -Mailbox $mbxObj
            Show-SharedMailboxAccess -Access $access
        }

        '6' {
            $newMailbox = Read-Host "`nEmail of the new shared mailbox"
            if (-not [string]::IsNullOrWhiteSpace($newMailbox)) {
                $tmp = Resolve-Recipient -Identity $newMailbox.Trim() -MailboxOnly
                if ($tmp) {
                    $mbxObj = $tmp
                    $access = Get-SharedMailboxAccess -Mailbox $mbxObj
                    Show-SharedMailboxAccess -Access $access
                }
            }
        }

        '0' { Write-Info 'Exiting. The EXO session stays open (Disconnect-ExchangeOnline if you want to close it).' }

        default { Write-Warn 'Invalid option.' }
    }

} while ($choice -ne '0')

#endregion
