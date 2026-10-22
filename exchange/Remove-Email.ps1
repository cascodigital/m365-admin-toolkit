<#
.SYNOPSIS
    Removes (soft-deletes) a specific email from all mailboxes in Microsoft 365 using a
    Purview compliance search and purge action.

.DESCRIPTION
    Interactive script for incident response (phishing, mail sent by mistake):
      1. Checks that ExchangeOnlineManagement 3.x is installed (it does NOT install it).
      2. Connects to Security & Compliance PowerShell (Connect-IPPSSession
         -EnableSearchOnlySession), reusing an existing session if there is one.
      3. Creates and runs a compliance search over all Exchange mailboxes for
         sender + subject.
      4. Shows the hit count, asks for confirmation and starts a SoftDelete purge.
      5. Polls the purge action until it completes and reports the result.

    PURGE LIMIT: a single purge action removes at most 10 items PER MAILBOX
    (Microsoft limit for New-ComplianceSearchAction -Purge). If a mailbox holds more
    than 10 matching items, run the script again (with the same sender/subject) until
    the search returns 0 items.

    SoftDelete moves the items to the Recoverable Items folder, where they stay for the
    deleted item retention period (14 days by default, up to 30) and can be recovered.

.PARAMETER Sender
    Sender address to match (from:). Prompted if omitted.

.PARAMETER Subject
    Subject to match (subject:). Prompted if omitted.

.PARAMETER PurgeTimeoutMinutes
    How long to wait for the purge action to finish before giving up polling. Default 30.

.EXAMPLE
    .\Remove-Email.ps1
    Prompts for the sender and subject.

.EXAMPLE
    .\Remove-Email.ps1 -Sender phishing@external.com -Subject "Invoice overdue"

.NOTES
    Requires: ExchangeOnlineManagement 3.x with Connect-IPPSSession -EnableSearchOnlySession
              (3.9.0 or later). Install/update it yourself, for example:
                  Install-Module ExchangeOnlineManagement -Scope CurrentUser
                  Update-Module ExchangeOnlineManagement
    Role    : Purview "Search And Purge" (e.g. Organization Management or a custom
              eDiscovery role group with Compliance Search + Search And Purge).
    Version : 3.2
#>

[CmdletBinding()]
param(
    [string]$Sender,
    [string]$Subject,
    [int]$PurgeTimeoutMinutes = 30
)

function Write-ColorMessage {
    param(
        [string]$Message,
        [string]$Type = "Info"
    )

    switch ($Type) {
        "Success"  { Write-Host "[OK]   $Message" -ForegroundColor Green }
        "Error"    { Write-Host "[ERR]  $Message" -ForegroundColor Red }
        "Warning"  { Write-Host "[WARN] $Message" -ForegroundColor Yellow }
        "Info"     { Write-Host "[INFO] $Message" -ForegroundColor Cyan }
        "Progress" { Write-Host "  ->   $Message" -ForegroundColor White }
        default    { Write-Host $Message }
    }
}

function Test-RequiredModule {
    Write-ColorMessage "=== DEPENDENCY CHECK ===" -Type "Info"
    Write-Host ""

    $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if (-not $module -or $module.Version -lt [version]'3.0.0') {
        if ($module) {
            Write-ColorMessage "ExchangeOnlineManagement $($module.Version) found - version 3.x or later is required." -Type "Error"
        } else {
            Write-ColorMessage "ExchangeOnlineManagement module not found." -Type "Error"
        }
        Write-Host ""
        Write-ColorMessage "Install or update it, then run this script again:" -Type "Info"
        Write-Host "  Install-Module ExchangeOnlineManagement -Scope CurrentUser" -ForegroundColor Gray
        Write-Host "  Update-Module ExchangeOnlineManagement" -ForegroundColor Gray
        exit 1
    }

    Write-ColorMessage "Installed version: $($module.Version)" -Type "Progress"
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Write-Host ""
}

function Connect-M365Services {
    Write-ColorMessage "=== CONNECTING ===" -Type "Info"
    Write-Host ""

    # Does Connect-IPPSSession support -EnableSearchOnlySession?
    $supportsSearchOnly = $false
    try {
        $params = (Get-Command Connect-IPPSSession -ErrorAction Stop).Parameters
        $supportsSearchOnly = $params.ContainsKey('EnableSearchOnlySession')
    } catch {
        Write-ColorMessage "Connect-IPPSSession is not available in this session." -Type "Error"
        Write-ColorMessage "If you just installed/updated the module, open a new PowerShell window." -Type "Warning"
        exit 1
    }

    if (-not $supportsSearchOnly) {
        Write-ColorMessage "Connect-IPPSSession does not support -EnableSearchOnlySession." -Type "Error"
        Write-ColorMessage "Update ExchangeOnlineManagement to 3.9.0 or later:" -Type "Info"
        Write-Host "  Update-Module ExchangeOnlineManagement" -ForegroundColor Gray
        Write-ColorMessage "Then close ALL PowerShell windows, open a new one and run the script again." -Type "Info"
        exit 1
    }

    Write-ColorMessage "-EnableSearchOnlySession support: OK" -Type "Success"
    Write-Host ""

    Write-ColorMessage "Connecting to Security & Compliance..." -Type "Progress"
    try {
        Get-ComplianceSearch -ResultSize 1 -ErrorAction Stop | Out-Null
        Write-ColorMessage "Already connected." -Type "Success"
    } catch {
        try {
            Connect-IPPSSession -EnableSearchOnlySession -ShowBanner:$false -ErrorAction Stop
            Write-ColorMessage "Connected." -Type "Success"
        } catch {
            Write-ColorMessage "Connection failed: $($_.Exception.Message)" -Type "Error"
            exit 1
        }
    }

    Write-Host ""
}

function Wait-PurgeAction {
    param(
        [Parameter(Mandatory)][string]$ActionName,
        [int]$TimeoutMinutes = 30
    )

    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $action = $null
    do {
        Start-Sleep -Seconds 15
        $action = Get-ComplianceSearchAction -Identity $ActionName -ErrorAction SilentlyContinue
        $status = if ($action) { $action.Status } else { '(not visible yet)' }
        Write-Host "    Purge status: $status" -ForegroundColor Gray
        if ($action -and $action.Status -in @('Completed', 'PartiallySucceeded', 'Failed')) { break }
    } while ((Get-Date) -lt $deadline)

    return $action
}

function Remove-EmailBySearch {
    Write-ColorMessage "=== EMAIL REMOVAL ===" -Type "Info"
    Write-Host ""

    $sender = $script:Sender
    if ([string]::IsNullOrWhiteSpace($sender)) { $sender = Read-Host "Sender address" }
    if ([string]::IsNullOrWhiteSpace($sender)) {
        Write-ColorMessage "Sender cannot be empty." -Type "Error"
        exit 1
    }

    $subject = $script:Subject
    if ([string]::IsNullOrWhiteSpace($subject)) { $subject = Read-Host "Email subject" }
    if ([string]::IsNullOrWhiteSpace($subject)) {
        Write-ColorMessage "Subject cannot be empty." -Type "Error"
        exit 1
    }

    # Unique search name
    $timestamp   = Get-Date -Format "yyyyMMdd_HHmmss"
    $searchName  = "EmailRemoval_$timestamp"
    $searchQuery = "(from:$sender) AND (subject:'$subject')"

    Write-Host ""
    Write-ColorMessage "OPERATION SUMMARY:" -Type "Info"
    Write-Host "  Sender : $sender"
    Write-Host "  Subject: $subject"
    Write-Host "  Query  : $searchQuery"
    Write-Host "  Search : $searchName"
    Write-Host ""

    $confirm = Read-Host "Run the search? (Y/N)"
    if ($confirm -notmatch '^[YySs]') {
        Write-ColorMessage "Operation cancelled." -Type "Warning"
        exit
    }

    Write-Host ""
    Write-ColorMessage "=== PROCESSING ===" -Type "Info"

    try {
        # 1. Create the search
        Write-ColorMessage "1/6 Creating compliance search..." -Type "Progress"
        New-ComplianceSearch -Name $searchName -ExchangeLocation All -ContentMatchQuery $searchQuery -ErrorAction Stop | Out-Null
        Write-ColorMessage "Search created." -Type "Success"

        # 2. Start the search
        Write-ColorMessage "2/6 Starting search..." -Type "Progress"
        Start-ComplianceSearch -Identity $searchName -ErrorAction Stop
        Write-ColorMessage "Search started." -Type "Success"

        # 3. Wait for completion
        Write-ColorMessage "3/6 Waiting for the search to complete..." -Type "Progress"
        $attempts = 0
        do {
            Start-Sleep -Seconds 10
            $searchStatus = Get-ComplianceSearch -Identity $searchName
            $attempts++
            Write-Host "    Status: $($searchStatus.Status) (attempt $attempts)" -ForegroundColor Gray

            if ($searchStatus.Status -eq 'Failed') {
                throw "Search failed."
            }
        } while ($searchStatus.Status -ne 'Completed' -and $attempts -lt 60)

        if ($searchStatus.Status -ne 'Completed') {
            throw "Timeout: the search did not complete within 10 minutes."
        }

        Write-ColorMessage "Search completed." -Type "Success"

        # 4. Check the results
        $itemCount = $searchStatus.Items
        Write-ColorMessage "4/6 Items found: $itemCount" -Type "Info"

        if ($itemCount -eq 0) {
            Write-ColorMessage "No email matched the criteria." -Type "Warning"
            Remove-ComplianceSearch -Identity $searchName -Confirm:$false
            exit
        }

        Write-Host ""
        Write-ColorMessage "Note: one purge removes at most 10 items per mailbox. Re-run if a mailbox has more." -Type "Warning"
        $confirmPurge = Read-Host "Soft-delete these $itemCount item(s) now? (Y/N)"
        if ($confirmPurge -notmatch '^[YySs]') {
            Write-ColorMessage "Purge cancelled. The search '$searchName' was kept for review in Purview." -Type "Warning"
            exit
        }

        # 5. Start the purge
        Write-ColorMessage "5/6 Starting purge (SoftDelete)..." -Type "Progress"
        $action = New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType SoftDelete -Confirm:$false -ErrorAction Stop
        $actionName = if ($action -and $action.Name) { $action.Name } else { "${searchName}_Purge" }

        # 6. Poll the purge until it finishes
        Write-ColorMessage "6/6 Waiting for the purge action '$actionName' to finish..." -Type "Progress"
        $final = Wait-PurgeAction -ActionName $actionName -TimeoutMinutes $script:PurgeTimeoutMinutes

        Write-Host ""
        if (-not $final) {
            Write-ColorMessage "Could not read the purge action status. Check it with:" -Type "Warning"
            Write-Host "  Get-ComplianceSearchAction -Identity '$actionName' | Format-List Status,Results,Errors" -ForegroundColor Gray
        } elseif ($final.Status -eq 'Completed') {
            Write-ColorMessage "=== PURGE COMPLETED ===" -Type "Success"
            Write-ColorMessage "Search hits: $itemCount item(s)." -Type "Info"
            if ($final.Results) { Write-ColorMessage "Purge results: $($final.Results)" -Type "Info" }
            Write-ColorMessage "Items were moved to Recoverable Items (recoverable during the deleted item retention period)." -Type "Info"
            Write-ColorMessage "Search name: $searchName" -Type "Info"
            Write-ColorMessage "If any mailbox had more than 10 matching items, run the script again to remove the rest." -Type "Warning"
        } elseif ($final.Status -in @('Failed', 'PartiallySucceeded')) {
            Write-ColorMessage "Purge finished with status '$($final.Status)'." -Type "Error"
            if ($final.Results) { Write-Host "  Results: $($final.Results)" -ForegroundColor Gray }
            if ($final.Errors)  { Write-Host "  Errors : $($final.Errors)"  -ForegroundColor Gray }
        } else {
            Write-ColorMessage "Purge still '$($final.Status)' after $($script:PurgeTimeoutMinutes) minutes. It keeps running in the background." -Type "Warning"
            Write-Host "  Check later: Get-ComplianceSearchAction -Identity '$actionName' | Format-List Status,Results,Errors" -ForegroundColor Gray
        }

    } catch {
        Write-Host ""
        Write-ColorMessage "ERROR: $($_.Exception.Message)" -Type "Error"

        # Cleanup
        try {
            if (Get-ComplianceSearch -Identity $searchName -ErrorAction SilentlyContinue) {
                Remove-ComplianceSearch -Identity $searchName -Confirm:$false -ErrorAction SilentlyContinue
                Write-ColorMessage "Search removed." -Type "Info"
            }
        } catch {}
    }
}

# ===== MAIN =====
Clear-Host
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "          EMAIL REMOVAL - MICROSOFT 365 (Purview purge)          " -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

try {
    # Step 1: check the module
    Test-RequiredModule

    # Step 2: connect
    Connect-M365Services

    # Step 3: search and purge
    Remove-EmailBySearch

} catch {
    Write-ColorMessage "Fatal error: $($_.Exception.Message)" -Type "Error"
}

Write-Host ""
$null = Read-Host "Press ENTER to exit"
