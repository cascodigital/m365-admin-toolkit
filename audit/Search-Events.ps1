<#
.SYNOPSIS
    Searches Windows event logs for one or more Event IDs (local or remote) and exports the results to Excel.

.DESCRIPTION
    Searches Windows event logs for several Event IDs at once within a date range.
    Local mode scans every .evtx file in the system log folder, including archived and
    rotated logs. Remote mode queries the active logs of another computer with
    Get-WinEvent -ComputerName. Results are exported to Excel with one worksheet per Event ID.

    Main features:
    - Multiple Event IDs combined with a logical OR
    - LOCAL mode: scans all .evtx files, including archived logs
    - REMOTE mode: queries active logs on a remote computer
    - Date-range filtering (UTC-based XPath)
    - Multi-worksheet Excel export grouped by Event ID
    - All EventData properties extracted as columns
    - Skips .evtx files last written before the start date
    - Ignores locked or corrupted log files

    Typical uses:
    - Security incident investigation
    - Logon and access auditing
    - System failure analysis
    - Compliance reporting

.PARAMETER OutputPath
    Folder for the report. Defaults to the folder containing this script.
    If that folder is not writable, $env:TEMP is used instead.

.EXAMPLE
    .\Search-Events.ps1
    # Search type: L
    # Event IDs: 4624,4625,4648
    # Start date: 2024-12-01
    # End date: 2024-12-31
    # Result: local logon analysis

.EXAMPLE
    .\Search-Events.ps1 -OutputPath D:\Reports
    # Search type: R
    # Computer name: SRV-DC01
    # Event IDs: 1074,1076,6005,6006
    # Start date: 2024-11-15
    # End date: 2024-11-30
    # Result: remote shutdown/restart history

.INPUTS
    None. The script prompts for:
    - Search type (L = Local, R = Remote)
    - Remote computer name (Remote only)
    - Comma-separated Event IDs (e.g. 4624,4625,4648)
    - Start and end date (YYYY-MM-DD)

.OUTPUTS
    - Excel file (.xlsx) in -OutputPath
    - One worksheet per Event ID, plus an "All Events" worksheet when there are several IDs
    - Progress and statistics in the console

.NOTES
    Version       : 8.1
    Compatibility : PowerShell 5.1+, Windows Server/Client

    Requirements:
    - Local administrator rights (required to read the Security log locally)
    - ImportExcel module (Install-Module ImportExcel -Scope CurrentUser)
    - Read access to the Windows event logs
    - Remote mode: Get-WinEvent -ComputerName uses the Event Log RPC interface
      (not WinRM). The target must allow "Remote Event Log Management" inbound
      (TCP 135 + dynamic RPC ports) and the account needs rights to read the logs
      (local admin or Event Log Readers).

    Technical details:
    - LOCAL: recursive scan of C:\Windows\System32\winevt\Logs\
    - REMOTE: active logs only, via Get-WinEvent -ComputerName
    - All EventData properties are extracted from the event XML

    Logs searched (Remote):
    - Application, System, Security
    - Microsoft-Windows-PowerShell/Operational
    - Microsoft-Windows-TaskScheduler/Operational
    - TerminalServices, Windows Defender and WindowsUpdateClient operational logs
    - Any other log on the target that contains records

    Known limitations:
    - LOCAL: administrator rights required for the Security log
    - REMOTE: archived logs are not searched, only active logs
    - REMOTE: requires the Remote Event Log Management firewall rules
    - Run time grows with the number of logs and the length of the period

.LINK
    https://learn.microsoft.com/windows/win32/wes/consuming-events

.LINK
    https://learn.microsoft.com/powershell/module/microsoft.powershell.diagnostics/get-winevent
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$OutputPath
)

# --- Output folder ---
function Resolve-OutputDirectory {
    param([string]$Path)

    $candidates = @()
    if (-not [string]::IsNullOrWhiteSpace($Path)) { $candidates += $Path }
    elseif ($PSScriptRoot) { $candidates += $PSScriptRoot }
    else { $candidates += (Get-Location).Path }
    $candidates += $env:TEMP

    foreach ($candidate in $candidates) {
        try {
            if (-not (Test-Path -LiteralPath $candidate)) {
                New-Item -ItemType Directory -Path $candidate -Force -ErrorAction Stop | Out-Null
            }
            $probe = Join-Path -Path $candidate -ChildPath (".write_test_{0}.tmp" -f [guid]::NewGuid())
            [System.IO.File]::WriteAllText($probe, '')
            Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
            return $candidate
        }
        catch {
            Write-Warning "Output folder '$candidate' is not writable. Trying the next location."
        }
    }
    throw "No writable output folder found."
}

# --- Administrator check ---
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "WARNING: This script is not running with administrator privileges."
    Write-Warning "LOCAL searches of the Security log require administrator privileges."
    Write-Warning "REMOTE searches may still work if your account has rights on the target."
    Write-Host ""
}

# --- Configuration and user input ---

# Import the ImportExcel module
try {
    Import-Module ImportExcel -ErrorAction Stop
}
catch {
    Write-Error "ERROR: The 'ImportExcel' module was not found or could not be loaded."
    Write-Error "Run 'Install-Module -Name ImportExcel -Scope CurrentUser' and try again."
    if ($Host.Name -eq "ConsoleHost") {
        Read-Host "Press Enter to exit."
    }
    return
}

# Folder for the exported report
$exportDirectory = Resolve-OutputDirectory -Path $OutputPath
Write-Host "Output folder: $exportDirectory" -ForegroundColor Cyan

# --- Choose local or remote ---
$searchType = $null
$computerName = $null
$isRemote = $false

while ($null -eq $searchType) {
    Write-Host "`nChoose the search type:" -ForegroundColor Cyan
    Write-Host "  [L] Local  - Search all logs on this computer (including archived logs)"
    Write-Host "  [R] Remote - Search active logs on another computer (Event Log RPC)"
    $choice = Read-Host "Enter L for Local or R for Remote"

    if ($choice -match '^[Ll]$') {
        $searchType = "Local"
        $computerName = $env:COMPUTERNAME
        Write-Host "Selected mode: LOCAL search on '$computerName'" -ForegroundColor Green
    }
    elseif ($choice -match '^[Rr]$') {
        $searchType = "Remote"
        $isRemote = $true

        while ([string]::IsNullOrWhiteSpace($computerName)) {
            $computerName = Read-Host "Enter the remote computer name or IP address"
            if ([string]::IsNullOrWhiteSpace($computerName)) {
                Write-Warning "Computer name cannot be empty."
            }
        }

        Write-Host "Selected mode: REMOTE search on '$computerName'" -ForegroundColor Green
        Write-Host "NOTE: The target must allow Remote Event Log Management (RPC, TCP 135 + dynamic ports)." -ForegroundColor Yellow
        Write-Host "NOTE: Remote search covers active logs only, not archived logs." -ForegroundColor Yellow
    }
    else {
        Write-Warning "Invalid option. Enter L or R."
    }
}

# --- Event ID input ---
$eventIds = $null
while ($null -eq $eventIds) {
    try {
        $idText = Read-Host "`nEnter the Event ID(s) to search for, comma-separated (e.g. 4624,4625,4648)"

        # Remove whitespace and split on commas
        $eventIdStrings = $idText -replace '\s+', '' -split ','

        # Convert each value to an integer
        $eventIds = @()
        foreach ($idString in $eventIdStrings) {
            if ([string]::IsNullOrWhiteSpace($idString)) {
                continue
            }
            $eventIds += [int]$idString
        }

        if ($eventIds.Count -eq 0) {
            throw "No valid Event ID was provided"
        }

        Write-Host "Selected Event IDs: $($eventIds -join ', ')" -ForegroundColor Green

    } catch { 
        Write-Warning "Invalid input. Enter numbers separated by commas."
        $eventIds = $null
    }
}

$inputStartDate = $null
$inputEndDate = $null
while ($null -eq $inputStartDate -or $null -eq $inputEndDate) {
    try {
        if ($null -eq $inputStartDate) { $dateText = Read-Host "Enter the START date (YYYY-MM-DD)"; $inputStartDate = [datetime]$dateText }
        if ($null -eq $inputEndDate) { $dateText = Read-Host "Enter the END date (YYYY-MM-DD)"; $inputEndDate = [datetime]$dateText }
        if ($inputEndDate -lt $inputStartDate) { Write-Warning "The end date cannot be earlier than the start date."; $inputStartDate = $null; $inputEndDate = $null }
    } catch { Write-Warning "Invalid date format. Use YYYY-MM-DD."; $inputStartDate = $null; $inputEndDate = $null }
}

$startDateFormatted = $inputStartDate.Date.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.000Z')
$endDateObject = $inputEndDate.Date.AddDays(1).AddTicks(-1)
$endDateFormatted = $endDateObject.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.000Z')

# --- XPath query for multiple Event IDs ---
$eventIdConditions = @()
foreach ($id in $eventIds) {
    $eventIdConditions += "EventID=$id"
}
$eventIdQuery = "(" + ($eventIdConditions -join " or ") + ")"

$xPathQuery = "*[System[$eventIdQuery and TimeCreated[@SystemTime>='$startDateFormatted' and @SystemTime<='$endDateFormatted']]]"

$allFoundEvents = @()

Write-Host "`nSearching for Event IDs '$($eventIds -join ', ')' from $($inputStartDate.ToString('yyyy-MM-dd')) to $($inputEndDate.ToString('yyyy-MM-dd'))" -ForegroundColor Green

# --- Search: local vs remote ---

if ($isRemote) {
    # ========== REMOTE SEARCH ==========
    Write-Host "[REMOTE] Querying active logs on '$computerName'..." -ForegroundColor Cyan
    Write-Host "NOTE: Only active logs are searched (archived logs are not available remotely)." -ForegroundColor Yellow

    # Common logs queried first
    $logsToQuery = @(
        'Application',
        'System',
        'Security',
        'Microsoft-Windows-PowerShell/Operational',
        'Microsoft-Windows-TaskScheduler/Operational',
        'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational',
        'Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational',
        'Microsoft-Windows-Windows Defender/Operational',
        'Microsoft-Windows-WindowsUpdateClient/Operational'
    )

    $logCount = 0
    foreach ($logName in $logsToQuery) {
        $logCount++
        Write-Host "  [$logCount/$($logsToQuery.Count)] Querying log: $logName" -ForegroundColor Gray

        try {
            $events = Get-WinEvent -ComputerName $computerName -LogName $logName -FilterXPath $xPathQuery -ErrorAction Stop
            if ($null -ne $events) {
                $allFoundEvents += $events
                Write-Host "      -> Found $($events.Count) events." -ForegroundColor White
            }
        }
        catch {
            # Ignore logs that do not exist or cannot be read
            if ($_.Exception.Message -notmatch "No events were found") {
                Write-Host "      -> Log not accessible or does not exist." -ForegroundColor DarkGray
            }
        }
    }

    # List all remaining logs on the target and query them too
    Write-Host "`n  [Extra] Checking additional logs on the remote computer..." -ForegroundColor Gray
    try {
        $allRemoteLogs = Get-WinEvent -ComputerName $computerName -ListLog * -ErrorAction Stop | Where-Object { $_.RecordCount -gt 0 -and $_.LogName -notin $logsToQuery }

        foreach ($log in $allRemoteLogs) {
            Write-Host "      Querying: $($log.LogName)" -ForegroundColor DarkGray
            try {
                $events = Get-WinEvent -ComputerName $computerName -LogName $log.LogName -FilterXPath $xPathQuery -ErrorAction Stop
                if ($null -ne $events) {
                    $allFoundEvents += $events
                    Write-Host "      -> Found $($events.Count) events." -ForegroundColor White
                }
            }
            catch {
                # Silently skip logs that cannot be queried
            }
        }
    }
    catch {
        Write-Host "      -> Could not list additional logs." -ForegroundColor DarkGray
    }
}
else {
    # ========== LOCAL SEARCH ==========
    Write-Host "[LOCAL] Searching all log files (.evtx) on '$env:COMPUTERNAME'..." -ForegroundColor Cyan

    $logsPath = "C:\Windows\System32\winevt\Logs"
    Write-Warning "ALL log files in '$logsPath' will be searched. This can take a while."

    Write-Host "[1/2] Searching all log files (.evtx)..." -ForegroundColor Cyan
    $allLogFiles = Get-ChildItem -Path $logsPath -Filter "*.evtx" -File -Recurse -ErrorAction SilentlyContinue

    if ($null -ne $allLogFiles) {
        Write-Host "    -> Found $($allLogFiles.Count) log files (.evtx) to check."

        $fileCount = 0
        foreach ($file in $allLogFiles) {
            $fileCount++

            # Performance: skip files last written before the start date
            if ($file.LastWriteTime -lt $inputStartDate.Date) {
                continue
            }

            Write-Host "    [$fileCount/$($allLogFiles.Count)] Checking: $($file.Name)" -ForegroundColor Gray
            try {
                $archivedEvents = Get-WinEvent -Path $file.FullName -FilterXPath $xPathQuery -ErrorAction Stop
                if ($null -ne $archivedEvents) {
                    $allFoundEvents += $archivedEvents
                    Write-Host "        -> Found $($archivedEvents.Count) events." -ForegroundColor White
                }
            }
            catch {
                # Skip files that cannot be read (locked, corrupted, etc.)
            }
        }
    }
    else {
        Write-Host "    -> No log files (.evtx) found in the log folder." -ForegroundColor Gray
    }
}

# --- Processing and export ---

Write-Host "`n[2/2] Processing and exporting results..." -ForegroundColor Cyan

if ($allFoundEvents.Count -gt 0) {
    Write-Host "$($allFoundEvents.Count) raw events found. Extracting details..." -ForegroundColor Green

    $detailedResults = @()

    foreach ($event in $allFoundEvents) {
        $xml = [xml]$event.ToXml()
        $properties = [ordered]@{
            LogName      = $event.LogName
            TimeCreated  = $event.TimeCreated
            EventID      = $event.Id
            ProviderName = $event.ProviderName
            ComputerName = $event.MachineName
            Level        = $event.LevelDisplayName
            UserID       = $event.UserId
        }

        $xml.Event.EventData.Data | ForEach-Object {
            $properties[$_.Name] = $_.'#text'
        }

        $detailedResults += New-Object -TypeName PSObject -Property $properties
    }

    # --- Output file name ---
    $eventIdsList = $eventIds -join '-'
    $computerNameSafe = $computerName -replace '[^a-zA-Z0-9]', '_'
    $fileName = "EventLog_$($searchType)_$($computerNameSafe)_IDs-$($eventIdsList)_$($inputStartDate.ToString('yyyyMMdd'))-to-$($inputEndDate.ToString('yyyyMMdd')).xlsx"
    $outputPath = Join-Path -Path $exportDirectory -ChildPath $fileName

    Write-Host "Exporting $($detailedResults.Count) events to '$outputPath'..." -ForegroundColor Cyan

    # Group results by Event ID (one worksheet each)
    $groupedResults = $detailedResults | Group-Object -Property EventID

    # Several Event IDs: one worksheet per ID
    if ($groupedResults.Count -gt 1) {
        foreach ($group in $groupedResults) {
            $worksheetName = "Event ID $($group.Name)"
            $group.Group | Export-Excel -Path $outputPath -AutoSize -TableName "EventData_$($group.Name)" -TableStyle Medium9 -WorksheetName $worksheetName
        }

        # Plus a consolidated worksheet
        $detailedResults | Export-Excel -Path $outputPath -AutoSize -TableName "AllEvents" -TableStyle Medium9 -WorksheetName "All Events"
    }
    else {
        # Single Event ID: one worksheet
        $detailedResults | Export-Excel -Path $outputPath -AutoSize -TableName "EventData" -TableStyle Medium9 -WorksheetName "Event ID $($eventIds[0])" -ClearSheet
    }

    Write-Host "`nExport completed successfully." -ForegroundColor Green
    Write-Host "File saved to: $outputPath"

    # Summary per Event ID
    Write-Host "`nEvents found:" -ForegroundColor Cyan
    Write-Host "  Computer: $computerName ($searchType)" -ForegroundColor White
    foreach ($group in $groupedResults) {
        Write-Host "  Event ID $($group.Name): $($group.Count) events" -ForegroundColor White
    }
}
else {
    Write-Host "`nSearch finished. No events with IDs '$($eventIds -join ', ')' were found in the selected period." -ForegroundColor Yellow
    if ($isRemote) {
        Write-Host "HINT: Check that Remote Event Log Management is allowed through the firewall and that you have rights on the remote computer." -ForegroundColor Yellow
    }
}

Write-Host "`nScript finished." -ForegroundColor Cyan
