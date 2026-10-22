<#
.SYNOPSIS
    Collects and analyzes user logon events from Active Directory domain computers.

.DESCRIPTION
    Searches for interactive (human) logon events (Event ID 4624) on Windows computers
    in the Active Directory domain. Collects logon types Interactive (2), Unlock (7),
    RemoteInteractive (10) and CachedInteractive (11) within a date range.

    Main features:
    - Query a single computer or every enabled Windows computer in the domain
    - Custom date range
    - Per-computer timeout and offline handling
    - Export to Excel (ImportExcel), with CSV fallback
    - Collection summary in the console

    Remote collection runs Get-WinEvent inside Invoke-Command, so it uses
    PowerShell Remoting (WinRM, TCP 5985/5986) on each target computer.

.PARAMETER OutputPath
    Folder for the report. Defaults to the folder containing this script.
    If that folder is not writable, $env:TEMP is used instead.

.EXAMPLE
    .\Get-LogonEvents.ps1
    Runs interactively and prompts for:
    - Target computer name, or 'T' for the whole domain
    - Start and end date (YYYY-MM-DD)

.EXAMPLE
    .\Get-LogonEvents.ps1 -OutputPath D:\Reports
    # Enter: SRV-FILE01
    # Enter: 2024-12-01
    # Enter: 2024-12-31

.INPUTS
    None. The script prompts for input.

.OUTPUTS
    - Excel file (.xlsx) in -OutputPath (CSV if the Excel export fails)
    - Collection statistics in the console

.NOTES
    Version       : 2.1
    Compatibility : PowerShell 5.1+, Windows Server/Client

    Requirements:
    - ActiveDirectory module (RSAT)
    - ImportExcel module
    - Run as Administrator
    - WinRM enabled and reachable on the target computers
    - Rights to read the Security log on the target computers (local admin or
      Event Log Readers membership)
    - Logon auditing enabled (Audit Logon - Success)

    Events collected:
    - Event ID 4624 (successful logon)
    - LogonType 2  (Interactive)
    - LogonType 7  (Unlock)
    - LogonType 10 (RemoteInteractive / RDP)
    - LogonType 11 (CachedInteractive)

.LINK
    https://learn.microsoft.com/windows/security/threat-protection/auditing/event-4624
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

# --- Prerequisites ---
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Import-Module ImportExcel -ErrorAction Stop
}
catch {
    Write-Error "ERROR: Module 'ActiveDirectory' or 'ImportExcel' not found."
    Write-Error "Run this script on a machine with RSAT (AD tools) and the ImportExcel module installed."
    Read-Host "Press Enter to exit."
    return
}

# Check for administrator privileges
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "ERROR: This script must be run with administrator privileges."
    Read-Host "Press Enter to exit."
    return
}

# --- Configuration ---
$outputDirectory = Resolve-OutputDirectory -Path $OutputPath
$commandTimeout = 60  # Timeout in seconds per remote computer

Write-Host "Output folder: $outputDirectory" -ForegroundColor Gray

# --- User input (target) ---
Write-Host "=== DOMAIN LOGON EVENT SEARCH ===" -ForegroundColor Cyan
$targetComputerName = $null
while ([string]::IsNullOrWhiteSpace($targetComputerName)) {
    $targetComputerName = Read-Host "Enter the target computer NAME, or 'T' to search the WHOLE domain"
}

# --- Computer lookup ---
$reportNamePart = ""
if ($targetComputerName -eq 'T') {
    Write-Host "`nRetrieving all computers in the domain..." -ForegroundColor Yellow
    $reportNamePart = "Domain"
    try {
        $domainComputers = Get-ADComputer -Filter {OperatingSystem -like "*Windows*"} -Properties Name, OperatingSystem, LastLogonDate |
                           Where-Object { $_.Enabled -eq $true } |
                           Sort-Object Name
        Write-Host "Found $($domainComputers.Count) enabled Windows computers in the domain." -ForegroundColor Green
    }
    catch {
        Write-Error "Error retrieving domain computers: $($_.Exception.Message)"
        Read-Host "Press Enter to exit."
        return
    }
}
else {
    Write-Host "`nLooking up target computer '$targetComputerName' in Active Directory..." -ForegroundColor Yellow
    $reportNamePart = $targetComputerName
    try {
        $singleComputer = Get-ADComputer -Identity $targetComputerName -Properties Name, OperatingSystem, LastLogonDate -ErrorAction Stop
        if ($singleComputer.Enabled -ne $true) {
            Write-Error "Computer '$targetComputerName' is disabled in Active Directory."
            Read-Host "Press Enter to exit."
            return
        }
        $domainComputers = @($singleComputer) # Wrap the single object in an array
        Write-Host "Target set to: $($singleComputer.Name)" -ForegroundColor Green
    }
    catch {
        Write-Error "Error: computer '$targetComputerName' not found in Active Directory, or the query failed."
        Read-Host "Press Enter to exit."
        return
    }
}

# --- User input (dates) ---
Write-Host "`nDate format: YYYY-MM-DD (example: 2025-01-15)" -ForegroundColor Yellow

$inputStartDate = $null
$inputEndDate = $null

while ($null -eq $inputStartDate -or $null -eq $inputEndDate) {
    try {
        if ($null -eq $inputStartDate) {
            $dateText = Read-Host "Enter the START date (YYYY-MM-DD)"
            $inputStartDate = [datetime]$dateText
        }
        if ($null -eq $inputEndDate) {
            $dateText = Read-Host "Enter the END date (YYYY-MM-DD)"
            $inputEndDate = [datetime]$dateText
        }
        if ($inputEndDate -lt $inputStartDate) {
            Write-Warning "The end date cannot be earlier than the start date."
            $inputStartDate = $null
            $inputEndDate = $null
        }
    }
    catch {
        Write-Warning "Invalid date format. Use YYYY-MM-DD."
        $inputStartDate = $null
        $inputEndDate = $null
    }
}

# --- Query preparation ---
$startDateFormatted = $inputStartDate.Date.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.000Z')
$endDateObject = $inputEndDate.Date.AddDays(1).AddTicks(-1)
$endDateFormatted = $endDateObject.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.000Z')

# Human logons: Interactive (2), RemoteInteractive (10), Unlock (7), CachedInteractive (11)
$xPathQuery = "*[System[EventID=4624 and TimeCreated[@SystemTime>='$startDateFormatted' and @SystemTime<='$endDateFormatted']] and EventData[Data[@Name='LogonType']='2' or Data[@Name='LogonType']='10' or Data[@Name='LogonType']='7' or Data[@Name='LogonType']='11']]"

Write-Host "`nPeriod: $($inputStartDate.ToString('yyyy-MM-dd')) to $($inputEndDate.ToString('yyyy-MM-dd'))" -ForegroundColor Yellow

# --- Event collection ---
$allFoundEvents = New-Object System.Collections.ArrayList
$successCount = 0
$failCount = 0
$offlineCount = 0
$timeoutCount = 0
$totalComputers = $domainComputers.Count

Write-Host "`nStarting event collection..." -ForegroundColor Yellow
Write-Host "Logon types: Interactive (2), RemoteInteractive (10), Unlock (7), CachedInteractive (11)" -ForegroundColor Yellow

foreach ($computer in $domainComputers) {
    $computerName = $computer.Name
    $currentIndex = $domainComputers.IndexOf($computer) + 1

    Write-Host "[$currentIndex/$totalComputers] Checking: $computerName" -ForegroundColor White

    if (Test-Connection -ComputerName $computerName -Count 1 -Quiet -ErrorAction SilentlyContinue) {
        try {
            # Run in a job with a timeout so one unresponsive computer cannot hang the script
            $job = Start-Job -ScriptBlock {
                param($computerName, $xPathQuery)
                try {
                    Invoke-Command -ComputerName $computerName -ScriptBlock {
                        param($xPathQuery)
                        try {
                            $events = Get-WinEvent -LogName Security -FilterXPath $xPathQuery -ErrorAction Stop

                            $results = @()
                            foreach ($event in $events) {
                                $xml = [xml]$event.ToXml()
                                $results += [PSCustomObject]@{
                                    LogonTime    = $event.TimeCreated
                                    ComputerName = $event.MachineName
                                    LogonType    = $xml.Event.EventData.Data[8].'#text'
                                    UserName     = $xml.Event.EventData.Data[5].'#text'
                                    UserDomain   = $xml.Event.EventData.Data[6].'#text'
                                    IpAddress    = if ($xml.Event.EventData.Data[18]) { $xml.Event.EventData.Data[18].'#text' } else { "N/A" }
                                    Workstation  = if ($xml.Event.EventData.Data[11]) { $xml.Event.EventData.Data[11].'#text' } else { "N/A" }
                                }
                            }
                            return $results
                        }
                        catch {
                            return @()
                        }
                    } -ArgumentList $xPathQuery -ErrorAction Stop
                }
                catch {
                    throw $_
                }
            } -ArgumentList $computerName, $xPathQuery

            # Wait for the job with a timeout
            $completed = $job | Wait-Job -Timeout $commandTimeout

            if ($completed) {
                $remoteEvents = Receive-Job -Job $job
                Remove-Job -Job $job -Force

                if ($null -ne $remoteEvents -and $remoteEvents.Count -gt 0) {
                    $remoteEvents | ForEach-Object { [void]$allFoundEvents.Add($_) }
                    Write-Host "   -> OK: found $($remoteEvents.Count) events on $computerName" -ForegroundColor Green
                } else {
                    Write-Host "   -> NO EVENTS: no matching logons on $computerName" -ForegroundColor Gray
                }
                $successCount++
            } else {
                # Timeout - stop the job and move on
                Write-Host "   -> TIMEOUT: $computerName took longer than ${commandTimeout}s - skipping..." -ForegroundColor Yellow
                Stop-Job -Job $job
                Remove-Job -Job $job -Force
                $timeoutCount++
            }
        }
        catch {
            Write-Warning "   -> FAILED: could not read logs on '$computerName': $($_.Exception.Message)"
            $failCount++
        }
    }
    else {
        Write-Host "   -> OFFLINE: $computerName is not responding" -ForegroundColor DarkYellow
        $offlineCount++
    }
}

# --- Collection summary ---
Write-Host "`n=== COLLECTION SUMMARY ===" -ForegroundColor Cyan
Write-Host "Computers processed successfully: $successCount" -ForegroundColor Green
Write-Host "Computers offline: $offlineCount" -ForegroundColor Yellow
Write-Host "Computers timed out: $timeoutCount" -ForegroundColor Yellow
Write-Host "Computers failed: $failCount" -ForegroundColor Red
Write-Host "Total logon events collected: $($allFoundEvents.Count)" -ForegroundColor White

# --- Processing and export ---
if ($allFoundEvents.Count -gt 0) {
    Write-Host "`nProcessing events for export..." -ForegroundColor Yellow

    $detailedResults = $allFoundEvents | ForEach-Object {
        $properties = [ordered]@{
            LogonTime    = $_.LogonTime
            ComputerName = $_.ComputerName
            LogonType    = switch ($_.LogonType) {
                2  { "Interactive" }
                7  { "Unlock" }
                10 { "RemoteInteractive" }
                11 { "CachedInteractive" }
                default { "Other ($($_.LogonType))" }
            }
            UserName     = $_.UserName
            UserDomain   = $_.UserDomain
            IpAddress    = $_.IpAddress
            Workstation  = $_.Workstation
        }
        New-Object -TypeName PSObject -Property $properties
    }

    $reportFile = Join-Path -Path $outputDirectory -ChildPath "LogonReport_$($reportNamePart)_$($inputStartDate.ToString('yyyyMMdd'))-to-$($inputEndDate.ToString('yyyyMMdd')).xlsx"

    try {
        $detailedResults | Export-Excel -Path $reportFile -AutoSize -TableName "Logons" -TableStyle Medium9 -ClearSheet
        Write-Host "`nDone. Excel report saved to:" -ForegroundColor Green
        Write-Host "$reportFile" -ForegroundColor White
    }
    catch {
        Write-Warning "Excel export failed. Saving as CSV..."
        $csvPath = $reportFile -replace '\.xlsx$', '.csv'
        $detailedResults | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host "CSV report saved to: $csvPath" -ForegroundColor Yellow
    }

    Write-Host "`nPreview of the first 10 events:" -ForegroundColor Cyan
    $detailedResults | Select-Object -First 10 | Format-Table -AutoSize
}
else {
    Write-Host "`nNo matching logon events found for the selected target and period." -ForegroundColor Yellow
}

Write-Host "`nScript finished." -ForegroundColor Green
Read-Host "Press Enter to exit"
