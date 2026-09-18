<#
.SYNOPSIS
    Exchange Online Message Trace GUI using Get-MessageTraceV2.
.DESCRIPTION
    Connects interactively to Exchange Online, lets the operator search by sender,
    recipient, subject, status, and date range, then opens hop-by-hop details for
    the selected message. Requires an active Exchange Online PIM role.
.NOTES
    PowerShell 7+ and ExchangeOnlineManagement are required.
    Searches are limited to 89 days by design and are split into API-safe windows.
    If a window remains capped at 5,000 results after subdivision, the GUI reports
    the result as incomplete instead of silently treating it as complete.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$AdminUPN
)

Set-StrictMode -Version Latest
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

function Show-Error {
    param([string]$Message)
    [void][System.Windows.Forms.MessageBox]::Show(
        $Message,
        'Message Trace GUI - Error',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
}

function Set-Status {
    param([string]$Text)
    $script:StatusLabel.Text = $Text
    $script:Form.Refresh()
}

try {
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        throw 'PowerShell 7 or newer is required.'
    }

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "ExchangeOnlineManagement is not installed. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }

    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false -ErrorAction Stop

    if (-not (Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue)) {
        throw "Get-MessageTraceV2 is unavailable. Confirm the active PIM role and update ExchangeOnlineManagement."
    }
}
catch {
    Show-Error $_.Exception.Message
    return
}

$script:TraceResults = @()
$script:Form = New-Object System.Windows.Forms.Form
$script:Form.Text = 'Exchange Online Message Trace'
$script:Form.Size = New-Object System.Drawing.Size(1180, 720)
$script:Form.StartPosition = 'CenterScreen'
$script:Form.MinimumSize = New-Object System.Drawing.Size(1000, 600)

$filterPanel = New-Object System.Windows.Forms.Panel
$filterPanel.Dock = 'Top'
$filterPanel.Height = 150
$script:Form.Controls.Add($filterPanel)

function Add-Label([string]$Text, [int]$X, [int]$Y) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.AutoSize = $true
    $filterPanel.Controls.Add($label)
}

function Add-TextBox([int]$X, [int]$Y, [int]$Width) {
    $box = New-Object System.Windows.Forms.TextBox
    $box.Location = New-Object System.Drawing.Point($X, $Y)
    $box.Width = $Width
    $filterPanel.Controls.Add($box)
    return $box
}

Add-Label 'Sender address (optional)' 15 15
$txtSender = Add-TextBox 15 35 330

Add-Label 'Recipient address (optional)' 365 15
$txtRecipient = Add-TextBox 365 35 330

Add-Label 'Subject contains (optional)' 715 15
$txtSubject = Add-TextBox 715 35 300

Add-Label 'Start date/time' 15 72
$dtStart = New-Object System.Windows.Forms.DateTimePicker
$dtStart.Location = New-Object System.Drawing.Point(15, 92)
$dtStart.Width = 210
$dtStart.Format = 'Custom'
$dtStart.CustomFormat = 'yyyy-MM-dd HH:mm'
$dtStart.Value = (Get-Date).AddDays(-1)
$filterPanel.Controls.Add($dtStart)

Add-Label 'End date/time' 245 72
$dtEnd = New-Object System.Windows.Forms.DateTimePicker
$dtEnd.Location = New-Object System.Drawing.Point(245, 92)
$dtEnd.Width = 210
$dtEnd.Format = 'Custom'
$dtEnd.CustomFormat = 'yyyy-MM-dd HH:mm'
$dtEnd.Value = Get-Date
$filterPanel.Controls.Add($dtEnd)

Add-Label 'Status' 475 72
$cmbStatus = New-Object System.Windows.Forms.ComboBox
$cmbStatus.Location = New-Object System.Drawing.Point(475, 92)
$cmbStatus.Width = 150
$cmbStatus.DropDownStyle = 'DropDownList'
[void]$cmbStatus.Items.AddRange(@('All','Delivered','Failed','Pending','Quarantined','FilteredAsSpam','Unknown'))
$cmbStatus.SelectedIndex = 0
$filterPanel.Controls.Add($cmbStatus)

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = 'Search'
$btnSearch.Location = New-Object System.Drawing.Point(650, 89)
$btnSearch.Size = New-Object System.Drawing.Size(110, 30)
$filterPanel.Controls.Add($btnSearch)

$btnDetails = New-Object System.Windows.Forms.Button
$btnDetails.Text = 'View details'
$btnDetails.Location = New-Object System.Drawing.Point(775, 89)
$btnDetails.Size = New-Object System.Drawing.Size(110, 30)
$filterPanel.Controls.Add($btnDetails)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = 'Export CSV'
$btnExport.Location = New-Object System.Drawing.Point(900, 89)
$btnExport.Size = New-Object System.Drawing.Size(110, 30)
$filterPanel.Controls.Add($btnExport)

$btn89Days = New-Object System.Windows.Forms.Button
$btn89Days.Text = 'Last 89 days'
$btn89Days.Location = New-Object System.Drawing.Point(1020, 89)
$btn89Days.Size = New-Object System.Drawing.Size(120, 30)
$filterPanel.Controls.Add($btn89Days)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = 'Fill'
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.MultiSelect = $false
$grid.SelectionMode = 'FullRowSelect'
$grid.AutoSizeColumnsMode = 'Fill'
$grid.AutoGenerateColumns = $true
$script:Form.Controls.Add($grid)
$grid.BringToFront()

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$script:StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:StatusLabel.Text = "Connected as $AdminUPN"
[void]$statusStrip.Items.Add($script:StatusLabel)
$script:Form.Controls.Add($statusStrip)

function Invoke-TraceSearch {
    try {
        $start = $dtStart.Value
        $end = $dtEnd.Value
        if ($end -le $start) { throw 'End date must be later than start date.' }
        if (($end - $start).TotalDays -gt 89) { throw 'The search period cannot exceed 89 days.' }
        if ([string]::IsNullOrWhiteSpace($txtSender.Text) -and [string]::IsNullOrWhiteSpace($txtRecipient.Text)) {
            throw 'Enter at least a sender or recipient address. Tenant-wide archaeology is not a search strategy.'
        }

        $common = @{
            ResultSize = 5000
            ErrorAction = 'Stop'
        }
        if (-not [string]::IsNullOrWhiteSpace($txtSender.Text))    { $common.SenderAddress = $txtSender.Text.Trim() }
        if (-not [string]::IsNullOrWhiteSpace($txtRecipient.Text)) { $common.RecipientAddress = $txtRecipient.Text.Trim() }
        if (-not [string]::IsNullOrWhiteSpace($txtSubject.Text)) {
            $common.Subject = $txtSubject.Text.Trim()
            $common.SubjectFilterType = 'Contains'
        }
        if ($cmbStatus.SelectedItem.ToString() -ne 'All') { $common.Status = $cmbStatus.SelectedItem.ToString() }

        $btnSearch.Enabled = $false
        $btn89Days.Enabled = $false
        Set-Status 'Searching Exchange Online...'

        $raw = [System.Collections.Generic.List[object]]::new()

        function Get-TraceWindow {
            param(
                [datetime]$WindowStart,
                [datetime]$WindowEnd
            )

            $params = @{} + $common
            $params.StartDate = $WindowStart
            $params.EndDate = $WindowEnd
            $batch = @(Get-MessageTraceV2 @params)

            if ($batch.Count -lt 5000) {
                foreach ($item in $batch) { [void]$raw.Add($item) }
                return
            }

            $span = $WindowEnd - $WindowStart
            if ($span.TotalHours -le 1) {
                [void]$raw.Add([pscustomobject]@{
                    TraceWarning = "Window $($WindowStart.ToString('yyyy-MM-dd HH:mm')) to $($WindowEnd.ToString('yyyy-MM-dd HH:mm')) remained capped at 5,000 results. The result set is incomplete; narrow the filters or dates."
                })
                foreach ($item in $batch) { [void]$raw.Add($item) }
                return
            }

            $midpoint = $WindowStart.AddTicks([int64]($span.Ticks / 2))
            Get-TraceWindow -WindowStart $WindowStart -WindowEnd $midpoint
            Get-TraceWindow -WindowStart $midpoint -WindowEnd $WindowEnd
        }

        $cursor = $start
        $windowNumber = 0
        $windowTotal = [Math]::Ceiling(($end - $start).TotalDays / 10)
        while ($cursor -lt $end) {
            $windowNumber++
            $windowEnd = $cursor.AddDays(10)
            if ($windowEnd -gt $end) { $windowEnd = $end }
            Set-Status ("Searching window {0}/{1}: {2:yyyy-MM-dd HH:mm} to {3:yyyy-MM-dd HH:mm}" -f $windowNumber, $windowTotal, $cursor, $windowEnd)
            Get-TraceWindow -WindowStart $cursor -WindowEnd $windowEnd
            $cursor = $windowEnd
        }

        $script:TraceResults = @($raw |
            Where-Object { $_.PSObject.Properties.Name -contains 'MessageTraceId' } |
            Group-Object { '{0}|{1}|{2:O}' -f $_.MessageTraceId, $_.RecipientAddress, ([datetime]$_.Received) } |
            ForEach-Object { $_.Group[0] } |
            Sort-Object Received -Descending)

        $display = $script:TraceResults | Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, MessageTraceId
        $grid.DataSource = [System.Collections.ArrayList]@($display)
        $warningCount = @($raw | Where-Object { $_.PSObject.Properties.Name -contains 'TraceWarning' }).Count
        if ($warningCount -gt 0) {
            Set-Status ("Found {0} message(s). WARNING: {1} result window(s) were capped at 5,000; output is incomplete for those windows." -f $script:TraceResults.Count, $warningCount)
        } else {
            Set-Status ("Found {0} message(s). Double-click a row for details." -f $script:TraceResults.Count)
        }
    }
    catch {
        Show-Error $_.Exception.Message
        Set-Status 'Search failed.'
    }
    finally {
        $btnSearch.Enabled = $true
        $btn89Days.Enabled = $true
    }
}

function Set-Last89Days {
    $dtEnd.Value = Get-Date
    $dtStart.Value = $dtEnd.Value.AddDays(-89)
    Set-Status 'Date range set to the last 89 days.'
}

function Show-TraceDetails {
    try {
        if ($grid.SelectedRows.Count -eq 0) { throw 'Select one message first.' }
        $row = $grid.SelectedRows[0].DataBoundItem
        if (-not $row.MessageTraceId -or -not $row.RecipientAddress) {
            throw 'The selected row is missing MessageTraceId or RecipientAddress.'
        }

        Set-Status 'Loading message details...'
        $detailStart = ([datetime]$row.Received).AddDays(-1)
        $detailEnd = ([datetime]$row.Received).AddDays(1)
        $details = @(Get-MessageTraceDetailV2 -MessageTraceId $row.MessageTraceId -RecipientAddress $row.RecipientAddress -StartDate $detailStart -EndDate $detailEnd -ErrorAction Stop |
            Select-Object Date, Event, Action, Detail)

        $detailGrid = New-Object System.Windows.Forms.DataGridView
        $detailGrid.ReadOnly = $true
        $detailGrid.AllowUserToAddRows = $false
        $detailGrid.SelectionMode = 'FullRowSelect'
        $detailGrid.AutoSizeColumnsMode = 'Fill'
        $detailGrid.Dock = 'Fill'
        $detailGrid.DataSource = [System.Collections.ArrayList]@($details)

        $detailForm = New-Object System.Windows.Forms.Form
        $detailForm.Text = "Trace details: $($row.Subject)"
        $detailForm.Size = New-Object System.Drawing.Size(1050, 520)
        $detailForm.StartPosition = 'CenterParent'
        $detailForm.Controls.Add($detailGrid)
        [void]$detailForm.ShowDialog($script:Form)
        Set-Status 'Details loaded.'
    }
    catch {
        Show-Error $_.Exception.Message
        Set-Status 'Could not load details.'
    }
}

$btnSearch.Add_Click({ Invoke-TraceSearch })
$btn89Days.Add_Click({ Set-Last89Days })
$btnDetails.Add_Click({ Show-TraceDetails })
$grid.Add_CellDoubleClick({ param($sender, $eventArgs) if ($eventArgs.RowIndex -ge 0) { Show-TraceDetails } })

$btnExport.Add_Click({
    try {
        if ($script:TraceResults.Count -eq 0) { throw 'Run a search before exporting.' }
        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Filter = 'CSV files (*.csv)|*.csv'
        $dialog.FileName = "MessageTrace_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss')
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:TraceResults |
                Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, MessageTraceId |
                Export-Csv -Path $dialog.FileName -NoTypeInformation -Encoding UTF8
            Set-Status "Exported to $($dialog.FileName)"
        }
    }
    catch { Show-Error $_.Exception.Message }
})

$script:Form.Add_FormClosed({
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
})

[void]$script:Form.ShowDialog()
