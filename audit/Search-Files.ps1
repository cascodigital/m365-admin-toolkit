<#
.SYNOPSIS
    Finds files in OneDrive for Business using the Microsoft Graph search API.

.DESCRIPTION
    Searches OneDrive for Business with two modes:

    1. Single user  - searches the OneDrive of one user (by UPN/email)
    2. Whole domain - searches every provisioned OneDrive of users in a domain

    The file name filter supports wildcards (e.g. *.pdf, report*, *v2*). A filter
    without '*' is treated as *filter*.

    How it works: for each user the script calls the drive search endpoint
    GET /users/{id}/drive/root/search(q='{term}'), follows @odata.nextLink to read
    every page, keeps only files (not folders) and then applies the wildcard filter
    to the file name locally. The search term sent to Graph is the longest
    non-wildcard part of the filter (for "*.pdf" the term is "pdf").

    Results include user, file name, path, size, last modified date and web URL.
    In domain mode the results are also exported to CSV in the current directory.

.EXAMPLE
    .\Search-Files.ps1
    # Interactive: choose single user or domain, then enter the file filter.

.EXAMPLE
    Connect-MgGraph -TenantId contoso.onmicrosoft.com -ClientId <appId> -CertificateThumbprint <thumbprint>
    .\Search-Files.ps1
    # Uses an existing app-only connection (Files.Read.All + User.Read.All application
    # permissions) so that every user's OneDrive can be searched.

.NOTES
    Version       : 5.0
    Requirements  : PowerShell 5.1+, Microsoft.Graph.Authentication, Microsoft.Graph.Users

    Permissions (least privilege):
    - Files.Read.All and User.Read.All
    - With DELEGATED permissions Graph only returns files the signed-in account can
      access. To search other users' OneDrives, either connect app-only (application
      permissions, admin consent) before running the script, or make the admin a site
      collection administrator of each OneDrive.

    Limitations:
    - Search relies on the SharePoint search index; very recent files may not appear yet.
    - Only each user's default drive (OneDrive) is searched.
#>

# Clear the console
Clear-Host

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "                 ONEDRIVE FILE FINDER" -ForegroundColor Cyan
Write-Host "                  MULTI-USER VERSION" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Reads user input from the console
function Get-UserInput {
    param(
        [string]$Prompt,
        [switch]$Required
    )

    do {
        Write-Host $Prompt -NoNewline -ForegroundColor Yellow
        Write-Host ": " -NoNewline
        $userInput = [Console]::ReadLine()
        if ($null -eq $userInput) { $userInput = '' }

        if ($Required -and [string]::IsNullOrWhiteSpace($userInput)) {
            Write-Host "Input cannot be empty. Try again." -ForegroundColor Red
            Write-Host ""
        }
    } while ($Required -and [string]::IsNullOrWhiteSpace($userInput))

    return $userInput.Trim()
}

# Builds the Graph search term from a wildcard filter (longest non-wildcard token)
function Get-SearchTerm {
    param([string]$Filter)

    $tokens = ($Filter -split '[\*\?]') |
        ForEach-Object { $_.Trim().TrimStart('.') } |
        Where-Object { $_ -ne '' } |
        Sort-Object Length -Descending

    return ($tokens | Select-Object -First 1)
}

# Searches the default drive (OneDrive) of one user
function Search-UserFiles {
    param(
        $User,
        [string]$FileFilter,
        [string]$SearchTerm
    )

    $foundFiles = @()

    # Check that the OneDrive is provisioned
    try {
        $drive = Invoke-MgGraphRequest -Method GET -Uri "v1.0/users/$($User.Id)/drive?`$select=id,name" -ErrorAction Stop
    }
    catch {
        return @{
            User   = $User
            Status = "OneDrive not provisioned or not accessible"
            Files  = @()
        }
    }

    try {
        $escapedTerm = [uri]::EscapeDataString($SearchTerm.Replace("'", "''"))
        $select = 'id,name,size,lastModifiedDateTime,parentReference,file,webUrl'
        $uri = "v1.0/users/$($User.Id)/drive/root/search(q='$escapedTerm')?`$select=$select&`$top=200"

        while ($uri) {
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop

            foreach ($item in $response.value) {
                # Files only (skip folders) and apply the wildcard filter to the name
                if ($null -eq $item.file) { continue }
                if ($item.name -notlike $FileFilter) { continue }

                $fullPath = $item.name
                if ($item.parentReference -and $item.parentReference.path) {
                    $relativePath = $item.parentReference.path -replace "^.+?:", ""
                    $fullPath = ("$relativePath/$($item.name)") -replace "^/", ""
                }

                $size = if ($item.size) { [int64]$item.size } else { 0 }

                $foundFiles += [PSCustomObject]@{
                    User         = $User.DisplayName
                    Email        = $User.Mail
                    FileName     = $item.name
                    FullPath     = $fullPath
                    Drive        = $drive.name
                    Size         = $size
                    SizeKB       = [math]::Round($size / 1KB, 2)
                    LastModified = $item.lastModifiedDateTime
                    WebUrl       = $item.webUrl
                    ID           = $item.id
                }
            }

            $uri = $response.'@odata.nextLink'
        }

        return @{
            User   = $User
            Status = "Success"
            Files  = $foundFiles
        }
    }
    catch {
        return @{
            User   = $User
            Status = "Error: $($_.Exception.Message)"
            Files  = $foundFiles
        }
    }
}

# Reads the file filter and derives the search term
function Read-FileFilter {
    while ($true) {
        $fileName = Get-UserInput -Prompt "Enter the file search filter" -Required

        if (-not ($fileName.Contains("*"))) {
            $filter = "*$fileName*"
            Write-Host "Filter applied: $filter" -ForegroundColor Gray
        } else {
            $filter = $fileName
        }

        $term = Get-SearchTerm -Filter $filter
        if ([string]::IsNullOrWhiteSpace($term)) {
            Write-Host "The filter must contain at least one character other than wildcards." -ForegroundColor Red
            continue
        }

        Write-Host "Graph search term: $term" -ForegroundColor Gray
        return @{ Filter = $filter; Term = $term }
    }
}

# Choose the operating mode
Write-Host "CHOOSE THE SEARCH MODE:" -ForegroundColor Cyan
Write-Host "1. Search the OneDrive of a specific user" -ForegroundColor White
Write-Host "2. Search all OneDrives in a domain" -ForegroundColor White
Write-Host ""

$option = Get-UserInput -Prompt "Enter 1 or 2" -Required

while ($option -notin @("1", "2")) {
    Write-Host "Invalid option. Enter 1 or 2." -ForegroundColor Red
    $option = Get-UserInput -Prompt "Enter 1 or 2" -Required
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan

# Check the Graph connection
try {
    $context = Get-MgContext
    if (-not $context) {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow

        $scopes = @(
            "Files.Read.All",
            "User.Read.All"
        )

        Connect-MgGraph -Scopes $scopes -ErrorAction Stop
        Write-Host "Connected to Microsoft Graph" -ForegroundColor Green
    } else {
        Write-Host "Already connected to Microsoft Graph ($($context.AuthType))" -ForegroundColor Green
    }
}
catch {
    Write-Host "ERROR connecting to Microsoft Graph:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    return
}

Write-Host ""

if ($option -eq "1") {
    # MODE 1: single user
    Write-Host "MODE: SINGLE USER SEARCH" -ForegroundColor Cyan
    Write-Host ""

    $targetUpn = Get-UserInput -Prompt "Enter the email/UPN of the user to search" -Required

    Write-Host ""
    Write-Host "FILE FILTER EXAMPLES:" -ForegroundColor Yellow
    Write-Host "  Example 1: mikrotik          (any file with 'mikrotik' in the name)"
    Write-Host "  Example 2: Mikrotik_V2.docx  (exact name)"
    Write-Host "  Example 3: Mikrotik_V2.*     (Mikrotik_V2 with any extension)"
    Write-Host "  Example 4: *v2*              (any file with 'v2' anywhere in the name)"
    Write-Host "  Example 5: *.docx            (all .docx files)"
    Write-Host ""

    $filterInfo = Read-FileFilter

    try {
        Write-Host ""
        Write-Host "Looking up user $targetUpn..." -ForegroundColor Yellow
        $targetUser = Get-MgUser -UserId $targetUpn -Property Id, DisplayName, Mail, UserPrincipalName -ErrorAction Stop

        Write-Host "User found: $($targetUser.DisplayName)" -ForegroundColor Green
        Write-Host ""
        Write-Host "Searching..." -ForegroundColor Yellow

        $result = Search-UserFiles -User $targetUser -FileFilter $filterInfo.Filter -SearchTerm $filterInfo.Term

        if ($result.Files.Count -gt 0) {
            Write-Host ""
            Write-Host "FILES FOUND: $($result.Files.Count)" -ForegroundColor Green
            Write-Host "============================================================" -ForegroundColor Green

            foreach ($file in $result.Files) {
                Write-Host ""
                Write-Host "FILE: $($file.FileName)" -ForegroundColor Yellow
                Write-Host "PATH: $($file.FullPath)" -ForegroundColor White
                Write-Host "DRIVE: $($file.Drive)" -ForegroundColor Cyan
                Write-Host "SIZE: $($file.SizeKB) KB"
                Write-Host "MODIFIED: $($file.LastModified)"
                Write-Host "URL: $($file.WebUrl)"
            }
        } else {
            Write-Host ""
            Write-Host "NO FILES FOUND" -ForegroundColor Yellow
            Write-Host "Status: $($result.Status)" -ForegroundColor Gray
        }
    }
    catch {
        Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    # MODE 2: all users in the domain
    Write-Host "MODE: SEARCH ALL USERS IN A DOMAIN" -ForegroundColor Cyan
    Write-Host ""

    $domain = Get-UserInput -Prompt "Enter the domain (e.g. contoso.com)" -Required

    Write-Host ""
    Write-Host "FILE FILTER EXAMPLES:" -ForegroundColor Yellow
    Write-Host "  Example 1: mikrotik          (any file with 'mikrotik' in the name)"
    Write-Host "  Example 2: Mikrotik_V2.docx  (exact name)"
    Write-Host "  Example 3: *.pdf             (all PDF files)"
    Write-Host ""

    $filterInfo = Read-FileFilter

    try {
        Write-Host ""
        Write-Host "Retrieving users in domain $domain..." -ForegroundColor Yellow

        # All users of the domain
        $allUsers = @(Get-MgUser -All -Property Id, DisplayName, Mail, UserPrincipalName -ErrorAction Stop | Where-Object {
            $_.Mail -like "*@$domain" -or $_.UserPrincipalName -like "*@$domain"
        })

        if ($allUsers.Count -eq 0) {
            Write-Host "No users found in domain $domain" -ForegroundColor Red
            return
        }

        Write-Host "Users found: $($allUsers.Count)" -ForegroundColor Green
        Write-Host ""
        Write-Host "============================================================"
        Write-Host "SEARCHING ALL USERS..."
        Write-Host "============================================================"
        Write-Host ""

        $allResults = @()
        $counter = 0

        foreach ($user in $allUsers) {
            $counter++
            Write-Host "[$counter/$($allUsers.Count)] Processing: $($user.DisplayName) ($($user.Mail))" -ForegroundColor Yellow

            $result = Search-UserFiles -User $user -FileFilter $filterInfo.Filter -SearchTerm $filterInfo.Term

            Write-Host "  Status: $($result.Status)" -ForegroundColor Gray
            if ($result.Files.Count -gt 0) {
                Write-Host "  Found: $($result.Files.Count) file(s)" -ForegroundColor Green
                $allResults += $result.Files
            }

            Write-Host ""
        }

        # Report
        Write-Host "============================================================"
        Write-Host "SEARCH COMPLETE" -ForegroundColor Green
        Write-Host "============================================================"
        Write-Host "Total files found: $($allResults.Count)" -ForegroundColor Green

        if ($allResults.Count -gt 0) {
            # CSV report
            $reportName = "OneDrive_Search_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').csv"
            $reportPath = Join-Path -Path (Get-Location).Path -ChildPath $reportName
            $allResults | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8

            Write-Host "Report saved to: $reportPath" -ForegroundColor Cyan
            Write-Host ""

            # Summary per user
            $summaryPerUser = $allResults | Group-Object User | Sort-Object Count -Descending

            Write-Host "SUMMARY PER USER:" -ForegroundColor Cyan
            foreach ($group in $summaryPerUser) {
                Write-Host "  $($group.Name): $($group.Count) file(s)" -ForegroundColor White
            }

            Write-Host ""
            Write-Host "FIRST 10 FILES FOUND:" -ForegroundColor Cyan
            $allResults | Select-Object -First 10 | ForEach-Object {
                Write-Host "  $($_.User): $($_.FileName)" -ForegroundColor White
            }

            if ($allResults.Count -gt 10) {
                Write-Host "  ... and $($allResults.Count - 10) more file(s). See the CSV for the full list." -ForegroundColor Gray
            }
        } else {
            Write-Host "No files found for any of the searched users." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "============================================================"
Write-Host "Press ENTER to close..." -ForegroundColor Gray
$null = [Console]::ReadLine()
