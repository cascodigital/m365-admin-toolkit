<#
.SYNOPSIS
    Bulk random password generator for Microsoft 365 users, filtered by domain.

.DESCRIPTION
    Generates and applies random passwords in bulk to Microsoft 365 / Entra ID users.
    Uses the Microsoft.Graph module and generates passwords in an easy-to-type format
    (AA1234qq): 2 uppercase letters, 4 digits and 2 lowercase letters.

    Main features:
    - Random passwords in a fixed format
    - Automatic filtering by a specific domain (enabled accounts only)
    - Bulk processing with a throttling pause
    - Detailed success/failure report
    - Automatic CSV export with timestamp
    - Permission hints and error handling

.PARAMETER None
    Interactive script - prompts for the target domain at run time.

.EXAMPLE
    .\Set-M365Passwords.ps1
    # Prompt: Enter the domain to apply the new passwords to (example: contoso.com)
    # Input:  contoso.com
    # Lists all enabled users of contoso.com, asks for confirmation,
    # applies the new passwords and writes a CSV report.

.INPUTS
    String - Target domain entered interactively.

.OUTPUTS
    - CSV file: M365_Passwords_[domain]_[timestamp].csv in the current directory
      Columns: Name, Email, NewPassword, Status ('Success' or the error message)
    - Console: processed users and status
    - Summary: success and failure counts

.NOTES
    Version       : 2.0
    Compatibility : PowerShell 5.1+, Windows/Linux/macOS

    Microsoft Graph requirements:
    - Microsoft.Graph module installed
    - Delegated permissions requested:
      * User.ReadWrite.All
      * Directory.ReadWrite.All
      * UserAuthenticationMethod.ReadWrite.All
      * Directory.AccessAsUser.All

    Required admin role (one of):
    - Global Administrator
    - User Administrator
    - Password Administrator

    Generated password format:
    - Pattern: AA1234qq (8 characters)
    - 2 random uppercase letters
    - 4 random digits
    - 2 random lowercase letters

    Security settings:
    - ForceChangePasswordNextSignIn: false
    - Passwords are final (users are not required to change them)
    - 1 second pause between users (throttling)

    WARNING: the CSV contains plain-text passwords. Store it securely and delete
    it once the passwords have been distributed.

.LINK
    https://learn.microsoft.com/graph/api/user-update

.LINK
    https://learn.microsoft.com/graph/permissions-reference
#>


# Install and import the module if needed
if (!(Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Authentication

# Generates a password in the AA1234qq format
function New-RandomPassword {
    # 2 uppercase letters
    $upperChars = -join ((65..90) | Get-Random -Count 2 | ForEach-Object {[char]$_})

    # 4 digits
    $numbers = -join ((0..9) | Get-Random -Count 4)

    # 2 lowercase letters
    $lowerChars = -join ((97..122) | Get-Random -Count 2 | ForEach-Object {[char]$_})

    return $upperChars + $numbers + $lowerChars
}

# Connect to Microsoft Graph with the required permissions
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Green
Write-Host "IMPORTANT: Sign in with an account that holds one of these roles:" -ForegroundColor Yellow
Write-Host "- Global Administrator OR" -ForegroundColor Yellow
Write-Host "- User Administrator OR" -ForegroundColor Yellow
Write-Host "- Password Administrator" -ForegroundColor Yellow

try {
    # Permissions required for password resets
    $scopes = @(
        "User.ReadWrite.All",
        "Directory.ReadWrite.All",
        "UserAuthenticationMethod.ReadWrite.All",
        "Directory.AccessAsUser.All"
    )

    Connect-MgGraph -Scopes $scopes
}
catch {
    Write-Host "Error while connecting: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Show the current context
$context = Get-MgContext
Write-Host "Connected as: $($context.Account)" -ForegroundColor Green

# Ask for the domain
$targetDomain = Read-Host "Enter the domain to apply the new passwords to (example: contoso.com)"

if ([string]::IsNullOrWhiteSpace($targetDomain)) {
    Write-Host "Domain cannot be empty!" -ForegroundColor Red
    exit
}

Write-Host "Looking up users in domain: $targetDomain" -ForegroundColor Yellow

# Retrieve all users
try {
    $allUsers = Get-MgUser -All -Property "Id,UserPrincipalName,DisplayName,AccountEnabled"
}
catch {
    Write-Host "Error while retrieving users: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Check that you have the required permissions." -ForegroundColor Red
    exit
}

# Keep only enabled users of the specified domain
$targetUsers = $allUsers | Where-Object {
    $_.UserPrincipalName -like "*@$targetDomain" -and
    $_.AccountEnabled -eq $true
}

if ($targetUsers.Count -eq 0) {
    Write-Host "No users found in domain $targetDomain" -ForegroundColor Red
    exit
}

Write-Host "Found $($targetUsers.Count) users in domain $targetDomain" -ForegroundColor Green

# List the users found
Write-Host "`nUsers whose password will be changed:" -ForegroundColor Cyan
foreach ($user in $targetUsers) {
    Write-Host "- $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor White
}

# Confirm (S is still accepted for backward compatibility)
$confirm = Read-Host "`nContinue and change the passwords of these users? (Y/N)"
if ($confirm -notin @('Y', 'y', 'S', 's')) {
    Write-Host "Operation cancelled by the user." -ForegroundColor Yellow
    exit
}

# Results collection
$results = @()

# Process each user
Write-Host "`nProcessing users..." -ForegroundColor Green
foreach ($user in $targetUsers) {
    try {
        # Generate a new password
        $newPassword = New-RandomPassword

        Write-Host "Processing: $($user.UserPrincipalName)..." -ForegroundColor Yellow

        # Password profile
        $passwordProfile = @{
            Password = $newPassword
            ForceChangePasswordNextSignIn = $false
        }

        # Update the password
        Update-MgUser -UserId $user.Id -PasswordProfile $passwordProfile -ErrorAction Stop

        # Record the result (CSV column names/values kept for compatibility)
        $results += [PSCustomObject]@{
            'Name' = $user.DisplayName
            'Email' = $user.UserPrincipalName
            'NewPassword' = $newPassword
            'Status' = 'Success'
        }

        Write-Host "[OK] Password changed for: $($user.UserPrincipalName)" -ForegroundColor Green

        # Short pause to avoid throttling
        Start-Sleep -Milliseconds 1000
    }
    catch {
        # On error
        $errorMessage = $_.Exception.Message

        $results += [PSCustomObject]@{
            'Name' = $user.DisplayName
            'Email' = $user.UserPrincipalName
            'NewPassword' = 'ERROR'
            'Status' = $errorMessage
        }

        Write-Host "[FAIL] Error changing password for: $($user.UserPrincipalName)" -ForegroundColor Red
        Write-Host "  Error: $errorMessage" -ForegroundColor Red

        # Hint on permission errors
        if ($errorMessage -like "*Insufficient privileges*" -or $errorMessage -like "*Authorization_RequestDenied*") {
            Write-Host "  HINT: Administrative privileges are required to change passwords!" -ForegroundColor Red
        }
    }
}

# Write the CSV report
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$excelPath = "M365_Passwords_$($targetDomain.Replace('.','_'))_$timestamp.csv"

try {
    $results | Export-Csv -Path $excelPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nFile created: $excelPath" -ForegroundColor Green

    # Offer to open the file
    if (Test-Path $excelPath) {
        $openFile = Read-Host "Open the file now? (Y/N)"
        if ($openFile -in @('Y', 'y', 'S', 's')) {
            Start-Process $excelPath
        }
    }
}
catch {
    Write-Host "Error writing CSV file: $($_.Exception.Message)" -ForegroundColor Red
}

# Final summary
Write-Host "`n=== SUMMARY ===" -ForegroundColor Cyan
Write-Host "Domain processed: $targetDomain" -ForegroundColor White
Write-Host "Total users: $($targetUsers.Count)" -ForegroundColor White
Write-Host "Succeeded: $(($results | Where-Object {$_.Status -eq 'Success'}).Count)" -ForegroundColor Green
Write-Host "Failed: $(($results | Where-Object {$_.Status -ne 'Success'}).Count)" -ForegroundColor Red

# List successful changes
$successes = $results | Where-Object {$_.Status -eq 'Success'}
if ($successes.Count -gt 0) {
    Write-Host "`nPasswords changed successfully:" -ForegroundColor Green
    foreach ($success in $successes) {
        Write-Host "[OK] $($success.Email) - New password: $($success.'NewPassword')" -ForegroundColor Green
    }
}

# Disconnect
Disconnect-MgGraph

Write-Host "`nScript finished!" -ForegroundColor Green
Write-Host "Passwords were set as final - users are not required to change them." -ForegroundColor Yellow

if (($results | Where-Object {$_.Status -ne 'Success'}).Count -gt 0) {
    Write-Host "`nIMPORTANT: Some users failed. Make sure you have:" -ForegroundColor Red
    Write-Host "- The Global Administrator, User Administrator or Password Administrator role" -ForegroundColor Red
    Write-Host "- Adequate permissions in the tenant" -ForegroundColor Red
}
