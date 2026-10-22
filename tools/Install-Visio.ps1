<#
.SYNOPSIS
    Installs Visio Professional (subscription) via the Office Deployment Tool.
.DESCRIPTION
    Detects the architecture of the installed Microsoft 365 Apps (32/64-bit) and installs
    VisioProRetail with the same bitness, MonthlyEnterprise channel, OS language.
    Gets the ODT setup.exe in this order: -OdtPath, a previously downloaded copy, the
    Office CDN, the Download Center. Every binary's Microsoft Authenticode signature is
    validated before it runs. Log in %TEMP%\VisioODT.
.PARAMETER OdtPath
    Local or UNC path to a known-good ODT setup.exe. Tried before any download.
.EXAMPLE
    .\Install-Visio.ps1
.EXAMPLE
    .\Install-Visio.ps1 -OdtPath '\\fileserver\it$\ODT\setup.exe'
.NOTES
    Requires administrator. Requires a Visio license assigned to the user.
#>
#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    # Optional: local path or UNC share holding a known-good ODT setup.exe.
    # Tried before any download. Example: -OdtPath '\\fileserver\it$\ODT\setup.exe'
    [string]$OdtPath
)

$ErrorActionPreference = 'Stop'
$WorkDir = Join-Path $env:TEMP 'VisioODT'
$LogFile = Join-Path $WorkDir ("Install-Visio-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))

function Write-Log {
    param([Parameter(Mandatory)][string]$Message)
    $line = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    if (-not (Test-Path -LiteralPath $WorkDir)) {
        New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
    }
    Add-Content -LiteralPath $LogFile -Value $line -Encoding UTF8
    Write-Host $line
}

function Test-OdtBinary {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return $false }

    $item = Get-Item -LiteralPath $Path
    if ($item.Length -lt 1MB) {
        Write-Log "Rejected ODT candidate '$Path': implausible size ($($item.Length) bytes)."
        return $false
    }

    # Authenticode is the real trust anchor here - not the URL it came from.
    $sig = Get-AuthenticodeSignature -LiteralPath $Path
    if ($sig.Status -ne 'Valid') {
        Write-Log "Rejected ODT candidate '$Path': signature status is $($sig.Status)."
        return $false
    }
    if ($sig.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
        Write-Log "Rejected ODT candidate '$Path': signer is $($sig.SignerCertificate.Subject)."
        return $false
    }

    $info = $item.VersionInfo
    Write-Log ("Verified ODT: {0} v{1}, signed by Microsoft Corporation." -f $item.Name, $info.FileVersion)
    return $true
}

function Get-OfficeDeploymentTool {
    param(
        [Parameter(Mandatory)][string]$Destination,
        [string]$PreferredSource
    )

    $setupPath = Join-Path $Destination 'setup.exe'

    # 1. Already downloaded and verified in this work directory.
    if ((Test-Path -LiteralPath $setupPath) -and (Test-OdtBinary -Path $setupPath)) {
        Write-Log "Reusing cached ODT at $setupPath."
        return $setupPath
    }

    [Net.ServicePointManager]::SecurityProtocol = `
        [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11

    # 2. Operator-supplied copy (internal share, USB, pinned build).
    if ($PreferredSource) {
        Write-Log "Copying ODT from supplied source: $PreferredSource"
        try {
            Copy-Item -LiteralPath $PreferredSource -Destination $setupPath -Force
            if (Test-OdtBinary -Path $setupPath) { return $setupPath }
            Write-Log 'Supplied ODT failed verification. Falling back.'
        }
        catch {
            Write-Log "Could not use supplied ODT: $($_.Exception.Message). Falling back."
        }
        Remove-Item -LiteralPath $setupPath -Force -ErrorAction SilentlyContinue
    }

    # 3. Evergreen Office CDN. This is the ODT setup.exe itself - Microsoft
    #    refreshes it in place, so there is no version, page or link to scrape.
    $cdnUrl = 'https://officecdn.microsoft.com/pr/wsus/setup.exe'
    Write-Log "Downloading ODT from Office CDN: $cdnUrl"

    $oldProgress = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'
    try {
        Invoke-WebRequest -Uri $cdnUrl -OutFile $setupPath -UseBasicParsing -TimeoutSec 300
        if (Test-OdtBinary -Path $setupPath) { return $setupPath }
        Write-Log 'CDN download failed verification. Falling back to download centre.'
    }
    catch {
        Write-Log "CDN download failed: $($_.Exception.Message). Falling back to download centre."
    }
    finally {
        $ProgressPreference = $oldProgress
    }
    Remove-Item -LiteralPath $setupPath -Force -ErrorAction SilentlyContinue

    # 4. Last resort: scrape the Download Center page for the self-extracting
    #    package. Fragile by nature - kept only so the script degrades instead
    #    of dying if the CDN path ever moves.
    $downloadPage = 'https://www.microsoft.com/en-us/download/details.aspx?id=49117'
    Write-Log 'Attempting Download Center fallback.'
    $page = Invoke-WebRequest -Uri $downloadPage -UseBasicParsing
    $pattern = 'https://download\.microsoft\.com/download/[^"'']+officedeploymenttool[^"'']+\.exe'
    $match = [regex]::Match($page.Content, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) {
        throw 'Could not obtain the Office Deployment Tool from the CDN or the Download Center. Supply one with -OdtPath.'
    }

    $odtInstaller = Join-Path $Destination 'officedeploymenttool.exe'
    Invoke-WebRequest -Uri $match.Value -OutFile $odtInstaller -UseBasicParsing
    if (-not (Test-OdtBinary -Path $odtInstaller)) {
        throw 'Download Center package failed signature verification.'
    }

    Write-Log 'Extracting Office Deployment Tool.'
    $extract = Start-Process -FilePath $odtInstaller -ArgumentList "/quiet /extract:`"$Destination`"" -Wait -PassThru
    if ($extract.ExitCode -ne 0 -or -not (Test-OdtBinary -Path $setupPath)) {
        throw "Office Deployment Tool extraction failed. Exit code: $($extract.ExitCode)"
    }

    return $setupPath
}

function Wait-OdtProcess {
    param(
        [Parameter(Mandatory)][System.Diagnostics.Process]$Process,
        [Parameter(Mandatory)][string]$Action
    )

    $started = Get-Date
    $frames = @('|', '/', '-', '\')
    $frame = 0

    while (-not $Process.HasExited) {
        $elapsed = (Get-Date) - $started
        $status = '{0} {1} | Elapsed {2:hh\:mm\:ss} | Do not close this window' -f $frames[$frame % $frames.Count], $Action, $elapsed
        Write-Progress -Activity 'Microsoft Visio' -Status $status -PercentComplete -1
        Write-Host ("`r{0}" -f $status) -NoNewline
        Start-Sleep -Seconds 2
        $frame++
        $Process.Refresh()
    }

    $Process.WaitForExit()
    Write-Progress -Activity 'Microsoft Visio' -Completed
    Write-Host ''
}

try {
    New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
    Write-Log 'Starting Visio installation.'

    $runningApps = @(Get-Process -Name WINWORD,EXCEL,OUTLOOK,POWERPNT,MSACCESS,ONENOTE,VISIO -ErrorAction SilentlyContinue)
    if ($runningApps.Count -gt 0) {
        throw "Close Microsoft 365 applications first: $(($runningApps.ProcessName | Sort-Object -Unique) -join ', ')"
    }

    $c2rPath = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    if (-not (Test-Path -LiteralPath $c2rPath)) {
        throw 'Microsoft 365 Apps Click-to-Run was not detected.'
    }

    $c2r = Get-ItemProperty -Path $c2rPath
    if ($c2r.Platform -notin @('x64', 'x86')) {
        throw "Unknown Office architecture: $($c2r.Platform)"
    }

    $edition = if ($c2r.Platform -eq 'x64') { '64' } else { '32' }
    $channel = 'MonthlyEnterprise'
    Write-Log "Detected Office architecture: $edition-bit. Target channel: $channel."

    $setupPath = Get-OfficeDeploymentTool -Destination $WorkDir -PreferredSource $OdtPath
    if ($setupPath -isnot [string]) {
        throw "Internal error: ODT path resolved to $($setupPath.GetType().FullName) instead of a string."
    }
    Write-Log "Using ODT: $setupPath"

    $xmlPath = Join-Path $WorkDir 'Install-Visio.xml'

    $xml = @"
<Configuration>
  <Add OfficeClientEdition="$edition" Channel="$channel">
    <Product ID="VisioProRetail">
      <Language ID="MatchOS" />
    </Product>
  </Add>
  <Property Name="AUTOACTIVATE" Value="1" />
  <Property Name="FORCEAPPSHUTDOWN" Value="FALSE" />
  <Display Level="Full" AcceptEULA="TRUE" />
  <Updates Enabled="TRUE" Channel="$channel" />
  <RemoveMSI />
</Configuration>
"@
    [System.IO.File]::WriteAllText($xmlPath, $xml, (New-Object System.Text.UTF8Encoding($false)))

    Write-Log 'Installing Visio Professional subscription desktop app.'
    $process = Start-Process -FilePath $setupPath -ArgumentList "/configure `"$xmlPath`"" -PassThru
    Wait-OdtProcess -Process $process -Action 'Installing Visio'

    if ($process.ExitCode -ne 0) {
        throw "Visio installation failed. Exit code: $($process.ExitCode)"
    }

    $products = (Get-ItemProperty -Path $c2rPath).ProductReleaseIds
    if ($products -notmatch 'VisioProRetail') {
        throw 'ODT returned exit code 0, but VisioProRetail was not detected in Click-to-Run configuration.'
    }

    Write-Log 'Visio installation completed successfully.'
    Write-Host "SUCCESS. Log: $LogFile" -ForegroundColor Green
}
catch {
    Write-Progress -Activity 'Microsoft Visio' -Completed
    Write-Log "ERROR: $($_.Exception.Message)"
    Write-Error "Installation failed. Log: $LogFile`n$($_.Exception.Message)"
    exit 1
}
