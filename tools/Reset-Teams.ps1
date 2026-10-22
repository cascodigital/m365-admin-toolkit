<#
.SYNOPSIS
    Fully removes Microsoft Teams (new + classic) and reinstalls the new Teams.
.DESCRIPTION
    For when Teams will not open, hangs at sign-in or is stuck on a corrupted cache.
    Removes the MSTeams package (all users, if admin), the provisioned package, classic
    Teams, the Machine-Wide Installer, the Outlook meeting add-in and the target profile's
    caches. Optionally clears the Office token caches (WAM/OneAuth), which forces a new
    sign-in. Reinstalls via winget, with the teamsbootstrapper as fallback.

    SIMULATES BY DEFAULT. The first run only lists what would be removed and saves the
    plan to plan-teams.json. The second run, with the APPLY key, applies that plan.
    Custom video backgrounds are copied to the output folder before deletion.

    Options (target user, clear tokens, reinstall) live in the SCOPE block of the script.
.EXAMPLE
    .\Reset-Teams.ps1
    Simulates and saves the plan.
.EXAMPLE
    $global:APPLY = $true; .\Reset-Teams.ps1
    Applies the plan produced by the simulation (valid for 60 minutes).
.NOTES
    Run as administrator to remove the package for all users.
    The legacy key $global:EXECUTAR is still accepted as an alias of $global:APPLY.
    Logs/verification CSVs go to the 'csv' subfolder next to the script ($env:TEMP if not writable).
#>
# =====================================================================
# Reset-Teams.ps1  -- full removal of Microsoft Teams (new + classic)
#                         and reinstall of the new Teams
#
# IMPACT: closes Teams; removes the MSTeams package (all users, if admin),
#          the provisioned package, classic Teams, the Machine-Wide Installer, the
#          Outlook meeting add-in, caches and temp files of the target profile and,
#          if enabled, the Office token caches (forces re-login in Office).
#          Reinstalls the new Teams at the end (winget; fallback teamsbootstrapper).
# TARGET:  profile of the user logged on to the machine (or $TargetUser). Literal list
#          printed as "TARGETS: n" and saved to plan-teams.json during simulation.
# ROLLBACK: the script's own reinstall + signing in again. Chats and files come from the
#           server. IRREVERSIBLE: custom video backgrounds -- the script
#           copies them to the output folder before deleting.
#
# USAGE (2 runs):
#     .\Reset-Teams.ps1                                 # 1) simulates and saves the plan
#     $global:APPLY = $true; .\Reset-Teams.ps1       # 2) applies the plan
# =====================================================================

# ---------- SAFETY PREAMBLE ----------
# This script SIMULATES by default. To APPLY, run everything in a single command:
#     $global:APPLY = $true; .\Reset-Teams.ps1
$Execute = $false                 # <<< FAIL CLOSED: NEVER remove this line
$__k  = Get-Variable APPLY    -Scope Global -ValueOnly -ErrorAction SilentlyContinue
$__k2 = Get-Variable EXECUTAR -Scope Global -ValueOnly -ErrorAction SilentlyContinue   # legacy key name
$Execute = (($__k -is [bool] -and $__k) -or ($__k2 -is [bool] -and $__k2))
$global:APPLY = $false         # consumes the key
$global:EXECUTAR = $false

$outDir = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.ProviderPath }
try { $__p = Join-Path $outDir "w$PID.tmp"; [IO.File]::WriteAllText($__p,'x'); [IO.File]::Delete($__p) }
catch { $outDir = $env:TEMP; Write-Host "WARNING: output redirected to $outDir" -Foreground Yellow }
$__root = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.ProviderPath }
$csvDir = Join-Path $__root 'csv'   # CSV always in the csv subfolder next to the script, even if the script is moved
if ((-not (Test-Path $csvDir)) -and (Test-Path (Join-Path (Split-Path $__root) 'csv'))) { $csvDir = Join-Path (Split-Path $__root) 'csv' }   # script in a subfolder: use the csv folder of the parent
try { if (-not (Test-Path $csvDir)) { $null = New-Item -ItemType Directory -Path $csvDir -ErrorAction Stop } } catch {}
try { $__p = Join-Path $csvDir "w$PID.tmp"; [IO.File]::WriteAllText($__p,'x'); [IO.File]::Delete($__p) }
catch { $csvDir = $env:TEMP; Write-Host "WARNING: CSV output redirected to $csvDir" -Foreground Yellow }

# ---------- SCOPE (edit here, these are the only switches) ----------
$TargetUser       = ''       # empty = user logged on to the machine. Or the folder name under C:\Users
$ClearTokens      = $true    # WAM/OneAuth/IdentityCache caches -> re-login in Teams AND Office
$ClearCredentials = $true    # Credential Manager (only works when running AS the user's account)
$Reinstall        = $true    # reinstalls the new Teams at the end
$PlanValidityMin  = 60       # a plan older than this is rejected

$PkgNames  = @('MSTeams','MicrosoftTeams')   # new Teams (work) and the Win11 "chat" Teams
$WingetId  = 'Microsoft.Teams'
$BootUrl   = 'https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409'
$ts        = Get-Date -f yyyyMMdd-HHmmss
$planFile     = Join-Path $outDir 'plan-teams.json'

if ($PSVersionTable.PSVersion.Major -ge 7) {
    Import-Module Appx -UseWindowsPowerShell -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
} else {
    Import-Module Appx -ErrorAction SilentlyContinue | Out-Null
}

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$mySid  = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value

# ================= WHO IS THE TARGET USER =================
# When running elevated with a separate admin account, $env:LOCALAPPDATA and HKCU belong to the ADMIN, not to the
# user. That is why the target profile is resolved by SID, never by environment variable.
$plKey  = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
$targetSid = $null
if (-not $isAdmin) {
    $targetSid = $mySid
} elseif ($TargetUser) {
    $hit = @(Get-ChildItem $plKey -ErrorAction SilentlyContinue | Where-Object {
        (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).ProfileImagePath -like "*\$TargetUser" })
    if ($hit.Count -eq 1) { $targetSid = $hit[0].PSChildName }
} else {
    $owners = @(Get-CimInstance Win32_Process -Filter "Name='explorer.exe'" -ErrorAction SilentlyContinue | ForEach-Object {
        $o = Invoke-CimMethod -InputObject $_ -MethodName GetOwnerSid -ErrorAction SilentlyContinue
        if ($o -and $o.Sid) { $o.Sid } } | Sort-Object -Unique)
    if ($owners.Count -eq 1) { $targetSid = $owners[0] }
    elseif ($owners.Count -gt 1) {
        Write-Host "STOP: $($owners.Count) users logged on. Set `$TargetUser at the top of the script." -Foreground Red
        return
    }
}
$prof = $null
if ($targetSid) { $prof = (Get-ItemProperty "$plKey\$targetSid" -ErrorAction SilentlyContinue).ProfileImagePath }
if (-not $prof -or -not (Test-Path -LiteralPath $prof)) {
    Write-Host "STOP: could not find the target user's profile. Set `$TargetUser (folder name under C:\Users)." -Foreground Red
    return
}
$targetName     = Split-Path $prof -Leaf
$sameUser = ($targetSid -eq $mySid)
$L   = Join-Path $prof 'AppData\Local'
$R   = Join-Path $prof 'AppData\Roaming'
$HKU = "Registry::HKEY_USERS\$targetSid"

Write-Host ""
Write-Host "TARGET USER: $targetName  ($targetSid)" -Foreground Cyan
Write-Host "ADMIN: $isAdmin   RUNNING AS THE USER THEMSELVES: $sameUser" -Foreground Cyan

# ================= READ-ONLY FUNCTIONS =================
function Get-TeamsMsi {
    $keys = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
              'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    if ($sameUser) { $keys += "$HKU\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" }
    Get-ItemProperty $keys -ErrorAction SilentlyContinue | Where-Object {
        $_.PSChildName -match '^\{[0-9A-Fa-f\-]{36}\}$' -and
        $_.DisplayName -match '^(Teams Machine-Wide Installer|Microsoft Teams Meeting Add-in|Microsoft Teams classic|Microsoft Teams)' }
}
function Get-TeamsAppx {
    if ($isAdmin) { Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Where-Object { $PkgNames -contains $_.Name } }
    else          { Get-AppxPackage -ErrorAction SilentlyContinue | Where-Object { $PkgNames -contains $_.Name } }
}
function Get-TeamsProv {
    if ($isAdmin) { Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $PkgNames -contains $_.DisplayName } }
}
function Get-State {
    # -User does not accept an Entra account SID (S-1-12-1-...): filter -AllUsers by SID instead
    if ($isAdmin) { $forUser = @(Get-AppxPackage -AllUsers -Name MSTeams -ErrorAction SilentlyContinue | Where-Object {
                        @($_.PackageUserInformation | Where-Object { $_.UserSecurityId.Sid -eq $targetSid -and "$($_.InstallState)" -eq 'Installed' }).Count }) }
    else          { $forUser = Get-AppxPackage -Name MSTeams -ErrorAction SilentlyContinue }
    [pscustomobject]@{
        TeamsForUser = [bool]$forUser
        Provisioned   = if ($isAdmin) { [bool](@(Get-TeamsProv) | Where-Object { $_.DisplayName -eq 'MSTeams' }) } else { 'n/a' }
        PackageFolder    = Test-Path -LiteralPath "$L\Packages\MSTeams_8wekyb3d8bbwe"
        ClassicFolder  = (Test-Path -LiteralPath "$L\Microsoft\Teams") -or (Test-Path -LiteralPath "$R\Microsoft\Teams")
        MsiTeams       = @(Get-TeamsMsi).Count
    }
}
function Get-WingetExe {
    $c = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($c) { return $c.Source }
    $p = @(Resolve-Path "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller_*_8wekyb3d8bbwe\winget.exe" -ErrorAction SilentlyContinue)
    foreach ($x in ($p | Sort-Object Path -Descending)) {
        # the WindowsApps exe may exist and return 'Access denied' on the admin account: it only counts if it RUNS
        try { $null = & $x.Path --version 2>&1; if ($LASTEXITCODE -eq 0) { return $x.Path } } catch {}
    }
    return $null
}

# ================= TARGET DISCOVERY (read-only) =================
$targets = New-Object System.Collections.Generic.List[object]
function Add-Target($order,$kind,$action,$target,$det) {
    $targets.Add([pscustomobject]@{ Order=$order; Kind=$kind; Action=$action; Target=$target; Detail=$det })
}

# 1) processes
foreach ($p in @('ms-teams','Teams','msteamsupdate','ms-teamsupdate')) {
    $procs = @(Get-Process -Name $p -ErrorAction SilentlyContinue)
    if ($procs.Count) { Add-Target 1 'Process' 'Stop-Process -Force' $p "$($procs.Count) instance(s)" }
}

# 2) MSIs: Machine-Wide Installer, meeting add-in, classic
if ($isAdmin) {
    foreach ($m in @(Get-TeamsMsi)) { Add-Target 2 'MSI' 'msiexec /x /qn' $m.PSChildName $m.DisplayName }
}

# 3) Appx package
foreach ($a in @(Get-TeamsAppx)) {
    $action = if ($isAdmin) { 'Remove-AppxPackage -AllUsers' } else { 'Remove-AppxPackage' }
    Add-Target 3 'Appx' $action $a.PackageFullName "version $($a.Version)"
}

# 4) provisioned package (otherwise it comes back on a new profile)
foreach ($pp in @(Get-TeamsProv)) { Add-Target 4 'AppxProv' 'Remove-AppxProvisionedPackage' $pp.PackageName 'provisioned in the image' }

# 5) data, cache and temp folders
$folders = @(
    "$L\Packages\MSTeams_8wekyb3d8bbwe"
    "$L\Packages\MicrosoftTeams_8wekyb3d8bbwe"
    "$L\Microsoft\Teams"
    "$L\Microsoft\TeamsMeetingAdd-in"
    "$L\Microsoft\TeamsMeetingAddin"
    "$L\Microsoft\TeamsPresenceAddin"
    "$L\SquirrelTemp"
    "$R\Microsoft\Teams"
    "$env:ProgramData\$targetName\Microsoft\Teams"
)
if ($isAdmin) { $folders += "${env:ProgramFiles(x86)}\Teams Installer" }
if ($ClearTokens) {
    $folders += @(
        "$L\Microsoft\OneAuth"
        "$L\Microsoft\IdentityCache"
        "$L\Microsoft\TokenBroker\Cache"
        "$L\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\AC\TokenBroker"
    )
}
$folders += @(Get-ChildItem -LiteralPath "$L\Temp" -Force -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -match 'teams' } | ForEach-Object { $_.FullName })
foreach ($f in $folders) {
    if (Test-Path -LiteralPath $f) {
        $sz = 0
        try { $sz = [math]::Round((Get-ChildItem -LiteralPath $f -Recurse -Force -ErrorAction SilentlyContinue |
                   Measure-Object Length -Sum).Sum / 1MB, 1) } catch {}
        Add-Target 5 'Folder' 'Remove-Item -Recurse -Force' $f "$sz MB"
    }
}

# 6) target user's registry (only if the hive is loaded = user logged on)
if (Test-Path -LiteralPath "$HKU\Software\Microsoft\Office\Teams") {
    Add-Target 6 'RegKey' 'Remove-Item (key)' "$HKU\Software\Microsoft\Office\Teams" 'classic Teams config'
}
$runKey = "$HKU\Software\Microsoft\Windows\CurrentVersion\Run"
foreach ($v in @('com.squirrel.Teams.Teams','Teams')) {
    $exists = $false
    try { $exists = $null -ne (Get-ItemProperty -LiteralPath $runKey -Name $v -ErrorAction Stop) } catch {}
    if ($exists) { Add-Target 6 'RegValue' 'Remove-ItemProperty' "$runKey|$v" 'classic autostart' }
}

# 7) credentials (the vault is per account: can only be cleared when running as the user themselves)
if ($ClearCredentials -and $sameUser) {
    # 'Destino:' = Portuguese-localized cmdkey output, 'Target:' = English
    $creds = @(cmdkey /list 2>$null | Select-String 'Destino:|Target:' |
        ForEach-Object { ($_ -split ':',2)[1].Trim() } |
        Where-Object { $_ -match 'msteams|teams|MicrosoftOffice|OneAuthAccount' })
    foreach ($c in ($creds | Sort-Object -Unique)) { Add-Target 7 'Credential' 'cmdkey /delete' $c 'saved credential' }
}

# ================= COUNT =================
Write-Host ""
Write-Host "TARGETS: $($targets.Count)" -Foreground Cyan
if (-not $isAdmin) {
    Write-Host "WARNING: NO admin - Machine-Wide Installer, add-in and provisioned package stay (Teams comes back on its own)." -Foreground Yellow
}
if ($ClearCredentials -and -not $sameUser) {
    Write-Host "INFO: vault credentials skipped (the vault belongs to the user's account, not the admin account)." -Foreground Yellow
}
$targets | Sort-Object Order | Format-Table Kind,Target,Detail -AutoSize | Out-Host

$before = Get-State

# ================= ACTION =================
$failures   = 0
$toApply  = @()
$log      = Join-Path $csvDir "nuke-teams-$ts.csv"
$backgroundsDir = $null

if (-not $Execute) {
    foreach ($a in ($targets | Sort-Object Order)) { Write-Host "SIMULATION: $($a.Action) '$($a.Target)'" }
    # .ToArray(): @(List[object]) inside a hashtable breaks on PS 5.1 ("Argument types do not match")
    [pscustomobject]@{ Created = (Get-Date).ToString('o'); Sid = $targetSid; Targets = $targets.ToArray() } |
        ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $planFile -Encoding UTF8
} else {
    # --- plan gate: only apply what was shown in the simulation ---
    if (-not (Test-Path -LiteralPath $planFile)) {
        Write-Host "STOP: no plan exists. First run WITHOUT the key:  .\Reset-Teams.ps1" -Foreground Red
        return
    }
    $pj    = Get-Content -LiteralPath $planFile -Raw | ConvertFrom-Json
    $ageMin = ((Get-Date) - [datetime]$pj.Created).TotalMinutes
    if ($ageMin -gt $PlanValidityMin) {
        Write-Host "STOP: plan is $([math]::Round($ageMin)) min old (limit $PlanValidityMin). Simulate again." -Foreground Red
        return
    }
    if ($pj.Sid -ne $targetSid) {
        Write-Host "STOP: the plan belongs to another user ($($pj.Sid)). Simulate again." -Foreground Red
        return
    }
    $inPlan = @($pj.Targets | ForEach-Object { "$($_.Kind)|$($_.Target)" })
    $toApply = @($targets | Where-Object { $inPlan -contains "$($_.Kind)|$($_.Target)" } | Sort-Object Order)
    foreach ($x in @($targets | Where-Object { $inPlan -notcontains "$($_.Kind)|$($_.Target)" })) {
        Write-Host "OUTSIDE THE PLAN (will not be touched): $($x.Kind) '$($x.Target)'" -Foreground Yellow
    }

    # --- back up video backgrounds BEFORE removing the Appx (which deletes the folder) ---
    $sources = @("$L\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Backgrounds\Uploads",
                "$R\Microsoft\Teams\Backgrounds\Uploads")
    foreach ($bg in $sources) {
        if (Test-Path -LiteralPath $bg) {
            try {
                $backgroundsDir = Join-Path $outDir "teams-backgrounds-$ts"
                New-Item -ItemType Directory -Path $backgroundsDir -Force -ErrorAction Stop | Out-Null
                Copy-Item -Path (Join-Path $bg '*') -Destination $backgroundsDir -Recurse -Force -ErrorAction Stop
            } catch { Write-Host "WARNING: background backup failed: $($_.Exception.Message)" -Foreground Yellow }
        }
    }

    foreach ($a in $toApply) {
        $res = 'OK'
        try {
            switch ($a.Kind) {
                'Process'    { Get-Process -Name $a.Target -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction Stop
                               Start-Sleep -Milliseconds 500 }
                'MSI'        { $pr = Start-Process msiexec.exe -ArgumentList "/x $($a.Target) /qn /norestart" -Wait -PassThru -ErrorAction Stop
                               if (@(0,1605,3010) -notcontains $pr.ExitCode) { throw "msiexec exited with $($pr.ExitCode)" } }
                'Appx'       { if ($isAdmin) { Remove-AppxPackage -Package $a.Target -AllUsers -ErrorAction Stop }
                               else          { Remove-AppxPackage -Package $a.Target -ErrorAction Stop } }
                'AppxProv'   { Remove-AppxProvisionedPackage -Online -PackageName $a.Target -ErrorAction Stop | Out-Null }
                'Folder'     { if (Test-Path -LiteralPath $a.Target) { Remove-Item -LiteralPath $a.Target -Recurse -Force -ErrorAction Stop }
                               else { $res = 'OK (already gone)' } }
                'RegKey'     { Remove-Item -LiteralPath $a.Target -Recurse -Force -ErrorAction Stop }
                'RegValue'   { $k,$v = $a.Target -split '\|',2
                               Remove-ItemProperty -LiteralPath $k -Name $v -Force -ErrorAction Stop }
                'Credential' { $r = cmdkey /delete:"$($a.Target)" 2>&1
                               if ($LASTEXITCODE -ne 0) { throw "cmdkey: $r" } }
            }
        } catch {
            $failures++
            $res = "FAILED: $($_.Exception.Message)"
        }
        $a | Add-Member -NotePropertyName Result -NotePropertyValue $res -Force
        $a | Export-Csv $log -Append -NoTypeInformation -Encoding UTF8
    }
    Remove-Item -LiteralPath $planFile -Force -ErrorAction SilentlyContinue   # the plan consumes itself
}

# ---------- REINSTALL ----------
$installResult = 'not attempted'
if ($Execute -and $Reinstall) {
    Write-Host ""
    Write-Host "Reinstalling Teams..." -Foreground Cyan
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $ProgressPreference = 'SilentlyContinue'   # script scope; makes the download 10x faster on 5.1

    # 1) does winget exist? if not, try registering it, otherwise install from the official GitHub
    $wg = Get-WingetExe
    if (-not $wg) {
        Write-Host "winget missing - trying to register App Installer..." -Foreground Yellow
        try { Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe -ErrorAction Stop } catch {}
        Start-Sleep -Seconds 3
        $wg = Get-WingetExe
    }
    if (-not $wg) {
        Write-Host "winget missing - downloading from Microsoft's GitHub..." -Foreground Yellow
        try {
            $tmp = Join-Path $env:TEMP "winget-inst-$ts"
            New-Item -ItemType Directory -Path $tmp -Force -ErrorAction Stop | Out-Null
            Invoke-WebRequest -Uri 'https://github.com/microsoft/winget-cli/releases/latest/download/DesktopAppInstaller_Dependencies.zip' `
                -OutFile "$tmp\deps.zip" -UseBasicParsing -ErrorAction Stop
            Invoke-WebRequest -Uri 'https://aka.ms/getwinget' -OutFile "$tmp\winget.msixbundle" -UseBasicParsing -ErrorAction Stop
            Expand-Archive -LiteralPath "$tmp\deps.zip" -DestinationPath "$tmp\deps" -Force -ErrorAction Stop
            $arch = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { 'arm64' } else { 'x64' }
            Get-ChildItem "$tmp\deps" -Recurse -Filter *.appx | Where-Object { $_.FullName -match "\\$arch\\" } |
                ForEach-Object { Add-AppxPackage -Path $_.FullName -ErrorAction SilentlyContinue }
            Add-AppxPackage -Path "$tmp\winget.msixbundle" -ErrorAction Stop
            Start-Sleep -Seconds 3
            $wg = Get-WingetExe
        } catch { Write-Host "WARNING: winget installation failed: $($_.Exception.Message)" -Foreground Yellow }
    }

    # 2) Teams via winget ('winget' source, NOT 'msstore': the Store is often blocked by policy)
    if ($wg) {
        Write-Host "winget: $wg"
        try {
            & $wg install --id $WingetId --exact --source winget --silent --disable-interactivity `
                --accept-package-agreements --accept-source-agreements 2>&1 | Out-Host
            $installResult = "winget exited with $LASTEXITCODE"
        } catch { $installResult = "winget failed: $($_.Exception.Message)" }
    } else { $installResult = 'winget unavailable' }

    # 3) did it work? if not, official fallback: teamsbootstrapper -p (provisions for the machine, needs admin)
    Start-Sleep -Seconds 5
    $ok = [bool](@(Get-TeamsAppx) | Where-Object { $_.Name -eq 'MSTeams' }) -or
          [bool](@(Get-TeamsProv) | Where-Object { $_.DisplayName -eq 'MSTeams' })
    if (-not $ok -and $isAdmin) {
        Write-Host "Teams did not show up - using teamsbootstrapper..." -Foreground Yellow
        try {
            $bs = Join-Path $env:TEMP "teamsbootstrapper-$ts.exe"
            Invoke-WebRequest -Uri $BootUrl -OutFile $bs -UseBasicParsing -ErrorAction Stop
            $out = (& $bs -p 2>&1 | Out-String).Trim()
            $installResult += " | bootstrapper: $out"
        } catch { $installResult += " | bootstrapper failed: $($_.Exception.Message)" }
        $ok = [bool](@(Get-TeamsProv) | Where-Object { $_.DisplayName -eq 'MSTeams' }) -or
              [bool](@(Get-TeamsAppx) | Where-Object { $_.Name -eq 'MSTeams' })
    }
    $installResult = if ($ok) { "OK ($installResult)" } else { "DID NOT REINSTALL ($installResult) - request it via Company Portal" }
}

# ============ VERIFICATION -- LAST BLOCK, NOTHING AFTER IT ============
$after = Get-State
$table = @(
  $before  | Select-Object @{n='Phase';e={'BEFORE'}},TeamsForUser,Provisioned,PackageFolder,ClassicFolder,MsiTeams
  $after | Select-Object @{n='Phase';e={'AFTER' }},TeamsForUser,Provisioned,PackageFolder,ClassicFolder,MsiTeams
)
Write-Host ""
if (-not $Execute) {
    Write-Host '### SIMULATION MODE - NOTHING WAS CHANGED ###' -Foreground Yellow
    if (Test-Path -LiteralPath $planFile) { Write-Host "### Plan saved: $planFile (valid for $PlanValidityMin min) ###" -Foreground Yellow }
    else { Write-Host '### ATTENTION: the plan was NOT saved - the execution will refuse ###' -Foreground Red }
    Write-Host '### Run:  $global:APPLY = $true; .\Reset-Teams.ps1  ###' -Foreground Yellow
} elseif ($failures -gt 0) {
    Write-Host "### PARTIAL - $failures of $($toApply.Count) items failed (see $log) ###" -Foreground Red
} else {
    Write-Host "### APPLIED - $($toApply.Count) items. Check BEFORE vs AFTER ###" -Foreground Green
}
if ($Execute -and $Reinstall) { Write-Host "Reinstall: $installResult" }
if ($backgroundsDir) { Write-Host "Video backgrounds saved to: $backgroundsDir" }
if ($Execute)  { Write-Host "Expected in AFTER: Provisioned=True, PackageFolder=False (created on first sign-in/launch of Teams)." }
$table | Format-Table -AutoSize | Out-Host
$csv = Join-Path $csvDir "verification-nuke-teams-$ts.csv"
$table | Export-Csv $csv -NoTypeInformation -Encoding UTF8
if (Test-Path $csv) { Write-Host "Summary saved to: $csv" } else { Write-Host "ATTENTION: the CSV was NOT saved" -Foreground Red }
