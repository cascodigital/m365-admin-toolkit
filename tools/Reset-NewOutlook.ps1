<#
.SYNOPSIS
    Fully removes the "new Outlook" (olk) and reinstalls it.
.DESCRIPTION
    For when the new Outlook will not open, hangs at sign-in or is stuck on a corrupted
    cache. Removes the Microsoft.OutlookForWindows package (all users, if admin), the
    provisioned package, the target user's local data folder (cache/index/sqlite), the
    WAM/OneAuth token caches and the migration registry values, then reinstalls.
    CLASSIC Outlook (profiles, PST, OST, signatures) is NOT touched.

    DRY RUN BY DEFAULT. The first run only lists what would be removed and saves the
    plan to plan-newoutlook.json. The second run, with the APPLY key, applies that plan.

    Options (target user, clear tokens, reinstall) live in the SCOPE block of the script.
.EXAMPLE
    .\Reset-NewOutlook.ps1
    Dry run; saves the plan.
.EXAMPLE
    $global:APPLY = $true; .\Reset-NewOutlook.ps1
    Applies the plan produced by the dry run (valid for 60 minutes).
.NOTES
    Run as administrator to remove the package for all users and reinstall via the
    official Setup.exe. The legacy key $global:EXECUTAR is still accepted as an alias
    of $global:APPLY. Logs/verification CSVs go to the 'csv' subfolder next to the
    script ($env:TEMP if not writable).
#>
# =====================================================================
# Reset-NewOutlook.ps1  -- full removal of the "new Outlook" (olk)
#                              and reinstall, running as ADMIN
#
# IMPACT:  closes olk; removes the Microsoft.OutlookForWindows package (all
#          users, if admin), the provisioned package, the target user's local
#          data folder (cache/index/sqlite), the WAM/OneAuth token caches
#          and the migration registry values. Reinstalls at the end.
#          CLASSIC Outlook (profiles, PST, OST, signatures) is NOT touched.
# TARGET:  profile of the user logged on to the machine (or $TargetUser). Literal list
#          printed as "TARGETS: n" and saved to plan-newoutlook.json.
# ROLLBACK: the script's own reinstall + sign in again. Mail comes from the server.
#           IRREVERSIBLE: unsynced drafts and local POP accounts in the
#           new Outlook, if any.
#
# USAGE (2 runs):
#     .\Reset-NewOutlook.ps1                                 # 1) dry run, saves the plan
#     $global:APPLY = $true; .\Reset-NewOutlook.ps1       # 2) applies the plan
# =====================================================================

# ---------- SAFETY PREAMBLE ----------
# This script runs as a DRY RUN by default. To APPLY, run it all in a single command:
#     $global:APPLY = $true; .\Reset-NewOutlook.ps1
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
$ClearTokens      = $true    # WAM/OneAuth/IdentityCache caches -> re-sign-in across all of Office
$ClearCredentials = $true    # Credential Manager (only works when running AS the user's account)
$Reinstall        = $true    # reinstalls the new Outlook at the end
$PlanValidityMin  = 60       # a plan older than this is rejected

$PkgName  = 'Microsoft.OutlookForWindows'
$PkgFam   = 'Microsoft.OutlookForWindows_8wekyb3d8bbwe'
$WingetId = '9NRX63209R7B'    # only exists in the msstore source (often blocked by policy) -> only used without admin
$SetupUrl = 'https://go.microsoft.com/fwlink/?linkid=2207851'   # official new Outlook Setup.exe
$ts       = Get-Date -f yyyyMMdd-HHmmss
$planFile    = Join-Path $outDir 'plan-newoutlook.json'

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
$HKU = "Registry::HKEY_USERS\$targetSid"

Write-Host ""
Write-Host "TARGET USER: $targetName  ($targetSid)" -Foreground Cyan
Write-Host "ADMIN: $isAdmin   RUNNING AS THE USER THEMSELVES: $sameUser" -Foreground Cyan

# ================= READ FUNCTIONS =================
function Get-OlkAppx {
    if ($isAdmin) { Get-AppxPackage -AllUsers -Name $PkgName -ErrorAction SilentlyContinue }
    else          { Get-AppxPackage -Name $PkgName -ErrorAction SilentlyContinue }
}
function Get-OlkProv {
    if ($isAdmin) { Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -eq $PkgName } }
}
function Get-State {
    # -User does not accept an Entra account SID (S-1-12-1-...): filter -AllUsers by SID instead
    if ($isAdmin) { $forUser = @(Get-OlkAppx | Where-Object {
                        @($_.PackageUserInformation | Where-Object { $_.UserSecurityId.Sid -eq $targetSid -and "$($_.InstallState)" -eq 'Installed' }).Count }) }
    else          { $forUser = Get-OlkAppx }
    [pscustomobject]@{
        OutlookForUser = [bool]$forUser
        Provisioned     = if ($isAdmin) { [bool](@(Get-OlkProv)) } else { 'n/a' }
        PackageFolder      = Test-Path -LiteralPath "$L\Packages\$PkgFam"
        OlkFolder         = Test-Path -LiteralPath "$L\Microsoft\Olk"
    }
}
function Get-WingetExe {
    $c = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($c) { return $c.Source }
    $p = @(Resolve-Path "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller_*_8wekyb3d8bbwe\winget.exe" -ErrorAction SilentlyContinue)
    foreach ($x in ($p | Sort-Object Path -Descending)) {
        # the WindowsApps exe may exist and still give 'Access denied' on the admin account: it only counts if it RUNS
        try { $null = & $x.Path --version 2>&1; if ($LASTEXITCODE -eq 0) { return $x.Path } } catch {}
    }
    return $null
}

# ================= TARGET INVENTORY (read only) =================
$targets = New-Object System.Collections.Generic.List[object]
function Add-Target($order,$kind,$action,$target,$det) {
    $targets.Add([pscustomobject]@{ Order=$order; Kind=$kind; Action=$action; Target=$target; Detail=$det })
}

# 1) process
$procs = @(Get-Process -Name 'olk' -ErrorAction SilentlyContinue)
if ($procs.Count) { Add-Target 1 'Process' 'Stop-Process -Force' 'olk' "$($procs.Count) instance(s)" }

# 2) Appx package
foreach ($a in @(Get-OlkAppx)) {
    $action = if ($isAdmin) { 'Remove-AppxPackage -AllUsers' } else { 'Remove-AppxPackage' }
    Add-Target 2 'Appx' $action $a.PackageFullName "version $($a.Version)"
}

# 3) provisioned package (otherwise it comes back on a new profile)
foreach ($pp in @(Get-OlkProv)) { Add-Target 3 'AppxProv' 'Remove-AppxProvisionedPackage' $pp.PackageName 'provisioned in the image' }

# 4) data, cache and temp folders
$folders = @(
    "$L\Packages\$PkgFam"
    "$L\Microsoft\Olk"
    "$L\Microsoft\OutlookForWindows"
    "$L\Temp\Outlook Logging"
)
if ($ClearTokens) {
    $folders += @(
        "$L\Microsoft\OneAuth"
        "$L\Microsoft\IdentityCache"
        "$L\Microsoft\TokenBroker\Cache"
        "$L\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy\AC\TokenBroker"
    )
}
foreach ($f in $folders) {
    if (Test-Path -LiteralPath $f) {
        $sz = 0
        try { $sz = [math]::Round((Get-ChildItem -LiteralPath $f -Recurse -Force -ErrorAction SilentlyContinue |
                   Measure-Object Length -Sum).Sum / 1MB, 1) } catch {}
        Add-Target 4 'Folder' 'Remove-Item -Recurse -Force' $f "$sz MB"
    }
}

# 5) new Outlook registry values in the USER'S HKCU (hive loaded = logged on).
#    VALUES, never the whole key: the key holds classic Outlook preferences.
$regVals = @(
    @{ K="$HKU\Software\Microsoft\Office\16.0\Outlook\Options\General"; V='DoNewOutlookAutoMigration' }
    @{ K="$HKU\Software\Microsoft\Office\16.0\Outlook\Options\General"; V='NewOutlookAutoMigrationRetryIntervals' }
    @{ K="$HKU\Software\Microsoft\Office\16.0\Outlook\Options\General"; V='HideNewOutlookToggle' }
    @{ K="$HKU\Software\Microsoft\Office\16.0\Outlook\Preferences";     V='NewOutlookMigrationUserSetting' }
    @{ K="$HKU\Software\Microsoft\Office\16.0\Outlook\Preferences";     V='UseNewOutlook' }
)
foreach ($r in $regVals) {
    $exists = $false
    try { $exists = $null -ne (Get-ItemProperty -LiteralPath $r.K -Name $r.V -ErrorAction Stop) } catch {}
    if ($exists) { Add-Target 5 'RegValue' 'Remove-ItemProperty' "$($r.K)|$($r.V)" 'value' }
}

# 6) credentials (the vault is per account: it can only be cleared when running as the user themselves)
if ($ClearCredentials -and $sameUser) {
    # 'Destino:' = Portuguese-localized cmdkey output, 'Target:' = English
    $creds = @(cmdkey /list 2>$null | Select-String 'Destino:|Target:' |
        ForEach-Object { ($_ -split ':',2)[1].Trim() } |
        Where-Object { $_ -match 'MicrosoftOffice|outlook|office365|OneAuthAccount' })
    foreach ($c in ($creds | Sort-Object -Unique)) { Add-Target 6 'Credential' 'cmdkey /delete' $c 'saved credential' }
}

# ================= COUNT =================
Write-Host ""
Write-Host "TARGETS: $($targets.Count)" -Foreground Cyan
if (-not $isAdmin) {
    Write-Host "WARNING: NO admin - the provisioned package stays (comes back on a new profile) and the reinstall depends on the Store." -Foreground Yellow
}
if ($ClearCredentials -and -not $sameUser) {
    Write-Host "INFO: vault credentials skipped (the vault belongs to the user's account, not the admin account)." -Foreground Yellow
}
$targets | Sort-Object Order | Format-Table Kind,Target,Detail -AutoSize | Out-Host

$before = Get-State

# ================= ACTION =================
$failures  = 0
$toApply = @()
$log     = Join-Path $csvDir "nuke-newoutlook-$ts.csv"

if (-not $Execute) {
    foreach ($a in ($targets | Sort-Object Order)) { Write-Host "DRY RUN: $($a.Action) '$($a.Target)'" }
    # .ToArray(): @(List[object]) inside a hashtable breaks on PS 5.1 ("Argument types do not match")
    [pscustomobject]@{ Created = (Get-Date).ToString('o'); Sid = $targetSid; Targets = $targets.ToArray() } |
        ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $planFile -Encoding UTF8
} else {
    # --- plan gate: only applies what was shown in the dry run ---
    if (-not (Test-Path -LiteralPath $planFile)) {
        Write-Host "STOP: no plan found. Run it first WITHOUT the key:  .\Reset-NewOutlook.ps1" -Foreground Red
        return
    }
    $pj    = Get-Content -LiteralPath $planFile -Raw | ConvertFrom-Json
    $age = ((Get-Date) - [datetime]$pj.Created).TotalMinutes
    if ($age -gt $PlanValidityMin) {
        Write-Host "STOP: plan is $([math]::Round($age)) min old (limit $PlanValidityMin). Run the dry run again." -Foreground Red
        return
    }
    if ($pj.Sid -ne $targetSid) {
        Write-Host "STOP: the plan belongs to another user ($($pj.Sid)). Run the dry run again." -Foreground Red
        return
    }
    $inPlan = @($pj.Targets | ForEach-Object { "$($_.Kind)|$($_.Target)" })
    $toApply = @($targets | Where-Object { $inPlan -contains "$($_.Kind)|$($_.Target)" } | Sort-Object Order)
    foreach ($x in @($targets | Where-Object { $inPlan -notcontains "$($_.Kind)|$($_.Target)" })) {
        Write-Host "NOT IN PLAN (will not be touched): $($x.Kind) '$($x.Target)'" -Foreground Yellow
    }

    foreach ($a in $toApply) {
        $res = 'OK'
        try {
            switch ($a.Kind) {
                'Process'   { Get-Process -Name $a.Target -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction Stop
                               Start-Sleep -Milliseconds 500 }
                'Appx'       { if ($isAdmin) { Remove-AppxPackage -Package $a.Target -AllUsers -ErrorAction Stop }
                               else          { Remove-AppxPackage -Package $a.Target -ErrorAction Stop } }
                'AppxProv'   { Remove-AppxProvisionedPackage -Online -PackageName $a.Target -ErrorAction Stop | Out-Null }
                'Folder'      { if (Test-Path -LiteralPath $a.Target) { Remove-Item -LiteralPath $a.Target -Recurse -Force -ErrorAction Stop }
                               else { $res = 'OK (already gone)' } }
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
    Write-Host "Reinstalling the new Outlook..." -Foreground Cyan
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $ProgressPreference = 'SilentlyContinue'   # script scope; makes the download 10x faster on 5.1

    if ($isAdmin) {
        # official Setup.exe, provisions for the whole machine (does not depend on the Store or winget)
        try {
            $st = Join-Path $env:TEMP "olk-setup-$ts.exe"
            Invoke-WebRequest -Uri $SetupUrl -OutFile $st -UseBasicParsing -ErrorAction Stop
            $pr = Start-Process -FilePath $st -ArgumentList '--provision true --quiet' -Wait -PassThru -ErrorAction Stop
            $installResult = "Setup.exe exited with $($pr.ExitCode)"
        } catch { $installResult = "Setup.exe failed: $($_.Exception.Message)" }
    } else {
        $wg = Get-WingetExe
        if ($wg) {
            try {
                & $wg install --id $WingetId --source msstore --silent --disable-interactivity `
                    --accept-package-agreements --accept-source-agreements 2>&1 | Out-Host
                $installResult = "winget exited with $LASTEXITCODE"
            } catch { $installResult = "winget failed: $($_.Exception.Message)" }
        } else { $installResult = 'winget unavailable' }
    }

    Start-Sleep -Seconds 5
    $ok = [bool](@(Get-OlkProv)) -or [bool](@(Get-OlkAppx))
    $installResult = if ($ok) { "OK ($installResult)" } else { "DID NOT REINSTALL ($installResult) - request it via Company Portal" }
}

# ============ VERIFICATION -- LAST BLOCK, NOTHING AFTER IT ============
$after = Get-State
$tab = @(
  $before  | Select-Object @{n='Phase';e={'BEFORE'}},OutlookForUser,Provisioned,PackageFolder,OlkFolder
  $after | Select-Object @{n='Phase';e={'AFTER' }},OutlookForUser,Provisioned,PackageFolder,OlkFolder
)
Write-Host ""
if (-not $Execute) {
    Write-Host '### DRY RUN MODE - NOTHING WAS CHANGED ###' -Foreground Yellow
    if (Test-Path -LiteralPath $planFile) { Write-Host "### Plan saved: $planFile (valid for $PlanValidityMin min) ###" -Foreground Yellow }
    else { Write-Host '### WARNING: the plan was NOT saved - the run will refuse ###' -Foreground Red }
    Write-Host '### Run:  $global:APPLY = $true; .\Reset-NewOutlook.ps1  ###' -Foreground Yellow
} elseif ($failures -gt 0) {
    Write-Host "### PARTIAL - $failures of $($toApply.Count) items failed (see $log) ###" -Foreground Red
} else {
    Write-Host "### APPLIED - $($toApply.Count) items. Check BEFORE vs AFTER ###" -Foreground Green
}
if ($Execute -and $Reinstall) { Write-Host "Reinstall: $installResult" }
if ($Execute)  { Write-Host "Expected in AFTER: Provisioned=True, OlkFolder=False. The package is registered for the user at next sign-in." }
$tab | Format-Table -AutoSize | Out-Host
$csv = Join-Path $csvDir "verification-nuke-newoutlook-$ts.csv"
$tab | Export-Csv $csv -NoTypeInformation -Encoding UTF8
if (Test-Path $csv) { Write-Host "Summary saved to: $csv" } else { Write-Host "WARNING: the CSV was NOT saved" -Foreground Red }
