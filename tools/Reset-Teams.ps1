<#
.SYNOPSIS
    Remove o Microsoft Teams (novo + classico) por completo e reinstala o Teams novo.
.DESCRIPTION
    Para quando o Teams nao abre, trava no login ou fica preso em cache corrompido.
    Remove o pacote MSTeams (todos os usuarios, se admin), o pacote provisionado, o
    Teams classico, o Machine-Wide Installer, o add-in de reuniao do Outlook e os
    caches do perfil alvo. Opcionalmente limpa caches de token do Office (WAM/OneAuth),
    o que forca relogin. Reinstala via winget, com fallback para o teamsbootstrapper.

    SIMULA POR PADRAO. A primeira rodada so lista o que seria removido e grava o
    plano em plano-teams.json. A segunda, com a chave EXECUTAR, aplica esse plano.
    Fundos de video personalizados sao copiados para a pasta de saida antes de apagar.

    Opcoes (usuario alvo, limpar tokens, reinstalar) ficam no bloco ESCOPO do script.
.EXAMPLE
    .\Reset-Teams.ps1
    Simula e grava o plano.
.EXAMPLE
    $global:EXECUTAR = $true; .\Reset-Teams.ps1
    Aplica o plano gerado na simulacao (valido por 60 minutos).
.NOTES
    Rodar como administrador para remover o pacote de todos os usuarios.
    Autor: Andre Kittler da Costa
#>
# =====================================================================
# Reset-Teams.ps1  -- remocao total do Microsoft Teams (novo + classico)
#                         e reinstalacao do Teams novo
#
# IMPACTO: fecha o Teams; remove o pacote MSTeams (todos os usuarios, se admin),
#          o pacote provisionado, o Teams classico, o Machine-Wide Installer, o
#          add-in de reuniao do Outlook, caches e temporarios do perfil alvo e,
#          se ligado, os caches de token do Office (forca relogin no Office).
#          Reinstala o Teams novo no fim (winget; fallback teamsbootstrapper).
# ALVO:    perfil do usuario logado na maquina (ou $UsuarioAlvo). Lista literal
#          impressa como "ALVOS: n" e gravada em plano-teams.json na simulacao.
# ROLLBACK: a reinstalacao do proprio script + relogar. Chats e arquivos vem do
#           servidor. IRREVERSIVEL: fundos de video personalizados -- o script
#           copia eles para a pasta de saida antes de apagar.
#
# USO (2 rodadas, como sempre):
#     .\Reset-Teams.ps1                                 # 1) simula e grava o plano
#     $global:EXECUTAR = $true; .\Reset-Teams.ps1       # 2) aplica o plano
# =====================================================================

# ---------- PREAMBULO DE SEGURANCA ----------
# Este script SIMULA por padrao. Para APLICAR, rode tudo num comando so:
#     $global:EXECUTAR = $true; .\Reset-Teams.ps1
$Execute = $false                 # <<< FALHA FECHADA: NUNCA remover esta linha
$__k = Get-Variable EXECUTAR -Scope Global -ValueOnly -ErrorAction SilentlyContinue
$Execute = ($__k -is [bool] -and $__k)
$global:EXECUTAR = $false         # consome a chave

$saida = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.ProviderPath }
try { $__p = Join-Path $saida "w$PID.tmp"; [IO.File]::WriteAllText($__p,'x'); [IO.File]::Delete($__p) }
catch { $saida = $env:TEMP; Write-Host "AVISO: saida redirecionada para $saida" -Foreground Yellow }

# ---------- ESCOPO (edite aqui, sao as unicas chaves) ----------
$UsuarioAlvo       = ''       # vazio = usuario logado na maquina. Ou o nome da pasta em C:\Users
$LimparTokens      = $true    # caches WAM/OneAuth/IdentityCache -> relogin em Teams E Office
$LimparCredenciais = $true    # Gerenciador de Credenciais (so funciona rodando NA conta do usuario)
$Reinstalar        = $true    # reinstala o Teams novo no fim
$ValidadePlanoMin  = 60       # plano mais velho que isso e recusado

$PkgNames  = @('MSTeams','MicrosoftTeams')   # Teams novo (trabalho) e o Teams "chat" do Win11
$WingetId  = 'Microsoft.Teams'
$BootUrl   = 'https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409'
$ts        = Get-Date -f yyyyMMdd-HHmmss
$plano     = Join-Path $saida 'plano-teams.json'

if ($PSVersionTable.PSVersion.Major -ge 7) {
    Import-Module Appx -UseWindowsPowerShell -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null
} else {
    Import-Module Appx -ErrorAction SilentlyContinue | Out-Null
}

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$meuSid  = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value

# ================= QUEM E O USUARIO ALVO =================
# Rodando elevado com a conta _a, $env:LOCALAPPDATA e HKCU sao do ADMIN, nao do
# usuario. Por isso o perfil alvo e resolvido pelo SID, nunca por variavel de ambiente.
$plKey  = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
$alvoSid = $null
if (-not $isAdmin) {
    $alvoSid = $meuSid
} elseif ($UsuarioAlvo) {
    $hit = @(Get-ChildItem $plKey -ErrorAction SilentlyContinue | Where-Object {
        (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).ProfileImagePath -like "*\$UsuarioAlvo" })
    if ($hit.Count -eq 1) { $alvoSid = $hit[0].PSChildName }
} else {
    $donos = @(Get-CimInstance Win32_Process -Filter "Name='explorer.exe'" -ErrorAction SilentlyContinue | ForEach-Object {
        $o = Invoke-CimMethod -InputObject $_ -MethodName GetOwnerSid -ErrorAction SilentlyContinue
        if ($o -and $o.Sid) { $o.Sid } } | Sort-Object -Unique)
    if ($donos.Count -eq 1) { $alvoSid = $donos[0] }
    elseif ($donos.Count -gt 1) {
        Write-Host "PARE: $($donos.Count) usuarios logados. Preencha `$UsuarioAlvo no topo do script." -Foreground Red
        return
    }
}
$prof = $null
if ($alvoSid) { $prof = (Get-ItemProperty "$plKey\$alvoSid" -ErrorAction SilentlyContinue).ProfileImagePath }
if (-not $prof -or -not (Test-Path -LiteralPath $prof)) {
    Write-Host "PARE: nao achei o perfil do usuario alvo. Preencha `$UsuarioAlvo (nome da pasta em C:\Users)." -Foreground Red
    return
}
$alvoNome     = Split-Path $prof -Leaf
$mesmoUsuario = ($alvoSid -eq $meuSid)
$L   = Join-Path $prof 'AppData\Local'
$R   = Join-Path $prof 'AppData\Roaming'
$HKU = "Registry::HKEY_USERS\$alvoSid"

Write-Host ""
Write-Host "USUARIO ALVO: $alvoNome  ($alvoSid)" -Foreground Cyan
Write-Host "ADMIN: $isAdmin   RODANDO COMO O PROPRIO USUARIO: $mesmoUsuario" -Foreground Cyan

# ================= FUNCOES DE LEITURA =================
function Get-TeamsMsi {
    $keys = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
              'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')
    if ($mesmoUsuario) { $keys += "$HKU\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" }
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
function Get-Estado {
    # -User nao aceita SID de conta Entra (S-1-12-1-...): filtra o -AllUsers pelo SID
    if ($isAdmin) { $doUser = @(Get-AppxPackage -AllUsers -Name MSTeams -ErrorAction SilentlyContinue | Where-Object {
                        @($_.PackageUserInformation | Where-Object { $_.UserSecurityId.Sid -eq $alvoSid -and "$($_.InstallState)" -eq 'Installed' }).Count }) }
    else          { $doUser = Get-AppxPackage -Name MSTeams -ErrorAction SilentlyContinue }
    [pscustomobject]@{
        TeamsNoUsuario = [bool]$doUser
        Provisionado   = if ($isAdmin) { [bool](@(Get-TeamsProv) | Where-Object { $_.DisplayName -eq 'MSTeams' }) } else { 'n/a' }
        PastaPacote    = Test-Path -LiteralPath "$L\Packages\MSTeams_8wekyb3d8bbwe"
        PastaClassico  = (Test-Path -LiteralPath "$L\Microsoft\Teams") -or (Test-Path -LiteralPath "$R\Microsoft\Teams")
        MsiTeams       = @(Get-TeamsMsi).Count
    }
}
function Get-WingetExe {
    $c = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($c) { return $c.Source }
    $p = @(Resolve-Path "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller_*_8wekyb3d8bbwe\winget.exe" -ErrorAction SilentlyContinue)
    foreach ($x in ($p | Sort-Object Path -Descending)) {
        # o exe do WindowsApps pode existir e dar 'Access denied' na conta admin: so vale se RODAR
        try { $null = & $x.Path --version 2>&1; if ($LASTEXITCODE -eq 0) { return $x.Path } } catch {}
    }
    return $null
}

# ================= LEVANTAMENTO DE ALVOS (so leitura) =================
$alvos = New-Object System.Collections.Generic.List[object]
function Add-Alvo($ordem,$tipo,$acao,$alvo,$det) {
    $alvos.Add([pscustomobject]@{ Ordem=$ordem; Tipo=$tipo; Acao=$acao; Alvo=$alvo; Detalhe=$det })
}

# 1) processos
foreach ($p in @('ms-teams','Teams','msteamsupdate','ms-teamsupdate')) {
    $procs = @(Get-Process -Name $p -ErrorAction SilentlyContinue)
    if ($procs.Count) { Add-Alvo 1 'Processo' 'Stop-Process -Force' $p "$($procs.Count) instancia(s)" }
}

# 2) MSIs: Machine-Wide Installer, add-in de reuniao, classico
if ($isAdmin) {
    foreach ($m in @(Get-TeamsMsi)) { Add-Alvo 2 'MSI' 'msiexec /x /qn' $m.PSChildName $m.DisplayName }
}

# 3) pacote Appx
foreach ($a in @(Get-TeamsAppx)) {
    $acao = if ($isAdmin) { 'Remove-AppxPackage -AllUsers' } else { 'Remove-AppxPackage' }
    Add-Alvo 3 'Appx' $acao $a.PackageFullName "versao $($a.Version)"
}

# 4) pacote provisionado (senao volta em perfil novo)
foreach ($pp in @(Get-TeamsProv)) { Add-Alvo 4 'AppxProv' 'Remove-AppxProvisionedPackage' $pp.PackageName 'provisionado na imagem' }

# 5) pastas de dados, cache e temporarios
$pastas = @(
    "$L\Packages\MSTeams_8wekyb3d8bbwe"
    "$L\Packages\MicrosoftTeams_8wekyb3d8bbwe"
    "$L\Microsoft\Teams"
    "$L\Microsoft\TeamsMeetingAdd-in"
    "$L\Microsoft\TeamsMeetingAddin"
    "$L\Microsoft\TeamsPresenceAddin"
    "$L\SquirrelTemp"
    "$R\Microsoft\Teams"
    "$env:ProgramData\$alvoNome\Microsoft\Teams"
)
if ($isAdmin) { $pastas += "${env:ProgramFiles(x86)}\Teams Installer" }
if ($LimparTokens) {
    $pastas += @(
        "$L\Microsoft\OneAuth"
        "$L\Microsoft\IdentityCache"
        "$L\Microsoft\TokenBroker\Cache"
        "$L\Packages\Microsoft.AAD.BrokerPlugin_cne1h2txyewy\AC\TokenBroker"
    )
}
$pastas += @(Get-ChildItem -LiteralPath "$L\Temp" -Force -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -match 'teams' } | ForEach-Object { $_.FullName })
foreach ($f in $pastas) {
    if (Test-Path -LiteralPath $f) {
        $sz = 0
        try { $sz = [math]::Round((Get-ChildItem -LiteralPath $f -Recurse -Force -ErrorAction SilentlyContinue |
                   Measure-Object Length -Sum).Sum / 1MB, 1) } catch {}
        Add-Alvo 5 'Pasta' 'Remove-Item -Recurse -Force' $f "$sz MB"
    }
}

# 6) registro do usuario alvo (so se a hive estiver carregada = usuario logado)
if (Test-Path -LiteralPath "$HKU\Software\Microsoft\Office\Teams") {
    Add-Alvo 6 'RegChave' 'Remove-Item (chave)' "$HKU\Software\Microsoft\Office\Teams" 'config do Teams classico'
}
$runKey = "$HKU\Software\Microsoft\Windows\CurrentVersion\Run"
foreach ($v in @('com.squirrel.Teams.Teams','Teams')) {
    $existe = $false
    try { $existe = $null -ne (Get-ItemProperty -LiteralPath $runKey -Name $v -ErrorAction Stop) } catch {}
    if ($existe) { Add-Alvo 6 'RegValor' 'Remove-ItemProperty' "$runKey|$v" 'autostart do classico' }
}

# 7) credenciais (o cofre e por conta: so da pra limpar rodando como o proprio usuario)
if ($LimparCredenciais -and $mesmoUsuario) {
    $creds = @(cmdkey /list 2>$null | Select-String 'Destino:|Target:' |
        ForEach-Object { ($_ -split ':',2)[1].Trim() } |
        Where-Object { $_ -match 'msteams|teams|MicrosoftOffice|OneAuthAccount' })
    foreach ($c in ($creds | Sort-Object -Unique)) { Add-Alvo 7 'Credencial' 'cmdkey /delete' $c 'credencial salva' }
}

# ================= CONTAGEM =================
Write-Host ""
Write-Host "ALVOS: $($alvos.Count)" -Foreground Cyan
if (-not $isAdmin) {
    Write-Host "AVISO: SEM admin - Machine-Wide Installer, add-in e pacote provisionado ficam (Teams volta sozinho)." -Foreground Yellow
}
if ($LimparCredenciais -and -not $mesmoUsuario) {
    Write-Host "INFO: credenciais do cofre puladas (cofre e da conta do usuario, nao da conta admin)." -Foreground Yellow
}
$alvos | Sort-Object Ordem | Format-Table Tipo,Alvo,Detalhe -AutoSize | Out-Host

$antes = Get-Estado

# ================= ACAO =================
$falhas   = 0
$aplicar  = @()
$log      = Join-Path $saida "nuke-teams-$ts.csv"
$fundosEm = $null

if (-not $Execute) {
    foreach ($a in ($alvos | Sort-Object Ordem)) { Write-Host "SIMULACAO: $($a.Acao) '$($a.Alvo)'" }
    # .ToArray(): @(List[object]) dentro de hashtable quebra no PS 5.1 ("Argument types do not match")
    [pscustomobject]@{ Criado = (Get-Date).ToString('o'); Sid = $alvoSid; Alvos = $alvos.ToArray() } |
        ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $plano -Encoding UTF8
} else {
    # --- gate do plano (R11): so aplica o que foi mostrado na simulacao ---
    if (-not (Test-Path -LiteralPath $plano)) {
        Write-Host "PARE: nao existe plano. Rode primeiro SEM a chave:  .\Reset-Teams.ps1" -Foreground Red
        return
    }
    $pj    = Get-Content -LiteralPath $plano -Raw | ConvertFrom-Json
    $idade = ((Get-Date) - [datetime]$pj.Criado).TotalMinutes
    if ($idade -gt $ValidadePlanoMin) {
        Write-Host "PARE: plano com $([math]::Round($idade)) min (limite $ValidadePlanoMin). Simule de novo." -Foreground Red
        return
    }
    if ($pj.Sid -ne $alvoSid) {
        Write-Host "PARE: o plano e de outro usuario ($($pj.Sid)). Simule de novo." -Foreground Red
        return
    }
    $noPlano = @($pj.Alvos | ForEach-Object { "$($_.Tipo)|$($_.Alvo)" })
    $aplicar = @($alvos | Where-Object { $noPlano -contains "$($_.Tipo)|$($_.Alvo)" } | Sort-Object Ordem)
    foreach ($x in @($alvos | Where-Object { $noPlano -notcontains "$($_.Tipo)|$($_.Alvo)" })) {
        Write-Host "FORA DO PLANO (nao sera tocado): $($x.Tipo) '$($x.Alvo)'" -Foreground Yellow
    }

    # --- backup dos fundos de video ANTES de remover o Appx (que apaga a pasta) ---
    $fontes = @("$L\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Backgrounds\Uploads",
                "$R\Microsoft\Teams\Backgrounds\Uploads")
    foreach ($bg in $fontes) {
        if (Test-Path -LiteralPath $bg) {
            try {
                $fundosEm = Join-Path $saida "teams-fundos-$ts"
                New-Item -ItemType Directory -Path $fundosEm -Force -ErrorAction Stop | Out-Null
                Copy-Item -Path (Join-Path $bg '*') -Destination $fundosEm -Recurse -Force -ErrorAction Stop
            } catch { Write-Host "AVISO: backup dos fundos falhou: $($_.Exception.Message)" -Foreground Yellow }
        }
    }

    foreach ($a in $aplicar) {
        $res = 'OK'
        try {
            switch ($a.Tipo) {
                'Processo'   { Get-Process -Name $a.Alvo -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction Stop
                               Start-Sleep -Milliseconds 500 }
                'MSI'        { $pr = Start-Process msiexec.exe -ArgumentList "/x $($a.Alvo) /qn /norestart" -Wait -PassThru -ErrorAction Stop
                               if (@(0,1605,3010) -notcontains $pr.ExitCode) { throw "msiexec saiu com $($pr.ExitCode)" } }
                'Appx'       { if ($isAdmin) { Remove-AppxPackage -Package $a.Alvo -AllUsers -ErrorAction Stop }
                               else          { Remove-AppxPackage -Package $a.Alvo -ErrorAction Stop } }
                'AppxProv'   { Remove-AppxProvisionedPackage -Online -PackageName $a.Alvo -ErrorAction Stop | Out-Null }
                'Pasta'      { if (Test-Path -LiteralPath $a.Alvo) { Remove-Item -LiteralPath $a.Alvo -Recurse -Force -ErrorAction Stop }
                               else { $res = 'OK (ja tinha sumido)' } }
                'RegChave'   { Remove-Item -LiteralPath $a.Alvo -Recurse -Force -ErrorAction Stop }
                'RegValor'   { $k,$v = $a.Alvo -split '\|',2
                               Remove-ItemProperty -LiteralPath $k -Name $v -Force -ErrorAction Stop }
                'Credencial' { $r = cmdkey /delete:"$($a.Alvo)" 2>&1
                               if ($LASTEXITCODE -ne 0) { throw "cmdkey: $r" } }
            }
        } catch {
            $falhas++
            $res = "FALHA: $($_.Exception.Message)"
        }
        $a | Add-Member -NotePropertyName Resultado -NotePropertyValue $res -Force
        $a | Export-Csv $log -Append -NoTypeInformation -Encoding UTF8
    }
    Remove-Item -LiteralPath $plano -Force -ErrorAction SilentlyContinue   # plano se autoconsome
}

# ---------- REINSTALACAO ----------
$instalou = 'nao tentado'
if ($Execute -and $Reinstalar) {
    Write-Host ""
    Write-Host "Reinstalando o Teams..." -Foreground Cyan
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $ProgressPreference = 'SilentlyContinue'   # escopo do script; deixa o download 10x mais rapido no 5.1

    # 1) winget existe? senao tenta registrar, senao instala do GitHub oficial
    $wg = Get-WingetExe
    if (-not $wg) {
        Write-Host "winget ausente - tentando registrar o App Installer..." -Foreground Yellow
        try { Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe -ErrorAction Stop } catch {}
        Start-Sleep -Seconds 3
        $wg = Get-WingetExe
    }
    if (-not $wg) {
        Write-Host "winget ausente - baixando do GitHub da Microsoft..." -Foreground Yellow
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
        } catch { Write-Host "AVISO: instalacao do winget falhou: $($_.Exception.Message)" -Foreground Yellow }
    }

    # 2) Teams via winget (fonte 'winget', NAO 'msstore': a Store costuma estar bloqueada em ambiente corporativo)
    if ($wg) {
        Write-Host "winget: $wg"
        try {
            & $wg install --id $WingetId --exact --source winget --silent --disable-interactivity `
                --accept-package-agreements --accept-source-agreements 2>&1 | Out-Host
            $instalou = "winget saiu com $LASTEXITCODE"
        } catch { $instalou = "winget falhou: $($_.Exception.Message)" }
    } else { $instalou = 'winget indisponivel' }

    # 3) conferiu? senao fallback oficial: teamsbootstrapper -p (provisiona pra maquina, precisa admin)
    Start-Sleep -Seconds 5
    $ok = [bool](@(Get-TeamsAppx) | Where-Object { $_.Name -eq 'MSTeams' }) -or
          [bool](@(Get-TeamsProv) | Where-Object { $_.DisplayName -eq 'MSTeams' })
    if (-not $ok -and $isAdmin) {
        Write-Host "Teams nao apareceu - usando teamsbootstrapper..." -Foreground Yellow
        try {
            $bs = Join-Path $env:TEMP "teamsbootstrapper-$ts.exe"
            Invoke-WebRequest -Uri $BootUrl -OutFile $bs -UseBasicParsing -ErrorAction Stop
            $out = (& $bs -p 2>&1 | Out-String).Trim()
            $instalou += " | bootstrapper: $out"
        } catch { $instalou += " | bootstrapper falhou: $($_.Exception.Message)" }
        $ok = [bool](@(Get-TeamsProv) | Where-Object { $_.DisplayName -eq 'MSTeams' }) -or
              [bool](@(Get-TeamsAppx) | Where-Object { $_.Name -eq 'MSTeams' })
    }
    $instalou = if ($ok) { "OK ($instalou)" } else { "NAO REINSTALOU ($instalou) - pedir pelo Company Portal" }
}

# ============ VERIFICACAO -- ULTIMO BLOCO, NADA DEPOIS ============
$depois = Get-Estado
$tab = @(
  $antes  | Select-Object @{n='Fase';e={'ANTES' }},TeamsNoUsuario,Provisionado,PastaPacote,PastaClassico,MsiTeams
  $depois | Select-Object @{n='Fase';e={'DEPOIS'}},TeamsNoUsuario,Provisionado,PastaPacote,PastaClassico,MsiTeams
)
Write-Host ""
if (-not $Execute) {
    Write-Host '### MODO SIMULACAO - NADA FOI ALTERADO ###' -Foreground Yellow
    if (Test-Path -LiteralPath $plano) { Write-Host "### Plano gravado: $plano (vale $ValidadePlanoMin min) ###" -Foreground Yellow }
    else { Write-Host '### ATENCAO: o plano NAO foi gravado - a execucao vai recusar ###' -Foreground Red }
    Write-Host '### Rode:  $global:EXECUTAR = $true; .\Reset-Teams.ps1  ###' -Foreground Yellow
} elseif ($falhas -gt 0) {
    Write-Host "### PARCIAL - $falhas de $($aplicar.Count) itens falharam (ver $log) ###" -Foreground Red
} else {
    Write-Host "### APLICADO - $($aplicar.Count) itens. Confira ANTES x DEPOIS ###" -Foreground Green
}
if ($Execute -and $Reinstalar) { Write-Host "Reinstalacao: $instalou" }
if ($fundosEm) { Write-Host "Fundos de video salvos em: $fundosEm" }
if ($Execute)  { Write-Host "Esperado no DEPOIS: Provisionado=True, PastaPacote=False (cria no 1o login/abertura do Teams)." }
$tab | Format-Table -AutoSize | Out-Host
$csv = Join-Path $saida "verificacao-nuke-teams-$ts.csv"
$tab | Export-Csv $csv -NoTypeInformation -Encoding UTF8
if (Test-Path $csv) { Write-Host "Resumo salvo em: $csv" } else { Write-Host "ATENCAO: o CSV NAO foi gravado" -Foreground Red }
