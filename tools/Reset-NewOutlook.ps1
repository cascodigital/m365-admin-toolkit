<#
.SYNOPSIS
    Remove o Novo Outlook (olk) por completo e reinstala.
.DESCRIPTION
    Para quando o Novo Outlook nao abre, fica preso em conta antiga ou com cache
    corrompido. Fecha o olk, remove o pacote Microsoft.OutlookForWindows (todos os
    usuarios, se admin), o pacote provisionado e os caches do perfil alvo.
    Opcionalmente limpa caches de token do Office (WAM/OneAuth), o que forca relogin.
    Reinstala pelo Setup.exe oficial (winget/msstore so quando sem admin).

    SIMULA POR PADRAO. A primeira rodada so lista o que seria removido e grava o
    plano em plano-newoutlook.json. A segunda, com a chave EXECUTAR, aplica esse plano.

    Opcoes (usuario alvo, limpar tokens, reinstalar) ficam no bloco ESCOPO do script.
.EXAMPLE
    .\Reset-NewOutlook.ps1
    Simula e grava o plano.
.EXAMPLE
    $global:EXECUTAR = $true; .\Reset-NewOutlook.ps1
    Aplica o plano gerado na simulacao (valido por 60 minutos).
.NOTES
    Rodar como administrador para remover o pacote de todos os usuarios.
    Autor: Andre Kittler da Costa
#>
# =====================================================================
# Reset-NewOutlook.ps1  -- remocao total do "Novo Outlook" (olk)
#                              e reinstalacao, rodando como ADMIN
#
# IMPACTO: fecha o olk; remove o pacote Microsoft.OutlookForWindows (todos os
#          usuarios, se admin), o pacote provisionado, a pasta de dados local do
#          usuario alvo (cache/indice/sqlite), os caches de token do WAM/OneAuth
#          e os valores de registro de migracao. Reinstala no fim.
#          O Outlook CLASSICO (perfis, PST, OST, assinaturas) NAO e tocado.
# ALVO:    perfil do usuario logado na maquina (ou $UsuarioAlvo). Lista literal
#          impressa como "ALVOS: n" e gravada em plano-newoutlook.json.
# ROLLBACK: a reinstalacao do proprio script + relogar. E-mail vem do servidor.
#           IRREVERSIVEL: rascunhos nao sincronizados e contas POP locais do
#           Novo Outlook, se existirem.
#
# USO (2 rodadas, como sempre):
#     .\Reset-NewOutlook.ps1                                 # 1) simula e grava o plano
#     $global:EXECUTAR = $true; .\Reset-NewOutlook.ps1       # 2) aplica o plano
# =====================================================================

# ---------- PREAMBULO DE SEGURANCA ----------
# Este script SIMULA por padrao. Para APLICAR, rode tudo num comando so:
#     $global:EXECUTAR = $true; .\Reset-NewOutlook.ps1
$Execute = $false                 # <<< FALHA FECHADA: NUNCA remover esta linha
$__k = Get-Variable EXECUTAR -Scope Global -ValueOnly -ErrorAction SilentlyContinue
$Execute = ($__k -is [bool] -and $__k)
$global:EXECUTAR = $false         # consome a chave

$saida = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.ProviderPath }
try { $__p = Join-Path $saida "w$PID.tmp"; [IO.File]::WriteAllText($__p,'x'); [IO.File]::Delete($__p) }
catch { $saida = $env:TEMP; Write-Host "AVISO: saida redirecionada para $saida" -Foreground Yellow }

# ---------- ESCOPO (edite aqui, sao as unicas chaves) ----------
$UsuarioAlvo       = ''       # vazio = usuario logado na maquina. Ou o nome da pasta em C:\Users
$LimparTokens      = $true    # caches WAM/OneAuth/IdentityCache -> relogin no Office inteiro
$LimparCredenciais = $true    # Gerenciador de Credenciais (so funciona rodando NA conta do usuario)
$Reinstalar        = $true    # reinstala o Novo Outlook no fim
$ValidadePlanoMin  = 60       # plano mais velho que isso e recusado

$PkgName  = 'Microsoft.OutlookForWindows'
$PkgFam   = 'Microsoft.OutlookForWindows_8wekyb3d8bbwe'
$WingetId = '9NRX63209R7B'    # so existe na fonte msstore (costuma estar bloqueada em ambiente corporativo) -> so usado sem admin
$SetupUrl = 'https://go.microsoft.com/fwlink/?linkid=2207851'   # Setup.exe oficial do Novo Outlook
$ts       = Get-Date -f yyyyMMdd-HHmmss
$plano    = Join-Path $saida 'plano-newoutlook.json'

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
$HKU = "Registry::HKEY_USERS\$alvoSid"

Write-Host ""
Write-Host "USUARIO ALVO: $alvoNome  ($alvoSid)" -Foreground Cyan
Write-Host "ADMIN: $isAdmin   RODANDO COMO O PROPRIO USUARIO: $mesmoUsuario" -Foreground Cyan

# ================= FUNCOES DE LEITURA =================
function Get-OlkAppx {
    if ($isAdmin) { Get-AppxPackage -AllUsers -Name $PkgName -ErrorAction SilentlyContinue }
    else          { Get-AppxPackage -Name $PkgName -ErrorAction SilentlyContinue }
}
function Get-OlkProv {
    if ($isAdmin) { Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -eq $PkgName } }
}
function Get-Estado {
    # -User nao aceita SID de conta Entra (S-1-12-1-...): filtra o -AllUsers pelo SID
    if ($isAdmin) { $doUser = @(Get-OlkAppx | Where-Object {
                        @($_.PackageUserInformation | Where-Object { $_.UserSecurityId.Sid -eq $alvoSid -and "$($_.InstallState)" -eq 'Installed' }).Count }) }
    else          { $doUser = Get-OlkAppx }
    [pscustomobject]@{
        OutlookNoUsuario = [bool]$doUser
        Provisionado     = if ($isAdmin) { [bool](@(Get-OlkProv)) } else { 'n/a' }
        PastaPacote      = Test-Path -LiteralPath "$L\Packages\$PkgFam"
        PastaOlk         = Test-Path -LiteralPath "$L\Microsoft\Olk"
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

# 1) processo
$procs = @(Get-Process -Name 'olk' -ErrorAction SilentlyContinue)
if ($procs.Count) { Add-Alvo 1 'Processo' 'Stop-Process -Force' 'olk' "$($procs.Count) instancia(s)" }

# 2) pacote Appx
foreach ($a in @(Get-OlkAppx)) {
    $acao = if ($isAdmin) { 'Remove-AppxPackage -AllUsers' } else { 'Remove-AppxPackage' }
    Add-Alvo 2 'Appx' $acao $a.PackageFullName "versao $($a.Version)"
}

# 3) pacote provisionado (senao volta em perfil novo)
foreach ($pp in @(Get-OlkProv)) { Add-Alvo 3 'AppxProv' 'Remove-AppxProvisionedPackage' $pp.PackageName 'provisionado na imagem' }

# 4) pastas de dados, cache e temporarios
$pastas = @(
    "$L\Packages\$PkgFam"
    "$L\Microsoft\Olk"
    "$L\Microsoft\OutlookForWindows"
    "$L\Temp\Outlook Logging"
)
if ($LimparTokens) {
    $pastas += @(
        "$L\Microsoft\OneAuth"
        "$L\Microsoft\IdentityCache"
        "$L\Microsoft\TokenBroker\Cache"
        "$L\Packages\Microsoft.AAD.BrokerPlugin_cne1h2txyewy\AC\TokenBroker"
    )
}
foreach ($f in $pastas) {
    if (Test-Path -LiteralPath $f) {
        $sz = 0
        try { $sz = [math]::Round((Get-ChildItem -LiteralPath $f -Recurse -Force -ErrorAction SilentlyContinue |
                   Measure-Object Length -Sum).Sum / 1MB, 1) } catch {}
        Add-Alvo 4 'Pasta' 'Remove-Item -Recurse -Force' $f "$sz MB"
    }
}

# 5) valores de registro do Novo Outlook no HKCU DO USUARIO (hive carregada = logado).
#    VALORES, nunca a chave inteira: a chave carrega preferencia do Outlook classico.
$regVals = @(
    @{ K="$HKU\Software\Microsoft\Office\16.0\Outlook\Options\General"; V='DoNewOutlookAutoMigration' }
    @{ K="$HKU\Software\Microsoft\Office\16.0\Outlook\Options\General"; V='NewOutlookAutoMigrationRetryIntervals' }
    @{ K="$HKU\Software\Microsoft\Office\16.0\Outlook\Options\General"; V='HideNewOutlookToggle' }
    @{ K="$HKU\Software\Microsoft\Office\16.0\Outlook\Preferences";     V='NewOutlookMigrationUserSetting' }
    @{ K="$HKU\Software\Microsoft\Office\16.0\Outlook\Preferences";     V='UseNewOutlook' }
)
foreach ($r in $regVals) {
    $existe = $false
    try { $existe = $null -ne (Get-ItemProperty -LiteralPath $r.K -Name $r.V -ErrorAction Stop) } catch {}
    if ($existe) { Add-Alvo 5 'RegValor' 'Remove-ItemProperty' "$($r.K)|$($r.V)" 'valor' }
}

# 6) credenciais (o cofre e por conta: so da pra limpar rodando como o proprio usuario)
if ($LimparCredenciais -and $mesmoUsuario) {
    $creds = @(cmdkey /list 2>$null | Select-String 'Destino:|Target:' |
        ForEach-Object { ($_ -split ':',2)[1].Trim() } |
        Where-Object { $_ -match 'MicrosoftOffice|outlook|office365|OneAuthAccount' })
    foreach ($c in ($creds | Sort-Object -Unique)) { Add-Alvo 6 'Credencial' 'cmdkey /delete' $c 'credencial salva' }
}

# ================= CONTAGEM =================
Write-Host ""
Write-Host "ALVOS: $($alvos.Count)" -Foreground Cyan
if (-not $isAdmin) {
    Write-Host "AVISO: SEM admin - o pacote provisionado fica (volta em perfil novo) e a reinstalacao depende da Store." -Foreground Yellow
}
if ($LimparCredenciais -and -not $mesmoUsuario) {
    Write-Host "INFO: credenciais do cofre puladas (cofre e da conta do usuario, nao da conta admin)." -Foreground Yellow
}
$alvos | Sort-Object Ordem | Format-Table Tipo,Alvo,Detalhe -AutoSize | Out-Host

$antes = Get-Estado

# ================= ACAO =================
$falhas  = 0
$aplicar = @()
$log     = Join-Path $saida "nuke-newoutlook-$ts.csv"

if (-not $Execute) {
    foreach ($a in ($alvos | Sort-Object Ordem)) { Write-Host "SIMULACAO: $($a.Acao) '$($a.Alvo)'" }
    # .ToArray(): @(List[object]) dentro de hashtable quebra no PS 5.1 ("Argument types do not match")
    [pscustomobject]@{ Criado = (Get-Date).ToString('o'); Sid = $alvoSid; Alvos = $alvos.ToArray() } |
        ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $plano -Encoding UTF8
} else {
    # --- gate do plano (R11): so aplica o que foi mostrado na simulacao ---
    if (-not (Test-Path -LiteralPath $plano)) {
        Write-Host "PARE: nao existe plano. Rode primeiro SEM a chave:  .\Reset-NewOutlook.ps1" -Foreground Red
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

    foreach ($a in $aplicar) {
        $res = 'OK'
        try {
            switch ($a.Tipo) {
                'Processo'   { Get-Process -Name $a.Alvo -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction Stop
                               Start-Sleep -Milliseconds 500 }
                'Appx'       { if ($isAdmin) { Remove-AppxPackage -Package $a.Alvo -AllUsers -ErrorAction Stop }
                               else          { Remove-AppxPackage -Package $a.Alvo -ErrorAction Stop } }
                'AppxProv'   { Remove-AppxProvisionedPackage -Online -PackageName $a.Alvo -ErrorAction Stop | Out-Null }
                'Pasta'      { if (Test-Path -LiteralPath $a.Alvo) { Remove-Item -LiteralPath $a.Alvo -Recurse -Force -ErrorAction Stop }
                               else { $res = 'OK (ja tinha sumido)' } }
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
    Write-Host "Reinstalando o Novo Outlook..." -Foreground Cyan
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $ProgressPreference = 'SilentlyContinue'   # escopo do script; deixa o download 10x mais rapido no 5.1

    if ($isAdmin) {
        # Setup.exe oficial, provisiona pra maquina inteira (nao depende da Store nem do winget)
        try {
            $st = Join-Path $env:TEMP "olk-setup-$ts.exe"
            Invoke-WebRequest -Uri $SetupUrl -OutFile $st -UseBasicParsing -ErrorAction Stop
            $pr = Start-Process -FilePath $st -ArgumentList '--provision true --quiet' -Wait -PassThru -ErrorAction Stop
            $instalou = "Setup.exe saiu com $($pr.ExitCode)"
        } catch { $instalou = "Setup.exe falhou: $($_.Exception.Message)" }
    } else {
        $wg = Get-WingetExe
        if ($wg) {
            try {
                & $wg install --id $WingetId --source msstore --silent --disable-interactivity `
                    --accept-package-agreements --accept-source-agreements 2>&1 | Out-Host
                $instalou = "winget saiu com $LASTEXITCODE"
            } catch { $instalou = "winget falhou: $($_.Exception.Message)" }
        } else { $instalou = 'winget indisponivel' }
    }

    Start-Sleep -Seconds 5
    $ok = [bool](@(Get-OlkProv)) -or [bool](@(Get-OlkAppx))
    $instalou = if ($ok) { "OK ($instalou)" } else { "NAO REINSTALOU ($instalou) - pedir pelo Company Portal" }
}

# ============ VERIFICACAO -- ULTIMO BLOCO, NADA DEPOIS ============
$depois = Get-Estado
$tab = @(
  $antes  | Select-Object @{n='Fase';e={'ANTES' }},OutlookNoUsuario,Provisionado,PastaPacote,PastaOlk
  $depois | Select-Object @{n='Fase';e={'DEPOIS'}},OutlookNoUsuario,Provisionado,PastaPacote,PastaOlk
)
Write-Host ""
if (-not $Execute) {
    Write-Host '### MODO SIMULACAO - NADA FOI ALTERADO ###' -Foreground Yellow
    if (Test-Path -LiteralPath $plano) { Write-Host "### Plano gravado: $plano (vale $ValidadePlanoMin min) ###" -Foreground Yellow }
    else { Write-Host '### ATENCAO: o plano NAO foi gravado - a execucao vai recusar ###' -Foreground Red }
    Write-Host '### Rode:  $global:EXECUTAR = $true; .\Reset-NewOutlook.ps1  ###' -Foreground Yellow
} elseif ($falhas -gt 0) {
    Write-Host "### PARCIAL - $falhas de $($aplicar.Count) itens falharam (ver $log) ###" -Foreground Red
} else {
    Write-Host "### APLICADO - $($aplicar.Count) itens. Confira ANTES x DEPOIS ###" -Foreground Green
}
if ($Execute -and $Reinstalar) { Write-Host "Reinstalacao: $instalou" }
if ($Execute)  { Write-Host "Esperado no DEPOIS: Provisionado=True, PastaOlk=False. O pacote entra pro usuario no proximo login." }
$tab | Format-Table -AutoSize | Out-Host
$csv = Join-Path $saida "verificacao-nuke-newoutlook-$ts.csv"
$tab | Export-Csv $csv -NoTypeInformation -Encoding UTF8
if (Test-Path $csv) { Write-Host "Resumo salvo em: $csv" } else { Write-Host "ATENCAO: o CSV NAO foi gravado" -Foreground Red }
