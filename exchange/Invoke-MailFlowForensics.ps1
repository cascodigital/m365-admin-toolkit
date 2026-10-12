<#
.SYNOPSIS
    Investiga reclamacao de "meu email caiu na quarentena" no Exchange Online.
    Interativo por padrao, parametrizavel para automacao.

.DESCRIPTION
    v3.0 - reescrito depois de um caso real de quarentena (10/09/2026), que expos
    a armadilha central desta investigacao:

        QUARENTENA indexa o P2  (From: do cabecalho, o que o usuario ve)
        MESSAGE TRACE indexa o P1 (envelope MAIL FROM / Return-Path)

    Quando o remetente e caixa compartilhada, alias ou automacao, P1 != P2 e
    buscar o trace pelo endereco reclamado devolve ZERO resultados - falso
    negativo com cara de conclusao. Esta versao NUNCA correlaciona por remetente:
    usa MessageId, que e imune a divergencia.

    Segunda armadilha coberta: Status=Delivered no trace NAO significa entrega
    limpa. Mensagem quarentenada e depois liberada aparece como Delivered.
    A prova da retencao esta no ReleaseStatus da quarentena.

.PARAMETER Recipient
    Caixa que reclamou. Opcional se -Sender for informado.

.PARAMETER Sender
    Remetente. Aceita 'user@dominio.com', '@dominio.com' ou 'dominio.com'.
    Mesmo com endereco completo, o filtro de trace opera por DOMINIO (anti-P1/P2).

.EXAMPLE
    .\Invoke-MailFlowForensics.ps1

.EXAMPLE
    .\Invoke-MailFlowForensics.ps1 -Recipient user@contoso.com -Sender @fornecedor.com -Days 7

.NOTES
    Ambiente : Exchange Online
    Requer   : ExchangeOnlineManagement 3.x + PIM ativo
    Autor    : Andre Kittler da Costa
    Versao   : 3.0
#>

[CmdletBinding()]
param(
    [string] $Recipient,
    [string] $Sender,
    [int]    $Days,
    [switch] $Detail,
    [switch] $NonInteractive,
    [Parameter(Mandatory = $true)][string] $UPN,
    [string] $OutRoot                       # default = subpasta ao lado do script
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

# Saida ao lado do script, nao no Desktop.
if (-not $OutRoot) {
    $base = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $OutRoot = Join-Path $base 'MailFlowForensics'
}

#region ---------------------------------------------------------- UI HELPERS
function Say  { param($m, $c = 'Gray') Write-Host $m -ForegroundColor $c }
function Head { param($m) Say ''; Say ('=' * 76) DarkCyan; Say "  $m" Cyan; Say ('=' * 76) DarkCyan }
function Ask {
    param([string]$Question, [string]$Default, [switch]$AllowEmpty)
    while ($true) {
        $p = if ($Default) { "$Question [$Default]" } else { $Question }
        $a = (Read-Host $p).Trim()
        if (-not $a -and $Default)    { return $Default }
        if (-not $a -and $AllowEmpty) { return '' }
        if ($a) { return $a }
        Say '   Resposta vazia.' DarkYellow
    }
}
function Norm-MsgId { param($v) if ($v) { $v.Trim().Trim('<','>').ToLower() } else { '' } }
#endregion

#region ------------------------------------------------- NORMALIZACAO INPUT
function Resolve-SenderInput {
    param([string]$Raw)
    if (-not $Raw) { return $null }
    $v = $Raw.Trim().TrimStart('*').Trim()
    if ($v -match '^[^@\s]+@([^@\s]+\.[^@\s]+)$') {
        return @{ Mode='Address'; Address=$v.ToLower(); Domain=$Matches[1].ToLower(); Pattern="*@$($Matches[1].ToLower())" }
    }
    $d = $v.TrimStart('@').ToLower()
    if ($d -match '^[^@\s]+\.[^@\s]+$') {
        return @{ Mode='Domain'; Address=$null; Domain=$d; Pattern="*@$d" }
    }
    return $null
}
function Test-EmailAddress { param([string]$v) $v -match '^[^@\s]+@[^@\s]+\.[^@\s]+$' }
#endregion

Clear-Host
Head 'MAIL FLOW FORENSICS v3.0  -  quarentena / filtragem inbound'
Say '  Exchange Online' DarkGray
Say '  Correlacao por MessageId (imune a divergencia P1/P2)' DarkGray

#region ------------------------------------------------------ COLETA DE INPUT
if (-not $NonInteractive) {
    Say ''
    Say ' Deixe em BRANCO o que nao souber. Pelo menos um dos dois e obrigatorio.' DarkYellow
    Say ''
    if (-not $Recipient) { $Recipient = Ask 'Email do DESTINATARIO (quem reclamou)' -AllowEmpty }
    if (-not $Sender) {
        Say '   (aceita "user@dominio.com" ou apenas "@dominio.com")' DarkGray
        $Sender = Ask 'Email/dominio do REMETENTE' -AllowEmpty
    }
    if (-not $Days)   { $Days = [int](Ask 'Janela em dias' -Default '7') }
    if (-not $Detail) { $Detail = (Ask 'Puxar detalhe de eventos por mensagem? (s/N)' -Default 'N') -match '^[sSyY]' }
}
if (-not $Days) { $Days = 7 }

if ($Recipient) { $Recipient = $Recipient.Trim() }
$Snd = Resolve-SenderInput $Sender

if (-not $Recipient -and -not $Snd) { throw 'Informe pelo menos o destinatario OU o remetente.' }
if ($Recipient -and -not (Test-EmailAddress $Recipient)) { throw "Destinatario invalido: '$Recipient'" }
if ($Sender -and -not $Snd) { throw "Remetente invalido: '$Sender'. Use 'user@dominio.com' ou '@dominio.com'." }

if ($Snd -and $Snd.Mode -eq 'Address') {
    Say ''
    Say " [i] Endereco completo informado: $($Snd.Address)" Cyan
    Say "     O filtro de trace vai operar por DOMINIO (*@$($Snd.Domain)), nao por" DarkGray
    Say '     endereco. Motivo: trace indexa o envelope (P1), que costuma diferir do' DarkGray
    Say '     From: do cabecalho (P2). Filtrar por endereco produz falso negativo.' DarkGray
}

$End   = (Get-Date).ToUniversalTime()
$Start = $End.AddDays(-$Days)
$Stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$Tag   = if ($Snd) { $Snd.Domain } elseif ($Recipient) { ($Recipient -split '@')[0] } else { 'query' }
$OutDir = Join-Path $OutRoot "$Tag`_$Stamp"
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
try { Start-Transcript -Path (Join-Path $OutDir 'console.log') -Force | Out-Null } catch {}

Head 'PARAMETROS'
Say ("  Destinatario : {0}" -f $(if ($Recipient) { $Recipient } else { '(nao informado)' })) White
Say ("  Remetente    : {0}   [modo {1}]" -f $(if ($Snd) { $Snd.Pattern } else { '(nao informado)' }), $(if ($Snd) { $Snd.Mode } else { 'n/a' })) White
Say ("  Janela UTC   : {0}  ->  {1}   ({2}d)" -f $Start.ToString('u'), $End.ToString('u'), $Days) White
Say ("  Saida        : {0}" -f $OutDir) White
#endregion

#region ---------------------------------------------------------- CONEXAO
Head 'CONEXAO'
if (-not (Get-Module ExchangeOnlineManagement -ListAvailable)) {
    throw 'Modulo ExchangeOnlineManagement ausente. Install-Module ExchangeOnlineManagement -Scope CurrentUser'
}
Import-Module ExchangeOnlineManagement -ErrorAction Stop
$conn = Get-ConnectionInformation -ErrorAction SilentlyContinue |
        Where-Object { $_.State -eq 'Connected' -and $_.TokenStatus -eq 'Active' }
if ($conn) { Say "[+] Sessao ativa: $($conn.UserPrincipalName)" Green }
else       { Say "[*] Conectando como $UPN ..." Yellow
             Connect-ExchangeOnline -UserPrincipalName $UPN -ShowBanner:$false
             Say '[+] Conectado.' Green }
#endregion

#region -------------------------------------------------------- PRE-FLIGHT
Head 'PRE-FLIGHT DE PERMISSAO'
$Can = [ordered]@{ Quarantine=$false; TraceV2=$false; TraceV1=$false; Detail=$false
                   JunkCfg=$false; InboxRule=$false; AllowBlock=$false }
$Denied = @()
function Probe {
    param([string]$Name, [scriptblock]$Test)
    try { & $Test | Out-Null; $script:Can[$Name] = $true; Say "  [OK]     $Name" Green }
    catch { $script:Denied += "$Name -> $($_.Exception.Message)"; Say "  [NEGADO] $Name" Red }
}
Probe 'Quarantine' { Get-QuarantineMessage -PageSize 1 -ErrorAction Stop }
Probe 'TraceV2'    { Get-MessageTraceV2 -StartDate $Start -EndDate $End -ResultSize 1 -ErrorAction Stop }
if (-not $Can.TraceV2) { Probe 'TraceV1' { Get-MessageTrace -StartDate $Start -EndDate $End -PageSize 1 -ErrorAction Stop } }
$Can.Detail = [bool]($Can.TraceV2 -and (Get-Command Get-MessageTraceDetailV2 -ErrorAction SilentlyContinue))
if ($Recipient) {
    Probe 'JunkCfg'   { Get-MailboxJunkEmailConfiguration -Identity $Recipient -ErrorAction Stop }
    Probe 'InboxRule' { Get-InboxRule -Mailbox $Recipient -ErrorAction Stop }
}
Probe 'AllowBlock' { Get-TenantAllowBlockListItems -ListType Sender -ErrorAction Stop }

if (-not $Can.Quarantine -and -not $Can.TraceV2 -and -not $Can.TraceV1) {
    Say ' Sem acesso a nenhuma fonte primaria. Reative o PIM; se persistir, peca' Red
    Say ' "Security Reader" ou "Hygiene Management" no Defender.' Yellow
    $Denied | ForEach-Object { Say "   $_" DarkGray }
    try { Stop-Transcript | Out-Null } catch {}
    throw 'Abortando.'
}
#endregion

#region ------------------------------------------- FASE 1: QUARENTENA (P2)
Head 'FASE 1/5  -  QUARENTENA   (fonte do P2 / header From)'
$Quar = @()
if ($Can.Quarantine) {
    $useDomain = (Get-Command Get-QuarantineMessage).Parameters.ContainsKey('Domain')
    $page = 1
    do {
        $p = @{ StartReceivedDate=$Start; EndReceivedDate=$End; Direction='Inbound'
                PageSize=1000; Page=$page; ErrorAction='Stop' }
        # Sempre por DOMINIO, nunca por endereco: o usuario reporta o P2 e pode
        # existir mais de um P2 no mesmo dominio (foi o caso: 9 remetentes).
        if ($Snd -and $useDomain) { $p.Domain = $Snd.Domain }
        if ($Recipient)           { $p.RecipientAddress = $Recipient }
        try { $batch = Get-QuarantineMessage @p }
        catch {
            if ($p.ContainsKey('Domain')) { Say '    [!] -Domain rejeitado; filtrando no cliente.' DarkYellow
                                            $useDomain=$false; $page=1; $Quar=@(); continue }
            Say "    [!] $($_.Exception.Message)" Red; break
        }
        if ($batch) { $Quar += $batch; Say "    pagina $page -> $($batch.Count)" DarkGray }
        $page++
    } while ($batch -and $batch.Count -eq 1000 -and $page -le 200)

    if ($Snd) { $Quar = @($Quar | Where-Object { $_.SenderAddress -like $Snd.Pattern }) }

    $qHeld = @($Quar | Where-Object { $_.ReleaseStatus -ne 'RELEASED' })
    $qRel  = @($Quar | Where-Object { $_.ReleaseStatus -eq 'RELEASED' })
    Say "[+] $($Quar.Count) em quarentena  |  $($qHeld.Count) ainda retidas  |  $($qRel.Count) ja liberadas" Green
    if ($qRel.Count) {
        Say "    [!] $($qRel.Count) liberacao(oes) manual(is) = falsos positivos ja reconhecidos pela propria organizacao." Yellow
    }
} else { Say '[!] Pulado (sem permissao) - sem isso nao se CONFIRMA quarentena.' Red }
#endregion

#region -------------------------------------- FASE 2: TRACE (por MessageId)
function Get-Trace {
    param([string]$SenderAddress, [string]$RecipientAddress, [string]$MessageId)
    $out = @(); $chunk = 10; $cursor = $End
    while ($cursor -gt $Start) {
        $from = $cursor.AddDays(-$chunk); if ($from -lt $Start) { $from = $Start }
        $p = @{ StartDate=$from; EndDate=$cursor; ErrorAction='Stop' }
        if ($SenderAddress)    { $p.SenderAddress    = $SenderAddress }
        if ($RecipientAddress) { $p.RecipientAddress = $RecipientAddress }
        if ($MessageId)        { $p.MessageId        = $MessageId }
        try {
            $batch = if ($Can.TraceV2) { Get-MessageTraceV2 @p -ResultSize 5000 }
                     else              { Get-MessageTrace   @p -PageSize   5000 }
        } catch { Say "    [!] trace $($from.ToString('MM-dd')): $($_.Exception.Message)" Red; break }
        if ($batch) {
            $out += $batch
            if ($batch.Count -ge 5000) { Say '    [!] bloco saturado (5000). Reduza -Days.' Yellow }
        }
        $cursor = $from
    }
    return $out
}

Head 'FASE 2/5  -  MESSAGE TRACE   (fonte do P1 / envelope sender)'
$Trace = @()
if ($Can.TraceV2 -or $Can.TraceV1) {

    # (a) ANCORA: MessageId vindo da quarentena. Imune a P1 != P2.
    $traceCmd = if ($Can.TraceV2) { 'Get-MessageTraceV2' } else { 'Get-MessageTrace' }
    $canById  = (Get-Command $traceCmd).Parameters.ContainsKey('MessageId')
    if (-not $canById) {
        Say "[!] $traceCmd nao expoe -MessageId nesta versao do modulo." DarkYellow
        Say '    A ancora por MessageId sera pulada; a correlacao vai depender da' DarkYellow
        Say '    varredura por destinatario. Informe -Recipient para nao ficar cego.' DarkYellow
    }

    $ids = @($Quar | ForEach-Object { $_.MessageId } | Where-Object { $_ } | Sort-Object -Unique)
    if ($canById -and $ids.Count) {
        Say "[*] Ancorando por MessageId: $($ids.Count) id(s) unico(s) da quarentena" Yellow
        $i = 0
        foreach ($id in ($ids | Select-Object -First 300)) {
            $i++; Write-Progress -Activity 'Trace por MessageId' -Status "$i/$($ids.Count)" -PercentComplete (100*$i/$ids.Count)
            $Trace += Get-Trace -MessageId $id
        }
        Write-Progress -Activity 'Trace por MessageId' -Completed
        Say "    -> $($Trace.Count) linha(s)" DarkGray
    }

    # (b) VARREDURA: caixa do reclamante, filtrada por DOMINIO (nunca por endereco).
    if ($Recipient) {
        Say "[*] Varrendo a caixa $Recipient (filtro por dominio no cliente)" Yellow
        $all = Get-Trace -RecipientAddress $Recipient
        Say "    $($all.Count) mensagem(ns) totais na janela" DarkGray
        $Trace += if ($Snd) { $all | Where-Object { $_.SenderAddress -like $Snd.Pattern } } else { $all }
    }

    # (c) COMPLEMENTO: remetentes P1 ja observados, para pegar outros destinatarios.
    $p1s = @($Trace | Select-Object -ExpandProperty SenderAddress -Unique | Where-Object { $_ })
    foreach ($s in $p1s) {
        if ($Snd -and $s -notlike $Snd.Pattern) { continue }
        Say "[*] Ampliando por remetente P1: $s" Yellow
        $Trace += Get-Trace -SenderAddress $s
    }

    $Trace = @($Trace | Sort-Object MessageId, RecipientAddress, Received -Unique)
    Say "[+] $($Trace.Count) evento(s) unico(s)." Green
} else { Say '[!] Pulado (sem permissao).' Red }
#endregion

#region ------------------------------------- FASE 3: CORRELACAO E DIVERGENCIA
Head 'FASE 3/5  -  CORRELACAO P1 x P2'

$qIndex = @{}
foreach ($q in $Quar) { $qIndex["$(Norm-MsgId $q.MessageId)|$($q.RecipientAddress.ToLower())"] = $q }

$Corr = foreach ($t in $Trace) {
    $k = "$(Norm-MsgId $t.MessageId)|$($t.RecipientAddress.ToLower())"
    $q = $qIndex[$k]
    $eff = if ($q) {
               if ($q.ReleaseStatus -eq 'RELEASED') { 'QUARANTINED -> released manually' } else { 'QUARANTINED - still held' }
           } elseif ($t.Status -match 'FilteredAsSpam') { 'DELIVERED to Junk (FilteredAsSpam)' }
           elseif ($t.Status -match 'Quarantined')      { 'QUARANTINED' }
           elseif ($t.Status -match 'Failed')           { 'FAILED at transport' }
           else                                          { $t.Status }
    [pscustomobject]@{
        Received             = $t.Received
        EnvelopeSender_P1    = $t.SenderAddress
        HeaderSender_P2      = if ($q) { $q.SenderAddress } else { '' }
        SenderDivergence     = if ($q -and $q.SenderAddress -and ($q.SenderAddress -ne $t.SenderAddress)) { 'YES' } else { '' }
        Recipient            = $t.RecipientAddress
        Subject              = $t.Subject
        TraceStatus          = $t.Status
        QuarantineType       = if ($q) { $q.Type } else { '' }
        QuarantinePolicy     = if ($q) { $q.PolicyName } else { '' }
        ReleaseStatus        = if ($q) { $q.ReleaseStatus } else { '' }
        EffectiveDisposition = $eff
        MessageId            = $t.MessageId
        MessageTraceId       = $t.MessageTraceId
        FromIP               = $t.FromIP
    }
}
$Corr = @($Corr)

# Quarentena sem linha correspondente no trace (permissao ou janela)
$traceKeys = @{}; foreach ($t in $Trace) { $traceKeys["$(Norm-MsgId $t.MessageId)|$($t.RecipientAddress.ToLower())"] = $true }
$Orphans = @($Quar | Where-Object { -not $traceKeys["$(Norm-MsgId $_.MessageId)|$($_.RecipientAddress.ToLower())"] })

$Diverg = @($Corr | Where-Object { $_.SenderDivergence -eq 'YES' })
$PairMap = $Diverg | Group-Object HeaderSender_P2, EnvelopeSender_P1 | ForEach-Object {
    [pscustomobject]@{ HeaderSender_P2 = $_.Group[0].HeaderSender_P2
                       EnvelopeSender_P1 = $_.Group[0].EnvelopeSender_P1
                       Messages = $_.Count }
}

if ($Diverg.Count) {
    Say "[!!] DIVERGENCIA P1/P2 em $($Diverg.Count) mensagem(ns):" Red
    $PairMap | ForEach-Object { Say ("     header {0}  ->  envelope {1}   ({2} msg)" -f $_.HeaderSender_P2, $_.EnvelopeSender_P1, $_.Messages) DarkYellow }
    Say '     Consequencia: allow-list por ENDERECO pode nao funcionar - o filtro' Yellow
    Say '     avalia o envelope. Pedir allow no nivel de DOMINIO.' Yellow
} else { Say '[+] Sem divergencia P1/P2 nas mensagens correlacionadas.' Green }

if ($Orphans.Count) { Say "[!] $($Orphans.Count) item(ns) de quarentena sem linha no trace (ver orphans.csv)." DarkYellow }
#endregion

#region ------------------------------------------------------ FASE 4: DETALHE
Head 'FASE 4/5  -  DETALHE DE EVENTOS'
$Details = @()
if ($Detail -and $Can.Detail -and $Trace.Count) {
    $alvos = @($Corr | Where-Object { $_.EffectiveDisposition -notmatch '^Delivered$' } | Select-Object -First 25)
    if (-not $alvos) { $alvos = @($Corr | Select-Object -First 10) }
    Say "[*] Detalhando $($alvos.Count) mensagem(ns)..." Yellow
    foreach ($m in $alvos) {
        try {
            $Details += Get-MessageTraceDetailV2 -MessageTraceId $m.MessageTraceId -RecipientAddress $m.Recipient -ErrorAction Stop |
                Select-Object @{n='MessageId';e={$m.MessageId}}, @{n='Sender';e={$m.EnvelopeSender_P1}},
                              @{n='Recipient';e={$m.Recipient}}, Date, Event, Action, Detail
        } catch { Say "    [!] $($m.MessageId): $($_.Exception.Message)" DarkGray }
    }
    Say "[+] $($Details.Count) evento(s)." Green
} elseif ($Detail) { Say '[!] Indisponivel.' DarkYellow } else { Say '[-] Nao solicitado (-Detail).' DarkGray }
#endregion

#region -------------------------------------- FASE 5: USUARIO E TENANT
Head 'FASE 5/5  -  LADO DO USUARIO E TENANT'
$JunkCfg=$null; $Rules=@(); $AllowBlock=@(); $blocked=@(); $trusted=@(); $suspeitas=@()

if ($Can.JunkCfg) {
    $JunkCfg = Get-MailboxJunkEmailConfiguration -Identity $Recipient
    $blocked = @($JunkCfg.BlockedSendersAndDomains | Where-Object { -not $Snd -or $_ -like "*$($Snd.Domain)*" })
    $trusted = @($JunkCfg.TrustedSendersAndDomains | Where-Object { -not $Snd -or $_ -like "*$($Snd.Domain)*" })
    Say "  Junk filter ..........: $($JunkCfg.Enabled)" White
    Say "  Blocked senders ......: $($JunkCfg.BlockedSendersAndDomains.Count)" White
    if ($blocked) { Say "  [!!] REMETENTE NA BLOCKED LIST DO USUARIO: $($blocked -join ', ')" Red }
    if ($trusted) { Say "  [+]  Ja consta como confiavel: $($trusted -join ', ')" Green }
}
if ($Can.InboxRule) {
    $Rules = @(Get-InboxRule -Mailbox $Recipient | Where-Object { $_.Enabled })
    $suspeitas = @($Rules | Where-Object {
        $t = ($_.Description + $_.From + $_.SubjectContainsWords + $_.FromAddressContainsWords) -join ' '
        (-not $Snd -or $t -like "*$($Snd.Domain)*") -and ($_.DeleteMessage -or $_.MoveToFolder -or $_.MarkAsRead)
    })
    Say "  Inbox rules ativas ...: $($Rules.Count)" White
    if ($suspeitas) { Say '  [!!] REGRA PODE ESTAR DESVIANDO:' Red
                      $suspeitas | ForEach-Object { Say "       - $($_.Name)" DarkYellow } }
}
if ($Can.AllowBlock -and $Snd) {
    foreach ($lt in 'Sender','Url') {
        try { $AllowBlock += Get-TenantAllowBlockListItems -ListType $lt -ErrorAction Stop |
                             Where-Object { $_.Value -like "*$($Snd.Domain)*" } } catch {}
    }
    if ($AllowBlock) { Say '  [!] Entrada existente na Tenant Allow/Block List:' Yellow
                       $AllowBlock | ForEach-Object { Say "       $($_.Value)  Action=$($_.Action)  Exp=$($_.ExpirationDate)" DarkYellow } }
    else             { Say '  Tenant Allow/Block ...: nenhuma entrada para o dominio' White }
}
#endregion

#region ---------------------------------------------------------- VEREDITO
Head 'VEREDITO'
$nQuar = $Quar.Count
$nHeld = @($Quar | Where-Object { $_.ReleaseStatus -ne 'RELEASED' }).Count
$nRel  = @($Quar | Where-Object { $_.ReleaseStatus -eq 'RELEASED' }).Count
$nJunk = @($Corr | Where-Object { $_.EffectiveDisposition -match 'Junk' }).Count
$nFail = @($Corr | Where-Object { $_.EffectiveDisposition -match 'FAILED' }).Count
$nClean= @($Corr | Where-Object { $_.EffectiveDisposition -eq 'Delivered' }).Count
$fpRate = if ($nQuar) { [math]::Round(100*$nRel/$nQuar) } else { 0 }

$Verdict, $Action =
    if ($nQuar -gt 0) {
        "CONFIRMED - SERVER-SIDE QUARANTINE. $nQuar message(s); $nHeld still held, $nRel manually released ($fpRate% false-positive rate).",
        $(if ($Diverg.Count) { 'Escalate to Cyber. Request DOMAIN-level allow (envelope sender differs from header sender - address-level allow may not match).' }
          else               { 'Escalate to Cyber with quarantine.csv + correlated.csv as evidence.' })
    }
    elseif ($nFail -gt 0)  { "CONFIRMED - TRANSPORT FAILURE. $nFail message(s) failed.", 'Nao e quarentena: investigar connector / edge / DBEB antes do Cyber.' }
    elseif ($nJunk -gt 0)  { "CONFIRMED - JUNK FOLDER. $nJunk message(s) FilteredAsSpam.", 'Entregue na pasta Lixo, nao retido. Fix e anti-spam policy, nao release.' }
    elseif ($blocked)      { 'NOT CONFIRMED - USER BLOCKED SENDER LIST.', 'Causa do proprio usuario. Remover da Blocked Senders. Nao acionar Cyber.' }
    elseif ($suspeitas)    { 'NOT CONFIRMED - INBOX RULE.', 'Regra do Outlook do usuario. Corrigir com ele.' }
    elseif ($nClean -gt 0) { "NOT CONFIRMED - DELIVERED. $nClean message(s) entregues.", 'Mail flow limpo. Verificar pasta / Focused Inbox / sync do cliente.' }
    elseif ($Trace.Count -eq 0) { 'NO DATA - nada desse remetente alcancou o tenant na janela.', 'Confirmar o dominio, ampliar -Days, ou suspeitar de bloqueio pre-trace.' }
    else                   { 'INCONCLUSIVE.', 'Rodar com -Detail para ver a cadeia de eventos.' }

$cor = if ($Verdict -like 'CONFIRMED*') { 'Red' } elseif ($Verdict -like 'NOT*') { 'Yellow' } else { 'DarkYellow' }
Say ''; Say "  $Verdict" $cor; Say ''; Say "  ACAO: $Action" Cyan; Say ''
Say '  ----------------------------------------------------------------' DarkGray
Say ("  Quarentena total .....: {0}   (retidas {1} / liberadas {2})" -f $nQuar,$nHeld,$nRel) White
Say ("  Trace correlacionado .: {0}" -f $Corr.Count) White
Say ("  Divergencia P1/P2 ....: {0}" -f $Diverg.Count) White
$Corr | Group-Object EffectiveDisposition | Sort-Object Count -Descending |
    ForEach-Object { Say ("      {0,-38} {1}" -f $_.Name, $_.Count) DarkGray }
#endregion

#region ------------------------------------------------------------ EXPORT
Head 'EXPORT'
function Dump { param($data,$file,$sel)
    if ($data -and $data.Count) {
        if ($sel) { $data | Select-Object $sel | Export-Csv (Join-Path $OutDir $file) -NoTypeInformation -Encoding UTF8 }
        else      { $data | Export-Csv (Join-Path $OutDir $file) -NoTypeInformation -Encoding UTF8 }
        Say "  [+] $file  ($($data.Count))" Green
    }
}
Dump $Corr 'correlated.csv'
Dump $Quar 'quarantine.csv' @('ReceivedTime','SenderAddress','RecipientAddress','Subject','Type','QuarantineTypes','PolicyName','PolicyType','Expires','Released','ReleaseStatus','MessageId','Identity')
Dump $Trace 'messagetrace.csv' @('Received','SenderAddress','RecipientAddress','Subject','Status','FromIP','ToIP','Size','MessageId','MessageTraceId')
Dump $Orphans 'orphans.csv' @('ReceivedTime','SenderAddress','RecipientAddress','Subject','ReleaseStatus','MessageId')
Dump $Details 'trace-detail.csv'
Dump $Rules 'inboxrules.csv' @('Name','Enabled','Priority','Description')
Dump $PairMap 'sender-divergence.csv'

$sb = [System.Text.StringBuilder]::new(); $n = { param($t) [void]$sb.AppendLine($t) }
& $n 'MAIL FLOW EVIDENCE PACKAGE'; & $n '=========================='; & $n ''
& $n "Recipient .......: $(if ($Recipient) { $Recipient } else { '(not specified)' })"
& $n "Sender ..........: $(if ($Snd) { $Snd.Pattern } else { '(not specified)' })"
& $n "Window (UTC) ....: $($Start.ToString('u'))  ->  $($End.ToString('u'))   ($Days days)"
& $n "Tenant ..........: $((Get-AcceptedDomain | Where-Object Default).DomainName)"
& $n "Collected by ....: $UPN"
& $n "Collected at ....: $((Get-Date).ToUniversalTime().ToString('u')) UTC"
& $n ''; & $n "VERDICT: $Verdict"; & $n "ACTION : $Action"; & $n ''
& $n 'QUARANTINE (server-side hold - message never reached the mailbox)'
& $n '-----------------------------------------------------------------'
& $n "  Total .................: $nQuar"
& $n "  Still held ............: $nHeld"
& $n "  Manually released .....: $nRel   ($fpRate% of total = confirmed false positives)"
$Quar | Group-Object Type       | Sort-Object Count -Descending | ForEach-Object { & $n ("  type/{0,-24} {1}" -f $_.Name,$_.Count) }
$Quar | Group-Object PolicyName | Sort-Object Count -Descending | ForEach-Object { & $n ("  policy/{0,-22} {1}" -f $_.Name,$_.Count) }
& $n ''
& $n 'EFFECTIVE DISPOSITION (quarantine x trace, correlated by MessageId)'
& $n '-------------------------------------------------------------------'
$Corr | Group-Object EffectiveDisposition | Sort-Object Count -Descending | ForEach-Object { & $n ("  {0,-40} {1}" -f $_.Name,$_.Count) }
& $n ''
& $n 'SENDER IDENTITY (P1 envelope vs P2 header)'
& $n '------------------------------------------'
if ($PairMap) {
    & $n '  DIVERGENCE DETECTED - the filter evaluates the ENVELOPE sender:'
    $PairMap | ForEach-Object { & $n "    header From: $($_.HeaderSender_P2)   ->   envelope: $($_.EnvelopeSender_P1)   ($($_.Messages) msg)" }
    & $n '  => Address-level allow entries may NOT match. Request DOMAIN-level allow.'
} else { & $n '  No divergence detected.' }
& $n ''
& $n 'AFFECTED RECIPIENTS'; & $n '-------------------'
$Quar | Group-Object RecipientAddress | Sort-Object Count -Descending | ForEach-Object { & $n ("  {0,-40} {1}" -f $_.Name,$_.Count) }
& $n ''
& $n 'OBSERVED SENDERS (header From)'; & $n '------------------------------'
$Quar | Group-Object SenderAddress | Sort-Object Count -Descending | ForEach-Object { & $n ("  {0,-46} {1}" -f $_.Name,$_.Count) }
& $n ''
& $n 'RECIPIENT-SIDE CHECKS'; & $n '---------------------'
if ($JunkCfg) {
    & $n "  Junk filter enabled ..: $($JunkCfg.Enabled)"
    & $n "  Blocked senders ......: $($JunkCfg.BlockedSendersAndDomains.Count)"
    if ($blocked) { & $n "  >> SENDER ON USER'S BLOCKED LIST: $($blocked -join ', ')" }
    if ($trusted) { & $n "  >> Sender on user's trusted list: $($trusted -join ', ')" }
} else { & $n '  (not collected)' }
if ($Rules)     { & $n "  Enabled inbox rules ..: $($Rules.Count)" }
if ($suspeitas) { $suspeitas | ForEach-Object { & $n "  >> SUSPECT RULE: $($_.Name)" } }
& $n ''
& $n 'TENANT ALLOW/BLOCK LIST'; & $n '-----------------------'
if ($AllowBlock) { $AllowBlock | ForEach-Object { & $n "  $($_.Value)  Action=$($_.Action)  Expires=$($_.ExpirationDate)" } }
else             { & $n '  (no entry for this domain)' }
& $n ''
if ($Denied) {
    & $n 'DATA SOURCES UNAVAILABLE (permission)'; & $n '-------------------------------------'
    $Denied | ForEach-Object { & $n "  $_" }
    & $n '  NOTE: verdict is based only on the sources above.'; & $n ''
}
& $n 'METHOD NOTES'; & $n '------------'
& $n '  * Quarantine reports the P2 (header From); message trace reports the P1'
& $n '    (envelope MAIL FROM). Correlation is done by MessageId, not by sender.'
& $n '  * Trace Status=Delivered does NOT mean clean delivery: a quarantined'
& $n '    message that was released later shows as Delivered. See ReleaseStatus.'
& $n ''
& $n 'Attachments: correlated.csv, quarantine.csv, messagetrace.csv,'
& $n '             sender-divergence.csv, orphans.csv, console.log'
$sb.ToString() | Set-Content (Join-Path $OutDir 'SUMMARY.txt') -Encoding UTF8
Say '  [+] SUMMARY.txt' Green

#region ------------------------------------------------------- JIRA.TXT
Head 'TICKET (jira.txt)'

$isServerSide = $Verdict -like 'CONFIRMED*'
$isUserSide   = ($blocked -and $blocked.Count) -or ($suspeitas -and $suspeitas.Count)

$who = if ($Recipient) {
           $lp = ($Recipient -split '@')[0]
           (($lp -split '[._-]') | Where-Object { $_ } |
             ForEach-Object { $_.Substring(0,1).ToUpper() + $_.Substring(1).ToLower() }) -join ' '
       } else { '' }
$first = if ($who) { ($who -split ' ')[0] } else { 'there' }

$dom     = if ($Snd) { $Snd.Domain } else { 'the reported sender' }
$winTxt  = "{0}-{1} {2} (UTC)" -f $Start.ToString('dd'), $End.ToString('dd MMM yyyy'), ''
$winTxt  = "$($Start.ToString('dd MMM')) - $($End.ToString('dd MMM yyyy')), UTC"
$types   = (($Quar | Group-Object Type | Sort-Object Count -Descending |
             ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ', ')
$pol     = (($Quar | Group-Object PolicyName | Sort-Object Count -Descending |
             Select-Object -First 1).Name)
$polType = ($Quar | Select-Object -First 1).PolicyType
$nSnd    = @($Quar | Select-Object -ExpandProperty SenderAddress -Unique).Count
$rcpts   = @($Quar | Group-Object RecipientAddress | Sort-Object Count -Descending)
$rcptTxt = (($rcpts | Select-Object -First 6 | ForEach-Object { $_.Name }) -join ', ')
if ($rcpts.Count -gt 6) { $rcptTxt += " (+$($rcpts.Count - 6) more)" }
$oldest  = ($Quar | Where-Object { $_.ReleaseStatus -ne 'RELEASED' } |
            Sort-Object { [datetime]$_.ReceivedTime } | Select-Object -First 1)
$files   = @('SUMMARY.txt','correlated.csv') +
           $(if ($Quar.Count)    { 'quarantine.csv' }) +
           $(if ($PairMap)       { 'sender-divergence.csv' }) | Where-Object { $_ }

# ---------------------------------------------------------------- INTERNAL
$i = [System.Text.StringBuilder]::new(); $w = { param($t) [void]$i.AppendLine($t) }

& $w "$dom - inbound mail flow review"
& $w ''
& $w "Findings from a $Days-day review ($winTxt), tenant $((Get-AcceptedDomain | Where-Object Default).DomainName):"
& $w ''
if ($nQuar -gt 0) {
    & $w "- $nQuar inbound message(s) from $dom quarantined."
    & $w "- Verdict: $types."
    & $w "- Policy: `"$pol`" ($polType), action = quarantine. Server-side hold,"
    & $w "  not Junk-folder filing."
    & $w "- $nSnd distinct sender address(es); $($rcpts.Count) recipient(s) affected: $rcptTxt."
    if ($nRel -gt 0) {
        & $w "- $nRel of $nQuar ($fpRate%) manually released by staff. $nHeld still held$(if ($oldest) { ", oldest since $([datetime]$oldest.ReceivedTime | Get-Date -Format 'dd MMM')" })."
    } else {
        & $w "- $nHeld still held$(if ($oldest) { ", oldest since $([datetime]$oldest.ReceivedTime | Get-Date -Format 'dd MMM')" })."
    }
}
if ($nJunk -gt 0)  { & $w "- $nJunk message(s) delivered to the Junk folder (FilteredAsSpam), not quarantined." }
if ($nFail -gt 0)  { & $w "- $nFail message(s) failed at transport." }
if ($nClean -gt 0) { & $w "- $nClean message(s) delivered normally." }
if ($nQuar -eq 0 -and $Trace.Count -eq 0) { & $w "- No inbound mail from $dom reached the tenant in the window." }
if ($blocked)   { & $w "- Sender is present on the recipient's own Blocked Senders list: $($blocked -join ', ')." }
if ($suspeitas) { & $w "- Recipient has an enabled inbox rule matching this sender: $(($suspeitas | ForEach-Object { $_.Name }) -join ', ')." }
if ($AllowBlock) { & $w "- Existing Tenant Allow/Block entry: $(($AllowBlock | ForEach-Object { "$($_.Value) [$($_.Action)]" }) -join ', ')." }

if ($PairMap) {
    $ex = $PairMap | Sort-Object Messages -Descending | Select-Object -First 1
    $exId = ($Diverg | Where-Object { $_.HeaderSender_P2 -eq $ex.HeaderSender_P2 } | Select-Object -First 1).MessageId
    & $w ''
    & $w 'One detail relevant to any remediation scope: header From and envelope'
    & $w "sender diverge. On Message-ID"
    & $w "$exId"
    & $w "the header From is $($ex.HeaderSender_P2) while the envelope sender is"
    & $w "$($ex.EnvelopeSender_P1). Message trace matches on the envelope, which is"
    & $w 'why a search on the reported address returns nothing.'
}

& $w ''
& $w "Attachments: $($files -join ', ')."
& $w 'Method: Get-QuarantineMessage + Get-MessageTrace, correlated by Message-ID.'
& $w 'Timestamps UTC.'
if ($isServerSide -and -not $isUserSide) { & $w ''; & $w 'Handing over to Cyber for assessment.' }

# ---------------------------------------------------------------- CUSTOMER
$c = [System.Text.StringBuilder]::new(); $v = { param($t) [void]$c.AppendLine($t) }

& $v "Hi $first,"
& $v ''
& $v "I've reviewed the mail flow for $dom over the last $Days days. Here's what I found."
& $v ''

if ($nQuar -gt 0) {
    & $v "$nQuar message(s) from $dom were held in the server-side quarantine and"
    & $v 'classified as spam by our mail filter. They were stopped before reaching'
    & $v "your mailbox, which is why they weren't in your Junk folder either."
    & $v ''
    if ($nRel -gt 0 -and $nHeld -gt 0) {
        & $v "$nRel of them had already been released manually, so those should be in"
        & $v "your mailbox now. The remaining $nHeld are still being held - I've released"
        & $v 'nothing on my own; please let me know if you need them delivered.'
    } elseif ($nRel -gt 0) {
        & $v "All $nRel of them had already been released manually, so they should be in"
        & $v 'your mailbox now.'
    } else {
        & $v "All $nHeld are still being held. I have not released anything yet - let me"
        & $v 'know if you need them delivered while the underlying issue is reviewed.'
    }
    & $v ''
    & $v 'Releasing individual messages only fixes them one at a time, so I have'
    & $v 'passed the full findings to our Cyber Security team, who own the mail'
    & $v 'filtering policy. Any permanent change to how this sender is treated'
    & $v 'will come from them.'
}
elseif ($nJunk -gt 0) {
    & $v "$nJunk message(s) from $dom were delivered, but our mail filter classified"
    & $v 'them as spam and filed them into your Junk Email folder. They are not'
    & $v 'blocked - please check that folder.'
    & $v ''
    & $v 'I have passed the findings to our Cyber Security team, who own the mail'
    & $v 'filtering policy, so this can be looked at properly.'
}
elseif ($nFail -gt 0) {
    & $v "$nFail message(s) from $dom failed to be accepted by our mail system and"
    & $v 'never reached the mailbox. This is a delivery failure rather than a spam'
    & $v 'classification.'
    & $v ''
    & $v 'I have passed the findings to our Cyber Security team for assessment.'
}
elseif ($blocked) {
    & $v "The messages are not being blocked by our mail filter. $dom is on the"
    & $v "Blocked Senders list inside your own mailbox, which is what's diverting"
    & $v 'them.'
    & $v ''
    & $v 'This one is fixable directly on your account - let me know and I can'
    & $v 'remove that entry for you. No security escalation is needed.'
}
elseif ($suspeitas) {
    & $v 'The messages are not being blocked by our mail filter. There is a rule in'
    & $v 'your Outlook that matches this sender and is moving the messages.'
    & $v ''
    & $v 'I can walk you through removing it, or remove it for you if you prefer.'
    & $v 'No security escalation is needed.'
}
elseif ($nClean -gt 0) {
    & $v "$nClean message(s) from $dom were delivered to your mailbox normally in"
    & $v 'this period. Nothing was quarantined or filtered.'
    & $v ''
    & $v "If you still can't see them, it's most likely a folder, rule or Focused"
    & $v 'Inbox issue on the client side - happy to take a look with you.'
}
else {
    & $v "I could not find any inbound mail from $dom reaching us in the last"
    & $v "$Days days - nothing quarantined and nothing delivered."
    & $v ''
    & $v 'Could you confirm the exact sender address, or forward me one of the'
    & $v 'messages they say was sent? That will let me widen the search correctly.'
}

& $v ''
& $v 'Best regards,'

# ---------------------------------------------------------------- WRITE
$j = [System.Text.StringBuilder]::new()
[void]$j.AppendLine('=============== INTERNAL COMMENT ==============='); [void]$j.AppendLine('')
[void]$j.Append($i.ToString()); [void]$j.AppendLine('')
[void]$j.AppendLine('=============== CUSTOMER COMMENT ==============='); [void]$j.AppendLine('')
[void]$j.Append($c.ToString())
$j.ToString() | Set-Content (Join-Path $OutDir 'jira.txt') -Encoding UTF8
Say '  [+] jira.txt   (internal + customer, ready to paste)' Green
Say ''
Write-Host $j.ToString() -ForegroundColor White
#endregion

try { Stop-Transcript | Out-Null } catch {}
Head 'FIM'
Say "  Evidencia em: $OutDir" Cyan; Say ''
if (-not $NonInteractive) {
    if ((Ask 'Abrir a pasta? (S/n)' -Default 'S') -match '^[sSyY]') { Start-Process explorer.exe $OutDir }
}
