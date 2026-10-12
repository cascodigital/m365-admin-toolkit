<#
.SYNOPSIS
    Auditoria completa de inbox rules: regras ocultas, sombreamento (regra morta),
    corrupcao, encaminhamento externo e forwarding a nivel de mailbox.

.DESCRIPTION
    Escrito a partir de um caso real (18/09/2026), onde o script antigo devolvia
    "25 rules, none of them forward" -- uma frase verdadeira e inutil que escondia
    tres problemas reais.

    O QUE ELE ACHA QUE OS OUTROS NAO ACHAVAM:

    [A] REGRA MORTA POR SOMBREAMENTO  <<< o achado principal
        Duas regras com a MESMA condicao, a primeira com 'stop processing more rules'.
        A segunda nunca executa. No caso real: prioridade 18 e 22 casavam ambas com
        o mesmo remetente noreply de um fornecedor e mandavam para pastas DIFERENTES. A 22 era codigo
        morto havia meses e nenhum humano viu, lendo a lista com o olho.
        O script compara ASSINATURA DE CONDICAO, nao o texto da Description.

    [B] REGRA OCULTA DE USUARIO
        Get-InboxRule sem -IncludeHidden esconde regras que RODAM. No caso real havia
        uma regra de usuario em prioridade 1, Enabled, com stop --
        invisivel no dialogo Rules and Alerts do Outlook. A usuaria nao conseguia ver,
        editar nem desligar. Junk E-mail Rule e Microsoft.Exchange.OOF.* sao de sistema
        e o script os separa como ruido.

    [C] REGRA QUE APAGA AVISO DO EXCHANGE
        Regra que deleta mensagem vinda de MicrosoftExchange<32hex>@ mata NDR, aviso de
        cota e notificacao de sistema. O usuario fica cego justamente para o canal que
        explicaria o problema dele. Encontrado em producao, prioridade 2.

    [D] REGRAS CORROMPIDAS
        Falham no parse e sao PULADAS do output -- o cmdlet so emite WARNING. Por isso
        InError e sempre False: as ruins nem viram objeto. Contar warning, nao linha.

    [E] ENCAMINHAMENTO EXTERNO, as duas metades
        Inbox rule E forwarding a nivel de mailbox, ambos avaliados contra os accepted
        domains do tenant (Get-AcceptedDomain). O script antigo so imprimia o de mailbox sem avaliar,
        e a lista fixa de dominios estava incompleta -- marcava dominio interno como externo.

    ARMADILHA DOCUMENTADA AQUI PARA NAO SER REDESCOBERTA:
    >> O campo Priority NAO e estavel entre a chamada com e sem -IncludeHidden. <<
    Medido em 18/09/2026: a mesma regra aparece como Priority 3 sem hidden e 4 com;
    outra como 6 sem e 5 com. Comparar prioridade entre as duas listagens produz
    "buracos" fantasmas e uma hipotese inteiramente falsa (eu cai nessa).
    A chave estavel e RuleIdentity. Toda analise de ordem aqui roda sobre UMA
    listagem so -- a completa, com -IncludeHidden.

.NOTES
    AUDITORIA -- R1: zero cmdlet de escrita no tenant. Sem gate $Execute porque nao ha
    nada para gatear. Pode colar no console sem medo.

    PRE-REQ: PIM "Exchange Recipient Administrator" ativo ANTES de conectar. Sem ele,
    Get-Mailbox mente dizendo "object couldn't be found" -- 404 falso, e permissao.

.EXAMPLE
    .\Audit-InboxRules.ps1 -MailboxAddress "user@contoso.com"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$MailboxAddress,

    [Parameter(Mandatory = $true)]
    [string]$AdminUPN,

    [Parameter(Mandatory = $false)]
    [string]$OutFolder
)

# Dominios internos: preenchidos com Get-AcceptedDomain logo apos conectar (sem lista fixa).
$TrustedDomains = @()

# regras de sistema que sao ocultas em TODA caixa -- ruido, nao achado
$SystemHiddenPattern = '^Junk E-mail Rule$|^Microsoft\.Exchange\.OOF\.'

# caixa de sistema do Exchange (NDR, cota, notificacoes)
$SystemMailboxPattern = 'MicrosoftExchange[0-9a-f]{32}@'

# ---------------------------------------------------------------------------
# 0. Pasta de saida -- PROVA de escrita, nao Test-Path.
#    Fallback e TEMP, nao Desktop: Desktop pode estar redirecionado por OneDrive
#    KFM, e CSV com PII nao deve sincronizar.
# ---------------------------------------------------------------------------
$saida = if ($OutFolder) { $OutFolder } elseif ($PSScriptRoot) { $PSScriptRoot } else { $PWD.ProviderPath }
try { $__p = Join-Path $saida "w$PID.tmp"; [IO.File]::WriteAllText($__p,'x'); [IO.File]::Delete($__p) }
catch { $saida = $env:TEMP; Write-Host "AVISO: saida redirecionada para $saida" -ForegroundColor Yellow }

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$slug  = ($MailboxAddress.Split('@')[0]) -replace '[^\w\.\-]','_'

# ---------------------------------------------------------------------------
# 1. Conexao
# ---------------------------------------------------------------------------
try {
    if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
        Write-Host "Conectando ao EXO como $AdminUPN ..." -ForegroundColor Cyan
        Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false
    } else { Write-Host "Sessao EXO ja ativa." -ForegroundColor Green }
    $TrustedDomains = @(Get-AcceptedDomain | ForEach-Object { [string]$_.DomainName })
    Write-Host "Dominios internos: $($TrustedDomains -join ', ')" -ForegroundColor DarkGray
}
catch {
    Write-Host "FALHA na conexao: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "'Error Acquiring Token' + RuntimeBroker = WAM broker quebrado no CLIENTE," -ForegroundColor Yellow
    Write-Host "nao e PIM nem role -- o crash acontece antes de existir sessao." -ForegroundColor Yellow
    return
}

try { $mailbox = Get-Mailbox -Identity $MailboxAddress -ErrorAction Stop }
catch {
    Write-Host "Mailbox nao encontrada: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "ATENCAO: 'couldn't be found' quase sempre e PIM inativo disfarcado de 404." -ForegroundColor Yellow
    return
}

Write-Host "`nMailbox: $($mailbox.DisplayName) <$($mailbox.PrimarySmtpAddress)>" -ForegroundColor Green
Write-Host "Tipo: $($mailbox.RecipientTypeDetails) | Saida: $saida" -ForegroundColor DarkGray

$findings = [System.Collections.Generic.List[object]]::new()
function Add-Finding {
    param($Severidade,$Origem,$Item,$Detalhe,$Acao)
    $findings.Add([PSCustomObject]@{
        Mailbox=$MailboxAddress; Severidade=$Severidade; Origem=$Origem
        Item=$Item; Detalhe=$Detalhe; Acao=$Acao
    })
}

# ---------------------------------------------------------------------------
# 2. Assinaturas -- comparar CONDICAO, nao a prosa da Description
#    (Description e texto localizado e muda de forma sem a regra mudar)
# ---------------------------------------------------------------------------
$CondProps = @(
    'From','SentTo','FromAddressContainsWords','RecipientAddressContainsWords'
    'SubjectContainsWords','SubjectOrBodyContainsWords','BodyContainsWords'
    'HeaderContainsWords','MessageTypeMatches','SentOnlyToMe','MyNameInToBox'
    'MyNameInCcBox','MyNameInToOrCcBox','MyNameNotInToBox','HasAttachment'
    'WithImportance','WithSensitivity','FromSubscription','HasClassification'
    'FlaggedForAction','WithinSizeRangeMinimum','WithinSizeRangeMaximum'
    'ReceivedAfterDate','ReceivedBeforeDate'
    'ExceptIfFrom','ExceptIfSentTo','ExceptIfSubjectContainsWords'
)
$ActionProps = @(
    'MoveToFolder','CopyToFolder','DeleteMessage','PermanentDelete'
    'ForwardTo','ForwardAsAttachmentTo','RedirectTo','SendTextMessageNotificationTo'
    'MarkAsRead','MarkImportance','ApplyCategory','ApplySystemCategory'
)

function Get-Signature {
    param($Rule,[string[]]$Props)
    $parts = foreach ($p in $Props) {
        $pr = $Rule.PSObject.Properties[$p]
        if (-not $pr -or $null -eq $pr.Value) { continue }
        if ($pr.Value -is [bool] -and -not $pr.Value) { continue }   # $false nao e condicao
        $vals = @($pr.Value | ForEach-Object { "$_".Trim() } | Where-Object { $_ }) | Sort-Object
        if ($vals.Count -eq 0) { continue }
        "$p=$($vals -join '|')"
    }
    ($parts | Sort-Object) -join ' && '
}

# ---------------------------------------------------------------------------
# 3. Leitura -- as duas listagens. Warnings capturados (regra corrompida).
#    Sem -ResultSize: o sentinel Unlimited e rejeitado por algumas sessoes EXO.
# ---------------------------------------------------------------------------
$wVis = @()
try { $visiveis = @(Get-InboxRule -Mailbox $MailboxAddress -WarningVariable wVis -WarningAction SilentlyContinue -ErrorAction Stop) }
catch { Write-Host "Falha ao ler inbox rules: $($_.Exception.Message)" -ForegroundColor Red; return }

$temHidden = (Get-Command Get-InboxRule).Parameters.Keys -contains 'IncludeHidden'
$wAll = @()
if ($temHidden) {
    try { $todas = @(Get-InboxRule -Mailbox $MailboxAddress -IncludeHidden -WarningVariable wAll -WarningAction SilentlyContinue -ErrorAction Stop) }
    catch { $todas = $visiveis; $temHidden = $false }
} else { $todas = $visiveis }

$corruptas = [Math]::Max(@($wVis).Count, @($wAll).Count)

# chave estavel: RuleIdentity. NUNCA Priority (instavel entre as listagens).
$idsVis  = @($visiveis | ForEach-Object { "$($_.RuleIdentity)" })
$ocultas = @($todas | Where-Object { "$($_.RuleIdentity)" -notin $idsVis })

Write-Host "`n=== INVENTARIO ===" -ForegroundColor Yellow
Write-Host ("Visiveis no Outlook ..: {0}" -f $visiveis.Count)
Write-Host ("Total com ocultas ....: {0}{1}" -f $todas.Count, $(if(-not $temHidden){' (sessao nao expoe -IncludeHidden)'}else{''}))
Write-Host ("Ocultas ..............: {0}" -f $ocultas.Count)
Write-Host ("Corrompidas ..........: {0}" -f $corruptas) -ForegroundColor $(if($corruptas -gt 0){'Red'}else{'Green'})

foreach ($w in @($wAll)) {
    Add-Finding 'CORROMPIDA' 'InboxRule' "$w" 'Pulada pelo parser -- conteudo e destino desconhecidos.' `
        'Rodar Remove-CorruptInboxRules.ps1 -DryRun primeiro.'
}

# ---------------------------------------------------------------------------
# 4. Tabela de trabalho
# ---------------------------------------------------------------------------
$map = foreach ($rule in $todas) {
    $rid    = "$($rule.RuleIdentity)"
    $oculta = $rid -notin $idsVis
    [PSCustomObject]@{
        Ordem        = [int]$rule.Priority
        Nome         = "$($rule.Name)"
        Oculta       = $oculta
        Sistema      = ($oculta -and "$($rule.Name)" -match $SystemHiddenPattern)
        Enabled      = $rule.Enabled
        Stop         = [bool]$rule.StopProcessingRules
        CondSig      = Get-Signature $rule $CondProps
        ActSig       = Get-Signature $rule $ActionProps
        MoveToFolder = "$($rule.MoveToFolder)"
        ForwardTo             = if ($rule.ForwardTo) { (@($rule.ForwardTo|%{"$_"}) -join '; ') } else { '' }
        ForwardAsAttachmentTo = if ($rule.ForwardAsAttachmentTo) { (@($rule.ForwardAsAttachmentTo|%{"$_"}) -join '; ') } else { '' }
        RedirectTo            = if ($rule.RedirectTo) { (@($rule.RedirectTo|%{"$_"}) -join '; ') } else { '' }
        TextNotifyTo          = if ($rule.SendTextMessageNotificationTo) { (@($rule.SendTextMessageNotificationTo|%{"$_"}) -join '; ') } else { '' }
        DeleteMessage = [bool]$rule.DeleteMessage
        RuleIdentity  = $rid
        Description   = (("$($rule.Description)" -replace "`r`n",' ') -replace '\s+',' ').Trim()
    }
}

# ---------------------------------------------------------------------------
# 5. [A] SOMBREAMENTO -- a regra morta
# ---------------------------------------------------------------------------
Write-Host "`n=== SOMBREAMENTO (regra que nunca executa) ===" -ForegroundColor Yellow
$mortas = 0
foreach ($g in ($map | Where-Object { $_.CondSig } | Group-Object CondSig | Where-Object Count -gt 1)) {
    $ord = @($g.Group | Sort-Object Ordem)
    for ($i = 0; $i -lt $ord.Count - 1; $i++) {
        if (-not $ord[$i].Stop) { continue }
        $bloq = $ord[$i]
        for ($j = $i + 1; $j -lt $ord.Count; $j++) {
            $morta = $ord[$j]
            $mesmaAcao = ($morta.ActSig -eq $bloq.ActSig)
            $sev  = if ($mesmaAcao) { 'REDUNDANTE' } else { 'REGRA MORTA' }
            $acao = if ($mesmaAcao) {
                "Duplicata inofensiva. Apagar '$($morta.Nome)' para limpar a lista."
            } else {
                "CONFIRMAR com o usuario qual pasta e a correta, e apagar a regra perdedora."
            }
            Add-Finding $sev 'InboxRule' "$($morta.Nome)  [ordem $($morta.Ordem)]" `
                "Condicao identica a '$($bloq.Nome)' (ordem $($bloq.Ordem)), que tem STOP. Destino desta: '$($morta.MoveToFolder)' | destino da que ganha: '$($bloq.MoveToFolder)'." $acao
            if (-not $mesmaAcao) {
                $mortas++
                Write-Host "  [MORTA] '$($morta.Nome)' nunca roda -- bloqueada por '$($bloq.Nome)'" -ForegroundColor Red
                Write-Host "          esperado em '$($morta.MoveToFolder)', vai para '$($bloq.MoveToFolder)'" -ForegroundColor Red
            } else {
                Write-Host "  [dup]   '$($morta.Nome)' e redundante com '$($bloq.Nome)' (mesmo destino)" -ForegroundColor DarkYellow
            }
        }
        break
    }
}
if ($mortas -eq 0) { Write-Host "Nenhuma regra morta por condicao identica." -ForegroundColor DarkGray }
Write-Host "Obs: so deteta condicao IDENTICA. Regra mais ampla engolindo uma mais estreita" -ForegroundColor DarkGray
Write-Host "     exige olho humano -- ver o mapa de execucao no fim." -ForegroundColor DarkGray

# ---------------------------------------------------------------------------
# 6. [B] REGRAS OCULTAS DE USUARIO
# ---------------------------------------------------------------------------
Write-Host "`n=== REGRAS OCULTAS ===" -ForegroundColor Yellow
$ocultasUser = @($map | Where-Object { $_.Oculta -and -not $_.Sistema })
$ocultasSys  = @($map | Where-Object { $_.Oculta -and $_.Sistema })
if ($ocultasSys.Count -gt 0) {
    Write-Host "De sistema (normal, ignorar): $(($ocultasSys.Nome) -join ', ')" -ForegroundColor DarkGray
}
if ($ocultasUser.Count -eq 0) {
    Write-Host "Nenhuma regra oculta de usuario." -ForegroundColor Green
} else {
    foreach ($o in $ocultasUser) {
        $sev = if ($o.Enabled) { 'OCULTA ATIVA' } else { 'OCULTA INATIVA' }
        Write-Host "  [$sev] '$($o.Nome)' (ordem $($o.Ordem), stop=$($o.Stop))" -ForegroundColor Red
        Write-Host "         $($o.Description)" -ForegroundColor DarkGray
        Add-Finding $sev 'InboxRule' "$($o.Nome)  [ordem $($o.Ordem)]" `
            "Nao aparece em Rules and Alerts do Outlook. Enabled=$($o.Enabled), Stop=$($o.Stop). $($o.Description)" `
            'Mostrar ao usuario e perguntar se o comportamento e o esperado. So sai por Remove-InboxRule (APPLY- com gate).'
    }
}

# ---------------------------------------------------------------------------
# 7. [C] REGRA QUE APAGA AVISO DO EXCHANGE
# ---------------------------------------------------------------------------
foreach ($m in $map) {
    if ($m.CondSig -match $SystemMailboxPattern -and ($m.DeleteMessage -or $m.MoveToFolder -match 'Deleted|Lixeira')) {
        Add-Finding 'CEGUEIRA' 'InboxRule' "$($m.Nome)  [ordem $($m.Ordem)]" `
            'Descarta mensagem da caixa de sistema do Exchange: NDR, aviso de cota e notificacao de falha.' `
            'Informar o usuario -- ele nao vera nenhum aviso do Exchange enquanto esta regra existir.'
    }
}

# ---------------------------------------------------------------------------
# 8. [E] ENCAMINHAMENTO -- mailbox E regras, avaliados de verdade
# ---------------------------------------------------------------------------
Write-Host "`n=== ENCAMINHAMENTO ===" -ForegroundColor Yellow
function Test-Targets {
    param([string]$Origem,[string]$Item,[string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return }
    $addrs = @([regex]::Matches($Raw,'[\w\.\-\+]+@[\w\.\-]+\.\w+') | ForEach-Object { $_.Value } | Select-Object -Unique)
    foreach ($a in $addrs) {
        if ($TrustedDomains -notcontains $a.Split('@')[1]) {
            Add-Finding 'EXTERNO' $Origem $Item "Destino fora dos dominios do tenant: $a" `
                'Causa provavel do bloqueio anti-autoforward. Alternativa: shared mailbox ou grupo interno.'
            Write-Host "  [EXTERNO] $Item -> $a" -ForegroundColor Red
        }
    }
    # alvo que nao resolve para SMTP (objeto de diretorio) nao pode ser dado como seguro
    if ($addrs.Count -eq 0 -or $Raw -match '\[EX:') {
        Add-Finding 'REVISAR' $Origem $Item "Alvo nao resolveu para SMTP: $Raw" 'Conferir manualmente para quem aponta.'
    }
}

$mbxFwdAddr = "$($mailbox.ForwardingAddress)"
$mbxFwdSmtp = "$($mailbox.ForwardingSmtpAddress)"
if ($mbxFwdAddr -or $mbxFwdSmtp) {
    Write-Host "Forwarding a nivel de mailbox ATIVO:" -ForegroundColor Red
    Write-Host "  ForwardingAddress .....: $mbxFwdAddr"
    Write-Host "  ForwardingSmtpAddress .: $mbxFwdSmtp"
    Write-Host "  DeliverToMailboxAndForward: $($mailbox.DeliverToMailboxAndForward)"
    Test-Targets 'Mailbox' 'ForwardingAddress'     $mbxFwdAddr
    Test-Targets 'Mailbox' 'ForwardingSmtpAddress' $mbxFwdSmtp
    if (-not $mailbox.DeliverToMailboxAndForward) {
        Add-Finding 'REVISAR' 'Mailbox' 'DeliverToMailboxAndForward=False' `
            'Encaminha SEM deixar copia local -- o usuario nao ve a mensagem na propria caixa.' 'Confirmar se e intencional.'
    }
} else { Write-Host "Sem forwarding a nivel de mailbox." -ForegroundColor DarkGray }

$comFwd = @($map | Where-Object { $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo -or $_.TextNotifyTo })
if ($comFwd.Count -eq 0) { Write-Host "Nenhuma regra encaminha ou redireciona." -ForegroundColor DarkGray }
foreach ($f in $comFwd) {
    Test-Targets 'InboxRule' "$($f.Nome) :: ForwardTo"       $f.ForwardTo
    Test-Targets 'InboxRule' "$($f.Nome) :: ForwardAsAttach" $f.ForwardAsAttachmentTo
    Test-Targets 'InboxRule' "$($f.Nome) :: RedirectTo"      $f.RedirectTo
    Test-Targets 'InboxRule' "$($f.Nome) :: TextNotifyTo"    $f.TextNotifyTo
}

# nomes duplicados -- carimbo do Outlook quando duplica regra
foreach ($d in ($map | Group-Object Nome | Where-Object Count -gt 1)) {
    Add-Finding 'DUPLICADA' 'InboxRule' "$($d.Name)" "$($d.Count) regras com o nome identico." 'Conferir qual manter.'
}
foreach ($s in ($map | Where-Object { $_.Nome -match '\(\d+\)\s*$' })) {
    Add-Finding 'DUPLICADA' 'InboxRule' "$($s.Nome)" 'Sufixo "(N)" e o carimbo do Outlook ao duplicar regra.' 'Conferir se a original ainda existe.'
}

# ---------------------------------------------------------------------------
# 9. Exportacao -- SEMPRE, inclusive no negativo (o ticket precisa da prova)
# ---------------------------------------------------------------------------
$csvMapa    = Join-Path $saida "rules-mapa-$slug-$stamp.csv"
$csvAchados = Join-Path $saida "rules-achados-$slug-$stamp.csv"
$map | Sort-Object Ordem | Export-Csv -Path $csvMapa -NoTypeInformation -Encoding UTF8
if ($findings.Count -gt 0) { $findings | Export-Csv -Path $csvAchados -NoTypeInformation -Encoding UTF8 }
else {
    [PSCustomObject]@{ Mailbox=$MailboxAddress; Severidade='NADA ENCONTRADO'; Origem='-'; Item='-'
        Detalhe="$($todas.Count) regras, 0 corrompidas, 0 ocultas de usuario, 0 mortas, 0 destino externo."; Acao='-'
    } | Export-Csv -Path $csvAchados -NoTypeInformation -Encoding UTF8
}

# ============ RESUMO -- ULTIMO BLOCO, NADA DEPOIS ============
$c = @{}
foreach ($k in 'CORROMPIDA','REGRA MORTA','OCULTA ATIVA','EXTERNO','CEGUEIRA','REDUNDANTE','DUPLICADA','REVISAR') {
    $c[$k] = @($findings | Where-Object Severidade -eq $k).Count
}

Write-Host "`n`n=============== MAPA DE EXECUCAO ===============" -ForegroundColor Cyan
Write-Host "Ordem em que as regras tocam cada mensagem. 'Stop' corta as de baixo" -ForegroundColor DarkGray
Write-Host "APENAS para a mensagem que casou com aquela regra." -ForegroundColor DarkGray
$map | Sort-Object Ordem | Select-Object Ordem,
    @{n='Ocu';e={ if ($_.Sistema) {'sys'} elseif ($_.Oculta) {'SIM'} else {''} }},
    @{n='Stop';e={ if ($_.Stop) {'STOP'} else {''} }}, Enabled,
    @{n='Regra';e={ $_.Nome.Substring(0,[Math]::Min(42,$_.Nome.Length)) }},
    @{n='Pasta';e={ $_.MoveToFolder.Substring(0,[Math]::Min(32,$_.MoveToFolder.Length)) }} |
    Format-Table -AutoSize

Write-Host "=================== RESUMO ===================" -ForegroundColor Cyan
Write-Host ("Mailbox ..............: {0}" -f $MailboxAddress)
Write-Host ("Regras (visiveis/total): {0} / {1}" -f $visiveis.Count, $todas.Count)
Write-Host ("Corrompidas ..........: {0}" -f $c['CORROMPIDA'])   -ForegroundColor $(if($c['CORROMPIDA']){'Red'}else{'Green'})
Write-Host ("REGRAS MORTAS ........: {0}" -f $c['REGRA MORTA'])  -ForegroundColor $(if($c['REGRA MORTA']){'Red'}else{'Green'})
Write-Host ("Ocultas de usuario ...: {0}" -f $c['OCULTA ATIVA']) -ForegroundColor $(if($c['OCULTA ATIVA']){'Red'}else{'Green'})
Write-Host ("Destinos externos ....: {0}" -f $c['EXTERNO'])      -ForegroundColor $(if($c['EXTERNO']){'Red'}else{'Green'})
Write-Host ("Apaga aviso do Exchange: {0}" -f $c['CEGUEIRA'])    -ForegroundColor $(if($c['CEGUEIRA']){'Yellow'}else{'Green'})
Write-Host ("Redundantes/duplicadas : {0}" -f ($c['REDUNDANTE'] + $c['DUPLICADA'])) -ForegroundColor DarkGray
Write-Host ("A revisar ............: {0}" -f $c['REVISAR'])      -ForegroundColor DarkGray

if ($findings.Count -gt 0) {
    Write-Host ""
    $findings | Where-Object { $_.Severidade -notin 'REDUNDANTE','DUPLICADA' } |
        Select-Object Severidade,Item,Detalhe | Format-Table -AutoSize -Wrap
}

Write-Host ""
if ($c['CORROMPIDA'] -gt 0) {
    Write-Host "VEREDITO: INCONCLUSIVO -- $($c['CORROMPIDA']) regra(s) ilegivel(eis). Rodar Remove-CorruptInboxRules.ps1 -DryRun." -ForegroundColor Red
} elseif ($c['REGRA MORTA'] -gt 0) {
    Write-Host "VEREDITO: $($c['REGRA MORTA']) regra(s) NUNCA executa(m). Essa e a queixa 'my rules are not working'." -ForegroundColor Red
} elseif ($c['OCULTA ATIVA'] -gt 0) {
    Write-Host "VEREDITO: ha regra ativa invisivel no Outlook do usuario. Mostrar a ele antes de mexer." -ForegroundColor Red
} elseif ($c['EXTERNO'] -gt 0) {
    Write-Host "VEREDITO: destino externo -- causa do bloqueio anti-autoforward." -ForegroundColor Red
} else {
    Write-Host "VEREDITO: limpo. Sem regra morta, oculta, corrompida ou destino externo." -ForegroundColor Green
}

Write-Host ""
if (Test-Path $csvMapa)    { Write-Host "Mapa ....: $csvMapa" }    else { Write-Host "ATENCAO: mapa NAO gravado" -ForegroundColor Red }
if (Test-Path $csvAchados) { Write-Host "Achados .: $csvAchados" } else { Write-Host "ATENCAO: achados NAO gravado" -ForegroundColor Red }
