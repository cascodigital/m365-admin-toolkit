<#
.SYNOPSIS
    Script pai interativo: audita e (opcionalmente) habilita Auto-Expanding
    Archive em mailboxes do Exchange Online.

.DESCRIPTION
    Roda sem argumento nenhum. Ele pergunta o email, verifica tudo e decide:

      [A] JA HABILITADO          -> reporta e encerra. Nada a fazer.
      [B] SEM LICENCA            -> avisa e encerra. Habilitar nao resolveria.
      [C] DESABILITADO + LICENCA -> oferece habilitar (S/N).
                                    Se SIM: executa, espera 30s com contador,
                                    reconsulta e mostra a prova.
      [D] SEM ARCHIVE            -> avisa e entrega o comando pra provisionar.

    Ao terminar, pergunta se quer processar outra mailbox. Loop ate voce sair.

    A licenca e inferida pela ArchiveQuota (>= 90 GB = E3/E5 ou add-on EOA),
    porque o Microsoft Graph PowerShell esta bloqueado por Conditional Access
    nesta maquina e nao da pra consultar o SKU direto.

.PARAMETER User
    Opcional. Pula a primeira pergunta e ja processa este email.

.PARAMETER Admin
    UPN admin usado no Connect-ExchangeOnline.

.PARAMETER Force
    Pula a confirmacao S/N. Use com juizo.

.EXAMPLE
    .\Manage-ArchiveAutoExpand.ps1

.EXAMPLE
    .\Manage-ArchiveAutoExpand.ps1 -User user@contoso.com

.NOTES
    PIM OBRIGATORIO: Entra -> Identity Governance -> PIM -> My roles ->
    Microsoft Entra roles -> Activate -> "Exchange Recipient Administrator".
    Sem PIM ativo o Get-Mailbox mente: retorna "object couldn't be found"
    em vez de "access denied".

    Enable-Mailbox -AutoExpandingArchive e IRREVERSIVEL.
#>

[CmdletBinding()]
param(
    [string] $User,
    [Parameter(Mandatory = $true)][string] $Admin,
    [switch] $Force
)

$ErrorActionPreference = 'Stop'

# ================================================================ helpers ===
function Write-Step  { param($m) Write-Host "[*] $m" -ForegroundColor Cyan }
function Write-Ok    { param($m) Write-Host "[+] $m" -ForegroundColor Green }
function Write-Warn2 { param($m) Write-Host "[!] $m" -ForegroundColor Yellow }
function Write-Err2  { param($m) Write-Host "[X] $m" -ForegroundColor Red }

function Write-Banner {
    param($Text, $Color = 'Magenta')
    $line = '=' * 62
    Write-Host ''
    Write-Host $line -ForegroundColor $Color
    Write-Host ("  " + $Text) -ForegroundColor $Color
    Write-Host $line -ForegroundColor $Color
}

function ConvertTo-Gb {
    param($Size)
    if ($null -eq $Size) { return $null }
    $s = $Size.ToString()
    if ($s -match '\(([\d,\.]+)\s*bytes\)') {
        return [math]::Round(([double]($Matches[1] -replace '[,\.]','')) / 1GB, 2)
    }
    if ($s -match 'Unlimited') { return 'Unlimited' }
    return $s
}

function Start-Countdown {
    param([int] $Seconds = 30, [string] $Activity = 'Aguardando propagacao do EXO')
    for ($i = $Seconds; $i -gt 0; $i--) {
        $pct = (($Seconds - $i) / $Seconds) * 100
        Write-Progress -Activity $Activity -Status "$i segundos restantes..." -PercentComplete $pct
        Start-Sleep -Seconds 1
    }
    Write-Progress -Activity $Activity -Completed
}

function Read-YesNo {
    param([string] $Prompt)
    while ($true) {
        $a = (Read-Host "$Prompt [S/N]").Trim().ToUpper()
        if ($a -in @('S','SIM','Y','YES')) { return $true  }
        if ($a -in @('N','NAO','NÃO','NO')) { return $false }
        Write-Warn2 "Responde S ou N, primata."
    }
}

function Resolve-TargetMailbox {
    param([string] $Id)
    try {
        return Get-Mailbox -Identity $Id -ErrorAction Stop
    } catch {
        Write-Warn2 "Get-Mailbox falhou para '$Id'. Tentando Get-Recipient..."
    }
    $local = ($Id -split '@')[0]
    $rcpt  = @(Get-Recipient -Filter "EmailAddresses -like '*$local*'" -ErrorAction SilentlyContinue)

    if ($rcpt.Count -eq 0) {
        throw "Nenhum recipient encontrado para '$Id'. Confere o PIM e o endereco."
    }
    if ($rcpt.Count -gt 1) {
        Write-Warn2 "Multiplos recipients casaram com '$local':"
        $rcpt | Select-Object Name, PrimarySmtpAddress, RecipientType | Format-Table -AutoSize | Out-Host
        throw "Ambiguo. Rode de novo com o PrimarySmtpAddress exato."
    }
    Write-Ok "Resolvido via Get-Recipient -> Name = $($rcpt[0].Name)"
    return Get-Mailbox -Identity $rcpt[0].Name -ErrorAction Stop
}

# =============================================================== conexao ====
function Connect-IfNeeded {
    try   { $conn = Get-ConnectionInformation -ErrorAction SilentlyContinue }
    catch { $conn = $null }

    if (-not $conn -or $conn.State -ne 'Connected') {
        Write-Step "Sem sessao EXO ativa. Conectando como $Admin ..."
        Write-Warn2 "PIM 'Exchange Recipient Administrator' precisa estar ATIVO agora."
        Connect-ExchangeOnline -UserPrincipalName $Admin -ShowBanner:$false
        Write-Ok "Conectado."
    } else {
        Write-Ok "Sessao EXO ja ativa como $($conn.UserPrincipalName)."
    }
}

# ========================================================= nucleo por user ==
function Invoke-ArchiveCheck {
    param([string] $Target)

    Write-Banner "AUDITANDO: $Target"

    try {
        $mbx = Resolve-TargetMailbox -Id $Target
    } catch {
        Write-Err2 $_.Exception.Message
        return
    }

    # ArchiveStatus MENTE para mailbox cloud (heranca on-prem).
    # A verdade e ArchiveState + ArchiveGuid.
    $hasArchive = ($mbx.ArchiveGuid -ne [guid]::Empty) -and ($mbx.ArchiveState -ne 'None')
    $quotaGb    = ConvertTo-Gb $mbx.ArchiveQuota
    $warnGb     = ConvertTo-Gb $mbx.ArchiveWarningQuota

    $usedGb = $null; $pct = $null; $items = $null
    if ($hasArchive) {
        $st = Get-MailboxStatistics -Identity $mbx.Identity -Archive -ErrorAction SilentlyContinue
        if ($st) {
            $usedGb = ConvertTo-Gb $st.TotalItemSize
            $items  = $st.ItemCount
            if (($usedGb -is [double]) -and ($quotaGb -is [double]) -and $quotaGb -gt 0) {
                $pct = [math]::Round(($usedGb / $quotaGb) * 100, 1)
            }
        }
    }

    $primary = Get-MailboxStatistics -Identity $mbx.Identity -ErrorAction SilentlyContinue

    $licensed = ($quotaGb -eq 'Unlimited') -or (($quotaGb -is [double]) -and ($quotaGb -ge 90))

    # ---------------------------------------------------------- painel ----
    Write-Host ''
    [pscustomobject]@{
        DisplayName                 = $mbx.DisplayName
        PrimarySmtpAddress          = $mbx.PrimarySmtpAddress
        RecipientTypeDetails        = $mbx.RecipientTypeDetails
        ArchiveState                = $mbx.ArchiveState
        HasArchive                  = $hasArchive
        AutoExpandingArchiveEnabled = $mbx.AutoExpandingArchiveEnabled
        ArchiveQuotaGB              = $quotaGb
        ArchiveWarningQuotaGB       = $warnGb
        ArchiveUsedGB               = $usedGb
        ArchivePercentUsed          = $pct
        ArchiveItemCount            = $items
        PrimarySizeGB               = (ConvertTo-Gb $primary.TotalItemSize)
        RetentionPolicy             = $mbx.RetentionPolicy
        LitigationHoldEnabled       = $mbx.LitigationHoldEnabled
        InPlaceHolds                = ($mbx.InPlaceHolds -join ', ')
    } | Format-List

    # ------------------------------------------------------- cenario [A] --
    if ($mbx.AutoExpandingArchiveEnabled -eq $true) {
        Write-Banner "JA ESTA HABILITADO. Parabens, era isso." 'Green'
        Write-Ok "Auto-Expanding Archive = True. Nao ha nada a fazer."
        Write-Host ''
        Write-Host " O archive cresce automaticamente ate 1.5 TB."
        if ($pct) { Write-Host " Consumo atual: $usedGb GB ($pct% da quota base de $quotaGb GB)." }
        Write-Host ''
        Write-Host " Se mesmo assim o usuario reclama de archive cheio:"
        Write-Host "   - o provisionamento da area auxiliar leva ate 30 dias;"
        Write-Host "   - ou a retention policy nao esta movendo item nenhum"
        Write-Host "     (RetentionPolicy atual: $($mbx.RetentionPolicy))."
        return
    }

    # ------------------------------------------------------- cenario [D] --
    if (-not $hasArchive) {
        Write-Banner "SEM ARCHIVE PROVISIONADO" 'Yellow'
        Write-Warn2 "ArchiveState = $($mbx.ArchiveState). Nao existe archive nesta mailbox."
        Write-Host ''
        Write-Host " Auto-expand nem entra em pauta. Provisione o archive primeiro:"
        Write-Host "   Enable-Mailbox -Identity '$($mbx.PrimarySmtpAddress)' -Archive" -ForegroundColor Yellow
        Write-Host ''
        Write-Host " Depois rode este script de novo."
        return
    }

    # ------------------------------------------------------- cenario [B] --
    if (-not $licensed) {
        Write-Banner "SEM LICENCA DE ARCHIVING" 'Red'
        Write-Err2 "ArchiveQuota = $quotaGb GB."
        Write-Host ''
        Write-Host " Quota de ~50 GB indica AUSENCIA de licenca de archiving."
        Write-Host " Auto-Expanding Archive exige E3, E5 ou o add-on"
        Write-Host " Exchange Online Archiving (EOA)."
        Write-Host ''
        Write-Warn2 "Habilitar aqui NAO daria espaco nenhum. Resolva a licenca primeiro."
        Write-Host ''
        Write-Host " Proximo passo: verificar o SKU do usuario no portal"
        Write-Host " admin.microsoft.com -> Users -> Licenses and apps."
        Write-Host " (Graph PowerShell esta bloqueado por Conditional Access nesta maquina.)"
        return
    }

    # ------------------------------------------------------- cenario [C] --
    Write-Banner "DESABILITADO - MAS TEM LICENCA. DA PRA HABILITAR." 'Yellow'
    Write-Ok "Licenca confirmada (ArchiveQuota = $quotaGb GB)."
    Write-Warn2 "AutoExpandingArchiveEnabled = False"
    if ($pct) {
        Write-Host ''
        Write-Host " Consumo atual: $usedGb GB de $quotaGb GB  ($pct%)"
        if ($warnGb -is [double] -and $usedGb -is [double] -and $usedGb -ge $warnGb) {
            Write-Warn2 " Ja cruzou a ArchiveWarningQuota de $warnGb GB - por isso o alerta."
        }
    }

    if ($mbx.LitigationHoldEnabled -eq $true -or $mbx.InPlaceHolds) {
        Write-Host ''
        Write-Warn2 "ATENCAO: MAILBOX SOB HOLD."
        Write-Host "   LitigationHold : $($mbx.LitigationHoldEnabled)"
        Write-Host "   InPlaceHolds   : $($mbx.InPlaceHolds -join ', ')"
        Write-Host "   Archives auxiliares tem limitacao em eDiscovery/Content Search."
        Write-Host "   Valide com Cyber/Legal antes de prosseguir." -ForegroundColor Yellow
    }

    Write-Host ''
    Write-Host ' ESTA ACAO E IRREVERSIVEL:' -ForegroundColor Red
    Write-Host '   - nao desabilita depois, nem com ticket na Microsoft' -ForegroundColor Red
    Write-Host '   - trava retorno da mailbox pro on-prem e complica move entre tenants' -ForegroundColor Red
    Write-Host '   - limitacoes em eDiscovery/Content Search nos archives auxiliares' -ForegroundColor Red
    Write-Host '   - cliente precisa ser Outlook 2016+ ou OWA' -ForegroundColor Red
    Write-Host ''

    if ($Force) {
        Write-Warn2 "-Force ativo: pulando confirmacao."
    } elseif (-not (Read-YesNo "Habilitar Auto-Expanding Archive para $($mbx.PrimarySmtpAddress) AGORA?")) {
        Write-Host ''
        Write-Warn2 "ABORTADO. Nada foi alterado."
        Write-Host " Quando decidir, o comando e:"
        Write-Host "   Enable-Mailbox -Identity '$($mbx.PrimarySmtpAddress)' -AutoExpandingArchive" -ForegroundColor Yellow
        return
    }

    # ---------------------------------------------------------- executar --
    Write-Host ''
    Write-Step "Executando Enable-Mailbox -AutoExpandingArchive ..."
    try {
        Enable-Mailbox -Identity $mbx.Identity -AutoExpandingArchive -ErrorAction Stop | Out-Null
        Write-Ok "Cmdlet executado sem erro."
    } catch {
        Write-Err2 "FALHOU: $($_.Exception.Message)"
        Write-Host ''
        Write-Host " Suspeito 1: PIM expirado (janela de 1-8h). Reative e tente de novo."
        Write-Host " Suspeito 2: exige Exchange Administrator FULL, que voce nao tem no"
        Write-Host "             cloud - so o subset Recipient. Nesse caso, escalar."
        return
    }

    Write-Host ''
    Write-Step "Aguardando 30s de propagacao do EXO antes de verificar..."
    Start-Countdown -Seconds 30

    # --------------------------------------------------------- verificar --
    Write-Step "Verificando resultado..."
    Write-Host ''
    Write-Host "Get-Mailbox `"$($mbx.PrimarySmtpAddress)`" | FL AutoExpandingArchiveEnabled,ArchiveState,ArchiveQuota" -ForegroundColor DarkGray
    Write-Host ''

    $after = Get-Mailbox -Identity $mbx.Identity -ErrorAction Stop
    $after | Format-List AutoExpandingArchiveEnabled, ArchiveState, ArchiveQuota

    if ($after.AutoExpandingArchiveEnabled -eq $true) {
        Write-Banner "CONFIRMADO: AutoExpandingArchiveEnabled = True" 'Green'
        Write-Host ''
        Write-Host " O archive agora cresce automaticamente ate 1.5 TB."
        Write-Host ''
        Write-Warn2 "MAS: o provisionamento da area auxiliar e ASSINCRONO e leva ATE 30 DIAS."
        Write-Host " Ate la a margem real continua sendo o que sobra dos $quotaGb GB."
        if ($usedGb -is [double] -and $quotaGb -is [double]) {
            Write-Host " Sobra hoje: aprox. $([math]::Round($quotaGb - $usedGb, 2)) GB."
        }
        Write-Host " Se bater o teto antes disso, o Managed Folder Assistant para de"
        Write-Host " mover itens e o usuario toma erro."
        Write-Host ''
        Write-Host " NAO feche o ticket hoje. Remonitore em 7 dias rodando este script." -ForegroundColor Yellow
    } else {
        Write-Banner "AINDA FALSE - propagacao pendente" 'Yellow'
        Write-Warn2 "O cmdlet nao deu erro, entao provavelmente e so replicacao do EXO."
        Write-Host " Reconsulte em ~15 minutos:"
        Write-Host "   Get-Mailbox '$($mbx.PrimarySmtpAddress)' | FL AutoExpandingArchiveEnabled" -ForegroundColor Yellow
        Write-Host ''
        Write-Host " Se continuar False depois de 1h, foi permissao silenciosa: escalar."
    }
}

# ================================================================= main =====
Clear-Host
Write-Banner 'AUTO-EXPANDING ARCHIVE - AUDITORIA E HABILITACAO'
Write-Host '  Exchange Online' -ForegroundColor DarkGray
Write-Host ''
Write-Warn2 'Confirme que o PIM "Exchange Recipient Administrator" esta ATIVO.'
Write-Host '  Entra -> Identity Governance -> PIM -> My roles -> Entra roles -> Activate' -ForegroundColor DarkGray

Connect-IfNeeded

$target = $User
while ($true) {

    if (-not $target) {
        Write-Host ''
        $target = (Read-Host 'Email da pessoa (ENTER em branco para sair)').Trim()
        if ([string]::IsNullOrWhiteSpace($target)) { break }
    }

    try {
        Invoke-ArchiveCheck -Target $target
    } catch {
        Write-Err2 "Erro inesperado: $($_.Exception.Message)"
    }

    $target = $null
    Write-Host ''
    if (-not (Read-YesNo 'Verificar outra mailbox?')) { break }
}

Write-Host ''
Write-Ok 'Encerrado.'
Write-Host 'Sessao EXO deixada ABERTA. Pra fechar:' -ForegroundColor DarkGray
Write-Host '  Disconnect-ExchangeOnline -Confirm:$false' -ForegroundColor DarkGray
