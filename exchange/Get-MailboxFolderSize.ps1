<#
.SYNOPSIS
    Diagnostico de tamanho + arquivamento (MRM) de uma ou mais mailboxes.
    SOMENTE LEITURA: nenhum cmdlet de escrita no tenant.

.DESCRIPTION
    Por mailbox mostra:
      - primaria x quota, recoverable items, archive x quota, auto-expanding
      - retention policy, a tag padrao (DPT) e o que bloqueia o MRM (holds, ElcProcessing)
      - pastas apontando pra tag que nao existe mais (o WARNING "Unable to retrieve...")
      - TOP N pastas (tamanho proprio, sem somar subpastas) com: item mais antigo, tag
        efetiva e prazo, status do MRM, entrada estimada GB/dia e quanto a pasta ficaria
        com tag de 30 dias.

    Leitura do status:
      MRM pendente   -> tag aplicada, o assistente ainda nao passou (ciclo ate 7 dias)
      ATRASADO       -> tem item mais velho que o prazo + 7 dias: investigar
      OK no prazo    -> o arquivamento ja fez o trabalho; o volume e e-mail recente.
                        So reduz com tag mais curta.

    GB/dia e estimativa: tamanho / dias entre o item mais antigo e o mais novo.

.PARAMETER Mailbox
    Um ou mais e-mails. Sem parametro, pergunta (aceita varios separados por virgula).

.PARAMETER Top
    Quantas pastas mostrar por mailbox. Padrao 5.

.EXAMPLE
    .\Get-MailboxFolderSize.ps1 -Mailbox user1@contoso.com,user2@contoso.com

.NOTES
    PIM "Exchange Recipient Administrator" ativo ANTES de rodar, senao o Get-Mailbox
    responde "couldn't be found" em vez de acesso negado.
#>
#requires -Version 5.1
[CmdletBinding()]
param(
    [string[]]$Mailbox,
    [int]$Top = 5,
    [Parameter(Mandatory = $true)][string]$AdminUPN
)

$ErrorActionPreference = 'Stop'

function Convert-ExchangeSizeToBytes {
    param([AllowNull()]$Size)
    if ($null -eq $Size) { return [int64]0 }
    try {
        if ($Size.PSObject.Methods.Name -contains 'ToBytes') { return [int64]$Size.ToBytes() }
        if ($null -ne $Size.Value -and $Size.Value.PSObject.Methods.Name -contains 'ToBytes') { return [int64]$Size.Value.ToBytes() }
    } catch { }
    # No EXO REST o tamanho vem como TEXTO "9.88 GB (10,608,936,405 bytes)".
    # Ordenar esse texto direto da ordem lexical errada; sempre converter antes.
    if ([string]$Size -match '\(([0-9,\.\s]+)\s+bytes\)') { return [int64]($Matches[1] -replace '[^0-9]', '') }
    return [int64]0
}

function Convert-ToDays {
    param([AllowNull()]$Age)
    if ($null -eq $Age -or [string]::IsNullOrWhiteSpace([string]$Age)) { return $null }
    try { return [int]([TimeSpan]::Parse([string]$Age).TotalDays) } catch { return $null }
}

function GB { param([int64]$b) [math]::Round($b / 1GB, 2) }

# ---------- saida: prova que da pra escrever, nao so que existe ----------
$saida = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.ProviderPath }
try { $__p = Join-Path $saida "w$PID.tmp"; [IO.File]::WriteAllText($__p, 'x'); [IO.File]::Delete($__p) }
catch { $saida = $env:TEMP; Write-Host "AVISO: saida redirecionada para $saida" -ForegroundColor Yellow }

# ---------- entrada ----------
if (-not $Mailbox) {
    $Mailbox = (Read-Host 'E-mail(s) COMPLETO(s), separados por virgula') -split '[,;\s]+' | Where-Object { $_ }
}
$Mailbox = @($Mailbox | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^[^\s@]+@[^\s@]+\.[^\s@]+$' })
if (-not $Mailbox) { throw 'Nenhum e-mail valido informado.' }

# ---------- conexao: reaproveita a sessao aberta ----------
$conectouAqui = $false
$conectado = $false
try { $conectado = [bool](Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object State -eq 'Connected') } catch { }
if (-not $conectado) {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Write-Host "Conecte-se com: $AdminUPN" -ForegroundColor Yellow
    Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false -ErrorAction Stop
    $conectouAqui = $true
}

$tagCache = @{}
function Get-TagInfo {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    if (-not $tagCache.ContainsKey($Name)) {
        $t = Get-RetentionPolicyTag -Identity $Name -ErrorAction SilentlyContinue
        $tagCache[$Name] = if ($t) {
            [PSCustomObject]@{
                Name    = $t.Name
                Type    = [string]$t.Type
                Action  = [string]$t.RetentionAction
                Days    = Convert-ToDays $t.AgeLimitForRetention
                Enabled = [bool]$t.RetentionEnabled
            }
        } else { $null }
    }
    return $tagCache[$Name]
}

$csv = New-Object System.Collections.Generic.List[object]
$agora = Get-Date

try {
    foreach ($mbx in $Mailbox) {
        Write-Host "`n==================== $mbx ====================" -ForegroundColor Cyan
        try {
            $m = Get-Mailbox -Identity $mbx -ErrorAction Stop
        } catch {
            Write-Host "Nao achei $mbx. PIM Exchange Recipient Administrator esta ativo? ($($_.Exception.Message))" -ForegroundColor Red
            continue
        }

        $st = Get-MailboxStatistics -Identity $mbx -ErrorAction Stop
        $usado = Convert-ExchangeSizeToBytes $st.TotalItemSize
        $quota = Convert-ExchangeSizeToBytes $m.ProhibitSendReceiveQuota
        $recov = Convert-ExchangeSizeToBytes $st.TotalDeletedItemSize

        $arqTxt = 'SEM ARCHIVE'
        if ([string]$m.ArchiveStatus -eq 'Active') {
            $ast = Get-MailboxStatistics -Identity $mbx -Archive -ErrorAction SilentlyContinue
            $aUsado = Convert-ExchangeSizeToBytes $ast.TotalItemSize
            $aQuota = Convert-ExchangeSizeToBytes $m.ArchiveQuota
            $aPct = if ($aQuota) { [math]::Round(100 * $aUsado / $aQuota) } else { '?' }
            $auto = if ($m.AutoExpandingArchiveEnabled) { 'auto-expanding LIGADO' } else { 'auto-expanding desligado' }
            $arqTxt = "$(GB $aUsado) GB de $(GB $aQuota) GB ($aPct%) | $auto"
        }

        # tag padrao (DPT) de arquivamento da policy
        $dpt = $null
        $pessoais = @()
        if ($m.RetentionPolicy) {
            $pol = Get-RetentionPolicy -Identity ([string]$m.RetentionPolicy) -ErrorAction SilentlyContinue
            foreach ($link in @($pol.RetentionPolicyTagLinks)) {
                $ti = Get-TagInfo ([string]$link)
                if (-not $ti) { continue }
                if ($ti.Type -eq 'All' -and $ti.Action -eq 'MoveToArchive') { $dpt = $ti }
                if ($ti.Type -eq 'Personal' -and $ti.Enabled -and $ti.Days) { $pessoais += "$($ti.Name)" }
            }
        }
        $dptTxt = if ($dpt) { "$($dpt.Name) ($($dpt.Days)d)" } else { 'NENHUMA - pasta sem tag pessoal nunca arquiva' }

        $bloq = @()
        if ($m.ElcProcessingDisabled) { $bloq += 'ElcProcessingDisabled' }
        if ($m.RetentionHoldEnabled)  { $bloq += 'RetentionHold' }
        if ($m.LitigationHoldEnabled) { $bloq += 'LitigationHold (nao impede mover, segura o que for apagado)' }
        $holds = @($m.InPlaceHolds | Where-Object { $_ })
        if ($holds.Count) { $bloq += "InPlaceHolds: $($holds.Count)" }
        $bloqTxt = if ($bloq) { $bloq -join ', ' } else { 'nenhum' }

        # pastas; os WARNINGs de tag inexistente sao contados, nao despejados na tela
        $avisos = $null
        $fs = Get-MailboxFolderStatistics -Identity $mbx -IncludeOldestAndNewestItems -WarningAction SilentlyContinue -WarningVariable avisos -ErrorAction Stop
        $orfas = @($avisos | Where-Object { "$_" -match 'policy tag' }).Count

        $pctTxt = if ($quota) { "$([math]::Round(100 * $usado / $quota))%" } else { 'quota ?' }
        Write-Host ("Primaria   : {0} GB de {1} GB ({2}) | {3} itens | Recoverable {4} GB" -f (GB $usado), (GB $quota), $pctTxt, $st.ItemCount, (GB $recov))
        Write-Host ("Archive    : {0}" -f $arqTxt)
        Write-Host ("Politica   : {0} | padrao (DPT): {1}" -f $m.RetentionPolicy, $dptTxt)
        if ($pessoais) { Write-Host ("Tags p/ pasta: {0}" -f ($pessoais -join ', ')) -ForegroundColor DarkGray }
        Write-Host ("Bloqueios  : {0}" -f $bloqTxt) -ForegroundColor $(if ($bloq) { 'Yellow' } else { 'Gray' })
        if ($orfas) { Write-Host ("Tag orfa   : {0} pasta(s) apontam pra tag que nao existe mais no tenant" -f $orfas) -ForegroundColor Yellow }

        $linhas = foreach ($f in $fs) {
            if ($f.RecoverableItemsFolder -or [string]$f.FolderType -eq 'Root') { continue }
            $b = Convert-ExchangeSizeToBytes $f.FolderSize
            if ($b -le 0) { continue }

            $tag = if ($f.ArchivePolicy) { Get-TagInfo ([string]$f.ArchivePolicy) } else { $dpt }
            $tagTxt = if ($f.ArchivePolicy -and -not $tag) { "$($f.ArchivePolicy) (ORFA)" }
                      elseif (-not $tag) { '-' }
                      elseif ($f.ArchivePolicy) { "$($tag.Name)" }
                      else { "$($tag.Name) [DPT]" }

            $velho = $f.OldestItemReceivedDate
            $novo  = $f.NewestItemReceivedDate
            $idade = if ($velho) { [int]($agora - [datetime]$velho).TotalDays } else { $null }
            $email = ([string]$f.ContainerClass -eq 'IPF.Note')

            $status = '-'
            if ([string]$f.FolderType -eq 'DeletedItems') { $status = 'Lixeira: idade conta da exclusao' }
            elseif (-not $email) { $status = 'nao e pasta de e-mail' }
            elseif (-not $tag -or -not $tag.Enabled -or -not $tag.Days) { $status = 'SEM ARQUIVAMENTO' }
            elseif ([string]$f.RetentionFlags -match 'NeedsRescan') { $status = 'MRM pendente (tag nova)' }
            elseif ($null -ne $idade -and $idade -gt ($tag.Days + 7)) { $status = "ATRASADO: item de ${idade}d" }
            else { $status = "OK no prazo ($($tag.Days)d)" }

            $porDia = $null; $com30 = $null
            if ($email -and $velho -and $novo) {
                $span = [math]::Max(1, ([datetime]$novo - [datetime]$velho).TotalDays)
                $porDia = [math]::Round(($b / 1GB) / $span, 2)
                $com30 = [math]::Min((GB $b), [math]::Round($porDia * 30, 1))
            }

            [PSCustomObject]@{
                Mailbox      = $mbx
                Pasta        = $f.FolderPath
                Itens        = $f.ItemsInFolder
                GB           = GB $b
                PctCaixa     = if ($usado) { [math]::Round(100 * $b / $usado) } else { 0 }
                MaisAntigo   = if ($velho) { ([datetime]$velho).ToString('yyyy-MM-dd') } else { '' }
                IdadeDias    = $idade
                Tag          = $tagTxt
                Status       = $status
                GBporDia     = $porDia
                GBcom30d     = $com30
                Bytes        = $b
            }
        }

        $topo = @($linhas | Sort-Object Bytes -Descending | Select-Object -First $Top)
        $somaTop = ($topo | Measure-Object Bytes -Sum).Sum
        Write-Host ("`nTOP {0} PASTAS ({1} GB = {2}% da primaria)" -f $topo.Count, (GB $somaTop), $(if ($usado) { [math]::Round(100 * $somaTop / $usado) } else { 0 })) -ForegroundColor Yellow
        $topo | Select-Object Pasta, Itens, GB, @{n='%';e={$_.PctCaixa}}, MaisAntigo, Tag, Status, @{n='GB/dia';e={$_.GBporDia}}, @{n='GB c/30d';e={$_.GBcom30d}} |
            Format-Table -AutoSize | Out-String -Width 300 | Write-Host

        foreach ($t in $topo) { $csv.Add(($t | Select-Object * -ExcludeProperty Bytes)) }
    }

    if ($csv.Count) {
        $arq = Join-Path $saida ("MailboxTopFolders_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
        $csv | Export-Csv -Path $arq -NoTypeInformation -Encoding UTF8
        Write-Host "CSV: $arq" -ForegroundColor Green
    }
}
finally {
    if ($conectouAqui) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }
}
