<#
.SYNOPSIS
    Remove inbox rules corrompidas de uma mailbox, pelo NOME.

.DESCRIPTION
    As regras corrompidas NAO sao retornadas por Get-InboxRule (falham no parse e sao
    puladas, emitindo apenas um WARNING). Por isso InError e sempre False e elas so
    podem ser removidas pelo NOME.

    Metrica de sucesso: a contagem de WARNINGs de Get-InboxRule deve cair para 0.
    O script mede isso antes e depois, entao voce tem prova, nao achismo.

.PARAMETER DryRun
    So mostra o que faria. NAO apaga nada. Rode assim primeiro.

.PARAMETER NukeOutlookBlob
    ULTIMO RECURSO. Adiciona -AlwaysDeleteOutlookRulesBlob, que apaga o blob de regras
    client-side do Outlook. Isso DESTROI regras client-only (tocar som, alerta de
    desktop, mover para PST local) que NAO estao em backup nenhum porque nunca chegam
    ao servidor. So use com autorizacao explicita do dono da mailbox.

.EXAMPLE
    .\Remove-CorruptInboxRules.ps1 -MailboxAddress user@contoso.com -CorruptRules 'Regra A','Regra B' -DryRun
    .\Remove-CorruptInboxRules.ps1 -MailboxAddress user@contoso.com -CorruptRules 'Regra A','Regra B'
#>

param(
    [Parameter(Mandatory = $true)][string]$MailboxAddress,
    # Nomes copiados dos WARNINGs "contains errors" do Get-InboxRule
    [Parameter(Mandatory = $true)][string[]]$CorruptRules,
    [switch]$DryRun,
    [switch]$NukeOutlookBlob
)

$ErrorActionPreference = 'Stop'

$stamp   = Get-Date -Format 'yyyyMMdd_HHmm'
$deskDir = [Environment]::GetFolderPath('Desktop')

# ---------------------------------------------------------------------------
# 0. Sessao viva?
# ---------------------------------------------------------------------------
$conn = Get-ConnectionInformation -ErrorAction SilentlyContinue |
        Where-Object { $_.State -eq 'Connected' -and $_.TokenStatus -eq 'Active' }
if (-not $conn) {
    Write-Error "Sem sessao EXO ativa. Conecte primeiro (e confirme o PIM Exchange Recipient Administrator)."
    exit 1
}

# ---------------------------------------------------------------------------
# 1. BASELINE - conta os warnings ANTES (esta e a metrica que importa)
# ---------------------------------------------------------------------------
Write-Host "`n=== BASELINE ===" -ForegroundColor Cyan

$before = Get-InboxRule -Mailbox $MailboxAddress -IncludeHidden `
                        -WarningVariable warnBefore -WarningAction SilentlyContinue

Write-Host ("Regras legiveis : {0}" -f $before.Count)
Write-Host ("Regras em erro  : {0}   <-- warnings do servidor" -f $warnBefore.Count) -ForegroundColor Yellow
$warnBefore | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkYellow }

# ---------------------------------------------------------------------------
# 2. BACKUP das legiveis (as corrompidas sao insalvaveis por definicao)
# ---------------------------------------------------------------------------
$xml = Join-Path $deskDir "JW_Rules_PreDelete_$stamp.xml"
$csv = Join-Path $deskDir "JW_Rules_PreDelete_$stamp.csv"

$before | Export-Clixml -Path $xml
$before | Select-Object Name, Enabled, Priority, InError, ErrorType, RuleIdentity, Description |
    Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8

Write-Host "`nBackup salvo:" -ForegroundColor Green
Write-Host "  $xml"
Write-Host "  $csv"

if ($before.Count -eq 0) {
    Write-Error "Backup vazio. Abortando por seguranca."
    exit 1
}

# ---------------------------------------------------------------------------
# 3. Remocao, uma por uma, com relatorio individual
# ---------------------------------------------------------------------------
Write-Host "`n=== REMOCAO ===" -ForegroundColor Cyan
if ($DryRun)          { Write-Host "MODO DRY-RUN: nada sera apagado." -ForegroundColor Magenta }
if ($NukeOutlookBlob) { Write-Warning "-AlwaysDeleteOutlookRulesBlob ATIVO: regras client-only serao destruidas." }

$results = foreach ($name in $CorruptRules) {

    if ($DryRun) {
        Write-Host "  [DRY] removeria: $name" -ForegroundColor Magenta
        [PSCustomObject]@{ Rule = $name; Status = 'DryRun'; Detail = '' }
        continue
    }

    $params = @{
        Mailbox     = $MailboxAddress
        Identity    = $name
        Confirm     = $false
        ErrorAction = 'Stop'
    }
    if ($NukeOutlookBlob) { $params['AlwaysDeleteOutlookRulesBlob'] = $true }

    try {
        Remove-InboxRule @params
        Write-Host "  [OK]   removida: $name" -ForegroundColor Green
        [PSCustomObject]@{ Rule = $name; Status = 'Removed'; Detail = '' }
    }
    catch {
        $msg = $_.Exception.Message
        Write-Host "  [FAIL] $name -> $msg" -ForegroundColor Red
        [PSCustomObject]@{ Rule = $name; Status = 'Failed'; Detail = $msg }
    }
}

$results | Export-Csv -Path (Join-Path $deskDir "JW_RemovalResult_$stamp.csv") `
                      -NoTypeInformation -Encoding UTF8

if ($DryRun) {
    Write-Host "`nDry-run concluido. Rode sem -DryRun para aplicar." -ForegroundColor Magenta
    exit 0
}

# ---------------------------------------------------------------------------
# 4. VERIFICACAO - os warnings cairam?
# ---------------------------------------------------------------------------
Write-Host "`n=== VERIFICACAO ===" -ForegroundColor Cyan
Start-Sleep -Seconds 5

$after = Get-InboxRule -Mailbox $MailboxAddress -IncludeHidden `
                       -WarningVariable warnAfter -WarningAction SilentlyContinue

Write-Host ("Regras em erro ANTES : {0}" -f $warnBefore.Count)
Write-Host ("Regras em erro DEPOIS: {0}" -f $warnAfter.Count)
Write-Host ("Regras legiveis      : {0} -> {1}" -f $before.Count, $after.Count)

if ($warnAfter.Count -eq 0) {
    Write-Host "`nSUCESSO: nenhuma regra corrompida restante no servidor." -ForegroundColor Green
    Write-Host "Proximo passo: o usuario precisa FECHAR o Outlook por completo" -ForegroundColor Cyan
    Write-Host "(confirmar que OUTLOOK.EXE morreu no Gerenciador de Tarefas) e reabrir." -ForegroundColor Cyan
}
else {
    Write-Warning "Ainda restam $($warnAfter.Count) regra(s) em erro:"
    $warnAfter | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkYellow }
    Write-Host "Se sobraram as mesmas, o Remove-InboxRule nao consegue toca-las." -ForegroundColor Yellow
    Write-Host "Ai o caminho e -NukeOutlookBlob, com autorizacao do usuario." -ForegroundColor Yellow
}
