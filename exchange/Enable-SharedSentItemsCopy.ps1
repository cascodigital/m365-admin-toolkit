<#
.SYNOPSIS
    Faz a shared mailbox guardar no proprio Sent Items o que e enviado como ela.
.DESCRIPTION
    Liga MessageCopyForSentAsEnabled e MessageCopyForSendOnBehalfEnabled. Sem isso,
    o email enviado "como" a shared so aparece no Sent Items de quem enviou.
    Recusa se o alvo nao for SharedMailbox. Grava CSV de verificacao antes/depois.
    SIMULA POR PADRAO; aplica so com a chave EXECUTAR.
.PARAMETER Mailbox
    Endereco da shared mailbox.
.EXAMPLE
    .\Enable-SharedSentItemsCopy.ps1 -Mailbox shared@contoso.com
    Simula.
.EXAMPLE
    $global:EXECUTAR = $true; .\Enable-SharedSentItemsCopy.ps1 -Mailbox shared@contoso.com
    Aplica.
.NOTES
    Requer sessao do Exchange Online ja conectada (Connect-ExchangeOnline).
    Rollback: Set-Mailbox <shared> -MessageCopyForSentAsEnabled $false -MessageCopyForSendOnBehalfEnabled $false
    Autor: Andre Kittler da Costa
#>
# IMPACTO: liga copia no Sent Items da shared para Send As e Send on Behalf. 1 mailbox. Reversivel.
# ALVO: a shared mailbox passada em -Mailbox
# ROLLBACK: Set-Mailbox <shared> -MessageCopyForSentAsEnabled $false -MessageCopyForSendOnBehalfEnabled $false

# ---------- PREAMBULO DE SEGURANCA ----------
# Este script SIMULA por padrao. Para APLICAR, rode tudo num comando so:
#     $global:EXECUTAR = $true; .\Enable-SharedSentItemsCopy.ps1 -Mailbox shared@contoso.com
param([Parameter(Mandatory = $true)][string]$Mailbox)

$Execute = $false                 # <<< FALHA FECHADA: NUNCA remover esta linha
$__k = Get-Variable EXECUTAR -Scope Global -ValueOnly -ErrorAction SilentlyContinue
$Execute = ($__k -is [bool] -and $__k)
$global:EXECUTAR = $false         # consome a chave

$saida = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.ProviderPath }
try { $__p = Join-Path $saida "w$PID.tmp"; [IO.File]::WriteAllText($__p,'x'); [IO.File]::Delete($__p) }
catch { $saida = $env:TEMP; Write-Host "AVISO: saida redirecionada para $saida" -Foreground Yellow }

$alvo   = $Mailbox
$campos = 'Name','PrimarySmtpAddress','RecipientTypeDetails','MessageCopyForSentAsEnabled','MessageCopyForSendOnBehalfEnabled'

$antes = Get-Mailbox -Identity $alvo -ErrorAction Stop | Select-Object $campos
Write-Host "ALVOS: $(@($antes).Count)" -Foreground Cyan
if ($antes.RecipientTypeDetails -ne 'SharedMailbox') { throw "PARE: $alvo nao e SharedMailbox ($($antes.RecipientTypeDetails))" }

$p = @{ Identity = $alvo; MessageCopyForSentAsEnabled = $true; MessageCopyForSendOnBehalfEnabled = $true; ErrorAction = 'Stop' }
if ($Execute) {
    try { Set-Mailbox @p } catch { Write-Host "FALHOU: $($_.Exception.Message)" -Foreground Red }
} else {
    Write-Host "SIMULACAO: Set-Mailbox $(($p.GetEnumerator() | Sort-Object Key | ForEach-Object { "-$($_.Key) '$($_.Value)'" }) -join ' ')"
}

# ============ VERIFICACAO - ULTIMO BLOCO, NADA DEPOIS ============
if (-not $antes) { Write-Host "ATENCAO: `$antes VAZIO - esta tabela NAO e uma verificacao" -Foreground Red }
$depois = Get-Mailbox -Identity $alvo | Select-Object $campos
$tab = @(
  $antes  | Select-Object @{n='Fase';e={'ANTES' }},*
  $depois | Select-Object @{n='Fase';e={'DEPOIS'}},*
)
Write-Host ""
if (-not $Execute) {
    Write-Host '### MODO SIMULACAO - NADA FOI ALTERADO ###' -Foreground Yellow
    Write-Host '### Rode  $global:EXECUTAR = $true; .\Enable-SharedSentItemsCopy.ps1  p/ aplicar ###' -Foreground Yellow
} else {
    Write-Host '### APLICADO - confira ANTES x DEPOIS ###' -Foreground Green
}
$tab | Format-Table -AutoSize
$csv = Join-Path $saida "verificacao-sentcopy-$(Get-Date -f yyyyMMdd-HHmmss).csv"
$tab | Export-Csv $csv -NoTypeInformation -Encoding UTF8
if (Test-Path $csv) { Write-Host "Resumo salvo em: $csv" }
else { Write-Host "ATENCAO: o CSV NAO foi gravado" -Foreground Red }
