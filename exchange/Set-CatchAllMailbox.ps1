<#
.SYNOPSIS
    Configuracao automatizada de catch-all para um ou multiplos dominios no Microsoft 365.

.DESCRIPTION
    Script PowerShell interativo para implementacao de coletor geral (catch-all) em tenants M365.
    - Configura dominios selecionados como InternalRelay (essencial para catch-all).
    - Cria Grupo de Distribuicao Dinamico que inclui automaticamente todos os usuarios reais do tenant.
    - Cria Regra de Transporte que redireciona emails para enderecos inexistentes.
    - Garante a exclusao de usuarios validos da regra (evita redirecionar quem ja possui caixa).
    - Realiza limpeza automatica de regras/grupos anteriores com os mesmos nomes/emails.

    Funcionalidades:
    - Suporte a multiplos dominios (separados por virgula).
    - Filtro OPATH otimizado para evitar erros de wildcard inicial (*).
    - Prioridade automatica (Prioridade 0) no fluxo de e-mail.
    - Reset total de configuracoes anteriores antes da nova aplicacao.

.PARAMETER None
    Script interativo via terminal.

.EXAMPLE
    .\Configure-CatchAll-MultiDomain.ps1
    # Dominios alvo: empresa1.com.br, empresa2.com.br
    # Email coletor: sac@empresa1.com.br
    # Resultado: Qualquer email enviado para enderecos inexistentes em ambos os dominios sera entregue no SAC.

.NOTES
    Autor          : Andre Kittler
    Versao         : 3.0
    Compatibilidade: PowerShell 5.1+ / ExchangeOnlineManagement V3
    Requisito      : Permissao de Global Admin ou Exchange Admin.

    IMPORTANTE: 
    - A propagacao da alteracao para 'InternalRelay' pode levar ate 60 minutos.
    - Durante este tempo, erros de 'Address not found' (NDR 5.1.10) sao normais.
#>

$ErrorActionPreference = 'Stop'
$WarningPreference = 'SilentlyContinue'

# --- INPUTS INTERATIVOS ---
$upnAdmin = Read-Host "Email do Administrador (UPN)"
$dominiosInput = Read-Host "Dominios alvo (separe por virgula: dom1.com, dom2.com)"
$emailColetor = Read-Host "Email que recebera tudo (Caixa Coletora)"
$emailGrupoExcecao = Read-Host "Email para o Grupo Dinamico de Excecao"

$nomeRegra = "Catch-All Global - MultiDominio"
$nomeGrupoExcecao = "Excecao Catch-all - Usuarios Reais"
$dominios = $dominiosInput.Split(',').Trim()

Write-Host "Iniciando conexao com Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
Connect-ExchangeOnline -UserPrincipalName $upnAdmin -ShowBanner:$false

try {
    # --- FASE 1: LIMPEZA (RESET) ---
    Write-Host "Limpando configuracoes anteriores para evitar duplicidade..." -ForegroundColor Yellow
    
    if (Get-TransportRule -Identity $nomeRegra -ErrorAction SilentlyContinue) {
        Remove-TransportRule -Identity $nomeRegra -Confirm:$false
        Write-Host "Regra antiga removida." -ForegroundColor Gray
    }

    if (Get-DynamicDistributionGroup -Identity $emailGrupoExcecao -ErrorAction SilentlyContinue) {
        Remove-DynamicDistributionGroup -Identity $emailGrupoExcecao -Confirm:$false
        Write-Host "Grupo antigo removido." -ForegroundColor Gray
    }

    # --- FASE 2: CONFIGURACAO DE DOMINIO ---
    foreach ($dom in $dominios) {
        Write-Host "Configurando dominio $dom como InternalRelay..." -ForegroundColor Cyan
        Set-AcceptedDomain -Identity $dom -DomainType InternalRelay
    }

    # --- FASE 3: GRUPO DINAMICO DE EXCECAO ---
    # Filtro simplificado para evitar erro de wildcard inicial no OPATH do Exchange
    $finalFilter = "(RecipientTypeDetails -eq 'UserMailbox')"
    
    Write-Host "Criando grupo de excecao para usuarios reais..." -ForegroundColor Cyan
    New-DynamicDistributionGroup -Name $nomeGrupoExcecao `
        -PrimarySmtpAddress $emailGrupoExcecao `
        -RecipientFilter $finalFilter `
        -ErrorAction Stop

    # --- FASE 4: REGRA DE TRANSPORTE CATCH-ALL ---
    Write-Host "Criando regra de transporte com prioridade 0..." -ForegroundColor Cyan
    $ruleParams = @{
        Name = $nomeRegra
        RecipientDomainIs = $dominios
        RedirectMessageTo = $emailColetor
        ExceptIfSentToMemberOf = $emailGrupoExcecao
        Priority = 0
        StopRuleProcessing = $true
    }
    New-TransportRule @ruleParams

    Write-Host "`n===========================================================" -ForegroundColor Green
    Write-Host "CONFIGURACAO CONCLUIDA COM SUCESSO" -ForegroundColor Green
    Write-Host "Dominios: $($dominios -join ', ')" -ForegroundColor White
    Write-Host "Destino: $emailColetor" -ForegroundColor White
    Write-Host "Propagacao: Aguarde 60 minutos para validacao final." -ForegroundColor Yellow
    Write-Host "===========================================================" -ForegroundColor Green
}
catch {
    Write-Host "ERRO NA EXECUCAO: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "Sessao encerrada." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false
