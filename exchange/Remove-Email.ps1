<#
.SYNOPSIS
    Remocao automatizada de emails em Microsoft 365 com verificacao e instalacao automatica de dependencias

.DESCRIPTION
    Script corporativo autocontido - versao 3.1 com correcao de desinstalacao
    
.NOTES
    Autor         : Andre Kittler
    Versao        : 3.1 (Corrigido erro AllowPrerelease)
    Data          : 2025-11-23
#>

function Write-ColorMessage {
    param(
        [string]$Message,
        [string]$Type = "Info"
    )
    
    switch ($Type) {
        "Success" { Write-Host "✓ $Message" -ForegroundColor Green }
        "Error" { Write-Host "✗ $Message" -ForegroundColor Red }
        "Warning" { Write-Host "⚠ $Message" -ForegroundColor Yellow }
        "Info" { Write-Host "ℹ $Message" -ForegroundColor Cyan }
        "Progress" { Write-Host "→ $Message" -ForegroundColor White }
        default { Write-Host $Message }
    }
}

function Test-ModuleVersion {
    $installedModule = Get-InstalledModule -Name ExchangeOnlineManagement -AllVersions -ErrorAction SilentlyContinue | 
        Sort-Object Version -Descending | 
        Select-Object -First 1
    
    if (!$installedModule) {
        return $null
    }
    
    return $installedModule.Version
}

function Install-RequiredModule {
    Write-ColorMessage "=== VERIFICAÇÃO DE DEPENDÊNCIAS ===" -Type "Info"
    Write-Host ""
    
    $currentVersion = Test-ModuleVersion
    
    if ($currentVersion) {
        Write-ColorMessage "Versão instalada: $currentVersion" -Type "Progress"
    } else {
        Write-ColorMessage "ExchangeOnlineManagement não encontrado" -Type "Warning"
    }
    
    # Verificar se precisa atualizar
    $needsUpdate = $false
    
    if (!$currentVersion) {
        $needsUpdate = $true
        Write-ColorMessage "Módulo não instalado - instalação necessária" -Type "Warning"
    } elseif ($currentVersion -lt [version]"3.9.1") {
        $needsUpdate = $true
        Write-ColorMessage "Versão antiga detectada - atualização necessária (mínimo: 3.9.1-Preview1)" -Type "Warning"
    }
    
    if ($needsUpdate) {
        Write-Host ""
        Write-ColorMessage "Para executar este script é necessário ExchangeOnlineManagement 3.9.1-Preview1+" -Type "Info"
        $confirm = Read-Host "Deseja instalar/atualizar automaticamente? (S/N)"
        
        if ($confirm -ne 'S' -and $confirm -ne 's') {
            Write-ColorMessage "Operação cancelada pelo usuário" -Type "Error"
            Write-Host ""
            Write-ColorMessage "Para instalar manualmente execute:" -Type "Info"
            Write-Host "Install-Module ExchangeOnlineManagement -AllowPrerelease -Force" -ForegroundColor Gray
            exit
        }
        
        Write-Host ""
        Write-ColorMessage "Instalando ExchangeOnlineManagement Preview..." -Type "Progress"
        
        try {
            # Remover módulo da sessão atual
            if (Get-Module ExchangeOnlineManagement) {
                Write-ColorMessage "Removendo módulo da sessão..." -Type "Progress"
                Remove-Module ExchangeOnlineManagement -Force -ErrorAction SilentlyContinue
            }
            
            # Desinstalar versões antigas (corrigido)
            if ($currentVersion) {
                Write-ColorMessage "Desinstalando versões antigas..." -Type "Progress"
                
                $allVersions = Get-InstalledModule ExchangeOnlineManagement -AllVersions -ErrorAction SilentlyContinue
                foreach ($version in $allVersions) {
                    try {
                        Write-Host "  Removendo versão $($version.Version)..." -ForegroundColor Gray
                        Uninstall-Module -Name ExchangeOnlineManagement -RequiredVersion $version.Version -Force -ErrorAction Stop
                    } catch {
                        Write-ColorMessage "  Aviso: Não foi possível remover versão $($version.Version)" -Type "Warning"
                    }
                }
            }
            
            # Aguardar um momento
            Start-Sleep -Seconds 2
            
            # Instalar versão Preview
            Write-ColorMessage "Instalando versão Preview mais recente..." -Type "Progress"
            Install-Module -Name ExchangeOnlineManagement -AllowPrerelease -Force -Scope CurrentUser -SkipPublisherCheck -ErrorAction Stop
            
            # Importar módulo
            Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop
            
            $newVersion = (Get-InstalledModule ExchangeOnlineManagement -AllVersions | Sort-Object Version -Descending | Select-Object -First 1).Version
            Write-ColorMessage "Instalado com sucesso: versão $newVersion" -Type "Success"
            
            Write-Host ""
            Write-ColorMessage "IMPORTANTE: Reinicie o PowerShell para garantir que a nova versão seja carregada" -Type "Warning"
            $restart = Read-Host "Deseja sair agora e reiniciar o PowerShell? (S/N)"
            
            if ($restart -eq 'S' -or $restart -eq 's') {
                Write-ColorMessage "Feche esta janela e abra um novo PowerShell, depois execute o script novamente" -Type "Info"
                exit
            }
            
        } catch {
            Write-ColorMessage "Erro ao instalar: $($_.Exception.Message)" -Type "Error"
            Write-Host ""
            Write-ColorMessage "SOLUÇÃO MANUAL:" -Type "Warning"
            Write-Host "1. Feche TODAS as janelas PowerShell abertas" -ForegroundColor Gray
            Write-Host "2. Abra nova janela PowerShell como Administrador" -ForegroundColor Gray
            Write-Host "3. Execute os seguintes comandos:" -ForegroundColor Gray
            Write-Host ""
            Write-Host "   # Remover versões antigas" -ForegroundColor Yellow
            Write-Host "   Get-InstalledModule ExchangeOnlineManagement -AllVersions | Uninstall-Module -Force" -ForegroundColor White
            Write-Host ""
            Write-Host "   # Instalar Preview" -ForegroundColor Yellow
            Write-Host "   Install-Module ExchangeOnlineManagement -AllowPrerelease -Force -SkipPublisherCheck" -ForegroundColor White
            Write-Host ""
            Write-Host "4. Feche e abra novo PowerShell" -ForegroundColor Gray
            Write-Host "5. Execute este script novamente" -ForegroundColor Gray
            Write-Host ""
            exit
        }
    } else {
        Write-ColorMessage "Versão adequada já instalada" -Type "Success"
        Import-Module ExchangeOnlineManagement -Force
    }
    
    Write-Host ""
}

function Connect-M365Services {
    Write-ColorMessage "=== CONEXÃO AOS SERVIÇOS ===" -Type "Info"
    Write-Host ""
    
    # Verificar se Connect-IPPSSession suporta -EnableSearchOnlySession
    $supportsSearchOnly = $false
    try {
        $params = (Get-Command Connect-IPPSSession -ErrorAction Stop).Parameters
        $supportsSearchOnly = $params.ContainsKey('EnableSearchOnlySession')
    } catch {
        Write-ColorMessage "Erro ao verificar comando Connect-IPPSSession" -Type "Error"
        Write-ColorMessage "Certifique-se de ter reiniciado o PowerShell após a instalação" -Type "Warning"
        exit
    }
    
    if (!$supportsSearchOnly) {
        Write-ColorMessage "ERRO: Parâmetro -EnableSearchOnlySession não disponível" -Type "Error"
        Write-ColorMessage "Isso pode significar que:" -Type "Warning"
        Write-Host "  1. A versão Preview não foi instalada corretamente" -ForegroundColor Gray
        Write-Host "  2. O PowerShell não foi reiniciado após a instalação" -ForegroundColor Gray
        Write-Host "  3. Uma versão antiga ainda está em cache" -ForegroundColor Gray
        Write-Host ""
        Write-ColorMessage "SOLUÇÃO: Feche TODAS janelas PowerShell, abra nova janela e execute o script novamente" -Type "Info"
        exit
    }
    
    Write-ColorMessage "Suporte a -EnableSearchOnlySession: OK" -Type "Success"
    Write-Host ""
    
    # Conectar ao Compliance
    Write-ColorMessage "Conectando ao Security & Compliance Center..." -Type "Progress"
    try {
        Get-ComplianceSearch -ResultSize 1 -ErrorAction Stop | Out-Null
        Write-ColorMessage "Já conectado!" -Type "Success"
    } catch {
        try {
            Connect-IPPSSession -EnableSearchOnlySession -ShowBanner:$false -ErrorAction Stop
            Write-ColorMessage "Conectado com sucesso!" -Type "Success"
        } catch {
            Write-ColorMessage "Erro ao conectar: $($_.Exception.Message)" -Type "Error"
            exit
        }
    }
    
    Write-Host ""
}

function Remove-EmailBySearch {
    Write-ColorMessage "=== REMOÇÃO DE E-MAILS ===" -Type "Info"
    Write-Host ""
    
    # Solicitar dados
    $sender = Read-Host "Digite o endereço do remetente"
    if ([string]::IsNullOrWhiteSpace($sender)) {
        Write-ColorMessage "Remetente não pode estar vazio!" -Type "Error"
        exit
    }
    
    $subject = Read-Host "Digite o assunto do e-mail"
    if ([string]::IsNullOrWhiteSpace($subject)) {
        Write-ColorMessage "Assunto não pode estar vazio!" -Type "Error"
        exit
    }
    
    # Gerar nome único
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $searchName = "RemocaoEmail_$timestamp"
    $searchQuery = "(from:$sender) AND (subject:'$subject')"
    
    Write-Host ""
    Write-ColorMessage "RESUMO DA OPERAÇÃO:" -Type "Info"
    Write-Host "  Remetente: $sender"
    Write-Host "  Assunto: $subject"
    Write-Host "  Query: $searchQuery"
    Write-Host "  Busca: $searchName"
    Write-Host ""
    
    $confirm = Read-Host "Prosseguir com a remoção? (S/N)"
    if ($confirm -ne 'S' -and $confirm -ne 's') {
        Write-ColorMessage "Operação cancelada" -Type "Warning"
        exit
    }
    
    Write-Host ""
    Write-ColorMessage "=== PROCESSANDO ===" -Type "Info"
    
    try {
        # 1. Criar busca
        Write-ColorMessage "1/5 Criando busca de conformidade..." -Type "Progress"
        New-ComplianceSearch -Name $searchName -ExchangeLocation All -ContentMatchQuery $searchQuery -ErrorAction Stop | Out-Null
        Write-ColorMessage "Busca criada" -Type "Success"
        
        # 2. Iniciar busca
        Write-ColorMessage "2/5 Iniciando busca..." -Type "Progress"
        Start-ComplianceSearch -Identity $searchName -ErrorAction Stop
        Write-ColorMessage "Busca iniciada" -Type "Success"
        
        # 3. Aguardar conclusão
        Write-ColorMessage "3/5 Aguardando conclusão da busca..." -Type "Progress"
        $attempts = 0
        do {
            Start-Sleep -Seconds 10
            $searchStatus = Get-ComplianceSearch -Identity $searchName
            $attempts++
            Write-Host "    Status: $($searchStatus.Status) (tentativa $attempts)" -ForegroundColor Gray
            
            if ($searchStatus.Status -eq 'Failed') {
                throw "Busca falhou"
            }
        } while ($searchStatus.Status -ne 'Completed' -and $attempts -lt 60)
        
        if ($searchStatus.Status -ne 'Completed') {
            throw "Timeout: busca não completou em 10 minutos"
        }
        
        Write-ColorMessage "Busca concluída" -Type "Success"
        
        # 4. Verificar resultados
        $itemCount = $searchStatus.Items
        Write-ColorMessage "4/5 Itens encontrados: $itemCount" -Type "Info"
        
        if ($itemCount -eq 0) {
            Write-ColorMessage "Nenhum e-mail encontrado com os critérios especificados" -Type "Warning"
            Remove-ComplianceSearch -Identity $searchName -Confirm:$false
            exit
        }
        
        # 5. Executar remoção
        Write-ColorMessage "5/5 Executando remoção (SoftDelete)..." -Type "Progress"
        New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType SoftDelete -Confirm:$false -ErrorAction Stop | Out-Null
        
        Write-Host ""
        Write-ColorMessage "=== OPERAÇÃO CONCLUÍDA ===" -Type "Success"
        Write-Host ""
        Write-ColorMessage "$itemCount e-mail(s) removido(s) com sucesso" -Type "Success"
        Write-ColorMessage "E-mails movidos para Itens Recuperáveis (retenção: 30 dias)" -Type "Info"
        Write-ColorMessage "Nome da busca: $searchName" -Type "Info"
        
    } catch {
        Write-Host ""
        Write-ColorMessage "ERRO: $($_.Exception.Message)" -Type "Error"
        
        # Cleanup
        try {
            if (Get-ComplianceSearch -Identity $searchName -ErrorAction SilentlyContinue) {
                Remove-ComplianceSearch -Identity $searchName -Confirm:$false -ErrorAction SilentlyContinue
                Write-ColorMessage "Busca de teste removida" -Type "Info"
            }
        } catch {}
    }
}

# ===== EXECUÇÃO PRINCIPAL =====
Clear-Host
Write-Host ""
Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║          REMOÇÃO DE E-MAILS - MICROSOFT 365                  ║" -ForegroundColor Cyan
Write-Host "║          Script com Auto-Configuração v3.1                   ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

try {
    # Etapa 1: Verificar/Instalar módulo
    Install-RequiredModule
    
    # Etapa 2: Conectar aos serviços
    Connect-M365Services
    
    # Etapa 3: Executar remoção
    Remove-EmailBySearch
    
} catch {
    Write-ColorMessage "Erro fatal: $($_.Exception.Message)" -Type "Error"
}

Write-Host ""
$null = Read-Host "Pressione ENTER para sair"
