<#
.SYNOPSIS
    Gerenciador interativo de aliases de email com ativacao automatica de Send From Alias no Microsoft 365

.DESCRIPTION
    Script corporativo para administracao completa de aliases de email em Exchange Online com interface
    interativa e configuracao automatica da funcionalidade SendFromAliasEnabled. Permite adicao e
    visualizacao de aliases secundarios para usuarios, com ativacao automatica da capacidade de envio
    a partir de aliases no Outlook Web App e clientes desktop.
    
    Funcionalidades principais:
    - Verificacao e ativacao automatica do SendFromAliasEnabled organizacional
    - Interface de menu interativo para gerenciamento continuo de aliases
    - Validacao automatica de usuarios e tratamento de erros robusto
    - Visualizacao clara de endereco principal vs aliases secundarios
    - Instrucoes detalhadas para uso pratico no Outlook Web
    - Atualizacao em tempo real da lista de aliases apos adicoes
    - Desconexao segura de sessoes Exchange Online
    
    Processo de configuracao:
    1. Verifica/habilita SendFromAliasEnabled no tenant (propagacao 60min-2h)
    2. Valida usuario alvo e exibe configuracao atual
    3. Menu interativo para adicoes multiplas sem reinicializacao
    4. Instrucoes step-by-step para uso no Outlook Web
    
    Casos de uso corporativos:
    - Configuracao de aliases para departamentos (vendas@, suporte@)
    - Setup de identidades multiplas para executivos 
    - Criacao de aliases funcionais para equipes
    - Configuracao de enderecos de resposta personalizados

.PARAMETER None
    Script totalmente interativo - solicita usuario alvo e aliases durante execucao

.EXAMPLE
    .\Manage-EmailAliases.ps1
    # Usuario: joao.silva@cascodigital.com.br
    # Novo alias: vendas@cascodigital.com.br
    # Resultado: Alias configurado + instrucoes Outlook Web

.EXAMPLE
    .\Manage-EmailAliases.ps1
    # Usuario: diretoria@cascodigital.com.br
    # Aliases: ceo@cascodigital.com.br, presidencia@cascodigital.com.br
    # Resultado: Multiplos aliases funcionais configurados

.EXAMPLE  
    # Setup de aliases para equipe de suporte
    .\Manage-EmailAliases.ps1
    # Usuario: suporte.tecnico@cascodigital.com.br
    # Aliases: helpdesk@cascodigital.com.br, ti@cascodigital.com.br
    # Resultado: Equipe pode responder com identidades especializadas

.INPUTS
    String - Email do usuario alvo para gerenciamento de aliases
    String - Novos aliases a serem adicionados (multiplos via menu interativo)

.OUTPUTS
    - Console: Interface interativa colorida com menu de opcoes
    - Exchange Online: Aliases configurados na mailbox do usuario
    - Tenant: SendFromAliasEnabled ativado se necessario
    - Instrucoes: Step-by-step para uso no Outlook Web

.NOTES
    Autor         : Andre Kittler
    Versao        : 1.0
    Compatibilidade: PowerShell 5.1+, Windows/Linux/macOS
    
    Requisitos obrigatorios:
    - Modulo ExchangeOnlineManagement (instalacao automatica)
    - Licenca Microsoft 365 valida (qualquer nivel)
    - Privilegios administrativos Exchange:
      * Organization Management OU
      * Recipient Management OU
      * Mail Recipients role
    
    Configuracoes aplicadas:
    - Set-OrganizationConfig -SendFromAliasEnabled $true
    - Set-Mailbox -EmailAddresses @{add="alias@domain.com"}
    - Propagacao: 60 minutos a 2 horas para ativacao completa
    
    Limitacoes e consideracoes:
    - SendFromAliasEnabled requer propagacao em todo tenant
    - Aliases devem ser de dominios aceitos (Accepted Domains)
    - Outlook Desktop pode precisar de reinicializacao pos-configuracao
    - Algunos clientes mobile podem nao suportar envio via alias
    - Limitacao de 400 aliases por mailbox (limite teorico)
    
    Instrucoes Outlook Web App:
    1. Novo email → Opcoes → Mostrar De
    2. Botao De → Outro endereco de email
    3. Digite alias → Alias fica salvo para proximos envios
    
    Instrucoes Outlook Desktop:
    1. Arquivo → Configuracoes de Conta → Configuracoes de Conta
    2. Selecionar conta → Alterar → Mais Configuracoes
    3. Aba Email → Endereco de Email → Adicionar alias
    
    Troubleshooting comum:
    - Erro "Alias already exists": Verificar se ja esta em uso
    - Propagacao lenta: Aguardar 2h ou reiniciar Outlook
    - Alias nao aparece: Verificar SendFromAliasEnabled no tenant

.LINK
    https://docs.microsoft.com/en-us/exchange/recipients-in-exchange-online/manage-user-mailboxes/add-or-remove-email-addresses

#>


# Define que o script para em caso de erro para o 'try/catch' funcionar
$ErrorActionPreference = 'Stop'

# Funcao para escrever mensagens coloridas
function Write-Log {
    param(
        [string]$Message,
        [string]$Color
    )
    Write-Host $Message -ForegroundColor $Color
}

# --- 1. Verificacao de Pre-requisitos ---
Write-Log "Verificando se o modulo 'ExchangeOnlineManagement' esta instalado..." -Color Cyan
if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
    Write-Log "Modulo nao encontrado. Instalando..." -Color Yellow
    try {
        Install-Module ExchangeOnlineManagement -Repository PSGallery -Force -AllowClobber -Scope CurrentUser
        Write-Log "Modulo instalado com sucesso." -Color Green
    }
    catch {
        Write-Log "ERRO: Ocorreu um erro ao instalar o modulo. Verifique sua conexao ou execute o PowerShell como Administrador." -Color Red
        return # Para a execucao
    }
}
else {
    Write-Log "Modulo ja instalado." -Color Green
}

# --- 2. Conexao com o Exchange Online ---
Write-Log "`nConectando ao Exchange Online..." -Color Cyan
try {
    # O -ShowBanner:$false apenas limpa a saida do terminal
    Connect-ExchangeOnline -ShowBanner:$false
    Write-Log "Conectado com sucesso!" -Color Green
}
catch {
    Write-Log "ERRO: Falha na autenticacao. Verifique as credenciais e tente novamente." -Color Red
    return
}

# --- Bloco principal de execucao ---
try {
    # --- 3. Habilitar Send As Alias ---
    Write-Log "`nVerificando a configuracao 'SendFromAliasEnabled' no tenant..." -Color Cyan
    $orgConfig = Get-OrganizationConfig
    if (-not $orgConfig.SendFromAliasEnabled) {
        Write-Log "A funcionalidade 'Enviar como Alias' esta desativada. Habilitando agora..." -Color Yellow
        Set-OrganizationConfig -SendFromAliasEnabled $true
        Write-Log "Funcionalidade habilitada com sucesso!" -Color Green
        Write-Log "AVISO: A alteracao pode levar de 60 minutos a algumas horas para ser propagada em todo o ambiente." -Color Yellow
    }
    else {
        Write-Log "A funcionalidade 'Enviar como Alias' ja esta ativa na organizacao." -Color Green
    }

    # --- 4. Solicitar Usuario Alvo ---
    while ($true) {
        try {
            $userEmail = Read-Host -Prompt "`nDigite o email do usuario para gerenciar os aliases"
            $mailbox = Get-Mailbox -Identity $userEmail
            Write-Log "Usuario '$($mailbox.DisplayName)' encontrado." -Color Green
            break # Sai do loop de validacao
        }
        catch {
            Write-Log "ERRO: Usuario nao encontrado. Por favor, verifique o email e tente novamente." -Color Red
        }
    }


    # --- 5. Menu Interativo de Gerenciamento ---
    $exitLoop = $false
    while (-not $exitLoop) {
        Clear-Host
        Write-Log "--- Gerenciador de Aliases para: $($userEmail) ---" -Color White

        $allAddresses = $mailbox | Select-Object -ExpandProperty EmailAddresses
        $aliases = $allAddresses | Where-Object { $_ -clike 'smtp:*' }
        $primaryAddress = $allAddresses | Where-Object { $_ -clike 'SMTP:*' }

        Write-Log "`nEndereco Principal:" -Color Cyan
        Write-Host "> $($primaryAddress.Replace('SMTP:', ''))"

        if ($aliases.Count -gt 0) {
            Write-Log "`nAliases Atuais:" -Color Cyan
            foreach ($alias in $aliases) {
                Write-Host "- $($alias.Replace('smtp:', ''))"
            }
        }
        else {
            Write-Log "`nEste usuario nao possui aliases secundarios." -Color Yellow
        }


        Write-Log "`nO que voce deseja fazer?" -Color White
        Write-Log "1. Adicionar novo alias" -Color White
        Write-Log "2. Sair" -Color White

        $choice = (Read-Host -Prompt "Escolha uma opcao").Trim()

        switch ($choice) {
            '1' {
                $newAlias = Read-Host -Prompt "Digite o novo alias (ex: novo.email@dominio.com)"
                try {
                    Set-Mailbox -Identity $userEmail -EmailAddresses @{add = $newAlias }
                    Write-Log "Alias '$($newAlias)' adicionado com sucesso!" -Color Green
                    $mailbox = Get-Mailbox -Identity $userEmail
                }
                catch {
                    Write-Log "ERRO: Nao foi possivel adicionar o alias. Mensagem: $($_.Exception.Message)" -Color Red
                }
                Read-Host "Pressione ENTER para continuar..."
            }
            '2' {
                $exitLoop = $true # Define a variavel de controle para sair do loop
            }
            default {
                Write-Log "Opcao invalida. Pressione ENTER para tentar novamente." -Color Red
                Read-Host
            }
        }
    }

    # --- 6. Instrucoes Finais ---
    Clear-Host
    Write-Log "----------------------------------------------------------------" -Color Green
    Write-Log " Como usar o alias para enviar emails no Outlook (Online) " -Color White
    Write-Log "----------------------------------------------------------------" -Color Green
    Write-Log "1. Abra o Outlook na web (outlook.office.com)." -Color White
    Write-Log "2. Clique em 'Novo email'." -Color White
    Write-Log "3. Na janela de nova mensagem, clique no menu 'Opcoes'." -Color White
    Write-Log "4. Selecione 'Mostrar De' (Show From). O campo 'De' (From) aparecera." -Color White
    Write-Log "5. Clique no botao 'De' e depois em 'Outro endereco de email...'." -Color White
    Write-Log "6. Digite o alias que voce deseja usar e envie o email." -Color White
    Write-Log "   (Na proxima vez, o alias ja estara na lista para ser selecionado diretamente)." -Color White
    Write-Log "----------------------------------------------------------------`n" -Color Green


}
catch {
    # Captura qualquer erro que possa ocorrer nos passos acima
    Write-Log "`nERRO INESPERADO: Ocorreu um problema durante a execucao." -Color Red
    Write-Log "Mensagem de erro: $($_.Exception.Message)" -Color Red
}
finally {
    # --- 7. Desconexao ---
    Write-Log "`nDesconectando da sessao do Exchange Online..." -Color Cyan
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Write-Log "Sessao desconectada." -Color Green
}

