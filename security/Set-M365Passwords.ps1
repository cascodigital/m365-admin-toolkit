<#
.SYNOPSIS
    Gerador automatizado de senhas aleatorias para usuarios Microsoft 365 por dominio

.DESCRIPTION
    Script automatizado para geracao e aplicacao em massa de senhas aleatorias em usuarios 
    Microsoft 365/Azure AD. Utiliza o modulo Microsoft.Graph para conectividade moderna e 
    gera senhas no formato memoravel (AA1234qq) com 2 caracteres maiusculos, 4 numeros 
    e 2 caracteres minusculos.
    
    Funcionalidades principais:
    - Geracao de senhas aleatorias com formato padronizado e seguro
    - Filtragem automatica por dominio especifico
    - Processamento em massa com controle de throttling
    - Relatorio detalhado de sucessos e falhas
    - Exportacao automatica para CSV com timestamp
    - Validacao de permissoes e tratamento de erros robusto

.PARAMETER None
    Script interativo - solicita dominio alvo durante execucao

.EXAMPLE
    .\Generate-RandomPasswords.ps1
    # Script solicita: Digite o dominio para aplicar a nova senha (exemplo: empresa.com.br)
    # Digite: cascodigital.com.br
    # Processa todos usuarios habilitados do dominio cascodigital.com.br
    # Confirma lista de usuarios encontrados
    # Aplica novas senhas e gera relatorio CSV

.INPUTS
    String - Dominio alvo inserido interativamente pelo usuario

.OUTPUTS
    - Arquivo CSV: Senhas_Usuarios_[dominio]_[timestamp].csv
    - Console: Lista de usuarios processados com status
    - Relatorio: Estatisticas de sucessos e falhas

.NOTES
    Autor         : Andre Kittler
    Versao        : 2.0
    Compatibilidade: PowerShell 5.1+, Windows/Linux/macOS
    
    Requisitos Microsoft Graph:
    - Modulo Microsoft.Graph instalado
    - Permissoes necessarias:
      * User.ReadWrite.All
      * Directory.ReadWrite.All
      * UserAuthenticationMethod.ReadWrite.All
      * Directory.AccessAsUser.All
    
    Privilegios administrativos necessarios:
    - Global Administrator OU
    - User Administrator OU  
    - Password Administrator
    
    Formato de senha gerada:
    - Padrao: AA1234qq (8 caracteres)
    - 2 caracteres maiusculos aleatorios
    - 4 numeros aleatorios
    - 2 caracteres minusculos aleatorios
    
    Configuracoes de seguranca:
    - ForceChangePasswordNextSignIn: false
    - Senhas definidas como finais (usuarios nao precisam alterar)
    - Throttling de 1 segundo entre processamentos

.LINK
    https://docs.microsoft.com/en-us/graph/api/user-update
    
.LINK
    https://docs.microsoft.com/en-us/graph/permissions-reference
#>


# Instalar e importar o modulo se necessario
if (!(Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Authentication

# Funcao para gerar senha no formato especificado
function New-RandomPassword {
    # 2 caracteres maiusculos
    $upperChars = -join ((65..90) | Get-Random -Count 2 | ForEach-Object {[char]$_})
    
    # 4 numeros
    $numbers = -join ((0..9) | Get-Random -Count 4)
    
    # 2 caracteres minusculos
    $lowerChars = -join ((97..122) | Get-Random -Count 2 | ForEach-Object {[char]$_})
    
    return $upperChars + $numbers + $lowerChars
}

# Conectar ao Microsoft Graph com permissoes adequadas
Write-Host "Conectando ao Microsoft Graph..." -ForegroundColor Green
Write-Host "IMPORTANTE: Certifique-se de estar logado com uma conta que tenha privilegios de:" -ForegroundColor Yellow
Write-Host "- Global Administrator OU" -ForegroundColor Yellow  
Write-Host "- User Administrator OU" -ForegroundColor Yellow
Write-Host "- Password Administrator" -ForegroundColor Yellow

try {
    # Permissoes necessarias para reset de senhas
    $scopes = @(
        "User.ReadWrite.All",
        "Directory.ReadWrite.All", 
        "UserAuthenticationMethod.ReadWrite.All",
        "Directory.AccessAsUser.All"
    )
    
    Connect-MgGraph -Scopes $scopes
}
catch {
    Write-Host "Erro ao conectar: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Verificar contexto atual
$context = Get-MgContext
Write-Host "Conectado como: $($context.Account)" -ForegroundColor Green

# Solicitar o dominio
$targetDomain = Read-Host "Digite o dominio para aplicar a nova senha (exemplo: empresa.com.br)"

if ([string]::IsNullOrWhiteSpace($targetDomain)) {
    Write-Host "Dominio nao pode estar vazio!" -ForegroundColor Red
    exit
}

Write-Host "Buscando usuarios do dominio: $targetDomain" -ForegroundColor Yellow

# Buscar todos os usuarios
try {
    $allUsers = Get-MgUser -All -Property "Id,UserPrincipalName,DisplayName,AccountEnabled"
}
catch {
    Write-Host "Erro ao buscar usuarios: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Verifique se voce tem permissoes adequadas." -ForegroundColor Red
    exit
}

# Filtrar apenas usuarios do dominio especificado e que estao habilitados
$targetUsers = $allUsers | Where-Object { 
    $_.UserPrincipalName -like "*@$targetDomain" -and 
    $_.AccountEnabled -eq $true 
}

if ($targetUsers.Count -eq 0) {
    Write-Host "Nenhum usuario encontrado no dominio $targetDomain" -ForegroundColor Red
    exit
}

Write-Host "Encontrados $($targetUsers.Count) usuarios no dominio $targetDomain" -ForegroundColor Green

# Mostrar usuarios encontrados
Write-Host "`nUsuarios que terao a senha alterada:" -ForegroundColor Cyan
foreach ($user in $targetUsers) {
    Write-Host "- $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor White
}

# Confirmar acao
$confirm = Read-Host "`nDeseja continuar e alterar as senhas destes usuarios? (S/N)"
if ($confirm -ne "S" -and $confirm -ne "s") {
    Write-Host "Operacao cancelada pelo usuario." -ForegroundColor Yellow
    exit
}

# Array para armazenar resultados
$results = @()

# Processar cada usuario
Write-Host "`nProcessando usuarios..." -ForegroundColor Green
foreach ($user in $targetUsers) {
    try {
        # Gerar nova senha
        $newPassword = New-RandomPassword
        
        Write-Host "Processando: $($user.UserPrincipalName)..." -ForegroundColor Yellow
        
        # Atualizar senha do usuario (metodo alternativo)
        $passwordProfile = @{
            Password = $newPassword
            ForceChangePasswordNextSignIn = $false
        }
        
        # Tentar atualizar a senha
        Update-MgUser -UserId $user.Id -PasswordProfile $passwordProfile -ErrorAction Stop
        
        # Adicionar ao resultado
        $results += [PSCustomObject]@{
            'Nome' = $user.DisplayName
            'Email' = $user.UserPrincipalName
            'Nova Senha' = $newPassword
            'Status' = 'Sucesso'
        }
        
        Write-Host "✓ Senha alterada para: $($user.UserPrincipalName)" -ForegroundColor Green
        
        # Pequena pausa para evitar throttling
        Start-Sleep -Milliseconds 1000
    }
    catch {
        # Em caso de erro
        $errorMessage = $_.Exception.Message
        
        $results += [PSCustomObject]@{
            'Nome' = $user.DisplayName
            'Email' = $user.UserPrincipalName
            'Nova Senha' = 'ERRO'
            'Status' = $errorMessage
        }
        
        Write-Host "✗ Erro ao alterar senha para: $($user.UserPrincipalName)" -ForegroundColor Red
        Write-Host "  Erro: $errorMessage" -ForegroundColor Red
        
        # Se for erro de permissao, dar dica
        if ($errorMessage -like "*Insufficient privileges*" -or $errorMessage -like "*Authorization_RequestDenied*") {
            Write-Host "  DICA: Voce precisa de privilegios administrativos para alterar senhas!" -ForegroundColor Red
        }
    }
}

# Gerar arquivo Excel
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$excelPath = "Senhas_Usuarios_$($targetDomain.Replace('.','_'))_$timestamp.csv"

try {
    $results | Export-Csv -Path $excelPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nArquivo gerado: $excelPath" -ForegroundColor Green
    
    # Tentar abrir o arquivo
    if (Test-Path $excelPath) {
        $openFile = Read-Host "Deseja abrir o arquivo agora? (S/N)"
        if ($openFile -eq "S" -or $openFile -eq "s") {
            Start-Process $excelPath
        }
    }
}
catch {
    Write-Host "Erro ao gerar arquivo CSV: $($_.Exception.Message)" -ForegroundColor Red
}

# Resumo final
Write-Host "`n=== RESUMO ===" -ForegroundColor Cyan
Write-Host "Dominio processado: $targetDomain" -ForegroundColor White
Write-Host "Total de usuarios: $($targetUsers.Count)" -ForegroundColor White
Write-Host "Sucessos: $(($results | Where-Object {$_.Status -eq 'Sucesso'}).Count)" -ForegroundColor Green
Write-Host "Erros: $(($results | Where-Object {$_.Status -ne 'Sucesso'}).Count)" -ForegroundColor Red

# Mostrar usuarios com sucesso
$sucessos = $results | Where-Object {$_.Status -eq 'Sucesso'}
if ($sucessos.Count -gt 0) {
    Write-Host "`nSenhas alteradas com sucesso:" -ForegroundColor Green
    foreach ($sucesso in $sucessos) {
        Write-Host "✓ $($sucesso.Email) - Nova senha: $($sucesso.'Nova Senha')" -ForegroundColor Green
    }
}

# Desconectar
Disconnect-MgGraph

Write-Host "`nScript finalizado!" -ForegroundColor Green
Write-Host "As senhas foram definidas como finais - usuarios nao precisam alterar." -ForegroundColor Yellow

if (($results | Where-Object {$_.Status -ne 'Sucesso'}).Count -gt 0) {
    Write-Host "`nIMPORTANTE: Alguns usuarios tiveram erro. Verifique se voce tem:" -ForegroundColor Red
    Write-Host "- Privilegios de Global Administrator, User Administrator ou Password Administrator" -ForegroundColor Red
    Write-Host "- Permissoes adequadas no tenant" -ForegroundColor Red
}
