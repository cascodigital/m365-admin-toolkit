<#
.SYNOPSIS
    Gerencia Security Defaults e SMTP AUTH para usuario especifico

.DESCRIPTION
    Script unificado para ativar/desativar Security Defaults (MFA forcado) e 
    configurar SMTP AUTH individual em tenant Microsoft 365.
    
    Funcoes:
    - DESATIVAR: Remove Security Defaults + Habilita SMTP AUTH
    - ATIVAR: Restaura Security Defaults + Desabilita SMTP AUTH

.PARAMETER UserEmail
    Email do usuario para configurar SMTP AUTH individual

.EXAMPLE
    .\Manage-SecurityDefaults-SMTP.ps1 -UserEmail "user@domain.com"
    
    Executa menu interativo solicitando acao (ativar/desativar)

.NOTES
    Autor: Andre Kittler
    Versao: 1.1
    Requisitos:
        - Global Admin ou Exchange Admin
        - Modulos: Microsoft.Graph, ExchangeOnlineManagement
    

#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, HelpMessage="Email do usuario para configurar SMTP AUTH")]
    [ValidatePattern("^[\w\.-]+@[\w\.-]+\.\w+$")]
    [string]$UserEmail
)

# ============================================
# FUNCOES
# ============================================

function Install-RequiredModules {
    Write-Host "`nVerificando modulos..." -ForegroundColor Yellow
    
    # Verificar Microsoft.Graph (submódulo específico)
    try {
        Get-Command Connect-MgGraph -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Host "Instalando Microsoft.Graph.Authentication..." -ForegroundColor Cyan
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
    }
    
    # Verificar ExchangeOnlineManagement
    try {
        Get-Command Connect-ExchangeOnline -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Host "Instalando ExchangeOnlineManagement..." -ForegroundColor Cyan
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
    }
    
    Write-Host "Modulos OK" -ForegroundColor Green
}

function Set-SecurityDefaults {
    param([bool]$Enable)
    
    $action = if ($Enable) { "HABILITANDO" } else { "DESABILITANDO" }
    
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        
        # Connect-MgGraph importa apenas modulos necessarios (rapido)
        Connect-MgGraph -Scopes "Policy.ReadWrite.SecurityDefaults","Policy.Read.All" -NoWelcome
        
        Write-Host "`nStatus Security Defaults ATUAL:" -ForegroundColor Cyan
        Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy | Select-Object Id,IsEnabled
        
        Write-Host "`n$action Security Defaults..." -ForegroundColor Yellow
        $body = @{ isEnabled = $Enable }
        Update-MgPolicyIdentitySecurityDefaultEnforcementPolicy -BodyParameter $body
        
        Write-Host "`nStatus Security Defaults NOVO:" -ForegroundColor Cyan
        Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy | Select-Object Id,IsEnabled
        
        $status = if ($Enable) { "HABILITADO (MFA forcado)" } else { "DESABILITADO" }
        Write-Host "[OK] Security Defaults $status" -ForegroundColor Green
        
        if (!$Enable) {
            Write-Host "Aguarde 5-10 minutos para propagacao completa" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Error "ERRO Security Defaults: $_"
        throw
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }
}

function Set-UserSmtpAuth {
    param(
        [string]$Email,
        [bool]$Enable
    )
    
    $action = if ($Enable) { "HABILITANDO" } else { "DESABILITANDO" }
    
    try {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        
        Write-Host "`nStatus SMTP AUTH ATUAL ($Email):" -ForegroundColor Cyan
        Get-CASMailbox -Identity $Email | Select-Object DisplayName,SmtpClientAuthenticationDisabled
        
        Write-Host "`n$action SMTP AUTH..." -ForegroundColor Yellow
        Set-CASMailbox -Identity $Email -SmtpClientAuthenticationDisabled (!$Enable)
        
        Write-Host "`nStatus SMTP AUTH NOVO ($Email):" -ForegroundColor Cyan
        Get-CASMailbox -Identity $Email | Select-Object DisplayName,SmtpClientAuthenticationDisabled
        
        $status = if ($Enable) { "HABILITADO" } else { "DESABILITADO" }
        Write-Host "[OK] SMTP AUTH $status para $Email" -ForegroundColor Green
    }
    catch {
        Write-Error "ERRO SMTP AUTH: $_"
        throw
    }
    finally {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
}

# ============================================
# MAIN
# ============================================

Clear-Host
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "  GERENCIAR SECURITY DEFAULTS + SMTP" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan

# Instalar modulos necessarios
Install-RequiredModules

# Menu acao
Write-Host "`nUsuario: $UserEmail" -ForegroundColor Yellow
Write-Host "`nSelecione a acao:" -ForegroundColor Cyan
Write-Host "1 - DESATIVAR Security Defaults + HABILITAR SMTP AUTH"
Write-Host "2 - ATIVAR Security Defaults + DESABILITAR SMTP AUTH"
Write-Host ""

do {
    $choice = Read-Host "Opcao (1 ou 2)"
} while ($choice -notin @('1','2'))

# Executar acao
try {
    switch ($choice) {
        '1' {
            Write-Host "`n=== MODO: DESATIVAR SEGURANCA ===" -ForegroundColor Yellow
            Set-SecurityDefaults -Enable $false
            Set-UserSmtpAuth -Email $UserEmail -Enable $true
        }
        '2' {
            Write-Host "`n=== MODO: ATIVAR SEGURANCA ===" -ForegroundColor Yellow
            Set-SecurityDefaults -Enable $true
            Set-UserSmtpAuth -Email $UserEmail -Enable $false
        }
    }
    
    Write-Host "`n======================================" -ForegroundColor Green
    Write-Host "  OPERACAO CONCLUIDA COM SUCESSO" -ForegroundColor Green
    Write-Host "======================================" -ForegroundColor Green
}
catch {
    Write-Host "`n======================================" -ForegroundColor Red
    Write-Host "  ERRO NA EXECUCAO" -ForegroundColor Red
    Write-Host "======================================" -ForegroundColor Red
    Write-Error $_
    exit 1
}
